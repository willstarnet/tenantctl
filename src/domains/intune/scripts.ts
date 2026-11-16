import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { attempt, compact } from '../../core/util.ts';
import { INTUNE_READ_ROLES, pickByName, READ_CONFIG, scan } from './common.ts';

const READ_SCRIPTS = 'DeviceManagementScripts.Read.All';
const ALL = Number.POSITIVE_INFINITY;
/** Script output and errors can be whole transcripts; keep the compact view readable. */
export const OUTPUT_LIMIT = 500;

type ScriptKind = 'platform' | 'remediation' | 'macosShell';

const KINDS: Record<ScriptKind, { path: string; label: string; select: string }> = {
  platform: {
    path: '/deviceManagement/deviceManagementScripts',
    label: 'platform scripts',
    select: 'id,displayName,description,runAsAccount,enforceSignatureCheck,runAs32Bit,fileName,lastModifiedDateTime',
  },
  remediation: {
    path: '/deviceManagement/deviceHealthScripts',
    label: 'remediations',
    select:
      'id,displayName,description,publisher,version,runAsAccount,enforceSignatureCheck,runAs32Bit,isGlobalScript,' +
      'deviceHealthScriptType,lastModifiedDateTime',
  },
  macosShell: {
    path: '/deviceManagement/deviceShellScripts',
    label: 'macOS shell scripts',
    select:
      'id,displayName,description,runAsAccount,executionFrequency,retryCount,blockExecutionNotifications,fileName,' +
      'lastModifiedDateTime',
  },
};
const KIND_NAMES = Object.keys(KINDS) as ScriptKind[];

interface Script {
  id: string;
  displayName?: string | null;
  runAsAccount?: string | null;
  enforceSignatureCheck?: boolean | null;
  executionFrequency?: string | null;
  isGlobalScript?: boolean | null;
  lastModifiedDateTime?: string | null;
  [key: string]: unknown;
}

interface KindedScript {
  kind: ScriptKind;
  id: string;
  name: string;
  script: Script;
}

interface ScriptAssignment {
  runSchedule?: { '@odata.type'?: string; interval?: number | null; time?: string | null; date?: string | null };
  [key: string]: unknown;
}

const truncate = (text: string | null | undefined) =>
  text && text.length > OUTPUT_LIMIT ? `${text.slice(0, OUTPUT_LIMIT)}… (${text.length} chars)` : text || null;

async function listKind(ctx: Context, kind: ScriptKind, full: boolean): Promise<KindedScript[]> {
  const spec = KINDS[kind];
  const { items } = await ctx.graph.list<Script>(spec.path, {
    version: 'beta',
    query: full ? {} : { $select: spec.select },
    limit: ALL,
  });
  return items.map((script) => ({ kind, id: script.id, name: script.displayName ?? script.id, script }));
}

/**
 * All three kinds; a kind the caller cannot read lands in `unavailable` (list output reports it in `meta`) instead
 * of failing the command. When no kind is readable the first error is thrown.
 */
async function listScripts(ctx: Context, full: boolean) {
  const loaded = await Promise.allSettled(KIND_NAMES.map((kind) => listKind(ctx, kind, full)));
  if (loaded.every((r) => r.status === 'rejected')) throw (loaded[0] as PromiseRejectedResult).reason;
  const scripts: KindedScript[] = [];
  const unavailable: { section: ScriptKind; error: { code: string; message: string } }[] = [];
  loaded.forEach((result, index) => {
    const kind = KIND_NAMES[index] as ScriptKind;
    if (result.status === 'fulfilled') {
      scripts.push(...result.value);
      return;
    }
    const reason: unknown = result.reason;
    unavailable.push({
      section: kind,
      error: {
        code: reason instanceof CliError ? reason.code : 'INTERNAL',
        message: reason instanceof Error ? reason.message : String(reason),
      },
    });
  });
  return { scripts, unavailable };
}

/** `#microsoft.graph.deviceHealthScriptDailySchedule` + interval 1 + time -> `daily, every 1, at 09:00:00`. */
function describeSchedule(schedule: ScriptAssignment['runSchedule']): string | null {
  if (!schedule) return null;
  const type = String(schedule['@odata.type'] ?? '')
    .replace(/^#?microsoft\.graph\.deviceHealthScript/, '')
    .replace(/Schedule$/, '')
    .toLowerCase();
  const parts = [type || 'schedule'];
  if (schedule.interval) parts.push(`every ${schedule.interval}`);
  if (schedule.date) parts.push(`on ${schedule.date}`);
  if (schedule.time) parts.push(`at ${schedule.time}`);
  return parts.join(', ');
}

/** Run frequency: shell scripts carry an ISO 8601 duration; remediations a schedule per assignment. */
function frequencyOf(item: KindedScript, assignments: ScriptAssignment[] | null) {
  if (item.kind === 'macosShell') return item.script.executionFrequency || 'once';
  if (item.kind === 'remediation') {
    const schedules = [...new Set((assignments ?? []).map((a) => describeSchedule(a.runSchedule)).filter(Boolean))];
    return schedules.length ? schedules.join('; ') : null;
  }
  return 'once';
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'scripts list',
    summary: 'List Intune platform scripts, Remediations and macOS shell scripts',
    description:
      'Merges Windows platform (PowerShell) scripts, Remediations (device health scripts) and macOS shell ' +
      'scripts. runFrequency is "once" for platform scripts, the execution interval (ISO 8601 duration) for ' +
      'shell scripts and the per-assignment schedule for Remediations. assigned is null when the assignment ' +
      'lookup failed. A kind the caller cannot read is listed in meta.unavailable. Use `intune scripts status` ' +
      'for per-device results. (beta API: these endpoints have no v1.0 equivalent.)',
    examples: ['tenantctl intune scripts list', 'tenantctl intune scripts list --kind remediation'],
    args: [
      { name: 'kind', type: 'string', choices: KIND_NAMES, description: 'Only scripts of this kind.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: [READ_SCRIPTS, READ_CONFIG],
      application: [READ_SCRIPTS, READ_CONFIG],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const limit = listLimit(args);
      const { scripts, unavailable } = await listScripts(ctx, ctx.full);
      const matching = scripts
        .filter((s) => !args.kind || s.kind === args.kind)
        .sort((a, b) => a.kind.localeCompare(b.kind) || a.name.localeCompare(b.name));
      const page = matching.slice(0, limit);
      const results = await ctx.graph.batch<{ value?: ScriptAssignment[] }>(
        page.map((s) => ({ path: `${KINDS[s.kind].path}/${seg(s.id)}/assignments` })),
        'beta',
      );
      const items = page.map((s, index) => {
        const result = results[index];
        const assignments = result && !result.error ? (result.body?.value ?? []) : null;
        if (ctx.full) return { kind: s.kind, ...s.script, assignments };
        return compact({
          kind: s.kind,
          id: s.id,
          name: s.name,
          runAs: s.script.runAsAccount ?? null,
          runFrequency: frequencyOf(s, assignments),
          enforceSignatureCheck: s.kind === 'macosShell' ? undefined : (s.script.enforceSignatureCheck ?? null),
          builtIn: s.kind === 'remediation' ? (s.script.isGlobalScript ?? null) : undefined,
          assigned: assignments ? assignments.length > 0 : null,
          lastModifiedDateTime: s.script.lastModifiedDateTime ?? null,
        });
      });
      return {
        items,
        truncated: matching.length > page.length,
        ...(unavailable.length ? { unavailable } : {}),
      } satisfies ListResult<unknown>;
    },
  }),
  defineCommand({
    domain: 'intune',
    name: 'scripts status',
    summary: 'Show the run summary of one script and the devices where it failed',
    description:
      'Accepts the script name (or part of it) or id across platform scripts, Remediations and macOS shell ' +
      'scripts. summary is the Intune run summary (platform/shell: success and error device counts; Graph has no ' +
      'pending count for them. Remediations: detection and remediation counts). devices lists per-device ' +
      'results with error and output truncated to 500 characters (--full for raw objects); --failed-only keeps ' +
      'devices whose script errored (platform/shell runState fail or scriptError; Remediations detection ' +
      'scriptError or remediation remediationFailed/scriptError) and is filtered client-side. (beta API.)',
    examples: [
      'tenantctl intune scripts status "Set time zone"',
      'tenantctl intune scripts status BitLocker --failed-only',
    ],
    args: [
      { name: 'script', type: 'string', required: true, positional: true, description: 'Script name or id.' },
      { name: 'failedOnly', type: 'boolean', description: 'Only devices where the script failed.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_SCRIPTS], application: [READ_SCRIPTS], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const { scripts, unavailable } = await listScripts(ctx, false);
      const script = pickByName(scripts, args.script, 'script', (s) => s.name);
      const base = `${KINDS[script.kind].path}/${seg(script.id)}`;
      const remediation = script.kind === 'remediation';
      const [summary, states] = await Promise.all([
        attempt(ctx.graph.get<Record<string, unknown>>(`${base}/runSummary`, { version: 'beta' })),
        deviceStates(ctx, base, remediation, args.failedOnly === true, listLimit(args)),
      ]);
      return {
        script: { kind: script.kind, id: script.id, name: script.name },
        summary: 'error' in summary ? summary : ctx.full ? summary : shapeSummary(summary, remediation),
        devices: ctx.full ? states.items : states.items.map((s) => shapeState(s, remediation)),
        truncated: states.truncated,
        ...(unavailable.length ? { unavailable } : {}),
      };
    },
  }),
];

interface DeviceState {
  id: string;
  runState?: string | null;
  detectionState?: string | null;
  remediationState?: string | null;
  resultMessage?: string | null;
  errorCode?: number | null;
  errorDescription?: string | null;
  preRemediationDetectionScriptOutput?: string | null;
  preRemediationDetectionScriptError?: string | null;
  remediationScriptError?: string | null;
  postRemediationDetectionScriptOutput?: string | null;
  postRemediationDetectionScriptError?: string | null;
  lastStateUpdateDateTime?: string | null;
  managedDevice?: { id?: string; deviceName?: string | null } | null;
  [key: string]: unknown;
}

export function failed(state: DeviceState, remediation: boolean): boolean {
  if (remediation) {
    return (
      state.detectionState === 'scriptError' ||
      state.remediationState === 'remediationFailed' ||
      state.remediationState === 'scriptError'
    );
  }
  return state.runState === 'fail' || state.runState === 'scriptError';
}

/**
 * Per-device run states with the device name expanded (`$expand=managedDevice` is what the admin center sends but
 * is not in the docs; on a 400 the states come back without names).
 */
async function deviceStates(
  ctx: Context,
  base: string,
  remediation: boolean,
  failedOnly: boolean,
  limit: number,
): Promise<ListResult<DeviceState>> {
  const path = `${base}/deviceRunStates`;
  const load = (query: Record<string, string>) =>
    failedOnly
      ? scan<DeviceState>(ctx, path, query, (s) => failed(s, remediation), limit, 'beta')
      : ctx.graph.list<DeviceState>(path, { version: 'beta', query, limit });
  try {
    return await load({ $expand: 'managedDevice($select=id,deviceName)' });
  } catch (error) {
    if (!(error instanceof CliError && error.status === 400)) throw error;
    ctx.log('Graph rejected $expand=managedDevice on deviceRunStates; device names are omitted.');
    return load({});
  }
}

function shapeSummary(summary: Record<string, unknown>, remediation: boolean) {
  const n = (key: string) => (typeof summary[key] === 'number' ? (summary[key] as number) : null);
  if (!remediation) {
    return {
      success: n('successDeviceCount'),
      failed: n('errorDeviceCount'),
      successUsers: n('successUserCount'),
      failedUsers: n('errorUserCount'),
    };
  }
  return compact({
    noIssueDetected: n('noIssueDetectedDeviceCount'),
    issueDetected: n('issueDetectedDeviceCount'),
    detectionErrors: n('detectionScriptErrorDeviceCount'),
    detectionPending: n('detectionScriptPendingDeviceCount'),
    detectionNotApplicable: n('detectionScriptNotApplicableDeviceCount'),
    remediated: n('issueRemediatedDeviceCount'),
    remediationSkipped: n('remediationSkippedDeviceCount'),
    issueReoccurred: n('issueReoccurredDeviceCount'),
    remediationErrors: n('remediationScriptErrorDeviceCount'),
    remediatedLast30Days: n('issueRemediatedCumulativeDeviceCount'),
    lastRunDateTime: (summary.lastScriptRunDateTime as string | undefined) ?? null,
  });
}

function shapeState(state: DeviceState, remediation: boolean) {
  const device = {
    deviceName: state.managedDevice?.deviceName ?? null,
    deviceId: state.managedDevice?.id ?? undefined,
  };
  if (remediation) {
    return compact({
      ...device,
      failed: failed(state, true),
      detectionState: state.detectionState ?? null,
      remediationState: state.remediationState ?? null,
      error: truncate(
        state.remediationScriptError ||
          state.preRemediationDetectionScriptError ||
          state.postRemediationDetectionScriptError,
      ),
      output: truncate(state.postRemediationDetectionScriptOutput || state.preRemediationDetectionScriptOutput),
      lastStateUpdateDateTime: state.lastStateUpdateDateTime ?? null,
    });
  }
  return compact({
    ...device,
    failed: failed(state, false),
    runState: state.runState ?? null,
    errorCode: state.errorCode || undefined,
    error: truncate(state.errorDescription),
    output: truncate(state.resultMessage),
    lastStateUpdateDateTime: state.lastStateUpdateDateTime ?? null,
  });
}
