import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { pickOne } from '../../core/resolve.ts';
import { type AnyCommand, type ArgSpec, type Context, defineCommand, type Risk } from '../../core/types.ts';
import { compact, daysAgo, daysSince, isGuid, settle } from '../../core/util.ts';

/**
 * Microsoft Defender for Endpoint API (`https://api.security.microsoft.com/api`) through `ctx.defender`:
 * device inventory and response actions (isolate, scan, restrict, collect package, stop and quarantine).
 * Permissions are WindowsDefenderATP permissions, named as in the MDE API docs.
 */

export interface Machine {
  id: string;
  computerDnsName?: string | null;
  osPlatform?: string | null;
  version?: string | null;
  healthStatus?: string | null;
  riskScore?: string | null;
  exposureLevel?: string | null;
  lastSeen?: string | null;
  firstSeen?: string | null;
  lastIpAddress?: string | null;
  lastExternalIpAddress?: string | null;
  aadDeviceId?: string | null;
  machineTags?: string[] | null;
  onboardingStatus?: string | null;
  rbacGroupName?: string | null;
  [key: string]: unknown;
}

export interface MachineAction {
  id: string;
  type?: string | null;
  scope?: string | null;
  requestor?: string | null;
  requestorComment?: string | null;
  status?: string | null;
  machineId?: string | null;
  computerDnsName?: string | null;
  creationDateTimeUtc?: string | null;
  lastUpdateDateTimeUtc?: string | null;
  relatedFileInfo?: { fileIdentifier?: string | null; fileIdentifierType?: string | null } | null;
  [key: string]: unknown;
}

const LICENSE =
  'Needs Microsoft Defender for Endpoint P1/P2 or Defender for Business. Delegated use requires your own app ' +
  'registration with WindowsDefenderATP delegated permissions.';
const READ_ROLES = ['Security Reader', 'Security Operator', 'Security Administrator'];
const RESPONSE_ROLES = ['Security Operator', 'Security Administrator'];
/** List machines (application name; the delegated equivalent in the MDE docs is Machine.Read). */
const READ_MACHINES = 'Machine.Read.All';

const MDE_ID = /^[0-9a-f]{40}$/i;
const SHA1 = /^[0-9a-f]{40}$/i;

const MACHINE_ARG = {
  name: 'machine',
  type: 'string',
  required: true,
  positional: true,
  description: 'Computer name, Defender machine id or Entra device id (aadDeviceId).',
} as const satisfies ArgSpec;

const COMMENT_ARG = {
  name: 'comment',
  type: 'string',
  required: true,
  description: 'Why (recorded in the Defender action center).',
} as const satisfies ArgSpec;

export function shapeMachine(machine: Machine) {
  return compact({
    id: machine.id,
    computerDnsName: machine.computerDnsName ?? null,
    osPlatform: machine.osPlatform ?? null,
    healthStatus: machine.healthStatus ?? null,
    riskScore: machine.riskScore ?? null,
    exposureLevel: machine.exposureLevel ?? null,
    lastSeen: machine.lastSeen ?? null,
    daysSinceLastSeen: daysSince(machine.lastSeen),
    lastIpAddress: machine.lastIpAddress ?? null,
    aadDeviceId: machine.aadDeviceId ?? null,
    machineTags: machine.machineTags ?? [],
    onboardingStatus: machine.onboardingStatus ?? null,
  });
}

function shapeAction(action: MachineAction) {
  return compact({
    id: action.id,
    type: action.type ?? null,
    scope: action.scope ?? undefined,
    status: action.status ?? null,
    machine: action.computerDnsName ?? action.machineId ?? null,
    requestor: action.requestor ?? null,
    comment: action.requestorComment ?? null,
    file: action.relatedFileInfo?.fileIdentifier ?? undefined,
    created: action.creationDateTimeUtc ?? null,
    lastUpdated: action.lastUpdateDateTimeUtc ?? undefined,
  });
}

function describeMachine(machine: Machine) {
  return {
    id: machine.id,
    computerDnsName: machine.computerDnsName ?? null,
    osPlatform: machine.osPlatform ?? null,
    lastSeen: machine.lastSeen ?? null,
    aadDeviceId: machine.aadDeviceId ?? null,
  };
}

/**
 * One Defender machine from what a person would say: MDE machine id (40 hex), Entra device id (aadDeviceId),
 * or computer name (case-insensitive; short name or FQDN). Several machines sharing a name are AMBIGUOUS.
 */
export async function resolveMachine(ctx: Context, ref: string): Promise<Machine> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A machine is required');
  const find = (filter: string) => ctx.defender.list<Machine>('/machines', { query: { $filter: filter }, limit: 25 });
  if (MDE_ID.test(value)) {
    const { items } = await find(`id eq ${odata(value.toLowerCase())}`);
    if (!items.length) throw new CliError('NOT_FOUND', `No Defender machine with id ${value}`);
    return items[0] as Machine;
  }
  if (isGuid(value)) {
    const { items } = await find(`aadDeviceId eq ${odata(value)}`);
    if (!items.length) {
      throw new CliError('NOT_FOUND', `No Defender machine with Entra device id ${value}`, {
        hint: 'The device may not be onboarded to Defender for Endpoint. Search by name with `tenantctl defender machines list --search`.',
      });
    }
    return pickOne(items, value, 'machine', (m) => [m.aadDeviceId], describeMachine);
  }
  const { items } = await find(`startswith(computerDnsName,${odata(value.toLowerCase())})`);
  return pickOne(
    items,
    value,
    'machine',
    (m) => [m.computerDnsName, m.computerDnsName?.split('.')[0]],
    describeMachine,
  );
}

const machineName = (machine: Machine) => machine.computerDnsName ?? machine.id;

async function recentActions(ctx: Context, machineId: string) {
  const { items } = await ctx.defender.list<MachineAction>('/machineactions', {
    query: { $filter: `machineId eq ${odata(machineId)}` },
    limit: 100,
  });
  return sortActions(items).slice(0, 10).map(shapeAction);
}

async function openAlerts(ctx: Context, machineId: string): Promise<number> {
  const count = await ctx.defender.get<string | number>(`/machines/${seg(machineId)}/alerts/$count`, {
    query: { $filter: "status ne 'Resolved'" },
  });
  return Number(count);
}

function sortActions(actions: MachineAction[]): MachineAction[] {
  return [...actions].sort((a, b) => (b.creationDateTimeUtc ?? '').localeCompare(a.creationDateTimeUtc ?? ''));
}

type Args = Record<string, unknown>;

interface ActionSpec {
  name: string;
  summary: string;
  description: string;
  examples: string[];
  risk: Exclude<Risk, 'read'>;
  /** Action segment after `/machines/{id}/`, as in the MDE docs. */
  endpoint: string;
  permission: string;
  /** osPlatform test; machines whose platform is known and fails it are refused before anything is sent. */
  platform?: { test: RegExp; label: string };
  args?: ArgSpec[];
  body: (args: Args) => Record<string, unknown>;
  describe: (machine: Machine, args: Args) => string;
  validate?: (machine: Machine, args: Args) => void;
}

function checkPlatform(spec: ActionSpec, machine: Machine): void {
  if (!spec.platform || !machine.osPlatform || spec.platform.test.test(machine.osPlatform)) return;
  throw new CliError(
    'USAGE',
    `"machines ${spec.name}" is not supported on ${machine.osPlatform} (${machineName(machine)})`,
    { hint: `Defender for Endpoint supports it on ${spec.platform.label} only.` },
  );
}

async function target(ctx: Context, spec: ActionSpec, args: Args): Promise<Machine> {
  const machine = await resolveMachine(ctx, String(args.machine));
  checkPlatform(spec, machine);
  spec.validate?.(machine, args);
  return machine;
}

/** Response actions share one shape: resolve the machine, refuse unsupported platforms, preview, POST. */
function machineAction(spec: ActionSpec): AnyCommand {
  const path = (machine: Machine) => `/machines/${seg(machine.id)}/${spec.endpoint}`;
  return defineCommand({
    domain: 'defender',
    name: `machines ${spec.name}`,
    summary: spec.summary,
    description: `${spec.description} ${LICENSE}`,
    examples: spec.examples,
    args: [MACHINE_ARG, ...(spec.args ?? []), COMMENT_ARG] as readonly ArgSpec[],
    risk: spec.risk,
    permissions: { defender: [READ_MACHINES, spec.permission], roles: RESPONSE_ROLES },
    async preview(ctx, args) {
      const machine = await target(ctx, spec, args);
      return {
        action: spec.describe(machine, args),
        machine: shapeMachine(machine),
        request: { method: 'POST', path: path(machine), body: spec.body(args) },
      };
    },
    async run(ctx, args) {
      const machine = await target(ctx, spec, args);
      const action = await ctx.defender.post<MachineAction>(path(machine), spec.body(args));
      if (ctx.full) return action;
      return {
        action: spec.describe(machine, args),
        machineAction: compact({
          id: action?.id,
          type: action?.type ?? undefined,
          status: action?.status ?? null,
          created: action?.creationDateTimeUtc ?? undefined,
        }),
        machine: { id: machine.id, computerDnsName: machine.computerDnsName ?? null },
        next: `tenantctl defender machines actions --machine ${machineName(machine)}`,
      };
    },
  });
}

const comment = (args: Args) => String(args.comment);

const actionCommands: AnyCommand[] = [
  machineAction({
    name: 'isolate',
    summary: 'Isolate a machine from the network (Defender keeps its connection to the service)',
    description:
      'Use when a device is compromised or actively attacked. Full isolation cuts all network traffic except ' +
      'to Defender for Endpoint; --selective (Windows only) still lets Outlook, Teams and Skype communicate. ' +
      'Machines behind a full-tunnel VPN can lose contact with Defender once isolated. Undo with `machines ' +
      'release`. Supported: Windows 10 1703+/11 (selective 1709+), Windows Server, Linux.',
    examples: ['tenantctl defender machines isolate LAPTOP-7 --comment "Ransomware alert 1234"'],
    risk: 'destructive',
    endpoint: 'isolate',
    permission: 'Machine.Isolate',
    platform: { test: /^(windows|linux)/i, label: 'Windows and Linux' },
    args: [{ name: 'selective', type: 'boolean', description: 'Selective isolation (Outlook/Teams keep working).' }],
    body: (args) => ({ Comment: comment(args), IsolationType: args.selective ? 'Selective' : 'Full' }),
    describe: (m, args) =>
      `Isolate ${machineName(m)} from the network (${args.selective ? 'Selective' : 'Full'} isolation)`,
    validate: (m, args) => {
      if (args.selective && m.osPlatform && !/^windows/i.test(m.osPlatform)) {
        throw new CliError('USAGE', `Selective isolation is Windows only; ${machineName(m)} runs ${m.osPlatform}`, {
          hint: 'Drop --selective for full isolation.',
        });
      }
    },
  }),
  machineAction({
    name: 'release',
    summary: 'Release a machine from network isolation',
    description: 'Undoes `machines isolate`. The machine reconnects once it receives the request.',
    examples: ['tenantctl defender machines release LAPTOP-7 --comment "Remediated, cleared by SOC"'],
    risk: 'write',
    endpoint: 'unisolate',
    permission: 'Machine.Isolate',
    body: (args) => ({ Comment: comment(args) }),
    describe: (m) => `Release ${machineName(m)} from network isolation`,
  }),
  machineAction({
    name: 'scan',
    summary: 'Run a Microsoft Defender Antivirus quick or full scan on a machine',
    description:
      'Quick scan by default; --full-scan scans every file and can take hours. Works with Defender Antivirus in ' +
      'passive mode too. Supported: Windows 10 1709+/11, Windows Server, Linux, macOS.',
    examples: ['tenantctl defender machines scan LAPTOP-7 --full-scan --comment "Alert 1234 follow-up"'],
    risk: 'write',
    endpoint: 'runAntiVirusScan',
    permission: 'Machine.Scan',
    platform: { test: /^(windows|linux|mac)/i, label: 'Windows, Linux and macOS' },
    args: [{ name: 'fullScan', type: 'boolean', description: 'Full scan instead of a quick scan.' }],
    body: (args) => ({ Comment: comment(args), ScanType: args.fullScan ? 'Full' : 'Quick' }),
    describe: (m, args) => `Run a ${args.fullScan ? 'full' : 'quick'} antivirus scan on ${machineName(m)}`,
  }),
  machineAction({
    name: 'collect-package',
    summary: 'Collect an investigation package (logs, processes, network state) from a machine',
    description:
      'Use for forensics. When the action status is Succeeded, get a download link with `machines package-url ' +
      '<actionId>`.',
    examples: ['tenantctl defender machines collect-package LAPTOP-7 --comment "Incident 42 forensics"'],
    risk: 'write',
    endpoint: 'collectInvestigationPackage',
    permission: 'Machine.CollectForensics',
    body: (args) => ({ Comment: comment(args) }),
    describe: (m) => `Collect an investigation package from ${machineName(m)}`,
  }),
  machineAction({
    name: 'restrict',
    summary: 'Restrict app execution on a machine to Microsoft-signed binaries',
    description:
      'Containment for an active threat: only files signed by Microsoft can run until `machines unrestrict`. ' +
      'Needs Microsoft Defender Antivirus. Supported: Windows 10 1709+/11.',
    examples: ['tenantctl defender machines restrict LAPTOP-7 --comment "Unknown binary spreading"'],
    risk: 'destructive',
    endpoint: 'restrictCodeExecution',
    permission: 'Machine.RestrictExecution',
    platform: { test: /^windows/i, label: 'Windows' },
    body: (args) => ({ Comment: comment(args) }),
    describe: (m) => `Restrict ${machineName(m)} to running Microsoft-signed applications only`,
  }),
  machineAction({
    name: 'unrestrict',
    summary: 'Remove the app execution restriction from a machine',
    description: 'Undoes `machines restrict`.',
    examples: ['tenantctl defender machines unrestrict LAPTOP-7 --comment "Threat removed"'],
    risk: 'write',
    endpoint: 'unrestrictCodeExecution',
    permission: 'Machine.RestrictExecution',
    platform: { test: /^windows/i, label: 'Windows' },
    body: (args) => ({ Comment: comment(args) }),
    describe: (m) => `Allow all applications to run on ${machineName(m)} again`,
  }),
  machineAction({
    name: 'stop-quarantine-file',
    summary: 'Stop a running file on a machine and quarantine it (by SHA-1)',
    description:
      'Kills processes of the file and moves it to quarantine. Refused by Defender for files signed by Microsoft ' +
      'or trusted publishers. Needs Defender Antivirus (passive mode is enough). Supported: Windows 10 1703+/11. ' +
      'Find the SHA-1 in alert evidence or `tenantctl defender hunt`.',
    examples: [
      'tenantctl defender machines stop-quarantine-file LAPTOP-7 --sha1 a0c659857ccbe457fdaf5fe21d54efdcbf6f6508 --comment "Malware from alert 99"',
    ],
    risk: 'destructive',
    endpoint: 'StopAndQuarantineFile',
    permission: 'Machine.StopAndQuarantine',
    platform: { test: /^windows/i, label: 'Windows' },
    args: [{ name: 'sha1', type: 'string', required: true, description: 'SHA-1 of the file (40 hex characters).' }],
    body: (args) => ({ Comment: comment(args), Sha1: String(args.sha1).trim().toLowerCase() }),
    describe: (m, args) => `Stop and quarantine file ${String(args.sha1).trim().toLowerCase()} on ${machineName(m)}`,
    validate: (_m, args) => {
      if (!SHA1.test(String(args.sha1).trim())) {
        throw new CliError('USAGE', `"${String(args.sha1)}" is not a SHA-1 hash`, {
          hint: 'Pass the 40-character hex SHA-1 of the file (not SHA-256).',
        });
      }
    },
  }),
];

const RISK_CHOICES = ['high', 'medium', 'low'] as const;
const HEALTH_CHOICES = ['active', 'inactive'] as const;
const ACTION_STATUSES = ['Pending', 'InProgress', 'Succeeded', 'Failed', 'TimeOut', 'Cancelled'] as const;
/** MDE enum spellings for the lower-case CLI choices. */
const RISK_SCORE: Record<(typeof RISK_CHOICES)[number], string> = { high: 'High', medium: 'Medium', low: 'Low' };
const HEALTH_STATUS: Record<(typeof HEALTH_CHOICES)[number], string> = { active: 'Active', inactive: 'Inactive' };

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'defender',
    name: 'machines list',
    summary: 'List Defender for Endpoint machines with health, risk, exposure and last seen',
    description:
      'Device inventory as Defender for Endpoint sees it. --search matches the start of the computer name; ' +
      '--risk filters riskScore (High/Medium/Low); --health active|inactive filters healthStatus; --os takes an ' +
      'osPlatform value (Windows11, Windows10, WindowsServer2022, Linux, macOS, iOS, Android...); --stale-days ' +
      `lists machines not seen for N days. ${LICENSE}`,
    examples: [
      'tenantctl defender machines list --risk high',
      'tenantctl defender machines list --health inactive --stale-days 30 --all',
      'tenantctl defender machines list --search laptop-7',
    ],
    args: [
      { name: 'search', type: 'string', description: 'Computer name prefix.' },
      { name: 'risk', type: 'string', choices: RISK_CHOICES, description: 'Only this risk score.' },
      { name: 'health', type: 'string', choices: HEALTH_CHOICES, description: 'Only this sensor health status.' },
      { name: 'os', type: 'string', description: 'osPlatform, e.g. Windows11, WindowsServer2022, Linux, macOS.' },
      { name: 'staleDays', type: 'number', description: 'Only machines not seen in the last N days.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { defender: [READ_MACHINES], roles: READ_ROLES },
    async run(ctx, args) {
      const clauses = [
        args.search ? `startswith(computerDnsName,${odata(args.search.trim().toLowerCase())})` : undefined,
        args.risk ? `riskScore eq ${odata(RISK_SCORE[args.risk])}` : undefined,
        args.health ? `healthStatus eq ${odata(HEALTH_STATUS[args.health])}` : undefined,
        args.os ? `osPlatform eq ${odata(args.os.trim())}` : undefined,
        typeof args.staleDays === 'number' ? `lastSeen lt ${daysAgo(args.staleDays)}` : undefined,
      ].filter((c): c is string => Boolean(c));
      const result = await ctx.defender.list<Machine>('/machines', {
        query: { $filter: clauses.length ? clauses.join(' and ') : undefined },
        limit: listLimit(args),
      });
      if (ctx.full) return result;
      return { items: result.items.map(shapeMachine), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'machines get',
    summary: 'One Defender machine: health, risk, open alerts and recent response actions',
    description:
      'Accepts the computer name, Defender machine id or Entra device id. openAlerts counts alerts not yet ' +
      'resolved (needs Alert.ReadWrite.All per the MDE docs); recentActions lists the last 10 response actions ' +
      '(needs Machine.ReadWrite.All). A section the token lacks degrades to { error }. ' +
      `Next: \`defender alerts for-device <name>\` or a response action. ${LICENSE}`,
    examples: [
      'tenantctl defender machines get LAPTOP-7',
      'tenantctl defender machines get 7f3a1d8e-0000-4c1b-9a1e-1234567890ab',
    ],
    args: [MACHINE_ARG],
    risk: 'read',
    permissions: { defender: [READ_MACHINES, 'Machine.ReadWrite.All', 'Alert.ReadWrite.All'], roles: READ_ROLES },
    async run(ctx, args) {
      const machine = await resolveMachine(ctx, args.machine);
      const extra = await settle({
        openAlerts: openAlerts(ctx, machine.id),
        recentActions: recentActions(ctx, machine.id),
      });
      if (ctx.full) return { ...machine, ...extra };
      return {
        ...shapeMachine(machine),
        version: machine.version ?? null,
        firstSeen: machine.firstSeen ?? null,
        lastExternalIpAddress: machine.lastExternalIpAddress ?? null,
        machineGroup: machine.rbacGroupName ?? null,
        ...extra,
      };
    },
  }),

  ...actionCommands,

  defineCommand({
    domain: 'defender',
    name: 'machines package-url',
    summary: 'Download link for a collected investigation package',
    description:
      'Pass the machine action id returned by `machines collect-package`. The link is short-lived: download it ' +
      'right away. Answers NOT_FOUND while the collection is still running; check with `machines actions`. ' +
      `Defender limits this call to 2 per minute. ${LICENSE}`,
    examples: ['tenantctl defender machines package-url 69dc3630-1ccc-4342-acf3-35286eec741d'],
    args: [{ name: 'actionId', type: 'string', required: true, positional: true, description: 'Machine action id.' }],
    risk: 'read',
    permissions: { defender: ['Machine.ReadWrite.All', 'Machine.CollectForensics'], roles: RESPONSE_ROLES },
    async run(ctx, args) {
      const response = await ctx.defender.get<{ value?: string }>(
        `/machineactions/${seg(args.actionId.trim())}/GetPackageUri`,
      );
      if (ctx.full) return response;
      const url = (response?.value ?? '').replace(/^"+|"+$/g, '');
      return { actionId: args.actionId.trim(), url, note: 'Short-lived link; download it now.' };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'machines actions',
    summary: 'Response actions (isolate, scan, collect...) and their status, newest first',
    description:
      'Use to follow up on an action id returned by a `machines` response command, or to see what was done to ' +
      'a machine. status is Pending, InProgress, Succeeded, Failed, TimeOut or Cancelled. ' +
      `Needs Machine.ReadWrite.All per the MDE docs. ${LICENSE}`,
    examples: [
      'tenantctl defender machines actions --machine LAPTOP-7',
      'tenantctl defender machines actions --status Pending',
    ],
    args: [
      { name: 'machine', type: 'string', description: 'Only actions on this machine (name or id).' },
      { name: 'status', type: 'string', choices: ACTION_STATUSES, description: 'Only actions in this status.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { defender: ['Machine.ReadWrite.All'], roles: READ_ROLES },
    async run(ctx, args) {
      const machine = args.machine ? await resolveMachine(ctx, args.machine) : undefined;
      const clauses = [
        machine ? `machineId eq ${odata(machine.id)}` : undefined,
        args.status ? `status eq ${odata(args.status)}` : undefined,
      ].filter((c): c is string => Boolean(c));
      const result = await ctx.defender.list<MachineAction>('/machineactions', {
        query: { $filter: clauses.length ? clauses.join(' and ') : undefined },
        limit: listLimit(args),
      });
      const items = sortActions(result.items);
      if (ctx.full) return { items, truncated: result.truncated };
      return { items: items.map(shapeAction), truncated: result.truncated };
    },
  }),
];
