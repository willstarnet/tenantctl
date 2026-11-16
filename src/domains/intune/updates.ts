import { CliError } from '../../core/errors.ts';
import { type ApiVersion, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact } from '../../core/util.ts';
import { type Assignment, shapeAssignmentLists } from './assignments.ts';
import { CONFIG_PATH, INTUNE_READ_ROLES, odataType, pickByName, READ_CONFIG, READ_GROUPS } from './common.ts';

const ALL = Number.POSITIVE_INFINITY;
const RING_TYPE = 'windowsUpdateForBusinessConfiguration';

type UpdateKind = 'ring' | 'feature' | 'quality' | 'driver';

const KINDS: Record<UpdateKind, { path: string; version: ApiVersion }> = {
  ring: { path: CONFIG_PATH, version: 'v1.0' },
  feature: { path: '/deviceManagement/windowsFeatureUpdateProfiles', version: 'beta' },
  quality: { path: '/deviceManagement/windowsQualityUpdateProfiles', version: 'beta' },
  driver: { path: '/deviceManagement/windowsDriverUpdateProfiles', version: 'beta' },
};
const KIND_NAMES = Object.keys(KINDS) as UpdateKind[];

interface UpdatePolicy {
  id: string;
  displayName?: string | null;
  description?: string | null;
  lastModifiedDateTime?: string | null;
  [key: string]: unknown;
}

interface Kinded {
  kind: UpdateKind;
  id: string;
  name: string;
  policy: UpdatePolicy;
}

/** Autopatch creates its rings and profiles itself and names them "Windows Autopatch ..."; edits get overwritten. */
export function managedByAutopatch(policy: UpdatePolicy): boolean {
  return (
    /^windows autopatch/i.test(policy.displayName ?? '') ||
    /created by windows autopatch/i.test(policy.description ?? '')
  );
}

/** Intune reports "not paused" as 0001-01-01; treat it as null. */
const realDate = (iso: unknown) => (typeof iso === 'string' && !iso.startsWith('0001-') ? iso : null);

/**
 * Update rings are deviceConfigurations of one type. `isof()` works in practice but is not documented for this
 * collection, so on a 400 the type is matched client-side.
 */
async function listRings(ctx: Context): Promise<UpdatePolicy[]> {
  try {
    const { items } = await ctx.graph.list<UpdatePolicy>(CONFIG_PATH, {
      query: { $filter: `isof('microsoft.graph.${RING_TYPE}')` },
      limit: ALL,
    });
    return items;
  } catch (error) {
    if (!(error instanceof CliError && error.status === 400)) throw error;
    ctx.log('Graph rejected the isof() filter; matching update rings client-side.');
    const { items } = await ctx.graph.list<UpdatePolicy>(CONFIG_PATH, { limit: ALL });
    return items.filter((p) => odataType(p) === RING_TYPE);
  }
}

async function listKind(ctx: Context, kind: UpdateKind): Promise<Kinded[]> {
  const items =
    kind === 'ring'
      ? await listRings(ctx)
      : (await ctx.graph.list<UpdatePolicy>(KINDS[kind].path, { version: 'beta', limit: ALL })).items;
  return items.map((policy) => ({ kind, id: policy.id, name: policy.displayName ?? policy.id, policy }));
}

/** Policies of each kind; an unreadable kind lands in `unavailable` instead of failing, unless all of them fail. */
async function listUpdatePolicies(ctx: Context, kinds: UpdateKind[]) {
  const loaded = await Promise.allSettled(kinds.map((kind) => listKind(ctx, kind)));
  if (loaded.every((r) => r.status === 'rejected')) throw (loaded[0] as PromiseRejectedResult).reason;
  const policies: Kinded[] = [];
  const unavailable: { section: UpdateKind; error: { code: string; message: string } }[] = [];
  loaded.forEach((result, index) => {
    const kind = kinds[index] as UpdateKind;
    if (result.status === 'fulfilled') {
      policies.push(...result.value);
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
  return { policies, unavailable };
}

/** Assignments of each policy through `$batch` (one per API version); null where the lookup failed. */
async function loadAssignments(ctx: Context, policies: Kinded[]): Promise<(Assignment[] | null)[]> {
  const out: (Assignment[] | null)[] = policies.map(() => null);
  for (const version of ['v1.0', 'beta'] as const) {
    const indexes = policies.flatMap((p, i) => (KINDS[p.kind].version === version ? [i] : []));
    if (!indexes.length) continue;
    const results = await ctx.graph.batch<{ value?: Assignment[] }>(
      indexes.map((i) => {
        const p = policies[i] as Kinded;
        return { path: `${KINDS[p.kind].path}/${seg(p.id)}/assignments` };
      }),
      version,
    );
    indexes.forEach((policyIndex, k) => {
      const result = results[k];
      out[policyIndex] = result && !result.error ? (result.body?.value ?? []) : null;
    });
  }
  return out;
}

const pick = (policy: UpdatePolicy, key: string) => (policy[key] === undefined ? null : policy[key]);

function shapeRing(p: UpdatePolicy, detail: boolean) {
  return compact({
    quality: {
      deferralDays: pick(p, 'qualityUpdatesDeferralPeriodInDays'),
      deadlineDays: pick(p, 'deadlineForQualityUpdatesInDays'),
      paused: pick(p, 'qualityUpdatesPaused'),
      pauseExpiryDateTime: realDate(p.qualityUpdatesPauseExpiryDateTime),
    },
    feature: {
      deferralDays: pick(p, 'featureUpdatesDeferralPeriodInDays'),
      deadlineDays: pick(p, 'deadlineForFeatureUpdatesInDays'),
      paused: pick(p, 'featureUpdatesPaused'),
      pauseExpiryDateTime: realDate(p.featureUpdatesPauseExpiryDateTime),
      rollbackWindowDays: pick(p, 'featureUpdatesRollbackWindowInDays'),
    },
    deadlineGracePeriodDays: pick(p, 'deadlineGracePeriodInDays'),
    postponeRebootUntilAfterDeadline: pick(p, 'postponeRebootUntilAfterDeadline'),
    automaticUpdateMode: pick(p, 'automaticUpdateMode'),
    allowWindows11Upgrade: pick(p, 'allowWindows11Upgrade'),
    driversExcluded: pick(p, 'driversExcluded'),
    microsoftUpdateServiceAllowed: detail ? pick(p, 'microsoftUpdateServiceAllowed') : undefined,
    installationSchedule: detail ? pick(p, 'installationSchedule') : undefined,
    businessReadyUpdatesOnly: detail ? pick(p, 'businessReadyUpdatesOnly') : undefined,
    userPauseAccess: detail ? pick(p, 'userPauseAccess') : undefined,
    userWindowsUpdateScanAccess: detail ? pick(p, 'userWindowsUpdateScanAccess') : undefined,
    updateNotificationLevel: detail ? pick(p, 'updateNotificationLevel') : undefined,
    skipChecksBeforeRestart: detail ? pick(p, 'skipChecksBeforeRestart') : undefined,
    engagedRestartDeadlineDays: detail ? pick(p, 'engagedRestartDeadlineInDays') : undefined,
  });
}

function shapeSettings(item: Kinded, detail: boolean) {
  const p = item.policy;
  switch (item.kind) {
    case 'ring':
      return shapeRing(p, detail);
    case 'feature': {
      const rollout = (p.rolloutSettings ?? {}) as Record<string, unknown>;
      const start = rollout.offerStartDateTimeInUTC ?? null;
      const end = rollout.offerEndDateTimeInUTC ?? null;
      const interval = rollout.offerIntervalInDays ?? null;
      return compact({
        targetVersion: pick(p, 'featureUpdateVersion'),
        rollout:
          start || end || interval
            ? { offerStartDateTime: start, offerEndDateTime: end, offerIntervalDays: interval }
            : 'immediate',
        endOfSupportDate: pick(p, 'endOfSupportDate'),
        installLatestWindows10OnWindows11IneligibleDevice: pick(p, 'installLatestWindows10OnWindows11IneligibleDevice'),
        installFeatureUpdatesOptional: detail ? pick(p, 'installFeatureUpdatesOptional') : undefined,
      });
    }
    case 'quality': {
      const expedite = (p.expeditedUpdateSettings ?? null) as Record<string, unknown> | null;
      return compact({
        release: pick(p, 'releaseDateDisplayName'),
        content: pick(p, 'deployableContentDisplayName'),
        expedite: expedite
          ? {
              qualityUpdateRelease: expedite.qualityUpdateRelease ?? null,
              daysUntilForcedReboot: expedite.daysUntilForcedReboot ?? null,
            }
          : null,
      });
    }
    case 'driver':
      return compact({
        approvalType: pick(p, 'approvalType'),
        deploymentDeferralDays: pick(p, 'deploymentDeferralInDays'),
        devicesReporting: pick(p, 'deviceReporting'),
        newUpdates: pick(p, 'newUpdates'),
        inventorySyncStatus: detail ? pick(p, 'inventorySyncStatus') : undefined,
      });
  }
}

async function shapePolicies(ctx: Context, policies: Kinded[], detail: boolean) {
  const lists = await loadAssignments(ctx, policies);
  const shaped = await shapeAssignmentLists(
    ctx,
    lists.map((list) => list ?? []),
  );
  return policies.map((item, index) => {
    const assignments = lists[index] === null ? null : shaped[index];
    if (ctx.full) return { kind: item.kind, ...item.policy, assignments: lists[index] };
    return compact({
      kind: item.kind,
      id: item.id,
      name: item.name,
      managedByAutopatch: managedByAutopatch(item.policy),
      description: detail ? item.policy.description || undefined : undefined,
      ...shapeSettings(item, detail),
      assignments,
      lastModifiedDateTime: item.policy.lastModifiedDateTime ?? null,
    });
  });
}

const PERMISSIONS = {
  delegated: [READ_CONFIG, READ_GROUPS],
  application: [READ_CONFIG, READ_GROUPS],
  roles: INTUNE_READ_ROLES,
};

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'updates list',
    summary: 'List Windows Update rings and feature, quality and driver update profiles with assignments',
    description:
      'kind is ring (Windows Update for Business ring: deferrals, deadlines, grace period, pause state), feature ' +
      '(target version and rollout), quality (expedited update) or driver (approval type). Assignments show group ' +
      'names. managedByAutopatch=true marks rings and profiles Windows Autopatch generated; change those in ' +
      'Autopatch, not by editing the policy, which Autopatch overwrites. A kind the caller cannot read is listed ' +
      'in meta.unavailable. (Rings use v1.0; feature, quality and driver profiles are beta API only.)',
    examples: ['tenantctl intune updates list', 'tenantctl intune updates list --kind ring'],
    args: [
      { name: 'kind', type: 'string', choices: KIND_NAMES, description: 'Only policies of this kind.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: PERMISSIONS,
    async run(ctx, args) {
      const kinds = args.kind ? [args.kind as UpdateKind] : KIND_NAMES;
      const { policies, unavailable } = await listUpdatePolicies(ctx, kinds);
      const page = policies.slice(0, listLimit(args));
      return {
        items: await shapePolicies(ctx, page, false),
        truncated: policies.length > page.length,
        ...(unavailable.length ? { unavailable } : {}),
      } satisfies ListResult<unknown>;
    },
  }),
  defineCommand({
    domain: 'intune',
    name: 'updates get',
    summary: 'Show one Windows Update ring or update profile with its settings and assignments',
    description:
      'Accepts the name (or part of it) or id of an update ring or a feature, quality or driver update profile. ' +
      'managedByAutopatch=true means Windows Autopatch owns it: change it in Autopatch, not here or in the ' +
      'policy. Per-device feature update state is only available through the asynchronous Intune reports export ' +
      '(FeatureUpdateDeviceState), which this command does not run. (beta API for feature, quality and driver ' +
      'profiles.)',
    examples: ['tenantctl intune updates get "Broad ring"', 'tenantctl intune updates get "Windows 11 26H2"'],
    args: [
      { name: 'policy', type: 'string', required: true, positional: true, description: 'Ring or profile name or id.' },
    ],
    risk: 'read',
    permissions: PERMISSIONS,
    async run(ctx, args) {
      const { policies, unavailable } = await listUpdatePolicies(ctx, KIND_NAMES);
      const policy = pickByName(policies, args.policy, 'update policy', (p) => p.name);
      const [shaped] = await shapePolicies(ctx, [policy], true);
      return unavailable.length ? { policy: shaped, unavailable } : { policy: shaped };
    },
  }),
];
