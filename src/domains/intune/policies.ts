import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact, daysSince, settle } from '../../core/util.ts';
import {
  attempt,
  DEVICES_PATH,
  INTUNE_READ_ROLES,
  listDevices,
  mapLimit,
  odataType,
  pickByName,
  READ_CONFIG,
  READ_DEVICES,
  scan,
  shapeManagedDevice,
} from './common.ts';
import { complianceSummary, type PolicyState, shapePolicyStates } from './devices.ts';

export const READ_APPS = 'DeviceManagementApps.Read.All';
const READ_SERVICE_CONFIG = 'DeviceManagementServiceConfig.Read.All';
export const COMPLIANCE_PATH = '/deviceManagement/deviceCompliancePolicies';
export const CONFIG_PATH = '/deviceManagement/deviceConfigurations';
export const SETTINGS_CATALOG_PATH = '/deviceManagement/configurationPolicies';
export const APPS_PATH = '/deviceAppManagement/mobileApps';
const CONCURRENCY = 5;

const FAILING_ONLY_ARG = {
  name: 'failingOnly',
  type: 'boolean',
  description: 'Only devices in nonCompliant, error or conflict state.',
} as const;

const FAILING_STATUS: Record<string, true> = { nonCompliant: true, error: true, conflict: true };

export interface Policy {
  id: string;
  displayName?: string | null;
  name?: string | null;
  description?: string | null;
  lastModifiedDateTime?: string | null;
  assignments?: unknown[];
  [key: string]: unknown;
}

interface DeviceStatus {
  id: string;
  deviceDisplayName?: string | null;
  userPrincipalName?: string | null;
  userName?: string | null;
  deviceModel?: string | null;
  status?: string | null;
  lastReportedDateTime?: string | null;
  complianceGracePeriodExpirationDateTime?: string | null;
}

interface Overview {
  successCount?: number;
  failedCount?: number;
  errorCount?: number;
  conflictCount?: number;
  pendingCount?: number;
  notApplicableCount?: number;
  lastUpdateDateTime?: string;
}

function shapeOverview(overview: Overview | { error: unknown }) {
  if ('error' in overview) return overview;
  return compact({
    compliant: overview.successCount,
    nonCompliant: overview.failedCount,
    error: overview.errorCount,
    conflict: overview.conflictCount,
    pending: overview.pendingCount,
    notApplicable: overview.notApplicableCount,
    lastUpdateDateTime: overview.lastUpdateDateTime,
  });
}

function shapeDeviceStatus(status: DeviceStatus) {
  return compact({
    device: status.deviceDisplayName ?? null,
    user: status.userPrincipalName || status.userName || null,
    model: status.deviceModel || undefined,
    status: status.status ?? null,
    lastReportedDateTime: status.lastReportedDateTime ?? null,
    graceExpires:
      status.complianceGracePeriodExpirationDateTime &&
      !status.complianceGracePeriodExpirationDateTime.startsWith('9999')
        ? status.complianceGracePeriodExpirationDateTime
        : undefined,
  });
}

/** Device statuses for a compliance policy or configuration profile; status filtering is client-side. */
async function policyDeviceStatuses(
  ctx: Context,
  path: string,
  failingOnly: boolean,
  limit: number,
): Promise<ListResult<DeviceStatus>> {
  return failingOnly
    ? scan<DeviceStatus>(ctx, path, {}, (s) => Boolean(FAILING_STATUS[s.status ?? '']), limit)
    : ctx.graph.list<DeviceStatus>(path, { limit });
}

export async function resolveCatalogItem(ctx: Context, path: string, ref: string, kind: string): Promise<Policy> {
  const { items } = await ctx.graph.list<Policy>(path, {
    query: { $select: 'id,displayName' },
    limit: Number.POSITIVE_INFINITY,
  });
  return pickByName(items, ref, kind, (p) => p.displayName);
}

function tally<T>(items: T[], key: (item: T) => string): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const item of items) counts[key(item)] = (counts[key(item)] ?? 0) + 1;
  return counts;
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'compliance list',
    summary: 'List compliance policies with platform, assignment count and device status counts',
    description:
      'Use to see which compliance policies exist and which ones devices are failing. assignments = 0 means the ' +
      'policy applies to nobody. devices.nonCompliant/error point at the policy to investigate with ' +
      '`compliance status <policy> --failing-only`.',
    examples: ['m365 intune compliance list'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: [READ_CONFIG], application: [READ_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const result = await ctx.graph.list<Policy>(COMPLIANCE_PATH, {
        query: { $expand: 'assignments' },
        limit: listLimit(args),
      });
      const overviews = await mapLimit(result.items, CONCURRENCY, (policy) =>
        attempt(ctx.graph.get<Overview>(`${COMPLIANCE_PATH}/${seg(policy.id)}/deviceStatusOverview`)),
      );
      if (ctx.full) {
        return {
          items: result.items.map((p, i) => ({ ...p, deviceStatusOverview: overviews[i] })),
          truncated: result.truncated,
        };
      }
      return {
        items: result.items.map((policy, index) =>
          compact({
            id: policy.id,
            name: policy.displayName ?? null,
            platform: odataType(policy),
            assignments: policy.assignments?.length ?? 0,
            lastModifiedDateTime: policy.lastModifiedDateTime ?? null,
            devices: shapeOverview(overviews[index] as Overview),
          }),
        ),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'compliance status',
    summary: 'Per-device results for one compliance policy',
    description:
      'Use to find every device failing a given policy. Accepts the policy name (or part of it) or id. Status ' +
      'is compliant, nonCompliant, error, conflict, notApplicable or unknown; graceExpires is when a ' +
      'non-compliant device gets blocked by Conditional Access. Next: `devices compliance <device>` for the ' +
      'failing settings.',
    examples: ['m365 intune compliance status "Windows 11 baseline" --failing-only'],
    args: [
      { name: 'policy', type: 'string', required: true, positional: true, description: 'Policy name or id.' },
      FAILING_ONLY_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_CONFIG], application: [READ_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const policy = await resolveCatalogItem(ctx, COMPLIANCE_PATH, args.policy, 'compliance policy');
      const result = await policyDeviceStatuses(
        ctx,
        `${COMPLIANCE_PATH}/${seg(policy.id)}/deviceStatuses`,
        args.failingOnly,
        listLimit(args),
      );
      if (ctx.full) return { policy, ...result };
      return {
        policy: { id: policy.id, name: policy.displayName ?? null },
        items: result.items.map(shapeDeviceStatus),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'config list',
    summary: 'List configuration profiles: classic profiles and settings catalog policies together',
    description:
      'Shows every device configuration profile with `kind`: "profile" (classic templates, v1.0) or ' +
      '"settingsCatalog" (settings catalog / endpoint security, beta API). assigned=false means it applies to ' +
      'nobody. If the settings catalog call fails, the classic profiles are still returned with a warning.',
    examples: ['m365 intune config list', 'm365 intune config list --all'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: [READ_CONFIG], application: [READ_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const limit = listLimit(args);
      const [profiles, catalogResult] = await Promise.all([
        ctx.graph.list<Policy>(CONFIG_PATH, { query: { $expand: 'assignments' }, limit }),
        attempt(
          ctx.graph.list<Policy>(SETTINGS_CATALOG_PATH, {
            version: 'beta',
            query: {
              $select: 'id,name,description,platforms,technologies,isAssigned,lastModifiedDateTime,settingCount',
            },
            limit,
          }),
        ),
      ]);
      // Classic profiles are the core answer; a failing beta call should not hide them.
      const catalog = 'error' in catalogResult ? { items: [], truncated: false } : catalogResult;
      if ('error' in catalogResult) ctx.log(`Settings catalog policies unavailable: ${catalogResult.error.message}`);
      const merged = [
        ...profiles.items.map((p) =>
          ctx.full
            ? { kind: 'profile', ...p }
            : compact({
                id: p.id,
                name: p.displayName ?? null,
                kind: 'profile',
                type: odataType(p),
                assigned: (p.assignments?.length ?? 0) > 0,
                lastModifiedDateTime: p.lastModifiedDateTime ?? null,
              }),
        ),
        ...catalog.items.map((p) =>
          ctx.full
            ? { kind: 'settingsCatalog', ...p }
            : compact({
                id: p.id,
                name: p.name ?? null,
                kind: 'settingsCatalog',
                type: [p.platforms, p.technologies].filter(Boolean).join(' / ') || null,
                assigned: typeof p.isAssigned === 'boolean' ? p.isAssigned : null,
                settingCount: typeof p.settingCount === 'number' ? p.settingCount : undefined,
                lastModifiedDateTime: p.lastModifiedDateTime ?? null,
              }),
        ),
      ];
      return {
        items: merged.slice(0, limit),
        truncated: profiles.truncated || catalog.truncated || merged.length > limit,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'config status',
    summary: 'Per-device results for one classic configuration profile',
    description:
      'Use to find devices where a configuration profile failed or conflicts. Works for classic profiles ' +
      '(kind "profile" in `config list`); settings catalog policies have no per-device status resource in ' +
      'Graph, so use the Intune admin center for those. Next: `devices config-states <device>` for the failing ' +
      'settings on one device.',
    examples: ['m365 intune config status "Corporate Wi-Fi" --failing-only'],
    args: [
      { name: 'policy', type: 'string', required: true, positional: true, description: 'Profile name or id.' },
      FAILING_ONLY_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_CONFIG], application: [READ_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      let profile: Policy;
      try {
        profile = await resolveCatalogItem(ctx, CONFIG_PATH, args.policy, 'configuration profile');
      } catch (error) {
        if (!(error instanceof CliError && error.code === 'NOT_FOUND')) throw error;
        const catalog = await ctx.graph
          .list<Policy>(SETTINGS_CATALOG_PATH, {
            version: 'beta',
            query: { $select: 'id,name' },
            limit: Number.POSITIVE_INFINITY,
          })
          .catch(() => ({ items: [] as Policy[] }));
        const wanted = args.policy.toLowerCase();
        const hit = catalog.items.find((p) => p.id === args.policy || (p.name ?? '').toLowerCase().includes(wanted));
        if (!hit) throw error;
        throw new CliError(
          'USAGE',
          `"${hit.name}" is a settings catalog policy; Graph has no per-device status for it`,
          {
            hint: 'Check its device status in the Intune admin center, or `devices config-states <device>` per device.',
          },
        );
      }
      const result = await policyDeviceStatuses(
        ctx,
        `${CONFIG_PATH}/${seg(profile.id)}/deviceStatuses`,
        args.failingOnly,
        listLimit(args),
      );
      if (ctx.full) return { profile, ...result };
      return {
        profile: { id: profile.id, name: profile.displayName ?? null },
        items: result.items.map(shapeDeviceStatus),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'apps list',
    summary: 'List apps in Intune with type, publisher and whether they are assigned',
    description:
      'Use to find an app before checking its deployment with `apps status`. type comes from the app kind ' +
      '(win32LobApp, iosVppApp, winGetApp, webApp...). --search matches name or publisher client-side, so it ' +
      'pages through the whole catalog.',
    examples: ['m365 intune apps list --search teams'],
    args: [{ name: 'search', type: 'string', description: 'Substring of the app name or publisher.' }, ...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: [READ_APPS], application: [READ_APPS], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const term = args.search?.toLowerCase();
      const query = { $expand: 'assignments' };
      const result = term
        ? await scan<Policy>(
            ctx,
            APPS_PATH,
            query,
            (app) => `${app.displayName ?? ''} ${app.publisher ?? ''}`.toLowerCase().includes(term),
            listLimit(args),
          )
        : await ctx.graph.list<Policy>(APPS_PATH, { query, limit: listLimit(args) });
      if (ctx.full) return result;
      return {
        items: result.items.map((app) =>
          compact({
            id: app.id,
            name: app.displayName ?? null,
            type: odataType(app),
            publisher: typeof app.publisher === 'string' ? app.publisher : null,
            isAssigned: (app.assignments?.length ?? 0) > 0,
            publishingState: app.publishingState,
            lastModifiedDateTime: app.lastModifiedDateTime ?? null,
          }),
        ),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'apps status',
    summary: 'Install summary of an app and the devices where it failed, with error codes',
    description:
      'Use for "did Zoom install on everyone" or "why did the app fail on Bob\'s laptop". summary has device and ' +
      'user counts (installed, failed, pending, notInstalled, notApplicable); failures lists devices with ' +
      'errorCode (hex, searchable in Microsoft docs) and installStateDetail. (beta API; per-device statuses ' +
      'are not populated for every app type.)',
    examples: ['m365 intune apps status "Company Portal"'],
    args: [
      { name: 'app', type: 'string', required: true, positional: true, description: 'App name or id.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_APPS], application: [READ_APPS], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const app = await resolveCatalogItem(ctx, APPS_PATH, args.app, 'app');
      const base = `${APPS_PATH}/${seg(app.id)}`;
      const sections = await settle({
        summary: ctx.graph.get<Record<string, unknown>>(`${base}/installSummary`, { version: 'beta' }),
        failures: scan<Record<string, unknown>>(
          ctx,
          `${base}/deviceStatuses`,
          {},
          (s) => s.installState === 'failed' || s.installState === 'uninstallFailed',
          listLimit(args),
          'beta',
        ),
      });
      if (ctx.full) return { app, ...sections };
      const summary = sections.summary;
      return {
        app: { id: app.id, name: app.displayName ?? null, type: odataType(app) },
        summary:
          'error' in summary
            ? summary
            : compact({
                installedDevices: summary.installedDeviceCount,
                failedDevices: summary.failedDeviceCount,
                pendingDevices: summary.pendingInstallDeviceCount,
                notInstalledDevices: summary.notInstalledDeviceCount,
                notApplicableDevices: summary.notApplicableDeviceCount,
                installedUsers: summary.installedUserCount,
                failedUsers: summary.failedUserCount,
              }),
        failures:
          'error' in sections.failures
            ? sections.failures
            : {
                items: sections.failures.items.map((s) =>
                  compact({
                    device: s.deviceName ?? null,
                    deviceId: s.deviceId ?? null,
                    user: s.userPrincipalName ?? null,
                    os: [s.osDescription, s.osVersion].filter(Boolean).join(' ') || null,
                    state: s.installState ?? null,
                    detail: s.installStateDetail ?? null,
                    errorCode:
                      typeof s.errorCode === 'number' && s.errorCode !== 0
                        ? `0x${(s.errorCode >>> 0).toString(16).toUpperCase()}`
                        : null,
                    lastSyncDateTime: s.lastSyncDateTime ?? null,
                  }),
                ),
                truncated: sections.failures.truncated,
              },
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot list',
    summary: 'List Windows Autopilot devices with group tag, enrollment state and profile status',
    description:
      'Use when a new laptop does not get the expected Autopilot experience: profileStatus must be ' +
      '"assignedInSync" before the device is reset/turned on. enrollmentState "notContacted" means it has ' +
      'never enrolled. --serial matches part of the serial; --user matches the assigned user client-side. ' +
      '(beta API, needed for the deployment profile status.)',
    examples: ['m365 intune autopilot list --serial 5CG1234', 'm365 intune autopilot list --user jane@contoso.com'],
    args: [
      { name: 'serial', type: 'string', description: 'Serial number or part of it.' },
      { name: 'user', type: 'string', description: 'Assigned user UPN or part of it.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_SERVICE_CONFIG], application: [READ_SERVICE_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const path = '/deviceManagement/windowsAutopilotDeviceIdentities';
      const query = { $filter: args.serial ? `contains(serialNumber,${odata(args.serial)})` : undefined };
      const user = args.user?.toLowerCase();
      const result = user
        ? await scan<Record<string, unknown>>(
            ctx,
            path,
            query,
            (d) => `${d.userPrincipalName ?? ''} ${d.addressableUserName ?? ''}`.toLowerCase().includes(user),
            listLimit(args),
            'beta',
          )
        : await ctx.graph.list<Record<string, unknown>>(path, { version: 'beta', query, limit: listLimit(args) });
      if (ctx.full) return result;
      return {
        items: result.items.map((d) =>
          compact({
            id: d.id,
            serialNumber: d.serialNumber ?? null,
            manufacturer: d.manufacturer || undefined,
            model: d.model ?? null,
            groupTag: d.groupTag || null,
            enrollmentState: d.enrollmentState ?? null,
            profileStatus: d.deploymentProfileAssignmentStatus ?? null,
            profileAssignedDateTime: d.deploymentProfileAssignedDateTime ?? undefined,
            assignedUser: d.userPrincipalName || null,
            lastContactedDateTime: d.lastContactedDateTime ?? null,
            daysSinceContact: daysSince(typeof d.lastContactedDateTime === 'string' ? d.lastContactedDateTime : null),
            managedDeviceId:
              typeof d.managedDeviceId === 'string' && !/^0{8}-/.test(d.managedDeviceId) ? d.managedDeviceId : null,
            azureAdDeviceId: d.azureAdDeviceId || d.azureActiveDirectoryDeviceId || null,
          }),
        ),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'report noncompliant',
    summary: 'Non-compliant devices grouped by OS and by the top failing policies and settings',
    description:
      'Use for "how many devices are non-compliant and why". Lists noncompliant and inGracePeriod devices ' +
      '(--top caps the device list), then inspects the failing settings of up to --inspect devices (batched, 20 ' +
      'per Graph call) to rank reasons. byReason counts how many inspected devices fail each policy/setting. Devices ' +
      'whose inspection failed carry `error`. Next: `devices compliance <device>` or `compliance status`.',
    examples: ['m365 intune report noncompliant', 'm365 intune report noncompliant --all --inspect 100'],
    args: [
      {
        name: 'inspect',
        type: 'number',
        default: 25,
        description: 'How many devices to inspect for failing settings.',
      },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: [READ_DEVICES, READ_CONFIG],
      application: [READ_DEVICES, READ_CONFIG],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const devices = await listDevices(
        ctx,
        [
          {
            filter: `complianceState eq 'noncompliant' or complianceState eq 'inGracePeriod'`,
            documented: true,
            test: () => true,
          },
        ],
        listLimit(args),
      );
      const inspected = devices.items.slice(0, Math.max(0, args.inspect));
      const states = await ctx.graph.batch<{ value?: PolicyState[] }>(
        inspected.map((device) => ({ path: `${DEVICES_PATH}/${seg(device.id)}/deviceCompliancePolicyStates` })),
      );
      const reasons: Record<string, { policy: string; setting: string | null; devices: number }> = {};
      const details: Record<string, { summary?: string; error?: unknown }> = {};
      inspected.forEach((device, index) => {
        const result = states[index];
        if (!result || result.error) {
          details[device.id] = {
            error: { code: result?.error?.code ?? 'INTERNAL', message: result?.error?.message ?? 'No batch response' },
          };
          return;
        }
        const policies = shapePolicyStates(result.body?.value ?? []);
        details[device.id] = { summary: complianceSummary(device, policies) };
        const seen: Record<string, true> = {};
        for (const policy of policies) {
          if (!FAILING_STATUS[policy.state ?? '']) continue;
          const settings = policy.failingSettings.length ? policy.failingSettings.map((s) => s.setting) : [null];
          for (const setting of settings) {
            const key = `${policy.policy}\u0000${setting ?? ''}`;
            if (seen[key]) continue;
            seen[key] = true;
            const entry = (reasons[key] ??= { policy: policy.policy, setting, devices: 0 });
            entry.devices++;
          }
        }
      });
      return {
        total: devices.items.length,
        truncated: devices.truncated,
        inspected: inspected.length,
        byState: tally(devices.items, (d) => d.complianceState ?? 'unknown'),
        byOs: tally(devices.items, (d) => d.operatingSystem ?? 'unknown'),
        byReason: Object.values(reasons).sort((a, b) => b.devices - a.devices),
        devices: devices.items.map((device) => ({ ...shapeManagedDevice(device), ...details[device.id] })),
      };
    },
  }),
];
