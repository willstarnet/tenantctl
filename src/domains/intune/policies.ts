import { existsSync, lstatSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve, sep } from 'node:path';
import { configDir } from '../../core/config.ts';
import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { attempt, compact, settle, shellQuote } from '../../core/util.ts';
import { listAssignments, resolveConfigProfile, shapeAssignmentLists } from './assignments.ts';
import {
  APPS_PATH,
  COMPLIANCE_PATH,
  CONFIG_PATH,
  DEVICES_PATH,
  INTUNE_ADMIN_ROLES,
  INTUNE_READ_ROLES,
  listDevices,
  mapLimit,
  odataType,
  type Policy,
  READ_APPS,
  READ_CONFIG,
  READ_DEVICES,
  READ_GROUPS,
  resolveCatalogItem,
  SETTINGS_CATALOG_PATH,
  scan,
  shapeManagedDevice,
} from './common.ts';
import { complianceSummary, fetchReportRows, type PolicyState, type ReportRow, shapePolicyStates } from './devices.ts';

const CONCURRENCY = 5;
const ALL = Number.POSITIVE_INFINITY;
const WRITE_CONFIG = 'DeviceManagementConfiguration.ReadWrite.All';

const FAILING_ONLY_ARG = {
  name: 'failingOnly',
  type: 'boolean',
  description: 'Only devices in nonCompliant, error or conflict state.',
} as const;

const FAILING_STATUS: Record<string, true> = { nonCompliant: true, error: true, conflict: true };

function num(value: unknown): number | undefined {
  return typeof value === 'number' ? value : undefined;
}

function text(value: unknown): string | null {
  return typeof value === 'string' && value !== '' ? value : null;
}

/** One getAppStatusOverviewReport row; null when Intune has no results for the app yet. */
function shapeAppSummary(row: ReportRow | undefined): Record<string, unknown> | null {
  if (!row) return null;
  return compact({
    installedDevices: num(row.InstalledDeviceCount),
    failedDevices: num(row.FailedDeviceCount),
    pendingDevices: num(row.PendingInstallDeviceCount),
    notInstalledDevices: num(row.NotInstalledDeviceCount),
    notApplicableDevices: num(row.NotApplicableDeviceCount),
    installedUsers: num(row.InstalledUserCount),
    failedUsers: num(row.FailedUserCount),
    pendingUsers: num(row.PendingInstallUserCount),
  });
}

const NO_USER = '00000000-0000-0000-0000-000000000000';

/** Sort key for AppInstallState_loc ("Failed", "Install Pending", "Not Installed", "Installed", "Not Applicable"). */
function appStateRank(state: string): number {
  const s = state.toLowerCase();
  if (s.includes('fail')) return 0;
  if (s.includes('pending')) return 1;
  if (s.includes('not installed')) return 2;
  if (s.includes('not applicable')) return 4;
  return 3;
}

function shapeAppDevices(rows: ReportRow[], args: { failedOnly?: boolean; top: number; all: boolean }) {
  const shaped = rows
    .map((r) => {
      const code = num(r.ErrorCode);
      return compact({
        device: text(r.DeviceName),
        deviceId: text(r.DeviceId),
        user: r.UserId === NO_USER ? null : text(r.UserPrincipalName),
        platform: text(r.Platform),
        appVersion: text(r.AppVersion),
        state: text(r.AppInstallState_loc) ?? 'unknown',
        detail: text(r.AppInstallStateDetails_loc),
        errorCode: text(r.HexErrorCode) ?? (code ? `0x${(code >>> 0).toString(16).toUpperCase()}` : null),
        lastModifiedDateTime: text(r.LastModifiedDateTime),
      });
    })
    .filter((d) => !args.failedOnly || appStateRank(String(d.state)) === 0)
    .sort((a, b) => appStateRank(String(a.state)) - appStateRank(String(b.state)));
  const limit = listLimit(args);
  return { items: shaped.slice(0, limit), truncated: shaped.length > limit };
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

function tally<T>(items: T[], key: (item: T) => string): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const item of items) counts[key(item)] = (counts[key(item)] ?? 0) + 1;
  return counts;
}

const PROFILE_ARG = {
  name: 'policy',
  type: 'string',
  required: true,
  positional: true,
  description: 'Profile name or id (classic profile or settings catalog policy).',
} as const;

/** Classic profile properties that describe the profile rather than configure the device. */
const PROFILE_METADATA: Record<string, true> = {
  id: true,
  displayName: true,
  description: true,
  createdDateTime: true,
  lastModifiedDateTime: true,
  version: true,
  roleScopeTagIds: true,
  supportsScopeTags: true,
  deviceManagementApplicabilityRuleOsEdition: true,
  deviceManagementApplicabilityRuleOsVersion: true,
  deviceManagementApplicabilityRuleDeviceMode: true,
  assignments: true,
};

/** A classic profile's configured settings: no metadata, annotations, nulls, empty lists or "notConfigured". */
function classicSettings(profile: Policy): Record<string, unknown> {
  const settings: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(profile)) {
    if (PROFILE_METADATA[key] || key.includes('@odata.')) continue;
    if (value === null || value === '' || value === 'notConfigured') continue;
    if (Array.isArray(value) && !value.length) continue;
    settings[key] = value;
  }
  return settings;
}

interface SettingInstance {
  '@odata.type'?: string;
  settingDefinitionId?: string;
  choiceSettingValue?: ChoiceValue | null;
  choiceSettingCollectionValue?: ChoiceValue[] | null;
  simpleSettingValue?: { value?: unknown } | null;
  simpleSettingCollectionValue?: { value?: unknown }[] | null;
  groupSettingValue?: { children?: SettingInstance[] | null } | null;
  groupSettingCollectionValue?: { children?: SettingInstance[] | null }[] | null;
}

interface ChoiceValue {
  value?: string | null;
  children?: SettingInstance[] | null;
}

interface CatalogSetting {
  id?: string;
  settingInstance?: SettingInstance;
  [key: string]: unknown;
}

interface FlatSetting {
  id: string;
  /** Simple value, choice value id, list of values, or child settings for groups (one list per group entry). */
  value: unknown;
  /** Choice values are `<settingDefinitionId>_<option>`; this is `<option>` (`block`, `1`...). */
  choice?: string;
  children?: FlatSetting[];
}

/** Flatten one settings catalog setting instance (choice, simple, collections, groups) into `{ id, value }`. */
function flattenSetting(instance: SettingInstance): FlatSetting {
  const id = instance.settingDefinitionId ?? '';
  if (instance.choiceSettingValue) {
    const value = instance.choiceSettingValue.value ?? null;
    const children = (instance.choiceSettingValue.children ?? []).map(flattenSetting);
    return compact({
      id,
      value,
      choice: value?.startsWith(`${id}_`) ? value.slice(id.length + 1) : undefined,
      children: children.length ? children : undefined,
    }) as FlatSetting;
  }
  if (instance.choiceSettingCollectionValue) {
    return { id, value: instance.choiceSettingCollectionValue.map((v) => v.value ?? null) };
  }
  if (instance.simpleSettingValue) return { id, value: instance.simpleSettingValue.value ?? null };
  if (instance.simpleSettingCollectionValue) {
    return { id, value: instance.simpleSettingCollectionValue.map((v) => v.value ?? null) };
  }
  if (instance.groupSettingValue) {
    return { id, value: (instance.groupSettingValue.children ?? []).map(flattenSetting) };
  }
  if (instance.groupSettingCollectionValue) {
    return {
      id,
      value: instance.groupSettingCollectionValue.map((group) => (group.children ?? []).map(flattenSetting)),
    };
  }
  return { id, value: null };
}

/** Deep copy without nulls and OData annotations, keeping the `@odata.type` Graph needs on setting instances. */
function stripAnnotations(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(stripAnnotations);
  if (typeof value !== 'object' || value === null) return value;
  const out: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value)) {
    if (entry === null || (key.includes('@odata.') && key !== '@odata.type')) continue;
    out[key] = stripAnnotations(entry);
  }
  return out;
}

/** Body of a settings catalog policy create: what `config export` writes and `config create` sends. */
interface PolicyBody {
  name: string;
  description: string;
  platforms: string;
  technologies: string;
  roleScopeTagIds?: unknown[];
  templateReference?: { templateId: string };
  settings: { settingInstance: unknown }[];
}

/**
 * The create body for a settings catalog policy: only the writable policy properties, settings without their ids.
 * Accepts an export file or a raw `config get --full` policy alike. templateFamily/DisplayName/Version are read-only.
 */
function policyBody(policy: Record<string, unknown>): PolicyBody {
  const template = policy.templateReference as { templateId?: string | null } | null | undefined;
  const body: PolicyBody = {
    name: String(policy.name ?? ''),
    description: typeof policy.description === 'string' ? policy.description : '',
    platforms: String(policy.platforms ?? ''),
    technologies: String(policy.technologies ?? ''),
    settings: ((policy.settings ?? []) as CatalogSetting[]).map((setting) => ({
      settingInstance: stripAnnotations(setting.settingInstance),
    })),
  };
  if (Array.isArray(policy.roleScopeTagIds)) body.roleScopeTagIds = policy.roleScopeTagIds;
  if (template?.templateId) body.templateReference = { templateId: template.templateId };
  return body;
}

/** Read and validate a policy file for `config create`, so a bad file fails before anything reaches Graph. */
function readPolicyFile(path: string, name: string | undefined) {
  if (!existsSync(path)) throw new CliError('USAGE', `Policy file not found: ${path}`);
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(path, 'utf8'));
  } catch (error) {
    throw new CliError('USAGE', `${path} is not valid JSON: ${(error as Error).message}`, {
      hint: 'Start from `tenantctl intune config export <policy> --out <file>`.',
    });
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new CliError('USAGE', `${path} must contain one JSON object (a settings catalog policy)`);
  }
  const policy: Record<string, unknown> = { ...(parsed as Record<string, unknown>), ...(name ? { name } : {}) };
  const missing = ['name', 'platforms', 'technologies'].filter(
    (key) => typeof policy[key] !== 'string' || !(policy[key] as string).trim(),
  );
  if (missing.length) {
    throw new CliError('USAGE', `${path} is missing ${missing.join(', ')}`, {
      hint: missing.includes('name') ? 'Add "name" to the file or pass --name.' : undefined,
    });
  }
  const settings = policy.settings;
  if (!Array.isArray(settings) || !settings.length) {
    throw new CliError('USAGE', `${path} has no settings; "settings" must be a non-empty array`);
  }
  const settingIds: string[] = [];
  settings.forEach((setting: CatalogSetting | null, index) => {
    const id = setting?.settingInstance?.settingDefinitionId;
    if (typeof id !== 'string') {
      throw new CliError('USAGE', `settings[${index}] in ${path} has no settingInstance.settingDefinitionId`);
    }
    settingIds.push(id);
  });
  return { body: policyBody(policy), settingIds };
}

/** A settings catalog policy with all its settings (`$expand=settings`, following a settings nextLink if any). */
async function getCatalogPolicy(ctx: Context, id: string): Promise<Policy & { settings: CatalogSetting[] }> {
  const policy = await ctx.graph.get<Policy>(`${SETTINGS_CATALOG_PATH}/${seg(id)}`, {
    version: 'beta',
    query: { $expand: 'settings' },
  });
  const settings = [...((policy.settings ?? []) as CatalogSetting[])];
  const next = policy['settings@odata.nextLink'];
  if (typeof next === 'string') {
    settings.push(...(await ctx.graph.list<CatalogSetting>(next, { version: 'beta', limit: ALL })).items);
  }
  return { ...policy, settings };
}

interface SettingSummary {
  settingName?: string | null;
  instancePath?: string | null;
  compliantDeviceCount?: number;
  nonCompliantDeviceCount?: number;
  errorDeviceCount?: number;
  conflictDeviceCount?: number;
  notApplicableDeviceCount?: number;
  remediatedDeviceCount?: number;
  unknownDeviceCount?: number;
}

function shapeSettingSummaries(items: SettingSummary[]) {
  return items
    .map((s) => {
      const nonCompliant = s.nonCompliantDeviceCount ?? 0;
      const error = s.errorDeviceCount ?? 0;
      const conflict = s.conflictDeviceCount ?? 0;
      return compact({
        setting: (s.settingName ?? '').split('.').pop() || null,
        settingName: s.settingName ?? null,
        instancePath: s.instancePath || undefined,
        failing: nonCompliant + error + conflict,
        compliant: s.compliantDeviceCount ?? 0,
        nonCompliant,
        error,
        conflict,
        notApplicable: s.notApplicableDeviceCount ?? 0,
        remediated: s.remediatedDeviceCount ?? 0,
        unknown: s.unknownDeviceCount ?? 0,
      });
    })
    .sort((a, b) => (b.failing ?? 0) - (a.failing ?? 0) || String(a.setting).localeCompare(String(b.setting)));
}

/**
 * Write an export without surprises: never into tenantctl's own config folder (audit log, tokens), never through a
 * symlink, never over an existing file unless `force`, readable only by the current user.
 */
function writeExport(path: string, content: string, force: boolean): string {
  const file = resolve(path);
  const configRoot = resolve(configDir());
  if (file === configRoot || file.startsWith(`${configRoot}${sep}`)) {
    throw new CliError('USAGE', `--out must not point into the tenantctl config folder (${configRoot})`);
  }
  const existing = lstatSync(file, { throwIfNoEntry: false });
  if (existing?.isSymbolicLink()) throw new CliError('USAGE', `--out ${file} is a symbolic link; pick a plain file`);
  if (existing && !force) {
    throw new CliError('USAGE', `${file} already exists`, {
      hint: 'Pick a new file name, or pass --force to replace it.',
    });
  }
  writeFileSync(file, content, { mode: 0o600, flag: force ? 'w' : 'wx' });
  return file;
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
    examples: ['tenantctl intune compliance list'],
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
    examples: ['tenantctl intune compliance status "Windows 11 baseline" --failing-only'],
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
    name: 'compliance settings',
    summary: 'Per-setting results for one compliance policy: how many devices pass or fail each setting',
    description:
      'Use for "which requirement are devices failing?" across a whole policy, e.g. right after rolling out a new ' +
      'one, when per-device results (`devices compliance`) may not list setting states yet. Each setting counts ' +
      'devices compliant, nonCompliant, error, conflict, notApplicable, remediated and unknown; failing = ' +
      'nonCompliant + error + conflict, and settings are sorted by it (most failing first). Next: ' +
      '`compliance status <policy> --failing-only` for the devices.',
    examples: ['tenantctl intune compliance settings "Windows 11 baseline"'],
    args: [{ name: 'policy', type: 'string', required: true, positional: true, description: 'Policy name or id.' }],
    risk: 'read',
    permissions: { delegated: [READ_CONFIG], application: [READ_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const policy = await resolveCatalogItem(ctx, COMPLIANCE_PATH, args.policy, 'compliance policy');
      const { items } = await ctx.graph.list<SettingSummary>(
        `${COMPLIANCE_PATH}/${seg(policy.id)}/deviceSettingStateSummaries`,
        { limit: ALL },
      );
      if (ctx.full) return { policy, settings: items };
      return {
        policy: { id: policy.id, name: policy.displayName ?? null },
        settings: shapeSettingSummaries(items),
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
    examples: ['tenantctl intune config list', 'tenantctl intune config list --all'],
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
    examples: ['tenantctl intune config status "Corporate Wi-Fi" --failing-only'],
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
    name: 'config get',
    summary: 'Show one configuration profile: its settings, assignments, platform and template',
    description:
      'Use to read a policy back after changing it in the Intune admin center, or to see what it actually sets. ' +
      'Works for classic profiles (v1.0; `settings` lists the configured properties, leaving out nulls, empty ' +
      'lists and "notConfigured") and settings catalog / endpoint security policies (beta API, as Microsoft has ' +
      'no v1.0 equivalent; `settings` is one `{ id, value }` per setting, with `choice` holding the readable ' +
      'option of a choice value and `children`/nested lists for groups). --full returns the raw Graph objects. ' +
      'Next: `config export` to copy a settings catalog policy, `config assign` to change who gets it.',
    examples: ['tenantctl intune config get "Edge settings"', 'tenantctl intune config get "Corporate Wi-Fi" --full'],
    args: [PROFILE_ARG],
    risk: 'read',
    permissions: {
      delegated: [READ_CONFIG, READ_GROUPS],
      application: [READ_CONFIG, READ_GROUPS],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const subject = await resolveConfigProfile(ctx, args.policy);
      const catalog = subject.kind === 'settingsCatalog';
      const [policy, assignments] = await Promise.all([
        catalog ? getCatalogPolicy(ctx, subject.id) : ctx.graph.get<Policy>(`${CONFIG_PATH}/${seg(subject.id)}`),
        attempt(listAssignments(ctx, subject)),
      ]);
      if (ctx.full) return { policy, assignments };
      const shaped = 'error' in assignments ? assignments : (await shapeAssignmentLists(ctx, [assignments]))[0];
      const common = {
        id: policy.id,
        description: policy.description || undefined,
        lastModifiedDateTime: policy.lastModifiedDateTime ?? null,
        assignments: shaped,
      };
      if (!catalog) {
        return compact({
          ...common,
          name: policy.displayName ?? null,
          kind: 'profile',
          type: odataType(policy),
          settings: classicSettings(policy),
        });
      }
      const template = (policy.templateReference ?? {}) as Record<string, string | null | undefined>;
      const settings = (policy.settings as CatalogSetting[]).flatMap((s) =>
        s.settingInstance ? [flattenSetting(s.settingInstance)] : [],
      );
      return compact({
        ...common,
        name: policy.name ?? null,
        kind: template.templateFamily?.startsWith('endpointSecurity') ? 'endpointSecurity' : 'settingsCatalog',
        platforms: policy.platforms ?? null,
        technologies: policy.technologies ?? null,
        templateReference: template.templateId
          ? compact({
              templateId: template.templateId,
              templateFamily: template.templateFamily ?? undefined,
              templateDisplayName: template.templateDisplayName ?? undefined,
              templateDisplayVersion: template.templateDisplayVersion ?? undefined,
            })
          : undefined,
        settingCount: settings.length,
        settings,
      });
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'config export',
    summary: 'Export a settings catalog policy as a JSON body ready for `config create`',
    description:
      'Settings catalog and endpoint security policies only (beta API). The output has name, description, ' +
      'platforms, technologies, roleScopeTagIds, templateReference (template-based policies) and settings, with ' +
      'ids, nulls and OData annotations removed except the @odata.type Graph needs on each setting instance. ' +
      "--out writes it to a file. To change a policy's settings: export, edit the file, `config create` it under " +
      'a new name, assign the new policy, then unassign the old one.',
    examples: [
      'tenantctl intune config export "Edge settings"',
      'tenantctl intune config export "Edge settings" --out edge.json',
    ],
    args: [
      { ...PROFILE_ARG, description: 'Settings catalog policy name or id.' },
      { name: 'out', type: 'string', description: 'Write the JSON to this new file instead of printing it.' },
      { name: 'force', type: 'boolean', description: 'Overwrite --out if it already exists (never a symlink).' },
    ],
    risk: 'read',
    permissions: { delegated: [READ_CONFIG], application: [READ_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const subject = await resolveConfigProfile(ctx, args.policy);
      if (subject.kind !== 'settingsCatalog') {
        throw new CliError(
          'USAGE',
          `"${subject.name}" is a classic configuration profile; only settings catalog policies can be exported`,
          {
            hint: `Read it with \`tenantctl intune config get ${subject.id}\`.`,
          },
        );
      }
      const policy = await getCatalogPolicy(ctx, subject.id);
      if (ctx.full && !args.out) return policy;
      const body = policyBody(policy);
      if (!args.out) return body;
      const file = writeExport(args.out, `${JSON.stringify(body, null, 2)}\n`, args.force);
      return {
        file,
        name: body.name,
        settingCount: body.settings.length,
        next: `tenantctl intune config create --file ${shellQuote(file)} --name '<new name>'`,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'config create',
    summary: 'Create a settings catalog policy from a JSON file (unassigned until you run `config assign`)',
    description:
      'Creates a settings catalog / endpoint security policy (beta API) from a `config export` file. The file ' +
      'must hold one object with name, platforms, technologies and a non-empty settings array; --name overrides ' +
      'the name. The new policy applies to nobody until you assign it with `tenantctl intune config assign`. ' +
      "Graph's documented PATCH on these policies updates only name, description, scope tags and similar " +
      'properties, not settings, so to change settings: `config export` the old policy, edit the file, ' +
      '`config create` it, `config assign` the new one, then `config unassign` (or delete) the old one.',
    examples: [
      'tenantctl intune config create --file edge.json --name "Edge settings v2"',
      'tenantctl intune config create --file edge.json --dry-run',
    ],
    args: [
      { name: 'file', type: 'string', required: true, description: 'Policy JSON, e.g. from `config export --out`.' },
      { name: 'name', type: 'string', description: 'Name for the new policy (overrides the file).' },
    ],
    risk: 'write',
    permissions: { delegated: [WRITE_CONFIG], application: [WRITE_CONFIG], roles: INTUNE_ADMIN_ROLES },
    async preview(_ctx, args) {
      const { body, settingIds } = readPolicyFile(args.file, args.name);
      return {
        action:
          `Create settings catalog policy "${body.name}" (${body.platforms}, ${body.technologies}) with ` +
          `${settingIds.length} setting(s). It applies to no one until assigned.`,
        name: body.name,
        platforms: body.platforms,
        technologies: body.technologies,
        settingCount: settingIds.length,
        settingIds,
      };
    },
    async run(ctx, args) {
      const { body } = readPolicyFile(args.file, args.name);
      const created = await ctx.graph.post<Policy>(SETTINGS_CATALOG_PATH, body, { version: 'beta' });
      return {
        id: created.id,
        name: created.name ?? body.name,
        settingCount: body.settings.length,
        next: `tenantctl intune config assign ${created.id} --group <group>`,
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
    examples: ['tenantctl intune apps list --search teams'],
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
    summary: 'Install summary of an app and its state on every targeted device, with error codes',
    description:
      'Use for "did Zoom install on everyone" or "why did the app fail on Bob\'s laptop". summary has device and ' +
      'user counts (installed, failed, pending, notInstalled, notApplicable), or is null until Intune has ' +
      'results for a newly assigned app. devices has one row per device and user context (the system row has ' +
      'no user), failed first, with state, detail and errorCode (hex, searchable in Microsoft docs). ' +
      '--failed-only keeps failed rows. Reads the Intune app install reports, the same data as the admin ' +
      'center (beta API).',
    examples: ['tenantctl intune apps status "Company Portal"', 'tenantctl intune apps status "Zoom" --failed-only'],
    args: [
      { name: 'app', type: 'string', required: true, positional: true, description: 'App name or id.' },
      { name: 'failedOnly', type: 'boolean', description: 'Only devices where the install failed.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_APPS], application: [READ_APPS], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const app = await resolveCatalogItem(ctx, APPS_PATH, args.app, 'app');
      const filter = `(ApplicationId eq ${odata(app.id)})`;
      const sections = await settle({
        summary: fetchReportRows(ctx, 'getAppStatusOverviewReport', { filter }, 'app status'),
        devices: fetchReportRows(ctx, 'retrieveDeviceAppInstallationStatusReport', { filter }, 'app install'),
      });
      if (ctx.full) return { app, ...sections };
      return {
        app: { id: app.id, name: app.displayName ?? null, type: odataType(app) },
        summary: 'error' in sections.summary ? sections.summary : shapeAppSummary(sections.summary[0]),
        devices: 'error' in sections.devices ? sections.devices : shapeAppDevices(sections.devices, args),
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
    examples: ['tenantctl intune report noncompliant', 'tenantctl intune report noncompliant --all --inspect 100'],
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
