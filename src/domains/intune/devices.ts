import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { bytesToGb, compact, daysAgo, settle } from '../../core/util.ts';
import {
  DEVICES_PATH,
  type DeviceClause,
  deviceSummary,
  INTUNE_READ_ROLES,
  listDevices,
  type ManagedDevice,
  platformOf,
  primaryUserClause,
  READ_CONFIG,
  READ_DEVICES,
  resolveManagedDevice,
  shapeManagedDevice,
} from './common.ts';

const DEVICE_ARG = {
  name: 'device',
  type: 'string',
  required: true,
  positional: true,
  description: 'Device name, serial number, Intune device id or Entra device id.',
} as const;

const REVEAL_ARG = {
  name: 'reveal',
  type: 'boolean',
  description: 'Include the secret itself. Every reveal is recorded in the Entra audit log.',
} as const;

const OS_FILTERS: Record<string, { filter: string; names: string[] }> = {
  windows: { filter: `operatingSystem eq 'Windows'`, names: ['windows'] },
  ios: { filter: `operatingSystem eq 'iOS' or operatingSystem eq 'iPadOS'`, names: ['ios', 'ipados'] },
  android: { filter: `startswith(operatingSystem,'Android')`, names: ['android'] },
  macos: { filter: `operatingSystem eq 'macOS'`, names: ['macos'] },
  linux: { filter: `operatingSystem eq 'Linux'`, names: ['linux'] },
};

/** CLI choice -> Graph complianceState enum value. */
const COMPLIANCE_STATES: Record<string, string> = {
  compliant: 'compliant',
  noncompliant: 'noncompliant',
  ingraceperiod: 'inGracePeriod',
  unknown: 'unknown',
  error: 'error',
  conflict: 'conflict',
};

/**
 * Built-in "Default Device Compliance Policy" settings fail for reasons the setting name does not explain.
 * Keyed by the last segment of `setting`.
 */
const BUILTIN_SETTING_REASONS: Record<string, string> = {
  RequireRemainContact: 'the device has not checked in with Intune within the compliance validity period',
  RequireUserExistence: 'the enrolled user no longer exists or is disabled',
  RequireDeviceCompliancePolicyAssigned: 'no compliance policy is assigned to this device or its user',
};

const FAILING_STATES: Record<string, true> = { nonCompliant: true, error: true, conflict: true };

interface SettingState {
  setting?: string | null;
  settingName?: string | null;
  state?: string | null;
  currentValue?: string | null;
  errorCode?: number | null;
  errorDescription?: string | null;
}

export interface PolicyState {
  id: string;
  displayName?: string | null;
  platformType?: string | null;
  state?: string | null;
  settingCount?: number | null;
  settingStates?: SettingState[] | null;
}

function shapeSetting(setting: SettingState): SettingView {
  const key = (setting.setting ?? '').split('.').pop() ?? '';
  return compact({
    setting: setting.settingName || setting.setting || null,
    state: setting.state ?? null,
    currentValue: setting.currentValue ?? null,
    errorCode: setting.errorCode ? setting.errorCode : undefined,
    errorDescription: setting.errorDescription || undefined,
    reason: BUILTIN_SETTING_REASONS[key],
  }) as SettingView;
}

export interface SettingView {
  setting: string | null;
  state: string | null;
  currentValue: string | null;
  errorCode?: number;
  errorDescription?: string;
  reason?: string;
}

export interface PolicyStateView {
  policy: string;
  policyId: string;
  platform: string | null;
  state: string | null;
  settingCount: number | null;
  failingSettings: SettingView[];
}

/** Per-policy state with only the settings that are not passing, for compliance and configuration alike. */
export function shapePolicyStates(states: PolicyState[]): PolicyStateView[] {
  return states.map((policy) => ({
    policy: policy.displayName ?? policy.id,
    policyId: policy.id,
    platform: policy.platformType ?? null,
    state: policy.state ?? null,
    settingCount: policy.settingCount ?? null,
    failingSettings: (policy.settingStates ?? []).filter((s) => FAILING_STATES[s.state ?? '']).map(shapeSetting),
  }));
}

/** One plain sentence an agent can repeat to the user. */
export function complianceSummary(device: ManagedDevice, policies: PolicyStateView[]): string {
  const name = device.deviceName ?? device.id;
  const failing = policies.filter((p) => FAILING_STATES[p.state ?? '']);
  if (!failing.length) {
    if (device.complianceState === 'compliant')
      return `${name} is compliant with all ${policies.length} assigned policies.`;
    return (
      `${name} reports "${device.complianceState ?? 'unknown'}" but no policy shows a failing setting; ` +
      'it may not have checked in since the policy changed. Sync it and re-check.'
    );
  }
  const reasons = failing.map((p) => {
    const settings = (p.failingSettings ?? []).map((s) => {
      const detail = s.reason ?? (s.currentValue ? `current value ${s.currentValue}` : s.errorDescription);
      return detail ? `${s.setting} (${detail})` : s.setting;
    });
    return settings.length ? `"${p.policy}" fails ${settings.join(', ')}` : `"${p.policy}" is ${p.state}`;
  });
  return `${name} is ${device.complianceState ?? 'noncompliant'}: ${reasons.join('; ')}.`;
}

/** Compliance policy states for one device, failing settings included (shared with `report noncompliant`). */
export async function devicePolicyStates(ctx: Context, deviceId: string): Promise<PolicyState[]> {
  const { items } = await ctx.graph.list<PolicyState>(`${DEVICES_PATH}/${seg(deviceId)}/deviceCompliancePolicyStates`, {
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  return items;
}

function shapeProtection(state: Record<string, unknown> | null) {
  if (!state || 'error' in state) return state;
  return compact({
    deviceState: state.deviceState,
    malwareProtectionEnabled: state.malwareProtectionEnabled,
    realTimeProtectionEnabled: state.realTimeProtectionEnabled,
    tamperProtectionEnabled: state.tamperProtectionEnabled,
    networkInspectionSystemEnabled: state.networkInspectionSystemEnabled,
    signatureVersion: state.signatureVersion,
    signatureUpdateOverdue: state.signatureUpdateOverdue,
    engineVersion: state.engineVersion,
    quickScanOverdue: state.quickScanOverdue,
    fullScanOverdue: state.fullScanOverdue,
    lastQuickScanDateTime: state.lastQuickScanDateTime,
    lastFullScanDateTime: state.lastFullScanDateTime,
    rebootRequired: state.rebootRequired,
    lastReportedDateTime: state.lastReportedDateTime,
  });
}

/** LAPS passwords are base64 of UTF-16LE on some builds and UTF-8 on others; pick by the zero-byte pattern. */
function decodeLapsPassword(base64: string): string {
  const bytes = Buffer.from(base64, 'base64');
  const utf16 = bytes.length % 2 === 0 && bytes.length > 0 && bytes.every((b, i) => i % 2 === 0 || b === 0);
  return bytes.toString(utf16 ? 'utf16le' : 'utf8');
}

function requireEntraId(device: ManagedDevice, what: string): string {
  if (!device.azureADDeviceId || /^0{8}-/.test(device.azureADDeviceId)) {
    throw new CliError('NOT_FOUND', `${device.deviceName ?? device.id} has no Entra device id, so it has no ${what}`, {
      hint: 'Only Entra-joined or registered devices back up to Entra. Check `tenantctl intune devices get` for joinType.',
    });
  }
  return device.azureADDeviceId;
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'devices list',
    summary: 'List Intune managed devices with filters for user, OS, compliance, staleness and ownership',
    description:
      'Start here for "which devices does Jane have", "show non-compliant Windows laptops" or "find devices that ' +
      'stopped checking in". daysSinceLastSync > 30 usually means the device is lost, retired or reimaged. ' +
      '--compliance and --stale-days are filtered by Graph; --user/--os/--ownership use filters Intune accepts but ' +
      'does not document, and fall back to client-side filtering (slower) if Graph rejects them. --search is ' +
      'always applied client-side, so pair it with another filter in a large tenant. ' +
      'Next: `devices get` for detail, `devices compliance` for why a device is non-compliant.',
    examples: [
      'tenantctl intune devices list --user jane@contoso.com',
      'tenantctl intune devices list --os windows --compliance noncompliant',
      'tenantctl intune devices list --stale-days 30 --all',
    ],
    args: [
      { name: 'user', type: 'string', description: 'Primary user (UPN, mail, name or id).' },
      {
        name: 'os',
        type: 'string',
        choices: ['windows', 'ios', 'android', 'macos', 'linux'],
        description: 'Platform.',
      },
      {
        name: 'compliance',
        type: 'string',
        choices: ['compliant', 'noncompliant', 'ingraceperiod', 'unknown', 'error', 'conflict'],
        description: 'Compliance state.',
      },
      { name: 'staleDays', type: 'number', description: 'Only devices that have not synced for at least N days.' },
      { name: 'search', type: 'string', description: 'Device name prefix or exact serial number.' },
      { name: 'ownership', type: 'string', choices: ['company', 'personal'], description: 'Corporate or personal.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: [READ_DEVICES, 'User.Read.All'],
      application: [READ_DEVICES, 'User.Read.All'],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const clauses: DeviceClause[] = [];
      if (args.user) clauses.push(primaryUserClause(await resolveUser(ctx, args.user)));
      if (args.os) {
        const os = OS_FILTERS[args.os] as { filter: string; names: string[] };
        clauses.push({
          filter: os.filter,
          documented: false,
          test: (d) => os.names.some((n) => (d.operatingSystem ?? '').toLowerCase().startsWith(n)),
        });
      }
      if (args.compliance) {
        const state = COMPLIANCE_STATES[args.compliance] as string;
        clauses.push({
          filter: `complianceState eq ${odata(state)}`,
          documented: true,
          test: (d) => d.complianceState === state,
        });
      }
      if (args.staleDays !== undefined) {
        const cutoff = daysAgo(args.staleDays);
        clauses.push({
          filter: `lastSyncDateTime lt ${cutoff}`,
          documented: true,
          test: (d) => (d.lastSyncDateTime ?? '') < cutoff,
        });
      }
      if (args.ownership) {
        const owner = args.ownership;
        clauses.push({
          filter: `managedDeviceOwnerType eq ${odata(owner)}`,
          documented: false,
          test: (d) => d.managedDeviceOwnerType === owner,
        });
      }
      if (args.search) {
        const term = args.search;
        const lower = term.toLowerCase();
        clauses.push({
          // Intune answers startswith() inside an `or` with an empty page instead of an error.
          filter: `startswith(deviceName,${odata(term)}) or serialNumber eq ${odata(term)}`,
          documented: false,
          clientOnly: true,
          test: (d) =>
            (d.deviceName ?? '').toLowerCase().startsWith(lower) || (d.serialNumber ?? '').toLowerCase() === lower,
        });
      }
      const result = await listDevices(ctx, clauses, listLimit(args));
      return ctx.full ? result : { items: result.items.map(shapeManagedDevice), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'devices get',
    summary: 'Show one device in detail: hardware, storage, encryption, enrollment, primary user, Defender health',
    description:
      'Use when troubleshooting a specific device. storage.percentFree under 10 often explains failed updates and ' +
      'app installs. defender (Windows only) shows antivirus health as Intune last saw it; signatureUpdateOverdue ' +
      'or realTimeProtectionEnabled=false are red flags. autopilotEnrolled/joinType come from the beta API. ' +
      'Sections that fail (e.g. missing permission) appear as { error } without hiding the rest.',
    examples: ['tenantctl intune devices get DESKTOP-4F2K9', 'tenantctl intune devices get 5CG1234XYZ'],
    args: [DEVICE_ARG],
    risk: 'read',
    permissions: { delegated: [READ_DEVICES], application: [READ_DEVICES], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const found = await resolveManagedDevice(ctx, args.device);
      const path = `${DEVICES_PATH}/${seg(found.id)}`;
      const isWindows = platformOf(found) === 'windows';
      const [device, extra] = await Promise.all([
        ctx.graph.get<ManagedDevice>(path),
        settle({
          beta: ctx.graph.get<Record<string, unknown>>(path, {
            version: 'beta',
            query: { $select: 'id,autopilotEnrolled,joinType,chassisType,skuFamily,enrolledByUserPrincipalName' },
          }),
          primaryUsers: ctx.graph
            .list<Record<string, unknown>>(`${path}/users`, {
              query: { $select: 'id,displayName,userPrincipalName' },
              pageSize: false,
            })
            .then((r) => r.items),
          protection: isWindows
            ? ctx.graph.get<Record<string, unknown>>(`${path}/windowsProtectionState`)
            : Promise.resolve(null),
        }),
      ]);
      if (ctx.full) return { device, ...extra };
      const beta = 'error' in extra.beta ? {} : extra.beta;
      const total = typeof device.totalStorageSpaceInBytes === 'number' ? device.totalStorageSpaceInBytes : null;
      const free = typeof device.freeStorageSpaceInBytes === 'number' ? device.freeStorageSpaceInBytes : null;
      return compact({
        ...shapeManagedDevice(device),
        userDisplayName: device.userDisplayName ?? null,
        primaryUsers: Array.isArray(extra.primaryUsers)
          ? extra.primaryUsers.map((u) => ({
              id: u.id,
              displayName: u.displayName,
              userPrincipalName: u.userPrincipalName,
            }))
          : extra.primaryUsers,
        storage: {
          totalGb: bytesToGb(total),
          freeGb: bytesToGb(free),
          percentFree: total && free !== null ? Math.round((free / total) * 100) : null,
        },
        jailBroken: device.jailBroken ?? null,
        isSupervised: device.isSupervised ?? null,
        autopilotEnrolled: beta.autopilotEnrolled ?? null,
        joinType: beta.joinType ?? null,
        chassisType: beta.chassisType ?? null,
        enrollmentType: device.deviceEnrollmentType ?? null,
        enrollmentProfileName: device.enrollmentProfileName || null,
        enrolledBy: beta.enrolledByUserPrincipalName ?? undefined,
        managementAgent: device.managementAgent ?? null,
        registrationState: device.deviceRegistrationState ?? null,
        deviceCategory: device.deviceCategoryDisplayName || null,
        complianceGracePeriodExpirationDateTime: device.complianceGracePeriodExpirationDateTime ?? null,
        partnerReportedThreatState: device.partnerReportedThreatState ?? null,
        exchangeAccessState: device.exchangeAccessState ?? null,
        managementCertificateExpirationDate: device.managementCertificateExpirationDate ?? null,
        androidSecurityPatchLevel: device.androidSecurityPatchLevel || undefined,
        imei: device.imei || undefined,
        phoneNumber: device.phoneNumber || undefined,
        wiFiMacAddress: device.wiFiMacAddress || undefined,
        defender: isWindows ? shapeProtection(extra.protection as Record<string, unknown> | null) : undefined,
      });
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'devices compliance',
    summary: 'Explain why a device is (non-)compliant: per-policy state and the failing settings',
    description:
      'Use for "why is my laptop non-compliant" or "why is Conditional Access blocking this device". `summary` is ' +
      'a sentence you can relay to the user; `policies[].failingSettings` lists each failing setting with its ' +
      'current value or error. Failures of the built-in "Default Device Compliance Policy" (no check-in, user ' +
      'deleted, no policy assigned) carry a `reason`. If nothing fails but the device is not compliant, run ' +
      '`devices sync` and re-check after a few minutes.',
    examples: ['tenantctl intune devices compliance DESKTOP-4F2K9'],
    args: [DEVICE_ARG],
    risk: 'read',
    permissions: {
      delegated: [READ_DEVICES, READ_CONFIG],
      application: [READ_DEVICES, READ_CONFIG],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const device = await resolveManagedDevice(ctx, args.device);
      const states = await devicePolicyStates(ctx, device.id);
      if (ctx.full) return { device, policyStates: states };
      const policies = shapePolicyStates(states);
      return { device: deviceSummary(device), summary: complianceSummary(device, policies), policies };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'devices config-states',
    summary: 'Show configuration profile states on a device, with errors and conflicts first',
    description:
      'Use when a setting (Wi-Fi, VPN, BitLocker, restrictions) did not apply to a device. `problems` lists ' +
      'profiles in error or conflict with their failing settings; a conflict means two profiles set the same ' +
      'setting to different values. Covers classic configuration profiles; settings catalog policies may not ' +
      'appear here.',
    examples: ['tenantctl intune devices config-states DESKTOP-4F2K9'],
    args: [DEVICE_ARG],
    risk: 'read',
    permissions: {
      delegated: [READ_DEVICES, READ_CONFIG],
      application: [READ_DEVICES, READ_CONFIG],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const device = await resolveManagedDevice(ctx, args.device);
      const { items } = await ctx.graph.list<PolicyState>(
        `${DEVICES_PATH}/${seg(device.id)}/deviceConfigurationStates`,
        {
          limit: Number.POSITIVE_INFINITY,
          pageSize: false,
        },
      );
      if (ctx.full) return { device, configurationStates: items };
      const profiles = shapePolicyStates(items);
      const problems = profiles.filter((p) => FAILING_STATES[p.state ?? '']);
      return {
        device: deviceSummary(device),
        counts: profiles.reduce<Record<string, number>>((acc, p) => {
          const state = p.state ?? 'unknown';
          acc[state] = (acc[state] ?? 0) + 1;
          return acc;
        }, {}),
        problems,
        profiles: profiles.filter((p) => !FAILING_STATES[p.state ?? '']).map(({ failingSettings: _, ...rest }) => rest),
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'devices apps',
    summary: 'List software Intune detected on a device (inventory, not install status)',
    description:
      'Answers "what version of Chrome is on this laptop" or "is Zoom installed". This is the discovered-apps ' +
      'inventory the device last reported, not whether an Intune-assigned app installed successfully (use ' +
      '`apps status <app>` for that). Personal iOS/Android devices only report managed apps. (beta API)',
    examples: ['tenantctl intune devices apps DESKTOP-4F2K9 --search chrome'],
    args: [
      DEVICE_ARG,
      { name: 'search', type: 'string', description: 'Substring of the app name or publisher.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_DEVICES], application: [READ_DEVICES], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const device = await resolveManagedDevice(ctx, args.device);
      const detail = await ctx.graph.get<{ detectedApps?: Record<string, unknown>[] }>(
        `${DEVICES_PATH}/${seg(device.id)}`,
        {
          version: 'beta',
          query: { $select: 'id', $expand: 'detectedApps' },
        },
      );
      const term = args.search?.toLowerCase();
      const matching = (detail.detectedApps ?? [])
        .filter((app) => !term || `${app.displayName ?? ''} ${app.publisher ?? ''}`.toLowerCase().includes(term))
        .sort((a, b) => String(a.displayName ?? '').localeCompare(String(b.displayName ?? '')));
      const limit = listLimit(args);
      const items = matching.slice(0, limit);
      if (ctx.full) return { items, truncated: matching.length > limit };
      return {
        items: items.map((app) =>
          compact({ name: app.displayName, version: app.version || null, publisher: app.publisher || undefined }),
        ),
        truncated: matching.length > limit,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'devices bitlocker-key',
    summary: 'List BitLocker recovery keys escrowed in Entra for a device; --reveal shows the keys',
    description:
      'Use when a user is stuck at the BitLocker recovery screen: match the key id shown on screen (first 8 ' +
      'characters) to `items[].id`, then rerun with --reveal. Without --reveal only key ids, creation time and ' +
      'volume type are returned. Each revealed key is logged as "Read BitLocker key" in the Entra audit log. ' +
      'Needs BitlockerKey.Read.All (BitlockerKey.ReadBasic.All suffices without --reveal) and, delegated, a role ' +
      'such as Cloud Device Administrator, Helpdesk Administrator or Intune Administrator. No keys usually means ' +
      'the device never backed up to Entra; rotate with `devices rotate-bitlocker`.',
    examples: [
      'tenantctl intune devices bitlocker-key DESKTOP-4F2K9',
      'tenantctl intune devices bitlocker-key DESKTOP-4F2K9 --reveal',
    ],
    args: [DEVICE_ARG, REVEAL_ARG],
    risk: 'read',
    permissions: {
      delegated: [READ_DEVICES, 'BitlockerKey.Read.All'],
      application: [READ_DEVICES, 'BitlockerKey.Read.All'],
      roles: ['Cloud Device Administrator', 'Helpdesk Administrator', 'Intune Administrator', 'Security Reader'],
    },
    async run(ctx, args) {
      const device = await resolveManagedDevice(ctx, args.device);
      const deviceId = requireEntraId(device, 'BitLocker keys in Entra');
      type RecoveryKey = { id: string; createdDateTime?: string; volumeType?: string; key?: string | null };
      const { items } = await ctx.graph.list<RecoveryKey>('/informationProtection/bitlocker/recoveryKeys', {
        query: { $filter: `deviceId eq ${odata(deviceId)}` },
        pageSize: false,
        limit: Number.POSITIVE_INFINITY,
      });
      const keys: RecoveryKey[] = args.reveal
        ? await Promise.all(
            items.map(async (key) => {
              const full = await ctx.graph.get<{ key?: string }>(
                `/informationProtection/bitlocker/recoveryKeys/${seg(key.id)}`,
                {
                  query: { $select: 'key' },
                },
              );
              return { ...key, key: full.key ?? null };
            }),
          )
        : items;
      if (ctx.full) return { device, keys };
      return {
        device: deviceSummary(device),
        keys: keys
          .sort((a, b) => String(b.createdDateTime ?? '').localeCompare(String(a.createdDateTime ?? '')))
          .map((k) =>
            compact({
              id: k.id,
              createdDateTime: k.createdDateTime ?? null,
              volumeType: k.volumeType ?? null,
              key: k.key,
            }),
          ),
        note: keys.length ? undefined : 'No recovery key is escrowed in Entra for this device.',
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'devices laps',
    summary: 'Show the Windows LAPS local admin account backed up to Entra; --reveal shows the password',
    description:
      'Use when someone needs local admin on a Windows device. Without --reveal returns the last backup and next ' +
      'rotation time only; --reveal adds the account name and decoded password (newest first). Reveals are ' +
      'logged in the Entra audit log. Needs DeviceLocalCredential.Read.All (ReadBasic.All is enough without ' +
      '--reveal) and, delegated, Cloud Device Administrator or Intune Administrator to reveal. Not found means ' +
      'Windows LAPS is not configured to back up to Entra for this device. Rotate with `devices rotate-laps`.',
    examples: ['tenantctl intune devices laps DESKTOP-4F2K9 --reveal'],
    args: [DEVICE_ARG, REVEAL_ARG],
    risk: 'read',
    permissions: {
      delegated: [READ_DEVICES, 'DeviceLocalCredential.Read.All'],
      application: [READ_DEVICES, 'DeviceLocalCredential.Read.All'],
      roles: ['Cloud Device Administrator', 'Intune Administrator'],
    },
    async run(ctx, args) {
      const device = await resolveManagedDevice(ctx, args.device);
      const deviceId = requireEntraId(device, 'LAPS password in Entra');
      const response = await ctx.graph.tryGet<Record<string, unknown>>(
        `/directory/deviceLocalCredentials/${seg(deviceId)}`,
        {
          query: args.reveal ? { $select: 'id,deviceName,lastBackupDateTime,refreshDateTime,credentials' } : undefined,
        },
      );
      if (!response) {
        throw new CliError(
          'NOT_FOUND',
          `No LAPS password is backed up to Entra for ${device.deviceName ?? device.id}`,
          {
            hint: 'Check that a Windows LAPS policy with backup directory "Azure AD" is assigned and the device has synced.',
          },
        );
      }
      // The docs show the object both bare and wrapped in `value`; accept either.
      const info = (typeof response.value === 'object' && response.value !== null ? response.value : response) as {
        lastBackupDateTime?: string;
        refreshDateTime?: string;
        credentials?: { accountName?: string; accountSid?: string; backupDateTime?: string; passwordBase64?: string }[];
      };
      if (ctx.full) return { device, credential: info };
      const credentials = [...(info.credentials ?? [])].sort((a, b) =>
        String(b.backupDateTime ?? '').localeCompare(String(a.backupDateTime ?? '')),
      );
      return compact({
        device: deviceSummary(device),
        lastBackupDateTime: info.lastBackupDateTime ?? null,
        nextRotationDateTime: info.refreshDateTime ?? null,
        credentials: args.reveal
          ? credentials.map((c, index) => ({
              accountName: c.accountName ?? null,
              backupDateTime: c.backupDateTime ?? null,
              current: index === 0,
              password: c.passwordBase64 ? decodeLapsPassword(c.passwordBase64) : null,
            }))
          : undefined,
      });
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'devices actions',
    summary: 'Show remote actions sent to a device (wipe, sync, lock...) and whether they completed',
    description:
      'Use after any remote action to check progress. state: pending (device has not picked it up; it must be ' +
      'online), active, done, failed, notSupported. A wipe stays pending until the device next checks in.',
    examples: ['tenantctl intune devices actions DESKTOP-4F2K9'],
    args: [DEVICE_ARG],
    risk: 'read',
    permissions: { delegated: [READ_DEVICES], application: [READ_DEVICES], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const found = await resolveManagedDevice(ctx, args.device);
      const device = await ctx.graph.get<ManagedDevice & { deviceActionResults?: Record<string, unknown>[] }>(
        `${DEVICES_PATH}/${seg(found.id)}`,
        {
          query: {
            $select:
              'id,deviceName,operatingSystem,osVersion,serialNumber,userPrincipalName,lastSyncDateTime,complianceState,' +
              'managementState,deviceActionResults',
          },
        },
      );
      const results = [...(device.deviceActionResults ?? [])].sort((a, b) =>
        String(b.startDateTime ?? '').localeCompare(String(a.startDateTime ?? '')),
      );
      if (ctx.full) return { device, actions: results };
      return {
        device: deviceSummary(device),
        actions: results.map((r) =>
          compact({
            action: r.actionName,
            state: r.actionState,
            startDateTime: r.startDateTime,
            lastUpdatedDateTime: r.lastUpdatedDateTime,
          }),
        ),
      };
    },
  }),
];
