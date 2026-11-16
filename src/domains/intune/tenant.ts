import { CliError } from '../../core/errors.ts';
import { type ApiVersion, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { compact, daysAgo, daysSince, isSettledError, type SettledError, settle } from '../../core/util.ts';
import { READ_SERVICE_CONFIG } from './autopilot.ts';
import {
  type DeviceClause,
  groupNamesById,
  INTUNE_ADMIN_ROLES,
  INTUNE_READ_ROLES,
  listDevices,
  type ManagedDevice,
  READ_CONFIG,
  READ_DEVICES,
  READ_GROUPS,
  WRITE_DEVICES,
} from './common.ts';

const ALL = Number.POSITIVE_INFINITY;
const DAY = 86_400_000;
const READ_RBAC = 'DeviceManagementRBAC.Read.All';
const CLEANUP_PATH = '/deviceManagement/managedDeviceCleanupRules';
/** Well-known id of the Microsoft Defender for Endpoint MTD connector; Graph returns partnerDisplayName null for it. */
const DEFENDER_CONNECTOR_ID = 'fc780465-2017-40d4-a0c5-307022471b92';

type Raw = Record<string, unknown>;

const str = (value: unknown): string | null => (typeof value === 'string' && value ? value : null);
const bool = (value: unknown): boolean | null => (typeof value === 'boolean' ? value : null);
const ids = (value: unknown): string[] => (Array.isArray(value) ? value.filter((v) => typeof v === 'string') : []);

/** Whole days until an ISO timestamp (negative once it has passed), or null when absent. */
function daysLeft(iso: string | null | undefined, now = Date.now()): number | null {
  if (!iso || iso.startsWith('0001-')) return null;
  const time = Date.parse(iso);
  return Number.isNaN(time) ? null : Math.floor((time - now) / DAY);
}

/** Optional section values: the shaped value, or the `{ error }` settle() produced. */
const mapSection = <T, R>(section: T | SettledError, fn: (value: T) => R): R | SettledError =>
  isSettledError(section) ? section : fn(section);

/** Every item of a small collection (settle() sections need a typed promise, not `.then` inferred as unknown). */
async function listAll<T>(ctx: Context, path: string, version: ApiVersion = 'v1.0'): Promise<T[]> {
  return (await ctx.graph.list<T>(path, { version, limit: ALL })).items;
}

function expiryWarning(what: string, iso: string | null, left: number | null, warnDays: number, fix: string) {
  if (left === null || left > warnDays) return null;
  const date = iso?.slice(0, 10) ?? 'unknown date';
  const when = left < 0 ? `expired ${-left} day(s) ago (${date})` : `expires in ${left} day(s) (${date})`;
  return `${what} ${when}. ${fix}`;
}

interface ComplianceView {
  markNoPolicyNoncompliant?: boolean | null;
  complianceValidityDays?: number | null;
  enhancedJailbreakDetection?: boolean | null;
}

interface ApnsView {
  configured: boolean;
  appleId?: string | null;
  expirationDateTime?: string | null;
  daysLeft?: number | null;
  serialNumber?: string | null;
  lastModifiedDateTime?: string | null;
}

interface DepTokenView {
  id?: string;
  name?: string | null;
  appleId?: string | null;
  tokenType?: string | null;
  expirationDateTime?: string | null;
  daysLeft?: number | null;
  lastSuccessfulSyncDateTime?: string | null;
  daysSinceSync?: number | null;
  lastSyncErrorCode?: number;
  syncedDeviceCount?: number | null;
}

interface VppTokenView {
  id: string;
  organization: string | null;
  appleId: string | null;
  accountType: string | null;
  state: string | null;
  expirationDateTime: string | null;
  daysLeft: number | null;
  lastSyncDateTime: string | null;
  lastSyncStatus: string | null;
}

interface GooglePlayView {
  bindStatus?: string | null;
  owner?: string | null;
  organization?: string | null;
  enterpriseType?: string | null;
  lastAppSyncDateTime?: string | null;
  lastAppSyncStatus?: string | null;
}

interface MtdPlatformView {
  platform?: string;
  compliance?: boolean | null;
  appProtection?: boolean | null;
  blockWithoutPartnerData?: boolean | null;
}

interface ConnectorView {
  id?: string;
  partner?: string | null;
  partnerState?: string | null;
  lastHeartbeatDateTime?: string | null;
  daysSinceHeartbeat?: number | null;
  unresponsiveAfterDays?: number | null;
  platforms?: MtdPlatformView[];
  blockUnsupportedOsVersions?: boolean | null;
  defenderSecurityConfiguration?: boolean | null;
}

interface TenantView {
  compliance: ComplianceView | SettledError;
  apns: ApnsView | SettledError;
  depTokens: DepTokenView[] | SettledError;
  vppTokens: VppTokenView[] | SettledError;
  managedGooglePlay: GooglePlayView | SettledError;
  mobileThreatDefense: ConnectorView[] | SettledError;
}

// ---- intune tenant ---------------------------------------------------------------------------------------------

function shapeSettings(s: Raw): ComplianceView {
  const days = typeof s.deviceComplianceCheckinThresholdDays === 'number' ? s.deviceComplianceCheckinThresholdDays : 0;
  return compact({
    markNoPolicyNoncompliant: bool(s.secureByDefault),
    complianceValidityDays: days > 0 ? days : null,
    enhancedJailbreakDetection: bool(s.enhancedJailBreak) ?? undefined,
  });
}

function shapeApns(c: Raw | null): ApnsView {
  if (!c || !str(c.expirationDateTime)) return { configured: false };
  const expires = str(c.expirationDateTime);
  return {
    configured: true,
    appleId: str(c.appleIdentifier),
    expirationDateTime: expires,
    daysLeft: daysLeft(expires),
    serialNumber: str(c.certificateSerialNumber),
    lastModifiedDateTime: str(c.lastModifiedDateTime),
  };
}

function shapeDepToken(t: Raw): DepTokenView {
  const expires = str(t.tokenExpirationDateTime);
  const synced = str(t.lastSuccessfulSyncDateTime);
  return compact({
    id: t.id as string,
    name: str(t.tokenName),
    appleId: str(t.appleIdentifier),
    tokenType: str(t.tokenType),
    expirationDateTime: expires,
    daysLeft: daysLeft(expires),
    lastSuccessfulSyncDateTime: synced,
    daysSinceSync: daysSince(synced),
    lastSyncErrorCode: typeof t.lastSyncErrorCode === 'number' && t.lastSyncErrorCode ? t.lastSyncErrorCode : undefined,
    syncedDeviceCount: typeof t.syncedDeviceCount === 'number' ? t.syncedDeviceCount : null,
  });
}

function shapeVppToken(t: Raw): VppTokenView {
  const expires = str(t.expirationDateTime);
  return {
    id: t.id as string,
    organization: str(t.organizationName),
    appleId: str(t.appleId),
    accountType: str(t.vppTokenAccountType),
    state: str(t.state),
    expirationDateTime: expires,
    daysLeft: daysLeft(expires),
    lastSyncDateTime: str(t.lastSyncDateTime),
    lastSyncStatus: str(t.lastSyncStatus),
  };
}

function shapeGooglePlay(s: Raw | null): GooglePlayView {
  if (!s) return { bindStatus: 'notBound' };
  return compact({
    bindStatus: str(s.bindStatus),
    owner: str(s.ownerUserPrincipalName),
    organization: str(s.ownerOrganizationName),
    enterpriseType: str(s.managedGooglePlayEnterpriseType),
    lastAppSyncDateTime: str(s.lastAppSyncDateTime),
    lastAppSyncStatus: str(s.lastAppSyncStatus),
  });
}

const MTD_PLATFORMS = [
  {
    platform: 'android',
    enabled: 'androidEnabled',
    mam: 'androidMobileApplicationManagementEnabled',
    prefix: 'android',
  },
  { platform: 'ios', enabled: 'iosEnabled', mam: 'iosMobileApplicationManagementEnabled', prefix: 'ios' },
  {
    platform: 'windows',
    enabled: 'windowsEnabled',
    mam: 'windowsMobileApplicationManagementEnabled',
    prefix: 'windows',
  },
  { platform: 'macos', enabled: 'macEnabled', mam: null, prefix: 'mac' },
] as const;

function shapeConnector(c: Raw): ConnectorView {
  const heartbeat = str(c.lastHeartbeatDateTime);
  return compact({
    id: c.id as string,
    partner: str(c.partnerDisplayName) ?? (c.id === DEFENDER_CONNECTOR_ID ? 'Microsoft Defender for Endpoint' : null),
    partnerState: str(c.partnerState),
    lastHeartbeatDateTime: heartbeat,
    daysSinceHeartbeat: daysSince(heartbeat),
    unresponsiveAfterDays:
      typeof c.partnerUnresponsivenessThresholdInDays === 'number' ? c.partnerUnresponsivenessThresholdInDays : null,
    platforms: MTD_PLATFORMS.map((p) =>
      compact({
        platform: p.platform,
        compliance: bool(c[p.enabled]),
        appProtection: p.mam ? bool(c[p.mam]) : undefined,
        blockWithoutPartnerData: bool(c[`${p.prefix}DeviceBlockedOnMissingPartnerData`]),
      }),
    ),
    blockUnsupportedOsVersions: bool(c.partnerUnsupportedOsVersionBlocked),
    defenderSecurityConfiguration: bool(c.microsoftDefenderForEndpointAttachEnabled) ?? undefined,
  });
}

/** Plain-language findings an admin should act on, each saying what to do. */
function tenantWarnings(view: TenantView, warnDays: number): string[] {
  const out: (string | null)[] = [];
  const ok = <T>(section: T | SettledError): section is T => !isSettledError(section);

  if (ok(view.compliance) && view.compliance.markNoPolicyNoncompliant === false) {
    out.push(
      'Devices with no compliance policy assigned are marked compliant (secureByDefault is off), so Conditional ' +
        'Access lets them in. Turn on "Mark devices with no compliance policy assigned as" Not compliant in ' +
        'Endpoint security > Device compliance > Compliance settings.',
    );
  }
  if (ok(view.apns) && view.apns.configured) {
    out.push(
      expiryWarning(
        `Apple MDM push certificate (${view.apns.appleId ?? 'unknown Apple ID'})`,
        view.apns.expirationDateTime ?? null,
        view.apns.daysLeft ?? null,
        warnDays,
        `Renew it in Devices > Enrollment > Apple > MDM Push Certificate with the SAME Apple ID ` +
          `(${view.apns.appleId ?? 'see appleId'}); a new certificate or Apple ID makes every Apple device re-enroll.`,
      ),
    );
  }
  if (ok(view.depTokens)) {
    for (const t of view.depTokens) {
      out.push(
        expiryWarning(
          `Apple enrollment (ADE) token "${t.name ?? t.id}" (${t.appleId ?? 'unknown Apple ID'})`,
          t.expirationDateTime ?? null,
          t.daysLeft ?? null,
          warnDays,
          'Download a new server token in Apple Business Manager and upload it to the same token in Intune ' +
            '(Devices > Enrollment > Apple > Enrollment program tokens).',
        ),
      );
      if (t.lastSyncErrorCode) {
        out.push(`Apple enrollment (ADE) token "${t.name ?? t.id}" last sync failed with code ${t.lastSyncErrorCode}.`);
      }
    }
  }
  if (ok(view.vppTokens)) {
    for (const t of view.vppTokens) {
      const label = `Apple VPP token "${t.organization ?? t.id}" (${t.appleId ?? 'unknown Apple ID'})`;
      out.push(
        expiryWarning(
          label,
          t.expirationDateTime,
          t.daysLeft,
          warnDays,
          'Download a new content token in Apple Business Manager and upload it to the same token in Intune ' +
            '(Tenant administration > Connectors and tokens > Apple VPP tokens).',
        ),
      );
      if (t.state && t.state !== 'valid' && !(t.daysLeft !== null && t.daysLeft < 0)) {
        out.push(`${label} is in state ${t.state}; Apple apps from it stop installing until it is fixed.`);
      }
    }
  }
  if (ok(view.managedGooglePlay)) {
    const gp = view.managedGooglePlay;
    const bound = gp.bindStatus === 'bound' || gp.bindStatus === 'boundAndValidated';
    if (bound && gp.lastAppSyncStatus && !['success', 'none'].includes(gp.lastAppSyncStatus)) {
      out.push(`Managed Google Play app sync last ended with ${gp.lastAppSyncStatus}.`);
    }
  }
  if (ok(view.mobileThreatDefense)) {
    for (const c of view.mobileThreatDefense) {
      const name = c.partner ?? c.id ?? 'MTD connector';
      if (c.partnerState && ['unresponsive', 'error', 'unavailable'].includes(c.partnerState)) {
        out.push(`${name} connector is ${c.partnerState}; Intune ignores its data in compliance until it recovers.`);
      }
      const threshold = c.unresponsiveAfterDays ?? null;
      if (c.daysSinceHeartbeat != null && threshold != null && c.daysSinceHeartbeat >= threshold) {
        out.push(`${name} connector has not sent a heartbeat for ${c.daysSinceHeartbeat} day(s).`);
      }
      const off = (c.platforms ?? []).filter((p) => p.compliance === false).map((p) => p.platform);
      if (off.length) {
        out.push(
          `${name} connector does not feed device risk into compliance for ${off.join(', ')}; turn it on per ` +
            'platform in Tenant administration > Connectors and tokens > Mobile Threat Defense.',
        );
      }
    }
  }
  return out.filter((w): w is string => w !== null);
}

// ---- intune cleanup-rules --------------------------------------------------------------------------------------

/** deviceCleanupRulePlatformType values (unknownFutureValue omitted). */
const CLEANUP_PLATFORMS = [
  'all',
  'windows',
  'windowsHolographic',
  'ios',
  'macOS',
  'androidAOSP',
  'androidDeviceAdministrator',
  'androidDedicatedAndFullyManagedCorporateOwnedWorkProfile',
  'androidPersonallyOwnedWorkProfile',
  'chromeOS',
  'visionOS',
  'tvOS',
] as const;
type CleanupPlatform = (typeof CLEANUP_PLATFORMS)[number];

const PLATFORM_LABELS: Record<CleanupPlatform, string> = {
  all: 'All platforms',
  windows: 'Windows',
  windowsHolographic: 'Windows Holographic',
  ios: 'iOS/iPadOS',
  macOS: 'macOS',
  androidAOSP: 'Android (AOSP)',
  androidDeviceAdministrator: 'Android (device administrator)',
  androidDedicatedAndFullyManagedCorporateOwnedWorkProfile:
    'Android (fully managed/dedicated/corporate-owned work profile)',
  androidPersonallyOwnedWorkProfile: 'Android (personally-owned work profile)',
  chromeOS: 'ChromeOS',
  visionOS: 'visionOS',
  tvOS: 'tvOS',
};

/**
 * managedDevice exposes only the free-text operatingSystem, so the preview matches the OS family. Android
 * management modes and HoloLens cannot be told apart from it; those previews count the whole family.
 */
const OS_FAMILY: Record<Exclude<CleanupPlatform, 'all'>, { test: (os: string) => boolean; exact: boolean }> = {
  windows: { test: (os) => os.startsWith('windows'), exact: false },
  windowsHolographic: { test: (os) => os.startsWith('windows'), exact: false },
  ios: { test: (os) => os === 'ios' || os === 'ipados', exact: true },
  macOS: { test: (os) => os.startsWith('mac'), exact: true },
  androidAOSP: { test: (os) => os.startsWith('android') || os === 'aosp', exact: false },
  androidDeviceAdministrator: { test: (os) => os.startsWith('android'), exact: false },
  androidDedicatedAndFullyManagedCorporateOwnedWorkProfile: { test: (os) => os.startsWith('android'), exact: false },
  androidPersonallyOwnedWorkProfile: { test: (os) => os.startsWith('android'), exact: false },
  chromeOS: { test: (os) => os.startsWith('chrome'), exact: true },
  visionOS: { test: (os) => os === 'visionos', exact: true },
  tvOS: { test: (os) => os === 'tvos', exact: true },
};

const MIN_DAYS = 30;
const MAX_DAYS = 270;
const SAMPLE_SIZE = 20;

interface CleanupRule {
  id: string;
  displayName?: string | null;
  description?: string | null;
  deviceCleanupRulePlatformType?: string | null;
  deviceInactivityBeforeRetirementInDays?: number | null;
  lastModifiedDateTime?: string | null;
  [key: string]: unknown;
}

const RULE_SELECT =
  'id,displayName,description,deviceCleanupRulePlatformType,deviceInactivityBeforeRetirementInDays,lastModifiedDateTime';

function shapeRule(rule: CleanupRule) {
  return compact({
    id: rule.id,
    name: rule.displayName ?? null,
    description: rule.description || undefined,
    platform: rule.deviceCleanupRulePlatformType ?? null,
    days: rule.deviceInactivityBeforeRetirementInDays ?? null,
    lastModifiedDateTime: rule.lastModifiedDateTime ?? null,
  });
}

function listRules(ctx: Context, limit: number) {
  return ctx.graph.list<CleanupRule>(CLEANUP_PATH, {
    version: 'beta',
    query: ctx.full ? undefined : { $select: RULE_SELECT },
    limit,
  });
}

function validateCleanupDays(days: number): number {
  if (!Number.isInteger(days) || days < MIN_DAYS || days > MAX_DAYS) {
    throw new CliError('USAGE', `--days must be a whole number from ${MIN_DAYS} to ${MAX_DAYS} (got ${days})`, {
      hint: 'Intune accepts 30–270 days without check-in for a device cleanup rule.',
    });
  }
  return days;
}

const STALE_SELECT = 'id,deviceName,userPrincipalName,operatingSystem,osVersion,serialNumber,lastSyncDateTime';

function staleDevice(d: ManagedDevice) {
  const synced = d.lastSyncDateTime && !d.lastSyncDateTime.startsWith('0001-') ? d.lastSyncDateTime : null;
  return {
    id: d.id,
    deviceName: d.deviceName,
    user: d.userPrincipalName || null,
    os: [d.operatingSystem, d.osVersion].filter(Boolean).join(' ') || null,
    serialNumber: d.serialNumber || null,
    lastSyncDateTime: synced,
    daysSinceLastSync: daysSince(synced),
  };
}

async function planCleanupRule(
  ctx: Context,
  args: { days: number; platform: CleanupPlatform; name?: string; description?: string },
) {
  const days = validateCleanupDays(args.days);
  const platform = args.platform;
  const label = PLATFORM_LABELS[platform];
  const name = args.name?.trim() || `${label} - ${days} days inactive`;
  const { items: rules } = await listRules(ctx, ALL);
  const samePlatform = rules.find((r) => r.deviceCleanupRulePlatformType === platform);
  if (samePlatform) {
    throw new CliError(
      'USAGE',
      `A ${label} cleanup rule already exists: "${samePlatform.displayName}" ` +
        `(${samePlatform.deviceInactivityBeforeRetirementInDays} days). Intune allows one rule per platform.`,
      { hint: 'Change that rule in the Intune admin center (Devices > Device cleanup rules) instead.' },
    );
  }
  const overlapping = rules.filter((r) =>
    platform === 'all' ? r.deviceCleanupRulePlatformType !== 'all' : r.deviceCleanupRulePlatformType === 'all',
  );

  const family = platform === 'all' ? null : OS_FAMILY[platform];
  const clauses: DeviceClause[] = [
    { filter: `lastSyncDateTime lt ${daysAgo(days)}`, documented: true, test: () => true },
  ];
  if (family) {
    clauses.push({
      filter: `operatingSystem in ${label}`,
      documented: false,
      clientOnly: true,
      test: (d) => family.test((d.operatingSystem ?? '').toLowerCase()),
    });
  }
  const { items: stale } = await listDevices(ctx, clauses, ALL, STALE_SELECT);
  const sorted = [...stale].sort((a, b) => (a.lastSyncDateTime ?? '').localeCompare(b.lastSyncDateTime ?? ''));
  const body = {
    '@odata.type': '#microsoft.graph.managedDeviceCleanupRule',
    displayName: name,
    description: args.description ?? '',
    deviceCleanupRulePlatformType: platform,
    deviceInactivityBeforeRetirementInDays: days,
  };
  const scope = platform === 'all' ? 'devices on any platform' : `${label} devices`;
  const action =
    `Create device cleanup rule "${name}" that hides ${scope} not checked in for ${days} days. ` +
    `${stale.length} device record(s) visible now would be hidden from the Intune admin center, reports and ` +
    'Graph device lists. Devices are not wiped or retired and stay in Entra ID; a hidden device reappears if it ' +
    'checks in before its certificate expires, otherwise it must re-enroll.';
  return {
    name,
    body,
    action,
    affected: compact({
      count: stale.length,
      approximate: family && !family.exact ? true : undefined,
      note:
        family && !family.exact
          ? `Counted every ${platform.startsWith('android') ? 'Android' : 'Windows'} device: Intune device records do not say which ${label.replace(/ \(.*/, '')} mode a device uses.`
          : undefined,
      devices: sorted.slice(0, SAMPLE_SIZE).map(staleDevice),
      more: stale.length > SAMPLE_SIZE ? stale.length - SAMPLE_SIZE : undefined,
    }),
    overlapping: overlapping.map(shapeRule),
  };
}

// ---- intune rbac -----------------------------------------------------------------------------------------------

interface RoleDefinition {
  id: string;
  displayName?: string | null;
  description?: string | null;
  isBuiltIn?: boolean | null;
  rolePermissions?: { resourceActions?: { allowedResourceActions?: string[] }[] }[];
  [key: string]: unknown;
}

interface RoleAssignment {
  id: string;
  displayName?: string | null;
  description?: string | null;
  members?: string[];
  resourceScopes?: string[];
  scopeMembers?: string[];
  scopeType?: string | null;
  roleScopeTagIds?: string[];
  [key: string]: unknown;
}

interface ScopeTag {
  id: string;
  displayName?: string | null;
  description?: string | null;
  isBuiltIn?: boolean | null;
}

interface ApprovalPolicy {
  id: string;
  displayName?: string | null;
  description?: string | null;
  policyType?: string | null;
  policyPlatform?: string | null;
  approverGroupIds?: string[];
  lastModifiedDateTime?: string | null;
}

async function rolesWithAssignments(ctx: Context) {
  const { items: roles } = await ctx.graph.list<RoleDefinition>('/deviceManagement/roleDefinitions', { limit: ALL });
  const results = await ctx.graph.batch<{ value?: RoleAssignment[] }>(
    roles.map((r) => ({ path: `/deviceManagement/roleDefinitions/${seg(r.id)}/roleAssignments` })),
    'beta',
  );
  return roles.map((role, index) => {
    const result = results[index];
    if (result?.error) throw result.error;
    return { role, assignments: result?.body?.value ?? [] };
  });
}

const SCOPE_TYPES: Record<string, string> = {
  allDevices: 'all devices',
  allLicensedUsers: 'all users',
  allDevicesAndLicensedUsers: 'all users and all devices',
};

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'tenant',
    summary: 'Intune tenant health: compliance settings, Apple tokens, Google Play and MTD connectors',
    description:
      'One view of the tenant-wide Intune plumbing that silently breaks enrollment when neglected. compliance: ' +
      'markNoPolicyNoncompliant (secureByDefault: devices with no compliance policy count as noncompliant), ' +
      'complianceValidityDays (null = not set, Intune default 30) and enhancedJailbreakDetection. apns: the Apple ' +
      'MDM push certificate with daysLeft; renew it with the SAME Apple ID shown in appleId, because a new Apple ID ' +
      'or a new certificate (instead of renewing) forces every iOS/iPadOS/macOS device to re-enroll, and an expired ' +
      'one stops Apple device management. depTokens (Apple ADE) and vppTokens (Apple apps) with daysLeft and last ' +
      'sync. managedGooglePlay: bind status and owner account. mobileThreatDefense: each MTD connector (the Defender ' +
      'for Endpoint connector included) with per-platform compliance and app protection (MAM) use, partner state ' +
      'and last heartbeat. warnings lists what needs action: anything expiring within --warn-days (default 30), ' +
      'secureByDefault off, an MTD platform not used in compliance. Each section degrades to { error } on its own. ' +
      '(beta API for compliance settings, ADE tokens, Managed Google Play and the macOS/Windows MTD flags.)',
    examples: ['tenantctl intune tenant', 'tenantctl intune tenant --warn-days 60'],
    args: [{ name: 'warnDays', type: 'number', default: 30, description: 'Warn about anything expiring this soon.' }],
    risk: 'read',
    permissions: {
      delegated: [READ_SERVICE_CONFIG, READ_CONFIG],
      application: [READ_SERVICE_CONFIG, READ_CONFIG],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const raw = await settle({
        compliance: ctx.graph.get<Raw>('/deviceManagement/settings', { version: 'beta' }),
        apns: ctx.graph.tryGet<Raw>('/deviceManagement/applePushNotificationCertificate'),
        depTokens: listAll<Raw>(ctx, '/deviceManagement/depOnboardingSettings', 'beta'),
        vppTokens: listAll<Raw>(ctx, '/deviceAppManagement/vppTokens'),
        managedGooglePlay: ctx.graph.tryGet<Raw>('/deviceManagement/androidManagedStoreAccountEnterpriseSettings', {
          version: 'beta',
        }),
        mobileThreatDefense: listAll<Raw>(ctx, '/deviceManagement/mobileThreatDefenseConnectors', 'beta'),
      });
      const view: TenantView = {
        compliance: mapSection(raw.compliance, shapeSettings),
        apns: mapSection(raw.apns, shapeApns),
        depTokens: mapSection(raw.depTokens, (items) => items.map(shapeDepToken)),
        vppTokens: mapSection(raw.vppTokens, (items) => items.map(shapeVppToken)),
        managedGooglePlay: mapSection(raw.managedGooglePlay, shapeGooglePlay),
        mobileThreatDefense: mapSection(raw.mobileThreatDefense, (items) => items.map(shapeConnector)),
      };
      const warnings = tenantWarnings(view, args.warnDays);
      return ctx.full ? { ...raw, warnings } : { ...view, warnings };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'cleanup-rules list',
    summary: 'List Intune device cleanup rules: platform and days without check-in',
    description:
      'Device cleanup rules hide Intune device records that have not checked in for `days`. Intune allows one rule ' +
      'per platform; where an "all" rule and a platform rule both apply, the one with fewer days wins. Rules may ' +
      'report the all-zero id. (beta API: cleanup rules have no v1.0 equivalent.)',
    examples: ['tenantctl intune cleanup-rules list'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: [READ_DEVICES], application: [READ_DEVICES], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const result = await listRules(ctx, listLimit(args));
      return { items: ctx.full ? result.items : result.items.map(shapeRule), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'cleanup-rules create',
    summary: 'Create a device cleanup rule that hides devices not checked in for N days',
    description:
      'Hides matching device records from the Intune admin center, reports and Graph device lists on a schedule ' +
      '(https://learn.microsoft.com/en-us/intune/governance/configure-cleanup-rules). It sends nothing to devices ' +
      '(no wipe or retire) and leaves Entra ID untouched; a hidden device reappears if it checks in before its ' +
      'device certificate expires, after that it must re-enroll. --days is 30–270. --platform takes the Intune ' +
      'platform names (default all); Intune allows one rule per platform, and where an "all" rule and a platform ' +
      'rule both apply the one with fewer days wins. The preview counts visible devices whose lastSyncDateTime is ' +
      'older than --days and lists up to 20, oldest first; Android modes and Windows Holographic count the whole OS ' +
      'family because device records do not say which mode a device uses. Graph answered a created rule with the ' +
      'all-zero id in a live check, so the result is read back by listing; check it with `tenantctl intune ' +
      'cleanup-rules list`. (beta API.)',
    examples: [
      'tenantctl intune cleanup-rules create --days 90',
      'tenantctl intune cleanup-rules create --days 180 --platform windows --name "Windows - 180 days inactive"',
    ],
    args: [
      { name: 'days', type: 'number', required: true, description: 'Days without check-in before hiding (30–270).' },
      {
        name: 'platform',
        type: 'string',
        choices: CLEANUP_PLATFORMS,
        default: 'all',
        description: 'Platform the rule applies to.',
      },
      { name: 'name', type: 'string', description: 'Rule name (default "<platform> - <days> days inactive").' },
      { name: 'description', type: 'string', description: 'Rule description.' },
    ],
    risk: 'destructive',
    permissions: { delegated: [WRITE_DEVICES], application: [WRITE_DEVICES], roles: INTUNE_ADMIN_ROLES },
    async preview(ctx, args) {
      const plan = await planCleanupRule(ctx, args);
      return {
        action: plan.action,
        rule: { name: plan.name, platform: args.platform, days: args.days },
        affected: plan.affected,
        ...(plan.overlapping.length
          ? { overlapping: plan.overlapping, note: 'Where two rules apply to a device, the one with fewer days wins.' }
          : {}),
        requests: [{ method: 'POST', path: `beta${CLEANUP_PATH}`, body: plan.body }],
      };
    },
    async run(ctx, args) {
      const plan = await planCleanupRule(ctx, args);
      const created = await ctx.graph.post<CleanupRule>(CLEANUP_PATH, plan.body, { version: 'beta' });
      const { items } = await listRules(ctx, ALL);
      const stored = items.find(
        (r) => r.displayName === plan.name && r.deviceCleanupRulePlatformType === args.platform,
      );
      const rule = stored ?? created;
      return {
        status: 'created',
        action: plan.action,
        rule: rule ? (ctx.full ? rule : shapeRule(rule)) : null,
        hiddenNow: plan.affected.count,
        next: 'Check it with `tenantctl intune cleanup-rules list`; Intune applies the rule on its next scheduled run.',
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'rbac',
    summary: 'Intune role-based access: roles with their assignments, scope tags and Multi Admin Approval',
    description:
      'roles: every Intune role (builtIn true/false) with its assignments: members (admin groups), scope (the ' +
      'groups whose users/devices they manage, or all users/devices) and scope tags, all by name. Intune roles are ' +
      'separate from Entra roles: Intune Administrator and Global Administrator in Entra get full Intune access ' +
      'without appearing here. scopeTags lists the tags. multiAdminApproval lists the access policies that make ' +
      'a second admin approve wipes, scripts, apps or other operations, with the approver groups. Needs ' +
      `${READ_RBAC}; without it each section shows a FORBIDDEN error and \`hint\` says how to consent. (beta API ` +
      'for assignment scope tags, scope tags and Multi Admin Approval.)',
    examples: ['tenantctl intune rbac'],
    args: [],
    risk: 'read',
    permissions: {
      delegated: [READ_RBAC, READ_GROUPS],
      application: [READ_RBAC, READ_GROUPS],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx) {
      const raw = await settle({
        roles: rolesWithAssignments(ctx),
        scopeTags: listAll<ScopeTag>(ctx, '/deviceManagement/roleScopeTags', 'beta'),
        multiAdminApproval: listAll<ApprovalPolicy>(ctx, '/deviceManagement/operationApprovalPolicies', 'beta'),
      });
      const forbidden = Object.values(raw).some((s) => isSettledError(s) && s.error.code === 'FORBIDDEN');
      const consent =
        ctx.profile.authMode === 'app'
          ? `Grant the app the ${READ_RBAC} application permission with admin consent, then rerun.`
          : `Run \`tenantctl auth login\` and consent to ${READ_RBAC} (an admin who can grant consent must approve), then rerun.`;
      const hint = forbidden ? { hint: consent } : {};
      if (ctx.full) return { ...raw, ...hint };

      const groupIds = [
        ...(isSettledError(raw.roles)
          ? []
          : raw.roles.flatMap(({ assignments }) =>
              assignments.flatMap((a) => [...ids(a.members), ...ids(a.resourceScopes), ...ids(a.scopeMembers)]),
            )),
        ...(isSettledError(raw.multiAdminApproval)
          ? []
          : raw.multiAdminApproval.flatMap((p) => ids(p.approverGroupIds))),
      ];
      const names = await groupNamesById(ctx, groupIds);
      const group = (id: string) => ({ name: names[id] ?? null, id });
      const tagNames: Record<string, string | null> = { '0': 'Default' };
      if (!isSettledError(raw.scopeTags)) for (const t of raw.scopeTags) tagNames[t.id] = t.displayName ?? null;

      const shapeAssignment = (a: RoleAssignment) => {
        const scopeGroups = [...new Set([...ids(a.resourceScopes), ...ids(a.scopeMembers)])];
        const scopeType = a.scopeType ?? 'resourceScope';
        return compact({
          name: a.displayName ?? null,
          description: a.description || undefined,
          members: ids(a.members).map(group),
          scope: SCOPE_TYPES[scopeType] ?? (scopeGroups.length ? undefined : scopeType),
          scopeGroups: scopeGroups.length ? scopeGroups.map(group) : undefined,
          scopeTags: ids(a.roleScopeTagIds).map((id) => tagNames[id] ?? id),
        });
      };
      const roles = mapSection(raw.roles, (list) =>
        list
          .map(({ role, assignments }) =>
            compact({
              id: role.id,
              name: role.displayName ?? null,
              builtIn: role.isBuiltIn ?? null,
              description: role.isBuiltIn ? undefined : role.description || undefined,
              permissionCount: role.isBuiltIn
                ? undefined
                : (role.rolePermissions ?? []).reduce(
                    (n, p) =>
                      n + (p.resourceActions ?? []).reduce((m, a) => m + (a.allowedResourceActions?.length ?? 0), 0),
                    0,
                  ),
              assignments: assignments.map(shapeAssignment),
            }),
          )
          .sort(
            (a, b) =>
              Number(a.builtIn) - Number(b.builtIn) ||
              (b.assignments?.length ?? 0) - (a.assignments?.length ?? 0) ||
              (a.name ?? '').localeCompare(b.name ?? ''),
          ),
      );
      return {
        roles,
        scopeTags: mapSection(raw.scopeTags, (tags) =>
          tags.map((t) =>
            compact({
              id: t.id,
              name: t.displayName ?? null,
              description: t.description || undefined,
              builtIn: t.isBuiltIn ?? null,
            }),
          ),
        ),
        multiAdminApproval: mapSection(raw.multiAdminApproval, (policies) =>
          policies.map((p) =>
            compact({
              id: p.id,
              name: p.displayName ?? null,
              description: p.description || undefined,
              policyType: p.policyType ?? null,
              platform: p.policyPlatform ?? null,
              approvers: ids(p.approverGroupIds).map(group),
              lastModifiedDateTime: p.lastModifiedDateTime ?? null,
            }),
          ),
        ),
        ...hint,
      };
    },
  }),
];
