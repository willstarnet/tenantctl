import { seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { compact } from '../../core/util.ts';
import { type Assignment, assignmentGroupNames, shapeAssignment } from './assignments.ts';
import { READ_SERVICE_CONFIG } from './autopilot.ts';
import { APPS_PATH, INTUNE_READ_ROLES, odataType, pickByName, READ_APPS, READ_GROUPS } from './common.ts';

const CONFIGS_PATH = '/deviceManagement/deviceEnrollmentConfigurations';
const ALL = Number.POSITIVE_INFINITY;

/** `@odata.type` (without `#microsoft.graph.`) -> the name the admin center uses. */
const FRIENDLY_TYPES: Record<string, string> = {
  windows10EnrollmentCompletionPageConfiguration: 'enrollmentStatusPage',
  deviceEnrollmentPlatformRestrictionsConfiguration: 'platformRestrictions',
  deviceEnrollmentPlatformRestrictionConfiguration: 'platformRestriction',
  deviceEnrollmentLimitConfiguration: 'deviceLimit',
  deviceEnrollmentWindowsHelloForBusinessConfiguration: 'windowsHelloForBusiness',
  deviceComanagementAuthorityConfiguration: 'coManagementAuthority',
  windowsRestoreDeviceEnrollmentConfiguration: 'windowsRestore',
  deviceEnrollmentNotificationConfiguration: 'enrollmentNotification',
};

const TYPES = [
  'enrollmentStatusPage',
  'platformRestrictions',
  'platformRestriction',
  'deviceLimit',
  'windowsHelloForBusiness',
  'coManagementAuthority',
  'windowsRestore',
  'enrollmentNotification',
] as const;

/** Properties every configuration type shares; the rest are type-specific settings. */
const BASE_KEYS: Record<string, true> = {
  '@odata.type': true,
  id: true,
  displayName: true,
  description: true,
  priority: true,
  createdDateTime: true,
  lastModifiedDateTime: true,
  version: true,
  roleScopeTagIds: true,
  deviceEnrollmentConfigurationType: true,
  assignments: true,
  'assignments@odata.context': true,
};

interface EnrollmentConfig {
  id: string;
  displayName?: string | null;
  description?: string | null;
  priority?: number | null;
  deviceEnrollmentConfigurationType?: string | null;
  lastModifiedDateTime?: string | null;
  assignments?: Assignment[];
  [key: string]: unknown;
}

interface PlatformRestriction {
  platformBlocked?: boolean | null;
  personalDeviceEnrollmentBlocked?: boolean | null;
  osMinimumVersion?: string | null;
  osMaximumVersion?: string | null;
  blockedManufacturers?: string[] | null;
}

function typeOf(config: EnrollmentConfig): string {
  const type = odataType(config) ?? 'unknown';
  return FRIENDLY_TYPES[type] ?? type;
}

function shapeRestriction(platform: string, r: PlatformRestriction) {
  return compact({
    platform,
    blocked: r.platformBlocked ?? null,
    personalBlocked: r.personalDeviceEnrollmentBlocked ?? null,
    osMinimumVersion: r.osMinimumVersion || undefined,
    osMaximumVersion: r.osMaximumVersion || undefined,
    blockedManufacturers: r.blockedManufacturers?.length ? r.blockedManufacturers : undefined,
  });
}

/** The settings an admin checks first, per configuration type; unknown types return undefined. */
function keySettings(config: EnrollmentConfig, apps: Record<string, string | null>) {
  const c = config as Record<string, unknown>;
  switch (typeOf(config)) {
    case 'enrollmentStatusPage': {
      const selected = Array.isArray(c.selectedMobileAppIds) ? (c.selectedMobileAppIds as string[]) : [];
      return {
        showProgress: c.showInstallationProgress ?? null,
        // The admin center's "Block device use until all apps and profiles are installed" is stored inverted.
        blockDeviceUseUntilInstalled:
          typeof c.blockDeviceSetupRetryByUser === 'boolean' ? !c.blockDeviceSetupRetryByUser : null,
        timeoutMinutes: c.installProgressTimeoutInMinutes ?? null,
        blockingApps: selected.length ? selected.map((id) => ({ id, name: apps[id] ?? null })) : 'all required apps',
        allowReset: c.allowDeviceResetOnInstallFailure ?? null,
        allowUseOnFailure: c.allowDeviceUseOnInstallFailure ?? null,
        allowLogCollection: c.allowLogCollectionOnInstallFailure ?? null,
        customErrorMessage: c.customErrorMessage || null,
        autopilotOnly: c.trackInstallProgressForAutopilotOnly ?? null,
        installQualityUpdates: c.installQualityUpdates ?? undefined,
      };
    }
    case 'platformRestrictions':
      return {
        platforms: Object.entries(c)
          .filter(([key, value]) => key.endsWith('Restriction') && value && typeof value === 'object')
          .map(([key, value]) => shapeRestriction(key.replace(/Restriction$/, ''), value as PlatformRestriction)),
      };
    case 'platformRestriction':
      return {
        platforms: [
          shapeRestriction(String(c.platformType ?? 'unknown'), (c.platformRestriction ?? {}) as PlatformRestriction),
        ],
      };
    case 'deviceLimit':
      return { limit: c.limit ?? null };
    case 'windowsHelloForBusiness':
      return compact({
        state: c.state ?? null,
        pinMinimumLength: c.pinMinimumLength ?? null,
        pinMaximumLength: c.pinMaximumLength ?? null,
        pinUppercase: c.pinUppercaseCharactersUsage ?? null,
        pinLowercase: c.pinLowercaseCharactersUsage ?? null,
        pinSpecialCharacters: c.pinSpecialCharactersUsage ?? null,
        pinExpirationInDays: c.pinExpirationInDays ?? undefined,
        securityDeviceRequired: c.securityDeviceRequired ?? null,
        biometrics: c.unlockWithBiometricsEnabled ?? null,
      });
    case 'windowsRestore':
      return { state: c.state ?? null };
    case 'coManagementAuthority':
      return {
        managedDeviceAuthority: c.managedDeviceAuthority ?? null,
        installConfigurationManagerAgent: c.installConfigurationManagerAgent ?? null,
      };
    default:
      return undefined;
  }
}

function shapeConfig(
  config: EnrollmentConfig,
  names: Record<string, string | null>,
  apps: Record<string, string | null>,
  detail = false,
) {
  const properties: Record<string, unknown> = {};
  if (detail) for (const [key, value] of Object.entries(config)) if (!BASE_KEYS[key]) properties[key] = value;
  return compact({
    id: config.id,
    name: config.displayName ?? null,
    type: typeOf(config),
    priority: config.priority ?? null,
    // Intune's built-in "All users and all devices" defaults have priority 0 (their type is not always default*).
    isDefault:
      config.priority === 0 || (config.deviceEnrollmentConfigurationType ?? '').startsWith('default') || undefined,
    description: detail ? config.description || null : undefined,
    settings: keySettings(config, apps),
    properties: detail ? properties : undefined,
    assignments: (config.assignments ?? []).map((a) => shapeAssignment(a, names)),
    lastModifiedDateTime: config.lastModifiedDateTime ?? null,
  });
}

/** Names of the ESP blocking apps, fetched through `$batch`; unreadable apps map to null. */
async function appNames(ctx: Context, configs: EnrollmentConfig[]) {
  const ids = [
    ...new Set(
      configs.flatMap((c) => (Array.isArray(c.selectedMobileAppIds) ? (c.selectedMobileAppIds as string[]) : [])),
    ),
  ];
  const names: Record<string, string | null> = {};
  if (!ids.length) return names;
  const results = await ctx.graph.batch<{ displayName?: string }>(
    ids.map((id) => ({ path: `${APPS_PATH}/${seg(id)}`, query: { $select: 'id,displayName' } })),
  );
  ids.forEach((id, index) => {
    const result = results[index];
    names[id] = result && !result.error ? (result.body?.displayName ?? null) : null;
  });
  return names;
}

async function shapeAll(ctx: Context, configs: EnrollmentConfig[], detail = false) {
  const [names, apps] = await Promise.all([
    assignmentGroupNames(
      ctx,
      configs.flatMap((c) => c.assignments ?? []),
    ),
    appNames(ctx, configs),
  ]);
  return configs.map((c) => shapeConfig(c, names, apps, detail));
}

const PERMISSIONS = {
  delegated: [READ_SERVICE_CONFIG, READ_GROUPS, READ_APPS],
  application: [READ_SERVICE_CONFIG, READ_GROUPS, READ_APPS],
  roles: INTUNE_READ_ROLES,
};

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'enrollment list',
    summary: 'List enrollment configurations: ESP, platform restrictions, device limit, Windows Hello',
    description:
      'Sorted by type, then priority (lower wins when a user is in several assigned groups; isDefault applies to ' +
      'everyone else). enrollmentStatusPage: blockDeviceUseUntilInstalled, timeoutMinutes, blockingApps ("all ' +
      'required apps" or the selected apps by name), allowReset/allowUseOnFailure on failure and the custom ' +
      'message. platformRestrictions: blocked and personalBlocked per platform. Assignments show group names. ' +
      'Details: `tenantctl intune enrollment get <name>`. (beta API: v1.0 lacks the ESP blocking apps and ' +
      'single-platform restrictions.)',
    examples: ['tenantctl intune enrollment list', 'tenantctl intune enrollment list --type enrollmentStatusPage'],
    args: [
      { name: 'type', type: 'string', choices: TYPES, description: 'Only configurations of this type.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: PERMISSIONS,
    async run(ctx, args) {
      const result = await ctx.graph.list<EnrollmentConfig>(CONFIGS_PATH, {
        version: 'beta',
        query: { $expand: 'assignments' },
        limit: args.type ? ALL : listLimit(args),
      });
      const matching = args.type ? result.items.filter((c) => typeOf(c) === args.type) : result.items;
      const limit = listLimit(args);
      const items = matching
        .slice()
        .sort((a, b) => typeOf(a).localeCompare(typeOf(b)) || (a.priority ?? 0) - (b.priority ?? 0))
        .slice(0, limit);
      const truncated = result.truncated || matching.length > items.length;
      if (ctx.full) return { items, truncated };
      return { items: await shapeAll(ctx, items), truncated };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'enrollment get',
    summary: 'Show one enrollment configuration with all its settings and assignments',
    description:
      'settings holds the key values (as in `tenantctl intune enrollment list`); properties holds every ' +
      'type-specific property Graph returns. Find names with `tenantctl intune enrollment list`. (beta API.)',
    examples: ['tenantctl intune enrollment get "All users ESP"'],
    args: [
      { name: 'config', type: 'string', required: true, positional: true, description: 'Configuration name or id.' },
    ],
    risk: 'read',
    permissions: PERMISSIONS,
    async run(ctx, args) {
      const { items } = await ctx.graph.list<EnrollmentConfig>(CONFIGS_PATH, {
        version: 'beta',
        query: { $expand: 'assignments' },
        limit: ALL,
      });
      const config = pickByName(items, args.config, 'enrollment configuration', (c) => c.displayName);
      if (ctx.full) return config;
      const [shaped] = await shapeAll(ctx, [config], true);
      return shaped;
    },
  }),
];
