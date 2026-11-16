import { createHash } from 'node:crypto';
import { readCsvRows } from '../../core/csvInput.ts';
import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { pickOne, type ResolvedGroup, resolveGroup, resolveUser } from '../../core/resolve.ts';
import { requireApprovedTarget } from '../../core/safety.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact, daysAgo, daysSince, isGuid, settle, shellQuote } from '../../core/util.ts';
import { type Assignment, assignmentGroupNames, memberCount, shapeAssignment } from './assignments.ts';
import {
  DEVICE_SELECT,
  DEVICES_PATH,
  deviceSummary,
  INTUNE_ADMIN_ROLES,
  INTUNE_READ_ROLES,
  listDevices,
  type ManagedDevice,
  odataType,
  pickByName,
  READ_DEVICES,
  READ_GROUPS,
  resolveManagedDevice,
  scan,
} from './common.ts';

export const READ_SERVICE_CONFIG = 'DeviceManagementServiceConfig.Read.All';
export const WRITE_SERVICE_CONFIG = 'DeviceManagementServiceConfig.ReadWrite.All';
export const AUTOPILOT_DEVICES_PATH = '/deviceManagement/windowsAutopilotDeviceIdentities';
const IMPORTS_PATH = '/deviceManagement/importedWindowsAutopilotDeviceIdentities';
const SETTINGS_PATH = '/deviceManagement/windowsAutopilotSettings';
const PROFILES_PATH = '/deviceManagement/windowsAutopilotDeploymentProfiles';
const EVENTS_PATH = '/deviceManagement/autopilotEvents';
const ALL = Number.POSITIVE_INFINITY;

const ENROLLMENT_STATES = ['unknown', 'enrolled', 'pendingReset', 'failed', 'notContacted', 'blocked'] as const;

export interface AutopilotDevice {
  id: string;
  serialNumber?: string | null;
  manufacturer?: string | null;
  model?: string | null;
  groupTag?: string | null;
  enrollmentState?: string | null;
  deploymentProfileAssignmentStatus?: string | null;
  deploymentProfileAssignmentDetailedStatus?: string | null;
  deploymentProfileAssignedDateTime?: string | null;
  userPrincipalName?: string | null;
  addressableUserName?: string | null;
  displayName?: string | null;
  lastContactedDateTime?: string | null;
  managedDeviceId?: string | null;
  azureAdDeviceId?: string | null;
  azureActiveDirectoryDeviceId?: string | null;
  [key: string]: unknown;
}

const lower = (value: string | null | undefined) => (value ?? '').toLowerCase();

/** Autopilot reports "no link" as the all-zero GUID. */
function realId(value: unknown): string | null {
  return typeof value === 'string' && value && !/^0{8}-0{4}-0{4}-0{4}-0{12}$/.test(value) ? value : null;
}

const entraId = (d: AutopilotDevice) => realId(d.azureAdDeviceId) ?? realId(d.azureActiveDirectoryDeviceId);

/**
 * Autopilot registrations linked to an Entra device id. The Autopilot service rejects `eq` filters on these ids
 * (only `contains(serialNumber,...)` works), so this pages client-side; tenants have hundreds, not thousands.
 */
export async function autopilotByEntraId(ctx: Context, deviceId: string): Promise<AutopilotDevice[]> {
  const wanted = deviceId.toLowerCase();
  const { items } = await scan<AutopilotDevice>(
    ctx,
    AUTOPILOT_DEVICES_PATH,
    {},
    (d) => lower(entraId(d)) === wanted,
    5,
    'beta',
  );
  return items;
}
const label = (d: AutopilotDevice) => d.serialNumber || d.id;

export function shapeAutopilotDevice(d: AutopilotDevice) {
  return compact({
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
    managedDeviceId: realId(d.managedDeviceId),
    azureAdDeviceId: entraId(d),
  });
}

function describeAutopilot(d: AutopilotDevice) {
  return {
    id: d.id,
    serialNumber: d.serialNumber ?? null,
    model: d.model ?? null,
    groupTag: d.groupTag || null,
    assignedUser: d.userPrincipalName || null,
    enrollmentState: d.enrollmentState ?? null,
  };
}

async function bySerial(ctx: Context, serial: string): Promise<AutopilotDevice[]> {
  const { items } = await ctx.graph.list<AutopilotDevice>(AUTOPILOT_DEVICES_PATH, {
    version: 'beta',
    query: { $filter: `contains(serialNumber,${odata(serial)})` },
    limit: 25,
  });
  return items;
}

/** The Autopilot record of an Intune device (by name, serial, Intune id or Entra id), matched on serial number. */
async function viaIntune(ctx: Context, ref: string): Promise<AutopilotDevice> {
  let managed: ManagedDevice;
  try {
    managed = await resolveManagedDevice(ctx, ref);
  } catch (error) {
    if (!(error instanceof CliError && error.code === 'NOT_FOUND')) throw error;
    throw new CliError('NOT_FOUND', `No Autopilot device matches "${ref}"`, {
      hint: 'Pass a serial number, or search with `tenantctl intune autopilot list --serial <part>`.',
    });
  }
  const serial = managed.serialNumber;
  const matches = serial ? (await bySerial(ctx, serial)).filter((d) => lower(d.serialNumber) === lower(serial)) : [];
  if (!serial || !matches.length) {
    throw new CliError(
      'NOT_FOUND',
      `Intune device ${managed.deviceName ?? managed.id} (serial ${serial ?? 'unknown'}) is not registered in Windows Autopilot`,
      { hint: 'Register it with `tenantctl intune autopilot import <csv>` using its hardware hash.' },
    );
  }
  return pickOne(matches, serial, 'Autopilot device', (d) => [d.serialNumber], describeAutopilot);
}

/**
 * Find one Windows Autopilot device identity from what a person would say: serial number (exact, case-insensitive;
 * for reads a unique partial serial also matches), Autopilot identity id, the linked Intune managed device id or
 * Entra device id, or the Intune device name (resolved to its serial). Several matches are AMBIGUOUS. Commands that
 * change a device pass `exact` so a partial serial never selects it.
 */
export async function resolveAutopilotDevice(
  ctx: Context,
  ref: string,
  options: { exact?: boolean } = {},
): Promise<AutopilotDevice> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A device is required');
  if (isGuid(value)) {
    const direct = await ctx.graph.tryGet<AutopilotDevice>(`${AUTOPILOT_DEVICES_PATH}/${seg(value)}`, {
      version: 'beta',
    });
    if (direct) return direct;
    const wanted = value.toLowerCase();
    const linked = await scan<AutopilotDevice>(
      ctx,
      AUTOPILOT_DEVICES_PATH,
      {},
      (d) => lower(realId(d.managedDeviceId)) === wanted || lower(entraId(d)) === wanted,
      10,
      'beta',
    );
    if (linked.items.length) {
      return pickOne(
        linked.items,
        value,
        'Autopilot device',
        (d) => [d.managedDeviceId, entraId(d)],
        describeAutopilot,
      );
    }
    return viaIntune(ctx, value);
  }
  const candidates = await bySerial(ctx, value);
  const exact = candidates.filter((d) => lower(d.serialNumber) === lower(value));
  if (exact.length) return pickOne(exact, value, 'Autopilot device', (d) => [d.serialNumber], describeAutopilot);
  if (candidates.length && !options.exact) {
    return pickOne(candidates, value, 'Autopilot device', (d) => [d.serialNumber], describeAutopilot);
  }
  if (candidates.length) {
    throw new CliError('USAGE', `"${value}" is only part of a serial number; changes need the whole serial`, {
      hint: `Matching serials: ${candidates.map((d) => d.serialNumber).join(', ')}`,
    });
  }
  return viaIntune(ctx, value);
}

/** Intune managed device records for an Autopilot device: its linked id, else any record with the same serial. */
async function intuneRecords(ctx: Context, device: AutopilotDevice): Promise<ManagedDevice[]> {
  const linked = realId(device.managedDeviceId);
  if (linked) {
    const direct = await ctx.graph.tryGet<ManagedDevice>(`${DEVICES_PATH}/${seg(linked)}`, {
      query: { $select: DEVICE_SELECT },
    });
    if (direct) return [direct];
  }
  const serial = device.serialNumber;
  if (!serial) return [];
  const { items } = await listDevices(
    ctx,
    [
      {
        filter: `serialNumber eq ${odata(serial)}`,
        documented: false,
        test: (d) => lower(d.serialNumber) === lower(serial),
      },
    ],
    10,
  );
  return items;
}

// ---------------------------------------------------------------------------------------------------------------
// Device update / delete

interface Change {
  field: string;
  before: string | null;
  after: string | null;
}

interface UpdatePlan {
  device: AutopilotDevice;
  changes: Change[];
  body: Record<string, string>;
  unassignUser: boolean;
  action: string;
}

interface UpdateArgs {
  device: string;
  groupTag?: string;
  clearGroupTag: boolean;
  user?: string;
  clearUser: boolean;
  displayName?: string;
}

const devicePath = (device: AutopilotDevice) => `${AUTOPILOT_DEVICES_PATH}/${seg(device.id)}`;

async function planUpdate(ctx: Context, args: UpdateArgs): Promise<UpdatePlan> {
  if (args.groupTag !== undefined && args.clearGroupTag) {
    throw new CliError('USAGE', 'Pass either --group-tag or --clear-group-tag, not both');
  }
  if (args.user !== undefined && args.clearUser) throw new CliError('USAGE', 'Pass either --user or --clear-user');
  if (args.groupTag === undefined && !args.clearGroupTag && !args.user && !args.clearUser && !args.displayName) {
    throw new CliError('USAGE', 'Nothing to change', {
      hint: 'Pass at least one of --group-tag, --clear-group-tag, --user, --clear-user or --display-name.',
    });
  }
  const [device, user] = await Promise.all([
    resolveAutopilotDevice(ctx, args.device, { exact: true }),
    args.user ? resolveUser(ctx, args.user) : Promise.resolve(undefined),
  ]);
  const changes: Change[] = [];
  const body: Record<string, string> = {};
  let unassignUser = false;
  if (args.groupTag !== undefined || args.clearGroupTag) {
    const next = args.clearGroupTag ? '' : (args.groupTag ?? '').trim();
    if ((device.groupTag ?? '') !== next) {
      body.groupTag = next;
      changes.push({ field: 'groupTag', before: device.groupTag || null, after: next || null });
    }
  }
  if (user && lower(device.userPrincipalName) !== lower(user.userPrincipalName)) {
    body.userPrincipalName = user.userPrincipalName;
    body.addressableUserName = user.displayName;
    changes.push({ field: 'assignedUser', before: device.userPrincipalName || null, after: user.userPrincipalName });
  }
  if (args.clearUser && device.userPrincipalName) {
    unassignUser = true;
    changes.push({ field: 'assignedUser', before: device.userPrincipalName, after: null });
  }
  if (args.displayName !== undefined && (device.displayName ?? '') !== args.displayName) {
    body.displayName = args.displayName;
    changes.push({ field: 'displayName', before: device.displayName || null, after: args.displayName });
  }
  const describe = (c: Change) => `${c.field} from ${c.before ?? '(none)'} to ${c.after ?? '(none)'}`;
  const action = changes.length
    ? `Change ${changes.map(describe).join(', ')} on Autopilot device ${label(device)}`
    : `No change: Autopilot device ${label(device)} already has these values`;
  return { device, changes, body, unassignUser, action };
}

function updateRequests(plan: UpdatePlan) {
  return [
    ...(plan.unassignUser
      ? [{ method: 'POST', path: `${devicePath(plan.device)}/unassignUserFromDevice`, version: 'v1.0' }]
      : []),
    ...(Object.keys(plan.body).length
      ? [
          {
            method: 'POST',
            path: `${devicePath(plan.device)}/updateDeviceProperties`,
            version: 'v1.0',
            body: plan.body,
          },
        ]
      : []),
  ];
}

function deleteAction(device: AutopilotDevice): string {
  const what = [device.manufacturer, device.model].filter(Boolean).join(' ');
  const details = [what, device.groupTag ? `group tag ${device.groupTag}` : null].filter(Boolean).join(', ');
  const user = device.userPrincipalName ? `, unassigning ${device.userPrincipalName} first` : '';
  return (
    `Deregister ${label(device)}${details ? ` (${details})` : ''} from Windows Autopilot${user}. It stops getting ` +
    'the Autopilot experience at its next reset; registering it again needs its hardware hash.'
  );
}

/** Microsoft's deregistration order: delete the Intune record first, then the Autopilot registration. */
async function planDelete(ctx: Context, ref: string) {
  const device = await resolveAutopilotDevice(ctx, ref, { exact: true });
  const records = await intuneRecords(ctx, device);
  if (records.length) {
    const first = records[0] as ManagedDevice;
    throw new CliError(
      'USAGE',
      `${label(device)} still has an Intune device record (${first.deviceName ?? first.id}); Microsoft requires ` +
        'deleting it from Intune before deregistering from Autopilot',
      {
        hint: `Run \`tenantctl intune devices delete ${first.id}\` first, then rerun this command.`,
        details: records.map(deviceSummary),
      },
    );
  }
  const requests = [
    ...(device.userPrincipalName
      ? [{ method: 'POST', path: `${devicePath(device)}/unassignUserFromDevice`, version: 'v1.0' }]
      : []),
    { method: 'DELETE', path: devicePath(device), version: 'v1.0' },
  ];
  return { device, requests, action: deleteAction(device) };
}

// ---------------------------------------------------------------------------------------------------------------
// Import

interface ImportRow {
  line: number;
  serialNumber: string;
  productKey: string;
  hardwareHash: string;
  groupTag: string;
  assignedUser: string;
}

const BASE64 = /^[A-Za-z0-9+/]+={0,2}$/;

/**
 * Rows from a `Get-WindowsAutopilotInfo` CSV. Every row is checked before anything is sent; any problem refuses the
 * whole file with one entry per bad row.
 */
export function readImportCsv(path: string, groupTag?: string): ImportRow[] {
  const rows = readCsvRows(path, ['Device Serial Number', 'Hardware Hash']);
  const problems: { line: number; serialNumber: string | null; problem: string }[] = [];
  const seen = new Map<string, number>();
  const out: ImportRow[] = [];
  for (const row of rows) {
    const serial = row.get('Device Serial Number') ?? '';
    const hash = (row.get('Hardware Hash') ?? '').replace(/\s+/g, '');
    const user = row.get('Assigned User') ?? '';
    const report = (problem: string) => problems.push({ line: row.line, serialNumber: serial || null, problem });
    if (!serial) report('Device Serial Number is empty');
    if (!hash) report('Hardware Hash is empty');
    else if (!BASE64.test(hash) || hash.length % 4 !== 0) report('Hardware Hash is not base64');
    if (user && !user.includes('@')) report(`Assigned User "${user}" is not a UPN`);
    if (serial) {
      const first = seen.get(serial.toLowerCase());
      if (first !== undefined) report(`serial ${serial} also appears on line ${first}`);
      else seen.set(serial.toLowerCase(), row.line);
    }
    out.push({
      line: row.line,
      serialNumber: serial,
      productKey: row.get('Windows Product ID') ?? '',
      hardwareHash: hash,
      groupTag: groupTag?.trim() || row.get('Group Tag') || '',
      assignedUser: user,
    });
  }
  if (problems.length) {
    throw new CliError(
      'USAGE',
      `${path}: ${problems.length} problem${problems.length > 1 ? 's' : ''}; nothing was imported`,
      {
        hint: 'Fix the listed rows (re-export the hash with Get-WindowsAutopilotInfo if it is damaged) and rerun.',
        details: problems,
      },
    );
  }
  return out;
}

/** Digest of the validated rows: approval of an import binds to it, so a file edited after the preview is refused. */
function importFingerprint(rows: ImportRow[]): string {
  return createHash('sha256').update(JSON.stringify(rows)).digest('hex');
}

function importBody(rows: ImportRow[]) {
  return {
    importedWindowsAutopilotDeviceIdentities: rows.map((row) => ({
      '@odata.type': '#microsoft.graph.importedWindowsAutopilotDeviceIdentity',
      groupTag: row.groupTag,
      serialNumber: row.serialNumber,
      productKey: row.productKey,
      hardwareIdentifier: row.hardwareHash,
      assignedUserPrincipalName: row.assignedUser,
      state: {
        '@odata.type': 'microsoft.graph.importedWindowsAutopilotDeviceIdentityState',
        deviceImportStatus: 'pending',
        deviceRegistrationId: '',
        deviceErrorCode: 0,
        deviceErrorName: '',
      },
    })),
  };
}

interface ImportedIdentity {
  id: string;
  serialNumber?: string | null;
  groupTag?: string | null;
  importId?: string | null;
  assignedUserPrincipalName?: string | null;
  state?: {
    deviceImportStatus?: string | null;
    deviceRegistrationId?: string | null;
    deviceErrorCode?: number | null;
    deviceErrorName?: string | null;
  } | null;
  [key: string]: unknown;
}

function shapeImported(item: ImportedIdentity) {
  const state = item.state ?? {};
  const failed = state.deviceImportStatus === 'error';
  return compact({
    id: item.id,
    serialNumber: item.serialNumber ?? null,
    groupTag: item.groupTag || null,
    assignedUser: item.assignedUserPrincipalName || null,
    importId: item.importId || undefined,
    status: state.deviceImportStatus ?? null,
    registrationId: state.deviceRegistrationId || undefined,
    errorCode: failed || state.deviceErrorCode ? (state.deviceErrorCode ?? null) : undefined,
    errorName: state.deviceErrorName || undefined,
  });
}

function countBy<T>(items: T[], key: (item: T) => string): Record<string, number> {
  const out: Record<string, number> = {};
  for (const item of items) out[key(item)] = (out[key(item)] ?? 0) + 1;
  return out;
}

// ---------------------------------------------------------------------------------------------------------------
// Deployment profiles

interface OobeSettings {
  privacySettingsHidden?: boolean | null;
  eulaHidden?: boolean | null;
  userType?: string | null;
  deviceUsageType?: string | null;
  keyboardSelectionPageSkipped?: boolean | null;
  escapeLinkHidden?: boolean | null;
  // Deprecated (outOfBoxExperienceSettings) names.
  hidePrivacySettings?: boolean | null;
  hideEULA?: boolean | null;
  skipKeyboardSelectionPage?: boolean | null;
  hideEscapeLink?: boolean | null;
}

interface DeploymentProfile {
  id: string;
  displayName?: string | null;
  description?: string | null;
  language?: string | null;
  locale?: string | null;
  deviceNameTemplate?: string | null;
  deviceType?: string | null;
  hardwareHashExtractionEnabled?: boolean | null;
  extractHardwareHash?: boolean | null;
  preprovisioningAllowed?: boolean | null;
  enableWhiteGlove?: boolean | null;
  outOfBoxExperienceSetting?: OobeSettings | null;
  outOfBoxExperienceSettings?: OobeSettings | null;
  assignments?: Assignment[];
  [key: string]: unknown;
}

function joinType(profile: Record<string, unknown>): string | null {
  const type = odataType(profile);
  if (type === 'azureADWindowsAutopilotDeploymentProfile') return 'entraJoined';
  if (type === 'activeDirectoryWindowsAutopilotDeploymentProfile') return 'hybridJoined';
  return type;
}

const convertsAll = (p: DeploymentProfile) => p.hardwareHashExtractionEnabled ?? p.extractHardwareHash ?? null;

function oobe(p: DeploymentProfile) {
  const now = p.outOfBoxExperienceSetting ?? {};
  const old = p.outOfBoxExperienceSettings ?? {};
  return {
    userType: now.userType ?? old.userType ?? null,
    deviceUsageType: now.deviceUsageType ?? old.deviceUsageType ?? null,
    privacySettingsHidden: now.privacySettingsHidden ?? old.hidePrivacySettings ?? null,
    eulaHidden: now.eulaHidden ?? old.hideEULA ?? null,
    keyboardSelectionPageSkipped: now.keyboardSelectionPageSkipped ?? old.skipKeyboardSelectionPage ?? null,
    escapeLinkHidden: now.escapeLinkHidden ?? old.hideEscapeLink ?? null,
  };
}

function shapeProfile(p: DeploymentProfile, names: Record<string, string | null>, detail = false) {
  const settings = oobe(p);
  const hybrid = joinType(p) === 'hybridJoined';
  return compact({
    id: p.id,
    name: p.displayName ?? null,
    description: detail ? p.description || null : undefined,
    joinType: joinType(p),
    // Intune stores self-deploying mode as a shared-device profile.
    deploymentMode: settings.deviceUsageType === 'shared' ? 'selfDeploying' : 'userDriven',
    userType: settings.userType,
    deviceNameTemplate: p.deviceNameTemplate || null,
    language: p.locale || p.language || null,
    convertAllTargetedDevices: convertsAll(p),
    preProvisioningAllowed: p.preprovisioningAllowed ?? p.enableWhiteGlove ?? null,
    deviceType: p.deviceType ?? undefined,
    oobe: detail ? settings : undefined,
    skipDomainConnectivityCheck: detail && hybrid ? (p.hybridAzureADJoinSkipConnectivityCheck ?? null) : undefined,
    enrollmentStatusScreen: detail ? (p.enrollmentStatusScreenSettings ?? null) : undefined,
    createdDateTime: detail ? (p.createdDateTime ?? null) : undefined,
    lastModifiedDateTime: p.lastModifiedDateTime ?? null,
    assignments: p.assignments ? p.assignments.map((a) => shapeAssignment(a, names)) : undefined,
  });
}

async function listProfiles(ctx: Context, limit: number, expand = true): Promise<ListResult<DeploymentProfile>> {
  return ctx.graph.list<DeploymentProfile>(PROFILES_PATH, {
    version: 'beta',
    query: expand ? { $expand: 'assignments' } : { $select: 'id,displayName' },
    limit,
  });
}

async function resolveProfile(ctx: Context, ref: string): Promise<DeploymentProfile> {
  const { items } = await listProfiles(ctx, ALL);
  return pickByName(items, ref, 'Autopilot deployment profile', (p) => p.displayName);
}

async function profileAssignments(ctx: Context, profileId: string): Promise<Assignment[]> {
  const { items } = await ctx.graph.list<Assignment>(`${PROFILES_PATH}/${seg(profileId)}/assignments`, {
    version: 'beta',
    limit: ALL,
    pageSize: false,
  });
  return items;
}

const TARGET_TYPES = {
  group: '#microsoft.graph.groupAssignmentTarget',
  exclude: '#microsoft.graph.exclusionGroupAssignmentTarget',
} as const;

type TargetKind = keyof typeof TARGET_TYPES;

interface ProfileTarget {
  kind: TargetKind;
  group: ResolvedGroup;
  memberCount: number | null;
}

function kindOf(assignment: Assignment): TargetKind | null {
  const type = String(assignment.target?.['@odata.type'] ?? '').replace(/^#?microsoft\.graph\./, '');
  if (type === 'groupAssignmentTarget') return 'group';
  if (type === 'exclusionGroupAssignmentTarget') return 'exclude';
  return null;
}

const describeTarget = (t: ProfileTarget) =>
  `members of ${t.group.displayName}${typeof t.memberCount === 'number' ? ` (${t.memberCount} members)` : ''}`;

interface ProfilePlan {
  profile: DeploymentProfile;
  targets: ProfileTarget[];
  current: Assignment[];
  add: ProfileTarget[];
  remove: Assignment[];
  action: string;
}

async function resolveProfileTargets(
  ctx: Context,
  args: { group?: string; exclude?: string },
  mode: 'assign' | 'unassign',
): Promise<ProfileTarget[]> {
  const count = (args.group ? 1 : 0) + (args.exclude ? 1 : 0);
  if (!count || (mode === 'unassign' && count > 1)) {
    throw new CliError(
      'USAGE',
      mode === 'assign'
        ? 'Pass --group <group>, --exclude <group> or both'
        : 'Pass exactly one of --group or --exclude',
    );
  }
  const target = async (ref: string, kind: TargetKind): Promise<ProfileTarget> => {
    const group = await resolveGroup(ctx, ref);
    return { kind, group, memberCount: await memberCount(ctx, group.id) };
  };
  return Promise.all([
    ...(args.group ? [target(args.group, 'group')] : []),
    ...(args.exclude ? [target(args.exclude, 'exclude')] : []),
  ]);
}

async function planProfile(
  ctx: Context,
  args: { deploymentProfile: string; group?: string; exclude?: string },
  mode: 'assign' | 'unassign',
): Promise<ProfilePlan> {
  const [profile, targets] = await Promise.all([
    resolveProfile(ctx, args.deploymentProfile),
    resolveProfileTargets(ctx, args, mode),
  ]);
  const current = await profileAssignments(ctx, profile.id);
  const name = `Autopilot profile "${profile.displayName ?? profile.id}"`;
  const sameGroup = (a: Assignment, t: ProfileTarget) => a.target?.groupId === t.group.id;
  if (mode === 'unassign') {
    const target = targets[0] as ProfileTarget;
    const remove = current.filter((a) => sameGroup(a, target) && kindOf(a) === target.kind);
    if (!remove.length) {
      throw new CliError(
        'NOT_FOUND',
        `${name} is not ${target.kind === 'exclude' ? 'excluding' : 'assigned to'} ${target.group.displayName}`,
        {
          hint: `See its assignments with \`tenantctl intune autopilot profiles get ${shellQuote(profile.displayName ?? profile.id)}\`.`,
        },
      );
    }
    const action =
      target.kind === 'exclude'
        ? `Stop excluding ${describeTarget(target)} from ${name}`
        : `Remove the assignment of ${name} from ${describeTarget(target)}`;
    return { profile, targets, current, add: [], remove, action };
  }
  const add: ProfileTarget[] = [];
  for (const target of targets) {
    const existing = current.find((a) => sameGroup(a, target));
    if (!existing) add.push(target);
    else if (kindOf(existing) !== target.kind) {
      throw new CliError(
        'USAGE',
        `${target.group.displayName} is already ${kindOf(existing) === 'exclude' ? 'excluded from' : 'assigned to'} ${name}`,
        {
          hint: `Remove that first with \`tenantctl intune autopilot profiles unassign ${shellQuote(profile.displayName ?? profile.id)} --${kindOf(existing) === 'exclude' ? 'exclude' : 'group'} ${shellQuote(target.group.displayName)}\`.`,
        },
      );
    }
  }
  const include = targets.find((t) => t.kind === 'group');
  const exclude = targets.find((t) => t.kind === 'exclude');
  let action = include
    ? `Assign ${name} to ${describeTarget(include)}`
    : `Exclude ${describeTarget(exclude as ProfileTarget)} from ${name}`;
  if (include && exclude) action += `, excluding ${describeTarget(exclude)}`;
  if (!add.length) action = `Already in place: ${action}`;
  return { profile, targets, current, add, remove: [], action };
}

const assignmentPath = (profile: DeploymentProfile) => `${PROFILES_PATH}/${seg(profile.id)}/assignments`;

const newAssignment = (t: ProfileTarget) => ({
  '@odata.type': '#microsoft.graph.windowsAutopilotDeploymentProfileAssignment',
  target: { '@odata.type': TARGET_TYPES[t.kind], groupId: t.group.id },
});

async function describeProfilePlan(ctx: Context, plan: ProfilePlan) {
  const result: Assignment[] = [
    ...plan.current.filter((a) => !plan.remove.includes(a)),
    ...plan.add.map((t) => newAssignment(t) as Assignment),
  ];
  const known: Record<string, string> = {};
  for (const t of plan.targets) known[t.group.id] = t.group.displayName;
  const names = { ...(await assignmentGroupNames(ctx, plan.current)), ...known };
  const converting = convertsAll(plan.profile) && plan.add.some((t) => t.kind === 'group');
  return compact({
    action: plan.action,
    profile: {
      id: plan.profile.id,
      name: plan.profile.displayName ?? null,
      joinType: joinType(plan.profile),
      convertAllTargetedDevices: convertsAll(plan.profile),
    },
    targets: plan.targets.map((t) => ({
      target: t.kind,
      group: t.group.displayName,
      groupId: t.group.id,
      memberCount: t.memberCount,
    })),
    warning: converting
      ? 'This profile converts all targeted devices: Intune-managed Windows devices in the group get their hardware ' +
        'hash uploaded and are registered in Windows Autopilot.'
      : undefined,
    changed: plan.add.length + plan.remove.length > 0,
    current: plan.current.map((a) => shapeAssignment(a, names)),
    result: result.map((a) => shapeAssignment(a, names)),
    requests: [
      ...plan.add.map((t) => ({
        method: 'POST',
        path: assignmentPath(plan.profile),
        version: 'beta',
        body: newAssignment(t),
      })),
      ...plan.remove.map((a) => ({
        method: 'DELETE',
        path: `${assignmentPath(plan.profile)}/${a.id}`,
        version: 'beta',
      })),
    ],
  });
}

const PROFILE_PERMISSIONS = {
  delegated: [WRITE_SERVICE_CONFIG, READ_GROUPS],
  application: [WRITE_SERVICE_CONFIG, READ_GROUPS],
  roles: INTUNE_ADMIN_ROLES,
};

// ---------------------------------------------------------------------------------------------------------------
// Deployment events

interface AutopilotEvent {
  id: string;
  deviceId?: string | null;
  eventDateTime?: string | null;
  deviceSerialNumber?: string | null;
  managedDeviceName?: string | null;
  userPrincipalName?: string | null;
  windowsAutopilotDeploymentProfileDisplayName?: string | null;
  windows10EnrollmentCompletionPageConfigurationDisplayName?: string | null;
  enrollmentState?: string | null;
  enrollmentType?: string | null;
  deploymentState?: string | null;
  deviceSetupStatus?: string | null;
  accountSetupStatus?: string | null;
  enrollmentFailureDetails?: string | null;
  deploymentDuration?: string | null;
  deploymentTotalDuration?: string | null;
  deviceSetupDuration?: string | null;
  accountSetupDuration?: string | null;
  osVersion?: string | null;
  [key: string]: unknown;
}

/** ISO 8601 duration (`PT1H2M30S`, `P1DT2H`) -> minutes with one decimal; null when absent or unparseable. */
export function durationMinutes(value: string | null | undefined): number | null {
  const match = value?.match(/^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/);
  if (!match) return null;
  const [, d, h, m, s] = match;
  const seconds = Number(d ?? 0) * 86_400 + Number(h ?? 0) * 3600 + Number(m ?? 0) * 60 + Number(s ?? 0);
  return Math.round((seconds / 60) * 10) / 10;
}

function eventFailed(e: AutopilotEvent): boolean {
  return (
    e.deploymentState === 'failure' ||
    e.deviceSetupStatus === 'failure' ||
    e.accountSetupStatus === 'failure' ||
    e.enrollmentState === 'failed'
  );
}

function shapeEvent(e: AutopilotEvent) {
  const failedPhase =
    e.deviceSetupStatus === 'failure' ? 'deviceSetup' : e.accountSetupStatus === 'failure' ? 'accountSetup' : undefined;
  return compact({
    time: e.eventDateTime ?? null,
    device: e.managedDeviceName || null,
    serialNumber: e.deviceSerialNumber || null,
    user: e.userPrincipalName || null,
    deploymentState: e.deploymentState ?? null,
    enrollmentState: e.enrollmentState ?? null,
    deviceSetupStatus: e.deviceSetupStatus ?? null,
    accountSetupStatus: e.accountSetupStatus ?? null,
    failedPhase,
    failureDetails: e.enrollmentFailureDetails || undefined,
    totalMinutes: durationMinutes(e.deploymentTotalDuration),
    deploymentMinutes: durationMinutes(e.deploymentDuration),
    deviceSetupMinutes: durationMinutes(e.deviceSetupDuration),
    accountSetupMinutes: durationMinutes(e.accountSetupDuration),
    profile: e.windowsAutopilotDeploymentProfileDisplayName || null,
    enrollmentStatusPage: e.windows10EnrollmentCompletionPageConfigurationDisplayName || null,
    enrollmentType: e.enrollmentType ?? null,
    osVersion: e.osVersion || null,
    managedDeviceId: realId(e.deviceId),
  });
}

// ---------------------------------------------------------------------------------------------------------------

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'autopilot list',
    summary: 'List Windows Autopilot devices with group tag, enrollment state and profile status',
    description:
      'Use when a new laptop does not get the expected Autopilot experience: profileStatus must be ' +
      '"assignedInSync" before the device is reset/turned on. enrollmentState "notContacted" means it has ' +
      'never enrolled. --serial matches part of the serial (server-side); --user (part of the assigned UPN), ' +
      '--group-tag (exact, case-insensitive) and --enrollment-state filter client-side. Details of one device: ' +
      '`tenantctl intune autopilot get <serial>`. (beta API, needed for the deployment profile status.)',
    examples: [
      'tenantctl intune autopilot list --serial 5CG1234',
      'tenantctl intune autopilot list --user jane@contoso.com',
      'tenantctl intune autopilot list --group-tag Sales --enrollment-state notContacted',
    ],
    args: [
      { name: 'serial', type: 'string', description: 'Serial number or part of it.' },
      { name: 'user', type: 'string', description: 'Assigned user UPN or part of it.' },
      { name: 'groupTag', type: 'string', description: 'Group tag (exact, case-insensitive).' },
      {
        name: 'enrollmentState',
        type: 'string',
        choices: ENROLLMENT_STATES,
        description: 'Only devices in this Intune enrollment state.',
      },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_SERVICE_CONFIG], application: [READ_SERVICE_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const path = AUTOPILOT_DEVICES_PATH;
      const query = { $filter: args.serial ? `contains(serialNumber,${odata(args.serial)})` : undefined };
      const user = args.user?.toLowerCase();
      const tag = args.groupTag?.trim().toLowerCase();
      const state = args.enrollmentState;
      const keep = (d: AutopilotDevice) =>
        (!user || `${d.userPrincipalName ?? ''} ${d.addressableUserName ?? ''}`.toLowerCase().includes(user)) &&
        (tag === undefined || lower(d.groupTag) === tag) &&
        (!state || d.enrollmentState === state);
      const result =
        user || tag !== undefined || state
          ? await scan<AutopilotDevice>(ctx, path, query, keep, listLimit(args), 'beta')
          : await ctx.graph.list<AutopilotDevice>(path, { version: 'beta', query, limit: listLimit(args) });
      if (ctx.full) return result;
      return { items: result.items.map(shapeAutopilotDevice), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot get',
    summary: 'Show one Autopilot device with its deployment profile and linked Intune device',
    description:
      'Accepts a serial number, Autopilot device id, Intune device id, Entra device id or Intune device name. ' +
      'deploymentProfile is the profile the device has; intendedDeploymentProfile is the one Intune is moving it to ' +
      '(they differ while profileStatus is not assignedInSync). intuneDevice is null when the device has never ' +
      'enrolled. Change group tag/user with `tenantctl intune autopilot update`; see deployments with ' +
      '`tenantctl intune autopilot events <serial>`. (beta API, needed for the profile fields.)',
    examples: ['tenantctl intune autopilot get 5CG1234XYZ', 'tenantctl intune autopilot get DESKTOP-123'],
    args: [
      {
        name: 'device',
        type: 'string',
        required: true,
        positional: true,
        description: 'Serial number, Autopilot id, Intune or Entra device id, or Intune device name.',
      },
    ],
    risk: 'read',
    permissions: {
      delegated: [READ_SERVICE_CONFIG, READ_DEVICES],
      application: [READ_SERVICE_CONFIG, READ_DEVICES],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const device = await resolveAutopilotDevice(ctx, args.device);
      const sections = await settle({
        expanded: ctx.graph.get<AutopilotDevice & Record<string, DeploymentProfile | null>>(
          `${AUTOPILOT_DEVICES_PATH}/${seg(device.id)}`,
          { version: 'beta', query: { $expand: 'deploymentProfile,intendedDeploymentProfile' } },
        ),
        intune: intuneRecords(ctx, device),
      });
      const expanded = 'error' in sections.expanded ? null : sections.expanded;
      const full = expanded ?? device;
      if (ctx.full) return { device: full, intuneDevices: sections.intune };
      const profile = (key: 'deploymentProfile' | 'intendedDeploymentProfile') => {
        if ('error' in sections.expanded) return sections.expanded;
        const p = sections.expanded[key] as DeploymentProfile | null | undefined;
        return p ? { id: p.id, name: p.displayName ?? null, joinType: joinType(p) } : null;
      };
      const intune = sections.intune;
      return {
        device: compact({
          ...shapeAutopilotDevice(full),
          displayName: full.displayName || null,
          profileDetailedStatus:
            full.deploymentProfileAssignmentDetailedStatus && full.deploymentProfileAssignmentDetailedStatus !== 'none'
              ? full.deploymentProfileAssignmentDetailedStatus
              : undefined,
          purchaseOrderIdentifier: full.purchaseOrderIdentifier || undefined,
          addressableUserName: full.addressableUserName || undefined,
          remediationState: full.remediationState ?? undefined,
          userlessEnrollmentStatus: full.userlessEnrollmentStatus ?? undefined,
        }),
        deploymentProfile: profile('deploymentProfile'),
        intendedDeploymentProfile: profile('intendedDeploymentProfile'),
        intuneDevice: 'error' in intune ? intune : intune[0] ? deviceSummary(intune[0]) : null,
        ...(Array.isArray(intune) && intune.length > 1
          ? { otherIntuneRecords: intune.slice(1).map(deviceSummary) }
          : {}),
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot update',
    summary: 'Change an Autopilot device’s group tag, assigned user or device name',
    description:
      'A group tag change can move the device into another dynamic device group (rules on ' +
      'device.devicePhysicalIds -any _ -eq "[OrderID]:<tag>") and so to another deployment profile; the device only ' +
      'applies it at its next reset/OOBE, after profileStatus returns to assignedInSync (check with ' +
      '`tenantctl intune autopilot get <serial>`). --clear-group-tag and --clear-user remove the value ' +
      '(--clear-user uses the unassignUserFromDevice action). --display-name is the name applied during Autopilot ' +
      '(ignored by hybrid join). Values that already match are skipped; the preview lists before/after values.',
    examples: [
      'tenantctl intune autopilot update 5CG1234XYZ --group-tag Sales',
      'tenantctl intune autopilot update 5CG1234XYZ --user jane@contoso.com --display-name SALES-JANE',
      'tenantctl intune autopilot update 5CG1234XYZ --clear-group-tag --clear-user',
    ],
    args: [
      {
        name: 'device',
        type: 'string',
        required: true,
        positional: true,
        description: 'Serial number, Autopilot id, Intune or Entra device id, or Intune device name.',
      },
      { name: 'groupTag', type: 'string', description: 'New group tag.' },
      { name: 'clearGroupTag', type: 'boolean', description: 'Remove the group tag.' },
      { name: 'user', type: 'string', description: 'User to assign (UPN, mail, name or id).' },
      { name: 'clearUser', type: 'boolean', description: 'Remove the assigned user.' },
      { name: 'displayName', type: 'string', description: 'Device name applied during Autopilot.' },
    ],
    risk: 'destructive',
    permissions: {
      delegated: [WRITE_SERVICE_CONFIG, READ_DEVICES, 'User.Read.All'],
      application: [WRITE_SERVICE_CONFIG, READ_DEVICES, 'User.Read.All'],
      roles: INTUNE_ADMIN_ROLES,
    },
    async preview(ctx, args) {
      const plan = await planUpdate(ctx, args);
      return {
        action: plan.action,
        targetId: plan.device.id,
        device: describeAutopilot(plan.device),
        changes: plan.changes,
        requests: updateRequests(plan),
      };
    },
    async run(ctx, args, approved) {
      const plan = await planUpdate(ctx, args);
      requireApprovedTarget(approved, plan.device.id, 'Autopilot device');
      const view = { action: plan.action, device: describeAutopilot(plan.device), changes: plan.changes };
      if (!plan.changes.length) return { status: 'unchanged', ...view };
      if (plan.unassignUser) {
        await ctx.graph.post(`${devicePath(plan.device)}/unassignUserFromDevice`, undefined);
      }
      if (Object.keys(plan.body).length) {
        await ctx.graph.post(`${devicePath(plan.device)}/updateDeviceProperties`, plan.body);
      }
      return {
        status: 'updated',
        ...view,
        next: `Check profileStatus with \`tenantctl intune autopilot get ${shellQuote(label(plan.device))}\`; a new group tag can take a while to change the assigned profile.`,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot delete',
    summary: 'Deregister a device from Windows Autopilot',
    description:
      'Microsoft requires deleting the Intune device record before deregistering, so this refuses while one ' +
      'exists (delete it with `tenantctl intune devices delete <device>` first). An assigned user is unassigned ' +
      'first, as the admin center does. Deregistration takes a few minutes; `tenantctl intune autopilot sync` speeds ' +
      'it up. Do not delete the Entra device object by hand for Entra-joined devices; for hybrid-joined devices ' +
      'delete the computer object in on-premises AD instead.',
    examples: ['tenantctl intune autopilot delete 5CG1234XYZ'],
    args: [
      {
        name: 'device',
        type: 'string',
        required: true,
        positional: true,
        description: 'Serial number, Autopilot id, Intune or Entra device id, or Intune device name.',
      },
    ],
    risk: 'destructive',
    permissions: {
      delegated: [WRITE_SERVICE_CONFIG, READ_DEVICES],
      application: [WRITE_SERVICE_CONFIG, READ_DEVICES],
      roles: INTUNE_ADMIN_ROLES,
    },
    async preview(ctx, args) {
      const plan = await planDelete(ctx, args.device);
      return {
        action: plan.action,
        targetId: plan.device.id,
        device: shapeAutopilotDevice(plan.device),
        intuneDevice: null,
        requests: plan.requests,
      };
    },
    async run(ctx, args, approved) {
      const plan = await planDelete(ctx, args.device);
      requireApprovedTarget(approved, plan.device.id, 'Autopilot device');
      if (plan.device.userPrincipalName) {
        await ctx.graph.post(`${devicePath(plan.device)}/unassignUserFromDevice`, undefined);
      }
      await ctx.graph.delete(devicePath(plan.device));
      return {
        status: 'deleted',
        action: plan.action,
        device: describeAutopilot(plan.device),
        next: `Run \`tenantctl intune autopilot sync\`, then \`tenantctl intune autopilot list --serial ${shellQuote(label(plan.device))}\` until it is gone.`,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot import',
    summary: 'Register devices in Windows Autopilot from a hardware hash CSV',
    description:
      'Reads the CSV Get-WindowsAutopilotInfo writes (columns "Device Serial Number", "Windows Product ID", ' +
      '"Hardware Hash", optional "Group Tag" and "Assigned User"; at most 500 rows). Every row is validated first ' +
      '(serial and base64 hash required, no duplicate serials, UPN-shaped user); any problem refuses the whole file ' +
      'with one entry per bad row. --group-tag sets the tag on every row. Returns the import records with status ' +
      '"pending": follow them with `tenantctl intune autopilot imports`, then run `tenantctl intune autopilot sync` ' +
      'and check `tenantctl intune autopilot list --group-tag <tag>`.',
    examples: [
      'tenantctl intune autopilot import ./AutopilotHWID.csv',
      'tenantctl intune autopilot import ./hashes.csv --group-tag Sales',
    ],
    args: [
      { name: 'csv', type: 'string', required: true, positional: true, description: 'Path to the hardware hash CSV.' },
      { name: 'groupTag', type: 'string', description: 'Group tag for every row (overrides the CSV column).' },
    ],
    risk: 'destructive',
    permissions: { delegated: [WRITE_SERVICE_CONFIG], application: [WRITE_SERVICE_CONFIG], roles: INTUNE_ADMIN_ROLES },
    async preview(_ctx, args) {
      const rows = readImportCsv(args.csv, args.groupTag);
      return {
        action:
          `Register ${rows.length} device${rows.length === 1 ? '' : 's'} in this tenant's Windows Autopilot ` +
          '(they join this tenant at their next reset; another tenant cannot register them meanwhile)',
        targetId: importFingerprint(rows),
        rows: rows.length,
        groupTags: countBy(rows, (r) => r.groupTag || '(none)'),
        devices: rows.map((r) =>
          compact({
            line: r.line,
            serialNumber: r.serialNumber,
            groupTag: r.groupTag || null,
            assignedUser: r.assignedUser || undefined,
          }),
        ),
        request: { method: 'POST', path: `${IMPORTS_PATH}/import`, version: 'v1.0' },
      };
    },
    async run(ctx, args, approved) {
      const rows = readImportCsv(args.csv, args.groupTag);
      // The file is read again here; refuse if it changed after the person approved its preview.
      requireApprovedTarget(approved, importFingerprint(rows), 'CSV file');
      const response = await ctx.graph.post<{ value?: ImportedIdentity[] }>(`${IMPORTS_PATH}/import`, importBody(rows));
      const imported = (response?.value ?? []).map(shapeImported);
      return {
        status: 'submitted',
        rows: rows.length,
        statuses: countBy(imported, (i) => String(i.status ?? 'unknown')),
        imported,
        next:
          'Import takes several minutes: track it with `tenantctl intune autopilot imports`, then run ' +
          '`tenantctl intune autopilot sync`.',
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot imports',
    summary: 'Track hardware hash imports: status and error per serial',
    description:
      'Lists the import records `tenantctl intune autopilot import` created. status pending/partial means the ' +
      'Autopilot service is still processing; complete means the device is registered (it appears in ' +
      '`tenantctl intune autopilot list` after `tenantctl intune autopilot sync`); error carries errorCode/errorName ' +
      'from the Autopilot device directory service (for example a serial already registered in this or another ' +
      'tenant). --failed-only keeps error rows.',
    examples: ['tenantctl intune autopilot imports', 'tenantctl intune autopilot imports --failed-only'],
    args: [{ name: 'failedOnly', type: 'boolean', description: 'Only rows whose import failed.' }, ...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: [READ_SERVICE_CONFIG], application: [READ_SERVICE_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const query = { $select: 'id,serialNumber,groupTag,importId,state,assignedUserPrincipalName' };
      const result = args.failedOnly
        ? await scan<ImportedIdentity>(
            ctx,
            IMPORTS_PATH,
            query,
            (i) => i.state?.deviceImportStatus === 'error',
            listLimit(args),
          )
        : await ctx.graph.list<ImportedIdentity>(IMPORTS_PATH, { query, limit: listLimit(args) });
      if (ctx.full) return result;
      return { items: result.items.map(shapeImported), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot sync',
    summary: 'Sync Windows Autopilot devices from the Autopilot service now',
    description:
      'Pulls newly imported or deregistered devices and profile assignment changes from the Autopilot deployment ' +
      'service instead of waiting for the periodic sync. Microsoft documents no rate limit for this action; the ' +
      'service may reject a sync while one is running, so check `tenantctl intune autopilot sync-status` before ' +
      'retrying. (beta API: the v1.0 windowsAutopilotSettings path is documented but answers "Resource not found".)',
    examples: ['tenantctl intune autopilot sync'],
    args: [],
    risk: 'write',
    permissions: { delegated: [WRITE_SERVICE_CONFIG], application: [WRITE_SERVICE_CONFIG], roles: INTUNE_ADMIN_ROLES },
    async preview(ctx) {
      const settings = await ctx.graph.get<Record<string, unknown>>(SETTINGS_PATH, { version: 'beta' });
      return {
        action: 'Start a Windows Autopilot sync',
        current: shapeSettings(settings),
        request: { method: 'POST', path: `${SETTINGS_PATH}/sync`, version: 'beta' },
      };
    },
    async run(ctx) {
      await ctx.graph.post(`${SETTINGS_PATH}/sync`, undefined, { version: 'beta' });
      return {
        status: 'requested',
        next: 'The sync takes a few minutes; follow it with `tenantctl intune autopilot sync-status`.',
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot sync-status',
    summary: 'Show when Windows Autopilot devices last synced and whether a sync is running',
    description:
      'syncStatus inProgress means a sync is running; completed with a recent lastSyncDateTime means new imports ' +
      'and deregistrations are visible in `tenantctl intune autopilot list`. Start one with ' +
      '`tenantctl intune autopilot sync`. (beta API: the v1.0 path answers "Resource not found".)',
    examples: ['tenantctl intune autopilot sync-status'],
    args: [],
    risk: 'read',
    permissions: { delegated: [READ_SERVICE_CONFIG], application: [READ_SERVICE_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx) {
      const settings = await ctx.graph.get<Record<string, unknown>>(SETTINGS_PATH, { version: 'beta' });
      return ctx.full ? settings : shapeSettings(settings);
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot profiles list',
    summary: 'List Autopilot deployment profiles with join type, mode, naming and assignments',
    description:
      'joinType entraJoined/hybridJoined comes from the profile type; deploymentMode selfDeploying is a ' +
      'shared-device profile; userType administrator makes the first user a local admin. convertAllTargetedDevices ' +
      'true registers every targeted Intune-managed Windows device in Autopilot. Assignments show group names. ' +
      'Details: `tenantctl intune autopilot profiles get <name>`; target a group with ' +
      '`tenantctl intune autopilot profiles assign`. (beta API: deployment profiles are not in v1.0.)',
    examples: ['tenantctl intune autopilot profiles list'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: {
      delegated: [READ_SERVICE_CONFIG, READ_GROUPS],
      application: [READ_SERVICE_CONFIG, READ_GROUPS],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const result = await listProfiles(ctx, listLimit(args));
      if (ctx.full) return result;
      const names = await assignmentGroupNames(
        ctx,
        result.items.flatMap((p) => p.assignments ?? []),
      );
      return { items: result.items.map((p) => shapeProfile(p, names)), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot profiles get',
    summary: 'Show one Autopilot deployment profile: OOBE settings, assignments and assigned device count',
    description:
      'oobe lists what the user sees (privacy/EULA pages hidden, keyboard page skipped, "start over" link hidden). ' +
      'assignedDeviceCount counts Autopilot devices that have this profile. Change targeting with ' +
      '`tenantctl intune autopilot profiles assign|unassign`; see the devices with ' +
      '`tenantctl intune autopilot list`. (beta API.)',
    examples: ['tenantctl intune autopilot profiles get "User driven"'],
    args: [
      {
        name: 'deploymentProfile',
        type: 'string',
        required: true,
        positional: true,
        description: 'Deployment profile name or id.',
      },
    ],
    risk: 'read',
    permissions: {
      delegated: [READ_SERVICE_CONFIG, READ_GROUPS],
      application: [READ_SERVICE_CONFIG, READ_GROUPS],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const profile = await resolveProfile(ctx, args.deploymentProfile);
      const sections = await settle({
        assignments: profileAssignments(ctx, profile.id),
        devices: ctx.graph.list(`${PROFILES_PATH}/${seg(profile.id)}/assignedDevices`, { version: 'beta', limit: ALL }),
      });
      const assignments = 'error' in sections.assignments ? [] : sections.assignments;
      if (ctx.full) return { profile, assignments: sections.assignments, assignedDevices: sections.devices };
      const names = await assignmentGroupNames(ctx, assignments);
      return {
        ...shapeProfile({ ...profile, assignments }, names, true),
        ...('error' in sections.assignments ? { assignments: sections.assignments } : {}),
        assignedDeviceCount: 'error' in sections.devices ? sections.devices : sections.devices.items.length,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot profiles assign',
    summary: 'Assign an Autopilot deployment profile to a device group, or exclude a group',
    description:
      'Adds a group (--group) and/or an excluded group (--exclude) to the profile; other assignments stay. ' +
      'Already-assigned targets are reported as unchanged. The preview shows the group member count, the current ' +
      'and resulting assignments, and warns when the profile converts all targeted devices (existing Intune ' +
      'devices in the group get registered in Autopilot). Devices pick the profile up when profileStatus turns ' +
      "assignedInSync in `tenantctl intune autopilot list`. (beta API; uses the profile's documented assignments " +
      'relationship.)',
    examples: [
      'tenantctl intune autopilot profiles assign "User driven" --group "Autopilot Devices"',
      'tenantctl intune autopilot profiles assign "User driven" --exclude "Kiosks"',
    ],
    args: [
      {
        name: 'deploymentProfile',
        type: 'string',
        required: true,
        positional: true,
        description: 'Deployment profile name or id.',
      },
      { name: 'group', type: 'string', description: 'Device group to target: name, mail or id.' },
      { name: 'exclude', type: 'string', description: 'Device group to exclude: name, mail or id.' },
    ],
    risk: 'destructive',
    permissions: PROFILE_PERMISSIONS,
    async preview(ctx, args) {
      return describeProfilePlan(ctx, await planProfile(ctx, args, 'assign'));
    },
    async run(ctx, args) {
      const plan = await planProfile(ctx, args, 'assign');
      const view = await describeProfilePlan(ctx, plan);
      if (!plan.add.length) return { status: 'unchanged', ...view, requests: undefined };
      for (const target of plan.add) {
        await ctx.graph.post(assignmentPath(plan.profile), newAssignment(target), { version: 'beta' });
      }
      return {
        status: 'assigned',
        action: view.action,
        profile: view.profile,
        targets: view.targets,
        assignments: view.result,
        next: `Watch profileStatus in \`tenantctl intune autopilot list\`; details with \`tenantctl intune autopilot profiles get ${shellQuote(plan.profile.displayName ?? plan.profile.id)}\`.`,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot profiles unassign',
    summary: 'Remove a group (or an exclusion) from an Autopilot deployment profile',
    description:
      'Deletes only the assignment for the given group (--group) or exclusion (--exclude); the others stay. ' +
      'Devices that already have the profile keep it until another profile is assigned to them. (beta API.)',
    examples: [
      'tenantctl intune autopilot profiles unassign "User driven" --group "Autopilot Devices"',
      'tenantctl intune autopilot profiles unassign "User driven" --exclude "Kiosks"',
    ],
    args: [
      {
        name: 'deploymentProfile',
        type: 'string',
        required: true,
        positional: true,
        description: 'Deployment profile name or id.',
      },
      { name: 'group', type: 'string', description: 'Assigned group to remove: name, mail or id.' },
      { name: 'exclude', type: 'string', description: 'Excluded group to remove: name, mail or id.' },
    ],
    risk: 'destructive',
    permissions: PROFILE_PERMISSIONS,
    async preview(ctx, args) {
      return describeProfilePlan(ctx, await planProfile(ctx, args, 'unassign'));
    },
    async run(ctx, args) {
      const plan = await planProfile(ctx, args, 'unassign');
      const view = await describeProfilePlan(ctx, plan);
      for (const assignment of plan.remove) {
        await ctx.graph.delete(`${assignmentPath(plan.profile)}/${seg(String(assignment.id))}`, { version: 'beta' });
      }
      return {
        status: 'unassigned',
        action: view.action,
        profile: view.profile,
        targets: view.targets,
        assignments: view.result,
        next: `Check with \`tenantctl intune autopilot profiles get ${shellQuote(plan.profile.displayName ?? plan.profile.id)}\`.`,
      };
    },
  }),

  defineCommand({
    domain: 'intune',
    name: 'autopilot events',
    summary: 'Autopilot deployment history: state, failure details and phase durations per device',
    description:
      'Newest first. Reading a failure: deploymentState "failure" with failedPhase deviceSetup means something ' +
      'blocking in the Enrollment Status Page device phase (required apps, policies, certificates) failed or timed ' +
      'out; accountSetup means the user phase; failureDetails carries the error text. "successWithTimeout" means ' +
      'the ESP timeout was hit (see `tenantctl intune enrollment list --type enrollmentStatusPage`). Durations are ' +
      'minutes. Optional device matches the serial or Intune device name; --days limits to recent events. Next: ' +
      '`tenantctl intune autopilot get <serial>` or the Intune device with `tenantctl intune devices get <name>`. ' +
      '(beta API; no v1.0 equivalent.)',
    examples: [
      'tenantctl intune autopilot events --failed-only --days 7',
      'tenantctl intune autopilot events 5CG1234XYZ',
    ],
    args: [
      { name: 'device', type: 'string', positional: true, description: 'Serial number or Intune device name.' },
      { name: 'failedOnly', type: 'boolean', description: 'Only failed deployments.' },
      { name: 'days', type: 'number', description: 'Only events from the last N days.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_DEVICES], application: [READ_DEVICES], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      if (args.days !== undefined && !(args.days > 0)) throw new CliError('USAGE', '--days must be a positive number');
      const since = args.days !== undefined ? daysAgo(args.days) : undefined;
      const device = args.device?.trim().toLowerCase();
      const keep = (e: AutopilotEvent) =>
        (!device || lower(e.deviceSerialNumber) === device || lower(e.managedDeviceName) === device) &&
        (!args.failedOnly || eventFailed(e)) &&
        (!since || (e.eventDateTime ?? '') >= since);
      const load = (filter: string | undefined) =>
        scan<AutopilotEvent>(ctx, EVENTS_PATH, { $filter: filter }, keep, ALL, 'beta');
      let all: ListResult<AutopilotEvent>;
      try {
        // eventDateTime filtering works but is undocumented; on a 400 the date test stays client-side only.
        all = await load(since ? `eventDateTime ge ${since}` : undefined);
      } catch (error) {
        if (!(since && error instanceof CliError && error.status === 400)) throw error;
        all = await load(undefined);
      }
      const sorted = [...all.items].sort((a, b) => (b.eventDateTime ?? '').localeCompare(a.eventDateTime ?? ''));
      const limit = listLimit(args);
      const items = sorted.slice(0, limit);
      const truncated = sorted.length > items.length;
      if (ctx.full) return { items, truncated };
      return { items: items.map(shapeEvent), truncated };
    },
  }),
];

function shapeSettings(settings: Record<string, unknown>) {
  const last = typeof settings.lastSyncDateTime === 'string' ? settings.lastSyncDateTime : null;
  const real = last && !last.startsWith('0001-') ? last : null;
  const manual =
    typeof settings.lastManualSyncTriggerDateTime === 'string' &&
    !settings.lastManualSyncTriggerDateTime.startsWith('0001-')
      ? settings.lastManualSyncTriggerDateTime
      : null;
  return {
    syncStatus: settings.syncStatus ?? null,
    lastSyncDateTime: real,
    minutesSinceLastSync: real ? Math.floor((Date.now() - Date.parse(real)) / 60_000) : null,
    lastManualSyncTriggerDateTime: manual,
  };
}
