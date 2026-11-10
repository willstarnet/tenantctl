import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { pickOne, resolveUser } from '../../core/resolve.ts';
import { signinErrorInfo, signinErrorLookupUrl } from '../../core/signinErrors.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact, daysAgo, daysSince, isGuid, settle } from '../../core/util.ts';

/**
 * User lifecycle (create, update, disable, delete, restore, licenses, passwords) and per-user troubleshooting
 * (sign-ins, MFA methods, groups, devices). Directory-wide views (groups, audit, roles, CA, risk) live in
 * directory.ts.
 */

interface SignInActivity {
  lastSignInDateTime?: string | null;
  lastNonInteractiveSignInDateTime?: string | null;
  lastSuccessfulSignInDateTime?: string | null;
}

interface LicenseAssignmentState {
  skuId: string;
  assignedByGroup: string | null;
  state: string | null;
  error: string | null;
  disabledPlans: string[];
  lastUpdatedDateTime?: string | null;
}

export interface GraphUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
  accountEnabled?: boolean;
  userType?: string | null;
  department?: string | null;
  jobTitle?: string | null;
  officeLocation?: string | null;
  mobilePhone?: string | null;
  employeeId?: string | null;
  createdDateTime?: string | null;
  lastPasswordChangeDateTime?: string | null;
  onPremisesSyncEnabled?: boolean | null;
  onPremisesLastSyncDateTime?: string | null;
  usageLocation?: string | null;
  assignedLicenses?: { skuId: string; disabledPlans: string[] }[];
  licenseAssignmentStates?: LicenseAssignmentState[];
  signInActivity?: SignInActivity | null;
}

interface ServicePlan {
  servicePlanId: string;
  servicePlanName: string;
  provisioningStatus?: string;
}

interface SubscribedSku {
  skuId: string;
  skuPartNumber: string;
  appliesTo?: string;
  capabilityStatus?: string;
  consumedUnits: number;
  prepaidUnits?: { enabled?: number };
  servicePlans: ServicePlan[];
}

interface AuthMethod {
  '@odata.type': string;
  id: string;
  createdDateTime?: string | null;
  [key: string]: unknown;
}

export interface SignIn {
  id: string;
  createdDateTime: string;
  userPrincipalName?: string;
  appDisplayName?: string | null;
  resourceDisplayName?: string | null;
  clientAppUsed?: string | null;
  isInteractive?: boolean;
  ipAddress?: string | null;
  location?: { city?: string | null; state?: string | null; countryOrRegion?: string | null } | null;
  deviceDetail?: {
    deviceId?: string | null;
    displayName?: string | null;
    operatingSystem?: string | null;
    browser?: string | null;
    isCompliant?: boolean | null;
    isManaged?: boolean | null;
    trustType?: string | null;
  } | null;
  status?: { errorCode?: number; failureReason?: string | null; additionalDetails?: string | null } | null;
  conditionalAccessStatus?: string | null;
  authenticationRequirement?: string | null;
  riskLevelDuringSignIn?: string | null;
  correlationId?: string | null;
}

type Failed = { error: { code: string; message: string } };

function failed<T>(value: T | Failed): value is Failed {
  return typeof value === 'object' && value !== null && 'error' in value && Object.keys(value).length === 1;
}

const USER_ARG = {
  name: 'user',
  type: 'string',
  required: true,
  positional: true,
  description: 'User UPN, email, display name, object id, or "me".',
} as const;

const LIST_SELECT =
  'id,displayName,userPrincipalName,mail,accountEnabled,userType,department,jobTitle,createdDateTime,assignedLicenses';
const DETAIL_SELECT =
  'id,displayName,userPrincipalName,mail,accountEnabled,userType,department,jobTitle,officeLocation,mobilePhone,' +
  'employeeId,createdDateTime,lastPasswordChangeDateTime,onPremisesSyncEnabled,onPremisesLastSyncDateTime,' +
  'usageLocation,assignedLicenses';
const STATE_SELECT = 'id,displayName,userPrincipalName,mail,accountEnabled,onPremisesSyncEnabled,assignedLicenses';
const DEVICE_SELECT =
  'id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,isCompliant,isManaged,' +
  'approximateLastSignInDateTime,accountEnabled,registrationDateTime';
const DELETED_RETENTION_DAYS = 30;

const SYNCED_WARNING =
  'This account is synced from on-premises Active Directory: change it in AD (Graph rejects edits to ' +
  'on-premises-mastered properties) or the next sync overwrites it.';

function identity(user: { id: string; displayName: string; userPrincipalName: string }) {
  return { id: user.id, displayName: user.displayName, userPrincipalName: user.userPrincipalName };
}

export function shapeUser(user: GraphUser) {
  return compact({
    id: user.id,
    displayName: user.displayName,
    userPrincipalName: user.userPrincipalName,
    mail: user.mail,
    accountEnabled: user.accountEnabled,
    userType: user.userType,
    department: user.department,
    jobTitle: user.jobTitle,
    licensed: user.assignedLicenses ? user.assignedLicenses.length > 0 : undefined,
    lastSignIn: user.signInActivity !== undefined ? (user.signInActivity?.lastSignInDateTime ?? null) : undefined,
    createdDateTime: user.createdDateTime,
  });
}

/** Strong random password meeting Entra complexity rules: 16 chars, all four character classes. */
export function generatePassword(length = 16): string {
  const sets = ['ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnopqrstuvwxyz', '23456789', '!@#$%^&*-_=+?'];
  const all = sets.join('');
  const random = (max: number) => (crypto.getRandomValues(new Uint32Array(1))[0] as number) % max;
  const chars = sets.map((set) => set[random(set.length)] as string);
  while (chars.length < length) chars.push(all[random(all.length)] as string);
  // Shuffle so the guaranteed character classes are not always in the first four positions.
  for (let i = chars.length - 1; i > 0; i--) {
    const j = random(i + 1);
    [chars[i], chars[j]] = [chars[j] as string, chars[i] as string];
  }
  return chars.join('');
}

async function loadSkus(ctx: Context): Promise<SubscribedSku[]> {
  const { items } = await ctx.graph.list<SubscribedSku>('/subscribedSkus', {
    query: { $select: 'skuId,skuPartNumber,appliesTo,capabilityStatus,consumedUnits,prepaidUnits,servicePlans' },
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  return items;
}

function availableUnits(sku: SubscribedSku): number {
  return (sku.prepaidUnits?.enabled ?? 0) - sku.consumedUnits;
}

function skuName(skus: SubscribedSku[] | Failed, skuId: string): string {
  if (failed(skus)) return skuId;
  return skus.find((s) => s.skuId === skuId)?.skuPartNumber ?? skuId;
}

/** Find a tenant SKU by part number or id; partial part-number matches ("E3") are allowed when unique. */
function resolveSku(skus: SubscribedSku[], ref: string): SubscribedSku {
  const wanted = ref.trim().toLowerCase();
  const exact = skus.filter((s) => s.skuId.toLowerCase() === wanted || s.skuPartNumber.toLowerCase() === wanted);
  const candidates = exact.length ? exact : skus.filter((s) => s.skuPartNumber.toLowerCase().includes(wanted));
  if (!candidates.length) {
    throw new CliError('NOT_FOUND', `No license SKU in this tenant matches "${ref}"`, {
      hint: `Available SKUs: ${skus.map((s) => s.skuPartNumber).join(', ') || '(none)'}`,
    });
  }
  return pickOne(
    candidates,
    ref,
    'license SKU',
    (s) => [s.skuPartNumber, s.skuId],
    (s) => ({ skuPartNumber: s.skuPartNumber, skuId: s.skuId, availableUnits: availableUnits(s) }),
  );
}

function resolvePlans(sku: SubscribedSku, refs: string[]): ServicePlan[] {
  return refs.map((ref) => {
    const wanted = ref.toLowerCase();
    const plan = sku.servicePlans.find(
      (p) => p.servicePlanName.toLowerCase() === wanted || p.servicePlanId.toLowerCase() === wanted,
    );
    if (!plan) {
      throw new CliError('USAGE', `${sku.skuPartNumber} has no service plan "${ref}"`, {
        hint: `Plans in ${sku.skuPartNumber}: ${sku.servicePlans.map((p) => p.servicePlanName).join(', ')}`,
      });
    }
    return plan;
  });
}

async function setManager(ctx: Context, userId: string, managerId: string): Promise<void> {
  await ctx.graph.put(`/users/${seg(userId)}/manager/$ref`, {
    '@odata.id': ctx.graph.url(`/directoryObjects/${seg(managerId)}`),
  });
}

/** The `users get` view: profile, sync state, manager, licenses and last sign-in, degrading per section. */
export async function getUserOverview(ctx: Context, ref: string) {
  const user = (await resolveUser(ctx, ref, DETAIL_SELECT)) as GraphUser;
  const path = `/users/${seg(user.id)}`;
  const extra = await settle({
    manager: ctx.graph.tryGet<GraphUser>(`${path}/manager`, { query: { $select: 'id,displayName,userPrincipalName' } }),
    // Fetched separately: selecting signInActivity without AuditLog.Read.All fails the whole request.
    signInActivity: ctx.graph
      .get<{ signInActivity?: SignInActivity }>(path, { query: { $select: 'signInActivity' } })
      .then<SignInActivity | null>((r) => r.signInActivity ?? null),
    skus: loadSkus(ctx),
  });
  if (ctx.full) return { ...user, manager: extra.manager, signInActivity: extra.signInActivity };

  const activity = extra.signInActivity;
  let signIns: unknown;
  let lastSignIn: string | null | undefined;
  if (failed(activity)) {
    signIns = { ...activity, hint: 'Sign-in activity needs AuditLog.Read.All and Entra ID P1/P2.' };
  } else if (activity) {
    lastSignIn =
      activity.lastSuccessfulSignInDateTime ??
      [activity.lastSignInDateTime, activity.lastNonInteractiveSignInDateTime].filter(Boolean).sort().at(-1) ??
      null;
    signIns = {
      lastInteractive: activity.lastSignInDateTime ?? null,
      lastNonInteractive: activity.lastNonInteractiveSignInDateTime ?? null,
      lastSuccessful: activity.lastSuccessfulSignInDateTime ?? null,
    };
  } else {
    lastSignIn = null;
    signIns = { lastInteractive: null, lastNonInteractive: null, lastSuccessful: null, note: 'No sign-in recorded.' };
  }

  const manager = extra.manager;
  return compact({
    id: user.id,
    displayName: user.displayName,
    userPrincipalName: user.userPrincipalName,
    mail: user.mail,
    accountEnabled: user.accountEnabled,
    userType: user.userType,
    department: user.department,
    jobTitle: user.jobTitle,
    officeLocation: user.officeLocation,
    mobilePhone: user.mobilePhone,
    employeeId: user.employeeId,
    usageLocation: user.usageLocation,
    createdDateTime: user.createdDateTime,
    manager: failed(manager)
      ? manager
      : manager
        ? { displayName: manager.displayName, userPrincipalName: manager.userPrincipalName }
        : null,
    licenses: (user.assignedLicenses ?? []).map((l) => skuName(extra.skus, l.skuId)),
    isSynced: user.onPremisesSyncEnabled === true,
    onPremisesLastSyncDateTime: user.onPremisesSyncEnabled ? user.onPremisesLastSyncDateTime : undefined,
    lastPasswordChangeDateTime: user.lastPasswordChangeDateTime,
    daysSincePasswordChange: daysSince(user.lastPasswordChangeDateTime),
    signIns,
    daysSinceLastSignIn: lastSignIn === undefined ? undefined : daysSince(lastSignIn),
  });
}

// Keyed by @odata.type. `segment` is the type-specific collection used to delete a method; password has none.
const AUTH_METHOD_TYPES: Record<string, { label: string; segment?: string; mfa: boolean }> = {
  '#microsoft.graph.passwordAuthenticationMethod': { label: 'Password', mfa: false },
  '#microsoft.graph.phoneAuthenticationMethod': { label: 'Phone', segment: 'phoneMethods', mfa: true },
  '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod': {
    label: 'Microsoft Authenticator',
    segment: 'microsoftAuthenticatorMethods',
    mfa: true,
  },
  '#microsoft.graph.fido2AuthenticationMethod': { label: 'Passkey (FIDO2)', segment: 'fido2Methods', mfa: true },
  '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod': {
    label: 'Windows Hello for Business',
    segment: 'windowsHelloForBusinessMethods',
    mfa: true,
  },
  '#microsoft.graph.temporaryAccessPassAuthenticationMethod': {
    label: 'Temporary Access Pass',
    segment: 'temporaryAccessPassMethods',
    mfa: false,
  },
  '#microsoft.graph.emailAuthenticationMethod': { label: 'Email (SSPR only)', segment: 'emailMethods', mfa: false },
  '#microsoft.graph.softwareOathAuthenticationMethod': {
    label: 'Software OATH token',
    segment: 'softwareOathMethods',
    mfa: true,
  },
  '#microsoft.graph.platformCredentialAuthenticationMethod': {
    label: 'Platform Credential (macOS)',
    segment: 'platformCredentialMethods',
    mfa: true,
  },
  '#microsoft.graph.externalAuthenticationMethod': {
    label: 'External MFA',
    segment: 'externalAuthenticationMethods',
    mfa: true,
  },
};

function authMethodDetail(method: AuthMethod): string | undefined {
  const m = method as Record<string, string | number | boolean | null | undefined>;
  switch (method['@odata.type']) {
    case '#microsoft.graph.phoneAuthenticationMethod':
      return `${m.phoneType}: ${m.phoneNumber}${m.smsSignInState === 'ready' ? ' (SMS sign-in ready)' : ''}`;
    case '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod':
      return [m.displayName, m.phoneAppVersion ? `app ${m.phoneAppVersion}` : null].filter(Boolean).join(', ');
    case '#microsoft.graph.fido2AuthenticationMethod':
      return [m.displayName, m.model, m.passkeyType].filter(Boolean).join(', ');
    case '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod':
      return [m.displayName, m.keyStrength ? `key ${m.keyStrength}` : null].filter(Boolean).join(', ');
    case '#microsoft.graph.temporaryAccessPassAuthenticationMethod':
      return `${m.lifetimeInMinutes} min from ${m.startDateTime}, ${m.isUsable ? 'usable' : `not usable (${m.methodUsabilityReason})`}`;
    case '#microsoft.graph.emailAuthenticationMethod':
      return String(m.emailAddress);
    case '#microsoft.graph.platformCredentialAuthenticationMethod':
      return [m.displayName, m.platform].filter(Boolean).join(', ');
    case '#microsoft.graph.externalAuthenticationMethod':
      return m.displayName ? String(m.displayName) : undefined;
    default:
      return undefined;
  }
}

function shapeAuthMethod(method: AuthMethod) {
  const known = AUTH_METHOD_TYPES[method['@odata.type']];
  return compact({
    id: method.id,
    type: known?.label ?? method['@odata.type'].replace('#microsoft.graph.', ''),
    detail: authMethodDetail(method),
    createdDateTime: method.createdDateTime ?? undefined,
  });
}

async function listAuthMethods(ctx: Context, userId: string): Promise<AuthMethod[]> {
  const { items } = await ctx.graph.list<AuthMethod>(`/users/${seg(userId)}/authentication/methods`, {
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  return items;
}

/** Registered methods plus the MFA/SSPR registration summary; each half degrades to `{ error }`. */
export async function getAuthMethodsSummary(ctx: Context, userId: string) {
  const result = await settle({
    methods: listAuthMethods(ctx, userId),
    registration: ctx.graph.get<Record<string, unknown>>(
      `/reports/authenticationMethods/userRegistrationDetails/${seg(userId)}`,
    ),
  });
  if (ctx.full) return result;
  const { methods, registration } = result;
  return {
    methods: failed(methods) ? methods : methods.map(shapeAuthMethod),
    hasStrongMethod: failed(methods)
      ? undefined
      : methods.some((m) => AUTH_METHOD_TYPES[m['@odata.type']]?.mfa === true),
    registration: failed(registration)
      ? { ...registration, hint: 'The registration report needs AuditLog.Read.All and Entra ID P1/P2.' }
      : compact({
          isMfaRegistered: registration.isMfaRegistered,
          isMfaCapable: registration.isMfaCapable,
          isSsprRegistered: registration.isSsprRegistered,
          isSsprEnabled: registration.isSsprEnabled,
          isPasswordlessCapable: registration.isPasswordlessCapable,
          defaultMfaMethod: registration.userPreferredMethodForSecondaryAuthentication,
          systemPreferredMethods: registration.systemPreferredAuthenticationMethods,
          methodsRegistered: registration.methodsRegistered,
          isAdmin: registration.isAdmin,
          lastUpdatedDateTime: registration.lastUpdatedDateTime,
        }),
  };
}

async function findAuthMethod(ctx: Context, userRef: string, methodId: string) {
  const user = await resolveUser(ctx, userRef);
  const methods = await listAuthMethods(ctx, user.id);
  const method = methods.find((m) => m.id.toLowerCase() === methodId.trim().toLowerCase());
  if (!method) {
    throw new CliError('NOT_FOUND', `${user.userPrincipalName} has no authentication method ${methodId}`, {
      hint: `Run \`m365 entra users auth-methods ${user.userPrincipalName}\` to list method ids.`,
    });
  }
  const segment = AUTH_METHOD_TYPES[method['@odata.type']]?.segment;
  if (!segment) {
    throw new CliError('USAGE', `${shapeAuthMethod(method).type} methods cannot be removed`, {
      hint: method['@odata.type'].includes('password')
        ? 'Reset the password instead: `m365 entra users reset-password`.'
        : 'Remove it in the Entra admin center.',
    });
  }
  return { user, methods, method, segment };
}

const LICENSE_ERRORS: Record<string, string> = {
  CountViolation: 'No licenses left for this SKU: buy more or free one up.',
  MutuallyExclusiveViolation: 'A service plan conflicts with another license the user has; disable one of them.',
  DependencyViolation: 'A plan depends on another plan that is not assigned or is disabled.',
  ProhibitedInUsageLocationViolation: "A plan isn't available in the user's usage location.",
  UniquenessViolation: 'A unique attribute (e.g. proxy address) conflicts with another object.',
  Others: 'Assignment failed for another reason; check the group licensing pane in the Entra admin center.',
};

function shapeDevice(device: Record<string, unknown>, relationships: string[]) {
  const trust: Record<string, string> = {
    AzureAd: 'Entra joined',
    ServerAd: 'Hybrid joined',
    Workplace: 'Entra registered',
  };
  const os = [device.operatingSystem, device.operatingSystemVersion].filter(Boolean).join(' ');
  return compact({
    id: device.id as string,
    deviceId: device.deviceId as string | undefined,
    displayName: device.displayName as string,
    os: os || null,
    joinType: device.trustType ? (trust[device.trustType as string] ?? device.trustType) : null,
    isCompliant: (device.isCompliant as boolean | null | undefined) ?? null,
    isManaged: (device.isManaged as boolean | null | undefined) ?? null,
    accountEnabled: device.accountEnabled as boolean | undefined,
    lastSignIn: (device.approximateLastSignInDateTime as string | null | undefined) ?? null,
    daysSinceLastSignIn: daysSince(device.approximateLastSignInDateTime as string | null | undefined),
    relationship: relationships,
  });
}

/** One sign-in log entry, with the error code explained. Shared by directory sign-ins and troubleshoot views. */
export function shapeSignin(signin: SignIn) {
  const code = signin.status?.errorCode ?? 0;
  const info = code ? signinErrorInfo(code) : undefined;
  const location = signin.location
    ? [signin.location.city, signin.location.state, signin.location.countryOrRegion].filter(Boolean).join(', ')
    : '';
  const device = signin.deviceDetail;
  const mfaRequired = signin.authenticationRequirement === 'multiFactorAuthentication';
  return compact({
    time: signin.createdDateTime,
    user: signin.userPrincipalName,
    app: signin.appDisplayName ?? null,
    resource: signin.resourceDisplayName ?? undefined,
    clientApp: signin.clientAppUsed ?? null,
    interactive: signin.isInteractive,
    ipAddress: signin.ipAddress ?? null,
    location: location || null,
    device: device
      ? compact({
          displayName: device.displayName || undefined,
          os: device.operatingSystem || undefined,
          browser: device.browser || undefined,
          isCompliant: device.isCompliant ?? undefined,
          isManaged: device.isManaged ?? undefined,
          trustType: device.trustType || undefined,
        })
      : undefined,
    result: code === 0 ? 'success' : 'failure',
    errorCode: code || undefined,
    failureReason: code ? (signin.status?.failureReason ?? null) : undefined,
    explanation: info
      ? compact({ name: info.name, meaning: info.meaning, action: info.action, benign: info.benign })
      : code
        ? { lookup: signinErrorLookupUrl(code) }
        : undefined,
    conditionalAccess: signin.conditionalAccessStatus ?? undefined,
    mfaRequired,
    // v1.0 has no per-step MFA detail; a successful sign-in that required MFA satisfied it.
    mfaSatisfied: mfaRequired ? code === 0 : undefined,
    risk:
      signin.riskLevelDuringSignIn && signin.riskLevelDuringSignIn !== 'none'
        ? signin.riskLevelDuringSignIn
        : undefined,
    correlationId: signin.correlationId ?? undefined,
  });
}

async function userWithLicenses(ctx: Context, ref: string) {
  return (await resolveUser(
    ctx,
    ref,
    'id,displayName,userPrincipalName,mail,usageLocation,assignedLicenses,licenseAssignmentStates',
  )) as GraphUser;
}

async function planLicenseAdd(ctx: Context, args: { user: string; sku: string; disablePlans?: string[] }) {
  const [user, skus] = await Promise.all([userWithLicenses(ctx, args.user), loadSkus(ctx)]);
  const sku = resolveSku(skus, args.sku);
  const plans = resolvePlans(sku, args.disablePlans ?? []);
  const alreadyAssigned = (user.assignedLicenses ?? []).some((l) => l.skuId === sku.skuId);
  return { user, sku, plans, alreadyAssigned };
}

async function planLicenseRemove(ctx: Context, args: { user: string; sku: string }) {
  const [user, skus] = await Promise.all([userWithLicenses(ctx, args.user), loadSkus(ctx)]);
  const sku = resolveSku(skus, args.sku);
  const states = (user.licenseAssignmentStates ?? []).filter((s) => s.skuId === sku.skuId);
  if (!states.length) {
    throw new CliError('NOT_FOUND', `${user.userPrincipalName} does not have ${sku.skuPartNumber}`, {
      hint: `Run \`m365 entra users licenses ${user.userPrincipalName}\` to see their licenses.`,
    });
  }
  const groups = states.flatMap((s) => (s.assignedByGroup ? [s.assignedByGroup] : []));
  if (groups.length === states.length) {
    throw new CliError('USAGE', `${sku.skuPartNumber} is inherited from a group, not assigned directly`, {
      hint: `Remove the user from group ${groups.join(', ')} (or remove the license from that group) instead.`,
      details: { groupIds: groups },
    });
  }
  return { user, sku, alsoInheritedFrom: groups };
}

async function resolveDeletedUser(ctx: Context, ref: string): Promise<GraphUser & { deletedDateTime?: string }> {
  const value = ref.trim();
  if (isGuid(value)) {
    const found = await ctx.graph.tryGet<GraphUser & { deletedDateTime?: string }>(
      `/directory/deletedItems/${seg(value)}`,
    );
    if (!found) {
      throw new CliError('NOT_FOUND', `No deleted user with id ${value}`, {
        hint: 'Deleted users are purged after 30 days. Run `m365 entra users deleted` to list restorable users.',
      });
    }
    return found;
  }
  // Deleted users' UPNs are rewritten to `<id without dashes><original UPN>`, so match on the suffix.
  const { items } = await ctx.graph.list<GraphUser & { deletedDateTime?: string }>(
    '/directory/deletedItems/microsoft.graph.user',
    { query: { $select: 'id,displayName,userPrincipalName,mail,deletedDateTime' }, limit: 999 },
  );
  const wanted = value.toLowerCase();
  const candidates = items.filter(
    (u) =>
      u.displayName?.toLowerCase().includes(wanted) ||
      u.mail?.toLowerCase() === wanted ||
      u.userPrincipalName?.toLowerCase().endsWith(wanted),
  );
  return pickOne(
    candidates,
    value,
    'deleted user',
    (u) => [u.displayName, u.mail],
    (u) => ({
      id: u.id,
      displayName: u.displayName,
      userPrincipalName: u.userPrincipalName,
      deletedDateTime: u.deletedDateTime,
    }),
  );
}

const UPDATE_FIELDS = [
  'displayName',
  'jobTitle',
  'department',
  'officeLocation',
  'mobilePhone',
  'usageLocation',
  'employeeId',
] as const;

function updateBody(args: Partial<Record<(typeof UPDATE_FIELDS)[number], string>>): Record<string, string> {
  const body: Record<string, string> = {};
  for (const field of UPDATE_FIELDS) {
    const value = args[field];
    if (value !== undefined) body[field] = field === 'usageLocation' ? checkUsageLocation(value) : value;
  }
  return body;
}

function checkUsageLocation(value: string): string {
  if (!/^[A-Za-z]{2}$/.test(value)) {
    throw new CliError('USAGE', `--usage-location must be a two-letter country code, got "${value}"`, {
      hint: 'Use ISO 3166 codes such as US, GB, DE.',
    });
  }
  return value.toUpperCase();
}

const accountEnabledPermissions = {
  delegated: ['User.EnableDisableAccount.All', 'User.Read.All'],
  application: ['User.EnableDisableAccount.All', 'User.Read.All'],
  roles: ['User Administrator', 'Authentication Administrator', 'Privileged Authentication Administrator'],
};

async function userState(ctx: Context, ref: string) {
  return (await resolveUser(ctx, ref, STATE_SELECT)) as GraphUser;
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'users list',
    summary: 'List or search users, with filters for disabled, guest, unlicensed and inactive accounts',
    description:
      'Use to find users or answer "who is disabled / unlicensed / a guest / inactive". Filters combine with AND. ' +
      '--search matches display name, UPN and mail. --inactive-days N returns users whose last interactive ' +
      'sign-in is older than N days and adds lastSignIn; it needs AuditLog.Read.All and Entra ID P1/P2, cannot be ' +
      'combined with other filters (Graph limitation), and omits users who have never signed in. Follow up with ' +
      '`m365 entra users get <upn>` for one user.',
    examples: [
      'm365 entra users list --search jane',
      'm365 entra users list --department Sales --disabled',
      'm365 entra users list --unlicensed --all',
      'm365 entra users list --inactive-days 90',
    ],
    args: [
      { name: 'search', type: 'string', description: 'Match display name, UPN or mail (word prefix).' },
      { name: 'department', type: 'string', description: 'Exact department name.' },
      { name: 'disabled', type: 'boolean', description: 'Only accounts with sign-in blocked.' },
      { name: 'guests', type: 'boolean', description: 'Only guest (B2B) accounts.' },
      { name: 'unlicensed', type: 'boolean', description: 'Only users with no license assigned.' },
      { name: 'inactiveDays', type: 'number', description: 'Only users with no interactive sign-in for N days.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['User.Read.All', 'AuditLog.Read.All'],
      application: ['User.Read.All', 'AuditLog.Read.All'],
    },
    async run(ctx, args): Promise<ListResult> {
      const filters: string[] = [];
      if (args.department) filters.push(`department eq ${odata(args.department)}`);
      if (args.disabled) filters.push('accountEnabled eq false');
      if (args.guests) filters.push("userType eq 'Guest'");
      if (args.unlicensed) filters.push('assignedLicenses/$count eq 0');
      let select = LIST_SELECT;
      if (args.inactiveDays !== undefined) {
        if (args.inactiveDays <= 0) throw new CliError('USAGE', '--inactive-days must be a positive number');
        if (filters.length || args.search) {
          throw new CliError('USAGE', '--inactive-days cannot be combined with other filters or --search', {
            hint: 'Graph does not allow sign-in activity filters with other filters. Run it alone and filter the output.',
          });
        }
        filters.push(`signInActivity/lastSignInDateTime le ${daysAgo(args.inactiveDays)}`);
        select += ',signInActivity';
      }
      const term = args.search?.replace(/"/g, '');
      const advanced = Boolean(term) || args.unlicensed;
      const result = await ctx.graph.list<GraphUser>('/users', {
        query: {
          $select: select,
          $filter: filters.length ? filters.join(' and ') : undefined,
          $search: term ? `"displayName:${term}" OR "userPrincipalName:${term}" OR "mail:${term}"` : undefined,
          $count: advanced ? true : undefined,
        },
        advanced,
        limit: listLimit(args),
      });
      return ctx.full ? result : { items: result.items.map(shapeUser), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users get',
    summary: 'Show a user: account state, sync source, manager, licenses, password age and last sign-in',
    description:
      'Start here for any question about one person. accountEnabled=false means sign-in is blocked. isSynced=true ' +
      'means the account is mastered in on-premises AD (change it there). daysSinceLastSignIn uses the last ' +
      'successful sign-in and needs AuditLog.Read.All plus Entra ID P1/P2; without them signIns shows an error and ' +
      'the rest still works. Next: `users signins` for failures, `users auth-methods` for MFA, `users licenses` ' +
      'for missing apps.',
    examples: ['m365 entra users get jane@contoso.com', 'm365 entra users get "Jane Doe"'],
    args: [USER_ARG],
    risk: 'read',
    permissions: {
      delegated: ['User.Read.All', 'AuditLog.Read.All', 'LicenseAssignment.Read.All'],
      application: ['User.Read.All', 'AuditLog.Read.All', 'LicenseAssignment.Read.All'],
    },
    run: (ctx, args) => getUserOverview(ctx, args.user),
  }),

  defineCommand({
    domain: 'entra',
    name: 'users groups',
    summary: "List a user's groups (including nested) and directory roles",
    description:
      'Shows every group the user is in directly or through nesting, with type (Security, Microsoft 365, ' +
      'Mail-enabled security, Distribution) and whether membership is dynamic, plus active directory (admin) roles. ' +
      'Use to explain why someone has access to a site/app or receives a policy. Eligible (PIM) roles are not shown.',
    examples: ['m365 entra users groups jane@contoso.com --all'],
    args: [USER_ARG, ...LIST_ARGS],
    risk: 'read',
    permissions: {
      delegated: ['User.Read.All', 'GroupMember.Read.All'],
      application: ['User.Read.All', 'GroupMember.Read.All'],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const result = await ctx.graph.list<Record<string, unknown>>(`/users/${seg(user.id)}/transitiveMemberOf`, {
        query: {
          $select:
            'id,displayName,mail,groupTypes,securityEnabled,mailEnabled,membershipRule,description,roleTemplateId',
        },
        limit: listLimit(args),
      });
      if (ctx.full) return result;
      const ofType = (type: string) => result.items.filter((o) => o['@odata.type'] === `#microsoft.graph.${type}`);
      return {
        user: identity(user),
        groups: ofType('group').map((g) => {
          const types = (g.groupTypes as string[] | undefined) ?? [];
          const kind = types.includes('Unified')
            ? 'Microsoft 365'
            : g.securityEnabled
              ? g.mailEnabled
                ? 'Mail-enabled security'
                : 'Security'
              : 'Distribution';
          return compact({
            id: g.id as string,
            displayName: g.displayName as string,
            type: kind,
            dynamic: types.includes('DynamicMembership'),
            mail: (g.mail as string | null | undefined) ?? null,
          });
        }),
        directoryRoles: ofType('directoryRole').map((r) => ({
          id: r.id as string,
          displayName: r.displayName as string,
          roleTemplateId: r.roleTemplateId as string,
        })),
        administrativeUnits: ofType('administrativeUnit').map((a) => ({
          id: a.id as string,
          displayName: a.displayName as string,
        })),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users licenses',
    summary: "Explain a user's licenses: direct vs group-assigned, errors, and disabled service plans",
    description:
      'Use for "why doesn\'t X have Teams/Exchange/etc.". Each assignment shows assignedVia (direct or group), ' +
      'state and error (CountViolation = no seats, MutuallyExclusiveViolation = conflicting plans, ...) with an ' +
      'explanation, disabledPlans (service plans switched off in that assignment) and plansNotReady (plans whose ' +
      'provisioning is pending or failed; delegated sign-in only). A missing usageLocation blocks assignment. ' +
      'Fix with `users license add/remove` or by changing group membership.',
    examples: ['m365 entra users licenses jane@contoso.com'],
    args: [USER_ARG],
    risk: 'read',
    permissions: {
      delegated: ['User.Read.All', 'LicenseAssignment.Read.All', 'GroupMember.Read.All'],
      application: ['User.Read.All', 'LicenseAssignment.Read.All', 'GroupMember.Read.All'],
    },
    async run(ctx, args) {
      const user = await userWithLicenses(ctx, args.user);
      const states = user.licenseAssignmentStates ?? [];
      const groupIds = [...new Set(states.flatMap((s) => (s.assignedByGroup ? [s.assignedByGroup] : [])))];
      const { skus, details, groups } = await settle({
        skus: loadSkus(ctx),
        details: ctx.graph
          .list<{ skuId: string; servicePlans: ServicePlan[] }>(`/users/${seg(user.id)}/licenseDetails`, {
            limit: Number.POSITIVE_INFINITY,
            pageSize: false,
          })
          .then<{ skuId: string; servicePlans: ServicePlan[] }[]>((r) => r.items),
        // Group names are a convenience; an unreadable group still shows its id.
        groups: Promise.all(
          groupIds.map((id) =>
            ctx.graph
              .tryGet<{ id: string; displayName: string }>(`/groups/${seg(id)}`, {
                query: { $select: 'id,displayName' },
              })
              .catch(() => null),
          ),
        ),
      });
      if (ctx.full) return { user, skus, licenseDetails: details };
      const planName = (skuId: string, planId: string) => {
        if (failed(skus)) return planId;
        const sku = skus.find((s) => s.skuId === skuId);
        return sku?.servicePlans.find((p) => p.servicePlanId === planId)?.servicePlanName ?? planId;
      };
      const assignments = states.map((s) => {
        const detail = failed(details) ? undefined : details.find((d) => d.skuId === s.skuId);
        const error = s.error && s.error !== 'None' ? s.error : null;
        return compact({
          sku: skuName(skus, s.skuId),
          skuId: s.skuId,
          assignedVia: s.assignedByGroup ? 'group' : 'direct',
          group: s.assignedByGroup
            ? compact({
                id: s.assignedByGroup,
                displayName: failed(groups) ? undefined : groups.find((g) => g?.id === s.assignedByGroup)?.displayName,
              })
            : undefined,
          state: s.state,
          error,
          explanation: error ? (LICENSE_ERRORS[error] ?? undefined) : undefined,
          disabledPlans: s.disabledPlans.map((id) => planName(s.skuId, id)),
          plansNotReady: detail?.servicePlans
            .filter((p) => p.provisioningStatus && !['Success', 'Disabled'].includes(p.provisioningStatus))
            .map((p) => ({ plan: p.servicePlanName, status: p.provisioningStatus })),
          lastUpdatedDateTime: s.lastUpdatedDateTime ?? undefined,
        });
      });
      const warnings: string[] = [];
      if (!user.usageLocation) {
        warnings.push(
          'usageLocation is not set: licenses cannot be assigned until it is (`users update --usage-location`).',
        );
      }
      if (assignments.some((a) => a.error)) warnings.push('Some assignments have errors; see explanation on each.');
      return compact({
        user: identity(user),
        usageLocation: user.usageLocation ?? null,
        assignments,
        warnings: warnings.length ? warnings : undefined,
        licenseDetails: failed(details) ? details : undefined,
      });
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users auth-methods',
    summary: "Show a user's MFA/sign-in methods and registration status",
    description:
      'Use for MFA problems ("not getting prompts", "new phone"). methods lists each registered method with its id ' +
      '(needed for `users auth-method remove`). registration (needs AuditLog.Read.All, Entra ID P1/P2) shows ' +
      'isMfaRegistered/isMfaCapable, SSPR state, passwordless capability and the default MFA method. ' +
      'hasStrongMethod=false means only password/email: issue a TAP (`users tap`) so they can register. ' +
      'Authentication Administrators see phone numbers masked.',
    examples: ['m365 entra users auth-methods jane@contoso.com'],
    args: [USER_ARG],
    risk: 'read',
    permissions: {
      delegated: ['User.Read.All', 'UserAuthenticationMethod.Read.All', 'AuditLog.Read.All'],
      application: ['User.Read.All', 'UserAuthenticationMethod.Read.All', 'AuditLog.Read.All'],
      roles: ['Authentication Administrator', 'Privileged Authentication Administrator', 'Global Reader'],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      return { user: identity(user), ...(await getAuthMethodsSummary(ctx, user.id)) };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users devices',
    summary: 'List the Entra device objects a user owns or registered',
    description:
      "Entra ID view of the user's devices: join type (Entra joined / Hybrid joined / Entra registered), " +
      'isCompliant and isManaged (as Conditional Access sees them), last sign-in, and whether the device object is ' +
      'enabled. Use when CA blocks a device. For Intune details (compliance reasons, sync) use the intune commands. ' +
      'Delegated sign-in only (Graph has no application permission for this).',
    examples: ['m365 entra users devices jane@contoso.com'],
    args: [USER_ARG],
    risk: 'read',
    permissions: { delegated: ['User.Read.All'] },
    async run(ctx, args): Promise<ListResult> {
      const user = await resolveUser(ctx, args.user);
      const options = { query: { $select: DEVICE_SELECT }, limit: Number.POSITIVE_INFINITY };
      const [owned, registered] = await Promise.all([
        ctx.graph.list<Record<string, unknown>>(`/users/${seg(user.id)}/ownedDevices`, options),
        ctx.graph.list<Record<string, unknown>>(`/users/${seg(user.id)}/registeredDevices`, options),
      ]);
      const byId = new Map<string, { device: Record<string, unknown>; relationships: string[] }>();
      for (const [relationship, list] of [
        ['owner', owned.items],
        ['registered', registered.items],
      ] as const) {
        for (const device of list) {
          const type = device['@odata.type'];
          if (type !== undefined && type !== '#microsoft.graph.device') continue;
          const id = device.id as string;
          const entry = byId.get(id) ?? { device, relationships: [] };
          entry.relationships.push(relationship);
          byId.set(id, entry);
        }
      }
      const entries = [...byId.values()];
      return {
        items: entries.map((e) => (ctx.full ? e.device : shapeDevice(e.device, e.relationships))),
        truncated: false,
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users signins',
    summary: "Show a user's recent sign-ins with failures explained",
    description:
      'Use for "X can\'t sign in". Newest first. Each entry shows app, client, IP, location, device (compliant? ' +
      'joined?), result, errorCode with a plain-language explanation and suggested action, Conditional Access ' +
      'result, whether MFA was required/satisfied and sign-in risk. Interactive and non-interactive (token refresh) ' +
      'sign-ins are both included. Needs AuditLog.Read.All and Entra ID P1/P2; logs are kept 30 days (7 on free).',
    examples: [
      'm365 entra users signins jane@contoso.com',
      'm365 entra users signins jane@contoso.com --failures-only --days 3',
    ],
    args: [
      USER_ARG,
      { name: 'days', type: 'number', default: 7, description: 'How many days back to look (max 30).' },
      { name: 'failuresOnly', type: 'boolean', description: 'Only failed sign-ins.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['User.Read.All', 'AuditLog.Read.All'],
      application: ['User.Read.All', 'AuditLog.Read.All'],
      roles: ['Reports Reader', 'Security Reader', 'Global Reader'],
    },
    async run(ctx, args): Promise<ListResult> {
      if (args.days <= 0 || args.days > 30) throw new CliError('USAGE', '--days must be between 1 and 30');
      const user = await resolveUser(ctx, args.user);
      const filters = [`userId eq ${odata(user.id)}`, `createdDateTime ge ${daysAgo(args.days)}`];
      if (args.failuresOnly) filters.push('status/errorCode ne 0');
      const result = await ctx.graph.list<SignIn>('/auditLogs/signIns', {
        query: { $filter: filters.join(' and ') },
        limit: listLimit(args),
      });
      return ctx.full ? result : { items: result.items.map(shapeSignin), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users create',
    summary: 'Create a cloud user account (returns the temporary password)',
    description:
      'Creates an enabled member account. Without --password a strong random password is generated and returned ' +
      'once as temporaryPassword: hand it to the user securely. The user must change it at first sign-in unless ' +
      '--no-force-change. Set --usage-location now if the user will be licensed, then run `users license add`. ' +
      'The UPN domain must be a verified tenant domain. --manager needs User.ReadWrite.All; if it fails the user is ' +
      'still created and the result says so.',
    examples: [
      'm365 entra users create --display-name "Jane Doe" --upn jane@contoso.com --usage-location US',
      'm365 entra users create --display-name "Bob Ray" --upn bob@contoso.com --department Sales --manager jane@contoso.com',
    ],
    args: [
      { name: 'displayName', type: 'string', required: true, description: 'Full name shown in the address book.' },
      { name: 'upn', type: 'string', required: true, description: 'Sign-in name, e.g. jane@contoso.com.' },
      { name: 'mailNickname', type: 'string', description: 'Mail alias; defaults to the UPN prefix.' },
      { name: 'password', type: 'string', secret: true, description: 'Initial password; generated when omitted.' },
      {
        name: 'forceChange',
        type: 'boolean',
        default: true,
        description: 'Require a password change at first sign-in.',
      },
      { name: 'usageLocation', type: 'string', description: 'Two-letter country code (needed for licenses).' },
      { name: 'department', type: 'string', description: 'Department.' },
      { name: 'jobTitle', type: 'string', description: 'Job title.' },
      { name: 'manager', type: 'string', description: 'Manager (UPN, email, name or id).' },
    ],
    risk: 'write',
    permissions: { delegated: ['User.Create'], application: ['User.Create'], roles: ['User Administrator'] },
    async preview(ctx, args) {
      const manager = args.manager ? await resolveUser(ctx, args.manager) : undefined;
      return compact({
        action: `Create user ${args.upn}`,
        displayName: args.displayName,
        userPrincipalName: args.upn,
        mailNickname: args.mailNickname ?? args.upn.split('@')[0],
        usageLocation: args.usageLocation ? checkUsageLocation(args.usageLocation) : undefined,
        department: args.department,
        jobTitle: args.jobTitle,
        manager: manager ? identity(manager) : undefined,
        password: args.password ? '(provided)' : '(generated when run)',
        forceChangePasswordNextSignIn: args.forceChange,
      });
    },
    async run(ctx, args) {
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(args.upn)) {
        throw new CliError('USAGE', `--upn must look like name@domain, got "${args.upn}"`);
      }
      const manager = args.manager ? await resolveUser(ctx, args.manager) : undefined;
      const password = args.password ?? generatePassword();
      const body = compact({
        accountEnabled: true,
        displayName: args.displayName,
        userPrincipalName: args.upn,
        // mailNickname rejects characters such as @ ( ) [ ] ; : < > , and spaces.
        mailNickname: args.mailNickname ?? (args.upn.split('@')[0] as string).replace(/[^A-Za-z0-9._\-'!#^~]/g, ''),
        passwordProfile: { password, forceChangePasswordNextSignIn: args.forceChange },
        usageLocation: args.usageLocation ? checkUsageLocation(args.usageLocation) : undefined,
        department: args.department,
        jobTitle: args.jobTitle,
      });
      const created = await ctx.graph.post<GraphUser>('/users', body);
      let managerResult: unknown;
      if (manager) {
        managerResult = await setManager(ctx, created.id, manager.id).then(
          () => identity(manager),
          (error: unknown) => ({
            error: error instanceof Error ? error.message : String(error),
            hint: `User created; set the manager later with \`m365 entra users update ${created.userPrincipalName} --manager ...\`.`,
          }),
        );
      }
      return compact({
        user: shapeUser({ ...created, usageLocation: body.usageLocation } as GraphUser),
        temporaryPassword: args.password ? undefined : password,
        forceChangePasswordNextSignIn: args.forceChange,
        manager: managerResult,
        next: `Assign a license: m365 entra users license add ${created.userPrincipalName} --sku <SKU>`,
      });
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users update',
    summary: "Change a user's profile fields or manager",
    description:
      'Only the options you pass are changed. Synced (on-premises) users must be edited in AD for most fields. ' +
      '--usage-location is required before licenses can be assigned. Dry run (--dry-run) shows old -> new values.',
    examples: [
      'm365 entra users update jane@contoso.com --job-title "Sales Manager" --department Sales',
      'm365 entra users update jane@contoso.com --usage-location GB',
      'm365 entra users update jane@contoso.com --manager bob@contoso.com',
    ],
    args: [
      USER_ARG,
      { name: 'displayName', type: 'string', description: 'New display name.' },
      { name: 'jobTitle', type: 'string', description: 'New job title.' },
      { name: 'department', type: 'string', description: 'New department.' },
      { name: 'officeLocation', type: 'string', description: 'New office location.' },
      { name: 'mobilePhone', type: 'string', description: 'New mobile phone number (profile, not MFA).' },
      { name: 'usageLocation', type: 'string', description: 'Two-letter country code.' },
      { name: 'employeeId', type: 'string', description: 'Employee id (max 16 chars).' },
      { name: 'manager', type: 'string', description: 'New manager (UPN, email, name or id).' },
    ],
    risk: 'write',
    permissions: {
      delegated: ['User.ReadWrite.All'],
      application: ['User.ReadWrite.All'],
      roles: ['User Administrator'],
    },
    async preview(ctx, args) {
      const body = updateBody(args);
      if (!Object.keys(body).length && !args.manager) throw new CliError('USAGE', 'Nothing to update');
      const user = (await resolveUser(
        ctx,
        args.user,
        ['id', 'displayName', 'userPrincipalName', 'mail', 'onPremisesSyncEnabled', ...Object.keys(body)].join(','),
      )) as GraphUser & Record<string, unknown>;
      const changes: Record<string, { from: unknown; to: unknown }> = {};
      for (const [field, value] of Object.entries(body)) changes[field] = { from: user[field] ?? null, to: value };
      if (args.manager) {
        const [current, next] = await Promise.all([
          ctx.graph.tryGet<GraphUser>(`/users/${seg(user.id)}/manager`, { query: { $select: 'userPrincipalName' } }),
          resolveUser(ctx, args.manager),
        ]);
        changes.manager = { from: current?.userPrincipalName ?? null, to: next.userPrincipalName };
      }
      return compact({
        user: identity(user),
        changes,
        warning: user.onPremisesSyncEnabled ? SYNCED_WARNING : undefined,
      });
    },
    async run(ctx, args) {
      const body = updateBody(args);
      if (!Object.keys(body).length && !args.manager) {
        throw new CliError('USAGE', 'Nothing to update', {
          hint: 'Pass at least one of --display-name, --job-title, --department, --office-location, --mobile-phone, --usage-location, --employee-id, --manager.',
        });
      }
      const user = await resolveUser(ctx, args.user);
      const manager = args.manager ? await resolveUser(ctx, args.manager) : undefined;
      if (Object.keys(body).length) await ctx.graph.patch(`/users/${seg(user.id)}`, body);
      if (manager) await setManager(ctx, user.id, manager.id);
      return compact({
        user: identity(user),
        updated: { ...body, ...(manager ? { manager: manager.userPrincipalName } : {}) },
      });
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users disable',
    summary: 'Block a user from signing in (accountEnabled = false)',
    description:
      'Use for offboarding or a compromised account. Existing sessions keep working until their tokens expire: ' +
      'follow with `users revoke-sessions`. Reversible with `users enable`. Synced users must be disabled in AD.',
    examples: ['m365 entra users disable jane@contoso.com'],
    args: [USER_ARG],
    risk: 'destructive',
    permissions: accountEnabledPermissions,
    async preview(ctx, args) {
      const user = await userState(ctx, args.user);
      return compact({
        user: identity(user),
        accountEnabled: user.accountEnabled,
        isSynced: user.onPremisesSyncEnabled === true,
        action:
          user.accountEnabled === false
            ? 'Already disabled; no change.'
            : `Block sign-in for ${user.userPrincipalName}.`,
        warning: user.onPremisesSyncEnabled ? SYNCED_WARNING : undefined,
      });
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      await ctx.graph.patch(`/users/${seg(user.id)}`, { accountEnabled: false });
      return {
        user: identity(user),
        accountEnabled: false,
        next: `Sign them out everywhere now: m365 entra users revoke-sessions ${user.userPrincipalName}`,
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users enable',
    summary: 'Allow a disabled user to sign in again (accountEnabled = true)',
    description: 'Reverses `users disable`. Synced users must be enabled in on-premises AD.',
    examples: ['m365 entra users enable jane@contoso.com'],
    args: [USER_ARG],
    risk: 'write',
    permissions: accountEnabledPermissions,
    async preview(ctx, args) {
      const user = await userState(ctx, args.user);
      return compact({
        user: identity(user),
        accountEnabled: user.accountEnabled,
        action: user.accountEnabled ? 'Already enabled; no change.' : `Allow sign-in for ${user.userPrincipalName}.`,
        warning: user.onPremisesSyncEnabled ? SYNCED_WARNING : undefined,
      });
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      await ctx.graph.patch(`/users/${seg(user.id)}`, { accountEnabled: true });
      return { user: identity(user), accountEnabled: true };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users reset-password',
    summary: "Reset a user's password (returns the temporary password once)",
    description:
      'Sets a new password; without --password a strong one is generated and returned once as temporaryPassword. ' +
      'The user must change it at next sign-in unless --no-force-change. Which admin can reset whom: Helpdesk ' +
      'Administrator resets non-admins and some readers, User Administrator also other user/helpdesk admins, ' +
      'Privileged Authentication Administrator anyone. App-only use also needs an admin role assigned to the app. ' +
      'Synced users need password writeback; reset them in AD otherwise. Consider `users revoke-sessions` after.',
    examples: ['m365 entra users reset-password jane@contoso.com'],
    args: [
      USER_ARG,
      { name: 'password', type: 'string', secret: true, description: 'New password; generated when omitted.' },
      { name: 'forceChange', type: 'boolean', default: true, description: 'Require a change at next sign-in.' },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['User-PasswordProfile.ReadWrite.All'],
      application: ['User-PasswordProfile.ReadWrite.All'],
      roles: ['Helpdesk Administrator', 'User Administrator', 'Privileged Authentication Administrator'],
    },
    async preview(ctx, args) {
      const user = (await resolveUser(
        ctx,
        args.user,
        'id,displayName,userPrincipalName,mail,accountEnabled,onPremisesSyncEnabled,lastPasswordChangeDateTime',
      )) as GraphUser;
      return compact({
        user: identity(user),
        accountEnabled: user.accountEnabled,
        lastPasswordChangeDateTime: user.lastPasswordChangeDateTime,
        action:
          `Replace the password of ${user.userPrincipalName} with ` +
          `${args.password ? 'the provided password' : 'a generated one'}` +
          `${args.forceChange ? '; they must change it at next sign-in' : ''}.`,
        warning: user.onPremisesSyncEnabled ? SYNCED_WARNING : undefined,
      });
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const password = args.password ?? generatePassword();
      await ctx.graph.patch(`/users/${seg(user.id)}`, {
        passwordProfile: { password, forceChangePasswordNextSignIn: args.forceChange },
      });
      return compact({
        user: identity(user),
        temporaryPassword: args.password ? undefined : password,
        forceChangePasswordNextSignIn: args.forceChange,
      });
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users revoke-sessions',
    summary: 'Sign a user out of every app and device by revoking refresh tokens',
    description:
      'Use after disabling a user, a password reset for compromise, or a lost device. Apps re-prompt for sign-in; ' +
      'access tokens already issued may work for up to an hour. Does not block sign-in (use `users disable`).',
    examples: ['m365 entra users revoke-sessions jane@contoso.com'],
    args: [USER_ARG],
    risk: 'destructive',
    permissions: {
      delegated: ['User.RevokeSessions.All', 'User.Read.All'],
      application: ['User.RevokeSessions.All', 'User.Read.All'],
      roles: ['Helpdesk Administrator', 'User Administrator'],
    },
    async preview(ctx, args) {
      const user = await userState(ctx, args.user);
      return {
        user: identity(user),
        accountEnabled: user.accountEnabled,
        action: `Revoke all refresh tokens and browser sessions for ${user.userPrincipalName}; they must sign in again everywhere.`,
      };
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      await ctx.graph.post(`/users/${seg(user.id)}/revokeSignInSessions`);
      return { user: identity(user), sessionsRevoked: true };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users delete',
    summary: 'Delete a user (soft delete: restorable for 30 days)',
    description:
      'Moves the user to deleted items, frees their licenses and soft-deletes their mailbox/OneDrive with them. ' +
      'Restore within 30 days with `users restore <id>`; after that it is permanent. Synced users must be deleted ' +
      'in AD. For offboarding, disabling first is usually safer.',
    examples: ['m365 entra users delete jane@contoso.com'],
    args: [USER_ARG],
    risk: 'destructive',
    permissions: {
      delegated: ['User.ReadWrite.All'],
      application: ['User.ReadWrite.All'],
      roles: ['User Administrator', 'Privileged Authentication Administrator'],
    },
    async preview(ctx, args) {
      const user = await userState(ctx, args.user);
      return compact({
        user: identity(user),
        accountEnabled: user.accountEnabled,
        licenses: (user.assignedLicenses ?? []).length,
        action:
          `Delete ${user.userPrincipalName}. It can be restored for ${DELETED_RETENTION_DAYS} days with ` +
          `\`m365 entra users restore ${user.id}\`.`,
        warning: user.onPremisesSyncEnabled ? SYNCED_WARNING : undefined,
      });
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      await ctx.graph.delete(`/users/${seg(user.id)}`);
      return { user: identity(user), deleted: true, restoreWith: `m365 entra users restore ${user.id}` };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users deleted',
    summary: 'List deleted users that can still be restored',
    description:
      `Deleted users stay restorable for ${DELETED_RETENTION_DAYS} days; daysUntilPurge shows how long is left. ` +
      'Restore with `users restore <id>`.',
    examples: ['m365 entra users deleted'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: ['User.Read.All'], application: ['User.Read.All'] },
    async run(ctx, args): Promise<ListResult> {
      const result = await ctx.graph.list<GraphUser & { deletedDateTime?: string }>(
        '/directory/deletedItems/microsoft.graph.user',
        { query: { $select: 'id,displayName,userPrincipalName,mail,deletedDateTime' }, limit: listLimit(args) },
      );
      if (ctx.full) return result;
      return {
        items: result.items.map((u) => {
          const age = daysSince(u.deletedDateTime);
          return {
            id: u.id,
            displayName: u.displayName,
            userPrincipalName: u.userPrincipalName,
            mail: u.mail ?? null,
            deletedDateTime: u.deletedDateTime ?? null,
            daysUntilPurge: age === null ? null : Math.max(0, DELETED_RETENTION_DAYS - age),
          };
        }),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users restore',
    summary: 'Restore a deleted user (within 30 days of deletion)',
    description:
      'Accepts the id from `users deleted`, or the display name / original UPN. Licenses, group memberships and ' +
      'mailbox come back with the account. Fails if the UPN or mail has been reused meanwhile.',
    examples: ['m365 entra users restore 6e7b768e-07e2-4810-8459-485f84f8f204', 'm365 entra users restore "Jane Doe"'],
    args: [{ ...USER_ARG, description: 'Deleted user id, display name or original UPN.' }],
    risk: 'write',
    permissions: {
      delegated: ['User.Read.All', 'User.DeleteRestore.All'],
      application: ['User.Read.All', 'User.DeleteRestore.All'],
      roles: ['User Administrator'],
    },
    async preview(ctx, args) {
      const user = await resolveDeletedUser(ctx, args.user);
      return {
        user: identity(user),
        deletedDateTime: user.deletedDateTime ?? null,
        action: `Restore ${user.displayName}`,
      };
    },
    async run(ctx, args) {
      const user = await resolveDeletedUser(ctx, args.user);
      const restored = await ctx.graph.post<GraphUser>(`/directory/deletedItems/${seg(user.id)}/restore`);
      return { user: restored ? identity(restored) : identity(user), restored: true };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users license add',
    summary: 'Assign a license SKU to a user, optionally with some service plans disabled',
    description:
      'SKU by part number (ENTERPRISEPACK, SPE_E3), unique fragment (E3) or skuId; `users licenses` and Graph ' +
      'subscribedSkus show names. --disable-plans takes service plan names or ids (e.g. TEAMS1) to switch off; ' +
      'running it again for an assigned SKU replaces its disabled plans. The user needs a usageLocation first. ' +
      'Dry run shows available seats.',
    examples: [
      'm365 entra users license add jane@contoso.com --sku SPE_E3',
      'm365 entra users license add jane@contoso.com --sku ENTERPRISEPACK --disable-plans YAMMER_ENTERPRISE,SWAY',
    ],
    args: [
      USER_ARG,
      { name: 'sku', type: 'string', required: true, description: 'SKU part number or skuId.' },
      {
        name: 'disablePlans',
        type: 'string[]',
        description: 'Service plans to disable (names or ids, comma separated).',
      },
    ],
    risk: 'write',
    permissions: {
      delegated: ['LicenseAssignment.ReadWrite.All', 'User.Read.All'],
      application: ['LicenseAssignment.ReadWrite.All', 'User.Read.All'],
      roles: ['License Administrator', 'User Administrator'],
    },
    async preview(ctx, args) {
      const plan = await planLicenseAdd(ctx, args);
      const warnings: string[] = [];
      if (!plan.user.usageLocation) warnings.push('usageLocation is not set; the assignment will be refused.');
      if (!plan.alreadyAssigned && availableUnits(plan.sku) <= 0) warnings.push('No seats available for this SKU.');
      return compact({
        user: { ...identity(plan.user), usageLocation: plan.user.usageLocation ?? null },
        sku: { skuPartNumber: plan.sku.skuPartNumber, skuId: plan.sku.skuId, availableUnits: availableUnits(plan.sku) },
        disabledPlans: plan.plans.map((p) => p.servicePlanName),
        action: plan.alreadyAssigned
          ? `Update ${plan.sku.skuPartNumber} on ${plan.user.userPrincipalName} (replace disabled plans)`
          : `Assign ${plan.sku.skuPartNumber} to ${plan.user.userPrincipalName}`,
        warnings: warnings.length ? warnings : undefined,
      });
    },
    async run(ctx, args) {
      const { user, sku, plans } = await planLicenseAdd(ctx, args);
      if (!user.usageLocation) {
        throw new CliError('USAGE', `${user.userPrincipalName} has no usage location, which licensing requires`, {
          hint: `Set it first: m365 entra users update ${user.userPrincipalName} --usage-location US (their country).`,
        });
      }
      await ctx.graph.post(`/users/${seg(user.id)}/assignLicense`, {
        addLicenses: [{ skuId: sku.skuId, disabledPlans: plans.map((p) => p.servicePlanId) }],
        removeLicenses: [],
      });
      return {
        user: identity(user),
        assigned: sku.skuPartNumber,
        disabledPlans: plans.map((p) => p.servicePlanName),
        next: `Check provisioning: m365 entra users licenses ${user.userPrincipalName}`,
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users license remove',
    summary: 'Remove a directly assigned license SKU from a user',
    description:
      'The user loses the apps and services in that SKU (data such as the mailbox is kept for a grace period). ' +
      'Group-inherited licenses cannot be removed per user: the error names the group to change instead.',
    examples: ['m365 entra users license remove jane@contoso.com --sku SPE_E3'],
    args: [USER_ARG, { name: 'sku', type: 'string', required: true, description: 'SKU part number or skuId.' }],
    risk: 'destructive',
    permissions: {
      delegated: ['LicenseAssignment.ReadWrite.All', 'User.Read.All'],
      application: ['LicenseAssignment.ReadWrite.All', 'User.Read.All'],
      roles: ['License Administrator', 'User Administrator'],
    },
    async preview(ctx, args) {
      const plan = await planLicenseRemove(ctx, args);
      return compact({
        user: identity(plan.user),
        sku: { skuPartNumber: plan.sku.skuPartNumber, skuId: plan.sku.skuId },
        action: `Remove the direct ${plan.sku.skuPartNumber} assignment from ${plan.user.userPrincipalName}`,
        note: plan.alsoInheritedFrom.length
          ? `Also inherited from group(s) ${plan.alsoInheritedFrom.join(', ')}; the user keeps the license through them.`
          : undefined,
      });
    },
    async run(ctx, args) {
      const { user, sku } = await planLicenseRemove(ctx, args);
      await ctx.graph.post(`/users/${seg(user.id)}/assignLicense`, { addLicenses: [], removeLicenses: [sku.skuId] });
      return { user: identity(user), removed: sku.skuPartNumber };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users tap',
    summary: 'Issue a Temporary Access Pass so a user can sign in and register MFA',
    description:
      'Use when a user lost their phone/MFA method or is new and passwordless. Returns the pass once; it replaces ' +
      'any existing pass. Lifetime 10-43200 minutes (policy limits apply). The TAP authentication method policy ' +
      'must be enabled for the user; isUsable=false with DisabledByPolicy means it is not. Multi-use passes need ' +
      'the policy to allow them.',
    examples: [
      'm365 entra users tap jane@contoso.com',
      'm365 entra users tap jane@contoso.com --lifetime-minutes 480 --one-time',
    ],
    args: [
      USER_ARG,
      { name: 'lifetimeMinutes', type: 'number', default: 60, description: 'Validity in minutes (10-43200).' },
      { name: 'oneTime', type: 'boolean', description: 'Pass can be used for one sign-in only.' },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['UserAuthMethod-TAP.ReadWrite.All', 'User.Read.All'],
      application: ['UserAuthMethod-TAP.ReadWrite.All', 'User.Read.All'],
      roles: ['Authentication Administrator', 'Privileged Authentication Administrator'],
    },
    async preview(ctx, args) {
      if (args.lifetimeMinutes < 10 || args.lifetimeMinutes > 43200) {
        throw new CliError('USAGE', '--lifetime-minutes must be between 10 and 43200');
      }
      const user = await userState(ctx, args.user);
      return {
        user: identity(user),
        accountEnabled: user.accountEnabled,
        action:
          `Create a ${args.oneTime ? 'one-time' : 'multi-use'} Temporary Access Pass valid ${args.lifetimeMinutes} ` +
          `minutes for ${user.userPrincipalName}, replacing any existing pass. Anyone holding it can sign in as them.`,
      };
    },
    async run(ctx, args) {
      if (args.lifetimeMinutes < 10 || args.lifetimeMinutes > 43200) {
        throw new CliError('USAGE', '--lifetime-minutes must be between 10 and 43200');
      }
      const user = await resolveUser(ctx, args.user);
      const pass = await ctx.graph.post<Record<string, unknown>>(
        `/users/${seg(user.id)}/authentication/temporaryAccessPassMethods`,
        { lifetimeInMinutes: args.lifetimeMinutes, isUsableOnce: args.oneTime },
      );
      const start = (pass.startDateTime as string | undefined) ?? (pass.createdDateTime as string | undefined);
      return compact({
        user: identity(user),
        temporaryAccessPass: pass.temporaryAccessPass as string,
        startDateTime: start,
        expiresDateTime: start
          ? new Date(Date.parse(start) + (pass.lifetimeInMinutes as number) * 60_000).toISOString()
          : undefined,
        lifetimeInMinutes: pass.lifetimeInMinutes as number,
        isUsableOnce: pass.isUsableOnce as boolean,
        isUsable: pass.isUsable as boolean,
        methodUsabilityReason: pass.methodUsabilityReason as string | undefined,
        next: 'User signs in at https://aka.ms/mysecurityinfo with the pass and registers a new method.',
      });
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'users auth-method remove',
    summary: 'Remove one MFA/sign-in method from a user (e.g. an old phone)',
    description:
      'Get the method id from `users auth-methods`. Typical "new phone" flow: remove the old Authenticator and ' +
      'phone methods, then `users tap` so they can register again. The password cannot be removed. Removing the ' +
      "default method can fail until it is the user's last of that kind.",
    examples: ['m365 entra users auth-method remove jane@contoso.com 3179e48a-750b-4051-897c-87b9720928f7'],
    args: [
      USER_ARG,
      {
        name: 'methodId',
        type: 'string',
        required: true,
        positional: true,
        description: 'Method id from auth-methods.',
      },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['UserAuthenticationMethod.ReadWrite.All', 'User.Read.All'],
      application: ['UserAuthenticationMethod.ReadWrite.All', 'User.Read.All'],
      roles: ['Authentication Administrator', 'Privileged Authentication Administrator'],
    },
    async preview(ctx, args) {
      const { user, methods, method } = await findAuthMethod(ctx, args.user, args.methodId);
      const remainingStrong = methods.filter(
        (m) => m.id !== method.id && AUTH_METHOD_TYPES[m['@odata.type']]?.mfa === true,
      ).length;
      return compact({
        user: identity(user),
        method: shapeAuthMethod(method),
        action: `Remove this ${shapeAuthMethod(method).type} method from ${user.userPrincipalName}`,
        warning: remainingStrong
          ? undefined
          : 'This leaves no MFA method: the user must register again (issue a TAP with `users tap`).',
      });
    },
    async run(ctx, args) {
      const { user, method, segment } = await findAuthMethod(ctx, args.user, args.methodId);
      // External MFA methods are removed through their $ref link (per the v1.0 API).
      const suffix = segment === 'externalAuthenticationMethods' ? '/$ref' : '';
      await ctx.graph.delete(`/users/${seg(user.id)}/authentication/${segment}/${seg(method.id)}${suffix}`);
      return { user: identity(user), removed: shapeAuthMethod(method) };
    },
  }),
];
