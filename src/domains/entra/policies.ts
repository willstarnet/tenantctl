import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { requireApprovedTarget } from '../../core/safety.ts';
import { type AnyCommand, defineCommand } from '../../core/types.ts';
import { compact, isSettledError, settle } from '../../core/util.ts';
import { CA_STATE, type CaPolicy, resolveCaPolicy, shapeCaPolicy } from './directory.ts';

/**
 * Tenant-wide Entra policies an admin changes during a security baseline: Conditional Access state, named locations,
 * and the authorization / consent / authentication-method settings behind SOC 2 identity controls.
 */

const ALL = Number.POSITIVE_INFINITY;
const CA_PATH = '/identity/conditionalAccess/policies';
const LOCATIONS_PATH = '/identity/conditionalAccess/namedLocations';

/** CLI state -> Graph conditionalAccessPolicyState. */
const STATES = { on: 'enabled', off: 'disabled', 'report-only': 'enabledForReportingButNotEnforced' } as const;

// ---- IP ranges --------------------------------------------------------------------------------------------------

/** An IPv4 or IPv6 address as a number plus its bit width, or null when it is not an address. */
function parseIp(text: string): { value: bigint; bits: 32 | 128 } | null {
  const value = text.trim();
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(value)) {
    const parts = value.split('.').map(Number);
    if (parts.some((p) => p > 255)) return null;
    return { value: parts.reduce((n, p) => (n << 8n) + BigInt(p), 0n), bits: 32 };
  }
  if (!value.includes(':') || !/^[0-9a-f:]+$/i.test(value)) return null;
  const halves = value.split('::');
  if (halves.length > 2) return null;
  const head = halves[0] ? halves[0].split(':') : [];
  const tail = halves.length === 2 && halves[1] ? halves[1].split(':') : [];
  const missing = 8 - head.length - tail.length;
  if (missing < 0 || (halves.length === 1 && missing !== 0)) return null;
  const groups = [...head, ...Array<string>(missing).fill('0'), ...tail];
  if (groups.some((g) => !/^[0-9a-f]{1,4}$/i.test(g))) return null;
  return { value: groups.reduce((n, g) => (n << 16n) + BigInt(Number.parseInt(g, 16)), 0n), bits: 128 };
}

/** Whether `ip` falls inside `cidr` (same address family only). */
export function cidrContains(cidr: string, ip: string): boolean {
  const [base, prefixText] = cidr.split('/');
  const network = parseIp(base ?? '');
  const address = parseIp(ip);
  if (!network || !address || network.bits !== address.bits) return false;
  const prefix = prefixText === undefined ? network.bits : Number(prefixText);
  if (!Number.isInteger(prefix) || prefix < 0 || prefix > network.bits) return false;
  const shift = BigInt(network.bits - prefix);
  return network.value >> shift === address.value >> shift;
}

interface NamedLocation {
  id: string;
  displayName?: string | null;
  '@odata.type'?: string;
  isTrusted?: boolean | null;
  ipRanges?: { cidrAddress?: string | null }[] | null;
  countriesAndRegions?: string[] | null;
  includeUnknownCountriesAndRegions?: boolean | null;
  countryLookupMethod?: string | null;
  modifiedDateTime?: string | null;
}

// ---- Tenant identity settings ---------------------------------------------------------------------------------

/** Built-in guest user role ids (authorizationPolicy.guestUserRoleId). */
const GUEST_ROLES: Record<string, string> = {
  'a0b1b346-4d3e-4e8b-98f8-753987be4970': 'same as members',
  '10dae51f-b6af-4016-8d66-8c2a99b929b3': 'limited (default)',
  '2af84b1e-32c8-42b7-82bc-daa82404023b': 'restricted',
};

/** Methods that are phishable and that Microsoft recommends retiring for sign-in. */
const WEAK_METHODS: Record<string, true> = { Sms: true, Voice: true };

interface AuthorizationPolicy {
  allowInvitesFrom?: string | null;
  allowedToSignUpEmailBasedSubscriptions?: boolean | null;
  allowEmailVerifiedUsersToJoinOrganization?: boolean | null;
  allowedToUseSSPR?: boolean | null;
  blockMsolPowerShell?: boolean | null;
  guestUserRoleId?: string | null;
  defaultUserRolePermissions?: {
    allowedToCreateApps?: boolean | null;
    allowedToCreateSecurityGroups?: boolean | null;
    allowedToCreateTenants?: boolean | null;
    allowedToReadOtherUsers?: boolean | null;
    permissionGrantPoliciesAssigned?: string[] | null;
  } | null;
}

interface AuthMethodsPolicy {
  authenticationMethodConfigurations?: { id: string; state?: string | null }[] | null;
  registrationEnforcement?: { authenticationMethodsRegistrationCampaign?: { state?: string | null } | null } | null;
  systemCredentialPreferences?: { state?: string | null } | null;
}

interface ConsentRequestPolicy {
  isEnabled?: boolean | null;
  notifyReviewers?: boolean | null;
  requestDurationInDays?: number | null;
  reviewers?: unknown[] | null;
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'ca set-state',
    summary: 'Turn a Conditional Access policy on, off or to report-only',
    description:
      'Moves a policy between report-only (logs what it would do), on (enforced) and off. Check report-only results ' +
      'before turning a policy on: sign-ins show what it would have blocked (`tenantctl entra signins` and the ' +
      "policy's report-only result in the Entra sign-in log). The preview shows who and what the policy covers and " +
      'warns when an enforced policy blocks every user. Already in the requested state: nothing is sent.',
    examples: [
      'tenantctl entra ca set-state "CA010-AllUsers-DeviceCodeFlow-Block" --state on',
      'tenantctl entra ca set-state "Legacy MFA policy" --state off',
    ],
    args: [
      { name: 'policy', type: 'string', required: true, positional: true, description: 'Policy id or name.' },
      {
        name: 'state',
        type: 'string',
        required: true,
        choices: Object.keys(STATES),
        description: 'on (enforced), off, or report-only.',
      },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['Policy.Read.All', 'Policy.ReadWrite.ConditionalAccess'],
      application: ['Policy.Read.All', 'Policy.ReadWrite.ConditionalAccess'],
      roles: ['Conditional Access Administrator', 'Security Administrator'],
    },
    async preview(ctx, args) {
      const policy = await resolveCaPolicy(ctx, args.policy);
      const wanted = STATES[args.state as keyof typeof STATES];
      const shaped = shapeCaPolicy(policy);
      const warnings: string[] = [];
      const everyone = policy.conditions?.users?.includeUsers?.includes('All');
      const blocks = policy.grantControls?.builtInControls?.includes('block');
      if (wanted === 'enabled' && everyone && blocks) {
        warnings.push('Blocks sign-in for every user it matches; make sure break-glass accounts are excluded.');
      }
      if (wanted === 'enabled' && everyone && !policy.conditions?.users?.excludeUsers?.length) {
        warnings.push('Applies to all users with no user excluded.');
      }
      const from = CA_STATE[policy.state] ?? policy.state;
      return {
        action:
          policy.state === wanted
            ? `No change: "${policy.displayName}" is already ${from}`
            : `Change Conditional Access policy "${policy.displayName}" from ${from} to ${CA_STATE[wanted]}`,
        targetId: policy.id,
        warnings,
        policy: shaped,
      };
    },
    async run(ctx, args, approved) {
      const policy = await resolveCaPolicy(ctx, args.policy);
      requireApprovedTarget(approved, policy.id, 'Conditional Access policy');
      const wanted = STATES[args.state as keyof typeof STATES];
      const from = CA_STATE[policy.state] ?? policy.state;
      if (policy.state === wanted) return { status: 'unchanged', id: policy.id, name: policy.displayName, state: from };
      await ctx.graph.patch(`${CA_PATH}/${seg(policy.id)}`, { state: wanted });
      return {
        status: 'updated',
        id: policy.id,
        name: policy.displayName,
        from,
        to: CA_STATE[wanted],
        next: 'Watch the effect with `tenantctl entra signins --failures-only --days 1`.',
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'ca locations',
    summary: 'List Conditional Access named locations with their IP ranges or countries, and check an IP',
    description:
      'type is ip (CIDR ranges, trusted or not) or country. usedBy names the Conditional Access policies that ' +
      'include or exclude the location. --ip answers "is this sign-in address one of ours?": only locations whose ' +
      'ranges contain it are returned, with the matching range. A public NAT address is what sign-ins show, so ' +
      'check the address from the sign-in log, not the device.',
    examples: ['tenantctl entra ca locations', 'tenantctl entra ca locations --ip 203.0.113.10'],
    args: [{ name: 'ip', type: 'string', description: 'Only locations whose IP ranges contain this address.' }],
    risk: 'read',
    permissions: {
      delegated: ['Policy.Read.All'],
      application: ['Policy.Read.All'],
      roles: ['Security Reader', 'Conditional Access Administrator'],
    },
    async run(ctx, args) {
      if (args.ip !== undefined && !parseIp(args.ip)) {
        throw new CliError('USAGE', `--ip "${args.ip}" is not an IPv4 or IPv6 address`);
      }
      const [locations, policies] = await Promise.all([
        ctx.graph.list<NamedLocation>(LOCATIONS_PATH, { limit: ALL }),
        settle({ list: ctx.graph.list<CaPolicy>(CA_PATH, { limit: ALL }) }),
      ]);
      const usedBy = (id: string) =>
        isSettledError(policies.list)
          ? undefined
          : policies.list.items
              .filter(
                (p) =>
                  p.conditions?.locations?.includeLocations?.includes(id) ||
                  p.conditions?.locations?.excludeLocations?.includes(id),
              )
              .map((p) => ({
                name: p.displayName,
                state: CA_STATE[p.state] ?? p.state,
                use: p.conditions?.locations?.excludeLocations?.includes(id) ? 'exclude' : 'include',
              }));
      const items = locations.items.flatMap<unknown>((l) => {
        const ranges = (l.ipRanges ?? []).flatMap((r) => (r.cidrAddress ? [r.cidrAddress] : []));
        const matched = args.ip ? ranges.filter((r) => cidrContains(r, args.ip as string)) : undefined;
        if (matched && !matched.length) return [];
        if (ctx.full) return [l];
        const country = (l['@odata.type'] ?? '').endsWith('countryNamedLocation');
        return [
          compact({
            id: l.id,
            name: l.displayName ?? null,
            type: country ? 'country' : 'ip',
            trusted: country ? undefined : (l.isTrusted ?? false),
            ranges: country ? undefined : ranges,
            matchedRange: matched?.[0],
            countries: country ? (l.countriesAndRegions ?? []) : undefined,
            includeUnknownCountries: country ? (l.includeUnknownCountriesAndRegions ?? false) : undefined,
            usedBy: usedBy(l.id),
            modifiedDateTime: l.modifiedDateTime ?? null,
          }),
        ];
      });
      return { items, truncated: false };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'policies',
    summary: 'Tenant identity settings: security defaults, user consent, guest invites, sign-up, auth methods',
    description:
      'The Entra settings a security baseline checks, in one view. authorization: who can invite guests, whether ' +
      'users can consent to apps (userConsent lists the consent policies; empty = consent off), self-service ' +
      'sign-up and what members may create. adminConsentWorkflow: whether users can request admin approval. ' +
      "authenticationMethods: each method's state, the registration campaign and system-preferred MFA. warnings " +
      "lists settings that differ from Microsoft's recommendations (SMS/voice on, user consent on, no admin " +
      'consent workflow, anyone can invite guests, email sign-up on). Each section degrades to { error } on its own; ' +
      'authentication methods need Policy.Read.AuthenticationMethod and Global Reader or Authentication Policy ' +
      'Administrator.',
    examples: ['tenantctl entra policies'],
    args: [],
    risk: 'read',
    permissions: {
      delegated: ['Policy.Read.All', 'Policy.Read.AuthenticationMethod'],
      application: ['Policy.Read.All', 'Policy.Read.AuthenticationMethod'],
      roles: ['Global Reader'],
    },
    async run(ctx) {
      const raw = await settle({
        securityDefaults: ctx.graph.get<{ isEnabled?: boolean }>('/policies/identitySecurityDefaultsEnforcementPolicy'),
        authorization: ctx.graph.get<AuthorizationPolicy>('/policies/authorizationPolicy'),
        adminConsentWorkflow: ctx.graph.get<ConsentRequestPolicy>('/policies/adminConsentRequestPolicy'),
        authenticationMethods: ctx.graph.get<AuthMethodsPolicy>('/policies/authenticationMethodsPolicy'),
      });
      if (ctx.full) return raw;
      const warnings: string[] = [];
      const authorization = isSettledError(raw.authorization)
        ? raw.authorization
        : (() => {
            const a = raw.authorization;
            const grants = a.defaultUserRolePermissions?.permissionGrantPoliciesAssigned ?? [];
            const consent = grants
              .filter((g) => g.startsWith('ManagePermissionGrantsForSelf.'))
              .map((g) => g.slice('ManagePermissionGrantsForSelf.'.length));
            if (consent.length) {
              warnings.push(
                `Users can consent to apps themselves (${consent.join(', ')}); Microsoft recommends "Do not allow ` +
                  'user consent" with the admin consent workflow.',
              );
            }
            if (a.allowInvitesFrom === 'everyone') warnings.push('Anyone, including guests, can invite guests.');
            if (a.allowedToSignUpEmailBasedSubscriptions) {
              warnings.push('Users can sign up for email-verified (self-service) subscriptions.');
            }
            return compact({
              guestInvitesFrom: a.allowInvitesFrom ?? null,
              guestAccess: a.guestUserRoleId ? (GUEST_ROLES[a.guestUserRoleId] ?? a.guestUserRoleId) : null,
              userConsent: consent,
              emailSignUpSubscriptions: a.allowedToSignUpEmailBasedSubscriptions ?? null,
              emailVerifiedUsersCanJoin: a.allowEmailVerifiedUsersToJoinOrganization ?? null,
              selfServicePasswordReset: a.allowedToUseSSPR ?? null,
              usersCanCreateApps: a.defaultUserRolePermissions?.allowedToCreateApps ?? null,
              usersCanCreateSecurityGroups: a.defaultUserRolePermissions?.allowedToCreateSecurityGroups ?? null,
              usersCanCreateTenants: a.defaultUserRolePermissions?.allowedToCreateTenants ?? null,
              usersCanReadOtherUsers: a.defaultUserRolePermissions?.allowedToReadOtherUsers ?? null,
              blockMsolPowerShell: a.blockMsolPowerShell ?? undefined,
            });
          })();
      const adminConsentWorkflow = isSettledError(raw.adminConsentWorkflow)
        ? raw.adminConsentWorkflow
        : (() => {
            const p = raw.adminConsentWorkflow;
            if (!p.isEnabled) warnings.push('The admin consent workflow is off: users cannot ask an admin for an app.');
            return {
              enabled: p.isEnabled ?? false,
              reviewers: p.reviewers?.length ?? 0,
              requestDurationInDays: p.requestDurationInDays || null,
              notifyReviewers: p.notifyReviewers ?? null,
            };
          })();
      const authenticationMethods = isSettledError(raw.authenticationMethods)
        ? raw.authenticationMethods
        : (() => {
            const p = raw.authenticationMethods;
            const methods = Object.fromEntries(
              (p.authenticationMethodConfigurations ?? []).map((m) => [m.id, m.state ?? null]),
            );
            const weak = Object.keys(WEAK_METHODS).filter((id) => methods[id] === 'enabled');
            if (weak.length) {
              warnings.push(`${weak.join(' and ')} sign-in is enabled; these methods are phishable.`);
            }
            return compact({
              methods,
              registrationCampaign: p.registrationEnforcement?.authenticationMethodsRegistrationCampaign?.state ?? null,
              systemPreferredMfa: p.systemCredentialPreferences?.state ?? undefined,
            });
          })();
      return {
        securityDefaults: isSettledError(raw.securityDefaults)
          ? raw.securityDefaults
          : { enabled: raw.securityDefaults.isEnabled ?? null },
        authorization,
        adminConsentWorkflow,
        authenticationMethods,
        warnings,
      };
    },
  }),
];
