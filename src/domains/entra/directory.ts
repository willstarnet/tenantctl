import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { pickOne, resolveGroup, resolveUser } from '../../core/resolve.ts';
import { SIGNIN_ERRORS, signinErrorInfo, signinErrorLookupUrl } from '../../core/signinErrors.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact, daysAgo, daysSince, isGuid, settle } from '../../core/util.ts';
import { type SignIn, shapeSignin } from './users.ts';

const ALL = Number.POSITIVE_INFINITY;
const DAY_MS = 86_400_000;

// ---------------------------------------------------------------------------------------------------------------
// Groups
// ---------------------------------------------------------------------------------------------------------------

export interface GraphGroup {
  id: string;
  displayName: string;
  description?: string | null;
  mail?: string | null;
  mailNickname?: string | null;
  groupTypes?: string[];
  securityEnabled?: boolean;
  mailEnabled?: boolean;
  membershipRule?: string | null;
  membershipRuleProcessingState?: string | null;
  createdDateTime?: string | null;
  resourceProvisioningOptions?: string[];
  visibility?: string | null;
  isAssignableToRole?: boolean | null;
  onPremisesSyncEnabled?: boolean | null;
  assignedLicenses?: { skuId: string; disabledPlans?: string[] }[];
}

const GROUP_SELECT =
  'id,displayName,description,mail,mailNickname,groupTypes,securityEnabled,mailEnabled,membershipRule,' +
  'membershipRuleProcessingState,createdDateTime,resourceProvisioningOptions,visibility,onPremisesSyncEnabled';

/** m365 = Microsoft 365 (Unified) group; mail-security = mail-enabled security group (Exchange-managed). */
export type GroupKind = 'm365' | 'security' | 'distribution' | 'mail-security';

export function groupKind(g: Pick<GraphGroup, 'groupTypes' | 'mailEnabled' | 'securityEnabled'>): GroupKind {
  if (g.groupTypes?.includes('Unified')) return 'm365';
  if (g.mailEnabled && g.securityEnabled) return 'mail-security';
  if (g.mailEnabled) return 'distribution';
  return 'security';
}

export function shapeGroup(g: GraphGroup) {
  const dynamic = g.groupTypes?.includes('DynamicMembership') ?? false;
  return compact({
    id: g.id,
    displayName: g.displayName,
    type: groupKind(g),
    membership: dynamic ? 'dynamic' : 'assigned',
    mail: g.mail ?? null,
    membershipRule: dynamic ? (g.membershipRule ?? null) : undefined,
    createdDateTime: g.createdDateTime ?? null,
    isTeam: g.resourceProvisioningOptions?.includes('Team') ?? false,
    syncedFromOnPrem: g.onPremisesSyncEnabled === true,
  });
}

const GROUP_TYPE_FILTERS: Record<string, string> = {
  security: 'securityEnabled eq true and mailEnabled eq false',
  m365: "groupTypes/any(t:t eq 'Unified')",
  dynamic: "groupTypes/any(t:t eq 'DynamicMembership')",
  distribution: "mailEnabled eq true and securityEnabled eq false and NOT groupTypes/any(t:t eq 'Unified')",
  'mail-security': 'mailEnabled eq true and securityEnabled eq true',
};

interface DirectoryObject {
  '@odata.type'?: string;
  id: string;
  displayName?: string | null;
  userPrincipalName?: string | null;
  mail?: string | null;
  accountEnabled?: boolean | null;
  deviceId?: string | null;
  operatingSystem?: string | null;
  appId?: string | null;
}

const MEMBER_SELECT = 'id,displayName,userPrincipalName,mail,accountEnabled,deviceId,operatingSystem,appId';

function shapeDirectoryObject(o: DirectoryObject) {
  return compact({
    type: o['@odata.type']?.replace('#microsoft.graph.', '') ?? 'unknown',
    id: o.id,
    displayName: o.displayName ?? null,
    userPrincipalName: o.userPrincipalName ?? undefined,
    mail: o.mail ?? undefined,
    accountEnabled: o.accountEnabled ?? undefined,
    deviceId: o.deviceId ?? undefined,
    operatingSystem: o.operatingSystem ?? undefined,
    appId: o.appId ?? undefined,
  });
}

export interface ResolvedMember {
  id: string;
  type: 'user' | 'group' | 'device';
  displayName: string;
  /** UPN, mail or deviceId: whatever identifies it to a person. */
  detail: string | null;
}

const MEMBER_KINDS: Record<string, ResolvedMember['type']> = { user: 'user', group: 'group', device: 'device' };

/**
 * A group member can be a user, group or device. `user:`, `group:` or `device:` prefixes force the kind; otherwise
 * ids are looked up in each collection and names matched exactly across all three before a fuzzy user search.
 */
export async function resolveMember(ctx: Context, ref: string): Promise<ResolvedMember> {
  const prefixed = /^(user|group|device):(.+)$/i.exec(ref.trim());
  const kind = prefixed ? MEMBER_KINDS[(prefixed[1] as string).toLowerCase()] : undefined;
  const value = (prefixed ? (prefixed[2] as string) : ref).trim();
  if (!value) throw new CliError('USAGE', 'A member is required');

  if (kind === 'user') return userMember(await resolveUser(ctx, value));
  if (kind === 'group') return groupMember(await resolveGroup(ctx, value));
  if (kind === 'device') return deviceMember(await resolveEntraDevice(ctx, value));

  if (isGuid(value)) {
    const [user, group] = await Promise.all([
      ctx.graph.tryGet<{ id: string; displayName: string; userPrincipalName: string }>(`/users/${seg(value)}`, {
        query: { $select: 'id,displayName,userPrincipalName' },
      }),
      ctx.graph.tryGet<GraphGroup>(`/groups/${seg(value)}`, { query: { $select: 'id,displayName,mail' } }),
    ]);
    if (user) return userMember(user);
    if (group) return groupMember(group);
    return deviceMember(await resolveEntraDevice(ctx, value));
  }

  if (value.includes('@')) {
    try {
      return userMember(await resolveUser(ctx, value));
    } catch (error) {
      if (error instanceof CliError && error.code === 'NOT_FOUND') return groupMember(await resolveGroup(ctx, value));
      throw error;
    }
  }

  const name = odata(value);
  const [users, groups, devices] = await Promise.all([
    ctx.graph.list<{ id: string; displayName: string; userPrincipalName: string }>('/users', {
      query: { $filter: `displayName eq ${name}`, $select: 'id,displayName,userPrincipalName' },
      limit: 10,
    }),
    ctx.graph.list<GraphGroup>('/groups', {
      query: { $filter: `displayName eq ${name}`, $select: 'id,displayName,mail' },
      limit: 10,
    }),
    ctx.graph.list<EntraDevice>('/devices', {
      query: { $filter: `displayName eq ${name}`, $select: 'id,displayName,deviceId' },
      limit: 10,
    }),
  ]);
  const candidates = [
    ...users.items.map(userMember),
    ...groups.items.map(groupMember),
    ...devices.items.map(deviceMember),
  ];
  if (!candidates.length) return userMember(await resolveUser(ctx, value));
  if (candidates.length > 1) {
    throw new CliError('AMBIGUOUS', `"${value}" matches ${candidates.length} directory objects`, {
      hint: 'Ask which one is meant, then pass its id or prefix the name with user:, group: or device:.',
      details: candidates,
    });
  }
  return candidates[0] as ResolvedMember;
}

function userMember(u: { id: string; displayName: string; userPrincipalName?: string | null }): ResolvedMember {
  return { id: u.id, type: 'user', displayName: u.displayName, detail: u.userPrincipalName ?? null };
}

function groupMember(g: { id: string; displayName: string; mail?: string | null }): ResolvedMember {
  return { id: g.id, type: 'group', displayName: g.displayName, detail: g.mail ?? null };
}

function deviceMember(d: { id: string; displayName: string; deviceId?: string | null }): ResolvedMember {
  return { id: d.id, type: 'device', displayName: d.displayName, detail: d.deviceId ?? null };
}

async function skuPartNumbers(ctx: Context): Promise<Record<string, string>> {
  const { items } = await ctx.graph.list<SubscribedSku>('/subscribedSkus', {
    query: { $select: 'skuId,skuPartNumber' },
    limit: ALL,
    pageSize: false,
  });
  return Object.fromEntries(items.map((s) => [s.skuId, s.skuPartNumber]));
}

const groupArg = {
  name: 'group',
  type: 'string',
  required: true,
  positional: true,
  description: 'Group name, mail address or object id.',
} as const;

const groupCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'groups list',
    summary: 'List or search groups, with derived type (security/m365/distribution) and Teams flag',
    description:
      'Use to find a group or survey groups of one kind. `type` is derived: m365 (Microsoft 365 group), security, ' +
      'distribution, mail-security (mail-enabled security; members managed in Exchange). `membership` is ' +
      'dynamic or assigned; dynamic groups show their rule. `isTeam` means the group backs a Microsoft Team. ' +
      '`syncedFromOnPrem` groups must be changed in on-premises AD. Next: `tenantctl entra groups get <group>`.',
    examples: ['tenantctl entra groups list --search finance', 'tenantctl entra groups list --type dynamic --all'],
    args: [
      { name: 'search', type: 'string', description: 'Words in the display name.' },
      {
        name: 'type',
        type: 'string',
        choices: ['security', 'm365', 'dynamic', 'distribution', 'mail-security'],
        description: 'Only groups of this kind.',
      },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['Group.Read.All'], application: ['Group.Read.All'] },
    async run(ctx, args) {
      const term = args.search?.replace(/"/g, '');
      const r = await ctx.graph.list<GraphGroup>('/groups', {
        query: {
          $select: GROUP_SELECT,
          $filter: args.type ? GROUP_TYPE_FILTERS[args.type] : undefined,
          $search: term ? `"displayName:${term}"` : undefined,
          $count: true,
        },
        advanced: true,
        limit: listLimit(args),
      });
      return ctx.full ? r : { items: r.items.map(shapeGroup), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'groups get',
    summary: 'Show a group with owners, member count, group-assigned licenses and dynamic rule state',
    description:
      'Use to explain what a group is and what it grants. `licenses` lists SKUs assigned through group-based ' +
      'licensing (Entra ID P1). For dynamic groups, `membershipRuleProcessingState` Paused means membership is ' +
      'frozen. `memberCount` counts direct members. Next: `tenantctl entra groups members <group>`.',
    examples: ['tenantctl entra groups get "Finance Team"', 'tenantctl entra groups get sales@contoso.com'],
    args: [groupArg],
    risk: 'read',
    permissions: {
      delegated: ['Group.Read.All', 'GroupMember.Read.All', 'LicenseAssignment.Read.All'],
      application: ['Group.Read.All', 'GroupMember.Read.All', 'LicenseAssignment.Read.All'],
    },
    async run(ctx, args) {
      const { id } = await resolveGroup(ctx, args.group);
      const group = await ctx.graph.get<GraphGroup>(`/groups/${seg(id)}`, {
        query: { $select: `${GROUP_SELECT},assignedLicenses,isAssignableToRole` },
      });
      const licensed = group.assignedLicenses?.length ?? 0;
      const extra = await settle({
        owners: ctx.graph
          .list<DirectoryObject>(`/groups/${seg(id)}/owners`, {
            query: { $select: 'id,displayName,userPrincipalName,mail' },
            limit: ALL,
          })
          .then((r) => r.items),
        memberCount: ctx.graph
          .get<{ '@odata.count'?: number }>(`/groups/${seg(id)}/members`, {
            query: { $count: true, $top: 1, $select: 'id' },
            advanced: true,
          })
          .then((r) => r['@odata.count'] ?? null),
        skus: licensed ? skuPartNumbers(ctx) : Promise.resolve({} as Record<string, string>),
      });
      if (ctx.full) return { ...group, owners: extra.owners, memberCount: extra.memberCount };
      const skus = 'error' in extra.skus ? {} : extra.skus;
      return compact({
        ...shapeGroup(group),
        description: group.description ?? null,
        visibility: group.visibility ?? null,
        isAssignableToRole: group.isAssignableToRole ?? false,
        membershipRuleProcessingState: group.groupTypes?.includes('DynamicMembership')
          ? (group.membershipRuleProcessingState ?? null)
          : undefined,
        memberCount: extra.memberCount,
        owners: Array.isArray(extra.owners)
          ? extra.owners.map((o) =>
              compact({ displayName: o.displayName ?? null, userPrincipalName: o.userPrincipalName ?? undefined }),
            )
          : extra.owners,
        licenses: (group.assignedLicenses ?? []).map((l) =>
          compact({
            sku: skus[l.skuId] ?? l.skuId,
            disabledPlans: l.disabledPlans?.length ? l.disabledPlans.length : undefined,
          }),
        ),
        licensesError: 'error' in extra.skus ? extra.skus.error : undefined,
      });
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'groups members',
    summary: 'List group members (users, devices, nested groups); --transitive flattens nesting',
    description:
      'Each item has `type` (user, device, group, servicePrincipal, orgContact). Use --transitive to answer "is ' +
      'this user in the group through a nested group?". Hidden-membership groups also need Member.Read.Hidden.',
    examples: [
      'tenantctl entra groups members "Finance Team"',
      'tenantctl entra groups members "All Laptops" --transitive --all',
    ],
    args: [
      groupArg,
      { name: 'transitive', type: 'boolean', description: 'Include members of nested groups.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['GroupMember.Read.All'], application: ['GroupMember.Read.All'] },
    async run(ctx, args) {
      const group = await resolveGroup(ctx, args.group);
      const r = await ctx.graph.list<DirectoryObject>(
        `/groups/${seg(group.id)}/${args.transitive ? 'transitiveMembers' : 'members'}`,
        { query: { $select: MEMBER_SELECT }, limit: listLimit(args) },
      );
      return ctx.full ? r : { items: r.items.map(shapeDirectoryObject), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'groups add-member',
    summary: 'Add one or more users, devices or groups to a security or Microsoft 365 group',
    description:
      'Members may be UPNs, names, device names or ids; prefix with user:, group: or device: to disambiguate. ' +
      'Each member is reported as added, alreadyMember or failed. Distribution and mail-enabled security groups ' +
      'cannot be changed through Graph (use Exchange Online). Adding devices also needs Device.Read.All ' +
      '(delegated) / Device.ReadWrite.All (application); role-assignable groups need Privileged Role Administrator.',
    examples: [
      'tenantctl entra groups add-member "Finance Team" jane@contoso.com bob@contoso.com',
      'tenantctl entra groups add-member "Autopilot Devices" device:DESKTOP-123',
    ],
    args: [
      groupArg,
      {
        name: 'members',
        type: 'string[]',
        required: true,
        positional: true,
        description: 'Members to add (user, device or group; comma separated or repeated).',
      },
    ],
    risk: 'write',
    permissions: {
      delegated: ['GroupMember.ReadWrite.All'],
      application: ['GroupMember.ReadWrite.All'],
      roles: ['Groups Administrator'],
    },
    async preview(ctx, args) {
      const group = await resolveGroup(ctx, args.group);
      const members = await Promise.all(args.members.map((m) => resolveMember(ctx, m)));
      return { group, members, action: `Add ${members.length} member(s) to group "${group.displayName}".` };
    },
    async run(ctx, args) {
      const group = await resolveGroup(ctx, args.group);
      const members = await Promise.all(args.members.map((m) => resolveMember(ctx, m)));
      const results = [];
      // One request per member so an existing member doesn't abort the others (a PATCH batch is all-or-nothing).
      for (const member of members) {
        try {
          await ctx.graph.post(`/groups/${seg(group.id)}/members/$ref`, {
            '@odata.id': ctx.graph.url(`/directoryObjects/${seg(member.id)}`),
          });
          results.push({ ...member, status: 'added' });
        } catch (error) {
          if (error instanceof CliError && error.status === 400 && /already exist/i.test(error.message)) {
            results.push({ ...member, status: 'alreadyMember' });
          } else {
            results.push({
              ...member,
              status: 'failed',
              error: error instanceof Error ? error.message : String(error),
            });
          }
        }
      }
      return { group: { id: group.id, displayName: group.displayName }, results };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'groups remove-member',
    summary: 'Remove a user, device or group from a group',
    description:
      'Removes a direct member only; members that come through a nested group or a dynamic rule cannot be removed ' +
      'this way. Distribution and mail-enabled security groups are managed in Exchange Online.',
    examples: ['tenantctl entra groups remove-member "Finance Team" jane@contoso.com'],
    args: [
      groupArg,
      { name: 'member', type: 'string', required: true, positional: true, description: 'Member to remove.' },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['GroupMember.ReadWrite.All'],
      application: ['GroupMember.ReadWrite.All'],
      roles: ['Groups Administrator'],
    },
    async preview(ctx, args) {
      const group = await resolveGroup(ctx, args.group);
      const member = await resolveMember(ctx, args.member);
      return {
        group,
        member,
        action: `Remove ${member.type} "${member.displayName}" from group "${group.displayName}".`,
      };
    },
    async run(ctx, args) {
      const group = await resolveGroup(ctx, args.group);
      const member = await resolveMember(ctx, args.member);
      await ctx.graph.delete(`/groups/${seg(group.id)}/members/${seg(member.id)}/$ref`);
      return { group: { id: group.id, displayName: group.displayName }, member, status: 'removed' };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'groups create',
    summary: 'Create a security or Microsoft 365 group, optionally with an owner and members',
    description:
      'tenantctl groups get a mailbox, SharePoint site and can become a Team; security groups are for access, ' +
      'licensing and Intune targeting. The mail nickname is derived from the name unless --mail-nickname is given. ' +
      'Owner plus members may total at most 20; add more with `groups add-member`. Dynamic groups and ' +
      'distribution lists are not created here.',
    examples: [
      'tenantctl entra groups create --name "Finance Laptops" --type security --members device:FIN-01,device:FIN-02',
      'tenantctl entra groups create --name "Project Phoenix" --type tenantctl --owner jane@contoso.com',
    ],
    args: [
      { name: 'name', type: 'string', required: true, description: 'Display name.' },
      { name: 'type', type: 'string', choices: ['security', 'm365'], default: 'security', description: 'Group kind.' },
      { name: 'description', type: 'string', description: 'Group description.' },
      { name: 'mailNickname', type: 'string', description: 'Mail alias (default: derived from the name).' },
      { name: 'owner', type: 'string', description: 'Owner (user).' },
      { name: 'members', type: 'string[]', description: 'Initial members (users, devices, groups).' },
    ],
    risk: 'write',
    permissions: { delegated: ['Group.ReadWrite.All'], application: ['Group.Create'], roles: ['Groups Administrator'] },
    async preview(ctx, args) {
      const { body, owner, members } = await groupCreateBody(ctx, args);
      return { body, owner, members, action: `Create ${args.type} group "${args.name}".` };
    },
    async run(ctx, args) {
      const { body } = await groupCreateBody(ctx, args);
      const created = await ctx.graph.post<GraphGroup>('/groups', body);
      return ctx.full ? created : shapeGroup(created);
    },
  }),
];

async function groupCreateBody(
  ctx: Context,
  args: {
    name: string;
    type: 'security' | 'm365';
    description?: string;
    mailNickname?: string;
    owner?: string;
    members?: string[];
  },
) {
  const nickname = (args.mailNickname ?? args.name).replace(/[^A-Za-z0-9._-]/g, '').slice(0, 64);
  if (!nickname) {
    throw new CliError('USAGE', 'Could not derive a mail nickname from the name', {
      hint: 'Pass --mail-nickname with letters and digits only.',
    });
  }
  const memberRefs = args.members ?? [];
  if (memberRefs.length + (args.owner ? 1 : 0) > 20) {
    throw new CliError('USAGE', 'At most 20 owners and members can be set when creating a group', {
      hint: 'Create the group with up to 20, then use `tenantctl entra groups add-member` for the rest.',
    });
  }
  const [owner, members] = await Promise.all([
    args.owner ? resolveUser(ctx, args.owner) : Promise.resolve(undefined),
    Promise.all(memberRefs.map((m) => resolveMember(ctx, m))),
  ]);
  const m365 = args.type === 'm365';
  const body = compact({
    displayName: args.name,
    description: args.description,
    mailNickname: nickname,
    groupTypes: m365 ? ['Unified'] : [],
    mailEnabled: m365,
    securityEnabled: !m365,
    'owners@odata.bind': owner ? [ctx.graph.url(`/users/${seg(owner.id)}`)] : undefined,
    'members@odata.bind': members.length
      ? members.map((m) => ctx.graph.url(`/directoryObjects/${seg(m.id)}`))
      : undefined,
  });
  return { body, owner, members };
}

// ---------------------------------------------------------------------------------------------------------------
// Licenses, sign-ins, audit
// ---------------------------------------------------------------------------------------------------------------

interface SubscribedSku {
  skuId: string;
  skuPartNumber: string;
  appliesTo?: string;
  capabilityStatus?: string;
  consumedUnits?: number;
  prepaidUnits?: { enabled?: number; suspended?: number; warning?: number; lockedOut?: number };
}

export function shapeSku(s: SubscribedSku) {
  const enabled = s.prepaidUnits?.enabled ?? 0;
  const consumed = s.consumedUnits ?? 0;
  const available = enabled - consumed;
  return {
    skuPartNumber: s.skuPartNumber,
    skuId: s.skuId,
    status: s.capabilityStatus ?? null,
    enabled,
    consumed,
    available,
    suspended: s.prepaidUnits?.suspended ?? 0,
    warning: s.prepaidUnits?.warning ?? 0,
    appliesTo: s.appliesTo ?? null,
    lowOnLicenses: enabled > 0 && (available <= 2 || available <= enabled * 0.05),
  };
}

interface DirectoryAudit {
  id: string;
  activityDateTime: string;
  activityDisplayName: string;
  category: string;
  result: string;
  resultReason?: string | null;
  loggedByService?: string | null;
  correlationId?: string | null;
  initiatedBy?: {
    user?: {
      id?: string | null;
      displayName?: string | null;
      userPrincipalName?: string | null;
      ipAddress?: string | null;
    } | null;
    app?: { appId?: string | null; displayName?: string | null; servicePrincipalName?: string | null } | null;
  };
  targetResources?: {
    id?: string | null;
    displayName?: string | null;
    type?: string | null;
    userPrincipalName?: string | null;
    modifiedProperties?: { displayName?: string | null }[];
  }[];
}

export function shapeAudit(a: DirectoryAudit) {
  const user = a.initiatedBy?.user;
  const app = a.initiatedBy?.app;
  return compact({
    time: a.activityDateTime,
    activity: a.activityDisplayName,
    category: a.category,
    result: a.result,
    resultReason: a.resultReason || undefined,
    initiatedBy: user?.userPrincipalName || user?.displayName || app?.displayName || app?.servicePrincipalName || null,
    initiatedFromIp: user?.ipAddress || undefined,
    targets: (a.targetResources ?? []).map((t) => {
      const modified = (t.modifiedProperties ?? []).map((p) => p.displayName).filter(Boolean);
      return compact({
        type: t.type ?? undefined,
        name: t.userPrincipalName || t.displayName || t.id || null,
        modified: modified.length ? modified : undefined,
      });
    }),
    service: a.loggedByService ?? undefined,
    correlationId: a.correlationId ?? null,
  });
}

const directoryLogCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'licenses',
    summary: 'Tenant license inventory: purchased, assigned and available per SKU',
    description:
      'Answers "do we have a spare license?". `available` = enabled - consumed. `lowOnLicenses` is true when ' +
      '2 or fewer, or at most 5%, remain. `warning` units are in the renewal grace period; `suspended` units ' +
      'cannot be assigned. Only SKUs with appliesTo=User can be given to people.',
    examples: ['tenantctl entra licenses'],
    args: [],
    risk: 'read',
    permissions: { delegated: ['LicenseAssignment.Read.All'], application: ['LicenseAssignment.Read.All'] },
    async run(ctx) {
      const r = await ctx.graph.list<SubscribedSku>('/subscribedSkus', { limit: ALL, pageSize: false });
      if (ctx.full) return r;
      const items = r.items.map(shapeSku).sort((a, b) => a.skuPartNumber.localeCompare(b.skuPartNumber));
      return { items, truncated: false } satisfies ListResult;
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'signins',
    summary: 'Search tenant-wide sign-in logs by user, app, IP, error code or country',
    description:
      'Use for "why can\'t X sign in", password-spray checks ("failures from one IP") or "who signed in from ' +
      'country Y". Defaults to the last day; logs are kept 30 days. Each item explains its error code; run ' +
      '`tenantctl entra signin-error <code>` for detail. Requires Entra ID P1 (free tenants get no sign-in API). ' +
      'Interactive and non-interactive user sign-ins both appear in v1.0 results only as interactive; large ' +
      'windows can time out, so narrow with filters.',
    examples: [
      'tenantctl entra signins --user jane@contoso.com --failures-only --days 3',
      'tenantctl entra signins --ip 203.0.113.7 --days 7',
      'tenantctl entra signins --error-code 50126 --top 200',
    ],
    args: [
      { name: 'user', type: 'string', description: 'Only this user (UPN, name or id).' },
      { name: 'app', type: 'string', description: 'App id, or start of the app display name.' },
      { name: 'ip', type: 'string', description: 'Client IP; end with "." to match a prefix (e.g. 203.0.113.).' },
      { name: 'errorCode', type: 'string', description: 'Sign-in error code, e.g. 50126 or AADSTS50126.' },
      { name: 'failuresOnly', type: 'boolean', description: 'Only failed sign-ins.' },
      { name: 'country', type: 'string', description: 'Two-letter country code, e.g. US.' },
      { name: 'days', type: 'number', default: 1, description: 'Look back this many days (max 30).' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['AuditLog.Read.All'],
      application: ['AuditLog.Read.All'],
      roles: ['Reports Reader', 'Security Reader'],
    },
    async run(ctx, args) {
      const filters = [`createdDateTime ge ${daysAgo(args.days)}`];
      if (args.user) filters.push(`userId eq ${odata((await resolveUser(ctx, args.user)).id)}`);
      if (args.app) {
        filters.push(
          isGuid(args.app) ? `appId eq ${odata(args.app)}` : `startswith(appDisplayName, ${odata(args.app)})`,
        );
      }
      if (args.ip) {
        filters.push(
          /[.:]$/.test(args.ip) ? `startswith(ipAddress, ${odata(args.ip)})` : `ipAddress eq ${odata(args.ip)}`,
        );
      }
      if (args.errorCode) {
        const code = args.errorCode.replace(/^AADSTS/i, '');
        if (!/^\d+$/.test(code)) throw new CliError('USAGE', `--error-code must be numeric, got "${args.errorCode}"`);
        filters.push(`status/errorCode eq ${code}`);
      }
      if (args.failuresOnly) filters.push('status/errorCode ne 0');
      if (args.country) filters.push(`location/countryOrRegion eq ${odata(args.country.toUpperCase())}`);
      const r = await ctx.graph.list<SignIn>('/auditLogs/signIns', {
        query: { $filter: filters.join(' and ') },
        limit: listLimit(args),
      });
      return ctx.full ? r : { items: r.items.map(shapeSignin), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'signin-error',
    summary: 'Explain an Entra sign-in error code (AADSTS...) and what to do about it',
    description:
      'Offline lookup, no Graph call. Pass the number from a sign-in log `errorCode` or an AADSTS message. ' +
      "`benign` codes are expected noise (e.g. MFA prompts). Unknown codes return Microsoft's lookup URL. " +
      'Without a code, lists the whole catalog.',
    examples: [
      'tenantctl entra signin-error 50126',
      'tenantctl entra signin-error AADSTS53003',
      'tenantctl entra signin-error',
    ],
    args: [{ name: 'code', type: 'string', positional: true, description: 'Error code, e.g. 50126 or AADSTS50126.' }],
    risk: 'read',
    permissions: {},
    async run(_ctx, args) {
      if (!args.code) return { items: [...SIGNIN_ERRORS], truncated: false } satisfies ListResult;
      const code = args.code.trim().replace(/^AADSTS/i, '');
      if (!/^\d+$/.test(code)) {
        throw new CliError('USAGE', `"${args.code}" is not a sign-in error code`, {
          hint: 'Pass the number from the sign-in log, e.g. 50126 or AADSTS50126.',
        });
      }
      const info = signinErrorInfo(code);
      return { ...(info ?? { code: Number(code) }), known: Boolean(info), lookupUrl: signinErrorLookupUrl(code) };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'audit',
    summary: 'Search the Entra directory audit log: who changed what, when',
    description:
      'Use for "who reset Jane\'s password", "who added Bob to Global Admins", "what changed on this app". ' +
      '`initiatedBy` is the admin UPN or app name; `targets` lists changed objects and the property names that ' +
      'changed. --activity is a prefix match on the activity name (e.g. "Add member", "Reset password", "Update ' +
      'user"); --category values include UserManagement, GroupManagement, RoleManagement, ApplicationManagement, ' +
      'Policy, Device. Logs are kept 7 days (free) or 30 days (Entra ID P1/P2).',
    examples: [
      'tenantctl entra audit --target jane@contoso.com --days 14',
      'tenantctl entra audit --category RoleManagement --activity "Add member to role"',
      'tenantctl entra audit --user admin@contoso.com',
    ],
    args: [
      { name: 'days', type: 'number', default: 7, description: 'Look back this many days.' },
      { name: 'user', type: 'string', description: 'Only changes made by this admin (UPN, name or id).' },
      {
        name: 'target',
        type: 'string',
        description: 'Only changes to this object: user UPN, object id, or name prefix.',
      },
      { name: 'category', type: 'string', description: 'Audit category, e.g. UserManagement.' },
      { name: 'activity', type: 'string', description: 'Activity name prefix, e.g. "Reset password".' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['AuditLog.Read.All'],
      application: ['AuditLog.Read.All'],
      roles: ['Reports Reader', 'Security Reader'],
    },
    async run(ctx, args) {
      const filters = [`activityDateTime ge ${daysAgo(args.days)}`];
      if (args.user) {
        const upn = args.user.includes('@') ? args.user : (await resolveUser(ctx, args.user)).userPrincipalName;
        filters.push(`initiatedBy/user/userPrincipalName eq ${odata(upn)}`);
      }
      if (args.target) {
        if (isGuid(args.target)) {
          filters.push(`targetResources/any(t:t/id eq ${odata(args.target)})`);
        } else if (args.target.includes('@')) {
          const user = await resolveUser(ctx, args.target);
          filters.push(`targetResources/any(t:t/id eq ${odata(user.id)})`);
        } else {
          filters.push(`targetResources/any(t:startswith(t/displayName, ${odata(args.target)}))`);
        }
      }
      if (args.category) filters.push(`category eq ${odata(args.category)}`);
      if (args.activity) filters.push(`startswith(activityDisplayName, ${odata(args.activity)})`);
      const r = await ctx.graph.list<DirectoryAudit>('/auditLogs/directoryAudits', {
        query: { $filter: filters.join(' and ') },
        limit: listLimit(args),
      });
      return ctx.full ? r : { items: r.items.map(shapeAudit), truncated: r.truncated };
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Directory roles
// ---------------------------------------------------------------------------------------------------------------

interface RoleDefinitionRef {
  id: string;
  displayName: string;
  templateId?: string | null;
}

interface RoleAssignment {
  id: string;
  principalId: string;
  roleDefinitionId: string;
  directoryScopeId?: string | null;
  principal?: DirectoryObject | null;
  roleDefinition?: RoleDefinitionRef | null;
}

interface RoleEligibilityInstance extends RoleAssignment {
  startDateTime?: string | null;
  endDateTime?: string | null;
  memberType?: string | null;
}

/** Role definition id and template id -> display name, for resolving role ids from other APIs (GDAP etc.). */
export async function roleNames(ctx: Context): Promise<Record<string, string>> {
  const { items } = await ctx.graph.list<RoleDefinitionRef>('/roleManagement/directory/roleDefinitions', {
    limit: ALL,
    pageSize: false,
  });
  const names: Record<string, string> = {};
  for (const d of items) {
    names[d.id] = d.displayName;
    if (d.templateId) names[d.templateId] = d.displayName;
  }
  return names;
}

function scopeOf(scope: string | null | undefined): string | undefined {
  return scope && scope !== '/' ? scope : undefined;
}

const roleCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'roles list',
    summary: 'Who holds each active Entra admin role (Global Administrator etc.)',
    description:
      'Use for admin reviews ("who are our Global Admins?"). Shows active assignments, including PIM roles that ' +
      'are currently activated, grouped by role; PIM eligible-but-inactive assignments are not listed (see ' +
      '`roles user`). Members can be users, groups (role-assignable) or service principals. `scope` appears ' +
      'only when the assignment is limited to an administrative unit or object.',
    examples: ['tenantctl entra roles list', 'tenantctl entra roles list --role "Global Administrator"'],
    args: [{ name: 'role', type: 'string', description: 'Only roles whose name contains this text.' }],
    risk: 'read',
    permissions: { delegated: ['RoleManagement.Read.Directory'], application: ['RoleManagement.Read.Directory'] },
    async run(ctx, args) {
      const [assignments, names] = await Promise.all([
        ctx.graph.list<RoleAssignment>('/roleManagement/directory/roleAssignments', {
          query: { $expand: 'principal' },
          limit: ALL,
          pageSize: false,
        }),
        roleNames(ctx),
      ]);
      if (ctx.full) return assignments;
      const byRole: Record<string, { role: string; roleId: string; members: Record<string, unknown>[] }> = {};
      for (const a of assignments.items) {
        const role = names[a.roleDefinitionId] ?? a.roleDefinitionId;
        byRole[a.roleDefinitionId] ??= { role, roleId: a.roleDefinitionId, members: [] };
        byRole[a.roleDefinitionId]?.members.push(
          compact({
            type: a.principal?.['@odata.type']?.replace('#microsoft.graph.', '') ?? 'unknown',
            displayName: a.principal?.displayName ?? a.principalId,
            userPrincipalName: a.principal?.userPrincipalName ?? undefined,
            scope: scopeOf(a.directoryScopeId),
          }),
        );
      }
      const wanted = args.role?.toLowerCase();
      const items = Object.values(byRole)
        .filter((r) => !wanted || r.role.toLowerCase().includes(wanted))
        .map((r) => ({ role: r.role, roleId: r.roleId, memberCount: r.members.length, members: r.members }))
        .sort((a, b) => a.role.localeCompare(b.role));
      return { items, truncated: false } satisfies ListResult;
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'roles user',
    summary: "A user's active admin roles and PIM-eligible roles",
    description:
      'Use for "is Jane an admin?" or before offboarding an admin. `active` are roles in effect now; `eligible` ' +
      'are PIM roles the user can activate (Entra ID P2). Roles held through a role-assignable group show under ' +
      'eligible with memberType Group, but active group-based roles are listed under the group, not the user. ' +
      'A section shows `{ error }` when that part needs a permission or license the tenant lacks.',
    examples: ['tenantctl entra roles user jane@contoso.com'],
    args: [{ name: 'user', type: 'string', required: true, positional: true, description: 'User UPN, name or id.' }],
    risk: 'read',
    permissions: {
      delegated: ['RoleManagement.Read.Directory', 'RoleEligibilitySchedule.Read.Directory'],
      application: ['RoleManagement.Read.Directory', 'RoleEligibilitySchedule.Read.Directory'],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const filter = `principalId eq ${odata(user.id)}`;
      const sections = await settle({
        active: ctx.graph
          .list<RoleAssignment>('/roleManagement/directory/roleAssignments', {
            query: { $filter: filter, $expand: 'roleDefinition' },
            limit: ALL,
            pageSize: false,
          })
          .then((r) => r.items),
        eligible: ctx.graph
          .list<RoleEligibilityInstance>('/roleManagement/directory/roleEligibilityScheduleInstances', {
            query: { $filter: filter, $expand: 'roleDefinition' },
            limit: ALL,
            pageSize: false,
          })
          .then((r) => r.items),
      });
      if (ctx.full) return { user, ...sections };
      return {
        user: { id: user.id, displayName: user.displayName, userPrincipalName: user.userPrincipalName },
        active: Array.isArray(sections.active)
          ? sections.active.map((a) =>
              compact({
                role: a.roleDefinition?.displayName ?? a.roleDefinitionId,
                scope: scopeOf(a.directoryScopeId),
              }),
            )
          : sections.active,
        eligible: Array.isArray(sections.eligible)
          ? sections.eligible.map((e) =>
              compact({
                role: e.roleDefinition?.displayName ?? e.roleDefinitionId,
                scope: scopeOf(e.directoryScopeId),
                memberType: e.memberType ?? undefined,
                start: e.startDateTime ?? null,
                end: e.endDateTime ?? null,
              }),
            )
          : sections.eligible,
      };
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Enterprise apps
// ---------------------------------------------------------------------------------------------------------------

export interface ServicePrincipal {
  id: string;
  appId: string;
  displayName: string;
  publisherName?: string | null;
  verifiedPublisher?: { displayName?: string | null } | null;
  servicePrincipalType?: string | null;
  accountEnabled?: boolean | null;
  appOwnerOrganizationId?: string | null;
  signInAudience?: string | null;
  appRoleAssignmentRequired?: boolean | null;
  tags?: string[];
  homepage?: string | null;
  appRoles?: { id: string; value?: string | null; displayName?: string | null }[];
}

const SP_SELECT =
  'id,appId,displayName,publisherName,verifiedPublisher,servicePrincipalType,accountEnabled,' +
  'appOwnerOrganizationId,signInAudience,appRoleAssignmentRequired,tags,homepage';

// Tenants that own Microsoft first-party apps.
const MICROSOFT_TENANTS = ['f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47'];

export function shapeServicePrincipal(sp: ServicePrincipal) {
  return compact({
    id: sp.id,
    appId: sp.appId,
    displayName: sp.displayName,
    publisher: sp.verifiedPublisher?.displayName || sp.publisherName || null,
    publisherVerified: Boolean(sp.verifiedPublisher?.displayName),
    isMicrosoftApp: MICROSOFT_TENANTS.includes(sp.appOwnerOrganizationId ?? ''),
    type: sp.servicePrincipalType ?? null,
    accountEnabled: sp.accountEnabled ?? null,
    signInAudience: sp.signInAudience ?? null,
    assignmentRequired: sp.appRoleAssignmentRequired ?? undefined,
  });
}

/** Enterprise app (service principal) by object id, appId (client id) or display name. */
export async function resolveServicePrincipal(
  ctx: Context,
  ref: string,
  select = SP_SELECT,
): Promise<ServicePrincipal> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'An app is required');
  if (isGuid(value)) {
    const direct = await ctx.graph.tryGet<ServicePrincipal>(`/servicePrincipals/${seg(value)}`, {
      query: { $select: select },
    });
    if (direct) return direct;
    const byAppId = await ctx.graph.list<ServicePrincipal>('/servicePrincipals', {
      query: { $filter: `appId eq ${odata(value)}`, $select: select },
      limit: 1,
    });
    if (byAppId.items[0]) return byAppId.items[0];
    throw new CliError('NOT_FOUND', `No enterprise app with id or appId ${value}`, {
      hint: 'App registrations without a service principal are not enterprise apps; check `tenantctl entra apps list`.',
    });
  }
  const describe = (sp: ServicePrincipal) => ({ id: sp.id, appId: sp.appId, displayName: sp.displayName });
  const exact = await ctx.graph.list<ServicePrincipal>('/servicePrincipals', {
    query: { $filter: `displayName eq ${odata(value)}`, $select: select },
    limit: 10,
  });
  if (exact.items.length) return pickOne(exact.items, value, 'app', (sp) => [sp.displayName], describe);
  const term = value.replace(/"/g, '');
  const fuzzy = await ctx.graph.list<ServicePrincipal>('/servicePrincipals', {
    query: { $search: `"displayName:${term}"`, $select: select, $count: true },
    advanced: true,
    limit: 25,
  });
  return pickOne(fuzzy.items, value, 'app', (sp) => [sp.displayName], describe);
}

interface AppRoleAssignment {
  appRoleId: string;
  resourceId: string;
  resourceDisplayName?: string | null;
  createdDateTime?: string | null;
}

interface PermissionGrant {
  resourceId: string;
  scope?: string | null;
  consentType: string;
  principalId?: string | null;
}

interface GraphApplication {
  id: string;
  appId: string;
  displayName: string;
  passwordCredentials?: {
    keyId: string;
    displayName?: string | null;
    hint?: string | null;
    endDateTime?: string | null;
  }[];
  keyCredentials?: { keyId: string; displayName?: string | null; type?: string | null; endDateTime?: string | null }[];
}

const NIL_GUID = '00000000-0000-0000-0000-000000000000';

const appCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'apps list',
    summary: 'List or search enterprise apps (service principals)',
    description:
      'Use to find an app by name or survey third-party apps. `isMicrosoftApp` marks Microsoft first-party apps; ' +
      '`publisherVerified` means a verified publisher. `assignmentRequired` false means any user can sign in. ' +
      'Next: `tenantctl entra apps get <app>` for the permissions it holds.',
    examples: ['tenantctl entra apps list --search zoom', 'tenantctl entra apps list --all'],
    args: [{ name: 'search', type: 'string', description: 'Words in the app name.' }, ...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: ['Application.Read.All'], application: ['Application.Read.All'] },
    async run(ctx, args) {
      const term = args.search?.replace(/"/g, '');
      const r = await ctx.graph.list<ServicePrincipal>('/servicePrincipals', {
        query: term ? { $select: SP_SELECT, $search: `"displayName:${term}"`, $count: true } : { $select: SP_SELECT },
        advanced: Boolean(term),
        limit: listLimit(args),
      });
      return ctx.full ? r : { items: r.items.map(shapeServicePrincipal), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'apps get',
    summary: 'Show an enterprise app with the application and delegated permissions it was granted',
    description:
      'Use for consent reviews or "what can this app access?". `applicationPermissions` are app-only roles ' +
      '(e.g. Mail.Read on Microsoft Graph) - powerful because they need no user. `delegatedPermissions` are ' +
      'scopes consented for users: consent "admin" applies to everyone, "user" to one user. Accepts name, appId ' +
      'or object id.',
    examples: ['tenantctl entra apps get "Zoom"', 'tenantctl entra apps get 00000003-0000-0000-c000-000000000000'],
    args: [
      { name: 'app', type: 'string', required: true, positional: true, description: 'App name, appId or object id.' },
    ],
    risk: 'read',
    permissions: {
      delegated: ['Application.Read.All', 'Directory.Read.All'],
      application: ['Application.Read.All', 'Directory.Read.All'],
    },
    async run(ctx, args) {
      const sp = await resolveServicePrincipal(ctx, args.app);
      const resources: Record<string, Promise<ServicePrincipal | null>> = {};
      const resource = (id: string) =>
        (resources[id] ??= ctx.graph
          .get<ServicePrincipal>(`/servicePrincipals/${seg(id)}`, { query: { $select: 'id,displayName,appRoles' } })
          .catch(() => null));
      const sections = await settle({
        applicationPermissions: ctx.graph
          .list<AppRoleAssignment>(`/servicePrincipals/${seg(sp.id)}/appRoleAssignments`, {
            limit: ALL,
            pageSize: false,
          })
          .then((r) =>
            Promise.all(
              r.items.map(async (a) => {
                const res = await resource(a.resourceId);
                const role = res?.appRoles?.find((x) => x.id === a.appRoleId);
                return {
                  resource: a.resourceDisplayName ?? res?.displayName ?? a.resourceId,
                  permission: a.appRoleId === NIL_GUID ? 'default access' : (role?.value ?? a.appRoleId),
                  grantedDateTime: a.createdDateTime ?? null,
                };
              }),
            ),
          ),
        delegatedPermissions: ctx.graph
          .list<PermissionGrant>(`/servicePrincipals/${seg(sp.id)}/oauth2PermissionGrants`, {
            limit: ALL,
            pageSize: false,
          })
          .then((r) =>
            Promise.all(
              r.items.map(async (g) =>
                compact({
                  resource: (await resource(g.resourceId))?.displayName ?? g.resourceId,
                  scopes: (g.scope ?? '').split(' ').filter(Boolean),
                  consent: g.consentType === 'AllPrincipals' ? 'admin' : 'user',
                  principalId: g.consentType === 'AllPrincipals' ? undefined : (g.principalId ?? undefined),
                }),
              ),
            ),
          ),
      });
      return { ...(ctx.full ? sp : shapeServicePrincipal(sp)), ...sections };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'apps expiring',
    summary: 'App registrations whose client secrets or certificates expire soon or already expired',
    description:
      'Use to prevent outages from expired secrets ("integration stopped working"). Checks every app ' +
      'registration; `daysLeft` < 0 means already expired. Rotate in the Entra admin center (App registrations > ' +
      'Certificates & secrets), then update the consuming service. SAML signing certificates live on enterprise ' +
      'apps and are not covered.',
    examples: ['tenantctl entra apps expiring', 'tenantctl entra apps expiring --days 90'],
    args: [
      { name: 'days', type: 'number', default: 30, description: 'Include credentials expiring within N days.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['Application.Read.All'], application: ['Application.Read.All'] },
    async run(ctx, args) {
      const apps = await ctx.graph.list<GraphApplication>('/applications', {
        query: { $select: 'id,appId,displayName,passwordCredentials,keyCredentials' },
        limit: ALL,
      });
      const now = Date.now();
      const found = [];
      for (const app of apps.items) {
        const creds = [
          ...(app.passwordCredentials ?? []).map((c) => ({ ...c, credentialType: 'secret', detail: c.hint })),
          ...(app.keyCredentials ?? []).map((c) => ({ ...c, credentialType: 'certificate', detail: c.type })),
        ];
        for (const c of creds) {
          if (!c.endDateTime) continue;
          const daysLeft = Math.floor((Date.parse(c.endDateTime) - now) / DAY_MS);
          if (daysLeft > args.days) continue;
          found.push(
            compact({
              app: app.displayName,
              appId: app.appId,
              credentialType: c.credentialType,
              name: c.displayName ?? null,
              hint: c.detail ?? undefined,
              keyId: c.keyId,
              endDateTime: c.endDateTime,
              daysLeft,
              expired: daysLeft < 0,
            }),
          );
        }
      }
      found.sort((a, b) => (a.daysLeft ?? 0) - (b.daysLeft ?? 0));
      const limit = listLimit(args);
      return { items: found.slice(0, limit), truncated: found.length > limit } satisfies ListResult;
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Conditional Access
// ---------------------------------------------------------------------------------------------------------------

interface CaPolicy {
  id: string;
  displayName: string;
  state: string;
  createdDateTime?: string | null;
  modifiedDateTime?: string | null;
  conditions?: {
    users?: Partial<
      Record<
        'includeUsers' | 'excludeUsers' | 'includeGroups' | 'excludeGroups' | 'includeRoles' | 'excludeRoles',
        string[]
      >
    > & {
      includeGuestsOrExternalUsers?: { guestOrExternalUserTypes?: string } | null;
      excludeGuestsOrExternalUsers?: { guestOrExternalUserTypes?: string } | null;
    };
    applications?: {
      includeApplications?: string[];
      excludeApplications?: string[];
      includeUserActions?: string[];
    } | null;
    platforms?: { includePlatforms?: string[]; excludePlatforms?: string[] } | null;
    locations?: { includeLocations?: string[]; excludeLocations?: string[] } | null;
    clientAppTypes?: string[];
    signInRiskLevels?: string[];
    userRiskLevels?: string[];
    devices?: { deviceFilter?: { mode?: string; rule?: string } | null } | null;
  };
  grantControls?: {
    operator?: string;
    builtInControls?: string[];
    customAuthenticationFactors?: string[];
    termsOfUse?: string[];
    authenticationStrength?: { displayName?: string } | null;
  } | null;
  sessionControls?: {
    signInFrequency?: { isEnabled?: boolean; value?: number; type?: string; frequencyInterval?: string } | null;
    persistentBrowser?: { isEnabled?: boolean; mode?: string } | null;
    cloudAppSecurity?: { isEnabled?: boolean; cloudAppSecurityType?: string } | null;
    applicationEnforcedRestrictions?: { isEnabled?: boolean } | null;
    disableResilienceDefaults?: boolean | null;
  } | null;
}

const CA_STATE: Record<string, string> = {
  enabled: 'enabled',
  disabled: 'disabled',
  enabledForReportingButNotEnforced: 'report-only',
};

/** Keywords (All, None, GuestsOrExternalUsers, Office365...) are kept; specific ids are counted. */
function scopeSummary(ids: string[] | undefined): string | undefined {
  if (!ids?.length) return undefined;
  const keywords = ids.filter((id) => !isGuid(id));
  const count = ids.length - keywords.length;
  return [...keywords, ...(count ? [`${count} specific`] : [])].join(' + ');
}

function nonEmpty(values: string[] | undefined): string[] | undefined {
  return values?.length ? values : undefined;
}

function grantSummary(p: CaPolicy): string | null {
  const g = p.grantControls;
  if (!g) return null;
  const parts = [...(g.builtInControls ?? [])];
  if (g.authenticationStrength?.displayName) parts.push(`authStrength:${g.authenticationStrength.displayName}`);
  if (g.termsOfUse?.length) parts.push('termsOfUse');
  if (g.customAuthenticationFactors?.length) parts.push('customControl');
  return parts.length ? parts.join(` ${g.operator ?? 'OR'} `) : null;
}

function sessionSummary(p: CaPolicy): string[] | undefined {
  const s = p.sessionControls;
  if (!s) return undefined;
  const parts: string[] = [];
  if (s.signInFrequency?.isEnabled) {
    parts.push(
      s.signInFrequency.frequencyInterval === 'everyTime'
        ? 'signInFrequency: every time'
        : `signInFrequency: ${s.signInFrequency.value} ${s.signInFrequency.type}`,
    );
  }
  if (s.persistentBrowser?.isEnabled) parts.push(`persistentBrowser: ${s.persistentBrowser.mode}`);
  if (s.cloudAppSecurity?.isEnabled) parts.push(`appControl: ${s.cloudAppSecurity.cloudAppSecurityType}`);
  if (s.applicationEnforcedRestrictions?.isEnabled) parts.push('appEnforcedRestrictions');
  if (s.disableResilienceDefaults) parts.push('resilienceDefaultsDisabled');
  return parts.length ? parts : undefined;
}

export function shapeCaPolicy(p: CaPolicy) {
  const c = p.conditions ?? {};
  const u = c.users ?? {};
  return compact({
    id: p.id,
    name: p.displayName,
    state: CA_STATE[p.state] ?? p.state,
    modifiedDateTime: p.modifiedDateTime ?? p.createdDateTime ?? null,
    users: compact({
      include: scopeSummary(u.includeUsers),
      exclude: scopeSummary(u.excludeUsers),
      includeGroups: scopeSummary(u.includeGroups),
      excludeGroups: scopeSummary(u.excludeGroups),
      includeRoles: scopeSummary(u.includeRoles),
      excludeRoles: scopeSummary(u.excludeRoles),
      includeGuests: u.includeGuestsOrExternalUsers?.guestOrExternalUserTypes,
      excludeGuests: u.excludeGuestsOrExternalUsers?.guestOrExternalUserTypes,
    }),
    apps: compact({
      include: scopeSummary(c.applications?.includeApplications),
      exclude: scopeSummary(c.applications?.excludeApplications),
      userActions: nonEmpty(c.applications?.includeUserActions),
    }),
    platforms: c.platforms
      ? compact({ include: nonEmpty(c.platforms.includePlatforms), exclude: nonEmpty(c.platforms.excludePlatforms) })
      : undefined,
    locations: c.locations
      ? compact({
          include: scopeSummary(c.locations.includeLocations),
          exclude: scopeSummary(c.locations.excludeLocations),
        })
      : undefined,
    clientAppTypes: nonEmpty(c.clientAppTypes),
    signInRiskLevels: nonEmpty(c.signInRiskLevels),
    userRiskLevels: nonEmpty(c.userRiskLevels),
    deviceFilter: c.devices?.deviceFilter?.rule
      ? `${c.devices.deviceFilter.mode}: ${c.devices.deviceFilter.rule}`
      : undefined,
    grant: grantSummary(p),
    session: sessionSummary(p),
  });
}

async function resolveCaPolicy(ctx: Context, ref: string): Promise<CaPolicy> {
  const value = ref.trim();
  if (isGuid(value)) {
    const policy = await ctx.graph.tryGet<CaPolicy>(`/identity/conditionalAccess/policies/${seg(value)}`);
    if (policy) return policy;
  }
  const { items } = await ctx.graph.list<CaPolicy>('/identity/conditionalAccess/policies', { limit: ALL });
  const wanted = value.toLowerCase();
  return pickOne(
    items.filter((p) => p.displayName.toLowerCase().includes(wanted)),
    value,
    'Conditional Access policy',
    (p) => [p.displayName],
    (p) => ({ id: p.id, name: p.displayName, state: CA_STATE[p.state] ?? p.state }),
  );
}

function named(ids: string[] | undefined, names: Record<string, string>): string[] | undefined {
  return ids?.length ? ids.map((id) => names[id] ?? id) : undefined;
}

const caCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'ca list',
    summary: 'List Conditional Access policies with a one-line summary of who, what and how',
    description:
      'Use to understand why a sign-in was blocked or MFA was required. `state` report-only policies log but ' +
      'never enforce. Scope fields read like "All", "3 specific" or "GuestsOrExternalUsers + 1 specific"; ' +
      '`grant` is the requirement (e.g. "mfa OR compliantDevice", "block"). Needs Entra ID P1. Next: ' +
      '`tenantctl entra ca get <policy>` to see names of included/excluded users, groups, roles and apps.',
    examples: ['tenantctl entra ca list'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: ['Policy.Read.All'], application: ['Policy.Read.All'], roles: ['Security Reader'] },
    async run(ctx, args) {
      const r = await ctx.graph.list<CaPolicy>('/identity/conditionalAccess/policies', { limit: listLimit(args) });
      return ctx.full ? r : { items: r.items.map(shapeCaPolicy), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'ca get',
    summary: 'Show one Conditional Access policy with users, groups, roles, apps and locations named',
    description:
      'Accepts the policy id or part of its name. `resolved` translates ids in the conditions to names ' +
      '(unresolvable ids stay as ids, e.g. deleted objects). Use with `tenantctl entra signins --user X` to match ' +
      'a failed sign-in (error 53003) to the policy that blocked it.',
    examples: ['tenantctl entra ca get "Require MFA for admins"'],
    args: [{ name: 'policy', type: 'string', required: true, positional: true, description: 'Policy id or name.' }],
    risk: 'read',
    permissions: {
      delegated: ['Policy.Read.All', 'Directory.Read.All'],
      application: ['Policy.Read.All', 'Directory.Read.All'],
      roles: ['Security Reader'],
    },
    async run(ctx, args) {
      const policy = await resolveCaPolicy(ctx, args.policy);
      if (ctx.full) return policy;
      const c = policy.conditions ?? {};
      const u = c.users ?? {};
      const guids = (lists: (string[] | undefined)[]) => [
        ...new Set(lists.flat().filter((id): id is string => !!id && isGuid(id))),
      ];
      const objectIds = guids([u.includeUsers, u.excludeUsers, u.includeGroups, u.excludeGroups]);
      const appIds = guids([c.applications?.includeApplications, c.applications?.excludeApplications]);
      const locationIds = guids([c.locations?.includeLocations, c.locations?.excludeLocations]);
      const lookups = await settle({
        objects: objectIds.length
          ? ctx.graph
              .post<{ value: DirectoryObject[] }>('/directoryObjects/getByIds', {
                ids: objectIds,
                types: ['user', 'group'],
              })
              .then((r) => Object.fromEntries(r.value.map((o) => [o.id, o.userPrincipalName ?? o.displayName ?? o.id])))
          : Promise.resolve({} as Record<string, string>),
        roles:
          u.includeRoles?.length || u.excludeRoles?.length
            ? roleNames(ctx)
            : Promise.resolve({} as Record<string, string>),
        apps: appIds.length
          ? ctx.graph
              .list<ServicePrincipal>('/servicePrincipals', {
                query: { $filter: `appId in (${appIds.map(odata).join(',')})`, $select: 'appId,displayName' },
                limit: ALL,
              })
              .then((r) => Object.fromEntries(r.items.map((sp) => [sp.appId, sp.displayName])))
          : Promise.resolve({} as Record<string, string>),
        locations: locationIds.length
          ? ctx.graph
              .list<{ id: string; displayName: string }>('/identity/conditionalAccess/namedLocations', { limit: ALL })
              .then((r) => Object.fromEntries(r.items.map((l) => [l.id, l.displayName])))
          : Promise.resolve({} as Record<string, string>),
      });
      const names: Record<string, string> = {};
      const errors: Record<string, unknown> = {};
      for (const [key, section] of Object.entries(lookups) as [string, Record<string, string> | { error: unknown }][]) {
        if ('error' in section) errors[key] = section.error;
        else Object.assign(names, section);
      }
      return compact({
        ...shapeCaPolicy(policy),
        resolved: compact({
          includeUsers: named(u.includeUsers, names),
          excludeUsers: named(u.excludeUsers, names),
          includeGroups: named(u.includeGroups, names),
          excludeGroups: named(u.excludeGroups, names),
          includeRoles: named(u.includeRoles, names),
          excludeRoles: named(u.excludeRoles, names),
          includeApps: named(c.applications?.includeApplications, names),
          excludeApps: named(c.applications?.excludeApplications, names),
          includeLocations: named(c.locations?.includeLocations, names),
          excludeLocations: named(c.locations?.excludeLocations, names),
        }),
        lookupErrors: Object.keys(errors).length ? errors : undefined,
      });
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Identity Protection
// ---------------------------------------------------------------------------------------------------------------

interface RiskyUser {
  id: string;
  userPrincipalName?: string | null;
  userDisplayName?: string | null;
  riskLevel?: string | null;
  riskState?: string | null;
  riskDetail?: string | null;
  riskLastUpdatedDateTime?: string | null;
  isProcessing?: boolean | null;
}

function shapeRiskyUser(r: RiskyUser) {
  return {
    id: r.id,
    userPrincipalName: r.userPrincipalName ?? null,
    displayName: r.userDisplayName ?? null,
    riskLevel: r.riskLevel ?? null,
    riskState: r.riskState ?? null,
    riskDetail: r.riskDetail ?? null,
    lastUpdated: r.riskLastUpdatedDateTime ?? null,
  };
}

interface RiskDetection {
  id: string;
  detectedDateTime?: string | null;
  userPrincipalName?: string | null;
  riskEventType?: string | null;
  riskLevel?: string | null;
  riskState?: string | null;
  riskDetail?: string | null;
  ipAddress?: string | null;
  location?: { city?: string | null; state?: string | null; countryOrRegion?: string | null } | null;
  activity?: string | null;
  detectionTimingType?: string | null;
  source?: string | null;
  correlationId?: string | null;
}

function shapeRiskDetection(d: RiskDetection) {
  const place = [d.location?.city, d.location?.state, d.location?.countryOrRegion].filter(Boolean).join(', ');
  return compact({
    time: d.detectedDateTime ?? null,
    user: d.userPrincipalName ?? null,
    riskEventType: d.riskEventType ?? null,
    riskLevel: d.riskLevel ?? null,
    riskState: d.riskState ?? null,
    riskDetail: d.riskDetail ?? undefined,
    ipAddress: d.ipAddress ?? null,
    location: place || null,
    activity: d.activity ?? undefined,
    timing: d.detectionTimingType ?? undefined,
    source: d.source ?? undefined,
    correlationId: d.correlationId ?? undefined,
  });
}

const userArg = {
  name: 'user',
  type: 'string',
  required: true,
  positional: true,
  description: 'User UPN, name or id.',
} as const;

async function riskyUserPreview(ctx: Context, ref: string) {
  const user = await resolveUser(ctx, ref);
  const risky = await ctx.graph.tryGet<RiskyUser>(`/identityProtection/riskyUsers/${seg(user.id)}`);
  return { user, risk: risky ? shapeRiskyUser(risky) : null };
}

const riskCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'risky-users',
    summary: 'Users Entra ID Protection flags as risky (leaked credentials, impossible travel...)',
    description:
      'Requires Entra ID P2. Use --state atRisk for users that still need action. After investigating, either ' +
      '`risky-users dismiss` (false positive) or `risky-users confirm-compromised`, and reset the password / ' +
      'revoke sessions with the entra users commands. Next: `tenantctl entra risk-detections --user <user>` for the why.',
    examples: ['tenantctl entra risky-users --state atRisk', 'tenantctl entra risky-users --level high'],
    args: [
      { name: 'level', type: 'string', choices: ['low', 'medium', 'high'], description: 'Only this risk level.' },
      {
        name: 'state',
        type: 'string',
        choices: ['atRisk', 'confirmedCompromised', 'remediated', 'dismissed'],
        description: 'Only this risk state.',
      },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['IdentityRiskyUser.Read.All'],
      application: ['IdentityRiskyUser.Read.All'],
      roles: ['Security Reader'],
    },
    async run(ctx, args) {
      const filters = [
        ...(args.level ? [`riskLevel eq ${odata(args.level)}`] : []),
        ...(args.state ? [`riskState eq ${odata(args.state)}`] : []),
      ];
      const r = await ctx.graph.list<RiskyUser>('/identityProtection/riskyUsers', {
        query: { $filter: filters.length ? filters.join(' and ') : undefined },
        limit: listLimit(args),
      });
      return ctx.full ? r : { items: r.items.map(shapeRiskyUser), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'risky-users dismiss',
    summary: "Dismiss a user's risk after confirming it was a false positive",
    description:
      'Sets the user risk to none (riskState dismissed) so risk-based Conditional Access stops challenging them. ' +
      'Only do this after reviewing `tenantctl entra risk-detections --user <user>`. Requires Entra ID P2.',
    examples: ['tenantctl entra risky-users dismiss jane@contoso.com'],
    args: [userArg],
    risk: 'write',
    permissions: {
      delegated: ['IdentityRiskyUser.ReadWrite.All'],
      application: ['IdentityRiskyUser.ReadWrite.All'],
      roles: ['Security Operator'],
    },
    async preview(ctx, args) {
      const current = await riskyUserPreview(ctx, args.user);
      return { ...current, action: `Dismiss user risk for ${current.user.userPrincipalName}.` };
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      await ctx.graph.post('/identityProtection/riskyUsers/dismiss', { userIds: [user.id] });
      return { user: user.userPrincipalName, id: user.id, status: 'dismissed' };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'risky-users confirm-compromised',
    summary: 'Mark a user as confirmed compromised (raises user risk to high)',
    description:
      "Use when an investigation shows the account is in an attacker's hands. Sets user risk to high, so " +
      'user-risk Conditional Access policies block or force a secure password change. Also reset the password ' +
      'and revoke sessions. Requires Entra ID P2.',
    examples: ['tenantctl entra risky-users confirm-compromised jane@contoso.com'],
    args: [userArg],
    risk: 'destructive',
    permissions: {
      delegated: ['IdentityRiskyUser.ReadWrite.All'],
      application: ['IdentityRiskyUser.ReadWrite.All'],
      roles: ['Security Operator'],
    },
    async preview(ctx, args) {
      const current = await riskyUserPreview(ctx, args.user);
      return {
        ...current,
        action:
          `Mark ${current.user.userPrincipalName} as confirmed compromised: user risk becomes high and risk-based ` +
          'Conditional Access policies will block or force a password change.',
      };
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      await ctx.graph.post('/identityProtection/riskyUsers/confirmCompromised', { userIds: [user.id] });
      return { user: user.userPrincipalName, id: user.id, status: 'confirmedCompromised' };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'risk-detections',
    summary: 'Identity Protection risk detections: what made a user or sign-in risky',
    description:
      'Each item is one detection (riskEventType such as unfamiliarFeatures, anonymizedIPAddress, ' +
      'leakedCredentials, impossibleTravel) with IP and location. Full detail requires Entra ID P2; P1 tenants ' +
      'see limited detections. Use `correlationId` to find the matching sign-in.',
    examples: [
      'tenantctl entra risk-detections --user jane@contoso.com --days 30',
      'tenantctl entra risk-detections --level high',
    ],
    args: [
      { name: 'user', type: 'string', description: 'Only this user (UPN, name or id).' },
      { name: 'days', type: 'number', default: 7, description: 'Look back this many days.' },
      { name: 'level', type: 'string', choices: ['low', 'medium', 'high'], description: 'Only this risk level.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['IdentityRiskEvent.Read.All'],
      application: ['IdentityRiskEvent.Read.All'],
      roles: ['Security Reader'],
    },
    async run(ctx, args) {
      const filters = [`detectedDateTime ge ${daysAgo(args.days)}`];
      if (args.user) filters.push(`userId eq ${odata((await resolveUser(ctx, args.user)).id)}`);
      if (args.level) filters.push(`riskLevel eq ${odata(args.level)}`);
      const r = await ctx.graph.list<RiskDetection>('/identityProtection/riskDetections', {
        query: { $filter: filters.join(' and ') },
        limit: listLimit(args),
      });
      return ctx.full ? r : { items: r.items.map(shapeRiskDetection), truncated: r.truncated };
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Entra devices
// ---------------------------------------------------------------------------------------------------------------

export interface EntraDevice {
  id: string;
  deviceId?: string | null;
  displayName: string;
  operatingSystem?: string | null;
  operatingSystemVersion?: string | null;
  accountEnabled?: boolean | null;
  trustType?: string | null;
  isCompliant?: boolean | null;
  isManaged?: boolean | null;
  approximateLastSignInDateTime?: string | null;
  registrationDateTime?: string | null;
  profileType?: string | null;
  managementType?: string | null;
  deviceOwnership?: string | null;
}

const DEVICE_SELECT =
  'id,deviceId,displayName,operatingSystem,operatingSystemVersion,accountEnabled,trustType,isCompliant,isManaged,' +
  'approximateLastSignInDateTime,registrationDateTime,profileType,managementType,deviceOwnership';

const TRUST_TYPES: Record<string, string> = {
  AzureAd: 'Entra joined',
  ServerAd: 'Hybrid joined',
  Workplace: 'Entra registered',
};

export function shapeEntraDevice(d: EntraDevice) {
  return compact({
    id: d.id,
    deviceId: d.deviceId ?? null,
    displayName: d.displayName,
    operatingSystem: d.operatingSystem ?? null,
    osVersion: d.operatingSystemVersion ?? null,
    join: d.trustType ? (TRUST_TYPES[d.trustType] ?? d.trustType) : null,
    accountEnabled: d.accountEnabled ?? null,
    isCompliant: d.isCompliant ?? null,
    isManaged: d.isManaged ?? null,
    ownership: d.deviceOwnership ?? undefined,
    lastSignIn: d.approximateLastSignInDateTime ?? null,
    daysSinceSignIn: daysSince(d.approximateLastSignInDateTime),
    registered: d.registrationDateTime ?? null,
  });
}

function describeDevice(d: EntraDevice) {
  return {
    id: d.id,
    deviceId: d.deviceId,
    displayName: d.displayName,
    operatingSystem: d.operatingSystem,
    join: d.trustType ? (TRUST_TYPES[d.trustType] ?? d.trustType) : null,
    lastSignIn: d.approximateLastSignInDateTime,
  };
}

/**
 * Entra device by object id, deviceId (the id Intune and dsregcmd show) or display name. Re-registered devices
 * often share a name, which surfaces as AMBIGUOUS with each candidate's last sign-in.
 */
export async function resolveEntraDevice(ctx: Context, ref: string, select = DEVICE_SELECT): Promise<EntraDevice> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A device is required');
  if (isGuid(value)) {
    const direct = await ctx.graph.tryGet<EntraDevice>(`/devices/${seg(value)}`, { query: { $select: select } });
    if (direct) return direct;
    const byDeviceId = await ctx.graph.list<EntraDevice>('/devices', {
      query: { $filter: `deviceId eq ${odata(value)}`, $select: select },
      limit: 1,
    });
    if (byDeviceId.items[0]) return byDeviceId.items[0];
    throw new CliError('NOT_FOUND', `No Entra device with id or deviceId ${value}`, {
      hint: 'Intune-only ids (managedDevice id) differ; use `tenantctl intune devices get` for those.',
    });
  }
  const exact = await ctx.graph.list<EntraDevice>('/devices', {
    query: { $filter: `displayName eq ${odata(value)}`, $select: select },
    limit: 10,
  });
  if (exact.items.length) return pickOne(exact.items, value, 'device', (d) => [d.displayName], describeDevice);
  const term = value.replace(/"/g, '');
  const fuzzy = await ctx.graph.list<EntraDevice>('/devices', {
    query: { $search: `"displayName:${term}"`, $select: select, $count: true },
    advanced: true,
    limit: 25,
  });
  return pickOne(fuzzy.items, value, 'device', (d) => [d.displayName], describeDevice);
}

const deviceArg = {
  name: 'device',
  type: 'string',
  required: true,
  positional: true,
  description: 'Device name, deviceId or object id.',
} as const;

const DEVICE_WRITE_PERMISSIONS = {
  delegated: ['Directory.AccessAsUser.All'],
  application: ['Device.ReadWrite.All'],
  roles: ['Cloud Device Administrator'],
};

const deviceCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'devices list',
    summary: 'List Entra ID device objects; find stale or disabled devices',
    description:
      'Entra device objects (identity), not Intune management records. `join` is Entra joined, Hybrid joined ' +
      'or Entra registered (BYOD). `daysSinceSignIn` is approximate (updated about every 14 days). Use ' +
      '--stale-days 90 to find cleanup candidates. For compliance, apps and remote actions use `tenantctl intune devices`.',
    examples: [
      'tenantctl entra devices list --stale-days 90 --all',
      'tenantctl entra devices list --os Windows --disabled',
    ],
    args: [
      { name: 'search', type: 'string', description: 'Words in the device name.' },
      { name: 'os', type: 'string', description: 'Operating system prefix, e.g. Windows, iOS, MacMDM, Android.' },
      { name: 'staleDays', type: 'number', description: 'Only devices with no sign-in for at least N days.' },
      { name: 'disabled', type: 'boolean', description: 'Only disabled devices.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['Device.Read.All'], application: ['Device.Read.All'] },
    async run(ctx, args) {
      const filters: string[] = [];
      if (args.os) filters.push(`startswith(operatingSystem, ${odata(args.os)})`);
      if (args.staleDays !== undefined) filters.push(`approximateLastSignInDateTime le ${daysAgo(args.staleDays)}`);
      if (args.disabled) filters.push('accountEnabled eq false');
      const term = args.search?.replace(/"/g, '');
      const r = await ctx.graph.list<EntraDevice>('/devices', {
        query: {
          $select: DEVICE_SELECT,
          $filter: filters.length ? filters.join(' and ') : undefined,
          $search: term ? `"displayName:${term}"` : undefined,
          $count: true,
        },
        advanced: true,
        limit: listLimit(args),
      });
      return ctx.full ? r : { items: r.items.map(shapeEntraDevice), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'devices get',
    summary: 'Show an Entra device with its registered owners and users',
    description:
      'Use to see who a device belongs to in Entra and whether it is enabled, compliant and recently active. ' +
      'Accepts name, deviceId (as shown by Intune / dsregcmd) or object id.',
    examples: ['tenantctl entra devices get DESKTOP-123'],
    args: [deviceArg],
    risk: 'read',
    permissions: { delegated: ['Device.Read.All'], application: ['Device.Read.All'] },
    async run(ctx, args) {
      const device = await resolveEntraDevice(ctx, args.device);
      const people = (rel: string) =>
        ctx.graph
          .list<DirectoryObject>(`/devices/${seg(device.id)}/${rel}`, { limit: ALL, pageSize: false })
          .then((r) =>
            r.items.map((o) =>
              compact({ displayName: o.displayName ?? null, userPrincipalName: o.userPrincipalName ?? undefined }),
            ),
          );
      const related = await settle({ owners: people('registeredOwners'), users: people('registeredUsers') });
      return { ...(ctx.full ? device : shapeEntraDevice(device)), ...related };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'devices disable',
    summary: 'Disable an Entra device so it can no longer authenticate (lost or stolen device)',
    description:
      'Blocks the device from Entra sign-in, device-based Conditional Access and SSO; users on it lose access to ' +
      'cloud resources. Does not wipe or retire it in Intune (use `tenantctl intune devices wipe|retire`). Reverse with ' +
      '`tenantctl entra devices enable`. With application permissions only Windows devices can be changed.',
    examples: ['tenantctl entra devices disable DESKTOP-123'],
    args: [deviceArg],
    risk: 'destructive',
    permissions: DEVICE_WRITE_PERMISSIONS,
    async preview(ctx, args) {
      const device = await resolveEntraDevice(ctx, args.device);
      return {
        device: shapeEntraDevice(device),
        action: `Disable Entra device "${device.displayName}": it can no longer authenticate to Entra ID.`,
      };
    },
    async run(ctx, args) {
      const device = await resolveEntraDevice(ctx, args.device);
      await ctx.graph.patch(`/devices/${seg(device.id)}`, { accountEnabled: false });
      return { id: device.id, displayName: device.displayName, accountEnabled: false };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'devices enable',
    summary: 'Re-enable a disabled Entra device',
    description: 'Restores Entra sign-in for a device disabled with `devices disable`.',
    examples: ['tenantctl entra devices enable DESKTOP-123'],
    args: [deviceArg],
    risk: 'write',
    permissions: DEVICE_WRITE_PERMISSIONS,
    async preview(ctx, args) {
      const device = await resolveEntraDevice(ctx, args.device);
      return { device: shapeEntraDevice(device), action: `Enable Entra device "${device.displayName}".` };
    },
    async run(ctx, args) {
      const device = await resolveEntraDevice(ctx, args.device);
      await ctx.graph.patch(`/devices/${seg(device.id)}`, { accountEnabled: true });
      return { id: device.id, displayName: device.displayName, accountEnabled: true };
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Tenant
// ---------------------------------------------------------------------------------------------------------------

interface Organization {
  id: string;
  displayName: string;
  tenantType?: string | null;
  countryLetterCode?: string | null;
  verifiedDomains?: { name: string; isDefault?: boolean; isInitial?: boolean }[];
  technicalNotificationMails?: string[];
  onPremisesSyncEnabled?: boolean | null;
  onPremisesLastSyncDateTime?: string | null;
}

interface Domain {
  id: string;
  isVerified?: boolean;
  isDefault?: boolean;
  isInitial?: boolean;
  authenticationType?: string | null;
  supportedServices?: string[];
}

const DIRSYNC_STALE_HOURS = 3;

const tenantCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'tenant',
    summary: 'Tenant overview: name, id, domains (managed/federated) and directory sync health',
    description:
      'Start here to orient in an unfamiliar tenant. `dirSyncStale` is true when on-premises sync is enabled ' +
      `and the last sync is more than ${DIRSYNC_STALE_HOURS} hours old (default cycle is 30 minutes) - check the ` +
      'Entra Connect / Cloud Sync server. Federated domains authenticate at an external IdP (e.g. AD FS), which ' +
      'matters when troubleshooting sign-in.',
    examples: ['tenantctl entra tenant'],
    args: [],
    risk: 'read',
    permissions: {
      delegated: ['User.Read', 'Domain.Read.All'],
      application: ['Organization.Read.All', 'Domain.Read.All'],
    },
    async run(ctx) {
      const [orgs, extra] = await Promise.all([
        ctx.graph.list<Organization>('/organization', { limit: 1, pageSize: false }),
        settle({
          domains: ctx.graph.list<Domain>('/domains', { limit: ALL, pageSize: false }).then((r) => r.items),
        }),
      ]);
      const org = orgs.items[0];
      if (!org) throw new CliError('NOT_FOUND', 'Graph returned no organization for this tenant');
      if (ctx.full) return { ...org, domains: extra.domains };
      const lastSync = org.onPremisesLastSyncDateTime ? Date.parse(org.onPremisesLastSyncDateTime) : Number.NaN;
      const hours = Number.isNaN(lastSync) ? null : Math.round(((Date.now() - lastSync) / 3_600_000) * 10) / 10;
      const syncOn = org.onPremisesSyncEnabled === true;
      return compact({
        id: org.id,
        displayName: org.displayName,
        tenantType: org.tenantType ?? undefined,
        country: org.countryLetterCode ?? undefined,
        defaultDomain: org.verifiedDomains?.find((d) => d.isDefault)?.name ?? null,
        initialDomain: org.verifiedDomains?.find((d) => d.isInitial)?.name ?? null,
        technicalNotificationMails: org.technicalNotificationMails ?? [],
        onPremisesSyncEnabled: syncOn,
        onPremisesLastSyncDateTime: org.onPremisesLastSyncDateTime ?? null,
        hoursSinceLastDirSync: syncOn ? hours : undefined,
        dirSyncStale: syncOn ? hours === null || hours > DIRSYNC_STALE_HOURS : undefined,
        domains: Array.isArray(extra.domains)
          ? extra.domains.map((d) => ({
              name: d.id,
              isVerified: d.isVerified ?? false,
              isDefault: d.isDefault ?? false,
              authentication: d.authenticationType?.toLowerCase() ?? null,
              supportedServices: d.supportedServices ?? [],
            }))
          : extra.domains,
      });
    },
  }),
];

export const commands: AnyCommand[] = [
  ...tenantCommands,
  ...groupCommands,
  ...directoryLogCommands,
  ...roleCommands,
  ...appCommands,
  ...caCommands,
  ...riskCommands,
  ...deviceCommands,
];
