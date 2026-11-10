import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { settle } from '../../core/util.ts';
import { commands as callCommands } from './calls.ts';
import { IS_TEAM_FILTER, resolveChannel, resolveTeam, type ResolvedTeam, TEAM_GROUP_SELECT } from './resolve.ts';

export { resolveChannel, resolveTeam, type ResolvedChannel, type ResolvedTeam } from './resolve.ts';
export { summarizeCallQuality } from './calls.ts';

const TEAM_ARG = {
  name: 'team',
  type: 'string',
  required: true,
  positional: true,
  description: 'Team display name, group id, group mail address or a Teams link containing groupId=.',
} as const;
const USER_ARG = {
  name: 'user',
  type: 'string',
  required: true,
  positional: true,
  description: 'User UPN, email, display name or id.',
} as const;

const TEAMS_ADMIN = ['Teams Administrator'];
// resolveUser filters on mail/proxyAddresses, which User.ReadBasic.All does not cover.
const USER_READ = 'User.Read.All';

export interface ConversationMember {
  id: string;
  displayName?: string | null;
  email?: string | null;
  userId?: string | null;
  roles?: string[];
}

export function shapeTeamGroup(group: ResolvedTeam) {
  return {
    id: group.id,
    displayName: group.displayName,
    description: group.description,
    visibility: group.visibility,
    mail: group.mail,
    createdDateTime: group.createdDateTime,
  };
}

/** Teams reports owners and guests in `roles`; ordinary members have an empty list. */
export function memberRole(member: ConversationMember): 'owner' | 'guest' | 'member' {
  const roles = member.roles ?? [];
  if (roles.includes('owner')) return 'owner';
  if (roles.includes('guest')) return 'guest';
  return 'member';
}

export function shapeMember(member: ConversationMember) {
  return {
    displayName: member.displayName ?? null,
    email: member.email ?? null,
    role: memberRole(member),
    userId: member.userId ?? null,
    membershipId: member.id,
  };
}

async function findMembership(ctx: Context, teamId: string, userId: string): Promise<ConversationMember | undefined> {
  const { items } = await ctx.graph.list<ConversationMember>(`/teams/${seg(teamId)}/members`, {
    query: { $filter: `(microsoft.graph.aadUserConversationMember/userId eq ${odata(userId)})` },
    limit: 1,
  });
  return items[0];
}

export interface TeamsOperation {
  teamId: string | null;
  operationId: string | null;
  /** Command that polls the teamsAsyncOperation (status, targetResourceId, error), or null without a Location. */
  check: string | null;
}

/**
 * Create/archive/unarchive answer 202 with `Location: /teams('{id}')/operations('{opId}')` and
 * `Content-Location: /teams('{id}')`; the ids exist nowhere else. Graph says to poll Location, >30 s apart.
 */
export function teamsOperation(headers: Headers): TeamsOperation {
  const location = headers.get('location') ?? '';
  const teamId = (location || headers.get('content-location') || '').match(/teams\('([^']+)'\)/)?.[1] ?? null;
  const operationId = location.match(/operations\('([^']+)'\)/)?.[1] ?? null;
  return {
    teamId,
    operationId,
    check: teamId && operationId ? `m365 graph get "/teams('${teamId}')/operations('${operationId}')"` : null,
  };
}

async function teamAndUser(ctx: Context, team: string, user: string) {
  const [resolvedTeam, resolvedUser] = await Promise.all([resolveTeam(ctx, team), resolveUser(ctx, user)]);
  return { team: resolvedTeam, user: resolvedUser };
}

const ARCHIVE_CHECK_BATCH = 20;

/**
 * Graph's team list never fills in isArchived, so archived teams are found by reading each team. Checks run in
 * small batches and stop once enough archived teams are found.
 */
async function archivedTeams(ctx: Context, groups: ResolvedTeam[], limit: number): Promise<ListResult<unknown>> {
  const found: unknown[] = [];
  for (let start = 0; start < groups.length; start += ARCHIVE_CHECK_BATCH) {
    const batch = groups.slice(start, start + ARCHIVE_CHECK_BATCH);
    const states = await Promise.all(
      batch.map((g) =>
        ctx.graph.tryGet<{ isArchived?: boolean }>(`/teams/${seg(g.id)}`, { query: { $select: 'isArchived' } }),
      ),
    );
    for (const [index, group] of batch.entries()) {
      if (!states[index]?.isArchived) continue;
      if (found.length >= limit) return { items: found, truncated: true };
      found.push(ctx.full ? group : { ...shapeTeamGroup(group), isArchived: true });
    }
  }
  return { items: found, truncated: false };
}

interface TeamDetail {
  id: string;
  displayName: string;
  description: string | null;
  visibility: string | null;
  isArchived: boolean | null;
  webUrl: string | null;
  createdDateTime: string | null;
  classification: string | null;
  memberSettings?: Record<string, boolean>;
  guestSettings?: Record<string, boolean>;
  messagingSettings?: Record<string, boolean>;
  summary?: { ownersCount?: number; membersCount?: number; guestsCount?: number } | null;
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'teams',
    name: 'list',
    summary: 'List teams in the tenant, optionally by name or archived state',
    description:
      'Lists Microsoft 365 groups that have Teams enabled. --search matches words in the display name. ' +
      '--archived reads each team to find read-only (archived) ones, so it is slower on large tenants. ' +
      'Next: `m365 teams get <team>` for settings, owners and counts.',
    examples: ['m365 teams list --search marketing', 'm365 teams list --archived --all'],
    args: [
      { name: 'search', type: 'string', description: 'Words in the team display name.' },
      { name: 'archived', type: 'boolean', description: 'Only archived (read-only) teams.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['Group.ReadBasic.All', 'Team.ReadBasic.All'],
      application: ['Group.ReadBasic.All', 'Team.ReadBasic.All'],
    },
    async run(ctx, args) {
      const term = args.search?.replace(/"/g, '');
      const limit = listLimit(args);
      const result = await ctx.graph.list<ResolvedTeam>('/groups', {
        query: {
          $filter: IS_TEAM_FILTER,
          $select: TEAM_GROUP_SELECT,
          ...(term ? { $search: `"displayName:${term}"`, $count: true } : {}),
        },
        advanced: Boolean(term),
        limit: args.archived ? Number.POSITIVE_INFINITY : limit,
        pageSize: args.archived ? 999 : 100,
      });
      if (args.archived) return archivedTeams(ctx, result.items, limit);
      return { items: ctx.full ? result.items : result.items.map(shapeTeamGroup), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'get',
    summary: 'Show a team: archived state, key settings, owners, member/guest/channel counts, link',
    description:
      'Use to answer "who owns this team", "can guests/members create channels", "is it archived". ' +
      '`settings` keeps only the member, guest and messaging switches that admins are usually asked about. ' +
      'Sections that fail (missing permission) show `{ error }` instead of failing the whole view.',
    examples: ['m365 teams get "Project Falcon"'],
    args: [TEAM_ARG],
    risk: 'read',
    permissions: {
      delegated: ['TeamSettings.Read.All', 'GroupMember.Read.All', 'Channel.ReadBasic.All', 'Group.ReadBasic.All'],
      application: ['TeamSettings.Read.All', 'GroupMember.Read.All', 'Channel.ReadBasic.All', 'Group.ReadBasic.All'],
      roles: TEAMS_ADMIN,
    },
    async run(ctx, args) {
      const team = await resolveTeam(ctx, args.team);
      const id = seg(team.id);
      const parts = await settle({
        team: ctx.graph.get<TeamDetail>(`/teams/${id}`),
        owners: ctx.graph.list<{ displayName: string; userPrincipalName?: string }>(`/groups/${id}/owners`, {
          query: { $select: 'id,displayName,userPrincipalName', $count: true },
          advanced: true,
          limit: 100,
        }),
        channels: ctx.graph.list<{ membershipType?: string }>(`/teams/${id}/channels`, {
          query: { $select: 'id,membershipType' },
          limit: Number.POSITIVE_INFINITY,
          pageSize: false,
        }),
      });
      if (ctx.full) return parts;
      const detail = 'error' in parts.team ? null : parts.team;
      const channels =
        'error' in parts.channels
          ? parts.channels
          : {
              total: parts.channels.items.length,
              private: parts.channels.items.filter((c) => c.membershipType === 'private').length,
              shared: parts.channels.items.filter((c) => c.membershipType === 'shared').length,
            };
      const pickSettings = (source: Record<string, boolean> | undefined, keys: string[]) =>
        source ? Object.fromEntries(keys.filter((k) => k in source).map((k) => [k, source[k]])) : null;
      return {
        id: team.id,
        displayName: team.displayName,
        description: team.description,
        visibility: detail?.visibility ?? team.visibility,
        mail: team.mail,
        isArchived: detail?.isArchived ?? null,
        webUrl: detail?.webUrl ?? null,
        createdDateTime: detail?.createdDateTime ?? team.createdDateTime,
        counts: detail?.summary
          ? {
              owners: detail.summary.ownersCount ?? null,
              members: detail.summary.membersCount ?? null,
              guests: detail.summary.guestsCount ?? null,
            }
          : null,
        owners:
          'error' in parts.owners ? parts.owners : parts.owners.items.map((o) => o.userPrincipalName ?? o.displayName),
        channels,
        settings: detail
          ? {
              members: pickSettings(detail.memberSettings, [
                'allowCreateUpdateChannels',
                'allowCreatePrivateChannels',
                'allowDeleteChannels',
                'allowAddRemoveApps',
              ]),
              guests: pickSettings(detail.guestSettings, ['allowCreateUpdateChannels', 'allowDeleteChannels']),
              messaging: pickSettings(detail.messagingSettings, [
                'allowUserEditMessages',
                'allowUserDeleteMessages',
                'allowOwnerDeleteMessages',
                'allowTeamMentions',
                'allowChannelMentions',
              ]),
            }
          : parts.team,
      };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'members',
    summary: 'List the owners, members and guests of a team',
    description:
      '`role` is owner, member or guest. `membershipId` is the team membership (not the user id). ' +
      'Next: `m365 teams members add|remove <team> <user>`.',
    examples: ['m365 teams members "Project Falcon" --role owner'],
    args: [
      TEAM_ARG,
      { name: 'role', type: 'string', choices: ['owner', 'member', 'guest'], description: 'Only this role.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['TeamMember.Read.All', 'Group.ReadBasic.All'],
      application: ['TeamMember.Read.All', 'Group.ReadBasic.All'],
    },
    async run(ctx, args) {
      const team = await resolveTeam(ctx, args.team);
      const limit = listLimit(args);
      // Graph only documents $filter on userId for team members, so roles are filtered here.
      const result = await ctx.graph.list<ConversationMember>(`/teams/${seg(team.id)}/members`, {
        limit: args.role ? Number.POSITIVE_INFINITY : limit,
        pageSize: 999,
      });
      const matching = args.role ? result.items.filter((m) => memberRole(m) === args.role) : result.items;
      const items = matching.slice(0, limit);
      return {
        items: ctx.full ? items : items.map(shapeMember),
        truncated: result.truncated || matching.length > limit,
      };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'members add',
    summary: 'Add a user to a team as member or owner',
    description:
      'Adds the user directly to the team (and its Microsoft 365 group). --owner makes them an owner. ' +
      'Guests can be added only with delegated sign-in; app-only adds of guests are not supported by Graph.',
    examples: ['m365 teams members add "Project Falcon" jane@contoso.com', 'm365 teams members add Sales bob --owner'],
    args: [TEAM_ARG, USER_ARG, { name: 'owner', type: 'boolean', description: 'Add as an owner.' }],
    risk: 'write',
    permissions: {
      delegated: ['TeamMember.ReadWrite.All', USER_READ],
      application: ['TeamMember.ReadWrite.All', USER_READ],
    },
    async preview(ctx, args) {
      const { team, user } = await teamAndUser(ctx, args.team, args.user);
      const existing = await findMembership(ctx, team.id, user.id);
      return {
        team: { id: team.id, displayName: team.displayName },
        user: { id: user.id, userPrincipalName: user.userPrincipalName },
        currentRole: existing ? memberRole(existing) : null,
        action: `Add ${user.userPrincipalName} to team "${team.displayName}" as ${args.owner ? 'owner' : 'member'}`,
      };
    },
    async run(ctx, args) {
      const { team, user } = await teamAndUser(ctx, args.team, args.user);
      const member = await ctx.graph.post<ConversationMember>(`/teams/${seg(team.id)}/members`, {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: args.owner ? ['owner'] : [],
        'user@odata.bind': ctx.graph.url(`/users('${user.id}')`),
      });
      return { team: team.displayName, added: shapeMember(member) };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'members remove',
    summary: 'Remove a user from a team',
    description:
      'Removes the team membership (and the Microsoft 365 group membership), so the user loses access to the ' +
      "team's channels and files. A team's last owner cannot be removed.",
    examples: ['m365 teams members remove "Project Falcon" jane@contoso.com'],
    args: [TEAM_ARG, USER_ARG],
    risk: 'destructive',
    permissions: {
      delegated: ['TeamMember.ReadWrite.All', USER_READ],
      application: ['TeamMember.ReadWrite.All', USER_READ],
    },
    async preview(ctx, args) {
      const { team, user } = await teamAndUser(ctx, args.team, args.user);
      const membership = await findMembership(ctx, team.id, user.id);
      if (!membership) {
        throw new CliError('NOT_FOUND', `${user.userPrincipalName} is not a member of "${team.displayName}"`, {
          hint: 'Check with `m365 teams members <team>`; access may come from a shared channel or group nesting.',
        });
      }
      return {
        team: { id: team.id, displayName: team.displayName },
        user: { id: user.id, userPrincipalName: user.userPrincipalName },
        currentRole: memberRole(membership),
        membershipId: membership.id,
        action: `Remove ${user.userPrincipalName} (${memberRole(membership)}) from team "${team.displayName}"`,
      };
    },
    async run(ctx, args) {
      const { team, user } = await teamAndUser(ctx, args.team, args.user);
      const membership = await findMembership(ctx, team.id, user.id);
      if (!membership)
        throw new CliError('NOT_FOUND', `${user.userPrincipalName} is not a member of "${team.displayName}"`);
      await ctx.graph.delete(`/teams/${seg(team.id)}/members/${seg(membership.id)}`);
      return { team: team.displayName, removed: user.userPrincipalName, previousRole: memberRole(membership) };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'channels',
    summary: "List a team's channels with their type (standard, private, shared)",
    description:
      'Private and shared channels have their own membership: use `m365 teams channels members <team> <channel>` ' +
      'when someone can see the team but not a channel.',
    examples: ['m365 teams channels "Project Falcon"'],
    args: [TEAM_ARG, ...LIST_ARGS],
    risk: 'read',
    permissions: {
      delegated: ['Channel.ReadBasic.All', 'Group.ReadBasic.All'],
      application: ['Channel.ReadBasic.All', 'Group.ReadBasic.All'],
    },
    async run(ctx, args) {
      const team = await resolveTeam(ctx, args.team);
      // $select keeps Graph from computing each channel's email address, which is slow.
      const result = await ctx.graph.list(`/teams/${seg(team.id)}/channels`, {
        query: { $select: 'id,displayName,description,membershipType,createdDateTime,webUrl' },
        limit: listLimit(args),
        pageSize: false,
      });
      return result;
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'channels members',
    summary: 'List who is in a private or shared channel',
    description:
      'Standard channels inherit team membership, so use `m365 teams members` for them. Shared channels can ' +
      'include people from other teams or tenants (tenantId differs).',
    examples: ['m365 teams channels members "Project Falcon" "Leadership"'],
    args: [
      TEAM_ARG,
      {
        name: 'channel',
        type: 'string',
        required: true,
        positional: true,
        description: 'Channel name or id (19:...).',
      },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['ChannelMember.Read.All', 'Channel.ReadBasic.All', 'Group.ReadBasic.All'],
      application: ['ChannelMember.Read.All', 'Channel.ReadBasic.All', 'Group.ReadBasic.All'],
    },
    async run(ctx, args) {
      const team = await resolveTeam(ctx, args.team);
      const channel = await resolveChannel(ctx, team.id, args.channel);
      const result = await ctx.graph.list<ConversationMember>(
        `/teams/${seg(team.id)}/channels/${seg(channel.id)}/members`,
        { limit: listLimit(args), pageSize: false },
      );
      return { items: ctx.full ? result.items : result.items.map(shapeMember), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'user',
    summary: 'List the teams a user belongs to',
    description:
      'Answers "which teams is X in". Covers direct team membership only; access through a shared channel ' +
      'does not appear here.',
    examples: ['m365 teams user jane@contoso.com'],
    args: [USER_ARG],
    risk: 'read',
    permissions: {
      delegated: ['Team.ReadBasic.All', USER_READ],
      application: ['Team.ReadBasic.All', USER_READ],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      // joinedTeams accepts no OData query options, including $top.
      const { items } = await ctx.graph.list<{ id: string; displayName: string; isArchived?: boolean | null }>(
        `/users/${seg(user.id)}/joinedTeams`,
        { limit: Number.POSITIVE_INFINITY, pageSize: false },
      );
      return {
        user: user.userPrincipalName,
        count: items.length,
        teams: ctx.full
          ? items
          : items.map((t) => ({ id: t.id, displayName: t.displayName, isArchived: t.isArchived ?? null })),
      };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'presence',
    summary: "Show users' Teams presence (availability, activity, status message, out of office)",
    description:
      'availability is the coloured dot (Available, Busy, DoNotDisturb, Away, Offline...); activity adds detail ' +
      '(InACall, InAMeeting, Presenting...). Useful before calling someone or when a user reports their status is ' +
      'stuck. Accepts several users at once.',
    examples: ['m365 teams presence jane@contoso.com bob@contoso.com'],
    args: [
      {
        name: 'users',
        type: 'string[]',
        required: true,
        positional: true,
        description: 'One or more users (UPN, email, name or id).',
      },
    ],
    risk: 'read',
    permissions: {
      delegated: ['Presence.Read.All', USER_READ],
      application: ['Presence.Read.All', USER_READ],
    },
    async run(ctx, args) {
      const users = await Promise.all(args.users.map((ref) => resolveUser(ctx, ref)));
      const ids = [...new Set(users.map((u) => u.id))];
      // getPresencesByUserId is a POST but only reads; it takes up to 650 ids per call.
      const result = await ctx.graph.post<{ value: Presence[] }>('/communications/getPresencesByUserId', { ids });
      if (ctx.full) return result.value;
      const byId = new Map(users.map((u) => [u.id, u]));
      return result.value.map((p) => ({
        user: byId.get(p.id)?.userPrincipalName ?? p.id,
        availability: p.availability ?? null,
        activity: p.activity ?? null,
        statusMessage: p.statusMessage?.message?.content || null,
        isOutOfOffice: p.outOfOfficeSettings?.isOutOfOffice ?? null,
      }));
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'create',
    summary: 'Create a new team from the standard template',
    description:
      'Creates the Microsoft 365 group and team together. Graph accepts the request and provisions the team in ' +
      'the background, which can take several minutes. The result has the new teamId and operationId; `next` ' +
      'polls the operation (status succeeded/failed, wait 30+ seconds between checks). ' +
      'With app-only auth --owner is required (Graph needs an owner); with delegated sign-in the caller becomes ' +
      'an owner and --owner adds another.',
    examples: ['m365 teams create --name "Project Falcon" --description "Falcon launch" --owner jane@contoso.com'],
    args: [
      { name: 'name', type: 'string', required: true, description: 'Team display name.' },
      { name: 'description', type: 'string', description: 'Team description.' },
      {
        name: 'visibility',
        type: 'string',
        choices: ['private', 'public'],
        default: 'private',
        description: 'Who can find and join.',
      },
      { name: 'owner', type: 'string', description: 'Owner (UPN, email, name or id). Required with app-only auth.' },
    ],
    risk: 'write',
    permissions: {
      delegated: ['Team.Create', USER_READ],
      application: ['Team.Create', USER_READ],
    },
    async preview(ctx, args) {
      const owner = args.owner ? await resolveUser(ctx, args.owner) : null;
      return {
        displayName: args.name,
        description: args.description ?? null,
        visibility: args.visibility,
        owner: owner?.userPrincipalName ?? (ctx.profile.authMode === 'app' ? null : 'signed-in user'),
        action: `Create ${args.visibility} team "${args.name}"`,
      };
    },
    async run(ctx, args) {
      if (!args.owner && ctx.profile.authMode === 'app') {
        throw new CliError('USAGE', 'Creating a team with app-only auth needs --owner', {
          hint: 'Pass --owner <user>; Graph requires at least one owner when no user is signed in.',
        });
      }
      const owner = args.owner ? await resolveUser(ctx, args.owner) : null;
      const response = await ctx.graph.requestWithHeaders('POST', '/teams', {
        body: {
          'template@odata.bind': ctx.graph.url("/teamsTemplates('standard')"),
          displayName: args.name,
          ...(args.description ? { description: args.description } : {}),
          visibility: args.visibility,
          ...(owner
            ? {
                members: [
                  {
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    roles: ['owner'],
                    'user@odata.bind': ctx.graph.url(`/users('${owner.id}')`),
                  },
                ],
              }
            : {}),
        },
      });
      const operation = teamsOperation(response.headers);
      return {
        status: 'accepted',
        teamId: operation.teamId,
        operationId: operation.operationId,
        displayName: args.name,
        visibility: args.visibility,
        owner: owner?.userPrincipalName ?? null,
        note: 'Provisioning runs in the background and can take a few minutes.',
        next: operation.check ?? `m365 teams get "${args.name}"`,
      };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'archive',
    summary: 'Archive a team (make it read-only)',
    description:
      'Members can still read but no longer post, and the team moves to the archived list. Reversible with ' +
      '`m365 teams unarchive`. --spo-read-only also makes the SharePoint site read-only for members (delegated ' +
      'sign-in only).',
    examples: ['m365 teams archive "Project Falcon"'],
    args: [
      TEAM_ARG,
      {
        name: 'spoReadOnly',
        type: 'boolean',
        description: 'Also make the team SharePoint site read-only for members.',
      },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['TeamSettings.ReadWrite.All', 'Group.ReadBasic.All'],
      application: ['TeamSettings.ReadWrite.All', 'Group.ReadBasic.All'],
      roles: TEAMS_ADMIN,
    },
    async preview(ctx, args) {
      const team = await resolveTeam(ctx, args.team);
      const detail = await ctx.graph.get<TeamDetail>(`/teams/${seg(team.id)}`, {
        query: { $select: 'id,displayName,isArchived,summary' },
      });
      return {
        team: {
          id: team.id,
          displayName: team.displayName,
          isArchived: detail.isArchived,
          members: detail.summary?.membersCount ?? null,
        },
        action: detail.isArchived
          ? `Team "${team.displayName}" is already archived; nothing will change`
          : `Archive team "${team.displayName}" so it becomes read-only${args.spoReadOnly ? ', including its SharePoint site for members' : ''}`,
      };
    },
    async run(ctx, args) {
      const team = await resolveTeam(ctx, args.team);
      if (args.spoReadOnly && ctx.profile.authMode === 'app') {
        throw new CliError('USAGE', '--spo-read-only is not supported with app-only auth', {
          hint: 'Run without --spo-read-only, or use a delegated profile.',
        });
      }
      const response = await ctx.graph.requestWithHeaders('POST', `/teams/${seg(team.id)}/archive`, {
        body: args.spoReadOnly ? { shouldSetSpoSiteReadOnlyForMembers: true } : undefined,
      });
      const { operationId, check } = teamsOperation(response.headers);
      return {
        status: 'accepted',
        team: team.displayName,
        teamId: team.id,
        operationId,
        next: check,
        note: 'Archiving completes in the background.',
      };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'unarchive',
    summary: 'Restore an archived team to read-write',
    examples: ['m365 teams unarchive "Project Falcon"'],
    args: [TEAM_ARG],
    risk: 'write',
    permissions: {
      delegated: ['TeamSettings.ReadWrite.All', 'Group.ReadBasic.All'],
      application: ['TeamSettings.ReadWrite.All', 'Group.ReadBasic.All'],
      roles: TEAMS_ADMIN,
    },
    async run(ctx, args) {
      const team = await resolveTeam(ctx, args.team);
      const response = await ctx.graph.requestWithHeaders('POST', `/teams/${seg(team.id)}/unarchive`);
      const { operationId, check } = teamsOperation(response.headers);
      return {
        status: 'accepted',
        team: team.displayName,
        teamId: team.id,
        operationId,
        next: check,
        note: 'Unarchiving completes in the background.',
      };
    },
  }),

  ...callCommands,
];

interface Presence {
  id: string;
  availability?: string;
  activity?: string;
  statusMessage?: { message?: { content?: string } } | null;
  outOfOfficeSettings?: { isOutOfOffice?: boolean } | null;
}
