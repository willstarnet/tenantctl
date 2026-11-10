import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { pickOne } from '../../core/resolve.ts';
import type { Context } from '../../core/types.ts';
import { isGuid } from '../../core/util.ts';

export interface ResolvedTeam {
  id: string;
  displayName: string;
  description: string | null;
  visibility: string | null;
  mail: string | null;
  createdDateTime: string | null;
  resourceProvisioningOptions?: string[];
}

export interface ResolvedChannel {
  id: string;
  displayName: string;
  membershipType: string | null;
}

export const TEAM_GROUP_SELECT =
  'id,displayName,description,visibility,mail,createdDateTime,resourceProvisioningOptions';
export const IS_TEAM_FILTER = "resourceProvisioningOptions/Any(x:x eq 'Team')";

function describeTeam(team: ResolvedTeam) {
  return { id: team.id, displayName: team.displayName, mail: team.mail, visibility: team.visibility };
}

/**
 * A team is the Microsoft 365 group behind it, so the group id is the team id. Accepts a group id, a Teams
 * link (`...?groupId=<id>`), the group's mail address or the team's display name.
 */
export async function resolveTeam(ctx: Context, ref: string): Promise<ResolvedTeam> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A team is required');
  const fromLink = value.match(/[?&]groupId=([0-9a-f-]{36})/i)?.[1];
  const id = fromLink ?? (isGuid(value) ? value : undefined);
  if (id) {
    const group = await ctx.graph.tryGet<ResolvedTeam>(`/groups/${seg(id)}`, { query: { $select: TEAM_GROUP_SELECT } });
    if (!group)
      throw new CliError('NOT_FOUND', `No team with id ${id}`, {
        hint: 'Find the team with `m365 teams list --search <name>`.',
      });
    if (group.resourceProvisioningOptions && !group.resourceProvisioningOptions.includes('Team')) {
      throw new CliError('NOT_FOUND', `Group "${group.displayName}" (${id}) is not a team`, {
        hint: 'It is a Microsoft 365 group without Teams. Use the entra group commands for it instead.',
      });
    }
    return group;
  }
  const match = value.includes('@') ? `mail eq ${odata(value)}` : `displayName eq ${odata(value)}`;
  const exact = await ctx.graph.list<ResolvedTeam>('/groups', {
    query: { $filter: `${IS_TEAM_FILTER} and ${match}`, $select: TEAM_GROUP_SELECT },
    limit: 10,
  });
  if (exact.items.length) return pickOne(exact.items, value, 'team', (t) => [t.displayName, t.mail], describeTeam);
  const fuzzy = await ctx.graph.list<ResolvedTeam>('/groups', {
    query: {
      $filter: IS_TEAM_FILTER,
      $search: `"displayName:${value.replace(/"/g, '')}"`,
      $select: TEAM_GROUP_SELECT,
      $count: true,
    },
    advanced: true,
    limit: 25,
  });
  return pickOne(fuzzy.items, value, 'team', (t) => [t.displayName], describeTeam);
}

/** Channel by id (`19:...@thread...`) or display name within a team. */
export async function resolveChannel(ctx: Context, teamId: string, ref: string): Promise<ResolvedChannel> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A channel is required');
  if (value.startsWith('19:')) {
    const channel = await ctx.graph.tryGet<ResolvedChannel>(`/teams/${seg(teamId)}/channels/${seg(value)}`, {
      query: { $select: 'id,displayName,membershipType' },
    });
    if (!channel) throw new CliError('NOT_FOUND', `No channel ${value} in this team`);
    return channel;
  }
  const { items } = await ctx.graph.list<ResolvedChannel>(`/teams/${seg(teamId)}/channels`, {
    query: { $select: 'id,displayName,membershipType' },
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  const wanted = value.toLowerCase();
  const candidates = items.filter((c) => c.displayName.toLowerCase().includes(wanted));
  return pickOne(
    candidates,
    value,
    'channel',
    (c) => [c.displayName],
    (c) => ({
      id: c.id,
      displayName: c.displayName,
      membershipType: c.membershipType,
    }),
  );
}
