import { CliError } from './errors.ts';
import { odata, seg } from './graph.ts';
import type { Context } from './types.ts';
import { isGuid } from './util.ts';

/**
 * Users type names, not object ids. Every command that targets a user, group, device, team or site accepts
 * whatever identifier the person would naturally say, and resolves it here. Ambiguity is an error that lists
 * the candidates, so the agent can ask the user which one they meant.
 */

export interface ResolvedUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
}

export interface ResolvedGroup {
  id: string;
  displayName: string;
  mail: string | null;
}

const USER_SELECT = 'id,displayName,userPrincipalName,mail';
const GROUP_SELECT = 'id,displayName,mail';

/**
 * Narrow candidates to one. A unique case-insensitive exact match on any of `keys` wins over partial matches.
 */
export function pickOne<T>(
  candidates: T[],
  ref: string,
  kind: string,
  keys: (item: T) => (string | null | undefined)[],
  describe: (item: T) => unknown,
): T {
  if (candidates.length === 1) return candidates[0] as T;
  if (candidates.length === 0) {
    throw new CliError('NOT_FOUND', `No ${kind} matches "${ref}"`, {
      hint: `Check the spelling, or search with the ${kind} list command.`,
    });
  }
  const wanted = ref.toLowerCase();
  const exact = candidates.filter((item) => keys(item).some((key) => key?.toLowerCase() === wanted));
  if (exact.length === 1) return exact[0] as T;
  const shown = (exact.length ? exact : candidates).slice(0, 10);
  throw new CliError('AMBIGUOUS', `"${ref}" matches ${exact.length || candidates.length} ${kind}s`, {
    hint: `Ask the user which ${kind} they mean, then pass its id.`,
    details: shown.map(describe),
  });
}

export async function resolveUser(ctx: Context, ref: string, select = USER_SELECT): Promise<ResolvedUser> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A user is required');
  if (value.toLowerCase() === 'me') return ctx.graph.get<ResolvedUser>('/me', { query: { $select: select } });
  if (isGuid(value) || value.includes('@')) {
    const direct = await ctx.graph.tryGet<ResolvedUser>(`/users/${seg(value)}`, { query: { $select: select } });
    if (direct) return direct;
    if (isGuid(value)) throw new CliError('NOT_FOUND', `No user with id ${value}`);
    // Not a UPN: try primary SMTP address and aliases.
    const { items } = await ctx.graph.list<ResolvedUser>('/users', {
      query: {
        $filter: `mail eq ${odata(value)} or proxyAddresses/any(p:p eq ${odata(`smtp:${value}`)})`,
        $select: select,
        $count: true,
      },
      advanced: true,
      limit: 10,
    });
    return pickOne(items, value, 'user', (u) => [u.mail, u.userPrincipalName], describeUser);
  }
  const term = value.replace(/"/g, '');
  const { items } = await ctx.graph.list<ResolvedUser>('/users', {
    query: {
      $search: `"displayName:${term}" OR "mailNickname:${term}" OR "userPrincipalName:${term}"`,
      $select: select,
      $count: true,
    },
    advanced: true,
    limit: 25,
  });
  return pickOne(items, value, 'user', (u) => [u.displayName, u.userPrincipalName.split('@')[0]], describeUser);
}

function describeUser(u: ResolvedUser) {
  return { id: u.id, displayName: u.displayName, userPrincipalName: u.userPrincipalName };
}

export async function resolveGroup(ctx: Context, ref: string, select = GROUP_SELECT): Promise<ResolvedGroup> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A group is required');
  if (isGuid(value)) {
    const group = await ctx.graph.tryGet<ResolvedGroup>(`/groups/${seg(value)}`, { query: { $select: select } });
    if (!group) throw new CliError('NOT_FOUND', `No group with id ${value}`);
    return group;
  }
  const filter = value.includes('@')
    ? `mail eq ${odata(value)}`
    : `displayName eq ${odata(value)} or mailNickname eq ${odata(value)}`;
  const exact = await ctx.graph.list<ResolvedGroup>('/groups', {
    query: { $filter: filter, $select: select },
    limit: 10,
  });
  if (exact.items.length) {
    return pickOne(exact.items, value, 'group', (g) => [g.displayName, g.mail], describeGroup);
  }
  const term = value.replace(/"/g, '');
  const fuzzy = await ctx.graph.list<ResolvedGroup>('/groups', {
    query: { $search: `"displayName:${term}"`, $select: select, $count: true },
    advanced: true,
    limit: 25,
  });
  return pickOne(fuzzy.items, value, 'group', (g) => [g.displayName], describeGroup);
}

function describeGroup(g: ResolvedGroup) {
  return { id: g.id, displayName: g.displayName, mail: g.mail };
}
