import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import type { Context } from '../../core/types.ts';
import { settle } from '../../core/util.ts';

export const USER_ARG = {
  name: 'user',
  type: 'string',
  required: true,
  positional: true,
  description: 'Mailbox owner: UPN, email/alias, display name, object id, or "me".',
} as const;

/** resolveUser looks users up by name, mail and proxyAddresses before any mailbox call. */
export const RESOLVE_PERMISSION = 'User.Read.All';

export const MAILBOX_ACCESS_NOTE =
  "Reading or changing another person's mailbox needs app-only auth (application permissions); a delegated " +
  'sign-in only reaches your own mailbox and mailboxes shared with you.';

/** Aliases people say -> Graph well-known folder names. */
export const WELL_KNOWN_FOLDERS: Record<string, string> = {
  inbox: 'inbox',
  sent: 'sentitems',
  sentitems: 'sentitems',
  'sent items': 'sentitems',
  deleted: 'deleteditems',
  deleteditems: 'deleteditems',
  'deleted items': 'deleteditems',
  junk: 'junkemail',
  junkemail: 'junkemail',
  'junk email': 'junkemail',
  archive: 'archive',
  drafts: 'drafts',
  outbox: 'outbox',
};

/**
 * Folder ids -> display names. Rules and messages only carry folder ids; one GET per distinct id keeps this
 * correct for nested folders too. A folder that can't be read maps to null so callers can fall back to the id.
 */
export async function resolveFolderNames(
  ctx: Context,
  userId: string,
  ids: Iterable<string | null | undefined>,
): Promise<Record<string, string | null>> {
  const unique = [...new Set([...ids].filter((id): id is string => Boolean(id)))];
  const tasks: Record<string, Promise<{ displayName?: string }>> = {};
  for (const id of unique) {
    tasks[id] = ctx.graph.get(`/users/${seg(userId)}/mailFolders/${seg(id)}`, { query: { $select: 'displayName' } });
  }
  const settled = await settle(tasks);
  const names: Record<string, string | null> = {};
  for (const id of unique) {
    const folder = settled[id] as { displayName?: string } | { error: unknown };
    names[id] = 'displayName' in folder && folder.displayName ? folder.displayName : null;
  }
  return names;
}

/** Folder argument -> path segment: a well-known name, or a top-level folder looked up by display name. */
export async function resolveFolder(ctx: Context, userId: string, ref: string): Promise<string> {
  const known = WELL_KNOWN_FOLDERS[ref.trim().toLowerCase()];
  if (known) return known;
  const { items } = await ctx.graph.list<{ id: string; displayName: string }>(`/users/${seg(userId)}/mailFolders`, {
    query: { $filter: `displayName eq ${odata(ref.trim())}`, $select: 'id,displayName' },
    limit: 5,
  });
  const folder = items[0];
  if (!folder) {
    throw new CliError('NOT_FOUND', `No top-level mail folder named "${ref}"`, {
      hint: 'Use a well-known folder (inbox, sent, deleted, junk, archive, drafts) or the exact top-level folder name.',
    });
  }
  return folder.id;
}

export interface Recipient {
  emailAddress?: { name?: string | null; address?: string | null } | null;
}

export function addresses(recipients: Recipient[] | null | undefined): string[] {
  return (recipients ?? []).map((r) => r.emailAddress?.address ?? r.emailAddress?.name ?? '').filter(Boolean);
}
