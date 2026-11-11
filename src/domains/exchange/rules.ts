import { CliError } from '../../core/errors.ts';
import { exoRun } from '../../core/exo.ts';
import { seg } from '../../core/graph.ts';
import { pickOne, resolveGroup, resolveUser, type ResolvedUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import {
  addresses,
  MAILBOX_ACCESS_NOTE,
  RESOLVE_PERMISSION,
  type Recipient,
  resolveFolderNames,
  USER_ARG,
} from './common.ts';

type Predicates = Record<string, unknown>;

interface RuleActions {
  assignCategories?: string[] | null;
  copyToFolder?: string | null;
  delete?: boolean | null;
  forwardAsAttachmentTo?: Recipient[] | null;
  forwardTo?: Recipient[] | null;
  markAsRead?: boolean | null;
  markImportance?: string | null;
  moveToFolder?: string | null;
  permanentDelete?: boolean | null;
  redirectTo?: Recipient[] | null;
  stopProcessingRules?: boolean | null;
}

export interface RawInboxRule {
  id: string;
  displayName?: string | null;
  sequence?: number | null;
  isEnabled?: boolean | null;
  hasError?: boolean | null;
  isReadOnly?: boolean | null;
  conditions?: Predicates | null;
  exceptions?: Predicates | null;
  actions?: RuleActions | null;
}

export interface ShapedInboxRule {
  id: string;
  name: string | null;
  enabled: boolean;
  sequence: number | null;
  hasError: boolean;
  conditions: string[];
  exceptions: string[];
  actions: string[];
  forwards: boolean;
  forwardTargets: string[];
  /** null when tenant domains were unavailable, so internal/external can't be told apart. */
  forwardsExternally: boolean | null;
  externalTargets: string[] | null;
  deletes: boolean;
  movesToFolder: string | null;
  marksRead: boolean;
  suspicious: boolean;
  reason: string | null;
}

export interface ShapeRuleOptions {
  /** Verified tenant domains (lowercase). Omit when unknown. */
  tenantDomains?: readonly string[] | null;
  /** Folder id -> display name, for move/copy actions. */
  folderNames?: Record<string, string | null>;
}

// Folders attackers use to hide replies from the victim: rarely opened, or emptied automatically.
const HIDING_FOLDER =
  /^(rss( feeds| subscriptions)?|archive|conversation history|junk e-?mail|deleted items|notes|sync issues|outbox)$/i;
// Business email compromise rules typically hide replies about payments or the compromise itself.
const BEC_KEYWORDS =
  /\b(invoice|payment|paid|wire|bank|remit|remittance|ach|transfer|swift|iban|payroll|w-?2|password|security|hack(ed)?|phish|compromis\w*|fraud|spam|virus|helpdesk|it support)\b/i;

function predicateSummary(predicates: Predicates | null | undefined): string[] {
  const out: string[] = [];
  for (const [key, value] of Object.entries(predicates ?? {})) {
    if (value === null || value === undefined || value === false) continue;
    if (Array.isArray(value)) {
      if (!value.length) continue;
      const items =
        typeof value[0] === 'object'
          ? addresses(value as Recipient[])
          : (value as unknown[]).map((v) => `"${String(v)}"`);
      out.push(`${key}: ${items.join(', ')}`);
    } else if (value === true) {
      out.push(key);
    } else if (key === 'withinSizeRange' && typeof value === 'object') {
      const range = value as { minimumSize?: number; maximumSize?: number };
      out.push(`size ${range.minimumSize ?? 0}-${range.maximumSize ?? '∞'} KB`);
    } else {
      out.push(`${key}: ${String(value)}`);
    }
  }
  return out;
}

/** True when the address is in a verified tenant domain (or a subdomain), or has no domain at all. */
export function isInternal(address: string, tenantDomains: readonly string[]): boolean {
  if (!address.includes('@')) return true; // Display-name-only recipients resolve inside the directory.
  const domain = address.slice(address.lastIndexOf('@') + 1).toLowerCase();
  return tenantDomains.some((d) => domain === d || domain.endsWith(`.${d}`));
}

/** Compact a raw messageRule and derive the BEC indicators admins look for. */
export function shapeInboxRule(rule: RawInboxRule, options: ShapeRuleOptions = {}): ShapedInboxRule {
  const actions = rule.actions ?? {};
  const folderName = (id: string | null | undefined) => (id ? (options.folderNames?.[id] ?? id) : null);
  const forwardTargets = [
    ...addresses(actions.forwardTo),
    ...addresses(actions.forwardAsAttachmentTo),
    ...addresses(actions.redirectTo),
  ];
  const domains = options.tenantDomains?.map((d) => d.toLowerCase());
  const externalTargets = domains ? forwardTargets.filter((a) => !isInternal(a, domains)) : null;
  const deletes = Boolean(actions.delete || actions.permanentDelete);
  const movesToFolder = folderName(actions.moveToFolder);
  const marksRead = Boolean(actions.markAsRead);

  const actionSummary: string[] = [];
  if (actions.forwardTo?.length) actionSummary.push(`forward to ${addresses(actions.forwardTo).join(', ')}`);
  if (actions.forwardAsAttachmentTo?.length) {
    actionSummary.push(`forward as attachment to ${addresses(actions.forwardAsAttachmentTo).join(', ')}`);
  }
  if (actions.redirectTo?.length) actionSummary.push(`redirect to ${addresses(actions.redirectTo).join(', ')}`);
  if (actions.permanentDelete) actionSummary.push('permanently delete');
  else if (actions.delete) actionSummary.push('delete');
  if (movesToFolder) actionSummary.push(`move to "${movesToFolder}"`);
  if (actions.copyToFolder) actionSummary.push(`copy to "${folderName(actions.copyToFolder)}"`);
  if (marksRead) actionSummary.push('mark as read');
  if (actions.markImportance) actionSummary.push(`mark importance ${actions.markImportance}`);
  if (actions.assignCategories?.length) actionSummary.push(`categorize ${actions.assignCategories.join(', ')}`);
  if (actions.stopProcessingRules) actionSummary.push('stop processing more rules');

  const conditions = predicateSummary(rule.conditions);
  const reasons: string[] = [];
  if (externalTargets?.length) reasons.push(`forwards to external ${externalTargets.join(', ')}`);
  const hides = deletes || (movesToFolder !== null && HIDING_FOLDER.test(movesToFolder));
  const hideVerb = deletes ? 'deletes' : `moves to "${movesToFolder}"`;
  if (hides && marksRead) reasons.push(`${hideVerb} and marks read, hiding mail from the user`);
  const keyword = hides ? conditions.join(' ').match(BEC_KEYWORDS) : null;
  if (keyword) reasons.push(`${hideVerb} mail mentioning "${keyword[0]}"`);
  const name = rule.displayName?.trim() ?? '';
  if ((hides || forwardTargets.length) && !/[a-z0-9]{3,}/i.test(name)) {
    reasons.push(`meaningless rule name "${name}"`);
  }

  return {
    id: rule.id,
    name: rule.displayName ?? null,
    enabled: Boolean(rule.isEnabled),
    sequence: rule.sequence ?? null,
    hasError: Boolean(rule.hasError),
    conditions: conditions.length ? conditions : ['all messages'],
    exceptions: predicateSummary(rule.exceptions),
    actions: actionSummary,
    forwards: forwardTargets.length > 0,
    forwardTargets,
    forwardsExternally: externalTargets ? externalTargets.length > 0 : null,
    externalTargets,
    deletes,
    movesToFolder,
    marksRead,
    suspicious: reasons.length > 0,
    reason: reasons.length ? reasons.join('; ') : null,
  };
}

/** Verified domains of the tenant, lowercase. Anything else is external. */
export async function listTenantDomains(ctx: Context): Promise<string[]> {
  const { items } = await ctx.graph.list<{ id: string; isVerified?: boolean }>('/domains', {
    query: { $select: 'id,isVerified' },
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  return items.filter((d) => d.isVerified).map((d) => d.id.toLowerCase());
}

/** Tenant domains, or null plus a note when /domains is not readable (forwardsExternally becomes unknown). */
export async function tryTenantDomains(ctx: Context): Promise<{ domains: string[] | null; note?: string }> {
  try {
    return { domains: await listTenantDomains(ctx) };
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    return {
      domains: null,
      note: `forwardsExternally is null: tenant domains unavailable (${reason}); needs Domain.Read.All`,
    };
  }
}

function fetchInboxRules(ctx: Context, userId: string): Promise<RawInboxRule[]> {
  return ctx.graph
    .list<RawInboxRule>(`/users/${seg(userId)}/mailFolders/inbox/messageRules`, {
      limit: Number.POSITIVE_INFINITY,
      pageSize: false,
    })
    .then((r) => r.items);
}

/** Inbox rules of one mailbox, shaped, with move/copy folder ids resolved to names. */
export async function getInboxRules(
  ctx: Context,
  userId: string,
  tenantDomains?: readonly string[] | null,
): Promise<ShapedInboxRule[]> {
  const rules = await fetchInboxRules(ctx, userId);
  const folderNames = await resolveFolderNames(
    ctx,
    userId,
    rules.flatMap((r) => [r.actions?.moveToFolder, r.actions?.copyToFolder]),
  );
  return rules
    .map((rule) => shapeInboxRule(rule, { tenantDomains, folderNames }))
    .sort((a, b) => (a.sequence ?? 0) - (b.sequence ?? 0));
}

/**
 * Inbox rules of many mailboxes via `$batch` (20 per call): one round for the rules, one for every move/copy
 * folder name. Each entry is the shaped rules or the CliError for that mailbox, in input order.
 */
export async function getInboxRulesBatch(
  ctx: Context,
  userIds: readonly string[],
  tenantDomains?: readonly string[] | null,
): Promise<(ShapedInboxRule[] | CliError)[]> {
  const responses = await ctx.graph.batch<{ value?: RawInboxRule[]; '@odata.nextLink'?: string }>(
    userIds.map((id) => ({ path: `/users/${seg(id)}/mailFolders/inbox/messageRules` })),
  );
  const raw = await Promise.all(
    responses.map(async (response, index): Promise<RawInboxRule[] | CliError> => {
      if (response.error) return response.error;
      // Rules rarely page; when they do, read that mailbox the normal way.
      if (response.body?.['@odata.nextLink']) {
        return fetchInboxRules(ctx, userIds[index] as string).catch((error: unknown) =>
          error instanceof CliError ? error : new CliError('INTERNAL', String(error)),
        );
      }
      return response.body?.value ?? [];
    }),
  );
  const folderRequests: { userId: string; folderId: string }[] = [];
  raw.forEach((rules, index) => {
    if (rules instanceof CliError) return;
    const ids = new Set(rules.flatMap((r) => [r.actions?.moveToFolder, r.actions?.copyToFolder]).filter(Boolean));
    for (const folderId of ids) folderRequests.push({ userId: userIds[index] as string, folderId: folderId as string });
  });
  const folderResponses = folderRequests.length
    ? await ctx.graph.batch<{ displayName?: string }>(
        folderRequests.map((f) => ({
          path: `/users/${seg(f.userId)}/mailFolders/${seg(f.folderId)}`,
          query: { $select: 'displayName' },
        })),
      )
    : [];
  const folderNames = new Map<string, Record<string, string | null>>();
  folderRequests.forEach((f, index) => {
    const names = folderNames.get(f.userId) ?? {};
    names[f.folderId] = folderResponses[index]?.error ? null : (folderResponses[index]?.body?.displayName ?? null);
    folderNames.set(f.userId, names);
  });
  return raw.map((rules, index) => {
    if (rules instanceof CliError) return rules;
    const names = folderNames.get(userIds[index] as string) ?? {};
    return rules
      .map((rule) => shapeInboxRule(rule, { tenantDomains, folderNames: names }))
      .sort((a, b) => (a.sequence ?? 0) - (b.sequence ?? 0));
  });
}

export interface MailboxForwarding {
  /** External-capable SMTP forwarding (set by admins or the user in OWA), without the `smtp:` prefix. */
  forwardingSmtpAddress: string | null;
  /** Forwarding to a recipient in the organization (mail user/contact/mailbox), as Exchange names it. */
  forwardingAddress: string | null;
  deliverToMailboxAndForward: boolean;
}

export function shapeMailboxForwarding(mailbox: Record<string, unknown>): MailboxForwarding {
  const smtp = typeof mailbox.ForwardingSmtpAddress === 'string' ? mailbox.ForwardingSmtpAddress.trim() : '';
  return {
    forwardingSmtpAddress: smtp ? smtp.replace(/^smtp:/i, '') : null,
    forwardingAddress:
      typeof mailbox.ForwardingAddress === 'string' && mailbox.ForwardingAddress ? mailbox.ForwardingAddress : null,
    deliverToMailboxAndForward: mailbox.DeliverToMailboxAndForward === true,
  };
}

const FORWARDING_SELECT = ['ForwardingSmtpAddress', 'ForwardingAddress', 'DeliverToMailboxAndForward'];

/** Mailbox-level forwarding via Exchange Online PowerShell (`Get-Mailbox`). Throws CliError when the bridge fails. */
export async function getMailboxForwarding(ctx: Context, identity: string): Promise<MailboxForwarding> {
  const [mailboxes] = await exoRun(ctx.exo, [
    { cmdlet: 'Get-Mailbox', params: { Identity: identity }, select: FORWARDING_SELECT },
  ]);
  const mailbox = mailboxes?.[0];
  if (!mailbox) throw new CliError('NOT_FOUND', `No Exchange mailbox "${identity}"`);
  return shapeMailboxForwarding(mailbox);
}

async function findRule(ctx: Context, userRef: string, ruleRef: string) {
  const user = await resolveUser(ctx, userRef);
  const rules = await fetchInboxRules(ctx, user.id);
  const wanted = ruleRef.trim().toLowerCase();
  const byId = rules.filter((r) => r.id === ruleRef.trim());
  const candidates = byId.length ? byId : rules.filter((r) => (r.displayName ?? '').toLowerCase().includes(wanted));
  const rule = pickOne(
    candidates,
    ruleRef,
    'inbox rule',
    (r) => [r.displayName, r.id],
    (r) => ({ id: r.id, name: r.displayName }),
  );
  const folderNames = await resolveFolderNames(ctx, user.id, [rule.actions?.moveToFolder, rule.actions?.copyToFolder]);
  return { user, rule, shaped: shapeInboxRule(rule, { folderNames }) };
}

function ruleTarget(user: ResolvedUser, shaped: ShapedInboxRule) {
  return {
    user: user.userPrincipalName,
    rule: {
      id: shaped.id,
      name: shaped.name,
      enabled: shaped.enabled,
      conditions: shaped.conditions,
      actions: shaped.actions,
    },
  };
}

const RULE_ARG = {
  name: 'rule',
  type: 'string',
  required: true,
  positional: true,
  description: 'Rule id or name (from `exchange rules list`).',
} as const;

const RULE_READ = {
  delegated: ['MailboxSettings.Read', 'Mail.ReadBasic', 'Domain.Read.All', RESOLVE_PERMISSION],
  application: ['MailboxSettings.Read', 'Mail.ReadBasic.All', 'Domain.Read.All', RESOLVE_PERMISSION],
};

interface AuditUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
}

async function auditTargets(
  ctx: Context,
  args: { users?: string[]; group?: string; maxUsers: number },
): Promise<{ users: AuditUser[]; truncated: boolean; errors: { user: string; code: string; message: string }[] }> {
  const errors: { user: string; code: string; message: string }[] = [];
  if (args.users?.length) {
    const users: AuditUser[] = [];
    for (const ref of args.users) {
      try {
        users.push(await resolveUser(ctx, ref));
      } catch (error) {
        errors.push({
          user: ref,
          code: error instanceof CliError ? error.code : 'INTERNAL',
          message: (error as Error).message,
        });
      }
    }
    return { users, truncated: false, errors };
  }
  const select = 'id,displayName,userPrincipalName';
  if (args.group) {
    const group = await resolveGroup(ctx, args.group);
    const r = await ctx.graph.list<AuditUser>(`/groups/${seg(group.id)}/transitiveMembers/microsoft.graph.user`, {
      query: { $select: select, $count: true },
      advanced: true,
      limit: args.maxUsers,
    });
    return { users: r.items, truncated: r.truncated, errors };
  }
  const r = await ctx.graph.list<AuditUser>('/users', {
    query: { $filter: "assignedLicenses/$count ne 0 and userType eq 'Member'", $select: select, $count: true },
    advanced: true,
    limit: args.maxUsers,
  });
  return { users: r.items, truncated: r.truncated, errors };
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'exchange',
    name: 'rules list',
    summary: "List a mailbox's inbox rules with forwarding and BEC red flags",
    description:
      'Use when mail goes missing, a user says replies vanish, or you suspect a compromised account. Each rule has ' +
      'readable conditions/actions plus derived flags: forwards, forwardsExternally (target domain not a verified ' +
      'tenant domain), deletes, movesToFolder, marksRead, and suspicious with a reason (external forwarding, or ' +
      'delete/move-to-RSS/Archive combined with mark-as-read or payment keywords: classic business email ' +
      'compromise). Suspicious rule on an account: check `tenantctl entra users signins <user>`, then ' +
      '`tenantctl exchange rules delete`. Only Outlook inbox rules are visible; mailbox-level forwarding is not (see ' +
      `\`tenantctl exchange forwarding\`). ${MAILBOX_ACCESS_NOTE}`,
    examples: ['tenantctl exchange rules list jane@contoso.com'],
    args: [USER_ARG],
    risk: 'read',
    permissions: RULE_READ,
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      if (ctx.full) return { items: await fetchInboxRules(ctx, user.id), truncated: false };
      const { domains, note } = await tryTenantDomains(ctx);
      if (note) ctx.log(note);
      return { items: await getInboxRules(ctx, user.id, domains), truncated: false };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'rules delete',
    summary: 'Delete an inbox rule (e.g. a malicious forwarding rule)',
    description:
      'Incident response: removes the rule permanently. Previews the rule first. Consider `rules disable` if you ' +
      `want to keep it for investigation. ${MAILBOX_ACCESS_NOTE}`,
    examples: [
      'tenantctl exchange rules delete jane@contoso.com "."',
      'tenantctl exchange rules delete jane@contoso.com AQAAAJ5...',
    ],
    args: [USER_ARG, RULE_ARG],
    risk: 'destructive',
    permissions: {
      delegated: ['MailboxSettings.ReadWrite', 'Mail.ReadBasic', RESOLVE_PERMISSION],
      application: ['MailboxSettings.ReadWrite', 'Mail.ReadBasic.All', RESOLVE_PERMISSION],
    },
    async preview(ctx, args) {
      const { user, shaped } = await findRule(ctx, args.user, args.rule);
      return {
        ...ruleTarget(user, shaped),
        action: `Permanently delete inbox rule "${shaped.name}" from ${user.userPrincipalName}`,
      };
    },
    async run(ctx, args) {
      const { user, shaped } = await findRule(ctx, args.user, args.rule);
      await ctx.graph.delete(`/users/${seg(user.id)}/mailFolders/inbox/messageRules/${seg(shaped.id)}`);
      return { deleted: true, ...ruleTarget(user, shaped) };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'rules disable',
    summary: 'Turn off an inbox rule without deleting it',
    description:
      'Stops a rule from running while keeping it for investigation; re-enable it in Outlook or with ' +
      `\`tenantctl graph patch\`. ${MAILBOX_ACCESS_NOTE}`,
    examples: ['tenantctl exchange rules disable jane@contoso.com "Move invoices"'],
    args: [USER_ARG, RULE_ARG],
    risk: 'write',
    permissions: {
      delegated: ['MailboxSettings.ReadWrite', 'Mail.ReadBasic', RESOLVE_PERMISSION],
      application: ['MailboxSettings.ReadWrite', 'Mail.ReadBasic.All', RESOLVE_PERMISSION],
    },
    async preview(ctx, args) {
      const { user, shaped } = await findRule(ctx, args.user, args.rule);
      return {
        ...ruleTarget(user, shaped),
        action: `Disable inbox rule "${shaped.name}" for ${user.userPrincipalName}`,
      };
    },
    async run(ctx, args) {
      const { user, shaped } = await findRule(ctx, args.user, args.rule);
      await ctx.graph.patch(`/users/${seg(user.id)}/mailFolders/inbox/messageRules/${seg(shaped.id)}`, {
        isEnabled: false,
      });
      return { disabled: true, ...ruleTarget(user, { ...shaped, enabled: false }) };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'rules audit',
    summary: 'Scan many mailboxes for forwarding or suspicious inbox rules',
    description:
      'Tenant-wide BEC sweep. Scans the given users, a group, or (default) licensed member users up to --max-users, ' +
      'batched 20 mailboxes per Graph call, and returns only mailboxes whose rules forward mail or look suspicious. `errors` ' +
      'lists mailboxes that could not be read (no Exchange mailbox, no access). `usersTruncated: true` means more ' +
      'users exist than --max-users. Needs app-only auth to read other mailboxes. Follow up with ' +
      '`tenantctl exchange rules list <user>`.',
    examples: [
      'tenantctl exchange rules audit',
      'tenantctl exchange rules audit --group "Finance Team"',
      'tenantctl exchange rules audit --users jane@contoso.com,bob@contoso.com',
    ],
    args: [
      { name: 'users', type: 'string[]', description: 'Comma-separated users to scan.' },
      { name: 'group', type: 'string', description: 'Scan members of this group (nested members included).' },
      { name: 'maxUsers', type: 'number', default: 200, description: 'Maximum mailboxes to scan.' },
    ],
    risk: 'read',
    permissions: {
      delegated: [...RULE_READ.delegated, 'GroupMember.Read.All'],
      application: [...RULE_READ.application, 'GroupMember.Read.All'],
    },
    async run(ctx, args) {
      if (args.users?.length && args.group) throw new CliError('USAGE', 'Pass --users or --group, not both');
      const [targets, { domains, note }] = await Promise.all([auditTargets(ctx, args), tryTenantDomains(ctx)]);
      const flagged: { user: string; displayName: string; rules: ShapedInboxRule[] }[] = [];
      const errors = [...targets.errors];
      const results = await getInboxRulesBatch(
        ctx,
        targets.users.map((u) => u.id),
        domains,
      );
      targets.users.forEach((user, index) => {
        const result = results[index];
        if (result instanceof CliError || !result) {
          errors.push({
            user: user.userPrincipalName,
            code: result ? (result.graphCode ?? result.code) : 'INTERNAL',
            message: result ? result.message : 'No response',
          });
          return;
        }
        const rules = result.filter((r) => r.forwards || r.suspicious);
        if (rules.length) flagged.push({ user: user.userPrincipalName, displayName: user.displayName, rules });
      });
      flagged.sort((a, b) => Number(b.rules.some((r) => r.suspicious)) - Number(a.rules.some((r) => r.suspicious)));
      return {
        scanned: targets.users.length,
        usersTruncated: targets.truncated,
        flaggedCount: flagged.length,
        suspiciousCount: flagged.filter((m) => m.rules.some((r) => r.suspicious)).length,
        mailboxes: flagged,
        errors,
        ...(note ? { note } : {}),
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'forwarding',
    summary: 'Show every mail forwarding path of a mailbox: inbox rules and mailbox-level forwarding',
    description:
      'Answers "is this mailbox forwarding mail anywhere?". Lists inbox rules that forward or redirect (Graph), and ' +
      'mailbox-level forwarding (ForwardingSmtpAddress / ForwardingAddress, set by admins or in OWA settings) read ' +
      'through the Exchange Online PowerShell bridge. forwardsExternally covers both (null when tenant domains are ' +
      'unreadable). When the bridge is unavailable (no pwsh, not signed in to Exchange, no Exchange role) ' +
      '`mailboxForwarding` is `{ error, hint }` and the inbox-rule result still stands. Change mailbox forwarding with ' +
      `\`tenantctl exchange forwarding set|clear\`. ${MAILBOX_ACCESS_NOTE}`,
    examples: ['tenantctl exchange forwarding jane@contoso.com'],
    args: [USER_ARG],
    risk: 'read',
    permissions: {
      delegated: RULE_READ.delegated,
      application: RULE_READ.application,
      exchange: true,
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const [{ domains, note }, mailboxResult] = await Promise.all([
        tryTenantDomains(ctx),
        getMailboxForwarding(ctx, user.userPrincipalName).catch((error: unknown) => error as Error),
      ]);
      const rules = (await getInboxRules(ctx, user.id, domains)).filter((r) => r.forwards);
      const rulesExternal = rules.some((r) => r.enabled && r.forwardsExternally);
      let mailboxForwarding: Record<string, unknown>;
      let mailboxExternal: boolean | null = false;
      if (mailboxResult instanceof Error) {
        mailboxForwarding = {
          error: mailboxResult.message,
          hint:
            (mailboxResult instanceof CliError && mailboxResult.hint) ||
            'Mailbox-level forwarding needs Exchange Online PowerShell (pwsh + ExchangeOnlineManagement) and an ' +
              'Exchange Administrator role; inbox-rule results above are still valid.',
        };
        mailboxExternal = null;
      } else {
        const smtp = mailboxResult.forwardingSmtpAddress;
        mailboxExternal = smtp ? (domains ? !isInternal(smtp, domains) : null) : false;
        mailboxForwarding = {
          ...mailboxResult,
          forwards: Boolean(smtp || mailboxResult.forwardingAddress),
          forwardsExternally: mailboxExternal,
        };
      }
      const forwardsExternally =
        rulesExternal || mailboxExternal === true ? true : domains && mailboxExternal !== null ? false : null;
      return {
        user: user.userPrincipalName,
        forwardsViaInboxRules: rules.some((r) => r.enabled),
        forwardsViaMailbox: mailboxResult instanceof Error ? null : mailboxForwarding.forwards,
        forwardsExternally,
        inboxRules: rules.map((r) => ({
          id: r.id,
          name: r.name,
          enabled: r.enabled,
          targets: r.forwardTargets,
          externalTargets: r.externalTargets,
          conditions: r.conditions,
          suspicious: r.suspicious,
          reason: r.reason,
        })),
        mailboxForwarding,
        ...(note ? { note } : {}),
      };
    },
  }),
];
