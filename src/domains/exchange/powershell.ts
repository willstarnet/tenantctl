import { CliError } from '../../core/errors.ts';
import { type ExoStep, exoErrorCode, exoRun } from '../../core/exo.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { bytesToGb, compact, daysAgo } from '../../core/util.ts';
import {
  getMailboxForwarding,
  isInternal,
  type MailboxForwarding,
  shapeMailboxForwarding,
  tryTenantDomains,
} from './rules.ts';

/*
 * Exchange admin features Graph lacks, through the Exchange Online PowerShell bridge (`ctx.exo`). Identities go
 * straight to `-Identity`: Exchange resolves UPN, email, alias and display name itself and reports not-found or
 * ambiguous matches, which `exo()` maps to NOT_FOUND / AMBIGUOUS with a next step.
 */

const EXO_PERMISSIONS = {
  exchange: true,
  roles: ['Exchange Administrator'],
};

/** Quarantine actions for all users: Security Administrator (Entra) or the Quarantine Administrator role group. */
const QUARANTINE_PERMISSIONS = { ...EXO_PERMISSIONS, roles: ['Security Administrator'] };

const BRIDGE_NOTE =
  'Runs through Exchange Online PowerShell: needs pwsh 7 and the ExchangeOnlineManagement module (3.1+).';

const MAILBOX_ARG = {
  name: 'mailbox',
  type: 'string',
  required: true,
  positional: true,
  description: 'Mailbox: UPN, email, alias or display name.',
} as const;

const IDENTITY_HINT =
  'Use the email address or alias; `m365 entra users list --search <name>` or `m365 exchange dl list --search` finds it.';

/** Re-raise bridge errors with the code and hint an agent can act on. */
function mapExoError(error: unknown): never {
  if (!(error instanceof CliError)) throw error;
  if (/matches multiple entries|multiple (objects|recipients) match/i.test(error.message)) {
    throw new CliError('AMBIGUOUS', error.message, { hint: `Several recipients match. ${IDENTITY_HINT}` });
  }
  if (error.code === 'NOT_FOUND' && !error.hint)
    throw new CliError('NOT_FOUND', error.message, { hint: IDENTITY_HINT });
  throw error;
}

/** `exoRun` with AMBIGUOUS / NOT_FOUND hints. One call = one PowerShell session; batch related steps. */
async function exo(ctx: Context, steps: ExoStep[]): Promise<Record<string, unknown>[][]> {
  try {
    return await exoRun(ctx.exo, steps);
  } catch (error) {
    mapExoError(error);
  }
}

function first(rows: Record<string, unknown>[] | undefined, what: string): Record<string, unknown> {
  const row = rows?.[0];
  if (!row) throw new CliError('NOT_FOUND', `No ${what} found`, { hint: IDENTITY_HINT });
  return row;
}

function text(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  const s = String(value).trim();
  return s ? s : null;
}

/** Exchange multi-valued properties arrive as arrays of strings, a single string, or null. */
function strings(value: unknown): string[] {
  if (Array.isArray(value)) return value.map(String).filter(Boolean);
  const s = text(value);
  return s ? [s] : [];
}

/** `"1.234 GB (1,325,000,000 bytes)"` -> 1325000000. */
export function parseExoBytes(value: unknown): number | null {
  const match = String(value ?? '').match(/\(([\d,]+) bytes\)/);
  return match ? Number((match[1] as string).replace(/,/g, '')) : null;
}

/** Truncation for cmdlets that take ResultSize/PageSize: ask for one more than wanted. */
function sizeFor(limit: number): number | 'Unlimited' {
  return Number.isFinite(limit) ? limit + 1 : 'Unlimited';
}

function toList<T>(items: T[], limit: number): ListResult<T> {
  return { items: items.slice(0, limit), truncated: items.length > limit };
}

// ---------------------------------------------------------------------------------------------- forwarding

async function forwardingState(ctx: Context, mailbox: string) {
  const [current, { domains }] = await Promise.all([
    getMailboxForwarding(ctx, mailbox).catch(mapExoError),
    tryTenantDomains(ctx),
  ]);
  const external = (address: string | null) => (address ? (domains ? !isInternal(address, domains) : null) : false);
  return { current, domains, external };
}

function describeForwarding(f: MailboxForwarding): string {
  const target = f.forwardingSmtpAddress ?? f.forwardingAddress;
  if (!target) return 'not forwarding';
  return `forwarding to ${target}${f.deliverToMailboxAndForward ? ' (keeps a copy)' : ' (no copy kept)'}`;
}

// ---------------------------------------------------------------------------------------------- permissions

export type MailboxAccess = 'FullAccess' | 'SendAs' | 'SendOnBehalf';

export interface MailboxPermissionEntry {
  user: string;
  access: MailboxAccess;
  inherited?: boolean;
}

const SELF = /^NT AUTHORITY\\SELF$/i;

/** Flatten Get-MailboxPermission / Get-RecipientPermission / GrantSendOnBehalfTo into one list. */
export function flattenMailboxPermissions(
  fullAccess: Record<string, unknown>[],
  sendAs: Record<string, unknown>[],
  mailbox: Record<string, unknown> | undefined,
  includeInherited = false,
): MailboxPermissionEntry[] {
  const out: MailboxPermissionEntry[] = [];
  for (const row of fullAccess) {
    const user = text(row.User);
    if (!user || SELF.test(user) || row.Deny === true || !strings(row.AccessRights).some((r) => /FullAccess/i.test(r)))
      continue;
    if (row.IsInherited === true && !includeInherited) continue;
    out.push({ user, access: 'FullAccess', ...(row.IsInherited === true ? { inherited: true } : {}) });
  }
  for (const row of sendAs) {
    const user = text(row.Trustee);
    if (!user || SELF.test(user) || row.AccessControlType === 'Deny') continue;
    if (!strings(row.AccessRights).some((r) => /SendAs/i.test(r))) continue;
    if (row.IsInherited === true && !includeInherited) continue;
    out.push({ user, access: 'SendAs', ...(row.IsInherited === true ? { inherited: true } : {}) });
  }
  for (const user of strings(mailbox?.GrantSendOnBehalfTo)) out.push({ user, access: 'SendOnBehalf' });
  return out;
}

async function loadPermissions(ctx: Context, mailbox: string, includeInherited = false) {
  const [fullAccess, sendAs, mailboxes] = await exo(ctx, [
    {
      cmdlet: 'Get-MailboxPermission',
      params: { Identity: mailbox },
      select: ['User', 'AccessRights', 'IsInherited', 'Deny'],
    },
    {
      cmdlet: 'Get-RecipientPermission',
      params: { Identity: mailbox },
      select: ['Trustee', 'AccessRights', 'IsInherited', 'AccessControlType'],
    },
    { cmdlet: 'Get-Mailbox', params: { Identity: mailbox }, select: ['PrimarySmtpAddress', 'GrantSendOnBehalfTo'] },
  ]);
  const box = first(mailboxes, `mailbox "${mailbox}"`);
  return {
    mailbox: text(box.PrimarySmtpAddress) ?? mailbox,
    permissions: flattenMailboxPermissions(fullAccess ?? [], sendAs ?? [], box, includeInherited),
  };
}

const ACCESS_CHOICES = ['full', 'sendas', 'sendonbehalf'] as const;
type AccessChoice = (typeof ACCESS_CHOICES)[number];
const ACCESS_NAME: Record<AccessChoice, MailboxAccess> = {
  full: 'FullAccess',
  sendas: 'SendAs',
  sendonbehalf: 'SendOnBehalf',
};

function permissionStep(mailbox: string, user: string, access: AccessChoice, add: boolean, automap = true): ExoStep {
  if (access === 'full') {
    return add
      ? {
          cmdlet: 'Add-MailboxPermission',
          params: {
            Identity: mailbox,
            User: user,
            AccessRights: 'FullAccess',
            InheritanceType: 'All',
            AutoMapping: automap,
          },
          select: ['Identity', 'User', 'AccessRights'],
        }
      : {
          cmdlet: 'Remove-MailboxPermission',
          params: { Identity: mailbox, User: user, AccessRights: 'FullAccess', InheritanceType: 'All', Confirm: false },
        };
  }
  if (access === 'sendas') {
    return {
      cmdlet: add ? 'Add-RecipientPermission' : 'Remove-RecipientPermission',
      params: { Identity: mailbox, Trustee: user, AccessRights: 'SendAs', Confirm: false },
      ...(add ? { select: ['Identity', 'Trustee', 'AccessRights'] } : {}),
    };
  }
  return {
    cmdlet: 'Set-Mailbox',
    params: { Identity: mailbox, GrantSendOnBehalfTo: add ? { Add: user } : { Remove: user } },
  };
}

/** Exchange shows trustees by display name or UPN; match either way, case-insensitively. */
function holds(list: MailboxPermissionEntry[], user: string, access: MailboxAccess): boolean {
  const wanted = user.trim().toLowerCase();
  return list.some((p) => p.access === access && p.user.toLowerCase() === wanted);
}

const PERMISSION_ARGS = [
  MAILBOX_ARG,
  {
    name: 'user',
    type: 'string',
    required: true,
    positional: true,
    description: 'Who gets (or loses) access: UPN, email, alias or display name.',
  },
  {
    name: 'access',
    type: 'string',
    required: true,
    choices: ACCESS_CHOICES,
    description: 'full, sendas or sendonbehalf.',
  },
] as const;

// ---------------------------------------------------------------------------------------------- mailbox

const DETAIL_SELECT = [
  'DisplayName',
  'RecipientTypeDetails',
  'PrimarySmtpAddress',
  'EmailAddresses',
  'HiddenFromAddressListsEnabled',
  'ForwardingSmtpAddress',
  'ForwardingAddress',
  'DeliverToMailboxAndForward',
  'LitigationHoldEnabled',
  'ArchiveStatus',
  'ProhibitSendQuota',
];

/** Secondary SMTP addresses (`smtp:` lowercase prefix); the primary is `SMTP:`. */
export function aliasesOf(emailAddresses: unknown): string[] {
  return strings(emailAddresses)
    .filter((a) => a.startsWith('smtp:'))
    .map((a) => a.slice(5));
}

const HOLD_SELECT = [
  'PrimarySmtpAddress',
  'RecipientTypeDetails',
  'LitigationHoldEnabled',
  'LitigationHoldDuration',
  'LitigationHoldOwner',
  'LitigationHoldDate',
  'RetentionComment',
  'ArchiveStatus',
];

function holdState(row: Record<string, unknown>) {
  return {
    mailbox: text(row.PrimarySmtpAddress),
    litigationHold: row.LitigationHoldEnabled === true,
    holdDuration: text(row.LitigationHoldDuration),
    holdOwner: text(row.LitigationHoldOwner),
    holdSince: text(row.LitigationHoldDate),
    retentionComment: text(row.RetentionComment),
    archiveStatus: text(row.ArchiveStatus),
  };
}

async function readMailboxForConvert(ctx: Context, mailbox: string) {
  const results = await ctx.exo
    .run([
      {
        cmdlet: 'Get-Mailbox',
        params: { Identity: mailbox },
        select: [
          'PrimarySmtpAddress',
          'RecipientTypeDetails',
          'ArchiveStatus',
          'LitigationHoldEnabled',
          'InPlaceHolds',
        ],
      },
      {
        cmdlet: 'Get-MailboxStatistics',
        params: { Identity: mailbox },
        select: ['TotalItemSize'],
        continueOnError: true,
      },
    ])
    .catch(mapExoError);
  const [boxResult, statsResult] = results;
  if (!boxResult?.ok) {
    mapExoError(new CliError(exoErrorCode(boxResult?.error ?? ''), `Get-Mailbox: ${boxResult?.error ?? 'failed'}`));
  }
  const box = first(boxResult.data, `mailbox "${mailbox}"`);
  const bytes = statsResult?.ok ? parseExoBytes(statsResult.data[0]?.TotalItemSize) : null;
  return {
    mailbox: text(box.PrimarySmtpAddress) ?? mailbox,
    type: text(box.RecipientTypeDetails),
    sizeGb: bytesToGb(bytes),
    archiveEnabled: !/^none$/i.test(text(box.ArchiveStatus) ?? 'None'),
    litigationHold: box.LitigationHoldEnabled === true,
    inPlaceHolds: strings(box.InPlaceHolds).length,
  };
}

// ---------------------------------------------------------------------------------------------- quarantine

const QUARANTINE_TYPES = [
  'Bulk',
  'DataLossPrevention',
  'FileTypeBlock',
  'HighConfPhish',
  'Malware',
  'Phish',
  'Spam',
  'SPOMalware',
  'TransportRule',
] as const;

const QUARANTINE_SELECT = [
  'Identity',
  'ReceivedTime',
  'SenderAddress',
  'RecipientAddress',
  'Subject',
  'QuarantineTypes',
  'Type',
  'ReleaseStatus',
  'Expires',
  'Direction',
];

export function shapeQuarantineMessage(row: Record<string, unknown>) {
  return {
    identity: text(row.Identity),
    received: text(row.ReceivedTime),
    sender: text(row.SenderAddress),
    recipients: strings(row.RecipientAddress),
    subject: text(row.Subject),
    type: text(row.QuarantineTypes) ?? text(row.Type),
    releaseStatus: text(row.ReleaseStatus),
    expires: text(row.Expires),
    direction: text(row.Direction),
  };
}

async function quarantineMessage(ctx: Context, identity: string) {
  const [rows] = await exo(ctx, [
    { cmdlet: 'Get-QuarantineMessage', params: { Identity: identity }, select: QUARANTINE_SELECT },
  ]);
  return shapeQuarantineMessage(first(rows, `quarantined message "${identity}"`));
}

const QUARANTINE_ID_ARG = {
  name: 'identity',
  type: 'string',
  required: true,
  positional: true,
  description: 'Quarantine Identity (GUID1\\GUID2) from `exchange quarantine list`.',
} as const;

// ---------------------------------------------------------------------------------------------- distribution lists

const DL_SELECT = [
  'Name',
  'DisplayName',
  'PrimarySmtpAddress',
  'RecipientTypeDetails',
  'ManagedBy',
  'MemberJoinRestriction',
];
const MEMBER_SELECT = ['Name', 'DisplayName', 'PrimarySmtpAddress', 'RecipientTypeDetails'];

function shapeRecipient(row: Record<string, unknown>) {
  return {
    name: text(row.DisplayName) ?? text(row.Name),
    email: text(row.PrimarySmtpAddress),
    type: text(row.RecipientTypeDetails),
  };
}

const GROUP_ARG = {
  name: 'group',
  type: 'string',
  required: true,
  positional: true,
  description: 'Distribution list or mail-enabled security group: email, alias or name.',
} as const;

/** Per-item outcome of a step run with continueOnError. */
function stepOutcome(result: { ok: boolean; error?: string } | undefined, skipIf?: RegExp) {
  if (result?.ok) return { status: 'done' as const };
  const error = result?.error ?? 'did not run';
  if (skipIf?.test(error)) return { status: 'skipped' as const, reason: error };
  return { status: 'failed' as const, error };
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'exchange',
    name: 'forwarding set',
    summary: 'Forward all mail of a mailbox to an address (mailbox-level forwarding)',
    description:
      "Sets ForwardingSmtpAddress so every message to the mailbox is forwarded, e.g. a leaver's mail to their " +
      'manager. Without --keep-copy the mailbox keeps no copy. Destructive because mail leaves the mailbox; the ' +
      'preview shows current forwarding and whether the target is external (external forwarding may also be ' +
      `blocked by the outbound spam policy). ${BRIDGE_NOTE}`,
    examples: ['m365 exchange forwarding set jane@contoso.com --to bob@contoso.com --keep-copy'],
    args: [
      MAILBOX_ARG,
      { name: 'to', type: 'string', required: true, description: 'SMTP address to forward to.' },
      { name: 'keepCopy', type: 'boolean', description: 'Also deliver to the mailbox (DeliverToMailboxAndForward).' },
    ],
    risk: 'destructive',
    permissions: { ...EXO_PERMISSIONS, delegated: ['Domain.Read.All'], application: ['Domain.Read.All'] },
    async preview(ctx, args) {
      if (!args.to.includes('@')) throw new CliError('USAGE', '--to must be an SMTP address');
      const { current, external } = await forwardingState(ctx, args.mailbox);
      return {
        mailbox: args.mailbox,
        current: { ...current, forwardsExternally: external(current.forwardingSmtpAddress) },
        target: args.to,
        targetIsExternal: external(args.to),
        keepCopy: args.keepCopy,
        action:
          `Forward all mail for ${args.mailbox} to ${args.to}${external(args.to) ? ' (EXTERNAL)' : ''}, ` +
          `${args.keepCopy ? 'keeping a copy in the mailbox' : 'keeping no copy'} (currently ${describeForwarding(current)})`,
      };
    },
    async run(ctx, args) {
      if (!args.to.includes('@')) throw new CliError('USAGE', '--to must be an SMTP address');
      await exo(ctx, [
        {
          cmdlet: 'Set-Mailbox',
          params: { Identity: args.mailbox, ForwardingSmtpAddress: args.to, DeliverToMailboxAndForward: args.keepCopy },
        },
      ]);
      return { mailbox: args.mailbox, forwardingSmtpAddress: args.to, deliverToMailboxAndForward: args.keepCopy };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'forwarding clear',
    summary: 'Remove mailbox-level forwarding',
    description:
      'Clears ForwardingSmtpAddress and ForwardingAddress and turns DeliverToMailboxAndForward off. Inbox-rule ' +
      `forwarding is separate: see \`m365 exchange forwarding <mailbox>\` and \`exchange rules delete\`. ${BRIDGE_NOTE}`,
    examples: ['m365 exchange forwarding clear jane@contoso.com'],
    args: [MAILBOX_ARG],
    risk: 'destructive',
    permissions: EXO_PERMISSIONS,
    async preview(ctx, args) {
      const { current, external } = await forwardingState(ctx, args.mailbox);
      return {
        mailbox: args.mailbox,
        current: { ...current, forwardsExternally: external(current.forwardingSmtpAddress) },
        action: `Stop mailbox-level forwarding for ${args.mailbox} (currently ${describeForwarding(current)})`,
      };
    },
    async run(ctx, args) {
      await exo(ctx, [
        {
          cmdlet: 'Set-Mailbox',
          params: {
            Identity: args.mailbox,
            ForwardingSmtpAddress: null,
            ForwardingAddress: null,
            DeliverToMailboxAndForward: false,
          },
        },
      ]);
      return { mailbox: args.mailbox, cleared: true };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'permissions list',
    summary: 'Who has FullAccess, Send As or Send on Behalf on a mailbox',
    description:
      'One flat list `{ user, access }` from Get-MailboxPermission (FullAccess), Get-RecipientPermission (SendAs) ' +
      'and GrantSendOnBehalfTo. The mailbox owner (NT AUTHORITY\\SELF) and inherited/system entries are left out ' +
      `unless --include-inherited. ${BRIDGE_NOTE}`,
    examples: ['m365 exchange permissions list shared-finance@contoso.com'],
    args: [
      MAILBOX_ARG,
      { name: 'includeInherited', type: 'boolean', description: 'Also show inherited permissions (inherited: true).' },
    ],
    risk: 'read',
    permissions: EXO_PERMISSIONS,
    async run(ctx, args) {
      const { permissions } = await loadPermissions(ctx, args.mailbox, args.includeInherited);
      return { items: permissions, truncated: false };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'permissions add',
    summary: 'Grant FullAccess, Send As or Send on Behalf on a mailbox',
    description:
      'full = open the mailbox (auto-mapped into Outlook unless --no-automap); sendas = send as the mailbox; ' +
      `sendonbehalf = send "on behalf of". Destructive because it opens a mailbox to someone. ${BRIDGE_NOTE}`,
    examples: [
      'm365 exchange permissions add shared-finance@contoso.com bob@contoso.com --access full',
      'm365 exchange permissions add jane@contoso.com bob@contoso.com --access sendas',
    ],
    args: [
      ...PERMISSION_ARGS,
      {
        name: 'automap',
        type: 'boolean',
        default: true,
        description: 'FullAccess only: add to Outlook automatically (--no-automap to skip).',
      },
    ],
    risk: 'destructive',
    permissions: EXO_PERMISSIONS,
    async preview(ctx, args) {
      const { mailbox, permissions } = await loadPermissions(ctx, args.mailbox);
      const access = ACCESS_NAME[args.access];
      return {
        mailbox,
        user: args.user,
        access,
        alreadyGranted: holds(permissions, args.user, access),
        current: permissions,
        action: `Give ${args.user} ${access} on ${mailbox}${args.access === 'full' && !args.automap ? ' (no auto-mapping)' : ''}`,
      };
    },
    async run(ctx, args) {
      await exo(ctx, [permissionStep(args.mailbox, args.user, args.access, true, args.automap)]);
      return { mailbox: args.mailbox, user: args.user, access: ACCESS_NAME[args.access], granted: true };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'permissions remove',
    summary: 'Revoke FullAccess, Send As or Send on Behalf on a mailbox',
    description: `Removes the permission; Outlook drops an auto-mapped mailbox within a few hours. ${BRIDGE_NOTE}`,
    examples: ['m365 exchange permissions remove shared-finance@contoso.com bob@contoso.com --access full'],
    args: PERMISSION_ARGS,
    risk: 'destructive',
    permissions: EXO_PERMISSIONS,
    async preview(ctx, args) {
      const { mailbox, permissions } = await loadPermissions(ctx, args.mailbox);
      const access = ACCESS_NAME[args.access];
      return {
        mailbox,
        user: args.user,
        access,
        currentlyGranted: holds(permissions, args.user, access),
        current: permissions,
        action: `Remove ${access} for ${args.user} on ${mailbox}`,
      };
    },
    async run(ctx, args) {
      await exo(ctx, [permissionStep(args.mailbox, args.user, args.access, false)]);
      return { mailbox: args.mailbox, user: args.user, access: ACCESS_NAME[args.access], removed: true };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'shared-mailbox create',
    summary: 'Create a shared mailbox and give members FullAccess + Send As',
    description:
      'New-Mailbox -Shared, then FullAccess (auto-mapped) and SendAs for each member, in one session. Shared ' +
      'mailboxes need no license up to 50 GB without archive or hold. A member grant can fail right after creation ' +
      `while Exchange replicates; retry it with \`exchange permissions add\`. ${BRIDGE_NOTE}`,
    examples: [
      'm365 exchange shared-mailbox create --name "Finance" --email finance@contoso.com --members jane@contoso.com,bob@contoso.com',
    ],
    args: [
      { name: 'name', type: 'string', required: true, description: 'Display name (also the unique Name).' },
      { name: 'email', type: 'string', required: true, description: 'Primary SMTP address.' },
      { name: 'members', type: 'string[]', description: 'Users who get FullAccess and SendAs.' },
    ],
    risk: 'write',
    permissions: EXO_PERMISSIONS,
    async preview(_ctx, args) {
      return {
        name: args.name,
        email: args.email,
        members: args.members ?? [],
        action: `Create shared mailbox ${args.email} ("${args.name}")${args.members?.length ? ` with FullAccess + SendAs for ${args.members.join(', ')}` : ''}`,
      };
    },
    async run(ctx, args) {
      const members = args.members ?? [];
      const steps: ExoStep[] = [
        {
          cmdlet: 'New-Mailbox',
          params: { Shared: true, Name: args.name, DisplayName: args.name, PrimarySmtpAddress: args.email },
          select: ['PrimarySmtpAddress', 'Alias', 'RecipientTypeDetails'],
        },
        ...members.flatMap((member): ExoStep[] => [
          { ...permissionStep(args.email, member, 'full', true), continueOnError: true },
          { ...permissionStep(args.email, member, 'sendas', true), continueOnError: true },
        ]),
      ];
      const results = await ctx.exo.run(steps).catch(mapExoError);
      const created = results[0];
      if (!created?.ok) {
        mapExoError(new CliError(exoErrorCode(created?.error ?? ''), `New-Mailbox: ${created?.error ?? 'failed'}`));
      }
      const box = created.data[0] ?? {};
      return {
        mailbox: text(box.PrimarySmtpAddress) ?? args.email,
        alias: text(box.Alias),
        type: text(box.RecipientTypeDetails),
        members: members.map((member, index) => ({
          member,
          fullAccess: stepOutcome(results[1 + index * 2]),
          sendAs: stepOutcome(results[2 + index * 2]),
        })),
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'mailbox convert',
    summary: 'Convert a mailbox between user (regular) and shared',
    description:
      'Typical leaver flow: convert to shared, grant access, then remove the license. A shared mailbox over 50 GB, ' +
      'with an archive, or on hold still needs a license; the preview shows size, archive and holds. Converting ' +
      `to regular needs a license on the account before it can be used. ${BRIDGE_NOTE}`,
    examples: ['m365 exchange mailbox convert jane@contoso.com --to shared'],
    args: [
      MAILBOX_ARG,
      { name: 'to', type: 'string', required: true, choices: ['shared', 'regular'], description: 'Target type.' },
    ],
    risk: 'destructive',
    permissions: EXO_PERMISSIONS,
    async preview(ctx, args) {
      const state = await readMailboxForConvert(ctx, args.mailbox);
      const target = args.to === 'shared' ? 'SharedMailbox' : 'UserMailbox';
      const needsLicense =
        args.to === 'shared' &&
        ((state.sizeGb ?? 0) > 50 || state.archiveEnabled || state.litigationHold || state.inPlaceHolds > 0);
      return {
        ...state,
        target,
        alreadyThatType: state.type === target,
        ...(args.to === 'shared'
          ? {
              licenseNote: needsLicense
                ? 'Keep a license: the mailbox is over 50 GB, has an archive, or is on hold.'
                : 'No license needed after conversion (under 50 GB, no archive, no hold).',
            }
          : { licenseNote: 'Assign an Exchange Online license before the user signs in.' }),
        action: `Convert ${state.mailbox} from ${state.type ?? 'unknown'} to ${target}`,
      };
    },
    async run(ctx, args) {
      await exo(ctx, [
        {
          cmdlet: 'Set-Mailbox',
          params: { Identity: args.mailbox, Type: args.to === 'shared' ? 'Shared' : 'Regular' },
        },
      ]);
      return { mailbox: args.mailbox, convertedTo: args.to };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'mailbox details',
    summary: 'Exchange-side mailbox facts: type, aliases, GAL visibility, forwarding, hold, archive, size, last logon',
    description:
      'Get-Mailbox + Get-MailboxStatistics in one session. Complements `m365 exchange mailbox` (Graph settings). ' +
      `\`statistics\` is { error } when the mailbox was never logged on to. ${BRIDGE_NOTE}`,
    examples: ['m365 exchange mailbox details jane@contoso.com'],
    args: [MAILBOX_ARG],
    risk: 'read',
    permissions: EXO_PERMISSIONS,
    async run(ctx, args) {
      const results = await ctx.exo
        .run([
          { cmdlet: 'Get-Mailbox', params: { Identity: args.mailbox }, select: DETAIL_SELECT },
          {
            cmdlet: 'Get-MailboxStatistics',
            params: { Identity: args.mailbox },
            select: ['TotalItemSize', 'ItemCount', 'LastLogonTime'],
            continueOnError: true,
          },
        ])
        .catch(mapExoError);
      const [boxResult, statsResult] = results;
      if (!boxResult?.ok) {
        mapExoError(new CliError(exoErrorCode(boxResult?.error ?? ''), `Get-Mailbox: ${boxResult?.error ?? 'failed'}`));
      }
      const box = first(boxResult.data, `mailbox "${args.mailbox}"`);
      const stats = statsResult?.ok ? statsResult.data[0] : undefined;
      if (ctx.full) return { mailbox: box, statistics: stats ?? { error: statsResult?.error ?? 'did not run' } };
      return {
        displayName: text(box.DisplayName),
        type: text(box.RecipientTypeDetails),
        primarySmtpAddress: text(box.PrimarySmtpAddress),
        aliases: aliasesOf(box.EmailAddresses),
        hiddenFromAddressLists: box.HiddenFromAddressListsEnabled === true,
        forwarding: shapeMailboxForwarding(box),
        litigationHold: box.LitigationHoldEnabled === true,
        archiveStatus: text(box.ArchiveStatus),
        prohibitSendQuota: text(box.ProhibitSendQuota),
        statistics: stats
          ? {
              totalSize: text(stats.TotalItemSize),
              totalSizeGb: bytesToGb(parseExoBytes(stats.TotalItemSize)),
              itemCount: typeof stats.ItemCount === 'number' ? stats.ItemCount : Number(stats.ItemCount ?? 0),
              lastLogon: text(stats.LastLogonTime),
            }
          : { error: statsResult?.error ?? 'did not run' },
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'dl list',
    summary: 'List distribution lists and mail-enabled security groups',
    description: `--search matches name, alias and email (ambiguous name resolution). ${BRIDGE_NOTE}`,
    examples: ['m365 exchange dl list --search finance'],
    args: [{ name: 'search', type: 'string', description: 'Name, alias or email prefix.' }, ...LIST_ARGS],
    risk: 'read',
    permissions: EXO_PERMISSIONS,
    async run(ctx, args) {
      const limit = listLimit(args);
      const [rows] = await exo(ctx, [
        {
          cmdlet: 'Get-DistributionGroup',
          params: compact({ ResultSize: sizeFor(limit), Anr: args.search }),
          select: DL_SELECT,
        },
      ]);
      const items = (rows ?? []).map((row) =>
        ctx.full
          ? row
          : {
              ...shapeRecipient(row),
              managedBy: strings(row.ManagedBy),
              joinRestriction: text(row.MemberJoinRestriction),
            },
      );
      return toList(items, limit);
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'dl members',
    summary: 'List members of a distribution list',
    description: BRIDGE_NOTE,
    examples: ['m365 exchange dl members finance@contoso.com'],
    args: [GROUP_ARG, ...LIST_ARGS],
    risk: 'read',
    permissions: EXO_PERMISSIONS,
    async run(ctx, args) {
      const limit = listLimit(args);
      const [rows] = await exo(ctx, [
        {
          cmdlet: 'Get-DistributionGroupMember',
          params: { Identity: args.group, ResultSize: sizeFor(limit) },
          select: MEMBER_SELECT,
        },
      ]);
      return toList(
        (rows ?? []).map((row) => (ctx.full ? row : shapeRecipient(row))),
        limit,
      );
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'dl add-member',
    summary: 'Add people to a distribution list',
    description:
      'Adds each member in one session and reports per member (`skipped` when already a member). Works even when ' +
      `you are not a group owner (BypassSecurityGroupManagerCheck). ${BRIDGE_NOTE}`,
    examples: ['m365 exchange dl add-member finance@contoso.com jane@contoso.com bob@contoso.com'],
    args: [
      GROUP_ARG,
      {
        name: 'members',
        type: 'string[]',
        required: true,
        positional: true,
        description: 'Members to add (email, alias or name).',
      },
    ],
    risk: 'write',
    permissions: EXO_PERMISSIONS,
    async preview(_ctx, args) {
      return { group: args.group, members: args.members, action: `Add ${args.members.join(', ')} to ${args.group}` };
    },
    async run(ctx, args) {
      const results = await ctx.exo
        .run(
          args.members.map((member) => ({
            cmdlet: 'Add-DistributionGroupMember',
            params: { Identity: args.group, Member: member, BypassSecurityGroupManagerCheck: true, Confirm: false },
            continueOnError: true,
          })),
        )
        .catch(mapExoError);
      return {
        group: args.group,
        members: args.members.map((member, index) => ({
          member,
          ...stepOutcome(results[index], /already a member/i),
        })),
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'dl remove-member',
    summary: 'Remove someone from a distribution list',
    description: BRIDGE_NOTE,
    examples: ['m365 exchange dl remove-member finance@contoso.com bob@contoso.com'],
    args: [
      GROUP_ARG,
      { name: 'member', type: 'string', required: true, positional: true, description: 'Member to remove.' },
    ],
    risk: 'destructive',
    permissions: EXO_PERMISSIONS,
    async preview(ctx, args) {
      const [groups, members] = await exo(ctx, [
        { cmdlet: 'Get-DistributionGroup', params: { Identity: args.group }, select: DL_SELECT },
        {
          cmdlet: 'Get-DistributionGroupMember',
          params: { Identity: args.group, ResultSize: 'Unlimited' },
          select: MEMBER_SELECT,
        },
      ]);
      const group = shapeRecipient(first(groups, `distribution list "${args.group}"`));
      const wanted = args.member.trim().toLowerCase();
      const shaped = (members ?? []).map(shapeRecipient);
      const match = shaped.find((m) => m.email?.toLowerCase() === wanted || m.name?.toLowerCase() === wanted);
      return {
        group,
        member: match ?? args.member,
        isMember: match ? true : null,
        memberCount: shaped.length,
        action: `Remove ${match?.email ?? args.member} from ${group.email ?? args.group}`,
      };
    },
    async run(ctx, args) {
      await exo(ctx, [
        {
          cmdlet: 'Remove-DistributionGroupMember',
          params: { Identity: args.group, Member: args.member, BypassSecurityGroupManagerCheck: true, Confirm: false },
        },
      ]);
      return { group: args.group, member: args.member, removed: true };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'dl create',
    summary: 'Create a distribution list',
    description: BRIDGE_NOTE,
    examples: [
      'm365 exchange dl create --name "Finance Team" --email finance@contoso.com --members jane@contoso.com,bob@contoso.com',
    ],
    args: [
      { name: 'name', type: 'string', required: true, description: 'Unique name (also the display name).' },
      { name: 'email', type: 'string', required: true, description: 'Primary SMTP address.' },
      { name: 'members', type: 'string[]', description: 'Initial members.' },
    ],
    risk: 'write',
    permissions: EXO_PERMISSIONS,
    async preview(_ctx, args) {
      return {
        name: args.name,
        email: args.email,
        members: args.members ?? [],
        action: `Create distribution list ${args.email} ("${args.name}")`,
      };
    },
    async run(ctx, args) {
      const [rows] = await exo(ctx, [
        {
          cmdlet: 'New-DistributionGroup',
          params: compact({
            Name: args.name,
            DisplayName: args.name,
            PrimarySmtpAddress: args.email,
            Members: args.members?.length ? args.members : undefined,
          }),
          select: DL_SELECT,
        },
      ]);
      return { ...shapeRecipient(first(rows, 'created group')), members: args.members ?? [] };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'quarantine list',
    summary: 'List quarantined messages (spam, phish, malware, policy)',
    description:
      'Use for "I never got mail from X". `identity` feeds `quarantine release|delete`. Admins see all users\' ' +
      'quarantine with Security Administrator or the Quarantine Administrator role group; Security Reader / Global ' +
      `Reader can view. ${BRIDGE_NOTE}`,
    examples: [
      'm365 exchange quarantine list --recipient jane@contoso.com',
      'm365 exchange quarantine list --sender billing@vendor.com --days 14 --type Phish',
    ],
    args: [
      { name: 'recipient', type: 'string', description: 'Recipient address.' },
      { name: 'sender', type: 'string', description: 'Sender address.' },
      {
        name: 'days',
        type: 'number',
        default: 7,
        description: 'Received in the last N days (quarantine keeps up to 30).',
      },
      { name: 'type', type: 'string', choices: QUARANTINE_TYPES, description: 'Why it was quarantined.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: QUARANTINE_PERMISSIONS,
    async run(ctx, args) {
      if (args.days <= 0) throw new CliError('USAGE', '--days must be positive');
      const limit = listLimit(args);
      const [rows] = await exo(ctx, [
        {
          cmdlet: 'Get-QuarantineMessage',
          params: compact({
            StartReceivedDate: daysAgo(args.days),
            EndReceivedDate: new Date().toISOString(),
            RecipientAddress: args.recipient ? [args.recipient] : undefined,
            SenderAddress: args.sender ? [args.sender] : undefined,
            QuarantineTypes: args.type ? [args.type] : undefined,
            PageSize: Math.min(Number.isFinite(limit) ? limit + 1 : 1000, 1000),
            Page: 1,
          }),
          select: QUARANTINE_SELECT,
        },
      ]);
      return toList(
        (rows ?? []).map((row) => (ctx.full ? row : shapeQuarantineMessage(row))),
        limit,
      );
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'quarantine release',
    summary: 'Release a quarantined message to its recipients',
    description:
      'Default and --to-all release to every original recipient; --user releases to specific addresses. Only ' +
      `release what you have checked is safe: phish/malware reach the inbox. ${BRIDGE_NOTE}`,
    examples: [
      'm365 exchange quarantine release "c14401cf-...\\4c2ca98e-..." --to-all',
      'm365 exchange quarantine release "c14401cf-...\\4c2ca98e-..." --user jane@contoso.com',
    ],
    args: [
      QUARANTINE_ID_ARG,
      { name: 'toAll', type: 'boolean', description: 'Release to all original recipients (default).' },
      { name: 'user', type: 'string[]', description: 'Release only to these addresses.' },
    ],
    risk: 'destructive',
    permissions: QUARANTINE_PERMISSIONS,
    async preview(ctx, args) {
      if (args.toAll && args.user?.length) throw new CliError('USAGE', 'Pass --to-all or --user, not both');
      const message = await quarantineMessage(ctx, args.identity);
      const to = args.user?.length ? args.user.join(', ') : 'all original recipients';
      return {
        message,
        releaseTo: args.user?.length ? args.user : 'all',
        action: `Release ${message.type ?? 'quarantined'} message "${message.subject ?? ''}" from ${message.sender ?? 'unknown sender'} to ${to}`,
      };
    },
    async run(ctx, args) {
      if (args.toAll && args.user?.length) throw new CliError('USAGE', 'Pass --to-all or --user, not both');
      await exo(ctx, [
        {
          cmdlet: 'Release-QuarantineMessage',
          params: args.user?.length
            ? { Identity: args.identity, User: args.user, Confirm: false }
            : { Identity: args.identity, ReleaseToAll: true, Confirm: false },
        },
      ]);
      return { identity: args.identity, released: true, to: args.user?.length ? args.user : 'all' };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'quarantine delete',
    summary: 'Delete a quarantined message',
    description: `Removes it from quarantine; it can no longer be released. ${BRIDGE_NOTE}`,
    examples: ['m365 exchange quarantine delete "c14401cf-...\\4c2ca98e-..."'],
    args: [QUARANTINE_ID_ARG],
    risk: 'destructive',
    permissions: QUARANTINE_PERMISSIONS,
    async preview(ctx, args) {
      const message = await quarantineMessage(ctx, args.identity);
      return {
        message,
        action: `Delete quarantined message "${message.subject ?? ''}" from ${message.sender ?? 'unknown sender'}`,
      };
    },
    async run(ctx, args) {
      await exo(ctx, [{ cmdlet: 'Delete-QuarantineMessage', params: { Identity: args.identity, Confirm: false } }]);
      return { identity: args.identity, deleted: true };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'hold',
    summary: 'Turn litigation hold on or off for a mailbox',
    description:
      'Litigation hold keeps every item, including deleted and edited ones, for --duration-days (default ' +
      'unlimited). Needs Exchange Online Plan 2 (or an archiving add-on) on the mailbox. Turning it off lets ' +
      `retention delete held items. --note sets the retention comment users see in Outlook. ${BRIDGE_NOTE}`,
    examples: [
      'm365 exchange hold jane@contoso.com --on --duration-days 2555 --note "Legal case 42"',
      'm365 exchange hold jane@contoso.com --off',
    ],
    args: [
      MAILBOX_ARG,
      { name: 'on', type: 'boolean', description: 'Place the mailbox on litigation hold.' },
      { name: 'off', type: 'boolean', description: 'Remove litigation hold.' },
      { name: 'durationDays', type: 'number', description: 'Keep items N days from receipt (default unlimited).' },
      { name: 'note', type: 'string', description: 'Retention comment shown in Outlook.' },
    ],
    risk: 'destructive',
    permissions: EXO_PERMISSIONS,
    async preview(ctx, args) {
      const enable = holdDirection(args);
      const [rows] = await exo(ctx, [
        { cmdlet: 'Get-Mailbox', params: { Identity: args.mailbox }, select: HOLD_SELECT },
      ]);
      const current = holdState(first(rows, `mailbox "${args.mailbox}"`));
      return {
        current,
        action: enable
          ? `Place ${current.mailbox ?? args.mailbox} on litigation hold for ${args.durationDays ? `${args.durationDays} days` : 'an unlimited time'}`
          : `Remove litigation hold from ${current.mailbox ?? args.mailbox}`,
      };
    },
    async run(ctx, args) {
      const enable = holdDirection(args);
      await exo(ctx, [
        {
          cmdlet: 'Set-Mailbox',
          params: compact({
            Identity: args.mailbox,
            LitigationHoldEnabled: enable,
            LitigationHoldDuration: enable && args.durationDays ? args.durationDays : undefined,
            RetentionComment: enable ? args.note : undefined,
          }),
        },
      ]);
      return {
        mailbox: args.mailbox,
        litigationHold: enable,
        durationDays: enable ? (args.durationDays ?? 'unlimited') : null,
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'archive enable',
    summary: 'Enable the online archive for a mailbox',
    description: `Enable-Mailbox -Archive. The archive needs no extra license up to 50 GB on most business plans. ${BRIDGE_NOTE}`,
    examples: ['m365 exchange archive enable jane@contoso.com'],
    args: [MAILBOX_ARG],
    risk: 'write',
    permissions: EXO_PERMISSIONS,
    async preview(ctx, args) {
      const [rows] = await exo(ctx, [
        {
          cmdlet: 'Get-Mailbox',
          params: { Identity: args.mailbox },
          select: ['PrimarySmtpAddress', 'ArchiveStatus', 'LitigationHoldEnabled'],
        },
      ]);
      const box = first(rows, `mailbox "${args.mailbox}"`);
      return {
        mailbox: text(box.PrimarySmtpAddress),
        archiveStatus: text(box.ArchiveStatus),
        litigationHold: box.LitigationHoldEnabled === true,
        action: `Enable the online archive for ${text(box.PrimarySmtpAddress) ?? args.mailbox}`,
      };
    },
    async run(ctx, args) {
      const [rows] = await exo(ctx, [
        {
          cmdlet: 'Enable-Mailbox',
          params: { Identity: args.mailbox, Archive: true },
          select: ['PrimarySmtpAddress', 'ArchiveStatus'],
        },
      ]);
      const box = rows?.[0] ?? {};
      return {
        mailbox: text(box.PrimarySmtpAddress) ?? args.mailbox,
        archiveStatus: text(box.ArchiveStatus) ?? 'Active',
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'transport-rules',
    summary: 'List mail flow (transport) rules in priority order',
    description:
      'Mail-flow troubleshooting: rules that redirect, block, add disclaimers or change spam scoring. `mode` Audit/' +
      `AuditAndNotify means the rule is only tested, not enforced. ${BRIDGE_NOTE}`,
    examples: ['m365 exchange transport-rules'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: EXO_PERMISSIONS,
    async run(ctx, args) {
      const limit = listLimit(args);
      const [rows] = await exo(ctx, [
        {
          cmdlet: 'Get-TransportRule',
          params: { ResultSize: sizeFor(limit) },
          select: ['Name', 'State', 'Priority', 'Mode', 'Description'],
        },
      ]);
      const items = [...(rows ?? [])]
        .sort((a, b) => Number(a.Priority) - Number(b.Priority))
        .map((row) =>
          ctx.full
            ? row
            : {
                name: text(row.Name),
                state: text(row.State),
                priority: typeof row.Priority === 'number' ? row.Priority : Number(row.Priority),
                mode: text(row.Mode),
                description: text(row.Description),
              },
        );
      return toList(items, limit);
    },
  }),
];

function holdDirection(args: { on: boolean; off: boolean }): boolean {
  if (args.on === args.off) throw new CliError('USAGE', 'Pass exactly one of --on or --off');
  return args.on;
}
