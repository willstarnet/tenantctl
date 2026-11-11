import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { compact, daysAgo, settle } from '../../core/util.ts';
import {
  addresses,
  MAILBOX_ACCESS_NOTE,
  RESOLVE_PERMISSION,
  type Recipient,
  resolveFolder,
  resolveFolderNames,
  USER_ARG,
} from './common.ts';

interface DateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

export interface AutomaticRepliesSetting {
  status?: 'disabled' | 'alwaysEnabled' | 'scheduled' | string;
  externalAudience?: 'none' | 'contactsOnly' | 'all' | string;
  internalReplyMessage?: string | null;
  externalReplyMessage?: string | null;
  scheduledStartDateTime?: DateTimeTimeZone | null;
  scheduledEndDateTime?: DateTimeTimeZone | null;
}

interface MailboxSettings {
  automaticRepliesSetting?: AutomaticRepliesSetting;
  timeZone?: string;
  language?: { locale?: string; displayName?: string };
  workingHours?: { daysOfWeek?: string[]; startTime?: string; endTime?: string; timeZone?: { name?: string } };
  userPurpose?: string;
  delegateMeetingMessageDeliveryOptions?: string;
}

const MESSAGE_PREVIEW_CHARS = 300;

function plainText(html: string | null | undefined): string | null {
  if (!html) return null;
  const text = html
    .replace(/<(style|script)[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>|<\/p>|<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return text.length > MESSAGE_PREVIEW_CHARS ? `${text.slice(0, MESSAGE_PREVIEW_CHARS)}…` : text;
}

/**
 * Graph reports scheduled times as wall-clock + zone; only UTC (what Exchange returns by default) can be compared
 * to now without a Windows time zone table, so other zones make autoReplyOn unknown (null).
 */
function isAutoReplyOn(setting: AutomaticRepliesSetting | undefined, now = Date.now()): boolean | null {
  if (!setting?.status || setting.status === 'disabled') return false;
  if (setting.status === 'alwaysEnabled') return true;
  const start = setting.scheduledStartDateTime;
  const end = setting.scheduledEndDateTime;
  if (!start || !end || start.timeZone !== 'UTC' || end.timeZone !== 'UTC') return null;
  return Date.parse(`${start.dateTime}Z`) <= now && now < Date.parse(`${end.dateTime}Z`);
}

export function shapeAutoReply(setting: AutomaticRepliesSetting | undefined) {
  return {
    status: setting?.status ?? null,
    autoReplyOn: isAutoReplyOn(setting),
    scheduledStart: setting?.scheduledStartDateTime ?? null,
    scheduledEnd: setting?.scheduledEndDateTime ?? null,
    externalAudience: setting?.externalAudience ?? null,
    internalMessage: plainText(setting?.internalReplyMessage),
    externalMessage: plainText(setting?.externalReplyMessage),
  };
}

const STATS_FOLDERS: Record<string, string> = {
  inbox: 'inbox',
  sentItems: 'sentitems',
  deletedItems: 'deleteditems',
  junkEmail: 'junkemail',
};

const ISO_OFFSET = /(Z|[+-]\d{2}:?\d{2})$/i;

/**
 * --start/--end: an ISO time with an offset is converted to UTC; a wall-clock time without one is sent as-is in
 * the given (or mailbox) time zone, which Graph interprets as that zone's local time.
 */
function scheduleTime(value: string, flag: string, timeZone: string): DateTimeTimeZone {
  const parsed = Date.parse(ISO_OFFSET.test(value) ? value : `${value}Z`);
  if (Number.isNaN(parsed)) {
    throw new CliError('USAGE', `${flag} "${value}" is not an ISO date/time`, {
      hint: 'Use e.g. 2026-12-20T17:00 (mailbox time zone) or 2026-12-20T17:00:00Z.',
    });
  }
  if (ISO_OFFSET.test(value)) return { dateTime: new Date(parsed).toISOString().replace(/Z$/, ''), timeZone: 'UTC' };
  return { dateTime: value.length === 10 ? `${value}T00:00:00` : value, timeZone };
}

const MESSAGE_SELECT =
  'id,receivedDateTime,from,toRecipients,subject,isRead,hasAttachments,parentFolderId,internetMessageId,webLink';

interface RawMessage {
  id: string;
  receivedDateTime?: string;
  from?: Recipient | null;
  toRecipients?: Recipient[];
  subject?: string | null;
  isRead?: boolean;
  hasAttachments?: boolean;
  parentFolderId?: string;
  internetMessageId?: string;
  webLink?: string;
}

function shapeMessage(m: RawMessage, folderNames: Record<string, string | null>) {
  return {
    received: m.receivedDateTime ?? null,
    from: m.from?.emailAddress?.address ?? null,
    fromName: m.from?.emailAddress?.name ?? null,
    to: addresses(m.toRecipients),
    subject: m.subject ?? null,
    isRead: m.isRead ?? null,
    hasAttachments: m.hasAttachments ?? null,
    folder: m.parentFolderId ? (folderNames[m.parentFolderId] ?? m.parentFolderId) : null,
    internetMessageId: m.internetMessageId ?? null,
    id: m.id,
    webLink: m.webLink ?? null,
  };
}

/** KQL dates are M/D/YYYY in Outlook search. */
function kqlDate(iso: string): string {
  const d = new Date(iso);
  return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
}

/** KQL term without quotes; multi-word values become one restriction per word (all must match). */
function kqlTerms(property: string, value: string): string[] {
  return value
    .replace(/["()]/g, ' ')
    .split(/\s+/)
    .filter(Boolean)
    .map((word) => `${property}:${word}`);
}

/** Message id from `messages search`, or an Internet Message-ID (`<...@...>`) as seen in message trace. */
async function findMessage(ctx: Context, userRef: string, ref: string) {
  const user = await resolveUser(ctx, userRef);
  const base = `/users/${seg(user.id)}/messages`;
  const select = 'id,subject,from,receivedDateTime,parentFolderId,internetMessageId';
  let message: RawMessage | null;
  if (ref.trim().startsWith('<')) {
    const { items } = await ctx.graph.list<RawMessage>(base, {
      query: { $filter: `internetMessageId eq ${odata(ref.trim())}`, $select: select },
      limit: 2,
    });
    if (items.length > 1) {
      throw new CliError('AMBIGUOUS', `Several copies of ${ref} exist in the mailbox`, {
        hint: 'Pass the Graph message id from `tenantctl exchange messages search` instead.',
        details: items.map((m) => ({ id: m.id, subject: m.subject, received: m.receivedDateTime })),
      });
    }
    message = items[0] ?? null;
  } else {
    message = await ctx.graph.tryGet<RawMessage>(`${base}/${seg(ref.trim())}`, { query: { $select: select } });
  }
  if (!message) {
    throw new CliError('NOT_FOUND', `No message ${ref} in ${user.userPrincipalName}'s mailbox`, {
      hint: 'Find it with `tenantctl exchange messages search <user> --subject ...`.',
    });
  }
  const folders = await resolveFolderNames(ctx, user.id, [message.parentFolderId]);
  return { user, message, shaped: shapeMessage(message, folders) };
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'exchange',
    name: 'mailbox',
    summary: 'Mailbox overview: type, time zone, auto-reply, working hours, folder counts',
    description:
      'First stop for any mailbox question. userPurpose tells user/shared/room/equipment mailboxes apart. ' +
      'autoReplyOn is the effective out-of-office state right now (null when scheduled in a non-UTC zone). ' +
      'folders gives total/unread item counts for Inbox, Sent, Deleted and Junk (folder sizes are not exposed by ' +
      'Graph v1.0; use `tenantctl exchange usage --user` for storage). Each section degrades to { error } on its own. ' +
      MAILBOX_ACCESS_NOTE,
    examples: ['tenantctl exchange mailbox jane@contoso.com', 'tenantctl exchange mailbox "Conference Room A"'],
    args: [USER_ARG],
    risk: 'read',
    permissions: {
      delegated: ['MailboxSettings.Read', 'Mail.ReadBasic', RESOLVE_PERMISSION],
      application: ['MailboxSettings.Read', 'Mail.ReadBasic.All', RESOLVE_PERMISSION],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const base = `/users/${seg(user.id)}`;
      const folderTasks: Record<string, Promise<Record<string, unknown>>> = {};
      for (const [key, folder] of Object.entries(STATS_FOLDERS)) {
        folderTasks[key] = ctx.graph.get(`${base}/mailFolders/${folder}`, {
          query: { $select: 'displayName,totalItemCount,unreadItemCount' },
        });
      }
      const [{ settings }, folders] = await Promise.all([
        settle({ settings: ctx.graph.get<MailboxSettings>(`${base}/mailboxSettings`) }),
        settle(folderTasks),
      ]);
      if (ctx.full) return { user, settings, folders };
      const folderCounts: Record<string, unknown> = {};
      for (const [key, folder] of Object.entries(folders)) {
        folderCounts[key] =
          'error' in folder ? folder : { total: folder.totalItemCount ?? null, unread: folder.unreadItemCount ?? null };
      }
      if ('error' in settings) {
        return { user: user.userPrincipalName, displayName: user.displayName, settings, folders: folderCounts };
      }
      const hours = settings.workingHours;
      const autoReply = shapeAutoReply(settings.automaticRepliesSetting);
      return {
        user: user.userPrincipalName,
        displayName: user.displayName,
        userPurpose: settings.userPurpose ?? null,
        timeZone: settings.timeZone ?? null,
        language: settings.language?.locale ?? null,
        workingHours: hours
          ? {
              days: hours.daysOfWeek ?? [],
              start: hours.startTime ?? null,
              end: hours.endTime ?? null,
              timeZone: hours.timeZone?.name ?? null,
            }
          : null,
        delegateMeetingMessageDeliveryOptions: settings.delegateMeetingMessageDeliveryOptions ?? null,
        autoReplyOn: autoReply.autoReplyOn,
        autoReply,
        folders: folderCounts,
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'auto-reply get',
    summary: "Show a mailbox's automatic replies (out of office)",
    description:
      'status is disabled, alwaysEnabled or scheduled; autoReplyOn is whether replies go out right now. Messages ' +
      `are shown as plain text (use --full for the HTML). Change it with \`tenantctl exchange auto-reply set\`. ${MAILBOX_ACCESS_NOTE}`,
    examples: ['tenantctl exchange auto-reply get jane@contoso.com'],
    args: [USER_ARG],
    risk: 'read',
    permissions: {
      delegated: ['MailboxSettings.Read', RESOLVE_PERMISSION],
      application: ['MailboxSettings.Read', RESOLVE_PERMISSION],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const setting = await ctx.graph.get<AutomaticRepliesSetting>(
        `/users/${seg(user.id)}/mailboxSettings/automaticRepliesSetting`,
      );
      if (ctx.full) return setting;
      return { user: user.userPrincipalName, ...shapeAutoReply(setting) };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'auto-reply set',
    summary: 'Turn automatic replies (out of office) on, off or scheduled',
    description:
      'Pass exactly one of --on, --off, --scheduled (or none to only change messages/audience). --scheduled needs ' +
      '--start and --end: times with Z or an offset are converted to UTC; times without one are read in --time-zone ' +
      'or, by default, the mailbox time zone. Messages accept plain text or HTML. --external-audience: none (no ' +
      'external replies), contactsOnly, all. Common for leavers: set it on with a redirect message, then ' +
      `\`tenantctl exchange auto-reply get\` to confirm. ${MAILBOX_ACCESS_NOTE}`,
    examples: [
      'tenantctl exchange auto-reply set jane@contoso.com --on --internal-message "I am out until Monday."',
      'tenantctl exchange auto-reply set jane@contoso.com --scheduled --start 2026-12-20T17:00 --end 2027-01-05T08:00 --external-audience all --external-message "Out of office."',
      'tenantctl exchange auto-reply set jane@contoso.com --off',
    ],
    args: [
      USER_ARG,
      { name: 'on', type: 'boolean', description: 'Always send automatic replies.' },
      { name: 'off', type: 'boolean', description: 'Stop automatic replies.' },
      { name: 'scheduled', type: 'boolean', description: 'Send automatic replies between --start and --end.' },
      { name: 'start', type: 'string', description: 'Schedule start, ISO (2026-12-20T17:00 or with Z/offset).' },
      { name: 'end', type: 'string', description: 'Schedule end, ISO.' },
      {
        name: 'timeZone',
        type: 'string',
        description: 'Time zone for --start/--end without an offset (default: mailbox time zone).',
      },
      { name: 'internalMessage', type: 'string', description: 'Reply sent to people inside the organization.' },
      { name: 'externalMessage', type: 'string', description: 'Reply sent to external senders.' },
      {
        name: 'externalAudience',
        type: 'string',
        choices: ['none', 'contactsOnly', 'all'],
        description: 'Who outside the organization gets the external reply.',
      },
    ],
    risk: 'write',
    permissions: {
      delegated: ['MailboxSettings.ReadWrite', RESOLVE_PERMISSION],
      application: ['MailboxSettings.ReadWrite', RESOLVE_PERMISSION],
    },
    async run(ctx, args) {
      const modes = [args.on && 'alwaysEnabled', args.off && 'disabled', args.scheduled && 'scheduled'].filter(Boolean);
      if (modes.length > 1) throw new CliError('USAGE', 'Pass only one of --on, --off, --scheduled');
      if ((args.start || args.end) && !args.scheduled) {
        throw new CliError('USAGE', '--start/--end only apply with --scheduled');
      }
      if (args.scheduled && !(args.start && args.end)) {
        throw new CliError('USAGE', '--scheduled needs --start and --end', {
          hint: 'e.g. --start 2026-12-20T17:00 --end 2027-01-05T08:00',
        });
      }
      const status = modes[0] as string | undefined;
      if (
        !status &&
        args.internalMessage === undefined &&
        args.externalMessage === undefined &&
        !args.externalAudience
      ) {
        throw new CliError('USAGE', 'Nothing to change', {
          hint: 'Pass --on, --off, --scheduled or a message/audience option.',
        });
      }
      const user = await resolveUser(ctx, args.user);
      const base = `/users/${seg(user.id)}/mailboxSettings`;
      const setting: AutomaticRepliesSetting = compact({
        status,
        internalReplyMessage: args.internalMessage,
        externalReplyMessage: args.externalMessage,
        externalAudience: args.externalAudience,
      });
      if (args.scheduled && args.start && args.end) {
        const needsZone = !ISO_OFFSET.test(args.start) || !ISO_OFFSET.test(args.end);
        const zone =
          args.timeZone ??
          (needsZone ? (await ctx.graph.get<{ value?: string }>(`${base}/timeZone`)).value : undefined) ??
          'UTC';
        setting.scheduledStartDateTime = scheduleTime(args.start, '--start', zone);
        setting.scheduledEndDateTime = scheduleTime(args.end, '--end', zone);
        const startMs = Date.parse(setting.scheduledStartDateTime.dateTime);
        const endMs = Date.parse(setting.scheduledEndDateTime.dateTime);
        if (setting.scheduledStartDateTime.timeZone === setting.scheduledEndDateTime.timeZone && endMs <= startMs) {
          throw new CliError('USAGE', '--end must be after --start');
        }
      }
      const updated = await ctx.graph.patch<MailboxSettings>(base, { automaticRepliesSetting: setting });
      return { user: user.userPrincipalName, ...shapeAutoReply(updated?.automaticRepliesSetting ?? setting) };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'messages search',
    summary: 'Find messages in a mailbox by sender, subject, date, folder (metadata only)',
    description:
      'Use for "did Jane receive the invoice from X?" or to locate a phishing message before purging it. Returns ' +
      'metadata only, never bodies. --from with a full address and no --subject uses an exact $filter sorted ' +
      'newest first; --subject or a sender name uses Outlook keyword search (word matches, sorted by sent date, at ' +
      'most 1000 results). Pass the `id` (or internetMessageId) to `tenantctl exchange messages delete`. For "did it ' +
      `arrive at all", use \`tenantctl exchange trace\` instead. ${MAILBOX_ACCESS_NOTE}`,
    examples: [
      'tenantctl exchange messages search jane@contoso.com --from billing@vendor.com --since 7',
      'tenantctl exchange messages search jane@contoso.com --subject "invoice overdue" --folder junk',
    ],
    args: [
      USER_ARG,
      { name: 'from', type: 'string', description: 'Sender address or name.' },
      { name: 'subject', type: 'string', description: 'Words that must appear in the subject.' },
      { name: 'since', type: 'number', default: 30, description: 'Only messages received in the last N days.' },
      {
        name: 'folder',
        type: 'string',
        description: 'inbox, sent, deleted, junk, archive, drafts, or a top-level folder name.',
      },
      { name: 'hasAttachments', type: 'boolean', description: 'Only messages with (non-inline) attachments.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['Mail.ReadBasic', RESOLVE_PERMISSION],
      application: ['Mail.ReadBasic.All', RESOLVE_PERMISSION],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const folder = args.folder ? await resolveFolder(ctx, user.id, args.folder) : undefined;
      const path = folder
        ? `/users/${seg(user.id)}/mailFolders/${seg(folder)}/messages`
        : `/users/${seg(user.id)}/messages`;
      const since = daysAgo(args.since);
      const fromAddress = args.from?.includes('@') ? args.from.trim() : undefined;
      let query: Record<string, string>;
      // $search can't be combined with $filter or $orderby on messages, so a keyword search carries every
      // criterion as KQL.
      if (args.subject || (args.from && !fromAddress)) {
        const terms = [
          ...(args.subject ? kqlTerms('subject', args.subject) : []),
          ...(args.from ? kqlTerms('from', args.from) : []),
          `received>=${kqlDate(since)}`,
          ...(args.hasAttachments ? ['hasAttachments:true'] : []),
        ];
        query = { $search: `"${terms.join(' AND ')}"`, $select: MESSAGE_SELECT };
      } else {
        // receivedDateTime must lead the filter for $orderby on it to be accepted (InefficientFilter otherwise).
        const filters = [`receivedDateTime ge ${since}`];
        if (fromAddress) filters.push(`from/emailAddress/address eq ${odata(fromAddress)}`);
        if (args.hasAttachments) filters.push('hasAttachments eq true');
        query = { $filter: filters.join(' and '), $orderby: 'receivedDateTime desc', $select: MESSAGE_SELECT };
      }
      const r = await ctx.graph.list<RawMessage>(path, { query, limit: listLimit(args) });
      if (ctx.full) return r;
      const folderNames = await resolveFolderNames(
        ctx,
        user.id,
        r.items.map((m) => m.parentFolderId),
      );
      return { items: r.items.map((m) => shapeMessage(m, folderNames)), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'messages delete',
    summary: 'Delete a message from a mailbox (incident response purge)',
    description:
      'Removes a phishing or misdirected message. Default moves it to Deleted Items (recoverable by the user); ' +
      '--hard permanently deletes it (moves it to Recoverable Items > Purges; the user cannot restore it). Accepts ' +
      'a Graph message id from `messages search` or an Internet Message-ID (<...@...>) from `trace`. One mailbox ' +
      'per call: repeat per recipient. For tenant-wide purges use Microsoft Purview / Defender ' +
      `Threat Explorer. ${MAILBOX_ACCESS_NOTE}`,
    examples: [
      'tenantctl exchange messages delete jane@contoso.com AAMkAGI2...',
      'tenantctl exchange messages delete jane@contoso.com "<abc123@phish.example>" --hard',
    ],
    args: [
      USER_ARG,
      {
        name: 'message',
        type: 'string',
        required: true,
        positional: true,
        description: 'Message id or Internet Message-ID.',
      },
      { name: 'hard', type: 'boolean', description: 'Permanently delete instead of moving to Deleted Items.' },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['Mail.ReadWrite', RESOLVE_PERMISSION],
      application: ['Mail.ReadWrite', RESOLVE_PERMISSION],
    },
    async preview(ctx, args) {
      const { user, shaped } = await findMessage(ctx, args.user, args.message);
      const where = args.hard ? 'permanently delete it (user cannot restore it)' : 'move it to Deleted Items';
      return {
        user: user.userPrincipalName,
        message: {
          subject: shaped.subject,
          from: shaped.from,
          received: shaped.received,
          folder: shaped.folder,
          id: shaped.id,
        },
        action: `Delete "${shaped.subject}" from ${user.userPrincipalName}: ${where}`,
      };
    },
    async run(ctx, args) {
      const { user, message, shaped } = await findMessage(ctx, args.user, args.message);
      const path = `/users/${seg(user.id)}/messages/${seg(message.id)}`;
      if (args.hard) await ctx.graph.post(`${path}/permanentDelete`);
      else await ctx.graph.delete(path);
      return {
        deleted: true,
        permanent: args.hard,
        user: user.userPrincipalName,
        message: { subject: shaped.subject, from: shaped.from, received: shaped.received, id: shaped.id },
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'calendar permissions',
    summary: "Who can see or edit a user's primary calendar",
    description:
      'role: freeBusyRead (free/busy only), limitedRead (+ titles/locations), read (all details), write (edit), ' +
      'delegateWithoutPrivateEventAccess / delegateWithPrivateEventAccess (delegates who also get meeting ' +
      'messages), custom. "My Organization" has no address and is the default for everyone internal. ' +
      MAILBOX_ACCESS_NOTE,
    examples: ['tenantctl exchange calendar permissions jane@contoso.com'],
    args: [USER_ARG],
    risk: 'read',
    permissions: {
      delegated: ['Calendars.ReadBasic', RESOLVE_PERMISSION],
      application: ['Calendars.ReadBasic', RESOLVE_PERMISSION],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const r = await ctx.graph.list<{
        id: string;
        role?: string;
        allowedRoles?: string[];
        emailAddress?: { name?: string; address?: string | null };
        isInsideOrganization?: boolean;
        isRemovable?: boolean;
      }>(`/users/${seg(user.id)}/calendar/calendarPermissions`, { limit: Number.POSITIVE_INFINITY, pageSize: false });
      if (ctx.full) return r;
      return {
        items: r.items.map((p) => ({
          who: p.emailAddress?.name ?? null,
          address: p.emailAddress?.address ?? null,
          role: p.role ?? null,
          isInsideOrganization: p.isInsideOrganization ?? null,
          isRemovable: p.isRemovable ?? null,
          id: p.id,
        })),
        truncated: false,
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'calendar events',
    summary: "List a user's meetings for the next N days, in the mailbox time zone",
    description:
      'Expands recurring meetings into occurrences. Times are in the mailbox time zone (shown in timeZone) when ' +
      'it can be read, else UTC. Useful for "is Jane free Thursday" or checking a room mailbox\'s bookings. ' +
      MAILBOX_ACCESS_NOTE,
    examples: ['tenantctl exchange calendar events jane@contoso.com --days 3'],
    args: [
      USER_ARG,
      { name: 'days', type: 'number', default: 7, description: 'Days ahead to include, from now.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['Calendars.ReadBasic', 'MailboxSettings.Read', RESOLVE_PERMISSION],
      application: ['Calendars.ReadBasic', 'MailboxSettings.Read', RESOLVE_PERMISSION],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const base = `/users/${seg(user.id)}`;
      const zone = await ctx.graph
        .get<{ value?: string }>(`${base}/mailboxSettings/timeZone`)
        .then((r) => r.value || 'UTC')
        .catch(() => 'UTC');
      const now = Date.now();
      const r = await ctx.graph.list<{
        id: string;
        subject?: string;
        start?: DateTimeTimeZone;
        end?: DateTimeTimeZone;
        organizer?: Recipient;
        location?: { displayName?: string };
        isOnlineMeeting?: boolean;
        showAs?: string;
        isCancelled?: boolean;
        isAllDay?: boolean;
      }>(`${base}/calendarView`, {
        query: {
          startDateTime: new Date(now).toISOString(),
          endDateTime: new Date(now + args.days * 86_400_000).toISOString(),
          $select: 'id,subject,start,end,organizer,location,isOnlineMeeting,showAs,isCancelled,isAllDay',
          $orderby: 'start/dateTime',
        },
        headers: { Prefer: `outlook.timezone="${zone.replace(/"/g, '')}"` },
        limit: listLimit(args),
      });
      if (ctx.full) return r;
      return {
        items: r.items.map((e) => ({
          subject: e.subject ?? null,
          start: e.start?.dateTime ?? null,
          end: e.end?.dateTime ?? null,
          timeZone: e.start?.timeZone ?? zone,
          isAllDay: e.isAllDay ?? false,
          organizer: e.organizer?.emailAddress?.address ?? e.organizer?.emailAddress?.name ?? null,
          location: e.location?.displayName || null,
          isOnlineMeeting: e.isOnlineMeeting ?? false,
          showAs: e.showAs ?? null,
          isCancelled: e.isCancelled ?? false,
          id: e.id,
        })),
        truncated: r.truncated,
      };
    },
  }),
];
