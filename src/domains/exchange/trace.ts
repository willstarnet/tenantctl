import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { bytesToGb } from '../../core/util.ts';

const TRACE_PATH = '/admin/exchange/tracing/messageTraces';
const DAY_MS = 86_400_000;
const MAX_WINDOW_DAYS = 10;
const MAX_LOOKBACK_DAYS = 90;
const NEAR_QUOTA_PERCENT = 90;

const TRACE_PREREQ =
  'Prerequisite: the tenant needs the Microsoft message-trace service principal (appId ' +
  '8bd644d1-64a1-4d4b-ae52-2e0cbf64e373; create it once with `New-MgServicePrincipal -AppId ' +
  '8bd644d1-64a1-4d4b-ae52-2e0cbf64e373`). Throttled at 100 requests per 5 minutes.';

const TRACE_STATUSES = [
  'delivered',
  'failed',
  'pending',
  'expanded',
  'quarantined',
  'filteredAsSpam',
  'gettingStatus',
] as const;

interface RawTrace {
  id: string;
  senderAddress?: string;
  recipientAddress?: string;
  subject?: string;
  status?: string;
  size?: number;
  messageId?: string;
  receivedDateTime?: string;
  fromIP?: string;
  toIP?: string;
}

export interface TraceFilterArgs {
  sender?: string;
  recipient?: string;
  subject?: string;
  status?: string;
  days: number;
  start?: string;
  end?: string;
}

/**
 * The API needs both receivedDateTime bounds, a window of at most 10 days, a start within the last 90 days and
 * an end no later than now; checking here gives a clear message instead of a Graph 400.
 */
export function buildTraceFilter(args: TraceFilterArgs, now = Date.now()): string {
  const end = args.end ? Date.parse(args.end) : now;
  const start = args.start ? Date.parse(args.start) : end - args.days * DAY_MS;
  if (Number.isNaN(start) || Number.isNaN(end)) {
    throw new CliError('USAGE', '--start/--end must be ISO date/times', { hint: 'e.g. --start 2026-11-01T00:00:00Z' });
  }
  if (end <= start) throw new CliError('USAGE', 'The trace window ends before it starts');
  if (end - start > MAX_WINDOW_DAYS * DAY_MS) {
    throw new CliError('USAGE', `Message trace covers at most ${MAX_WINDOW_DAYS} days per query`, {
      hint: 'Split the search into windows with --start/--end, each 10 days or less.',
    });
  }
  if (start < now - MAX_LOOKBACK_DAYS * DAY_MS) {
    throw new CliError('USAGE', `Message trace only reaches back ${MAX_LOOKBACK_DAYS} days`, {
      hint: 'For older mail use a Purview content search or historical search (Start-HistoricalSearch) in Exchange Online PowerShell.',
    });
  }
  const iso = (ms: number) => new Date(Math.min(ms, now)).toISOString().replace(/\.\d{3}Z$/, 'Z');
  const parts = [`receivedDateTime ge ${iso(start)}`, `receivedDateTime le ${iso(end)}`];
  if (args.sender) parts.push(`senderAddress eq ${odata(args.sender.trim())}`);
  if (args.recipient) parts.push(`recipientAddress eq ${odata(args.recipient.trim())}`);
  if (args.subject) parts.push(`contains(subject, ${odata(args.subject)})`);
  if (args.status) parts.push(`status eq ${odata(args.status)}`);
  return parts.join(' and ');
}

function shapeTrace(t: RawTrace) {
  return {
    received: t.receivedDateTime ?? null,
    sender: t.senderAddress ?? null,
    recipient: t.recipientAddress ?? null,
    subject: t.subject ?? null,
    status: t.status ?? null,
    size: t.size ?? null,
    fromIP: t.fromIP || null,
    messageId: t.messageId ?? null,
    id: t.id,
  };
}

async function traceRecipient(ctx: Context, traceId: string): Promise<string> {
  // Without a date filter the API searches the default last 48 hours.
  const { items } = await ctx.graph.list<RawTrace>(TRACE_PATH, {
    query: { $filter: `id eq ${odata(traceId)}` },
    limit: 50,
  });
  if (!items.length) {
    throw new CliError('NOT_FOUND', `No message trace ${traceId} in the last 48 hours`, {
      hint: 'Pass --recipient with the recipient address shown by `m365 exchange trace`.',
    });
  }
  const recipients = [...new Set(items.map((t) => t.recipientAddress ?? ''))].filter(Boolean);
  if (recipients.length !== 1) {
    throw new CliError('AMBIGUOUS', `Trace ${traceId} has ${recipients.length} recipients`, {
      hint: 'Pass --recipient with the address you are investigating.',
      details: recipients,
    });
  }
  return recipients[0] as string;
}

interface UsageRow {
  user: string;
  displayName: string;
  storageUsedGb: number | null;
  itemCount: number | null;
  percentUsed: number | null;
  warningQuotaGb: number | null;
  prohibitSendQuotaGb: number | null;
  prohibitSendReceiveQuotaGb: number | null;
  deletedItemCount: number | null;
  deletedItemSizeGb: number | null;
  hasArchive: boolean | null;
  lastActivityDate: string | null;
  isDeleted: boolean;
  reportRefreshDate: string | null;
}

function num(value: string | undefined): number | null {
  if (value === undefined || value === '') return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function bool(value: string | undefined): boolean | null {
  if (!value) return null;
  return value.toLowerCase() === 'true';
}

/** One getMailboxUsageDetail CSV row -> compact usage; percentUsed is against the Prohibit Send quota. */
export function shapeMailboxUsage(row: Record<string, string>): UsageRow {
  const used = num(row['Storage Used (Byte)']);
  const sendQuota = num(row['Prohibit Send Quota (Byte)']);
  return {
    user: row['User Principal Name'] ?? '',
    displayName: row['Display Name'] ?? '',
    storageUsedGb: bytesToGb(used),
    itemCount: num(row['Item Count']),
    percentUsed: used !== null && sendQuota ? Math.round((used / sendQuota) * 1000) / 10 : null,
    warningQuotaGb: bytesToGb(num(row['Issue Warning Quota (Byte)'])),
    prohibitSendQuotaGb: bytesToGb(sendQuota),
    prohibitSendReceiveQuotaGb: bytesToGb(num(row['Prohibit Send/Receive Quota (Byte)'])),
    deletedItemCount: num(row['Deleted Item Count']),
    deletedItemSizeGb: bytesToGb(num(row['Deleted Item Size (Byte)'])),
    hasArchive: bool(row['Has Archive']),
    lastActivityDate: row['Last Activity Date'] || null,
    isDeleted: bool(row['Is Deleted']) ?? false,
    reportRefreshDate: row['Report Refresh Date'] || null,
  };
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'exchange',
    name: 'trace',
    summary: 'Message trace: did an email reach Exchange Online, and what happened to it',
    description:
      'Start here for "my email didn\'t arrive". One row per recipient with status: delivered, failed, pending, ' +
      'quarantined, filteredAsSpam, expanded (sent to a group). Searches the last --days (default 2, max 10 per ' +
      'query) or an explicit --start/--end window within the last 90 days. --sender/--recipient are exact SMTP ' +
      'addresses; --subject matches a substring. Then run `m365 exchange trace details <id> --recipient <addr>` ' +
      `for the hop-by-hop events and failure reason. ${TRACE_PREREQ}`,
    examples: [
      'm365 exchange trace --recipient jane@contoso.com --sender billing@vendor.com',
      'm365 exchange trace --sender jane@contoso.com --status failed --days 7',
      'm365 exchange trace --subject "Q3 report" --start 2026-10-01T00:00:00Z --end 2026-10-08T00:00:00Z',
    ],
    args: [
      { name: 'sender', type: 'string', description: 'Exact sender SMTP address.' },
      { name: 'recipient', type: 'string', description: 'Exact recipient SMTP address.' },
      { name: 'subject', type: 'string', description: 'Text contained in the subject.' },
      {
        name: 'status',
        type: 'string',
        choices: TRACE_STATUSES,
        description: 'Only messages with this delivery status.',
      },
      { name: 'days', type: 'number', default: 2, description: 'Search the last N days (max 10).' },
      { name: 'start', type: 'string', description: 'Window start, ISO (overrides --days).' },
      { name: 'end', type: 'string', description: 'Window end, ISO (default now).' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['ExchangeMessageTrace.Read.All'], application: ['ExchangeMessageTrace.Read.All'] },
    async run(ctx, args) {
      const r = await ctx.graph.list<RawTrace>(TRACE_PATH, {
        query: { $filter: buildTraceFilter(args) },
        limit: listLimit(args),
      });
      if (ctx.full) return r;
      return { items: r.items.map(shapeTrace), truncated: r.truncated };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'trace details',
    summary: 'Delivery events for one traced message and recipient (why it failed or was filtered)',
    description:
      'Answers "why didn\'t my email arrive": each event (Receive, Deliver, Fail, Spam/Malware detection, Transport ' +
      'rule, Defer...) with its action and description; the last event usually explains the outcome. traceId is ' +
      'the `id` from `m365 exchange trace`. --recipient is required for traces older than 48 hours or with several ' +
      `recipients. ${TRACE_PREREQ}`,
    examples: ['m365 exchange trace details 4451a062-48cb-e80d-e8c0-196330437ae6 --recipient jane@contoso.com'],
    args: [
      {
        name: 'traceId',
        type: 'string',
        required: true,
        positional: true,
        description: 'Trace id from `exchange trace`.',
      },
      { name: 'recipient', type: 'string', description: 'Recipient SMTP address of the traced message.' },
    ],
    risk: 'read',
    permissions: { delegated: ['ExchangeMessageTrace.Read.All'], application: ['ExchangeMessageTrace.Read.All'] },
    async run(ctx, args) {
      const recipient = args.recipient?.trim() || (await traceRecipient(ctx, args.traceId));
      const r = await ctx.graph.list<{
        dateTime?: string;
        event?: string;
        action?: string;
        description?: string;
        messageId?: string;
        data?: string;
      }>(`${TRACE_PATH}/${seg(args.traceId)}/getDetailsByRecipient(recipientAddress=${odata(recipient)})`, {
        limit: Number.POSITIVE_INFINITY,
        pageSize: false,
      });
      if (ctx.full) return r;
      const events = r.items
        .map((e) => ({
          date: e.dateTime ?? null,
          event: e.event ?? null,
          action: e.action || null,
          description: e.description ?? null,
        }))
        .sort((a, b) => (a.date ?? '').localeCompare(b.date ?? ''));
      return {
        traceId: args.traceId,
        recipient,
        messageId: r.items[0]?.messageId ?? null,
        finalEvent: events.at(-1) ?? null,
        events,
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'usage',
    summary: 'Mailbox storage per user, quotas and percent used (usage report)',
    description:
      'Answers "whose mailbox is full?". percentUsed is storage used against the Prohibit Send quota (the point ' +
      'where the user can no longer send); --near-quota keeps mailboxes at 90% or more. Sorted fullest first. Data ' +
      'comes from the Microsoft 365 usage report (refreshed daily, see reportRefreshDate). If the tenant hides user ' +
      'details in reports (Microsoft 365 admin center > Settings > Org settings > Reports), names and UPNs are ' +
      'concealed hashes and --user cannot match. Delegated callers need a reports role such as Reports Reader.',
    examples: ['m365 exchange usage --near-quota', 'm365 exchange usage --user jane@contoso.com'],
    args: [
      { name: 'user', type: 'string', description: 'Only rows whose UPN or display name contains this text.' },
      {
        name: 'nearQuota',
        type: 'boolean',
        description: `Only mailboxes at ${NEAR_QUOTA_PERCENT}% or more of the send quota.`,
      },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['Reports.Read.All'],
      application: ['Reports.Read.All'],
      roles: ['Reports Reader', 'Exchange Administrator', 'Global Reader'],
    },
    async run(ctx, args) {
      const rows = await ctx.graph.report("/reports/getMailboxUsageDetail(period='D7')");
      if (rows.length && rows.every((row) => !(row['User Principal Name'] ?? '').includes('@'))) {
        ctx.log('Usage report user names are concealed by the tenant privacy setting; --user matching is unreliable.');
      }
      const wanted = args.user?.trim().toLowerCase();
      const matching = rows
        .map((raw) => ({ raw, row: shapeMailboxUsage(raw) }))
        .filter(({ row }) => {
          if (wanted && !row.user.toLowerCase().includes(wanted) && !row.displayName.toLowerCase().includes(wanted)) {
            return false;
          }
          return !args.nearQuota || (row.percentUsed ?? 0) >= NEAR_QUOTA_PERCENT;
        })
        .sort((a, b) => (b.row.percentUsed ?? -1) - (a.row.percentUsed ?? -1))
        .map(({ raw, row }) => (ctx.full ? raw : row));
      const limit = listLimit(args);
      return { items: matching.slice(0, limit), truncated: matching.length > limit };
    },
  }),
];
