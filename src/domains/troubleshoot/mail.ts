import { odata, seg } from '../../core/graph.ts';
import { resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { isSettledError, settle } from '../../core/util.ts';
import {
  buildTraceFilter,
  commands as exchangeCommands,
  getInboxRules,
  type ShapedInboxRule,
} from '../exchange/index.ts';
import { type Finding, permissionsOf, sortFindings, unavailable, verdict } from './findings.ts';

const TRACE_PATH = '/admin/exchange/tracing/messageTraces';
/** Delivery detail is one Graph call per message; enough to explain a handful of failures. */
const MAX_DETAILS = 5;
const NOT_DELIVERED: Record<string, true> = { failed: true, quarantined: true, filteredAsSpam: true, pending: true };

interface Trace {
  id: string;
  senderAddress?: string;
  recipientAddress?: string;
  subject?: string;
  status?: string;
  receivedDateTime?: string;
  messageId?: string;
}

interface TraceEvent {
  dateTime?: string;
  event?: string;
  action?: string;
  description?: string;
}

async function traceEvents(ctx: Context, trace: Trace) {
  const { items } = await ctx.graph.list<TraceEvent>(
    `${TRACE_PATH}/${seg(trace.id)}/getDetailsByRecipient(recipientAddress=${odata(trace.recipientAddress ?? '')})`,
    { limit: Number.POSITIVE_INFINITY, pageSize: false },
  );
  return items
    .map((e) => ({
      date: e.dateTime ?? null,
      event: e.event ?? null,
      action: e.action || null,
      description: e.description ?? null,
    }))
    .sort((a, b) => (a.date ?? '').localeCompare(b.date ?? ''));
}

async function recipientRules(ctx: Context, recipient: string): Promise<ShapedInboxRule[]> {
  const user = await resolveUser(ctx, recipient);
  return getInboxRules(ctx, user.id);
}

function describe(trace: Trace): string {
  return `"${trace.subject ?? '(no subject)'}" from ${trace.senderAddress ?? 'unknown'} at ${trace.receivedDateTime ?? 'unknown time'}`;
}

function traceFindings(
  traces: Trace[],
  details: Record<string, { description: string | null }[]>,
  recipient: string,
  days: number,
): Finding[] {
  if (!traces.length) {
    return [
      {
        severity: 'medium',
        area: 'trace',
        issue: `No matching message to ${recipient} reached Exchange Online in the last ${days} days`,
        detail:
          'The message never arrived at Microsoft 365: check the sender address and spelling, and ask the sender ' +
          'for a bounce (NDR). Traces can lag 5-10 minutes behind delivery.',
      },
    ];
  }
  const findings: Finding[] = [];
  for (const trace of traces) {
    const status = trace.status ?? 'unknown';
    const lastEvent = details[trace.id]?.at(-1)?.description;
    const next = `m365 exchange trace details ${trace.id} --recipient ${trace.recipientAddress ?? recipient}`;
    if (status === 'failed') {
      findings.push({
        severity: 'high',
        area: 'delivery',
        issue: `Delivery failed: ${describe(trace)}`,
        detail: lastEvent ?? undefined,
        next,
      });
    } else if (status === 'quarantined') {
      findings.push({
        severity: 'high',
        area: 'delivery',
        issue: `Quarantined: ${describe(trace)}`,
        detail:
          `${lastEvent ?? ''} Release it from quarantine in the Microsoft Defender portal (not available via Graph).`.trim(),
        next,
      });
    } else if (status === 'filteredAsSpam') {
      findings.push({
        severity: 'medium',
        area: 'delivery',
        issue: `Delivered to Junk Email: ${describe(trace)}`,
        detail: lastEvent ?? 'Marked as spam by filtering.',
        next: `m365 exchange messages search ${recipient} --folder junk --from ${trace.senderAddress ?? '<sender>'}`,
      });
    } else if (status === 'pending' || status === 'gettingStatus') {
      findings.push({
        severity: 'info',
        area: 'delivery',
        issue: `Still in transit (${status}): ${describe(trace)}`,
        detail: lastEvent ?? undefined,
        next,
      });
    }
  }
  const delivered = traces.filter((t) => t.status === 'delivered' || t.status === 'expanded');
  if (delivered.length) {
    findings.push({
      severity: 'info',
      area: 'delivery',
      issue: `${delivered.length} message${delivered.length === 1 ? ' was' : 's were'} delivered to the mailbox`,
      detail: 'If the user cannot see it, check inbox rules (below), Junk, Focused/Other and Deleted Items.',
      next: `m365 exchange messages search ${recipient}${delivered[0]?.senderAddress ? ` --from ${delivered[0].senderAddress}` : ''}`,
    });
  }
  return findings;
}

function ruleFindings(rules: ShapedInboxRule[], recipient: string): Finding[] {
  return rules
    .filter((r) => r.enabled && (r.deletes || r.movesToFolder || r.forwards))
    .map((r) => ({
      severity: r.suspicious ? ('high' as const) : ('medium' as const),
      area: 'inbox rules',
      issue: `Inbox rule "${r.name ?? r.id}" ${r.deletes ? 'deletes' : r.movesToFolder ? `moves mail to "${r.movesToFolder}"` : 'forwards'} matching mail`,
      detail: [r.conditions.length ? `When: ${r.conditions.join('; ')}.` : 'Applies to all mail.', r.reason]
        .filter(Boolean)
        .join(' '),
      next: `m365 exchange rules list ${recipient}`,
    }));
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'troubleshoot',
    name: 'mail',
    summary: 'Why an email did not arrive: message trace, delivery events for failures, recipient inbox rules',
    description:
      'Start here for "I didn\'t get the email from X". Traces messages to the recipient (optionally from a sender ' +
      'or with a subject), pulls the delivery events for up to 5 undelivered ones (failed, quarantined, junked, ' +
      "pending) so the last event explains the outcome, and checks the recipient's inbox rules that delete, move " +
      'or forward mail. Message trace covers the last 10 days per query (90 days back) and needs the one-time ' +
      'message-trace service principal (see `m365 exchange trace --help`); inbox rules of another user need app-only ' +
      'auth.',
    examples: [
      'm365 troubleshoot mail jane@contoso.com --sender billing@vendor.com',
      'm365 troubleshoot mail jane@contoso.com --subject "Invoice" --days 5',
    ],
    args: [
      { name: 'recipient', type: 'string', required: true, positional: true, description: 'Recipient email address.' },
      { name: 'sender', type: 'string', description: 'Sender email address.' },
      { name: 'subject', type: 'string', description: 'Text the subject contains.' },
      { name: 'days', type: 'number', default: 2, description: 'How far back to trace, in days (max 10).' },
    ],
    risk: 'read',
    permissions: permissionsOf(
      exchangeCommands,
      ['exchange trace', 'exchange trace details', 'exchange rules list'],
      ['Exchange Administrator', 'Global Reader'],
    ),
    async run(ctx, args) {
      const filter = buildTraceFilter({
        recipient: args.recipient,
        sender: args.sender,
        subject: args.subject,
        days: args.days,
      });
      const sections = await settle({
        traces: ctx.graph.list<Trace>(TRACE_PATH, { query: { $filter: filter }, limit: 50 }),
        inboxRules: recipientRules(ctx, args.recipient),
      });
      if (isSettledError(sections.traces)) {
        return {
          recipient: args.recipient,
          verdict: `Could not run a message trace: ${sections.traces.error.message}`,
          unavailable: unavailable(sections),
        };
      }
      const traces = sections.traces.items.sort((a, b) =>
        (b.receivedDateTime ?? '').localeCompare(a.receivedDateTime ?? ''),
      );
      const undelivered = traces.filter((t) => NOT_DELIVERED[t.status ?? '']).slice(0, MAX_DETAILS);
      const details = Object.fromEntries(
        await Promise.all(undelivered.map(async (t) => [t.id, await traceEvents(ctx, t).catch(() => [])] as const)),
      );
      const rules = isSettledError(sections.inboxRules) ? [] : sections.inboxRules;
      const findings = sortFindings([
        ...traceFindings(traces, details, args.recipient, args.days),
        ...(traces.some((t) => t.status === 'delivered') || !traces.length ? ruleFindings(rules, args.recipient) : []),
      ]);
      const missing = unavailable(sections);
      return {
        recipient: args.recipient,
        verdict: verdict(
          args.recipient,
          findings,
          missing.map((m) => m.section),
        ),
        findings,
        unavailable: missing,
        statusCounts: traces.reduce<Record<string, number>>((counts, t) => {
          const status = t.status ?? 'unknown';
          counts[status] = (counts[status] ?? 0) + 1;
          return counts;
        }, {}),
        messages: traces.map((t) => ({
          received: t.receivedDateTime ?? null,
          sender: t.senderAddress ?? null,
          subject: t.subject ?? null,
          status: t.status ?? null,
          traceId: t.id,
          events: details[t.id],
        })),
        truncated: sections.traces.truncated,
      };
    },
  }),
];
