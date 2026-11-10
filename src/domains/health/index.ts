import { odata, seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { type AnyCommand, defineCommand } from '../../core/types.ts';
import { daysAgo } from '../../core/util.ts';

const HEALTH_PERMISSIONS = {
  delegated: ['ServiceHealth.Read.All'],
  application: ['ServiceHealth.Read.All'],
  roles: ['Service Support Administrator', 'Global Reader'],
};
const MESSAGE_PERMISSIONS = {
  delegated: ['ServiceMessage.Read.All'],
  application: ['ServiceMessage.Read.All'],
  roles: ['Message Center Reader', 'Global Reader'],
};

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };

/**
 * Service health posts and message center bodies are HTML. Agents need the words: keep paragraph and list
 * structure as line breaks and bullets, drop every tag, decode entities.
 */
export function htmlToText(html: string | null | undefined): string {
  if (!html) return '';
  return html
    .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<li[^>]*>/gi, '\n- ')
    .replace(/<\/(p|div|h[1-6]|tr|ul|ol|table)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (entity, code: string) => {
      if (code.startsWith('#x') || code.startsWith('#X'))
        return String.fromCodePoint(Number.parseInt(code.slice(2), 16));
      if (code.startsWith('#')) return String.fromCodePoint(Number(code.slice(1)));
      return ENTITIES[code.toLowerCase()] ?? entity;
    })
    .replace(/[ \t\u00a0]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Statuses that mean no current impact: operational, or a closed issue (restored, false positive, PIR
 * published). Anything else (investigating, degradation, interruption, extended recovery...) is ongoing.
 */
const SETTLED_STATUSES: Record<string, true> = {
  serviceoperational: true,
  servicerestored: true,
  falsepositive: true,
  postincidentreviewpublished: true,
  resolved: true,
  resolvedexternal: true,
};

interface ServiceHealth {
  id: string;
  service: string;
  status: string;
}

export interface HealthSummary {
  summary: string;
  degraded: { service: string; status: string }[];
  recentlyRestored: { service: string; status: string }[];
}

/** Graph has returned both `serviceOperational` and `ServiceOperational`; compare case-insensitively. */
export function summarizeHealth(services: ServiceHealth[]): HealthSummary {
  const degraded = services
    .filter((s) => !SETTLED_STATUSES[s.status.toLowerCase()])
    .map((s) => ({ service: s.service, status: s.status }));
  const recentlyRestored = services
    .filter((s) => s.status.toLowerCase() !== 'serviceoperational' && SETTLED_STATUSES[s.status.toLowerCase()])
    .map((s) => ({ service: s.service, status: s.status }));
  const summary = degraded.length
    ? `${degraded.length} of ${services.length} services have an active issue: ${degraded.map((d) => `${d.service} (${d.status})`).join(', ')}`
    : `All ${services.length} services are healthy`;
  return { summary, degraded, recentlyRestored };
}

interface HealthIssue {
  id: string;
  title: string;
  service: string;
  status: string;
  classification: string;
  feature?: string | null;
  featureGroup?: string | null;
  startDateTime: string | null;
  endDateTime?: string | null;
  lastModifiedDateTime: string | null;
  isResolved: boolean;
  impactDescription: string | null;
  posts?: { createdDateTime: string; postType?: string; description?: { content?: string } }[];
}

export function shapeIssue(issue: HealthIssue) {
  return {
    id: issue.id,
    title: issue.title,
    service: issue.service,
    status: issue.status,
    classification: issue.classification,
    featureGroup: issue.featureGroup ?? null,
    startDateTime: issue.startDateTime,
    lastModifiedDateTime: issue.lastModifiedDateTime,
    isResolved: issue.isResolved,
    impactDescription: issue.impactDescription,
  };
}

interface ServiceMessage {
  id: string;
  title: string;
  services?: string[];
  category?: string;
  severity?: string;
  isMajorChange?: boolean | null;
  actionRequiredByDateTime?: string | null;
  startDateTime?: string | null;
  endDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  tags?: string[];
  body?: { content?: string };
  details?: { name: string; value: string }[];
  viewPoint?: { isRead?: boolean | null } | null;
}

export function shapeMessage(message: ServiceMessage) {
  return {
    id: message.id,
    title: message.title,
    services: message.services ?? [],
    category: message.category ?? null,
    severity: message.severity ?? null,
    isMajorChange: message.isMajorChange ?? null,
    actionRequiredByDateTime: message.actionRequiredByDateTime ?? null,
    lastModifiedDateTime: message.lastModifiedDateTime ?? null,
    tags: message.tags ?? [],
  };
}

const ISSUE_SELECT =
  'id,title,service,status,classification,featureGroup,startDateTime,lastModifiedDateTime,isResolved,impactDescription';
const MESSAGE_SELECT =
  'id,title,services,category,severity,isMajorChange,actionRequiredByDateTime,lastModifiedDateTime,tags';

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'health',
    name: 'overview',
    summary: 'Current health of every Microsoft 365 service the tenant uses',
    description:
      'First check when several users report the same problem ("Outlook is down", "Teams calls failing"). ' +
      '`degraded` lists services with an ongoing issue; `recentlyRestored` were fixed but not yet back to ' +
      'operational. Next: `m365 health issues --service "<service>" --active` for details.',
    examples: ['m365 health overview'],
    args: [],
    risk: 'read',
    permissions: HEALTH_PERMISSIONS,
    async run(ctx) {
      const { items } = await ctx.graph.list<ServiceHealth>('/admin/serviceAnnouncement/healthOverviews', {
        query: { $select: 'id,service,status' },
        limit: Number.POSITIVE_INFINITY,
        pageSize: false,
      });
      if (ctx.full) return items;
      return { ...summarizeHealth(items), services: items.map((s) => ({ service: s.service, status: s.status })) };
    },
  }),

  defineCommand({
    domain: 'health',
    name: 'issues',
    summary: 'List service health incidents and advisories',
    description:
      'incident = significant user impact; advisory = limited impact or a subset of users. --active hides ' +
      'resolved ones. --service takes the name shown by `m365 health overview` (e.g. "Exchange Online", ' +
      '"Microsoft Teams"). --days counts by last update. Next: `m365 health issue <id>` for the timeline.',
    examples: ['m365 health issues --active', 'm365 health issues --service "Microsoft Teams" --days 7'],
    args: [
      { name: 'service', type: 'string', description: 'Service name, e.g. "Exchange Online".' },
      { name: 'active', type: 'boolean', description: 'Only unresolved issues.' },
      { name: 'days', type: 'number', description: 'Only issues updated in the last N days.' },
      {
        name: 'classification',
        type: 'string',
        choices: ['incident', 'advisory'],
        description: 'Incident or advisory.',
      },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: HEALTH_PERMISSIONS,
    async run(ctx, args) {
      const filters: string[] = [];
      if (args.service) filters.push(`service eq ${odata(args.service)}`);
      if (args.active) filters.push('isResolved eq false');
      if (args.days !== undefined) filters.push(`lastModifiedDateTime ge ${daysAgo(args.days)}`);
      if (args.classification) filters.push(`classification eq ${odata(args.classification)}`);
      const result = await ctx.graph.list<HealthIssue>('/admin/serviceAnnouncement/issues', {
        query: { $filter: filters.length ? filters.join(' and ') : undefined, $select: ISSUE_SELECT },
        limit: listLimit(args),
      });
      return { items: ctx.full ? result.items : result.items.map(shapeIssue), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'health',
    name: 'issue',
    summary: 'Show one service health issue with its update timeline as plain text',
    description:
      "`posts` are Microsoft's updates, newest first, with HTML removed. The latest post usually has the " +
      'current status, scope of impact and next update time.',
    examples: ['m365 health issue EX123456'],
    args: [
      {
        name: 'id',
        type: 'string',
        required: true,
        positional: true,
        description: 'Issue id, e.g. EX123456 or TM654321.',
      },
    ],
    risk: 'read',
    permissions: HEALTH_PERMISSIONS,
    async run(ctx, args) {
      const issue = await ctx.graph.get<HealthIssue>(`/admin/serviceAnnouncement/issues/${seg(args.id.trim())}`);
      if (ctx.full) return issue;
      const posts = [...(issue.posts ?? [])].sort(
        (a, b) => Date.parse(b.createdDateTime) - Date.parse(a.createdDateTime),
      );
      return {
        ...shapeIssue(issue),
        feature: issue.feature ?? null,
        endDateTime: issue.endDateTime ?? null,
        posts: posts.map((p) => ({
          createdDateTime: p.createdDateTime,
          postType: p.postType ?? null,
          text: htmlToText(p.description?.content),
        })),
      };
    },
  }),

  defineCommand({
    domain: 'health',
    name: 'messages',
    summary: 'List message center posts (upcoming changes, retirements, required actions)',
    description:
      'Use for "what is changing in Teams/Exchange", "what do we need to act on". --action-required keeps ' +
      'posts with a deadline; --major keeps major changes (30 days notice). --days counts by last update. ' +
      'Categories: preventOrFixIssue, planForChange, stayInformed. Next: `m365 health message <id>`.',
    examples: [
      'm365 health messages --action-required',
      'm365 health messages --service "Microsoft Teams" --days 30 --major',
    ],
    args: [
      { name: 'service', type: 'string', description: 'Service name, e.g. "Microsoft Teams".' },
      { name: 'days', type: 'number', description: 'Only posts updated in the last N days.' },
      { name: 'major', type: 'boolean', description: 'Only major changes.' },
      {
        name: 'category',
        type: 'string',
        choices: ['preventOrFixIssue', 'planForChange', 'stayInformed'],
        description: 'Message category.',
      },
      { name: 'actionRequired', type: 'boolean', description: 'Only posts with an action-required-by date.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: MESSAGE_PERMISSIONS,
    async run(ctx, args) {
      const filters: string[] = [];
      if (args.service) filters.push(`services/any(s:s eq ${odata(args.service)})`);
      if (args.days !== undefined) filters.push(`lastModifiedDateTime ge ${daysAgo(args.days)}`);
      if (args.major) filters.push('isMajorChange eq true');
      if (args.category) filters.push(`category eq ${odata(args.category)}`);
      if (args.actionRequired) filters.push('actionRequiredByDateTime ne null');
      const result = await ctx.graph.list<ServiceMessage>('/admin/serviceAnnouncement/messages', {
        query: { $filter: filters.length ? filters.join(' and ') : undefined, $select: MESSAGE_SELECT },
        limit: listLimit(args),
      });
      return { items: ctx.full ? result.items : result.items.map(shapeMessage), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'health',
    name: 'message',
    summary: 'Show one message center post with its body as plain text',
    examples: ['m365 health message MC123456'],
    args: [{ name: 'id', type: 'string', required: true, positional: true, description: 'Message id, e.g. MC123456.' }],
    risk: 'read',
    permissions: MESSAGE_PERMISSIONS,
    async run(ctx, args) {
      const message = await ctx.graph.get<ServiceMessage>(`/admin/serviceAnnouncement/messages/${seg(args.id.trim())}`);
      if (ctx.full) return message;
      return {
        ...shapeMessage(message),
        startDateTime: message.startDateTime ?? null,
        endDateTime: message.endDateTime ?? null,
        isRead: message.viewPoint?.isRead ?? null,
        details: Object.fromEntries((message.details ?? []).map((d) => [d.name, d.value])),
        body: htmlToText(message.body?.content),
      };
    },
  }),
];
