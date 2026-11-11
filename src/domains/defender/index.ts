import { CliError } from '../../core/errors.ts';
import { seg, odata } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact, daysAgo } from '../../core/util.ts';
import { commands as endpointCommands } from './endpoint.ts';
import { HUNT_TEMPLATES, parseParams, renderTemplate } from './templates.ts';

export { type Machine, type MachineAction, resolveMachine, shapeMachine } from './endpoint.ts';

type Json = Record<string, unknown>;

interface UserAccount {
  accountName?: string | null;
  domainName?: string | null;
  userPrincipalName?: string | null;
  displayName?: string | null;
}

interface FileDetails {
  fileName?: string | null;
  filePath?: string | null;
  sha256?: string | null;
}

/** Union of the alertEvidence subtypes we read; every field is optional because the subtype varies. */
interface Evidence {
  '@odata.type'?: string;
  userAccount?: UserAccount | null;
  deviceDnsName?: string | null;
  hostName?: string | null;
  ipAddress?: string | null;
  primaryAddress?: string | null;
  upn?: string | null;
  fileDetails?: FileDetails | null;
  url?: string | null;
  urls?: string[] | null;
  subject?: string | null;
  recipientEmailAddress?: string | null;
  deliveryAction?: string | null;
  networkMessageId?: string | null;
  p1Sender?: { emailAddress?: string | null } | null;
  displayName?: string | null;
}

interface AlertComment {
  comment?: string | null;
  createdByDisplayName?: string | null;
  createdDateTime?: string | null;
}

interface Alert {
  id: string;
  title?: string | null;
  description?: string | null;
  severity?: string | null;
  status?: string | null;
  classification?: string | null;
  determination?: string | null;
  assignedTo?: string | null;
  serviceSource?: string | null;
  detectionSource?: string | null;
  productName?: string | null;
  category?: string | null;
  categories?: string[] | null;
  mitreTechniques?: string[] | null;
  threatDisplayName?: string | null;
  recommendedActions?: string | null;
  createdDateTime?: string | null;
  lastUpdateDateTime?: string | null;
  firstActivityDateTime?: string | null;
  lastActivityDateTime?: string | null;
  resolvedDateTime?: string | null;
  incidentId?: string | null;
  alertWebUrl?: string | null;
  evidence?: Evidence[] | null;
  comments?: AlertComment[] | null;
}

interface Incident {
  id: string;
  displayName?: string | null;
  severity?: string | null;
  status?: string | null;
  classification?: string | null;
  determination?: string | null;
  assignedTo?: string | null;
  priorityScore?: number | null;
  summary?: string | null;
  description?: string | null;
  customTags?: string[] | null;
  systemTags?: string[] | null;
  redirectIncidentId?: string | null;
  resolvingComment?: string | null;
  createdDateTime?: string | null;
  lastUpdateDateTime?: string | null;
  incidentWebUrl?: string | null;
  alerts?: Alert[];
  comments?: AlertComment[] | null;
}

export interface EvidenceFile {
  name: string | null;
  path?: string;
  sha256?: string;
}

export interface EvidenceEmail {
  subject: string | null;
  sender: string | null;
  recipient: string | null;
  deliveryAction: string | null;
}

/** Entities named in alert evidence, deduplicated case-insensitively. Empty categories are omitted. */
export interface EvidenceSummary {
  users?: string[];
  devices?: string[];
  ips?: string[];
  mailboxes?: string[];
  urls?: string[];
  files?: EvidenceFile[];
  emails?: EvidenceEmail[];
  apps?: string[];
}

export interface HuntResult {
  columns: { name: string; type: string }[];
  rows: Json[];
}

// Newer serviceSource values (Sentinel, Defender for Cloud...) and evidence enums are only returned and
// filterable when the caller opts in to evolvable enum members.
const PREFER_ENUMS = { Prefer: 'include-unknown-enum-members' };

const SEVERITIES = ['informational', 'low', 'medium', 'high'] as const;
const CLASSIFICATIONS = ['unknown', 'falsePositive', 'truePositive', 'informationalExpectedActivity'] as const;
const DETERMINATIONS = [
  'unknown',
  'apt',
  'malware',
  'securityPersonnel',
  'securityTesting',
  'unwantedSoftware',
  'other',
  'multiStagedAttack',
  'compromisedAccount',
  'phishing',
  'maliciousUserActivity',
  'notMalicious',
  'notEnoughDataToValidate',
  'confirmedActivity',
  'lineOfBusinessApplication',
] as const;
const SERVICE_SOURCES = [
  'microsoftDefenderForEndpoint',
  'microsoftDefenderForIdentity',
  'microsoftDefenderForCloudApps',
  'microsoftDefenderForOffice365',
  'microsoft365Defender',
  'azureAdIdentityProtection',
  'microsoftAppGovernance',
  'dataLossPrevention',
  'microsoftDefenderForCloud',
  'microsoftSentinel',
  'microsoftThreatIntelligence',
  'microsoftSecurityForAI',
] as const;

const READ_INCIDENTS = {
  delegated: ['SecurityIncident.Read.All'],
  application: ['SecurityIncident.Read.All'],
  roles: ['Security Reader', 'Security Operator', 'Security Administrator'],
};
const WRITE_INCIDENTS = {
  delegated: ['SecurityIncident.ReadWrite.All'],
  application: ['SecurityIncident.ReadWrite.All'],
  roles: ['Security Operator', 'Security Administrator'],
};
const READ_ALERTS = {
  delegated: ['SecurityAlert.Read.All'],
  application: ['SecurityAlert.Read.All'],
  roles: ['Security Reader', 'Security Operator', 'Security Administrator'],
};
const WRITE_ALERTS = {
  delegated: ['SecurityAlert.ReadWrite.All'],
  application: ['SecurityAlert.ReadWrite.All'],
  roles: ['Security Operator', 'Security Administrator'],
};
const HUNTING = {
  delegated: ['ThreatHunting.Read.All'],
  application: ['ThreatHunting.Read.All'],
  roles: ['Security Reader', 'Security Operator', 'Security Administrator'],
};

const DAYS_ARG = { name: 'days', type: 'number', description: 'Only items created in the last N days.' } as const;
const SEVERITY_ARG = {
  name: 'severity',
  type: 'string',
  choices: SEVERITIES,
  description: 'Only this severity.',
} as const;
const CLASSIFICATION_ARG = {
  name: 'classification',
  type: 'string',
  choices: CLASSIFICATIONS,
  description: 'Verdict: truePositive (real threat), falsePositive, informationalExpectedActivity (benign), unknown.',
} as const;
const DETERMINATION_ARG = {
  name: 'determination',
  type: 'string',
  choices: DETERMINATIONS,
  description: 'What it was, e.g. phishing, malware, compromisedAccount, notMalicious, securityTesting.',
} as const;
const COMMENT_ARG = { name: 'comment', type: 'string', description: 'Add a comment to the activity log.' } as const;
const TIMESPAN_ARG = {
  name: 'timespan',
  type: 'string',
  description:
    'How far back to look: ISO 8601 duration (P7D, PT12H), an interval (2026-01-01T00:00:00Z/P7D), or shorthand ' +
    'like 7d. Default 30 days.',
} as const;

class Bucket<T> {
  private readonly seen = new Map<string, T>();

  add(key: string | null | undefined, value: T): void {
    if (!key) return;
    const normalized = key.toLowerCase();
    if (!this.seen.has(normalized)) this.seen.set(normalized, value);
  }

  get values(): T[] | undefined {
    return this.seen.size ? [...this.seen.values()] : undefined;
  }
}

function accountName(account: UserAccount | null | undefined): string | null {
  if (!account) return null;
  if (account.userPrincipalName) return account.userPrincipalName;
  if (!account.accountName) return account.displayName ?? null;
  return account.domainName ? `${account.domainName}\\${account.accountName}` : account.accountName;
}

/**
 * Reduce alert evidence (any mix of subtypes, across one or many alerts) to the people, machines and
 * indicators an admin acts on. Process evidence contributes only a real user (UPN), not service accounts.
 */
export function summarizeEvidence(evidence: readonly Evidence[]): EvidenceSummary {
  const users = new Bucket<string>();
  const devices = new Bucket<string>();
  const ips = new Bucket<string>();
  const mailboxes = new Bucket<string>();
  const urls = new Bucket<string>();
  const files = new Bucket<EvidenceFile>();
  const emails = new Bucket<EvidenceEmail>();
  const apps = new Bucket<string>();
  for (const item of evidence) {
    const type = (item['@odata.type'] ?? '').replace('#microsoft.graph.security.', '');
    switch (type) {
      case 'userEvidence': {
        const name = accountName(item.userAccount);
        users.add(name, name as string);
        break;
      }
      case 'processEvidence': {
        const upn = item.userAccount?.userPrincipalName;
        users.add(upn, upn as string);
        break;
      }
      case 'deviceEvidence': {
        const name = item.deviceDnsName ?? item.hostName;
        devices.add(name, name as string);
        break;
      }
      case 'ipEvidence':
        ips.add(item.ipAddress, item.ipAddress as string);
        break;
      case 'mailboxEvidence': {
        const address = item.primaryAddress ?? item.upn;
        mailboxes.add(address, address as string);
        break;
      }
      case 'urlEvidence':
        urls.add(item.url, item.url as string);
        break;
      case 'fileEvidence': {
        const file = item.fileDetails;
        if (!file) break;
        files.add(file.sha256 ?? `${file.filePath ?? ''}\\${file.fileName ?? ''}`, {
          name: file.fileName ?? null,
          ...(file.filePath ? { path: file.filePath } : {}),
          ...(file.sha256 ? { sha256: file.sha256 } : {}),
        });
        break;
      }
      case 'analyzedMessageEvidence': {
        mailboxes.add(item.recipientEmailAddress, item.recipientEmailAddress as string);
        for (const url of item.urls ?? []) urls.add(url, url);
        emails.add(`${item.networkMessageId ?? item.subject}|${item.recipientEmailAddress}`, {
          subject: item.subject ?? null,
          sender: item.p1Sender?.emailAddress ?? null,
          recipient: item.recipientEmailAddress ?? null,
          deliveryAction: item.deliveryAction ?? null,
        });
        break;
      }
      case 'oauthApplicationEvidence':
      case 'cloudApplicationEvidence':
        apps.add(item.displayName, item.displayName as string);
        break;
    }
  }
  return compact({
    users: users.values,
    devices: devices.values,
    ips: ips.values,
    mailboxes: mailboxes.values,
    urls: urls.values,
    files: files.values,
    emails: emails.values,
    apps: apps.values,
  });
}

/** Compact list row for an alert, with the first user and device from its evidence. */
export function shapeAlert(alert: Alert) {
  const entities = summarizeEvidence(alert.evidence ?? []);
  return {
    id: alert.id,
    title: alert.title ?? null,
    severity: alert.severity ?? null,
    status: alert.status ?? null,
    serviceSource: alert.serviceSource ?? null,
    category: alert.categories?.[0] ?? alert.category ?? null,
    createdDateTime: alert.createdDateTime ?? null,
    incidentId: alert.incidentId ?? null,
    user: entities.users?.[0] ?? entities.mailboxes?.[0] ?? null,
    device: entities.devices?.[0] ?? null,
    alertWebUrl: alert.alertWebUrl ?? null,
  };
}

function shapeComments(comments: AlertComment[] | null | undefined) {
  return (comments ?? []).map((c) => ({
    comment: c.comment ?? null,
    by: c.createdByDisplayName ?? null,
    at: c.createdDateTime ?? null,
  }));
}

function shapeAlertDetail(alert: Alert) {
  return {
    id: alert.id,
    title: alert.title ?? null,
    description: alert.description ?? null,
    severity: alert.severity ?? null,
    status: alert.status ?? null,
    classification: alert.classification ?? null,
    determination: alert.determination ?? null,
    assignedTo: alert.assignedTo ?? null,
    serviceSource: alert.serviceSource ?? null,
    detectionSource: alert.detectionSource ?? null,
    categories: alert.categories?.length ? alert.categories : alert.category ? [alert.category] : [],
    mitreTechniques: alert.mitreTechniques ?? [],
    threat: alert.threatDisplayName ?? null,
    createdDateTime: alert.createdDateTime ?? null,
    firstActivityDateTime: alert.firstActivityDateTime ?? null,
    lastActivityDateTime: alert.lastActivityDateTime ?? null,
    resolvedDateTime: alert.resolvedDateTime ?? null,
    incidentId: alert.incidentId ?? null,
    recommendedActions: alert.recommendedActions ?? null,
    entities: summarizeEvidence(alert.evidence ?? []),
    comments: shapeComments(alert.comments),
    alertWebUrl: alert.alertWebUrl ?? null,
  };
}

function shapeIncident(incident: Incident) {
  return {
    id: incident.id,
    displayName: incident.displayName ?? null,
    severity: incident.severity ?? null,
    status: incident.status ?? null,
    classification: incident.classification ?? null,
    determination: incident.determination ?? null,
    assignedTo: incident.assignedTo ?? null,
    priorityScore: incident.priorityScore ?? null,
    createdDateTime: incident.createdDateTime ?? null,
    lastUpdateDateTime: incident.lastUpdateDateTime ?? null,
    alertCount: incident.alerts?.length ?? null,
    incidentWebUrl: incident.incidentWebUrl ?? null,
  };
}

/** Run a KQL query through Graph advanced hunting. Rows are returned as-is, keyed by column name. */
export async function runHunting(ctx: Context, kql: string, timespan?: string): Promise<HuntResult> {
  const response = await ctx.graph.post<{ schema?: { name: string; type: string }[]; results?: Json[] }>(
    '/security/runHuntingQuery',
    compact({ Query: kql, Timespan: timespan }),
    // Subjects and file names with lookalike characters need an explicit charset, per the API docs.
    { headers: { 'Content-Type': 'application/json; charset=utf-8' } },
  );
  return {
    columns: (response.schema ?? []).map((column) => ({ name: column.name, type: column.type })),
    rows: response.results ?? [],
  };
}

function normalizeTimespan(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const trimmed = value.trim();
  const shorthand = trimmed.match(/^(\d+)\s*([dh])$/i);
  if (shorthand) {
    const [, amount, unit] = shorthand;
    return (unit as string).toLowerCase() === 'd' ? `P${amount}D` : `PT${amount}H`;
  }
  if (/^(P|\d{4}-\d{2}-\d{2})/i.test(trimmed)) return trimmed;
  throw new CliError('USAGE', `--timespan "${value}" is not an ISO 8601 duration or interval`, {
    hint: 'Use P7D (7 days), PT12H (12 hours), 2026-01-01T00:00:00Z/P7D, or shorthand 7d.',
  });
}

interface UpdateArgs {
  status?: string;
  assignTo?: string;
  classification?: string;
  determination?: string;
  comment?: string;
}

/** PATCH body with only the fields the caller passed. */
function updateBody(label: string, args: UpdateArgs): Json {
  const owner = args.assignTo?.trim();
  const body: Json = compact({
    status: args.status,
    // An empty or "none" owner clears the assignment.
    assignedTo: owner === undefined ? undefined : ['', 'none', 'null'].includes(owner.toLowerCase()) ? null : owner,
    classification: args.classification,
    determination: args.determination,
  });
  if (!Object.keys(body).length && !args.comment) {
    throw new CliError('USAGE', `Nothing to update on the ${label}`, {
      hint: 'Pass at least one of --status, --assign-to, --classification, --determination, --comment.',
    });
  }
  return body;
}

async function addComment(ctx: Context, path: string, comment: string): Promise<void> {
  await ctx.graph.post(path, { '@odata.type': 'microsoft.graph.security.alertComment', comment });
}

/** Page alerts and keep those matching a client-side predicate (category isn't filterable server-side). */
async function listAlerts(
  ctx: Context,
  query: Json,
  limit: number,
  category: string | undefined,
): Promise<ListResult<Alert>> {
  if (!category) {
    return ctx.graph.list<Alert>('/security/alerts_v2', {
      query: query as Record<string, string>,
      limit,
      headers: PREFER_ENUMS,
    });
  }
  const wanted = category.toLowerCase();
  const items: Alert[] = [];
  let next: string | undefined = ctx.graph.url('/security/alerts_v2', {
    query: { ...(query as Record<string, string>), $top: 500 },
  });
  while (next) {
    const page: { value?: Alert[]; '@odata.nextLink'?: string } = await ctx.graph.get(next, { headers: PREFER_ENUMS });
    for (const alert of page.value ?? []) {
      const categories = [...(alert.categories ?? []), alert.category ?? ''].map((c) => c.toLowerCase());
      if (!categories.includes(wanted)) continue;
      if (items.length >= limit) return { items, truncated: true };
      items.push(alert);
    }
    next = page['@odata.nextLink'];
    if (next && items.length >= limit) return { items, truncated: true };
  }
  return { items, truncated: false };
}

/**
 * Hunting gives AlertInfo rows without incident or status; look each alert up in alerts_v2 so the agent can
 * jump straight to the incident. Alerts the API can't return keep null fields.
 */
async function alertsFromHunting(ctx: Context, rows: Json[]) {
  return Promise.all(
    rows.map(async (row) => {
      const id = String(row.AlertId ?? '');
      const alert = id
        ? await ctx.graph.tryGet<Alert>(`/security/alerts_v2/${seg(id)}`, { headers: PREFER_ENUMS }).catch(() => null)
        : null;
      return {
        alertId: id,
        title: row.Title ?? null,
        severity: typeof row.Severity === 'string' ? row.Severity.toLowerCase() : null,
        category: row.Category ?? null,
        serviceSource: row.ServiceSource ?? null,
        time: row.Timestamp ?? null,
        status: alert?.status ?? null,
        incidentId: alert?.incidentId ?? null,
        alertWebUrl: alert?.alertWebUrl ?? null,
      };
    }),
  );
}

const HUNT_ALERTS_PERMISSIONS = {
  delegated: ['ThreatHunting.Read.All', 'SecurityAlert.Read.All'],
  application: ['ThreatHunting.Read.All', 'SecurityAlert.Read.All'],
  roles: ['Security Reader', 'Security Operator', 'Security Administrator'],
};

const LICENSE_NOTE =
  'Needs Microsoft Defender XDR (any Defender product: Defender for Endpoint P2, Office 365 P2, Identity or ' +
  'Cloud Apps, e.g. via Microsoft 365 E5 / E5 Security).';

const securityCommands: AnyCommand[] = [
  defineCommand({
    domain: 'defender',
    name: 'incidents list',
    summary: 'List Defender XDR incidents (correlated alerts), newest first',
    description:
      'Start here for "is anything going on?" or "what is open?". Each row is an incident: a group of related ' +
      'alerts across Defender products. alertCount shows how many alerts it groups; status active means nobody ' +
      'closed it; redirected means it was merged into another incident. Use --unassigned to find work nobody owns. ' +
      `Next: \`tenantctl defender incidents get <id>\` for alerts and affected users/devices. ${LICENSE_NOTE}`,
    examples: [
      'tenantctl defender incidents list --status active --severity high',
      'tenantctl defender incidents list --days 7 --unassigned',
    ],
    args: [
      {
        name: 'status',
        type: 'string',
        choices: ['active', 'inProgress', 'resolved', 'redirected'],
        description: 'Only this status.',
      },
      SEVERITY_ARG,
      DAYS_ARG,
      {
        name: 'assignedTo',
        type: 'string',
        description: 'Only incidents owned by this person (UPN as shown in Defender).',
      },
      { name: 'unassigned', type: 'boolean', description: 'Only incidents with no owner.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: READ_INCIDENTS,
    async run(ctx, args) {
      if (args.unassigned && args.assignedTo) {
        throw new CliError('USAGE', '--unassigned and --assigned-to cannot be combined');
      }
      const filter = [
        args.status && `status eq ${odata(args.status)}`,
        args.severity && `severity eq ${odata(args.severity)}`,
        args.days !== undefined && `createdDateTime ge ${daysAgo(args.days)}`,
        args.assignedTo && `assignedTo eq ${odata(args.assignedTo)}`,
        args.unassigned && 'assignedTo eq null',
      ]
        .filter(Boolean)
        .join(' and ');
      // The incidents API serves at most 50 per page; $expand=alerts is how alertCount is known.
      const result = await ctx.graph.list<Incident>('/security/incidents', {
        query: { $filter: filter || undefined, $expand: 'alerts' },
        limit: listLimit(args),
        pageSize: 50,
        headers: PREFER_ENUMS,
      });
      if (ctx.full) return result;
      return { items: result.items.map(shapeIncident), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'incidents get',
    summary: 'Show an incident with its alerts, affected entities and comments',
    description:
      'Use to explain an incident: what fired (alerts with severity, product, category and MITRE techniques), ' +
      'who and what is involved (entities: users, devices, IPs, mailboxes, files, URLs, emails, apps, merged ' +
      'across all alerts) and what analysts said (comments). Next: investigate a user with ' +
      '`tenantctl entra users get`, a device with `tenantctl intune devices get`, or close the incident with ' +
      '`tenantctl defender incidents update <id> --status resolved --classification ...`.',
    examples: ['tenantctl defender incidents get 3962396'],
    args: [{ name: 'id', type: 'string', required: true, positional: true, description: 'Incident id.' }],
    risk: 'read',
    permissions: READ_INCIDENTS,
    async run(ctx, args) {
      const incident = await ctx.graph.get<Incident>(`/security/incidents/${seg(args.id.trim())}`, {
        query: { $expand: 'alerts' },
        headers: PREFER_ENUMS,
      });
      if (ctx.full) return incident;
      const alerts = incident.alerts ?? [];
      return {
        ...shapeIncident(incident),
        summary: incident.summary ?? null,
        description: incident.description ?? null,
        tags: [...(incident.customTags ?? []), ...(incident.systemTags ?? [])],
        redirectIncidentId: incident.redirectIncidentId ?? null,
        resolvingComment: incident.resolvingComment ?? null,
        entities: summarizeEvidence(alerts.flatMap((alert) => alert.evidence ?? [])),
        alerts: alerts.map((alert) => ({
          id: alert.id,
          title: alert.title ?? null,
          severity: alert.severity ?? null,
          status: alert.status ?? null,
          serviceSource: alert.serviceSource ?? null,
          category: alert.categories?.[0] ?? alert.category ?? null,
          mitreTechniques: alert.mitreTechniques ?? [],
          createdDateTime: alert.createdDateTime ?? null,
        })),
        comments: shapeComments(incident.comments),
      };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'incidents update',
    summary: 'Assign, triage, resolve or comment on an incident',
    description:
      'Only the options you pass change. Resolving: --status resolved with --classification (truePositive, ' +
      'falsePositive, informationalExpectedActivity) and a matching --determination. --assign-to takes the owner ' +
      'as free text (normally a UPN); "none" clears it. --comment is added to the incident activity log. ' +
      'Updating an incident also updates its alerts.',
    examples: [
      'tenantctl defender incidents update 3962396 --assign-to secops@contoso.com --status inProgress',
      'tenantctl defender incidents update 3962396 --status resolved --classification falsePositive --determination notMalicious --comment "Pen test"',
    ],
    args: [
      { name: 'id', type: 'string', required: true, positional: true, description: 'Incident id.' },
      { name: 'status', type: 'string', choices: ['active', 'inProgress', 'resolved'], description: 'New status.' },
      { name: 'assignTo', type: 'string', description: 'Owner (UPN or free text); "none" unassigns.' },
      CLASSIFICATION_ARG,
      DETERMINATION_ARG,
      COMMENT_ARG,
    ],
    risk: 'write',
    permissions: WRITE_INCIDENTS,
    async preview(ctx, args) {
      const changes = updateBody('incident', args);
      const current = await ctx.graph.get<Incident>(`/security/incidents/${seg(args.id.trim())}`, {
        headers: PREFER_ENUMS,
      });
      return { incident: shapeIncident(current), changes, comment: args.comment ?? null };
    },
    async run(ctx, args) {
      const path = `/security/incidents/${seg(args.id.trim())}`;
      const body = updateBody('incident', args);
      let incident: Incident | undefined;
      if (Object.keys(body).length) incident = await ctx.graph.patch<Incident>(path, body, { headers: PREFER_ENUMS });
      if (args.comment) await addComment(ctx, `${path}/comments`, args.comment);
      incident ??= await ctx.graph.get<Incident>(path, { headers: PREFER_ENUMS });
      return { updated: Object.keys(body), commented: Boolean(args.comment), incident: shapeIncident(incident) };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'alerts list',
    summary: 'List Defender alerts from all Microsoft security products',
    description:
      'Alerts are individual detections (an incident groups several). user/device are the first ones named in ' +
      'the evidence. Filter by --source to one product (e.g. microsoftDefenderForOffice365 for email threats, ' +
      'azureAdIdentityProtection for risky sign-ins). --category (e.g. InitialAccess, Execution, ' +
      'CredentialAccess) is matched client-side, so it pages until enough matches are found; combine it with ' +
      `--days. Next: \`tenantctl defender alerts get <id>\` or \`incidents get <incidentId>\`. ${LICENSE_NOTE}`,
    examples: [
      'tenantctl defender alerts list --status new --severity high',
      'tenantctl defender alerts list --source microsoftDefenderForOffice365 --days 7',
    ],
    args: [
      SEVERITY_ARG,
      { name: 'status', type: 'string', choices: ['new', 'inProgress', 'resolved'], description: 'Only this status.' },
      { name: 'source', type: 'string', choices: SERVICE_SOURCES, description: 'Only alerts from this product.' },
      DAYS_ARG,
      { name: 'category', type: 'string', description: 'MITRE ATT&CK tactic, e.g. InitialAccess, Execution.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: READ_ALERTS,
    async run(ctx, args) {
      const filter = [
        args.severity && `severity eq ${odata(args.severity)}`,
        args.status && `status eq ${odata(args.status)}`,
        args.source && `serviceSource eq ${odata(args.source)}`,
        args.days !== undefined && `createdDateTime ge ${daysAgo(args.days)}`,
      ]
        .filter(Boolean)
        .join(' and ');
      const result = await listAlerts(ctx, filter ? { $filter: filter } : {}, listLimit(args), args.category);
      if (ctx.full) return result;
      return { items: result.items.map(shapeAlert), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'alerts get',
    summary: 'Show one alert with evidence, MITRE techniques and recommended actions',
    description:
      'entities summarizes the evidence (users, devices, IPs, mailboxes, files, URLs, emails, apps). ' +
      "recommendedActions is the product's remediation advice. Next: `incidents get <incidentId>` for the " +
      'wider picture, or `alerts update <id>` to triage.',
    examples: ['tenantctl defender alerts get da637551227677560813_-961444813'],
    args: [{ name: 'id', type: 'string', required: true, positional: true, description: 'Alert id.' }],
    risk: 'read',
    permissions: READ_ALERTS,
    async run(ctx, args) {
      const alert = await ctx.graph.get<Alert>(`/security/alerts_v2/${seg(args.id.trim())}`, {
        headers: PREFER_ENUMS,
      });
      return ctx.full ? alert : shapeAlertDetail(alert);
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'alerts update',
    summary: 'Assign, triage, resolve or comment on an alert',
    description:
      'Only the options you pass change. --assign-to takes the owner as free text (normally a UPN); "none" ' +
      'clears it. To close a whole attack prefer `incidents update`, which updates all of its alerts.',
    examples: [
      'tenantctl defender alerts update da637551227677560813_-961444813 --status resolved --classification truePositive --determination malware',
    ],
    args: [
      { name: 'id', type: 'string', required: true, positional: true, description: 'Alert id.' },
      { name: 'status', type: 'string', choices: ['new', 'inProgress', 'resolved'], description: 'New status.' },
      { name: 'assignTo', type: 'string', description: 'Owner (UPN or free text); "none" unassigns.' },
      CLASSIFICATION_ARG,
      DETERMINATION_ARG,
      COMMENT_ARG,
    ],
    risk: 'write',
    permissions: WRITE_ALERTS,
    async preview(ctx, args) {
      const changes = updateBody('alert', args);
      const current = await ctx.graph.get<Alert>(`/security/alerts_v2/${seg(args.id.trim())}`, {
        headers: PREFER_ENUMS,
      });
      return { alert: shapeAlert(current), changes, comment: args.comment ?? null };
    },
    async run(ctx, args) {
      const path = `/security/alerts_v2/${seg(args.id.trim())}`;
      const body = updateBody('alert', args);
      let alert: Alert | undefined;
      if (Object.keys(body).length) alert = await ctx.graph.patch<Alert>(path, body, { headers: PREFER_ENUMS });
      if (args.comment) await addComment(ctx, `${path}/comments`, args.comment);
      alert ??= await ctx.graph.get<Alert>(path, { headers: PREFER_ENUMS });
      return { updated: Object.keys(body), commented: Boolean(args.comment), alert: shapeAlert(alert) };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'alerts for-user',
    summary: 'Alerts that involve a user, across all Defender products',
    description:
      'Runs the alerts-for-user hunting query (AlertEvidence joined to AlertInfo) and looks each alert up to ' +
      'add its status and incident. Use when asked "has anything flagged Jane?". Matches evidence by UPN or ' +
      `object id. Hunting data covers at most 30 days. ${LICENSE_NOTE}`,
    examples: ['tenantctl defender alerts for-user jane@contoso.com --days 7'],
    args: [
      { name: 'user', type: 'string', required: true, positional: true, description: 'UPN, email, name or id.' },
      { name: 'days', type: 'number', default: 30, description: 'Look back this many days (max 30).' },
      { name: 'top', type: 'number', default: 25, description: 'Maximum alerts to return.' },
    ],
    risk: 'read',
    permissions: {
      ...HUNT_ALERTS_PERMISSIONS,
      delegated: [...HUNT_ALERTS_PERMISSIONS.delegated, 'User.Read.All'],
      application: [...HUNT_ALERTS_PERMISSIONS.application, 'User.Read.All'],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const kql = `${renderTemplate('alerts-for-user', { user: user.userPrincipalName })}\n| take ${args.top + 1}`;
      const result = await runHunting(ctx, kql, `P${args.days}D`);
      const rows = result.rows.slice(0, args.top);
      return {
        user: { id: user.id, userPrincipalName: user.userPrincipalName, displayName: user.displayName },
        days: args.days,
        alerts: ctx.full ? rows : await alertsFromHunting(ctx, rows),
        truncated: result.rows.length > args.top,
      };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'alerts for-device',
    summary: 'Alerts that involve a device, across all Defender products',
    description:
      "Runs the alerts-for-device hunting query and adds each alert's status and incident. The device is " +
      'matched by name (short name or FQDN) or Defender device id, as Defender knows it; no Intune lookup. ' +
      `Hunting data covers at most 30 days. ${LICENSE_NOTE}`,
    examples: ['tenantctl defender alerts for-device DESKTOP-123 --days 14'],
    args: [
      { name: 'device', type: 'string', required: true, positional: true, description: 'Device name or id.' },
      { name: 'days', type: 'number', default: 30, description: 'Look back this many days (max 30).' },
      { name: 'top', type: 'number', default: 25, description: 'Maximum alerts to return.' },
    ],
    risk: 'read',
    permissions: HUNT_ALERTS_PERMISSIONS,
    async run(ctx, args) {
      const device = args.device.trim();
      const kql = `${renderTemplate('alerts-for-device', { device })}\n| take ${args.top + 1}`;
      const result = await runHunting(ctx, kql, `P${args.days}D`);
      const rows = result.rows.slice(0, args.top);
      return {
        device,
        days: args.days,
        alerts: ctx.full ? rows : await alertsFromHunting(ctx, rows),
        truncated: result.rows.length > args.top,
      };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'hunt',
    summary: 'Run an advanced hunting (KQL) query or a curated template',
    description:
      'Pass KQL directly, or --template <name> with --param key=value (see `tenantctl defender hunt templates`). ' +
      'Template values are inserted as escaped KQL string literals. Output: columns (name, type) and rows. ' +
      'Graph returns up to 100,000 rows; this command shows the first --top, so add `| take`, `| summarize` or ' +
      '`| project` to keep results small. Data reaches back 30 days (default timespan). Long-running or ' +
      'expensive queries can be throttled (429) per tenant CPU quota. Read-only even though it is a POST. ' +
      LICENSE_NOTE,
    examples: [
      'tenantctl defender hunt "DeviceInfo | summarize arg_max(Timestamp, *) by DeviceName | take 10"',
      'tenantctl defender hunt --template emails-from-sender --param sender=bad@evil.example --timespan P7D',
    ],
    args: [
      { name: 'query', type: 'string', positional: true, description: 'KQL query (omit when using --template).' },
      { name: 'template', type: 'string', choices: Object.keys(HUNT_TEMPLATES), description: 'Curated query name.' },
      {
        name: 'param',
        type: 'string[]',
        noSplit: true,
        description: 'Template parameter as name=value; repeat --param for each key=value.',
      },
      TIMESPAN_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: HUNTING,
    async run(ctx, args) {
      if (Boolean(args.query) === Boolean(args.template)) {
        throw new CliError('USAGE', 'Pass either a KQL query or --template, not both or neither', {
          hint: 'See `tenantctl defender hunt templates` for ready-made queries.',
        });
      }
      const kql = args.template ? renderTemplate(args.template, parseParams(args.param ?? [])) : (args.query as string);
      const result = await runHunting(ctx, kql, normalizeTimespan(args.timespan));
      const limit = listLimit(args);
      const truncated = result.rows.length > limit;
      return compact({
        query: args.template ? kql : undefined,
        columns: result.columns,
        rows: truncated ? result.rows.slice(0, limit) : result.rows,
        rowCount: result.rows.length,
        truncated,
        note: truncated
          ? `Showing ${limit} of ${result.rows.length} rows. Narrow the query or raise --top.`
          : undefined,
      });
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'hunt templates',
    summary: 'List curated hunting queries and their parameters',
    description:
      'Each template answers a common question (sign-ins for a user, emails from a sender, processes on a ' +
      'device...). Run one with `tenantctl defender hunt --template <name> --param key=value`. requires names the ' +
      'product whose data the table needs; an empty result can mean that product is not licensed.',
    examples: ['tenantctl defender hunt templates'],
    args: [],
    risk: 'read',
    permissions: {},
    async run() {
      return {
        items: Object.entries(HUNT_TEMPLATES).map(([name, template]) => ({
          name,
          description: template.description,
          params: template.params,
          requires: template.requires,
          kql: template.kql,
        })),
        truncated: false,
      } satisfies ListResult;
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'secure-score',
    summary: 'Show Microsoft Secure Score and the improvement actions worth the most points',
    description:
      'percent is currentScore / maxScore. comparison holds average scores for all tenants, tenants of similar ' +
      'size (totalSeats) and the same industry. topActions ranks controls by points still available ' +
      '(remainingPoints) with cost, user impact and a portal link (actionUrl). Use for "how secure are we?" and ' +
      '"what should we fix first?". Scores are calculated daily.',
    examples: ['tenantctl defender secure-score', 'tenantctl defender secure-score --top-actions 5'],
    args: [{ name: 'topActions', type: 'number', default: 10, description: 'How many improvement actions to show.' }],
    risk: 'read',
    permissions: {
      delegated: ['SecurityEvents.Read.All'],
      application: ['SecurityEvents.Read.All'],
      roles: ['Security Reader', 'Global Reader', 'Security Administrator'],
    },
    async run(ctx, args) {
      interface ControlScore {
        controlName?: string;
        controlCategory?: string;
        score?: number | null;
      }
      interface SecureScore {
        createdDateTime?: string;
        currentScore?: number;
        maxScore?: number;
        licensedUserCount?: number;
        activeUserCount?: number;
        enabledServices?: string[];
        averageComparativeScores?: { basis?: string; averageScore?: number }[];
        controlScores?: ControlScore[];
      }
      interface ControlProfile {
        id: string;
        title?: string;
        maxScore?: number | null;
        controlCategory?: string;
        implementationCost?: string;
        userImpact?: string;
        actionUrl?: string;
        service?: string;
        deprecated?: boolean;
        controlStateUpdates?: { state?: string }[];
      }
      const [scores, profiles] = await Promise.all([
        ctx.graph.list<SecureScore>('/security/secureScores', { limit: 1 }),
        ctx.graph.list<ControlProfile>('/security/secureScoreControlProfiles', { limit: Number.POSITIVE_INFINITY }),
      ]);
      const latest = scores.items[0];
      if (!latest) {
        throw new CliError('NOT_FOUND', 'No Secure Score has been calculated for this tenant yet', {
          hint: 'Secure Score is computed daily; new tenants may need 24-48 hours.',
        });
      }
      if (ctx.full) return { secureScore: latest, controlProfiles: profiles.items };
      const byId: Record<string, ControlProfile> = Object.fromEntries(profiles.items.map((p) => [p.id, p]));
      const actions = (latest.controlScores ?? [])
        .map((control) => {
          const profile = byId[control.controlName ?? ''];
          if (!profile || profile.deprecated) return null;
          const max = profile.maxScore ?? 0;
          const current = control.score ?? 0;
          const state = profile.controlStateUpdates?.at(-1)?.state ?? null;
          return {
            control: profile.id,
            title: profile.title ?? null,
            service: profile.service ?? null,
            category: profile.controlCategory ?? control.controlCategory ?? null,
            currentPoints: current,
            maxScore: max,
            remainingPoints: Math.round((max - current) * 100) / 100,
            implementationCost: profile.implementationCost ?? null,
            userImpact: profile.userImpact ?? null,
            state,
            actionUrl: profile.actionUrl ?? null,
          };
        })
        .filter((action): action is NonNullable<typeof action> => action !== null && action.remainingPoints > 0)
        .sort((a, b) => b.remainingPoints - a.remainingPoints);
      const current = latest.currentScore ?? 0;
      const max = latest.maxScore ?? 0;
      return {
        calculatedAt: latest.createdDateTime ?? null,
        currentScore: current,
        maxScore: max,
        percent: max ? Math.round((current / max) * 1000) / 10 : null,
        comparison: Object.fromEntries(
          (latest.averageComparativeScores ?? []).map((c) => [
            (c.basis ?? 'unknown').replace(/^./, (ch) => ch.toLowerCase()),
            c.averageScore ?? null,
          ]),
        ),
        licensedUserCount: latest.licensedUserCount ?? null,
        activeUserCount: latest.activeUserCount ?? null,
        enabledServices: latest.enabledServices ?? [],
        topActions: actions.slice(0, args.topActions),
        remainingActions: Math.max(actions.length - args.topActions, 0),
      };
    },
  }),

  defineCommand({
    domain: 'defender',
    name: 'threat-submit email',
    summary: 'Submit an email to Microsoft for analysis as phishing, spam, malware or not junk (beta API)',
    description:
      'Use when a user reports a suspicious message, or a legitimate message was junked. Identify the message ' +
      'by the recipient and its Graph message id (from `tenantctl exchange` message search or Graph). Microsoft ' +
      're-scans it; the result (status, result.category, result.detail) appears on the returned submission ' +
      'after a while. Needs Defender for Office 365 or Exchange Online Protection. (beta API)',
    examples: ['tenantctl defender threat-submit email jane@contoso.com AAMkADU3MWUxOTU0... --category phishing'],
    args: [
      { name: 'user', type: 'string', required: true, positional: true, description: 'Recipient mailbox owner.' },
      { name: 'messageId', type: 'string', required: true, positional: true, description: 'Graph message id.' },
      {
        name: 'category',
        type: 'string',
        required: true,
        choices: ['phishing', 'spam', 'malware', 'notJunk'],
        description: 'What the message is.',
      },
    ],
    risk: 'write',
    permissions: {
      delegated: ['ThreatSubmission.ReadWrite', 'User.Read.All'],
      application: ['ThreatSubmission.ReadWrite.All', 'User.Read.All'],
      roles: ['Security Administrator'],
    },
    async preview(ctx, args) {
      return { submission: await submissionBody(ctx, args.user, args.messageId, args.category) };
    },
    async run(ctx, args) {
      const body = await submissionBody(ctx, args.user, args.messageId, args.category);
      const created = await ctx.graph.post<Json>('/security/threatSubmission/emailThreats', body, { version: 'beta' });
      if (ctx.full) return created;
      return compact({
        id: created.id,
        status: created.status,
        category: created.category,
        recipientEmailAddress: created.recipientEmailAddress,
        subject: created.emailSubject ?? created.subject,
        sender: created.sender,
        result: created.result ?? null,
        createdDateTime: created.createdDateTime,
      });
    },
  }),
];

export const commands: AnyCommand[] = [...securityCommands, ...endpointCommands];

async function submissionBody(ctx: Context, userRef: string, messageId: string, category: string): Promise<Json> {
  const user = await resolveUser(ctx, userRef);
  return {
    '@odata.type': '#microsoft.graph.security.emailUrlThreatSubmission',
    category,
    recipientEmailAddress: user.mail ?? user.userPrincipalName,
    messageUrl: ctx.graph.url(`/users/${seg(user.id)}/messages/${seg(messageId.trim())}`, { version: 'beta' }),
  };
}
