import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact, daysAgo } from '../../core/util.ts';
import { resolveSite } from '../sharepoint/resolve.ts';

/**
 * Microsoft Purview unified audit log through the Graph Audit Search API (v1.0 /security/auditLog/queries).
 * A search is asynchronous: create a query, poll it until it succeeds, then page its records. Creating a query
 * changes no tenant configuration, so every command here is `read`.
 */

const QUERIES = '/security/auditLog/queries';
const POLL_MS = 5000;
const FAILED_STATUSES: Record<string, true> = { failed: true, cancelled: true };

/** Cross-service search: per-service scopes would silently drop other services' records. */
const AUDIT_ALL = {
  delegated: ['AuditLogsQuery.Read.All'],
  application: ['AuditLogsQuery.Read.All'],
  roles: ['Audit Reader', 'Compliance Administrator'],
} as const;

export interface AuditFilters {
  /** ISO start of the window. */
  start: string;
  /** ISO end of the window. */
  end: string;
  /** UPNs of the user who performed the activity. */
  users?: string[];
  operations?: string[];
  /** auditLogRecordType values, e.g. `exchangeAdmin`, `sharePointFileOperation`. */
  recordTypes?: string[];
  /** Free text matched against the workload-specific auditData (not operation names). */
  keyword?: string;
  /** SharePoint/OneDrive: full file or folder URL; Exchange admin: name of the object the cmdlet changed. */
  objectIds?: string[];
  ips?: string[];
  displayName?: string;
}

export interface AuditQuery {
  id: string;
  displayName?: string;
  status?: string;
  filterStartDateTime?: string;
  filterEndDateTime?: string;
  operationFilters?: string[];
  userPrincipalNameFilters?: string[];
  recordTypeFilters?: string[];
  keywordFilter?: string;
  objectIdFilters?: string[];
  ipAddressFilters?: string[];
  approximateReturnedRecordCount?: number;
  isRecordCountLimitExceeded?: boolean;
}

export interface AuditRecord {
  id: string;
  createdDateTime?: string;
  operation?: string;
  service?: string;
  userPrincipalName?: string;
  userId?: string;
  userType?: string;
  clientIp?: string;
  objectId?: string;
  auditLogRecordType?: string;
  auditData?: Record<string, unknown>;
}

/** Finished query with its records (raw Graph objects; shape with `shapeRecord`). */
export interface AuditCompleted {
  queryId: string;
  status: 'succeeded';
  records: AuditRecord[];
  truncated: boolean;
  recordCountLimitExceeded: boolean;
}

/** Query still running when the wait ran out; `next` re-reads it later. */
export interface AuditPending {
  queryId: string;
  status: string;
  next: string;
}

export interface AuditRunOptions {
  /** Max records to fetch (default 50; Infinity for all). */
  limit?: number;
  pollMs?: number;
  sleep?: (ms: number) => Promise<void>;
}

/** POST body for an audit query: only the filters that were given. */
export function auditQueryBody(filters: AuditFilters): Record<string, unknown> {
  const list = (values: string[] | undefined) => (values?.length ? values : undefined);
  return compact({
    '@odata.type': '#microsoft.graph.security.auditLogQuery',
    displayName: filters.displayName ?? `tenantctl audit ${new Date().toISOString()}`,
    filterStartDateTime: filters.start,
    filterEndDateTime: filters.end,
    recordTypeFilters: list(filters.recordTypes),
    keywordFilter: filters.keyword || undefined,
    operationFilters: list(filters.operations),
    userPrincipalNameFilters: list(filters.users),
    ipAddressFilters: list(filters.ips),
    objectIdFilters: list(filters.objectIds),
  });
}

export function resultsCommand(queryId: string): string {
  return `tenantctl purview audit results ${queryId}`;
}

async function fetchRecords(ctx: Context, queryId: string, limit: number): Promise<ListResult<AuditRecord>> {
  return ctx.graph.list<AuditRecord>(`${QUERIES}/${seg(queryId)}/records`, { limit, pageSize: false });
}

function failed(query: AuditQuery): never {
  throw new CliError('GRAPH_ERROR', `Audit query ${query.id} ended with status ${query.status}`, {
    hint: 'Narrow the window or filters and run the search again; Purview occasionally fails very broad searches.',
  });
}

/**
 * Wait for a query to finish (polling every ~5 s) and fetch its records, or return it as pending when
 * `waitSeconds` runs out. Throws when the query failed or was cancelled.
 */
export async function awaitAuditQuery(
  ctx: Context,
  queryId: string,
  waitSeconds: number,
  options: AuditRunOptions = {},
): Promise<AuditCompleted | AuditPending> {
  const pollMs = options.pollMs ?? POLL_MS;
  const sleep = options.sleep ?? ((ms: number) => Bun.sleep(ms));
  const deadline = Date.now() + Math.max(0, waitSeconds) * 1000;
  for (let polls = 0; ; polls++) {
    const query = await ctx.graph.get<AuditQuery>(`${QUERIES}/${seg(queryId)}`);
    const status = query.status ?? 'unknown';
    if (status === 'succeeded') {
      const { items, truncated } = await fetchRecords(ctx, queryId, options.limit ?? 50);
      return {
        queryId,
        status,
        records: items,
        truncated,
        recordCountLimitExceeded: query.isRecordCountLimitExceeded ?? false,
      };
    }
    if (FAILED_STATUSES[status]) failed({ ...query, id: queryId, status });
    if (Date.now() + pollMs > deadline) return { queryId, status, next: resultsCommand(queryId) };
    if (polls % 6 === 0) ctx.log(`Audit query ${queryId} is ${status}; waiting...`);
    await sleep(pollMs);
  }
}

/** Create an audit query from `filters` and wait up to `waitSeconds` for its records. Shared by every preset. */
export async function runAuditQuery(
  ctx: Context,
  filters: AuditFilters,
  waitSeconds: number,
  options: AuditRunOptions = {},
): Promise<AuditCompleted | AuditPending> {
  const created = await ctx.graph.post<AuditQuery>(QUERIES, auditQueryBody(filters));
  ctx.log(`Audit query ${created.id} created; audit searches usually take 1-5 minutes.`);
  return awaitAuditQuery(ctx, created.id, waitSeconds, options);
}

function str(value: unknown): string | undefined {
  return typeof value === 'string' && value ? value : undefined;
}

/** `Path` of an Exchange folder object inside auditData. */
function path(value: unknown): string | undefined {
  return typeof value === 'object' && value !== null && 'Path' in value ? str(value.Path) : undefined;
}

const LOGON_TYPES = ['owner', 'admin', 'delegate', 'transport', 'systemService'];

/** The few auditData fields an admin reads for common workloads; everything else needs --full. */
export function recordDetails(data: Record<string, unknown> | undefined): Record<string, unknown> | undefined {
  if (!data) return undefined;
  if (data.SiteUrl !== undefined || data.SourceFileName !== undefined) {
    return compact({
      file: str(data.SourceFileName),
      site: str(data.SiteUrl),
      folder: str(data.SourceRelativeUrl),
      destination: str(data.DestinationFileName),
    });
  }
  if (data.MailboxOwnerUPN !== undefined) {
    const folders = Array.isArray(data.Folders) ? data.Folders.map(path).filter(Boolean) : [];
    const affected = Array.isArray(data.AffectedItems) ? (data.AffectedItems as Record<string, unknown>[]) : [];
    const item = data.Item as Record<string, unknown> | undefined;
    const subjects = [...affected.map((a) => str(a.Subject)), str(item?.Subject)].filter(Boolean) as string[];
    if (!folders.length) {
      for (const a of affected) folders.push(path(a.ParentFolder));
      folders.push(path(data.Folder), path(item?.ParentFolder));
    }
    const properties = Array.isArray(data.OperationProperties)
      ? (data.OperationProperties as { Name?: string; Value?: unknown }[])
      : [];
    const logon = typeof data.LogonType === 'number' ? (LOGON_TYPES[data.LogonType] ?? data.LogonType) : undefined;
    return compact({
      mailbox: str(data.MailboxOwnerUPN),
      logonType: logon,
      external: data.ExternalAccess === true ? true : undefined,
      folders: [...new Set(folders.filter(Boolean))].slice(0, 5),
      subjects: subjects.length ? subjects.slice(0, 5) : undefined,
      accessType: str(properties.find((p) => p.Name === 'MailAccessType')?.Value),
      client: str(data.ClientInfoString),
    });
  }
  if (Array.isArray(data.Parameters)) {
    const parameters: Record<string, unknown> = {};
    for (const p of data.Parameters as { Name?: string; Value?: unknown }[]) {
      if (p.Name) parameters[p.Name] = p.Value;
    }
    return compact({ parameters, result: data.ResultStatus });
  }
  if (Array.isArray(data.Target) || Array.isArray(data.ModifiedProperties)) {
    const targets = Array.isArray(data.Target) ? (data.Target as { ID?: string; Type?: number }[]) : [];
    const modified = Array.isArray(data.ModifiedProperties)
      ? (data.ModifiedProperties as { Name?: string }[]).map((m) => m.Name).filter(Boolean)
      : [];
    return compact({
      targets: targets
        .map((t) => t.ID)
        .filter(Boolean)
        .slice(0, 5),
      modified: modified.length ? modified : undefined,
      result: str(data.ResultStatus),
    });
  }
  return undefined;
}

export function shapeRecord(record: AuditRecord) {
  return compact({
    time: record.createdDateTime ?? null,
    operation: record.operation ?? null,
    service: record.service ?? null,
    user: record.userPrincipalName ?? record.userId ?? null,
    ip: record.clientIp ?? null,
    objectId: record.objectId ?? null,
    auditLogRecordType: record.auditLogRecordType ?? null,
    details: recordDetails(record.auditData),
  });
}

/** Command output: a ListResult of shaped records, or the pending query with its follow-up command. */
function present(
  ctx: Context,
  run: AuditCompleted | AuditPending,
  keep: (record: AuditRecord) => boolean = () => true,
): ListResult | AuditPending {
  if (!('records' in run)) {
    ctx.log(`Still running; re-check with: ${run.next}`);
    return run;
  }
  if (run.recordCountLimitExceeded) {
    ctx.log('The search hit the Purview per-search record limit; narrow the window or filters for complete results.');
  }
  ctx.log(`Audit query ${run.queryId}: re-read these results with ${resultsCommand(run.queryId)}`);
  const records = run.records.filter(keep);
  return { items: ctx.full ? records : records.map(shapeRecord), truncated: run.truncated };
}

function window(args: { days: number; start?: string; end?: string }): { start: string; end: string } {
  const parse = (value: string, flag: string) => {
    const time = Date.parse(value);
    if (Number.isNaN(time))
      throw new CliError('USAGE', `--${flag} "${value}" is not a date`, {
        hint: 'Use ISO 8601, e.g. 2026-11-01 or 2026-11-01T08:00:00Z.',
      });
    return new Date(time).toISOString();
  };
  const end = args.end ? parse(args.end, 'end') : new Date().toISOString();
  const start = args.start ? parse(args.start, 'start') : daysAgo(args.days, Date.parse(end));
  if (Date.parse(start) >= Date.parse(end)) throw new CliError('USAGE', '--start must be before --end');
  return { start, end };
}

const WINDOW_ARGS = [
  { name: 'days', type: 'number', default: 7, description: 'Search the last N days (ignored when --start is given).' },
  { name: 'start', type: 'string', description: 'Window start (ISO date/time). Overrides --days.' },
  { name: 'end', type: 'string', description: 'Window end (ISO date/time). Default: now.' },
] as const;
const WAIT_ARG = {
  name: 'waitSeconds',
  type: 'number',
  default: 120,
  description: 'How long to wait for the search; if it is still running, the query id is returned to re-check.',
} as const;

const AUDIT_NOTE =
  'Searches run asynchronously in Purview (typically 1-5 minutes). If the search outlives --wait-seconds the ' +
  'result is { queryId, status, next } and `next` fetches the records later; nothing is lost. Records are kept ' +
  '180 days with Audit (Standard), one year for Entra/Exchange/SharePoint/OneDrive with Audit (Premium).';

const SEARCH_SUMMARY_NOTE =
  'details holds the useful auditData fields (file/site for SharePoint and OneDrive, mailbox/folders/subjects ' +
  'for Exchange mailbox activity, cmdlet parameters for Exchange admin, targets for Entra); --full returns raw ' +
  'records including all auditData.';

/** Exact operation names from the Purview audit activity reference. */
export const FILE_DELETED_OPERATIONS = [
  'FileDeleted',
  'FileDeletedFirstStageRecycleBin',
  'FileDeletedSecondStageRecycleBin',
  'FolderDeleted',
  'FolderDeletedFirstStageRecycleBin',
  'FolderDeletedSecondStageRecycleBin',
];
export const MAILBOX_ACCESS_OPERATIONS = [
  'MailItemsAccessed',
  'FolderBind',
  'SendAs',
  'SendOnBehalf',
  'MoveToDeletedItems',
  'SoftDelete',
  'HardDelete',
  'UpdateInboxRules',
  'UpdateFolderPermissions',
];
export const ADMIN_CHANGE_RECORD_TYPES = ['exchangeAdmin', 'azureActiveDirectory'];

function lower(value: string | null | undefined): string {
  return (value ?? '').toLowerCase();
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'purview',
    name: 'audit search',
    summary: 'Search the unified audit log: who did what, when, from where, across Microsoft 365',
    description:
      'Answers "who deleted this file", "who changed this mailbox", "what did Jane do yesterday". Combine ' +
      '--user (UPN of the person who acted), --operation (exact names such as FileDeleted, MailItemsAccessed, ' +
      'Set-Mailbox, Add member to group.), --record-type (exchangeAdmin, exchangeItem, sharePointFileOperation, ' +
      'oneDrive, azureActiveDirectory, ...), --object (full file URL or Exchange object name), --ip and --keyword ' +
      '(free text inside the workload-specific audit data, not operation names). ' +
      `${AUDIT_NOTE} ${SEARCH_SUMMARY_NOTE} Presets: \`audit file-deleted\`, \`audit mailbox-access\`, \`audit admin-changes\`.`,
    examples: [
      'tenantctl purview audit search --user jane@contoso.com --days 2',
      'tenantctl purview audit search --operation FileDeleted --object "https://contoso.sharepoint.com/sites/finance/Shared Documents/Q1.xlsx"',
      'tenantctl purview audit search --operation Set-Mailbox,Add-MailboxPermission --days 30',
    ],
    args: [
      { name: 'user', type: 'string[]', description: 'UPN(s) of the person who performed the activity.' },
      { name: 'operation', type: 'string[]', description: 'Operation name(s), e.g. FileDeleted, MailItemsAccessed.' },
      {
        name: 'recordType',
        type: 'string[]',
        description: 'Audit record type(s), e.g. exchangeAdmin, sharePointFileOperation.',
      },
      { name: 'keyword', type: 'string', description: 'Free text matched inside the workload-specific audit data.' },
      {
        name: 'object',
        type: 'string[]',
        noSplit: true,
        description: 'Full file/folder URL (SharePoint/OneDrive) or object name (Exchange admin). Repeat the flag.',
      },
      { name: 'ip', type: 'string[]', description: 'Client IP address(es).' },
      ...WINDOW_ARGS,
      WAIT_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: AUDIT_ALL,
    async run(ctx, args) {
      const run = await runAuditQuery(
        ctx,
        {
          ...window(args),
          users: args.user,
          operations: args.operation,
          recordTypes: args.recordType,
          keyword: args.keyword,
          objectIds: args.object,
          ips: args.ip,
        },
        args.waitSeconds,
        { limit: listLimit(args) },
      );
      return present(ctx, run);
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'audit results',
    summary: 'Status and records of an existing audit search',
    description:
      'Use the queryId returned when `audit search` (or a preset) outlived its wait, or one listed by `audit ' +
      `queries\`. Waits up to --wait-seconds if it is still running (0 = just check). ${SEARCH_SUMMARY_NOTE}`,
    examples: ['tenantctl purview audit results 168ec429-084b-a489-90d8-504a87846305'],
    args: [
      { name: 'queryId', type: 'string', required: true, positional: true, description: 'Audit query id.' },
      { ...WAIT_ARG, default: 0 },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: AUDIT_ALL,
    async run(ctx, args) {
      return present(ctx, await awaitAuditQuery(ctx, args.queryId, args.waitSeconds, { limit: listLimit(args) }));
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'audit queries',
    summary: 'Recent audit searches with their status and filters',
    description:
      'Lists audit log queries created in this tenant (by anyone using the API), so an earlier search can be ' +
      're-read with `audit results <queryId>` instead of re-running it.',
    examples: ['tenantctl purview audit queries'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: AUDIT_ALL,
    async run(ctx, args) {
      const { items, truncated } = await ctx.graph.list<AuditQuery>(QUERIES, {
        limit: listLimit(args),
        pageSize: false,
      });
      if (ctx.full) return { items, truncated };
      return {
        items: items.map((q) =>
          compact({
            queryId: q.id,
            name: q.displayName ?? null,
            status: q.status ?? null,
            start: q.filterStartDateTime ?? null,
            end: q.filterEndDateTime ?? null,
            operations: q.operationFilters?.length ? q.operationFilters : undefined,
            users: q.userPrincipalNameFilters?.length ? q.userPrincipalNameFilters : undefined,
            recordTypes: q.recordTypeFilters?.length ? q.recordTypeFilters : undefined,
            keyword: q.keywordFilter || undefined,
            approximateRecords: q.approximateReturnedRecordCount,
          }),
        ),
        truncated,
      };
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'audit file-deleted',
    summary: 'Who deleted files on a SharePoint site or OneDrive (including recycle bin purges)',
    description:
      `Preset over \`audit search\` with operations ${FILE_DELETED_OPERATIONS.join(', ')}. The site URL (and ` +
      '--file name) is sent as the keyword and records are then kept only when their site matches, so the ' +
      'result is exact. FirstStageRecycleBin = emptied from the site recycle bin; SecondStage = purged from the ' +
      `site collection recycle bin (unrecoverable after that). ${AUDIT_NOTE}`,
    examples: [
      'tenantctl purview audit file-deleted --site https://contoso.sharepoint.com/sites/finance --file Q1.xlsx',
      'tenantctl purview audit file-deleted --site https://contoso-my.sharepoint.com/personal/jane_contoso_com --days 30',
    ],
    args: [
      { name: 'site', type: 'string', required: true, description: 'Site or OneDrive URL, or site name.' },
      { name: 'file', type: 'string', description: 'Only files/folders with this name.' },
      { name: 'user', type: 'string[]', description: 'Only deletions by these UPN(s).' },
      ...WINDOW_ARGS,
      WAIT_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['AuditLogsQuery-SharePoint.Read.All', 'AuditLogsQuery-OneDrive.Read.All', 'Sites.Read.All'],
      application: ['AuditLogsQuery-SharePoint.Read.All', 'AuditLogsQuery-OneDrive.Read.All', 'Sites.Read.All'],
      roles: AUDIT_ALL.roles,
    },
    async run(ctx, args) {
      const site = await resolveSite(ctx, args.site);
      const siteUrl = lower(site.webUrl).replace(/\/+$/, '');
      const file = args.file ? lower(args.file) : undefined;
      const run = await runAuditQuery(
        ctx,
        {
          ...window(args),
          operations: FILE_DELETED_OPERATIONS,
          users: args.user,
          keyword: args.file ?? site.webUrl,
          displayName: `tenantctl file-deleted ${site.webUrl}`,
        },
        args.waitSeconds,
        { limit: listLimit(args) },
      );
      return present(ctx, run, (record) => {
        const data = record.auditData ?? {};
        const recordSite = lower(str(data.SiteUrl)).replace(/\/+$/, '');
        const onSite = recordSite ? recordSite === siteUrl : lower(record.objectId).startsWith(`${siteUrl}/`);
        return onSite && (!file || lower(str(data.SourceFileName)) === file);
      });
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'audit mailbox-access',
    summary: 'Who opened, sent from or deleted items in a mailbox (owner, delegates, admins, apps)',
    description:
      `Preset over \`audit search\` with mailbox audit operations ${MAILBOX_ACCESS_OPERATIONS.join(', ')}. ` +
      'Mailbox audit records are filed under the person who acted, so the mailbox owner is matched through the ' +
      'audit data (MailboxOwnerUPN), not --user. --accessed-by narrows to one actor; --non-owner hides the ' +
      "owner's own activity (the usual question after a compromise or a delegate dispute). details.logonType is " +
      'owner, delegate or admin; MailItemsAccessed accessType is Bind (one item opened) or Sync (folder synced ' +
      `by a client). ${AUDIT_NOTE}`,
    examples: [
      'tenantctl purview audit mailbox-access ceo@contoso.com --non-owner --days 14',
      'tenantctl purview audit mailbox-access shared-finance@contoso.com --accessed-by bob@contoso.com',
    ],
    args: [
      {
        name: 'mailbox',
        type: 'string',
        required: true,
        positional: true,
        description: 'Mailbox owner: UPN, email, name or id.',
      },
      { name: 'accessedBy', type: 'string[]', description: 'Only activity by these UPN(s).' },
      { name: 'nonOwner', type: 'boolean', description: "Hide the mailbox owner's own activity." },
      ...WINDOW_ARGS,
      WAIT_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['AuditLogsQuery-Exchange.Read.All', 'User.ReadBasic.All'],
      application: ['AuditLogsQuery-Exchange.Read.All', 'User.Read.All'],
      roles: AUDIT_ALL.roles,
    },
    async run(ctx, args) {
      const owner = await resolveUser(ctx, args.mailbox);
      const upn = lower(owner.userPrincipalName);
      const run = await runAuditQuery(
        ctx,
        {
          ...window(args),
          operations: MAILBOX_ACCESS_OPERATIONS,
          users: args.accessedBy,
          keyword: owner.userPrincipalName,
          displayName: `tenantctl mailbox-access ${owner.userPrincipalName}`,
        },
        args.waitSeconds,
        { limit: listLimit(args) },
      );
      return present(ctx, run, (record) => {
        if (lower(str(record.auditData?.MailboxOwnerUPN)) !== upn) return false;
        return !args.nonOwner || lower(record.userPrincipalName ?? record.userId) !== upn;
      });
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'audit admin-changes',
    summary: 'Recent Exchange admin cmdlets and Entra directory changes, optionally by one admin',
    description:
      `Preset over \`audit search\` with record types ${ADMIN_CHANGE_RECORD_TYPES.join(', ')}: every ` +
      'Exchange admin cmdlet (Set-Mailbox, Add-MailboxPermission, New-TransportRule...) with its parameters, ' +
      'and Entra changes (users, groups, roles, apps). Answers "who changed this", "what did this admin do". ' +
      `Use --operation to narrow to specific cmdlets/activities. ${AUDIT_NOTE}`,
    examples: [
      'tenantctl purview audit admin-changes --days 3',
      'tenantctl purview audit admin-changes --user admin@contoso.com --operation Set-Mailbox',
    ],
    args: [
      { name: 'user', type: 'string[]', description: 'Only changes by these admin UPN(s).' },
      { name: 'operation', type: 'string[]', description: 'Only these cmdlets/activities.' },
      ...WINDOW_ARGS,
      WAIT_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['AuditLogsQuery-Exchange.Read.All', 'AuditLogsQuery-Entra.Read.All'],
      application: ['AuditLogsQuery-Exchange.Read.All', 'AuditLogsQuery-Entra.Read.All'],
      roles: AUDIT_ALL.roles,
    },
    async run(ctx, args) {
      const run = await runAuditQuery(
        ctx,
        {
          ...window(args),
          recordTypes: ADMIN_CHANGE_RECORD_TYPES,
          users: args.user,
          operations: args.operation,
          displayName: 'tenantctl admin-changes',
        },
        args.waitSeconds,
        { limit: listLimit(args) },
      );
      return present(ctx, run);
    },
  }),
];
