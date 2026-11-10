import { describe, expect, test } from 'bun:test';
import { commands, runAuditQuery } from '../../src/domains/purview/index.ts';
import { commandById, type MockRequest, type MockRoute, mockContext, page, runCommand } from '../helpers/graphMock.ts';

const QUERY_ID = 'q-1';
const CASE_ID = '11111111-2222-3333-4444-555555555555';

function auditRoutes(
  statuses: string[],
  records: Record<string, unknown>[][] = [[]],
  created: MockRequest[] = [],
): MockRoute[] {
  let poll = 0;
  return [
    {
      method: 'POST',
      path: '/security/auditLog/queries',
      reply: (req) => {
        created.push(req);
        return { status: 201, body: { id: QUERY_ID, status: 'notStarted' } };
      },
    },
    {
      path: `/security/auditLog/queries/${QUERY_ID}`,
      reply: () => ({ body: { id: QUERY_ID, status: statuses[Math.min(poll++, statuses.length - 1)] } }),
    },
    {
      path: `/security/auditLog/queries/${QUERY_ID}/records`,
      reply: (req) => {
        const index = Number(req.query.page ?? 0);
        const next =
          index + 1 < records.length
            ? `https://graph.test/v1.0/security/auditLog/queries/${QUERY_ID}/records?page=${index + 1}`
            : undefined;
        return page(records[index] ?? [], next);
      },
    },
  ];
}

describe('audit search', () => {
  const command = commandById(commands, 'purview audit search');

  test('builds the query body from flags and returns shaped records', async () => {
    const created: MockRequest[] = [];
    const record = {
      id: 'r1',
      createdDateTime: '2026-11-09T10:00:00Z',
      operation: 'FileDeleted',
      service: 'SharePoint',
      userPrincipalName: 'bob@contoso.com',
      clientIp: '203.0.113.5',
      objectId: 'https://contoso.sharepoint.com/sites/finance/Shared Documents/Q1.xlsx',
      auditLogRecordType: 'sharePointFileOperation',
      auditData: {
        SourceFileName: 'Q1.xlsx',
        SiteUrl: 'https://contoso.sharepoint.com/sites/finance/',
        SourceRelativeUrl: 'Shared Documents',
        Noise: 'x',
      },
    };
    const { outcome } = await runCommand(
      command,
      [
        '--user',
        'bob@contoso.com,jane@contoso.com',
        '--operation',
        'FileDeleted',
        '--record-type',
        'sharePointFileOperation',
        '--keyword',
        'Q1',
        '--object',
        'https://contoso.sharepoint.com/sites/finance/Shared Documents/Q1, final.xlsx',
        '--ip',
        '203.0.113.5',
        '--start',
        '2026-11-01',
        '--end',
        '2026-11-08',
      ],
      auditRoutes(['succeeded'], [[record]], created),
    );
    expect(created[0]?.body).toMatchObject({
      filterStartDateTime: '2026-11-01T00:00:00.000Z',
      filterEndDateTime: '2026-11-08T00:00:00.000Z',
      userPrincipalNameFilters: ['bob@contoso.com', 'jane@contoso.com'],
      operationFilters: ['FileDeleted'],
      recordTypeFilters: ['sharePointFileOperation'],
      keywordFilter: 'Q1',
      objectIdFilters: ['https://contoso.sharepoint.com/sites/finance/Shared Documents/Q1, final.xlsx'],
      ipAddressFilters: ['203.0.113.5'],
    });
    expect(outcome).toEqual({
      kind: 'result',
      result: {
        items: [
          {
            time: '2026-11-09T10:00:00Z',
            operation: 'FileDeleted',
            service: 'SharePoint',
            user: 'bob@contoso.com',
            ip: '203.0.113.5',
            objectId: record.objectId,
            auditLogRecordType: 'sharePointFileOperation',
            details: {
              file: 'Q1.xlsx',
              site: 'https://contoso.sharepoint.com/sites/finance/',
              folder: 'Shared Documents',
            },
          },
        ],
        truncated: false,
      },
    });
  });

  test('--days sets a window ending now and omits unused filters', async () => {
    const created: MockRequest[] = [];
    await runCommand(command, ['--days', '2'], auditRoutes(['succeeded'], [[]], created));
    const body = created[0]?.body as Record<string, string>;
    const span = Date.parse(body.filterEndDateTime ?? '') - Date.parse(body.filterStartDateTime ?? '');
    expect(span).toBe(2 * 86_400_000);
    expect(Math.abs(Date.parse(body.filterEndDateTime ?? '') - Date.now())).toBeLessThan(60_000);
    expect(Object.keys(body).filter((k) => k.endsWith('Filters') || k === 'keywordFilter')).toEqual([]);
  });

  test('a search still running after --wait-seconds returns the query id and the follow-up command', async () => {
    const { outcome } = await runCommand(
      command,
      ['--operation', 'MailItemsAccessed', '--wait-seconds', '0'],
      auditRoutes(['running']),
    );
    expect(outcome).toEqual({
      kind: 'result',
      result: { queryId: QUERY_ID, status: 'running', next: `m365 purview audit results ${QUERY_ID}` },
    });
  });

  test('a failed search is an error with a hint', async () => {
    const error = await runCommand(command, [], auditRoutes(['failed'])).catch((e: unknown) => e);
    expect((error as Error).message).toContain('failed');
  });
});

describe('runAuditQuery', () => {
  test('polls until the query succeeds, then pages through the records', async () => {
    const records = [[{ id: 'a' }, { id: 'b' }], [{ id: 'c' }]];
    const { ctx, calls, logs } = mockContext(auditRoutes(['notStarted', 'running', 'succeeded'], records));
    const sleeps: number[] = [];
    const run = await runAuditQuery(
      ctx,
      { start: '2026-11-01T00:00:00Z', end: '2026-11-02T00:00:00Z', operations: ['FileDeleted'] },
      60,
      { limit: Number.POSITIVE_INFINITY, sleep: async (ms) => void sleeps.push(ms) },
    );
    expect(run).toMatchObject({ queryId: QUERY_ID, status: 'succeeded', truncated: false });
    expect('records' in run && run.records.map((r) => r.id)).toEqual(['a', 'b', 'c']);
    expect(sleeps).toEqual([5000, 5000]);
    expect(calls.filter((c) => c.path === `/security/auditLog/queries/${QUERY_ID}`)).toHaveLength(3);
    expect(logs.some((l) => l.includes(QUERY_ID))).toBe(true);
  });

  test('stops at the record limit and reports truncation', async () => {
    const { ctx } = mockContext(auditRoutes(['succeeded'], [[{ id: 'a' }, { id: 'b' }], [{ id: 'c' }]]));
    const run = await runAuditQuery(ctx, { start: 'x', end: 'y' }, 0, { limit: 2 });
    expect(run).toMatchObject({ truncated: true });
    expect('records' in run && run.records).toHaveLength(2);
  });
});

describe('audit results / queries', () => {
  test('results re-reads an existing query', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'purview audit results'),
      [QUERY_ID],
      auditRoutes(
        ['succeeded'],
        [[{ id: 'r', operation: 'Set-Mailbox', auditData: { Parameters: [{ Name: 'Identity', Value: 'bob' }] } }]],
      ),
    );
    expect(calls.some((c) => c.method === 'POST')).toBe(false);
    expect(outcome).toMatchObject({
      result: { items: [{ operation: 'Set-Mailbox', details: { parameters: { Identity: 'bob' } } }] },
    });
  });

  test('queries lists status and filters', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'purview audit queries'),
      [],
      [
        {
          path: '/security/auditLog/queries',
          reply: page([
            {
              id: 'q',
              displayName: 'n',
              status: 'running',
              operationFilters: ['FileDeleted'],
              userPrincipalNameFilters: [],
            },
          ]),
        },
      ],
    );
    expect(outcome).toMatchObject({
      result: { items: [{ queryId: 'q', status: 'running', operations: ['FileDeleted'] }] },
    });
    expect(JSON.stringify(outcome)).not.toContain('users');
  });
});

describe('audit presets', () => {
  test('file-deleted searches the delete operations on the resolved site and keeps only that site', async () => {
    const created: MockRequest[] = [];
    const site = { id: 'contoso.sharepoint.com,a,b', webUrl: 'https://contoso.sharepoint.com/sites/finance' };
    const records = [
      {
        id: '1',
        operation: 'FileDeleted',
        auditData: { SiteUrl: 'https://contoso.sharepoint.com/sites/finance/', SourceFileName: 'Q1.xlsx' },
      },
      {
        id: '2',
        operation: 'FileDeleted',
        auditData: { SiteUrl: 'https://contoso.sharepoint.com/sites/finance-archive/', SourceFileName: 'Q1.xlsx' },
      },
    ];
    const { outcome } = await runCommand(
      commandById(commands, 'purview audit file-deleted'),
      ['--site', 'https://contoso.sharepoint.com/sites/finance/Shared Documents', '--file', 'Q1.xlsx'],
      [
        { path: '/sites/contoso.sharepoint.com:/sites/finance', reply: { body: site } },
        ...auditRoutes(['succeeded'], [records], created),
      ],
    );
    expect(created[0]?.body).toMatchObject({
      operationFilters: [
        'FileDeleted',
        'FileDeletedFirstStageRecycleBin',
        'FileDeletedSecondStageRecycleBin',
        'FolderDeleted',
        'FolderDeletedFirstStageRecycleBin',
        'FolderDeletedSecondStageRecycleBin',
      ],
      keywordFilter: 'Q1.xlsx',
    });
    expect(outcome).toMatchObject({
      result: { items: [{ details: { site: 'https://contoso.sharepoint.com/sites/finance/' } }] },
    });
    expect((outcome as { result: { items: unknown[] } }).result.items).toHaveLength(1);
  });

  test('mailbox-access matches the mailbox owner in the audit data and can hide the owner', async () => {
    const created: MockRequest[] = [];
    const owner = { id: 'u1', displayName: 'CEO', userPrincipalName: 'ceo@contoso.com', mail: 'ceo@contoso.com' };
    const records = [
      {
        id: '1',
        operation: 'MailItemsAccessed',
        userPrincipalName: 'ceo@contoso.com',
        auditData: { MailboxOwnerUPN: 'ceo@contoso.com', LogonType: 0 },
      },
      {
        id: '2',
        operation: 'MailItemsAccessed',
        userPrincipalName: 'eve@contoso.com',
        auditData: {
          MailboxOwnerUPN: 'ceo@contoso.com',
          LogonType: 2,
          Folders: [{ Path: '\\Inbox' }],
          OperationProperties: [{ Name: 'MailAccessType', Value: 'Bind' }],
        },
      },
      {
        id: '3',
        operation: 'SendAs',
        userPrincipalName: 'eve@contoso.com',
        auditData: { MailboxOwnerUPN: 'other@contoso.com' },
      },
    ];
    const { outcome } = await runCommand(
      commandById(commands, 'purview audit mailbox-access'),
      ['ceo@contoso.com', '--non-owner'],
      [{ path: '/users/ceo@contoso.com', reply: { body: owner } }, ...auditRoutes(['succeeded'], [records], created)],
    );
    expect(created[0]?.body).toMatchObject({ keywordFilter: 'ceo@contoso.com' });
    expect((created[0]?.body as { operationFilters: string[] }).operationFilters).toContain('MailItemsAccessed');
    expect((created[0]?.body as Record<string, unknown>).userPrincipalNameFilters).toBeUndefined();
    expect(outcome).toEqual({
      kind: 'result',
      result: {
        items: [
          {
            time: null,
            operation: 'MailItemsAccessed',
            service: null,
            user: 'eve@contoso.com',
            ip: null,
            objectId: null,
            auditLogRecordType: null,
            details: { mailbox: 'ceo@contoso.com', logonType: 'delegate', folders: ['\\Inbox'], accessType: 'Bind' },
          },
        ],
        truncated: false,
      },
    });
  });

  test('admin-changes filters Exchange admin and Entra record types by admin', async () => {
    const created: MockRequest[] = [];
    await runCommand(
      commandById(commands, 'purview audit admin-changes'),
      ['--user', 'admin@contoso.com'],
      auditRoutes(['succeeded'], [[]], created),
    );
    expect(created[0]?.body).toMatchObject({
      recordTypeFilters: ['exchangeAdmin', 'azureActiveDirectory'],
      userPrincipalNameFilters: ['admin@contoso.com'],
    });
  });
});

describe('ediscovery', () => {
  const kase = { id: CASE_ID, displayName: 'HR-2026-014', status: 'active' };
  const bob = { id: 'u-bob', displayName: 'Bob', userPrincipalName: 'bob@contoso.com', mail: 'bob@contoso.com' };
  const jane = { id: 'u-jane', displayName: 'Jane', userPrincipalName: 'jane@contoso.com', mail: 'jane@contoso.com' };
  const base = `/security/cases/ediscoveryCases/${CASE_ID}/custodians`;
  const routes = (): MockRoute[] => [
    { path: '/security/cases/ediscoveryCases', reply: page([kase, { id: 'x', displayName: 'Other' }]) },
    { path: '/users/bob@contoso.com', reply: { body: bob } },
    { path: '/users/jane@contoso.com', reply: { body: jane } },
    {
      path: base,
      reply: page([{ id: 'c-jane', email: 'jane@contoso.com', holdStatus: 'notApplied', status: 'active' }]),
    },
    { path: `${base}/c-jane/userSources`, reply: page([{ id: 's1', email: 'jane@contoso.com' }]) },
    { method: 'POST', path: base, reply: { status: 201, body: { id: 'c-bob', email: 'bob@contoso.com' } } },
    { method: 'POST', path: `${base}/c-bob/userSources`, reply: { status: 201, body: { id: 's2' } } },
    { method: 'POST', path: `${base}/applyHold`, reply: { status: 202 } },
    { method: 'POST', path: `${base}/c-jane/removeHold`, reply: { status: 202 } },
  ];

  test('hold previews without changing anything, then adds custodians, sources and applies the hold', async () => {
    const command = commandById(commands, 'purview ediscovery hold');
    const argv = ['hr-2026', 'bob@contoso.com', 'jane@contoso.com'];
    const first = await runCommand(command, argv, routes(), { allowWrites: true });
    expect(first.outcome).toMatchObject({
      kind: 'confirm',
      preview: {
        case: { id: CASE_ID, name: 'HR-2026-014' },
        people: [
          { user: 'bob@contoso.com', alreadyCustodian: false },
          { user: 'jane@contoso.com', alreadyCustodian: true, currentHoldStatus: 'notApplied' },
        ],
      },
    });
    expect(JSON.stringify(first.outcome)).toContain('on hold in eDiscovery case');
    expect(first.calls.some((c) => c.method === 'POST')).toBe(false);

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(command, argv, routes(), { allowWrites: true, confirm: token });
    const posts = second.calls.filter((c) => c.method === 'POST');
    expect(posts.map((c) => [c.path, c.body])).toEqual([
      [base, { email: 'bob@contoso.com' }],
      [`${base}/c-bob/userSources`, { email: 'bob@contoso.com', includedSources: 'mailbox,site' }],
      [`${base}/applyHold`, { ids: ['c-bob', 'c-jane'] }],
    ]);
    expect(second.outcome).toMatchObject({
      result: { holdStatus: 'applying', people: [{ addedToCase: true }, { addedToCase: false }] },
    });
  });

  test('release removes the hold from that custodian only after confirmation', async () => {
    const command = commandById(commands, 'purview ediscovery release');
    const first = await runCommand(
      command,
      [CASE_ID, 'jane@contoso.com'],
      [{ path: `/security/cases/ediscoveryCases/${CASE_ID}`, reply: { body: kase } }, ...routes()],
      { allowWrites: true },
    );
    expect(first.outcome).toMatchObject({ kind: 'confirm', preview: { custodian: { id: 'c-jane' } } });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(
      command,
      [CASE_ID, 'jane@contoso.com'],
      [{ path: `/security/cases/ediscoveryCases/${CASE_ID}`, reply: { body: kase } }, ...routes()],
      { allowWrites: true, confirm: token },
    );
    expect(second.calls.filter((c) => c.method === 'POST').map((c) => c.path)).toEqual([`${base}/c-jane/removeHold`]);
  });

  test('a name matching several cases is ambiguous', async () => {
    const error = await runCommand(
      commandById(commands, 'purview ediscovery custodians'),
      ['o'],
      [
        {
          path: '/security/cases/ediscoveryCases',
          reply: page([
            { id: 'a', displayName: 'Ops' },
            { id: 'b', displayName: 'Other' },
          ]),
        },
      ],
    ).catch((e: unknown) => e);
    expect((error as { code: string }).code).toBe('AMBIGUOUS');
  });

  test('create posts the case name and description', async () => {
    const { calls } = await runCommand(
      commandById(commands, 'purview ediscovery create'),
      ['--name', 'Case A', '--description', 'd'],
      [
        {
          method: 'POST',
          path: '/security/cases/ediscoveryCases',
          reply: { status: 201, body: { id: 'n', displayName: 'Case A' } },
        },
      ],
      { allowWrites: true },
    );
    expect(calls.find((c) => c.method === 'POST')?.body).toEqual({ displayName: 'Case A', description: 'd' });
  });
});
