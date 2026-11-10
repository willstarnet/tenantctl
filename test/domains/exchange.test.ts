import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import type { Outcome } from '../../src/core/execute.ts';
import {
  buildTraceFilter,
  commands,
  type ShapedInboxRule,
  shapeInboxRule,
  shapeMailboxUsage,
} from '../../src/domains/exchange/index.ts';
import { commandById, fakeExo, graphError, page, runCommand, type MockRoute } from '../helpers/graphMock.ts';

const JANE = { id: 'u1', displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com', mail: 'jane@contoso.com' };
const janeRoute: MockRoute = { path: '/users/jane@contoso.com', reply: { body: JANE } };
const domainsRoute: MockRoute = {
  path: '/domains',
  reply: page([
    { id: 'contoso.com', isVerified: true },
    { id: 'contoso.onmicrosoft.com', isVerified: true },
    { id: 'pending.example', isVerified: false },
  ]),
};

const MALICIOUS_RULE = {
  id: 'r1',
  displayName: '.',
  sequence: 1,
  isEnabled: true,
  conditions: { bodyOrSubjectContains: ['invoice', 'payment'] },
  actions: { moveToFolder: 'f-rss', markAsRead: true, stopProcessingRules: true },
};
const FORWARD_RULE = {
  id: 'r2',
  displayName: 'Copy to personal',
  sequence: 2,
  isEnabled: true,
  conditions: {},
  actions: { forwardTo: [{ emailAddress: { name: 'Jane', address: 'jane.doe@gmail.com' } }] },
};
const INTERNAL_RULE = {
  id: 'r3',
  displayName: 'To assistant',
  sequence: 3,
  isEnabled: true,
  conditions: { fromAddresses: [{ emailAddress: { address: 'ceo@contoso.com' } }] },
  actions: { redirectTo: [{ emailAddress: { address: 'pa@mail.contoso.com' } }] },
};

const rulesRoute = (rules: unknown[]): MockRoute => ({
  path: '/users/u1/mailFolders/inbox/messageRules',
  reply: page(rules),
});
const rssFolder: MockRoute = { path: '/users/u1/mailFolders/f-rss', reply: { body: { displayName: 'RSS Feeds' } } };

async function run(id: string, argv: string[], routes: MockRoute[], options = {}) {
  return runCommand(commandById(commands, id), argv, routes, options);
}

/** The command result, typed as the shape the test asserts on. */
function result<T = unknown>(outcome: Outcome): T {
  if (outcome.kind !== 'result') throw new Error(`expected result, got ${outcome.kind}`);
  return outcome.result as T;
}

describe('inbox rule heuristics', () => {
  test('external forwarding is detected against verified tenant domains, subdomains count as internal', () => {
    const domains = ['contoso.com', 'contoso.onmicrosoft.com'];
    const external = shapeInboxRule(FORWARD_RULE, { tenantDomains: domains });
    expect(external).toMatchObject({
      forwards: true,
      forwardsExternally: true,
      externalTargets: ['jane.doe@gmail.com'],
      suspicious: true,
      conditions: ['all messages'],
      actions: ['forward to jane.doe@gmail.com'],
    });
    expect(external.reason).toContain('forwards to external jane.doe@gmail.com');

    const internal = shapeInboxRule(INTERNAL_RULE, { tenantDomains: domains });
    expect(internal).toMatchObject({ forwards: true, forwardsExternally: false, suspicious: false, reason: null });
  });

  test('without tenant domains external forwarding is unknown, not guessed', () => {
    expect(shapeInboxRule(FORWARD_RULE)).toMatchObject({
      forwards: true,
      forwardsExternally: null,
      externalTargets: null,
      suspicious: false,
    });
  });

  test('move to RSS + mark read + payment keywords + meaningless name is flagged with every reason', () => {
    const shaped = shapeInboxRule(MALICIOUS_RULE, { folderNames: { 'f-rss': 'RSS Feeds' } });
    expect(shaped).toMatchObject({ movesToFolder: 'RSS Feeds', marksRead: true, deletes: false, suspicious: true });
    expect(shaped.reason).toContain('moves to "RSS Feeds" and marks read');
    expect(shaped.reason).toContain('mentioning "invoice"');
    expect(shaped.reason).toContain('meaningless rule name');
    expect(shaped.conditions).toEqual(['bodyOrSubjectContains: "invoice", "payment"']);
  });

  test('ordinary filing rules are not suspicious', () => {
    const shaped = shapeInboxRule(
      {
        id: 'r4',
        displayName: 'Newsletters',
        isEnabled: true,
        conditions: { senderContains: ['news'] },
        actions: { moveToFolder: 'f1', markAsRead: true },
      },
      { folderNames: { f1: 'Newsletters' } },
    );
    expect(shaped).toMatchObject({
      suspicious: false,
      movesToFolder: 'Newsletters',
      actions: ['move to "Newsletters"', 'mark as read'],
    });
  });

  test('delete + mark read is flagged', () => {
    const shaped = shapeInboxRule({ id: 'r5', displayName: 'cleanup', actions: { delete: true, markAsRead: true } });
    expect(shaped.deletes).toBe(true);
    expect(shaped.reason).toContain('deletes and marks read');
  });
});

describe('exchange rules', () => {
  test('rules list resolves folder names and tenant domains', async () => {
    const { outcome, calls } = await run(
      'exchange rules list',
      ['jane@contoso.com'],
      [janeRoute, domainsRoute, rulesRoute([FORWARD_RULE, MALICIOUS_RULE]), rssFolder],
    );
    const { items } = result<{ items: ShapedInboxRule[] }>(outcome);
    expect(items.map((r) => r.id)).toEqual(['r1', 'r2']);
    expect(items[0]).toMatchObject({ movesToFolder: 'RSS Feeds', suspicious: true });
    expect(items[1]).toMatchObject({ forwardsExternally: true });
    expect(calls.some((c) => c.path === '/users/u1/mailFolders/f-rss')).toBe(true);
  });

  test('rules audit reads mailboxes via $batch and returns only flagged mailboxes and errors', async () => {
    const users = [
      JANE,
      { id: 'u2', displayName: 'Bob', userPrincipalName: 'bob@contoso.com' },
      { id: 'u3', displayName: 'Room', userPrincipalName: 'room@contoso.com' },
    ];
    const sub: Record<string, { status: number; body: unknown }> = {
      '/users/u1/mailFolders/inbox/messageRules': { status: 200, body: { value: [FORWARD_RULE] } },
      '/users/u2/mailFolders/inbox/messageRules': {
        status: 200,
        body: { value: [{ id: 'x', displayName: 'Filing', actions: { moveToFolder: 'f' } }] },
      },
      '/users/u2/mailFolders/f': { status: 200, body: { displayName: 'Projects' } },
      '/users/u3/mailFolders/inbox/messageRules': {
        status: 404,
        body: { error: { code: 'MailboxNotEnabledForRESTAPI', message: 'no mailbox' } },
      },
    };
    const { outcome, calls } = await run(
      'exchange rules audit',
      [],
      [
        { path: '/users', reply: page(users) },
        domainsRoute,
        {
          method: 'POST',
          path: '/$batch',
          reply: (req) => ({
            body: {
              responses: (req.body as { requests: { id: string; url: string }[] }).requests.map((r) => ({
                id: r.id,
                ...(sub[r.url.split('?')[0] as string] ?? { status: 404, body: {} }),
              })),
            },
          }),
        },
      ],
    );
    const audit = result<{ mailboxes: unknown[]; errors: unknown[] }>(outcome);
    expect(audit).toMatchObject({ scanned: 3, flaggedCount: 1, suspiciousCount: 1 });
    expect(audit.mailboxes[0]).toMatchObject({
      user: 'jane@contoso.com',
      rules: [{ id: 'r2', forwardsExternally: true }],
    });
    expect(audit.errors).toEqual([
      expect.objectContaining({ user: 'room@contoso.com', code: 'MailboxNotEnabledForRESTAPI' }),
    ]);
    const userQuery = calls.find((c) => c.path === '/users');
    expect(userQuery?.query.$filter).toBe("assignedLicenses/$count ne 0 and userType eq 'Member'");
    expect(userQuery?.headers.ConsistencyLevel).toBe('eventual');
    expect(userQuery?.query.$top).toBe('100');
    // One $batch for the three mailboxes' rules, one for the folder name; no per-mailbox GETs.
    expect(calls.filter((c) => c.path === '/$batch')).toHaveLength(2);
    expect(calls.some((c) => c.method === 'GET' && c.path.includes('/messageRules'))).toBe(false);
  });

  test('rules delete previews the rule by name, then deletes it once confirmed', async () => {
    const routes = [
      janeRoute,
      rulesRoute([MALICIOUS_RULE, FORWARD_RULE]),
      rssFolder,
      { method: 'DELETE', path: '/users/u1/mailFolders/inbox/messageRules/r2', reply: { status: 204 } },
    ];
    const first = await run('exchange rules delete', ['jane@contoso.com', 'copy to personal'], routes, {
      allowWrites: true,
    });
    expect(first.outcome.kind).toBe('confirm');
    if (first.outcome.kind !== 'confirm') return;
    expect(first.outcome.preview).toMatchObject({
      user: 'jane@contoso.com',
      rule: { id: 'r2', name: 'Copy to personal' },
    });
    expect(first.calls.some((c) => c.method === 'DELETE')).toBe(false);

    const second = await run('exchange rules delete', ['jane@contoso.com', 'copy to personal'], routes, {
      allowWrites: true,
      confirm: first.outcome.token,
    });
    expect(result(second.outcome)).toMatchObject({ deleted: true });
    expect(second.calls.filter((c) => c.method === 'DELETE').map((c) => c.path)).toEqual([
      '/users/u1/mailFolders/inbox/messageRules/r2',
    ]);
  });

  test('a rule name matching several rules is AMBIGUOUS', async () => {
    const error = (await run(
      'exchange rules disable',
      ['jane@contoso.com', 'to'],
      [janeRoute, rulesRoute([FORWARD_RULE, INTERNAL_RULE])],
      { allowWrites: true },
    ).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
  });

  test('rules disable patches isEnabled false', async () => {
    const { calls } = await run(
      'exchange rules disable',
      ['jane@contoso.com', 'r3'],
      [
        janeRoute,
        rulesRoute([FORWARD_RULE, INTERNAL_RULE]),
        { method: 'PATCH', path: '/users/u1/mailFolders/inbox/messageRules/r3', reply: { body: {} } },
      ],
      { allowWrites: true },
    );
    const patch = calls.find((c) => c.method === 'PATCH');
    expect(patch?.body).toEqual({ isEnabled: false });
  });

  test('forwarding combines inbox rules with mailbox-level forwarding from Exchange', async () => {
    const exo = fakeExo(() => [
      { ForwardingSmtpAddress: 'smtp:jane@gmail.com', ForwardingAddress: null, DeliverToMailboxAndForward: false },
    ]);
    const { outcome } = await run(
      'exchange forwarding',
      ['jane@contoso.com'],
      [janeRoute, domainsRoute, rulesRoute([FORWARD_RULE, MALICIOUS_RULE]), rssFolder],
      { exo },
    );
    const out = result<{ inboxRules: { id: string }[] }>(outcome);
    expect(out).toMatchObject({
      forwardsViaInboxRules: true,
      forwardsViaMailbox: true,
      forwardsExternally: true,
      mailboxForwarding: { forwardingSmtpAddress: 'jane@gmail.com', forwardsExternally: true },
    });
    expect(out.inboxRules.map((r) => r.id)).toEqual(['r2']);
    expect(exo.steps[0]).toMatchObject({ cmdlet: 'Get-Mailbox', params: { Identity: 'jane@contoso.com' } });
  });
});

describe('auto-reply', () => {
  test('scheduled with offsets converts to UTC and PATCHes mailboxSettings', async () => {
    const { calls } = await run(
      'exchange auto-reply set',
      [
        'jane@contoso.com',
        '--scheduled',
        '--start',
        '2026-12-20T17:00:00+01:00',
        '--end',
        '2027-01-05T08:00:00Z',
        '--external-audience',
        'contactsOnly',
        '--internal-message',
        'Out until January',
      ],
      [janeRoute, { method: 'PATCH', path: '/users/u1/mailboxSettings', reply: { body: {} } }],
      { allowWrites: true },
    );
    expect(calls.find((c) => c.method === 'PATCH')?.body).toEqual({
      automaticRepliesSetting: {
        status: 'scheduled',
        internalReplyMessage: 'Out until January',
        externalAudience: 'contactsOnly',
        scheduledStartDateTime: { dateTime: '2026-12-20T16:00:00.000', timeZone: 'UTC' },
        scheduledEndDateTime: { dateTime: '2027-01-05T08:00:00.000', timeZone: 'UTC' },
      },
    });
  });

  test('wall-clock times use the mailbox time zone', async () => {
    const { calls } = await run(
      'exchange auto-reply set',
      ['jane@contoso.com', '--scheduled', '--start', '2026-12-20T17:00', '--end', '2026-12-21'],
      [
        janeRoute,
        { path: '/users/u1/mailboxSettings/timeZone', reply: { body: { value: 'Pacific Standard Time' } } },
        { method: 'PATCH', path: '/users/u1/mailboxSettings', reply: { body: {} } },
      ],
      { allowWrites: true },
    );
    expect(calls.find((c) => c.method === 'PATCH')?.body).toMatchObject({
      automaticRepliesSetting: {
        scheduledStartDateTime: { dateTime: '2026-12-20T17:00', timeZone: 'Pacific Standard Time' },
        scheduledEndDateTime: { dateTime: '2026-12-21T00:00:00', timeZone: 'Pacific Standard Time' },
      },
    });
  });

  test('--off sends only the status; conflicting modes are a usage error', async () => {
    const { calls } = await run(
      'exchange auto-reply set',
      ['jane@contoso.com', '--off'],
      [janeRoute, { method: 'PATCH', path: '/users/u1/mailboxSettings', reply: { body: {} } }],
      { allowWrites: true },
    );
    expect(calls.find((c) => c.method === 'PATCH')?.body).toEqual({ automaticRepliesSetting: { status: 'disabled' } });
    const error = (await run('exchange auto-reply set', ['jane@contoso.com', '--on', '--off'], [], {
      allowWrites: true,
    }).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('USAGE');
  });

  test('mailbox derives autoReplyOn and degrades a forbidden folder to an error', async () => {
    const now = Date.now();
    const iso = (ms: number) => new Date(ms).toISOString().replace('Z', '');
    const { outcome } = await run(
      'exchange mailbox',
      ['jane@contoso.com'],
      [
        janeRoute,
        {
          path: '/users/u1/mailboxSettings',
          reply: {
            body: {
              userPurpose: 'shared',
              timeZone: 'UTC',
              automaticRepliesSetting: {
                status: 'scheduled',
                scheduledStartDateTime: { dateTime: iso(now - 3_600_000), timeZone: 'UTC' },
                scheduledEndDateTime: { dateTime: iso(now + 3_600_000), timeZone: 'UTC' },
                internalReplyMessage: '<p>Away&nbsp;today</p>',
              },
            },
          },
        },
        { path: '/users/u1/mailFolders/inbox', reply: { body: { totalItemCount: 10, unreadItemCount: 2 } } },
        { path: '/users/u1/mailFolders/sentitems', reply: { body: { totalItemCount: 5, unreadItemCount: 0 } } },
        { path: '/users/u1/mailFolders/deleteditems', reply: { body: { totalItemCount: 1, unreadItemCount: 0 } } },
        { path: '/users/u1/mailFolders/junkemail', reply: graphError(403, 'ErrorAccessDenied', 'denied') },
      ],
    );
    const out = result<{ folders: Record<string, { total?: number; unread?: number; error?: { code: string } }> }>(
      outcome,
    );
    expect(out).toMatchObject({
      userPurpose: 'shared',
      autoReplyOn: true,
      autoReply: { internalMessage: 'Away today' },
    });
    expect(out.folders.inbox).toEqual({ total: 10, unread: 2 });
    expect(out.folders.junkEmail?.error?.code).toBe('FORBIDDEN');
  });
});

describe('message trace', () => {
  const now = Date.parse('2026-11-10T12:00:00Z');

  test('builds both date bounds and exact/contains filters', () => {
    expect(
      buildTraceFilter(
        { days: 2, sender: 'a@vendor.com', recipient: "o'neil@contoso.com", subject: 'Q3', status: 'failed' },
        now,
      ),
    ).toBe(
      'receivedDateTime ge 2026-11-08T12:00:00Z and receivedDateTime le 2026-11-10T12:00:00Z and ' +
        "senderAddress eq 'a@vendor.com' and recipientAddress eq 'o''neil@contoso.com' and " +
        "contains(subject, 'Q3') and status eq 'failed'",
    );
  });

  test('rejects windows over 10 days or older than 90 days', () => {
    expect(() => buildTraceFilter({ days: 11 }, now)).toThrow('at most 10 days');
    expect(() =>
      buildTraceFilter({ days: 2, start: '2026-01-01T00:00:00Z', end: '2026-01-03T00:00:00Z' }, now),
    ).toThrow('90 days');
  });

  test('trace lists shaped rows; details looks up the recipient when omitted', async () => {
    const trace = {
      id: 't1',
      senderAddress: 'a@vendor.com',
      recipientAddress: 'jane@contoso.com',
      subject: 'Hi',
      status: 'failed',
      size: 10,
      messageId: '<m@x>',
      receivedDateTime: '2026-11-10T10:00:00Z',
      toIP: '',
    };
    const list = await run(
      'exchange trace',
      ['--recipient', 'jane@contoso.com'],
      [{ path: '/admin/exchange/tracing/messageTraces', reply: page([trace]) }],
    );
    expect(result<{ items: unknown[] }>(list.outcome).items[0]).toMatchObject({
      sender: 'a@vendor.com',
      status: 'failed',
      id: 't1',
    });
    expect(list.calls[0]?.query.$filter).toContain("recipientAddress eq 'jane@contoso.com'");

    const details = await run(
      'exchange trace details',
      ['t1'],
      [
        { path: '/admin/exchange/tracing/messageTraces', reply: page([trace]) },
        {
          path: "/admin/exchange/tracing/messageTraces/t1/getDetailsByRecipient(recipientAddress='jane@contoso.com')",
          reply: page([
            {
              dateTime: '2026-11-10T10:00:05Z',
              event: 'Fail',
              action: '',
              description: '550 5.1.1 User unknown',
              messageId: '<m@x>',
            },
            { dateTime: '2026-11-10T10:00:00Z', event: 'Receive', action: '', description: 'Received' },
          ]),
        },
      ],
    );
    const out = result<{ events: { event: string }[]; finalEvent: unknown }>(details.outcome);
    expect(details.calls[0]?.query.$filter).toBe("id eq 't1'");
    expect(out.events.map((e) => e.event)).toEqual(['Receive', 'Fail']);
    expect(out.finalEvent).toMatchObject({ event: 'Fail', description: '550 5.1.1 User unknown', action: null });
  });
});

describe('messages', () => {
  test('search by sender address uses an exact $filter ordered by date and resolves folder names', async () => {
    const { outcome, calls } = await run(
      'exchange messages search',
      ['jane@contoso.com', '--from', 'billing@vendor.com', '--has-attachments', '--folder', 'junk'],
      [
        janeRoute,
        {
          path: '/users/u1/mailFolders/junkemail/messages',
          reply: page([
            {
              id: 'm1',
              subject: 'Invoice',
              from: { emailAddress: { address: 'billing@vendor.com', name: 'Billing' } },
              toRecipients: [{ emailAddress: { address: 'jane@contoso.com' } }],
              parentFolderId: 'fj',
              isRead: false,
            },
          ]),
        },
        { path: '/users/u1/mailFolders/fj', reply: { body: { displayName: 'Junk Email' } } },
      ],
    );
    const call = calls.find((c) => c.path.endsWith('/messages'));
    expect(call?.query.$filter).toMatch(
      /^receivedDateTime ge \S+ and from\/emailAddress\/address eq 'billing@vendor.com' and hasAttachments eq true$/,
    );
    expect(call?.query.$orderby).toBe('receivedDateTime desc');
    expect(call?.query.$select).not.toContain('body');
    expect(result<{ items: unknown[] }>(outcome).items[0]).toMatchObject({
      from: 'billing@vendor.com',
      to: ['jane@contoso.com'],
      folder: 'Junk Email',
    });
  });

  test('search by subject uses KQL $search without $filter', async () => {
    const { calls } = await run(
      'exchange messages search',
      ['jane@contoso.com', '--subject', 'invoice "overdue"', '--since', '7'],
      [janeRoute, { path: '/users/u1/messages', reply: page([]) }],
    );
    const call = calls.find((c) => c.path === '/users/u1/messages');
    expect(call?.query.$search).toMatch(/^"subject:invoice AND subject:overdue AND received>=\d+\/\d+\/\d{4}"$/);
    expect(call?.query.$filter).toBeUndefined();
  });

  test('delete by Internet Message-ID previews, then hard delete calls permanentDelete', async () => {
    const routes: MockRoute[] = [
      janeRoute,
      {
        path: '/users/u1/messages',
        reply: page([
          { id: 'm1', subject: 'Phish', from: { emailAddress: { address: 'x@evil.example' } }, parentFolderId: 'fi' },
        ]),
      },
      { path: '/users/u1/mailFolders/fi', reply: { body: { displayName: 'Inbox' } } },
      { method: 'POST', path: '/users/u1/messages/m1/permanentDelete', reply: { status: 204 } },
    ];
    const argv = ['jane@contoso.com', '<abc@evil.example>', '--hard'];
    const first = await run('exchange messages delete', argv, routes, { allowWrites: true });
    if (first.outcome.kind !== 'confirm') throw new Error('expected confirm');
    expect(first.outcome.preview).toMatchObject({ message: { subject: 'Phish', folder: 'Inbox', id: 'm1' } });
    expect(first.calls.find((c) => c.path === '/users/u1/messages')?.query.$filter).toBe(
      "internetMessageId eq '<abc@evil.example>'",
    );
    expect(first.calls.some((c) => c.method === 'POST')).toBe(false);

    const second = await run('exchange messages delete', argv, routes, {
      allowWrites: true,
      confirm: first.outcome.token,
    });
    expect(result(second.outcome)).toMatchObject({ deleted: true, permanent: true });
    expect(second.calls.some((c) => c.method === 'POST' && c.path === '/users/u1/messages/m1/permanentDelete')).toBe(
      true,
    );
  });
});

describe('calendar', () => {
  test('events request the window in the mailbox time zone', async () => {
    const { outcome, calls } = await run(
      'exchange calendar events',
      ['jane@contoso.com', '--days', '3'],
      [
        janeRoute,
        { path: '/users/u1/mailboxSettings/timeZone', reply: { body: { value: 'W. Europe Standard Time' } } },
        {
          path: '/users/u1/calendarView',
          reply: page([
            {
              id: 'e1',
              subject: 'Standup',
              start: { dateTime: '2026-11-11T09:00:00', timeZone: 'W. Europe Standard Time' },
              end: { dateTime: '2026-11-11T09:15:00', timeZone: 'W. Europe Standard Time' },
              organizer: { emailAddress: { address: 'bob@contoso.com' } },
              location: { displayName: '' },
              isOnlineMeeting: true,
              showAs: 'busy',
            },
          ]),
        },
      ],
    );
    const call = calls.find((c) => c.path === '/users/u1/calendarView');
    expect(call?.headers.Prefer).toBe('outlook.timezone="W. Europe Standard Time"');
    expect(Date.parse(call?.query.endDateTime ?? '') - Date.parse(call?.query.startDateTime ?? '')).toBe(
      3 * 86_400_000,
    );
    expect(result<{ items: unknown[] }>(outcome).items[0]).toMatchObject({
      subject: 'Standup',
      organizer: 'bob@contoso.com',
      location: null,
      timeZone: 'W. Europe Standard Time',
    });
  });
});

describe('usage report', () => {
  const header =
    'Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Created Date,Last Activity Date,Item Count,Storage Used (Byte),Issue Warning Quota (Byte),Prohibit Send Quota (Byte),Prohibit Send/Receive Quota (Byte),Deleted Item Count,Deleted Item Size (Byte),Deleted Item Quota (Byte),Has Archive,Report Period';
  const gb = 1024 ** 3;
  const csv = [
    header,
    `2026-11-09,jane@contoso.com,Jane Doe,False,,,2026-11-08,1200,${95 * gb},${98 * gb},${99 * gb},${100 * gb},40,${gb},${30 * gb},True,7`,
    `2026-11-09,bob@contoso.com,Bob,False,,,2026-11-08,300,${10 * gb},${98 * gb},${99 * gb},${100 * gb},1,0,${30 * gb},False,7`,
  ].join('\n');

  test('shapes CSV rows with GB values and percent of the send quota', () => {
    const [header0, jane] = csv.split('\n');
    const row = Object.fromEntries((header0 ?? '').split(',').map((h, i) => [h, (jane ?? '').split(',')[i] ?? '']));
    expect(shapeMailboxUsage(row)).toMatchObject({
      user: 'jane@contoso.com',
      storageUsedGb: 95,
      prohibitSendQuotaGb: 99,
      percentUsed: 96,
      itemCount: 1200,
      hasArchive: true,
      isDeleted: false,
    });
  });

  test('--near-quota keeps only mailboxes at 90%+, via the report redirect', async () => {
    const { outcome, calls } = await run(
      'exchange usage',
      ['--near-quota'],
      [
        {
          path: "/reports/getMailboxUsageDetail(period='D7')",
          reply: { status: 302, headers: { location: 'https://graph.test/download/report.csv' } },
        },
        { path: '/download/report.csv', reply: { body: csv } },
      ],
    );
    expect(calls[0]?.path).toBe("/reports/getMailboxUsageDetail(period='D7')");
    expect(result(outcome)).toEqual({
      items: [expect.objectContaining({ user: 'jane@contoso.com', percentUsed: 96 })],
      truncated: false,
    });
  });
});
