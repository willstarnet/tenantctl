import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands, summarizeEvidence } from '../../src/domains/defender/index.ts';
import { kqlString, parseParams, renderTemplate } from '../../src/domains/defender/templates.ts';
import { commandById, page, runCommand } from '../helpers/graphMock.ts';

const cmd = (id: string) => commandById(commands, `defender ${id}`);
const result = (outcome: { kind: string } & Record<string, unknown>) => {
  expect(outcome.kind).toBe('result');
  return outcome.result as Record<string, unknown>;
};

const userEvidence = (upn: string) => ({
  '@odata.type': '#microsoft.graph.security.userEvidence',
  userAccount: { accountName: upn.split('@')[0], userPrincipalName: upn },
});
const deviceEvidence = (name: string) => ({
  '@odata.type': '#microsoft.graph.security.deviceEvidence',
  deviceDnsName: name,
});

describe('summarizeEvidence', () => {
  test('groups entities by kind and dedupes case-insensitively across alerts', () => {
    const summary = summarizeEvidence([
      userEvidence('jane@contoso.com'),
      userEvidence('JANE@contoso.com'),
      deviceEvidence('pc1.contoso.com'),
      deviceEvidence('PC1.contoso.com'),
      { '@odata.type': '#microsoft.graph.security.ipEvidence', ipAddress: '203.0.113.5' },
      { '@odata.type': '#microsoft.graph.security.ipEvidence', ipAddress: '203.0.113.5' },
      { '@odata.type': '#microsoft.graph.security.mailboxEvidence', primaryAddress: 'jane@contoso.com' },
      {
        '@odata.type': '#microsoft.graph.security.fileEvidence',
        fileDetails: { fileName: 'evil.exe', filePath: 'C:\\Temp', sha256: 'ABC' },
      },
      {
        '@odata.type': '#microsoft.graph.security.fileEvidence',
        fileDetails: { fileName: 'evil-copy.exe', filePath: 'D:\\', sha256: 'abc' },
      },
      { '@odata.type': '#microsoft.graph.security.urlEvidence', url: 'https://evil.example/login' },
      {
        '@odata.type': '#microsoft.graph.security.analyzedMessageEvidence',
        networkMessageId: 'm1',
        subject: 'Reset your password',
        recipientEmailAddress: 'bob@contoso.com',
        deliveryAction: 'delivered',
        p1Sender: { emailAddress: 'it@evil.example' },
        urls: ['https://evil.example/login', 'https://evil.example/2'],
      },
      // Service accounts running a process are noise; only real users count.
      {
        '@odata.type': '#microsoft.graph.security.processEvidence',
        userAccount: { accountName: 'SYSTEM', domainName: 'NT AUTHORITY', userPrincipalName: null },
      },
      { '@odata.type': '#microsoft.graph.security.oauthApplicationEvidence', displayName: 'Mail Sync App' },
    ]);
    expect(summary).toEqual({
      users: ['jane@contoso.com'],
      devices: ['pc1.contoso.com'],
      ips: ['203.0.113.5'],
      mailboxes: ['jane@contoso.com', 'bob@contoso.com'],
      urls: ['https://evil.example/login', 'https://evil.example/2'],
      files: [{ name: 'evil.exe', path: 'C:\\Temp', sha256: 'ABC' }],
      emails: [
        {
          subject: 'Reset your password',
          sender: 'it@evil.example',
          recipient: 'bob@contoso.com',
          deliveryAction: 'delivered',
        },
      ],
      apps: ['Mail Sync App'],
    });
  });

  test('on-prem accounts without a UPN show as DOMAIN\\name; empty input gives an empty summary', () => {
    expect(
      summarizeEvidence([
        {
          '@odata.type': '#microsoft.graph.security.userEvidence',
          userAccount: { accountName: 'jdoe', domainName: 'CORP' },
        },
      ]),
    ).toEqual({ users: ['CORP\\jdoe'] });
    expect(summarizeEvidence([])).toEqual({});
  });
});

describe('incidents', () => {
  test('list builds the filter, expands alerts for alertCount and shapes rows', async () => {
    const { outcome, calls } = await runCommand(
      cmd('incidents list'),
      ['--status', 'active', '--severity', 'high', '--days', '7', '--unassigned'],
      [
        {
          path: '/security/incidents',
          reply: page([
            {
              id: '42',
              displayName: 'Multi-stage incident',
              severity: 'high',
              status: 'active',
              assignedTo: null,
              createdDateTime: '2026-11-09T10:00:00Z',
              incidentWebUrl: 'https://security.microsoft.com/incidents/42',
              alerts: [{ id: 'a1' }, { id: 'a2' }],
              '@odata.etag': 'noise',
            },
          ]),
        },
      ],
    );
    const filter = calls[0]?.query.$filter ?? '';
    expect(filter).toStartWith("status eq 'active' and severity eq 'high' and createdDateTime ge ");
    expect(filter).toEndWith(' and assignedTo eq null');
    expect(calls[0]?.query.$expand).toBe('alerts');
    expect(calls[0]?.query.$top).toBe('50');
    const items = (result(outcome) as unknown as { items: Record<string, unknown>[] }).items;
    expect(items[0]).toMatchObject({ id: '42', alertCount: 2, status: 'active', assignedTo: null });
    expect(items[0]).not.toHaveProperty('@odata.etag');
  });

  test('get merges entities from every alert and includes comments', async () => {
    const { outcome, calls } = await runCommand(
      cmd('incidents get'),
      ['42'],
      [
        {
          path: '/security/incidents/42',
          reply: {
            body: {
              id: '42',
              displayName: 'Phish',
              severity: 'medium',
              status: 'active',
              comments: [{ comment: 'Looking', createdByDisplayName: 'Ana', createdDateTime: '2026-11-09T11:00:00Z' }],
              alerts: [
                {
                  id: 'a1',
                  title: 'Phish delivered',
                  serviceSource: 'microsoftDefenderForOffice365',
                  categories: ['InitialAccess'],
                  mitreTechniques: ['T1566'],
                  evidence: [userEvidence('jane@contoso.com')],
                },
                {
                  id: 'a2',
                  title: 'Suspicious sign-in',
                  evidence: [userEvidence('Jane@contoso.com'), deviceEvidence('pc1')],
                },
              ],
            },
          },
        },
      ],
    );
    expect(calls[0]?.query.$expand).toBe('alerts');
    const shaped = result(outcome);
    expect(shaped.entities).toEqual({ users: ['jane@contoso.com'], devices: ['pc1'] });
    expect(shaped.alertCount).toBe(2);
    expect((shaped.alerts as Record<string, unknown>[])[0]).toMatchObject({
      category: 'InitialAccess',
      mitreTechniques: ['T1566'],
      serviceSource: 'microsoftDefenderForOffice365',
    });
    expect(shaped.comments).toEqual([{ comment: 'Looking', by: 'Ana', at: '2026-11-09T11:00:00Z' }]);
  });

  test('update PATCHes only the given fields, "none" unassigns, and the comment is POSTed separately', async () => {
    const { outcome, calls } = await runCommand(
      cmd('incidents update'),
      [
        '42',
        '--status',
        'resolved',
        '--classification',
        'falsePositive',
        '--assign-to',
        'none',
        '--comment',
        'Pen test',
      ],
      [
        { method: 'PATCH', path: '/security/incidents/42', reply: { body: { id: '42', status: 'resolved' } } },
        { method: 'POST', path: '/security/incidents/42/comments', reply: page([]) },
      ],
      { allowWrites: true },
    );
    expect(calls.map((c) => `${c.method} ${c.path}`)).toEqual([
      'PATCH /security/incidents/42',
      'POST /security/incidents/42/comments',
    ]);
    expect(calls[0]?.body).toEqual({ status: 'resolved', classification: 'falsePositive', assignedTo: null });
    expect(calls[1]?.body).toEqual({ '@odata.type': 'microsoft.graph.security.alertComment', comment: 'Pen test' });
    expect(result(outcome)).toMatchObject({ commented: true, incident: { id: '42', status: 'resolved' } });
  });

  test('update needs writes enabled and at least one change', async () => {
    const disabled = (await runCommand(cmd('incidents update'), ['42', '--status', 'resolved'], []).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(disabled.code).toBe('WRITES_DISABLED');
    const empty = (await runCommand(cmd('incidents update'), ['42'], [], { allowWrites: true }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(empty.code).toBe('USAGE');
  });
});

describe('alerts', () => {
  const alert = (id: string, categories: string[]) => ({
    id,
    title: `Alert ${id}`,
    severity: 'high',
    status: 'new',
    serviceSource: 'microsoftDefenderForEndpoint',
    categories,
    incidentId: '42',
    evidence: [deviceEvidence('pc1.contoso.com'), userEvidence('jane@contoso.com')],
  });

  test('list filters server-side and shapes the first user and device from evidence', async () => {
    const { outcome, calls } = await runCommand(
      cmd('alerts list'),
      ['--severity', 'high', '--status', 'new', '--source', 'microsoftSentinel'],
      [{ path: '/security/alerts_v2', reply: page([alert('a1', ['Execution'])]) }],
    );
    expect(calls[0]?.query.$filter).toBe(
      "severity eq 'high' and status eq 'new' and serviceSource eq 'microsoftSentinel'",
    );
    // Newer serviceSource values are only filterable/returned with this opt-in.
    expect(calls[0]?.headers.Prefer).toBe('include-unknown-enum-members');
    const items = (result(outcome) as unknown as { items: Record<string, unknown>[] }).items;
    expect(items[0]).toMatchObject({
      id: 'a1',
      category: 'Execution',
      user: 'jane@contoso.com',
      device: 'pc1.contoso.com',
      incidentId: '42',
    });
  });

  test('--category is matched client-side across pages until --top matches', async () => {
    const { outcome, calls } = await runCommand(
      cmd('alerts list'),
      ['--category', 'credentialaccess', '--top', '2'],
      [
        {
          path: '/security/alerts_v2',
          reply: (req) =>
            req.query.$skiptoken
              ? page([alert('a3', ['CredentialAccess']), alert('a4', ['CredentialAccess'])])
              : page(
                  [alert('a1', ['CredentialAccess']), alert('a2', ['Execution'])],
                  'https://graph.test/v1.0/security/alerts_v2?$skiptoken=x',
                ),
        },
      ],
    );
    expect(calls).toHaveLength(2);
    expect(calls[0]?.query.$filter).toBeUndefined();
    const shaped = result(outcome) as unknown as { items: { id: string }[]; truncated: boolean };
    expect(shaped.items.map((a) => a.id)).toEqual(['a1', 'a3']);
    expect(shaped.truncated).toBe(true);
  });

  test('update sends the alert PATCH body with assignedTo', async () => {
    const { calls } = await runCommand(
      cmd('alerts update'),
      ['a1', '--status', 'inProgress', '--assign-to', 'secops@contoso.com', '--determination', 'malware'],
      [{ method: 'PATCH', path: '/security/alerts_v2/a1', reply: { body: alert('a1', []) } }],
      { allowWrites: true },
    );
    expect(calls).toHaveLength(1);
    expect(calls[0]?.body).toEqual({
      status: 'inProgress',
      assignedTo: 'secops@contoso.com',
      determination: 'malware',
    });
  });

  test('for-user resolves the user, hunts by UPN over --days and adds incident ids', async () => {
    const { outcome, calls } = await runCommand(
      cmd('alerts for-user'),
      ['jane@contoso.com', '--days', '7'],
      [
        {
          path: '/users/jane@contoso.com',
          reply: { body: { id: 'u1', userPrincipalName: 'jane@contoso.com', displayName: 'Jane' } },
        },
        {
          method: 'POST',
          path: '/security/runHuntingQuery',
          reply: {
            body: {
              schema: [{ name: 'AlertId', type: 'String' }],
              results: [
                {
                  Timestamp: '2026-11-09T10:00:00Z',
                  AlertId: 'da1',
                  Title: 'Phish',
                  Severity: 'High',
                  ServiceSource: 'Microsoft Defender for Office 365',
                },
              ],
            },
          },
        },
        { path: '/security/alerts_v2/da1', reply: { body: { id: 'da1', status: 'new', incidentId: '42' } } },
      ],
    );
    const hunt = calls.find((c) => c.path === '/security/runHuntingQuery');
    const body = hunt?.body as { Query: string; Timespan: string };
    expect(body.Timespan).toBe('P7D');
    expect(body.Query).toContain("AccountUpn =~ 'jane@contoso.com'");
    expect(body.Query).toEndWith('| take 26');
    expect(result(outcome).alerts).toEqual([
      {
        alertId: 'da1',
        title: 'Phish',
        severity: 'high',
        category: null,
        serviceSource: 'Microsoft Defender for Office 365',
        time: '2026-11-09T10:00:00Z',
        status: 'new',
        incidentId: '42',
        alertWebUrl: null,
      },
    ]);
  });
});

describe('hunting', () => {
  test('template values become escaped KQL string literals', () => {
    expect(kqlString("O'Brien\\x\nnext")).toBe("'O\\'Brien\\\\x\\nnext'");
    const kql = renderTemplate('emails-from-sender', { sender: "x' or 1==1 //" });
    expect(kql).toContain("SenderFromAddress =~ 'x\\' or 1==1 //'");
    expect(kql).not.toContain('{sender}');
  });

  test('templates reject missing and unknown params', () => {
    expect(() => renderTemplate('device-logons', {})).toThrow('--param device=');
    expect(() => renderTemplate('device-logons', { device: 'pc1', user: 'x' })).toThrow('no parameter user');
    expect(() => renderTemplate('nope', {})).toThrow('Unknown hunting template');
  });

  test('--param keeps the whole value after the first =, commas included', () => {
    expect(parseParams(['subject=Invoice, final notice', 'other=a=b'])).toEqual({
      subject: 'Invoice, final notice',
      other: 'a=b',
    });
    expect(() => parseParams(['novalue'])).toThrow('name=value');
  });

  test('hunt posts Query and normalized Timespan and caps rows at --top', async () => {
    const rows = [{ n: 1 }, { n: 2 }, { n: 3 }];
    const { outcome, calls } = await runCommand(
      cmd('hunt'),
      ['--template', 'email-by-subject', '--param', 'subject=Invoice, final notice', '--timespan', '7d', '--top', '2'],
      [
        {
          method: 'POST',
          path: '/security/runHuntingQuery',
          reply: { body: { schema: [{ name: 'n', type: 'Int32' }], results: rows } },
        },
      ],
    );
    const body = calls[0]?.body as { Query: string; Timespan: string };
    expect(body.Timespan).toBe('P7D');
    expect(body.Query).toContain("Subject contains 'Invoice, final notice'");
    expect(result(outcome)).toMatchObject({
      columns: [{ name: 'n', type: 'Int32' }],
      rows: [{ n: 1 }, { n: 2 }],
      rowCount: 3,
      truncated: true,
    });
  });

  test('hunt needs exactly one of a query or a template, and a valid timespan', async () => {
    const neither = (await runCommand(cmd('hunt'), [], []).catch((e: unknown) => e)) as CliError;
    expect(neither.code).toBe('USAGE');
    const badSpan = (await runCommand(cmd('hunt'), ['DeviceInfo', '--timespan', 'last week'], []).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(badSpan.code).toBe('USAGE');
  });

  test('templates command lists every template with its params without calling Graph', async () => {
    const { outcome, calls } = await runCommand(cmd('hunt templates'), [], []);
    const items = (result(outcome) as unknown as { items: { name: string; params: Record<string, string> }[] }).items;
    expect(calls).toHaveLength(0);
    expect(items.map((t) => t.name)).toContain('inbox-rule-changes');
    expect(items.find((t) => t.name === 'file-hash-prevalence')?.params).toHaveProperty('sha256');
  });
});

describe('secure-score', () => {
  test('joins control scores with profiles and ranks actions by remaining points', async () => {
    const { outcome, calls } = await runCommand(
      cmd('secure-score'),
      ['--top-actions', '1'],
      [
        {
          path: '/security/secureScores',
          reply: page([
            {
              createdDateTime: '2026-11-09T00:00:00Z',
              currentScore: 40,
              maxScore: 80,
              averageComparativeScores: [{ basis: 'AllTenants', averageScore: 37 }],
              controlScores: [
                { controlName: 'AdminMFAV2', score: 2 },
                { controlName: 'DLPEnabled', score: 0 },
                { controlName: 'OldControl', score: 0 },
                { controlName: 'Done', score: 5 },
              ],
            },
          ]),
        },
        {
          path: '/security/secureScoreControlProfiles',
          reply: page([
            {
              id: 'AdminMFAV2',
              title: 'Require MFA for admins',
              maxScore: 10,
              userImpact: 'Low',
              actionUrl: 'https://x',
            },
            { id: 'DLPEnabled', title: 'Turn on DLP', maxScore: 3 },
            { id: 'OldControl', title: 'Old', maxScore: 50, deprecated: true },
            { id: 'Done', title: 'Done', maxScore: 5 },
          ]),
        },
      ],
    );
    expect(calls.find((c) => c.path === '/security/secureScores')?.query.$top).toBe('1');
    expect(result(outcome)).toMatchObject({
      currentScore: 40,
      maxScore: 80,
      percent: 50,
      comparison: { allTenants: 37 },
      topActions: [{ control: 'AdminMFAV2', currentPoints: 2, maxScore: 10, remainingPoints: 8, userImpact: 'Low' }],
      remainingActions: 1,
    });
  });
});

describe('threat-submit email', () => {
  test('resolves the recipient and submits the message URL to the beta endpoint', async () => {
    const { calls } = await runCommand(
      cmd('threat-submit email'),
      ['jane@contoso.com', 'AAMk/abc=', '--category', 'phishing'],
      [
        {
          path: '/users/jane@contoso.com',
          reply: { body: { id: 'u1', userPrincipalName: 'jane@contoso.com', mail: 'jane@contoso.com' } },
        },
        {
          method: 'POST',
          path: '/security/threatSubmission/emailThreats',
          reply: { status: 201, body: { id: 's1', status: 'running' } },
        },
      ],
      { allowWrites: true },
    );
    const submit = calls.find((c) => c.method === 'POST');
    expect(submit?.version).toBe('beta');
    expect(submit?.body).toEqual({
      '@odata.type': '#microsoft.graph.security.emailUrlThreatSubmission',
      category: 'phishing',
      recipientEmailAddress: 'jane@contoso.com',
      messageUrl: 'https://graph.microsoft.com/beta/users/u1/messages/AAMk%2Fabc%3D',
    });
  });
});
