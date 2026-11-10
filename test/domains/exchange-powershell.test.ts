import { describe, expect, test } from 'bun:test';
import { CliError } from '../../src/core/errors.ts';
import type { Outcome } from '../../src/core/execute.ts';
import type { ExoStep } from '../../src/core/exo.ts';
import { commands, flattenMailboxPermissions } from '../../src/domains/exchange/index.ts';
import { commandById, fakeExo, page, runCommand, type MockRoute } from '../helpers/graphMock.ts';

const JANE = { id: 'u1', displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com', mail: 'jane@contoso.com' };
const janeRoute: MockRoute = { path: '/users/jane@contoso.com', reply: { body: JANE } };
const domainsRoute: MockRoute = {
  path: '/domains',
  reply: page([
    { id: 'contoso.com', isVerified: true },
    { id: 'contoso.onmicrosoft.com', isVerified: true },
  ]),
};
const noRules: MockRoute = { path: '/users/u1/mailFolders/inbox/messageRules', reply: page([]) };

async function run(id: string, argv: string[], routes: MockRoute[], options = {}) {
  return runCommand(commandById(commands, id), argv, routes, options);
}

function result<T = unknown>(outcome: Outcome): T {
  if (outcome.kind !== 'result') throw new Error(`expected result, got ${outcome.kind}`);
  return outcome.result as T;
}

const isWrite = (step: ExoStep) => !step.cmdlet.startsWith('Get-');

describe('exchange forwarding (mailbox level)', () => {
  test('reads Get-Mailbox with Identity and select, and flags external forwarding', async () => {
    const exo = fakeExo(() => [
      { ForwardingSmtpAddress: 'smtp:jane.doe@gmail.com', ForwardingAddress: null, DeliverToMailboxAndForward: true },
    ]);
    const { outcome } = await run('exchange forwarding', ['jane@contoso.com'], [janeRoute, domainsRoute, noRules], {
      exo,
    });
    expect(exo.steps).toEqual([
      {
        cmdlet: 'Get-Mailbox',
        params: { Identity: 'jane@contoso.com' },
        select: ['ForwardingSmtpAddress', 'ForwardingAddress', 'DeliverToMailboxAndForward'],
      },
    ]);
    expect(result(outcome)).toMatchObject({
      forwardsViaInboxRules: false,
      forwardsViaMailbox: true,
      forwardsExternally: true,
      mailboxForwarding: {
        forwardingSmtpAddress: 'jane.doe@gmail.com',
        deliverToMailboxAndForward: true,
        forwardsExternally: true,
      },
    });
  });

  test('internal mailbox forwarding is not external', async () => {
    const exo = fakeExo(() => [{ ForwardingSmtpAddress: 'SMTP:bob@sales.contoso.com' }]);
    const { outcome } = await run('exchange forwarding', ['jane@contoso.com'], [janeRoute, domainsRoute, noRules], {
      exo,
    });
    expect(result(outcome)).toMatchObject({
      forwardsExternally: false,
      mailboxForwarding: { forwardsExternally: false },
    });
  });

  test('degrades to { error, hint } when the bridge fails, keeping inbox-rule results', async () => {
    const exo = {
      async run(): Promise<never> {
        throw new CliError('CONFIG', 'Exchange admin commands need PowerShell 7 (pwsh)', { hint: 'Install pwsh' });
      },
    };
    const { outcome } = await run('exchange forwarding', ['jane@contoso.com'], [janeRoute, domainsRoute, noRules], {
      exo,
    });
    expect(result(outcome)).toMatchObject({
      forwardsViaInboxRules: false,
      forwardsViaMailbox: null,
      forwardsExternally: null,
      mailboxForwarding: { error: 'Exchange admin commands need PowerShell 7 (pwsh)', hint: 'Install pwsh' },
    });
  });

  test('forwarding set previews current state and external target, writes only after confirm', async () => {
    const exo = fakeExo((step) => (step.cmdlet === 'Get-Mailbox' ? [{ ForwardingSmtpAddress: null }] : []));
    const argv = ['jane@contoso.com', '--to', 'x@gmail.com', '--keep-copy'];
    const first = await run('exchange forwarding set', argv, [domainsRoute], { exo, allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    if (first.outcome.kind !== 'confirm') return;
    expect(first.outcome.preview).toMatchObject({ targetIsExternal: true, current: { forwardingSmtpAddress: null } });
    expect((first.outcome.preview as { action: string }).action).toContain('EXTERNAL');
    expect(exo.steps.some(isWrite)).toBe(false);

    await run('exchange forwarding set', argv, [domainsRoute], {
      exo,
      allowWrites: true,
      confirm: first.outcome.token,
    });
    expect(exo.steps.find(isWrite)).toEqual({
      cmdlet: 'Set-Mailbox',
      params: { Identity: 'jane@contoso.com', ForwardingSmtpAddress: 'x@gmail.com', DeliverToMailboxAndForward: true },
    });
  });

  test('forwarding clear nulls both forwarding properties', async () => {
    const exo = fakeExo((step) =>
      step.cmdlet === 'Get-Mailbox' ? [{ ForwardingSmtpAddress: 'smtp:x@gmail.com' }] : [],
    );
    const first = await run('exchange forwarding clear', ['jane'], [domainsRoute], { exo, allowWrites: true });
    if (first.outcome.kind !== 'confirm') throw new Error('expected confirm');
    expect((first.outcome.preview as { action: string }).action).toContain('forwarding to x@gmail.com');
    await run('exchange forwarding clear', ['jane'], [domainsRoute], {
      exo,
      allowWrites: true,
      confirm: first.outcome.token,
    });
    expect(exo.steps.find(isWrite)?.params).toEqual({
      Identity: 'jane',
      ForwardingSmtpAddress: null,
      ForwardingAddress: null,
      DeliverToMailboxAndForward: false,
    });
  });
});

const PERMISSION_DATA: Record<string, Record<string, unknown>[]> = {
  'Get-MailboxPermission': [
    { User: 'NT AUTHORITY\\SELF', AccessRights: ['FullAccess', 'ReadPermission'], IsInherited: false, Deny: false },
    { User: 'bob@contoso.com', AccessRights: ['FullAccess'], IsInherited: false, Deny: false },
    { User: 'NAMPR01\\Admins', AccessRights: ['FullAccess'], IsInherited: true, Deny: false },
  ],
  'Get-RecipientPermission': [
    { Trustee: 'NT AUTHORITY\\SELF', AccessRights: ['SendAs'], IsInherited: false, AccessControlType: 'Allow' },
    { Trustee: 'amy@contoso.com', AccessRights: ['SendAs'], IsInherited: false, AccessControlType: 'Allow' },
  ],
  'Get-Mailbox': [{ PrimarySmtpAddress: 'finance@contoso.com', GrantSendOnBehalfTo: ['Carl Jones'] }],
};

describe('exchange permissions', () => {
  test('list runs three cmdlets in one session and flattens without SELF or inherited', async () => {
    const exo = fakeExo((step) => PERMISSION_DATA[step.cmdlet] ?? []);
    const { outcome } = await run('exchange permissions list', ['finance'], [], { exo });
    expect(exo.steps.map((s) => [s.cmdlet, s.params?.Identity])).toEqual([
      ['Get-MailboxPermission', 'finance'],
      ['Get-RecipientPermission', 'finance'],
      ['Get-Mailbox', 'finance'],
    ]);
    expect(exo.steps.every((s) => s.select?.length)).toBe(true);
    expect(result<{ items: unknown[] }>(outcome).items).toEqual([
      { user: 'bob@contoso.com', access: 'FullAccess' },
      { user: 'amy@contoso.com', access: 'SendAs' },
      { user: 'Carl Jones', access: 'SendOnBehalf' },
    ]);
  });

  test('inherited entries are marked when included', () => {
    const list = flattenMailboxPermissions(PERMISSION_DATA['Get-MailboxPermission'] ?? [], [], undefined, true);
    expect(list).toContainEqual({ user: 'NAMPR01\\Admins', access: 'FullAccess', inherited: true });
  });

  test('add full previews current permissions, then Add-MailboxPermission with AutoMapping', async () => {
    const exo = fakeExo((step) => PERMISSION_DATA[step.cmdlet] ?? []);
    const argv = ['finance', 'bob@contoso.com', '--access', 'full', '--no-automap'];
    const first = await run('exchange permissions add', argv, [], { exo, allowWrites: true });
    if (first.outcome.kind !== 'confirm') throw new Error('expected confirm');
    expect(first.outcome.preview).toMatchObject({ access: 'FullAccess', alreadyGranted: true });
    expect(exo.steps.some(isWrite)).toBe(false);
    await run('exchange permissions add', argv, [], { exo, allowWrites: true, confirm: first.outcome.token });
    expect(exo.steps.find(isWrite)).toMatchObject({
      cmdlet: 'Add-MailboxPermission',
      params: { Identity: 'finance', User: 'bob@contoso.com', AccessRights: 'FullAccess', AutoMapping: false },
    });
  });

  test('remove sendonbehalf uses Set-Mailbox GrantSendOnBehalfTo @{Remove}', async () => {
    const exo = fakeExo((step) => PERMISSION_DATA[step.cmdlet] ?? []);
    const argv = ['finance', 'Carl Jones', '--access', 'sendonbehalf'];
    const first = await run('exchange permissions remove', argv, [], { exo, allowWrites: true });
    if (first.outcome.kind !== 'confirm') throw new Error('expected confirm');
    expect(first.outcome.preview).toMatchObject({ currentlyGranted: true });
    await run('exchange permissions remove', argv, [], { exo, allowWrites: true, confirm: first.outcome.token });
    expect(exo.steps.find(isWrite)).toEqual({
      cmdlet: 'Set-Mailbox',
      params: { Identity: 'finance', GrantSendOnBehalfTo: { Remove: 'Carl Jones' } },
    });
  });

  test('an ambiguous identity becomes AMBIGUOUS', async () => {
    const exo = fakeExo(() => {
      throw new Error("The operation couldn't be performed because 'jane' matches multiple entries.");
    });
    const error = (await run('exchange permissions list', ['jane'], [], { exo }).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
  });
});

describe('exchange mailbox and admin cmdlets', () => {
  test('shared-mailbox create grants FullAccess and SendAs per member in one session', async () => {
    const exo = fakeExo((step) =>
      step.cmdlet === 'New-Mailbox'
        ? [{ PrimarySmtpAddress: 'finance@contoso.com', RecipientTypeDetails: 'SharedMailbox' }]
        : [],
    );
    const { outcome } = await run(
      'exchange shared-mailbox create',
      ['--name', 'Finance', '--email', 'finance@contoso.com', '--members', 'jane@contoso.com'],
      [],
      { exo, allowWrites: true },
    );
    expect(exo.steps.map((s) => s.cmdlet)).toEqual(['New-Mailbox', 'Add-MailboxPermission', 'Add-RecipientPermission']);
    expect(exo.steps[0]?.params).toMatchObject({
      Shared: true,
      Name: 'Finance',
      PrimarySmtpAddress: 'finance@contoso.com',
    });
    expect(result(outcome)).toMatchObject({
      type: 'SharedMailbox',
      members: [{ member: 'jane@contoso.com', fullAccess: { status: 'done' }, sendAs: { status: 'done' } }],
    });
  });

  test('mailbox convert previews type, size and license note before Set-Mailbox -Type', async () => {
    const exo = fakeExo((step) =>
      step.cmdlet === 'Get-Mailbox'
        ? [
            {
              PrimarySmtpAddress: 'jane@contoso.com',
              RecipientTypeDetails: 'UserMailbox',
              ArchiveStatus: 'None',
              LitigationHoldEnabled: false,
            },
          ]
        : step.cmdlet === 'Get-MailboxStatistics'
          ? [{ TotalItemSize: '60 GB (64,424,509,440 bytes)' }]
          : [],
    );
    const argv = ['jane@contoso.com', '--to', 'shared'];
    const first = await run('exchange mailbox convert', argv, [], { exo, allowWrites: true });
    if (first.outcome.kind !== 'confirm') throw new Error('expected confirm');
    expect(first.outcome.preview).toMatchObject({ type: 'UserMailbox', target: 'SharedMailbox', sizeGb: 60 });
    expect((first.outcome.preview as { licenseNote: string }).licenseNote).toContain('Keep a license');
    expect(exo.steps.some(isWrite)).toBe(false);
    await run('exchange mailbox convert', argv, [], { exo, allowWrites: true, confirm: first.outcome.token });
    expect(exo.steps.find(isWrite)).toEqual({
      cmdlet: 'Set-Mailbox',
      params: { Identity: 'jane@contoso.com', Type: 'Shared' },
    });
  });

  test('mailbox details shapes aliases and statistics from one session', async () => {
    const exo = fakeExo((step) =>
      step.cmdlet === 'Get-Mailbox'
        ? [
            {
              PrimarySmtpAddress: 'jane@contoso.com',
              EmailAddresses: ['SMTP:jane@contoso.com', 'smtp:jd@contoso.com', 'SIP:jane@contoso.com'],
              LitigationHoldEnabled: true,
            },
          ]
        : [{ TotalItemSize: '1.5 GB (1,610,612,736 bytes)', ItemCount: 1200, LastLogonTime: '2026-11-01T08:00:00Z' }],
    );
    const { outcome } = await run('exchange mailbox details', ['jane'], [], { exo });
    expect(exo.steps.map((s) => s.cmdlet)).toEqual(['Get-Mailbox', 'Get-MailboxStatistics']);
    expect(result(outcome)).toMatchObject({
      aliases: ['jd@contoso.com'],
      litigationHold: true,
      statistics: { totalSizeGb: 1.5, itemCount: 1200 },
    });
  });

  test('hold --on previews current hold, then sets duration and comment', async () => {
    const exo = fakeExo((step) =>
      step.cmdlet === 'Get-Mailbox' ? [{ PrimarySmtpAddress: 'jane@contoso.com', LitigationHoldEnabled: false }] : [],
    );
    const argv = ['jane', '--on', '--duration-days', '365', '--note', 'Case 42'];
    const first = await run('exchange hold', argv, [], { exo, allowWrites: true });
    if (first.outcome.kind !== 'confirm') throw new Error('expected confirm');
    expect(first.outcome.preview).toMatchObject({ current: { litigationHold: false } });
    expect(exo.steps.some(isWrite)).toBe(false);
    await run('exchange hold', argv, [], { exo, allowWrites: true, confirm: first.outcome.token });
    expect(exo.steps.find(isWrite)?.params).toEqual({
      Identity: 'jane',
      LitigationHoldEnabled: true,
      LitigationHoldDuration: 365,
      RetentionComment: 'Case 42',
    });
  });

  test('dl add-member reports already-members as skipped', async () => {
    const exo = fakeExo((step) => {
      if (step.params?.Member === 'bob') throw new Error('The recipient "bob" is already a member of the group.');
      return [];
    });
    const { outcome } = await run('exchange dl add-member', ['finance', 'jane', 'bob'], [], { exo, allowWrites: true });
    expect(exo.steps[0]).toMatchObject({
      cmdlet: 'Add-DistributionGroupMember',
      params: { Identity: 'finance', Member: 'jane', BypassSecurityGroupManagerCheck: true },
    });
    expect(result(outcome)).toMatchObject({
      members: [
        { member: 'jane', status: 'done' },
        { member: 'bob', status: 'skipped' },
      ],
    });
  });

  test('quarantine list filters by recipient and type and shapes rows', async () => {
    const exo = fakeExo(() => [
      {
        Identity: 'a\\b',
        SenderAddress: 'x@bad.com',
        RecipientAddress: ['jane@contoso.com'],
        Subject: 'Invoice',
        QuarantineTypes: 'Phish',
      },
    ]);
    const { outcome } = await run(
      'exchange quarantine list',
      ['--recipient', 'jane@contoso.com', '--type', 'Phish'],
      [],
      { exo },
    );
    expect(exo.steps[0]).toMatchObject({
      cmdlet: 'Get-QuarantineMessage',
      params: { RecipientAddress: ['jane@contoso.com'], QuarantineTypes: ['Phish'], PageSize: 51 },
    });
    expect(result<{ items: unknown[] }>(outcome).items[0]).toMatchObject({
      identity: 'a\\b',
      type: 'Phish',
      recipients: ['jane@contoso.com'],
    });
  });

  test('quarantine release previews the message and releases to all after confirm', async () => {
    const exo = fakeExo((step) =>
      step.cmdlet === 'Get-QuarantineMessage'
        ? [{ Identity: 'a\\b', Subject: 'Invoice', SenderAddress: 'x@v.com', QuarantineTypes: 'Spam' }]
        : [],
    );
    const first = await run('exchange quarantine release', ['a\\b'], [], { exo, allowWrites: true });
    if (first.outcome.kind !== 'confirm') throw new Error('expected confirm');
    expect((first.outcome.preview as { action: string }).action).toContain('Invoice');
    expect(exo.steps.some(isWrite)).toBe(false);
    await run('exchange quarantine release', ['a\\b'], [], { exo, allowWrites: true, confirm: first.outcome.token });
    expect(exo.steps.find(isWrite)).toEqual({
      cmdlet: 'Release-QuarantineMessage',
      params: { Identity: 'a\\b', ReleaseToAll: true, Confirm: false },
    });
  });
});

describe('exchange rooms', () => {
  const rooms = [
    { id: 'r1', displayName: 'Atlas', emailAddress: 'atlas@contoso.com', building: 'HQ', capacity: 10 },
    { id: 'r2', displayName: 'Boreas', emailAddress: 'boreas@contoso.com', building: 'HQ', capacity: 8 },
    { id: 'r3', displayName: 'Cosmo', emailAddress: 'cosmo@contoso.com', building: 'HQ', capacity: 4 },
    { id: 'r4', displayName: 'Delta', emailAddress: 'delta@contoso.com', building: 'Annex', capacity: 12 },
  ];
  const placesRoute: MockRoute = { path: '/places/microsoft.graph.room', reply: page(rooms) };

  test('rooms list filters by building and capacity', async () => {
    const { outcome } = await run('exchange rooms list', ['--building', 'hq', '--capacity', '8'], [placesRoute]);
    expect(result<{ items: { name: string }[] }>(outcome).items.map((r) => r.name)).toEqual(['Atlas', 'Boreas']);
  });

  test('rooms free returns only rooms with no busy time in the window', async () => {
    const { outcome, calls } = await run(
      'exchange rooms free',
      ['--start', '2026-11-12T14:00:00Z', '--end', '2026-11-12T15:00:00Z', '--building', 'HQ', '--capacity', '6'],
      [
        placesRoute,
        {
          method: 'POST',
          path: '/me/calendar/getSchedule',
          reply: {
            body: {
              value: [
                { scheduleId: 'atlas@contoso.com', scheduleItems: [{ status: 'busy' }] },
                { scheduleId: 'boreas@contoso.com', scheduleItems: [{ status: 'free' }] },
              ],
            },
          },
        },
      ],
    );
    const request = calls.find((c) => c.path === '/me/calendar/getSchedule');
    expect(request?.body).toMatchObject({
      schedules: ['atlas@contoso.com', 'boreas@contoso.com'],
      startTime: { dateTime: '2026-11-12T14:00:00.000', timeZone: 'UTC' },
    });
    expect(result(outcome)).toMatchObject({ checked: 2, freeCount: 1, free: [{ name: 'Boreas' }], busy: ['Atlas'] });
  });
});

describe('exchange rules audit batching', () => {
  test('25 mailboxes are read in two $batch calls (20 + 5), with per-mailbox errors', async () => {
    const users = Array.from({ length: 25 }, (_, i) => ({
      id: `u${i}`,
      displayName: `U${i}`,
      userPrincipalName: `u${i}@contoso.com`,
    }));
    const batchSizes: number[] = [];
    const { outcome, calls } = await run(
      'exchange rules audit',
      [],
      [
        { path: '/users', reply: page(users) },
        domainsRoute,
        {
          method: 'POST',
          path: '/$batch',
          reply: (req) => {
            const requests = (req.body as { requests: { id: string; url: string }[] }).requests;
            batchSizes.push(requests.length);
            return {
              body: {
                responses: requests.map((r) =>
                  r.url.startsWith('/users/u3/')
                    ? {
                        id: r.id,
                        status: 404,
                        body: { error: { code: 'MailboxNotEnabledForRESTAPI', message: 'no mailbox' } },
                      }
                    : { id: r.id, status: 200, body: { value: [] } },
                ),
              },
            };
          },
        },
      ],
    );
    expect(batchSizes).toEqual([20, 5]);
    expect(calls.some((c) => c.method === 'GET' && c.path.includes('/messageRules'))).toBe(false);
    expect(result(outcome)).toMatchObject({
      scanned: 25,
      flaggedCount: 0,
      errors: [{ user: 'u3@contoso.com', code: 'MailboxNotEnabledForRESTAPI' }],
    });
  });
});
