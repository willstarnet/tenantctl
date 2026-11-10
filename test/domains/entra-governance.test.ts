import { describe, expect, test } from 'bun:test';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import type { CliError } from '../../src/core/errors.ts';
import type { Outcome } from '../../src/core/execute.ts';
import { COMMANDS } from '../../src/registry.ts';
import { shapeSubscription } from '../../src/domains/entra/governance.ts';
import {
  commandById,
  graphError,
  type MockReply,
  type MockRequest,
  page,
  runCommand,
  tempConfigDir,
} from '../helpers/graphMock.ts';

const cmd = (id: string) => commandById(COMMANDS, `entra ${id}`);

const USER_ID = '22222222-2222-2222-2222-222222222222';
const GROUP_ID = '11111111-1111-1111-1111-111111111111';
const SP_ID = '55555555-5555-5555-5555-555555555555';
const GRAPH_SP_ID = '66666666-6666-6666-6666-666666666666';
const MAIL_READ_ROLE = '810c84a8-4a9e-49e6-bf7d-12d183f40d01';
const DAY = 86_400_000;

function resultOf(outcome: Outcome): unknown {
  if (outcome.kind !== 'result') throw new Error(`expected a result, got ${outcome.kind}`);
  return outcome.result;
}

interface Row {
  line: number;
  status: string;
  errors?: string[];
  temporaryPassword?: string;
  followUpErrors?: string[];
}
interface BulkPreview {
  preview: { summary: unknown; rows: Row[] };
}
interface BulkResult {
  summary: unknown;
  results: Row[];
}

/** Find a row by CSV line, failing the test when it is missing. */
function rowAt(rows: Row[], line: number): Row {
  const row = rows.find((r) => r.line === line);
  if (!row) throw new Error(`no row for line ${line}`);
  return row;
}

function csv(name: string, text: string): string {
  const path = join(tempConfigDir(), name);
  writeFileSync(path, text);
  return path;
}

interface BatchItem {
  id: string;
  method: string;
  url: string;
  body?: { userPrincipalName?: string; '@odata.id'?: string; passwordProfile?: { password?: string } };
}

/** A `$batch` route answering each inner request; records every inner request in `seen`. */
function batchRoute(answer: (item: BatchItem) => { status: number; body?: unknown }, seen: BatchItem[] = []) {
  return {
    method: 'POST',
    path: '/$batch',
    reply: (req: MockRequest): MockReply => {
      // Body shape is fixed by GraphClient.batch, which built it.
      const payload = req.body as { requests: BatchItem[] };
      const requests = payload.requests;
      const decoded = requests.map((r) => ({ ...r, url: decodeURIComponent(r.url) }));
      seen.push(...decoded);
      return { body: { responses: decoded.map((r) => ({ id: r.id, ...answer(r) })) } };
    },
  };
}

const orgRoute = {
  path: '/organization',
  reply: page([{ verifiedDomains: [{ name: 'contoso.com' }, { name: 'contoso.onmicrosoft.com' }] }]),
};

describe('users invite', () => {
  test('posts the invitation body, then adds the guest to the group', async () => {
    const { outcome, calls } = await runCommand(
      cmd('users invite'),
      ['pat@fabrikam.com', '--name', 'Pat Partner', '--message', 'Welcome to Project X', '--group', GROUP_ID],
      [
        orgRoute,
        { path: `/groups/${GROUP_ID}`, reply: { body: { id: GROUP_ID, displayName: 'Project X', mail: null } } },
        {
          method: 'POST',
          path: '/invitations',
          reply: {
            status: 201,
            body: {
              id: 'inv-1',
              status: 'PendingAcceptance',
              inviteRedeemUrl: 'https://login.microsoftonline.com/redeem?rd=x',
              invitedUser: { id: USER_ID, userPrincipalName: 'pat_fabrikam.com#EXT#@contoso.onmicrosoft.com' },
            },
          },
        },
        { method: 'POST', path: `/groups/${GROUP_ID}/members/$ref`, reply: { status: 204 } },
      ],
      { allowWrites: true },
    );
    const invite = calls.find((c) => c.method === 'POST' && c.path === '/invitations');
    expect(invite?.body).toEqual({
      invitedUserEmailAddress: 'pat@fabrikam.com',
      inviteRedirectUrl: 'https://myapps.microsoft.com',
      sendInvitationMessage: true,
      invitedUserDisplayName: 'Pat Partner',
      invitedUserMessageInfo: { customizedMessageBody: 'Welcome to Project X' },
    });
    const add = calls.find((c) => c.path === `/groups/${GROUP_ID}/members/$ref`);
    expect(add?.body).toEqual({ '@odata.id': expect.stringMatching(new RegExp(`/directoryObjects/${USER_ID}$`)) });
    expect(resultOf(outcome)).toMatchObject({
      id: USER_ID,
      userPrincipalName: 'pat_fabrikam.com#EXT#@contoso.onmicrosoft.com',
      status: 'PendingAcceptance',
      inviteRedeemUrl: 'https://login.microsoftonline.com/redeem?rd=x',
      emailSent: true,
      group: { displayName: 'Project X', status: 'added' },
    });
  });

  test('--no-email sends no invitation message', async () => {
    const { calls } = await runCommand(
      cmd('users invite'),
      ['pat@fabrikam.com', '--no-email'],
      [orgRoute, { method: 'POST', path: '/invitations', reply: { body: { invitedUser: { id: USER_ID } } } }],
      { allowWrites: true },
    );
    expect(calls.find((c) => c.path === '/invitations')?.body).toMatchObject({ sendInvitationMessage: false });
  });

  test('refuses an address in a verified tenant domain without inviting', async () => {
    let caught: CliError | undefined;
    const run = runCommand(cmd('users invite'), ['Jane@Contoso.com'], [orgRoute], { allowWrites: true });
    await run.catch((error: CliError) => {
      caught = error;
    });
    expect(caught?.code).toBe('USAGE');
    expect(caught?.message).toContain('internal user');
    expect(caught?.hint).toContain('users create');
  });
});

describe('users bulk-create', () => {
  const skus = page([{ skuId: 'sku-e3', skuPartNumber: 'SPE_E3', consumedUnits: 5, prepaidUnits: { enabled: 10 } }]);
  const groupsRoute = {
    path: '/groups',
    reply: (req: MockRequest): MockReply =>
      req.query.$filter?.includes("'Finance Team'")
        ? page([{ id: GROUP_ID, displayName: 'Finance Team', mail: null }])
        : page([]),
  };
  const bob = { id: 'bob-id', displayName: 'Bob Ray', userPrincipalName: 'bob@contoso.com', mail: null };
  const file = () =>
    csv(
      'users.csv',
      [
        'DisplayName,UserPrincipalName,UsageLocation,License,Manager,Groups',
        'Jane Doe,jane@contoso.com,us,SPE_E3,bob@contoso.com,Finance Team',
        ',nameless@contoso.com,,,,',
        'Jane Again,jane@contoso.com,,,,',
        'Sam Sku,sam@contoso.com,GB,NOPE_SKU,,',
        'Old Timer,old@contoso.com,,,,',
        'Gail Group,gail@contoso.com,,,,Finance Team;Missing Group',
        'Mia Mgr,mia@contoso.com,,,jane@contoso.com,',
      ].join('\n'),
    );
  const existing = new Set(['old@contoso.com']);

  test('dry run validates every row and creates nothing', async () => {
    const seen: BatchItem[] = [];
    const { outcome, calls } = await runCommand(
      cmd('users bulk-create'),
      [file()],
      [
        batchRoute((item) => {
          const upn = item.url.split('?')[0]?.replace('/users/', '') ?? '';
          return existing.has(upn)
            ? { status: 200, body: { id: 'old-id', userPrincipalName: upn } }
            : { status: 404, body: { error: { code: 'Request_ResourceNotFound', message: 'not found' } } };
        }, seen),
        { path: '/subscribedSkus', reply: skus },
        { path: '/users/bob@contoso.com', reply: { body: bob } },
        groupsRoute,
      ],
      { dryRun: true },
    );
    expect(calls.some((c) => c.method === 'POST' && c.path === '/users')).toBe(false);
    expect(seen.every((item) => item.method === 'GET')).toBe(true);
    const { preview } = resultOf(outcome) as BulkPreview;
    expect(preview.summary).toEqual({ rows: 7, valid: 2, skipped: 1, invalid: 4 });
    const row = (line: number) => rowAt(preview.rows, line);
    expect(row(2)).toMatchObject({
      status: 'valid',
      license: 'SPE_E3',
      manager: 'bob@contoso.com',
      groups: ['Finance Team'],
      password: 'generated when run',
    });
    expect(row(3).errors).toContain('DisplayName is empty');
    expect(row(4).errors?.[0]).toContain('duplicate UserPrincipalName (also on line 2)');
    expect(row(5).errors?.[0]).toContain('License "NOPE_SKU" is not a SKU');
    expect(row(6)).toMatchObject({ status: 'skipped', existingId: 'old-id' });
    expect(row(7).errors?.[0]).toContain('Group "Missing Group"');
    // A manager that is another row in the file is accepted.
    expect(row(8)).toMatchObject({ status: 'valid', manager: 'jane@contoso.com' });
  });

  test('creates valid rows via $batch, skips existing UPNs and reports per-row results', async () => {
    const seen: BatchItem[] = [];
    const { outcome } = await runCommand(
      cmd('users bulk-create'),
      [file()],
      [
        batchRoute((item) => {
          if (item.method === 'GET') {
            const upn = item.url.split('?')[0]?.replace('/users/', '') ?? '';
            return existing.has(upn)
              ? { status: 200, body: { id: 'old-id' } }
              : { status: 404, body: { error: { code: 'Request_ResourceNotFound', message: 'not found' } } };
          }
          if (item.method === 'POST' && item.url === '/users') {
            return { status: 201, body: { id: `id-${item.body?.userPrincipalName?.split('@')[0]}` } };
          }
          if (item.url.endsWith('/assignLicense')) {
            return { status: 400, body: { error: { code: 'Request_BadRequest', message: 'No seats' } } };
          }
          return { status: 204 };
        }, seen),
        { path: '/subscribedSkus', reply: skus },
        { path: '/users/bob@contoso.com', reply: { body: bob } },
        groupsRoute,
      ],
      { allowWrites: true },
    );
    const creates = seen.filter((i) => i.method === 'POST' && i.url === '/users');
    expect(creates.map((i) => i.body?.userPrincipalName)).toEqual(['jane@contoso.com', 'mia@contoso.com']);
    expect(creates[0]?.body).toMatchObject({
      accountEnabled: true,
      displayName: 'Jane Doe',
      mailNickname: 'jane',
      usageLocation: 'US',
      passwordProfile: { forceChangePasswordNextSignIn: true },
    });
    // Mia's manager is Jane, created in the same run.
    const managerPut = seen.find((i) => i.method === 'PUT' && i.url === '/users/id-mia/manager/$ref');
    expect(managerPut?.body?.['@odata.id']).toEndWith('/directoryObjects/id-jane');
    expect(seen.some((i) => i.url === `/groups/${GROUP_ID}/members/$ref`)).toBe(true);

    const result = resultOf(outcome) as BulkResult;
    expect(result.summary).toEqual({ created: 2, skipped: 1, failed: 4 });
    const row = (line: number) => rowAt(result.results, line);
    expect(row(2)).toMatchObject({ status: 'created', id: 'id-jane' });
    expect(row(2).temporaryPassword).toHaveLength(16);
    expect(row(2).temporaryPassword).toBe(creates[0]?.body?.passwordProfile?.password);
    expect(row(2).followUpErrors).toEqual(['license SPE_E3: POST /users/id-jane/assignLicense: No seats']);
    expect(row(6)).toMatchObject({ status: 'skipped', userPrincipalName: 'old@contoso.com' });
    expect(row(3)).toMatchObject({ status: 'failed', error: 'DisplayName is empty' });
  });
});

describe('groups add-members', () => {
  test('adds CSV members through $batch and maps existing members to alreadyMember', async () => {
    const seen: BatchItem[] = [];
    const path = csv('members.csv', 'UserPrincipalName\njane@contoso.com\nbob@contoso.com\nghost@contoso.com\n');
    const ids: Record<string, string> = { 'jane@contoso.com': 'jane-id', 'bob@contoso.com': 'bob-id' };
    const { outcome, calls } = await runCommand(
      cmd('groups add-members'),
      [GROUP_ID, '--from-csv', path],
      [
        { path: `/groups/${GROUP_ID}`, reply: { body: { id: GROUP_ID, displayName: 'Finance Team', mail: null } } },
        batchRoute((item) => {
          if (item.method === 'GET') {
            const upn = item.url.split('?')[0]?.replace('/users/', '') ?? '';
            const id = ids[upn];
            return id
              ? { status: 200, body: { id, displayName: upn, userPrincipalName: upn } }
              : { status: 404, body: { error: { code: 'Request_ResourceNotFound', message: 'missing' } } };
          }
          return item.body?.['@odata.id']?.endsWith('/bob-id')
            ? {
                status: 400,
                body: {
                  error: {
                    code: 'Request_BadRequest',
                    message: 'One or more added object references already exist for the following modified properties.',
                  },
                },
              }
            : { status: 204 };
        }, seen),
        { path: '/users', reply: page([]) },
      ],
      { allowWrites: true },
    );
    const posts = seen.filter((i) => i.method === 'POST');
    expect(posts.map((i) => i.url)).toEqual([`/groups/${GROUP_ID}/members/$ref`, `/groups/${GROUP_ID}/members/$ref`]);
    expect(calls.filter((c) => c.path === '/$batch')).toHaveLength(2);
    const result = resultOf(outcome) as BulkResult;
    expect(result.summary).toEqual({ added: 1, alreadyMember: 1, failed: 1 });
    expect(result.results.map((r) => [r.line, r.status])).toEqual([
      [2, 'added'],
      [3, 'alreadyMember'],
      [4, 'failed'],
    ]);
  });
});

describe('apps consents and revoke-consent', () => {
  const routes = [
    {
      path: `/servicePrincipals/${SP_ID}`,
      reply: { body: { id: SP_ID, appId: 'app-guid', displayName: 'Shady Mail', accountEnabled: true } },
    },
    {
      path: `/servicePrincipals/${GRAPH_SP_ID}`,
      reply: {
        body: {
          id: GRAPH_SP_ID,
          displayName: 'Microsoft Graph',
          appRoles: [{ id: MAIL_READ_ROLE, value: 'Mail.Read', displayName: 'Read mail in all mailboxes' }],
        },
      },
    },
    {
      path: '/oauth2PermissionGrants',
      reply: page([
        {
          id: 'grant-admin',
          clientId: SP_ID,
          resourceId: GRAPH_SP_ID,
          consentType: 'AllPrincipals',
          scope: 'User.Read Mail.Read',
        },
        {
          id: 'grant-user',
          clientId: SP_ID,
          resourceId: GRAPH_SP_ID,
          consentType: 'Principal',
          principalId: USER_ID,
          scope: 'Files.Read',
        },
      ]),
    },
    {
      method: 'POST',
      path: '/directoryObjects/getByIds',
      reply: { body: { value: [{ id: USER_ID, displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com' }] } },
    },
    {
      path: `/servicePrincipals/${SP_ID}/appRoleAssignments`,
      reply: page([
        { id: 'assign-1', appRoleId: MAIL_READ_ROLE, resourceId: GRAPH_SP_ID, resourceDisplayName: 'Microsoft Graph' },
      ]),
    },
  ];

  test('consents lists delegated grants with principals and named application permissions', async () => {
    const { outcome, calls } = await runCommand(cmd('apps consents'), [SP_ID], routes);
    expect(calls.find((c) => c.path === '/oauth2PermissionGrants')?.query.$filter).toBe(`clientId eq '${SP_ID}'`);
    expect(resultOf(outcome)).toMatchObject({
      summary: { adminConsentedGrants: 1, userConsentedGrants: 1, applicationPermissions: 1 },
      delegated: [
        { grantId: 'grant-admin', consent: 'admin', resource: 'Microsoft Graph', scopes: ['User.Read', 'Mail.Read'] },
        {
          grantId: 'grant-user',
          consent: 'user',
          principal: { userPrincipalName: 'jane@contoso.com' },
          scopes: ['Files.Read'],
        },
      ],
      application: [
        { assignmentId: 'assign-1', permission: 'Mail.Read', permissionName: 'Read mail in all mailboxes' },
      ],
    });
  });

  test('revoke-consent previews exactly what goes, then deletes it once confirmed', async () => {
    const seen: BatchItem[] = [];
    const all = [...routes, batchRoute(() => ({ status: 204 }), seen)];
    const argv = [SP_ID, '--all-delegated', '--all-application'];
    const first = await runCommand(cmd('apps revoke-consent'), argv, all, { allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.method === 'DELETE' || c.path === '/$batch')).toBe(false);
    if (first.outcome.kind !== 'confirm') return;
    const preview = first.outcome.preview as { action: string; delegated: unknown[]; application: unknown[] };
    expect(preview.action).toContain('Revoke from app "Shady Mail"');
    expect(preview.action).toContain('Microsoft Graph: User.Read Mail.Read [all users]');
    expect(preview.action).toContain('Microsoft Graph: Mail.Read');
    expect(preview.delegated).toHaveLength(2);

    const second = await runCommand(cmd('apps revoke-consent'), argv, all, {
      allowWrites: true,
      confirm: first.outcome.token,
    });
    expect(seen.map((i) => `${i.method} ${i.url}`)).toEqual([
      'DELETE /oauth2PermissionGrants/grant-admin',
      'DELETE /oauth2PermissionGrants/grant-user',
      `DELETE /servicePrincipals/${SP_ID}/appRoleAssignments/assign-1`,
    ]);
    expect(resultOf(second.outcome)).toMatchObject({
      delegated: [{ status: 'revoked' }, { status: 'revoked' }],
      application: [{ permission: 'Mail.Read', status: 'revoked' }],
      next: [`m365 entra apps disable "${SP_ID}"`, expect.any(String), expect.any(String)],
    });
  });

  test('revoke-consent rejects an id the app does not hold', async () => {
    let caught: CliError | undefined;
    await runCommand(cmd('apps revoke-consent'), [SP_ID, '--grant', 'nope'], routes, { allowWrites: true }).catch(
      (error: CliError) => {
        caught = error;
      },
    );
    expect(caught?.code).toBe('NOT_FOUND');
    expect(caught?.hint).toContain('grant-admin');
  });

  test('apps disable previews, then patches accountEnabled false', async () => {
    const all = [...routes, { method: 'PATCH', path: `/servicePrincipals/${SP_ID}`, reply: { status: 204 } }];
    const first = await runCommand(cmd('apps disable'), [SP_ID], all, { allowWrites: true });
    if (first.outcome.kind !== 'confirm') throw new Error('expected confirm');
    const second = await runCommand(cmd('apps disable'), [SP_ID], all, {
      allowWrites: true,
      confirm: first.outcome.token,
    });
    expect(second.calls.find((c) => c.method === 'PATCH')?.body).toEqual({ accountEnabled: false });
    expect(resultOf(second.outcome)).toMatchObject({ status: 'disabled' });
  });
});

describe('subscriptions and partners', () => {
  test('renewalSoon flags subscriptions within 30 days of their next lifecycle step', () => {
    const now = Date.parse('2026-01-01T00:00:00Z');
    const at = (days: number) => new Date(now + days * DAY).toISOString();
    expect(shapeSubscription({ id: 'a', status: 'Enabled', nextLifecycleDateTime: at(12) }, now)).toMatchObject({
      daysToNextLifecycle: 12,
      renewalSoon: true,
    });
    expect(shapeSubscription({ id: 'b', status: 'Enabled', nextLifecycleDateTime: at(200) }, now)).toMatchObject({
      daysToNextLifecycle: 200,
      renewalSoon: false,
    });
    expect(shapeSubscription({ id: 'c', status: 'Deleted', nextLifecycleDateTime: at(3) }, now).renewalSoon).toBe(
      false,
    );
    expect(shapeSubscription({ id: 'd', status: 'Enabled', nextLifecycleDateTime: null }, now)).toMatchObject({
      daysToNextLifecycle: null,
      renewalSoon: false,
    });
  });

  test('subscriptions command sorts by next lifecycle and filters --renewal-soon', async () => {
    const now = Date.now();
    const { outcome } = await runCommand(
      cmd('subscriptions'),
      ['--renewal-soon'],
      [
        {
          path: '/directory/subscriptions',
          reply: page([
            {
              id: '1',
              skuPartNumber: 'SPE_E3',
              status: 'Enabled',
              totalLicenses: 25,
              nextLifecycleDateTime: new Date(now + 300 * DAY).toISOString(),
            },
            {
              id: '2',
              skuPartNumber: 'POWER_BI_PRO',
              status: 'Warning',
              isTrial: true,
              totalLicenses: 5,
              nextLifecycleDateTime: new Date(now + 10.5 * DAY).toISOString(),
            },
          ]),
        },
      ],
    );
    expect(resultOf(outcome)).toEqual({
      truncated: false,
      items: [
        expect.objectContaining({
          skuPartNumber: 'POWER_BI_PRO',
          daysToNextLifecycle: 10,
          renewalSoon: true,
          isTrial: true,
        }),
      ],
    });
  });

  test('partners resolves role template ids and derives expiry', async () => {
    const now = Date.now();
    const { outcome } = await runCommand(
      cmd('partners'),
      [],
      [
        {
          path: '/tenantRelationships/delegatedAdminRelationships',
          reply: page([
            {
              id: 'rel-1',
              displayName: 'Contoso GDAP',
              status: 'active',
              duration: 'P730D',
              autoExtendDuration: 'PT0S',
              endDateTime: new Date(now + 20.5 * DAY).toISOString(),
              customer: { displayName: 'Contoso Inc', tenantId: 'tenant-1' },
              accessDetails: { unifiedRoles: [{ roleDefinitionId: '729827e3-9c14-49f7-bb1b-9608f156bbb8' }] },
            },
          ]),
        },
        {
          path: '/roleManagement/directory/roleDefinitions',
          reply: page([{ id: '729827e3-9c14-49f7-bb1b-9608f156bbb8', displayName: 'Helpdesk Administrator' }]),
        },
      ],
    );
    expect(resultOf(outcome)).toMatchObject({
      items: [
        {
          customer: 'Contoso Inc',
          roles: ['Helpdesk Administrator'],
          durationDays: 730,
          daysLeft: 20,
          autoExtend: false,
          expiringSoon: true,
        },
      ],
    });
  });

  test('access-reviews list surfaces a missing license or permission as FORBIDDEN', async () => {
    let caught: CliError | undefined;
    await runCommand(
      cmd('access-reviews list'),
      [],
      [{ path: '/identityGovernance/accessReviews/definitions', reply: graphError(403, 'Forbidden', 'needs P2') }],
    ).catch((error: CliError) => {
      caught = error;
    });
    expect(caught?.code).toBe('FORBIDDEN');
  });
});
