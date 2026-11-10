import { describe, expect, test } from 'bun:test';
import { commands, resolveEntraDevice, resolveServicePrincipal } from '../../src/domains/entra/directory.ts';
import type { CliError } from '../../src/core/errors.ts';
import type { Outcome } from '../../src/core/execute.ts';
import { commandById, graphError, mockContext, page, runCommand } from '../helpers/graphMock.ts';

const GROUP_ID = '11111111-1111-1111-1111-111111111111';
const USER_ID = '22222222-2222-2222-2222-222222222222';
const DEVICE_ID = '33333333-3333-3333-3333-333333333333';
const DEVICE_DEVICEID = '44444444-4444-4444-4444-444444444444';
const SP_ID = '55555555-5555-5555-5555-555555555555';
const GRAPH_SP_ID = '66666666-6666-6666-6666-666666666666';

const cmd = (id: string) => commandById(commands, `entra ${id}`);

function resultOf(outcome: Outcome): unknown {
  if (outcome.kind !== 'result') throw new Error(`expected a result, got ${outcome.kind}`);
  return outcome.result;
}

const jane = { id: USER_ID, displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com', mail: 'jane@contoso.com' };
const groupRoute = {
  path: `/groups/${GROUP_ID}`,
  reply: { body: { id: GROUP_ID, displayName: 'Finance Team', mail: null } },
};

describe('groups', () => {
  test('list filters by type and search with advanced query, derives type and Teams flag', async () => {
    const { outcome, calls } = await runCommand(
      cmd('groups list'),
      ['--type', 'm365', '--search', 'fin'],
      [
        {
          path: '/groups',
          reply: page([
            {
              id: 'g1',
              displayName: 'Finance',
              groupTypes: ['Unified', 'DynamicMembership'],
              mailEnabled: true,
              securityEnabled: false,
              mail: 'finance@contoso.com',
              membershipRule: 'user.department -eq "Finance"',
              resourceProvisioningOptions: ['Team'],
            },
          ]),
        },
      ],
    );
    expect(calls[0]?.query.$filter).toBe("groupTypes/any(t:t eq 'Unified')");
    expect(calls[0]?.query.$search).toBe('"displayName:fin"');
    expect(calls[0]?.headers.ConsistencyLevel).toBe('eventual');
    expect(resultOf(outcome)).toMatchObject({
      items: [
        {
          type: 'm365',
          membership: 'dynamic',
          membershipRule: 'user.department -eq "Finance"',
          isTeam: true,
          syncedFromOnPrem: false,
        },
      ],
    });
  });

  test('get adds owners, member count and license SKU names', async () => {
    const { outcome, calls } = await runCommand(
      cmd('groups get'),
      [GROUP_ID],
      [
        {
          path: `/groups/${GROUP_ID}`,
          reply: (req) =>
            req.query.$select?.includes('assignedLicenses')
              ? {
                  body: {
                    id: GROUP_ID,
                    displayName: 'Finance Team',
                    groupTypes: [],
                    securityEnabled: true,
                    mailEnabled: false,
                    assignedLicenses: [{ skuId: 'sku-e3', disabledPlans: [] }],
                  },
                }
              : { body: { id: GROUP_ID, displayName: 'Finance Team', mail: null } },
        },
        {
          path: `/groups/${GROUP_ID}/owners`,
          reply: page([{ id: USER_ID, displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com' }]),
        },
        { path: `/groups/${GROUP_ID}/members`, reply: { body: { '@odata.count': 42, value: [{ id: USER_ID }] } } },
        { path: '/subscribedSkus', reply: page([{ skuId: 'sku-e3', skuPartNumber: 'ENTERPRISEPACK' }]) },
      ],
    );
    expect(resultOf(outcome)).toMatchObject({
      type: 'security',
      memberCount: 42,
      owners: [{ displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com' }],
      licenses: [{ sku: 'ENTERPRISEPACK' }],
    });
    expect(calls.find((c) => c.path.endsWith('/members'))?.headers.ConsistencyLevel).toBe('eventual');
  });

  test('add-member resolves each member and reports already-present ones without failing the rest', async () => {
    const { outcome, calls } = await runCommand(
      cmd('groups add-member'),
      [GROUP_ID, 'jane@contoso.com', `device:${DEVICE_DEVICEID}`],
      [
        groupRoute,
        { path: '/users/jane@contoso.com', reply: { body: jane } },
        { path: `/devices/${DEVICE_DEVICEID}`, reply: graphError(404, 'Request_ResourceNotFound', 'nope') },
        {
          path: '/devices',
          reply: page([{ id: DEVICE_ID, deviceId: DEVICE_DEVICEID, displayName: 'FIN-LAPTOP-01' }]),
        },
        {
          method: 'POST',
          path: `/groups/${GROUP_ID}/members/$ref`,
          reply: (req) =>
            JSON.stringify(req.body).includes(USER_ID)
              ? graphError(
                  400,
                  'Request_BadRequest',
                  "One or more added object references already exist for the following modified properties: 'members'.",
                )
              : { status: 204 },
        },
      ],
      { allowWrites: true },
    );
    expect(calls.filter((c) => c.method === 'POST').map((c) => c.body)).toEqual([
      { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${USER_ID}` },
      { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${DEVICE_ID}` },
    ]);
    expect(resultOf(outcome)).toMatchObject({
      results: [
        { type: 'user', status: 'alreadyMember' },
        { type: 'device', status: 'added' },
      ],
    });
  });

  test('add-member with a name matching both a user and a device is AMBIGUOUS', async () => {
    const error = (await runCommand(
      cmd('groups add-member'),
      [GROUP_ID, 'Reception'],
      [
        groupRoute,
        {
          path: '/users',
          reply: page([{ id: USER_ID, displayName: 'Reception', userPrincipalName: 'reception@contoso.com' }]),
        },
        { path: '/groups', reply: page([]) },
        { path: '/devices', reply: page([{ id: DEVICE_ID, displayName: 'Reception', deviceId: DEVICE_DEVICEID }]) },
      ],
      { allowWrites: true },
    ).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
    expect(error.hint).toContain('device:');
  });

  test('remove-member previews, then deletes the member reference once confirmed', async () => {
    const routes = [
      groupRoute,
      { path: `/users/${USER_ID}`, reply: { body: jane } },
      { path: `/groups/${USER_ID}`, reply: graphError(404, 'Request_ResourceNotFound', 'nope') },
      { method: 'DELETE', path: `/groups/${GROUP_ID}/members/${USER_ID}/$ref`, reply: { status: 204 } },
    ];
    const preview = await runCommand(cmd('groups remove-member'), [GROUP_ID, USER_ID], routes, { allowWrites: true });
    expect(preview.outcome.kind).toBe('confirm');
    expect(preview.calls.some((c) => c.method === 'DELETE')).toBe(false);
    const token = preview.outcome.kind === 'confirm' ? preview.outcome.token : '';
    const { calls } = await runCommand(cmd('groups remove-member'), [GROUP_ID, USER_ID], routes, {
      allowWrites: true,
      confirm: token,
    });
    expect(calls.at(-1)).toMatchObject({ method: 'DELETE', path: `/groups/${GROUP_ID}/members/${USER_ID}/$ref` });
  });

  test('create builds a Microsoft 365 group body with owner binding and derived nickname', async () => {
    const { outcome, calls } = await runCommand(
      cmd('groups create'),
      ['--name', 'Project Phoenix!', '--type', 'm365', '--owner', 'jane@contoso.com'],
      [
        { path: '/users/jane@contoso.com', reply: { body: jane } },
        {
          method: 'POST',
          path: '/groups',
          reply: (req) => ({ status: 201, body: { id: 'new', ...(req.body as object) } }),
        },
      ],
      { allowWrites: true },
    );
    expect(calls.at(-1)?.body).toEqual({
      displayName: 'Project Phoenix!',
      mailNickname: 'ProjectPhoenix',
      groupTypes: ['Unified'],
      mailEnabled: true,
      securityEnabled: false,
      'owners@odata.bind': [`https://graph.microsoft.com/v1.0/users/${USER_ID}`],
    });
    expect(resultOf(outcome)).toMatchObject({ id: 'new', type: 'm365' });
  });

  test('writes are refused while writes are disabled', async () => {
    const error = (await runCommand(cmd('groups remove-member'), [GROUP_ID, USER_ID], []).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(error.code).toBe('WRITES_DISABLED');
  });
});

describe('licenses, sign-ins, audit', () => {
  test('licenses computes availability and flags low stock', async () => {
    const { outcome } = await runCommand(
      cmd('licenses'),
      [],
      [
        {
          path: '/subscribedSkus',
          reply: page([
            {
              skuId: 'a',
              skuPartNumber: 'SPE_E3',
              consumedUnits: 98,
              prepaidUnits: { enabled: 100, suspended: 0, warning: 0 },
              appliesTo: 'User',
            },
            {
              skuId: 'b',
              skuPartNumber: 'EMS',
              consumedUnits: 10,
              prepaidUnits: { enabled: 50, suspended: 5, warning: 0 },
              appliesTo: 'User',
            },
          ]),
        },
      ],
    );
    expect(resultOf(outcome)).toMatchObject({
      items: [
        expect.objectContaining({ skuPartNumber: 'EMS', available: 40, suspended: 5, lowOnLicenses: false }),
        expect.objectContaining({ skuPartNumber: 'SPE_E3', available: 2, lowOnLicenses: true }),
      ],
    });
  });

  test('signins combines filters and resolves the user to an id', async () => {
    const { calls } = await runCommand(
      cmd('signins'),
      [
        '--user',
        'jane@contoso.com',
        '--failures-only',
        '--error-code',
        'AADSTS50126',
        '--country',
        'us',
        '--ip',
        '203.0.113.',
      ],
      [
        { path: '/users/jane@contoso.com', reply: { body: jane } },
        { path: '/auditLogs/signIns', reply: page([]) },
      ],
    );
    const filter = calls.at(-1)?.query.$filter ?? '';
    expect(filter).toStartWith('createdDateTime ge ');
    expect(filter).toContain(`userId eq '${USER_ID}'`);
    expect(filter).toContain('status/errorCode eq 50126');
    expect(filter).toContain('status/errorCode ne 0');
    expect(filter).toContain("location/countryOrRegion eq 'US'");
    expect(filter).toContain("startswith(ipAddress, '203.0.113.')");
  });

  test('signin-error explains known codes, links unknown ones and rejects junk without calling Graph', async () => {
    const known = resultOf((await runCommand(cmd('signin-error'), ['AADSTS50126'], [])).outcome);
    expect(known).toMatchObject({
      code: 50126,
      known: true,
      meaning: expect.any(String),
      lookupUrl: 'https://login.microsoftonline.com/error?code=50126',
    });
    const unknown = resultOf((await runCommand(cmd('signin-error'), ['99999999'], [])).outcome);
    expect(unknown).toMatchObject({ code: 99999999, known: false });
    const catalog = resultOf((await runCommand(cmd('signin-error'), [], [])).outcome);
    expect(catalog).toMatchObject({ items: expect.arrayContaining([expect.objectContaining({ code: 50126 })]) });
    const error = (await runCommand(cmd('signin-error'), ['oops'], []).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('USAGE');
  });

  test('audit filters by target user id and activity prefix and shapes initiator and targets', async () => {
    const { outcome, calls } = await runCommand(
      cmd('audit'),
      ['--target', 'jane@contoso.com', '--activity', 'Reset password', '--user', 'admin@contoso.com'],
      [
        { path: '/users/jane@contoso.com', reply: { body: jane } },
        {
          path: '/auditLogs/directoryAudits',
          reply: page([
            {
              id: 'a1',
              activityDateTime: '2026-11-09T10:00:00Z',
              activityDisplayName: 'Reset password (by admin)',
              category: 'UserManagement',
              result: 'success',
              correlationId: 'c1',
              initiatedBy: { user: { userPrincipalName: 'admin@contoso.com', ipAddress: '10.0.0.1' }, app: null },
              targetResources: [
                {
                  id: USER_ID,
                  type: 'User',
                  userPrincipalName: 'jane@contoso.com',
                  modifiedProperties: [{ displayName: 'PasswordProfile' }],
                },
              ],
            },
          ]),
        },
      ],
    );
    const filter = calls.at(-1)?.query.$filter ?? '';
    expect(filter).toContain("initiatedBy/user/userPrincipalName eq 'admin@contoso.com'");
    expect(filter).toContain(`targetResources/any(t:t/id eq '${USER_ID}')`);
    expect(filter).toContain("startswith(activityDisplayName, 'Reset password')");
    expect(resultOf(outcome)).toEqual({
      truncated: false,
      items: [
        {
          time: '2026-11-09T10:00:00Z',
          activity: 'Reset password (by admin)',
          category: 'UserManagement',
          result: 'success',
          initiatedBy: 'admin@contoso.com',
          initiatedFromIp: '10.0.0.1',
          targets: [{ type: 'User', name: 'jane@contoso.com', modified: ['PasswordProfile'] }],
          correlationId: 'c1',
        },
      ],
    });
  });
});

describe('roles', () => {
  test('list groups active assignments by role name', async () => {
    const { outcome } = await runCommand(
      cmd('roles list'),
      ['--role', 'global'],
      [
        {
          path: '/roleManagement/directory/roleAssignments',
          reply: page([
            {
              id: 'r1',
              principalId: USER_ID,
              roleDefinitionId: 'ga',
              directoryScopeId: '/',
              principal: {
                '@odata.type': '#microsoft.graph.user',
                id: USER_ID,
                displayName: 'Jane Doe',
                userPrincipalName: 'jane@contoso.com',
              },
            },
            {
              id: 'r2',
              principalId: 'sp',
              roleDefinitionId: 'ga',
              directoryScopeId: '/',
              principal: { '@odata.type': '#microsoft.graph.servicePrincipal', id: 'sp', displayName: 'Automation' },
            },
            { id: 'r3', principalId: USER_ID, roleDefinitionId: 'ua', directoryScopeId: '/administrativeUnits/x' },
          ]),
        },
        {
          path: '/roleManagement/directory/roleDefinitions',
          reply: page([
            { id: 'ga', displayName: 'Global Administrator', templateId: 'ga' },
            { id: 'ua', displayName: 'User Administrator', templateId: 'ua' },
          ]),
        },
      ],
    );
    expect(resultOf(outcome)).toEqual({
      truncated: false,
      items: [
        {
          role: 'Global Administrator',
          roleId: 'ga',
          memberCount: 2,
          members: [
            { type: 'user', displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com' },
            { type: 'servicePrincipal', displayName: 'Automation' },
          ],
        },
      ],
    });
  });

  test('user shows active roles and degrades the PIM section when forbidden', async () => {
    const { outcome, calls } = await runCommand(
      cmd('roles user'),
      ['jane@contoso.com'],
      [
        { path: '/users/jane@contoso.com', reply: { body: jane } },
        {
          path: '/roleManagement/directory/roleAssignments',
          reply: page([
            {
              id: 'r1',
              principalId: USER_ID,
              roleDefinitionId: 'ga',
              directoryScopeId: '/',
              roleDefinition: { id: 'ga', displayName: 'Global Administrator' },
            },
          ]),
        },
        {
          path: '/roleManagement/directory/roleEligibilityScheduleInstances',
          reply: graphError(403, 'Forbidden', 'needs P2'),
        },
      ],
    );
    expect(calls.find((c) => c.path.endsWith('roleAssignments'))?.query.$filter).toBe(`principalId eq '${USER_ID}'`);
    expect(resultOf(outcome)).toMatchObject({
      active: [{ role: 'Global Administrator' }],
      eligible: { error: { code: 'FORBIDDEN' } },
    });
  });
});

describe('apps', () => {
  test('resolveServicePrincipal falls back from object id to appId', async () => {
    const { ctx, calls } = mockContext([
      { path: `/servicePrincipals/${SP_ID}`, reply: graphError(404, 'Request_ResourceNotFound', 'nope') },
      { path: '/servicePrincipals', reply: page([{ id: 'sp-object', appId: SP_ID, displayName: 'Zoom' }]) },
    ]);
    expect((await resolveServicePrincipal(ctx, SP_ID)).id).toBe('sp-object');
    expect(calls[1]?.query.$filter).toBe(`appId eq '${SP_ID}'`);
  });

  test('get resolves application permission names and delegated consent', async () => {
    const { outcome } = await runCommand(
      cmd('apps get'),
      [SP_ID],
      [
        {
          path: `/servicePrincipals/${SP_ID}`,
          reply: {
            body: {
              id: SP_ID,
              appId: 'app-1',
              displayName: 'Backup Tool',
              appOwnerOrganizationId: 'other',
              verifiedPublisher: { displayName: null },
              publisherName: 'Backup Inc',
            },
          },
        },
        {
          path: `/servicePrincipals/${SP_ID}/appRoleAssignments`,
          reply: page([
            {
              appRoleId: 'role-mail',
              resourceId: GRAPH_SP_ID,
              resourceDisplayName: 'Microsoft Graph',
              createdDateTime: '2026-01-01T00:00:00Z',
            },
          ]),
        },
        {
          path: `/servicePrincipals/${SP_ID}/oauth2PermissionGrants`,
          reply: page([{ resourceId: GRAPH_SP_ID, scope: ' User.Read offline_access ', consentType: 'AllPrincipals' }]),
        },
        {
          path: `/servicePrincipals/${GRAPH_SP_ID}`,
          reply: {
            body: {
              id: GRAPH_SP_ID,
              displayName: 'Microsoft Graph',
              appRoles: [{ id: 'role-mail', value: 'Mail.Read' }],
            },
          },
        },
      ],
    );
    expect(resultOf(outcome)).toMatchObject({
      displayName: 'Backup Tool',
      publisher: 'Backup Inc',
      publisherVerified: false,
      isMicrosoftApp: false,
      applicationPermissions: [{ resource: 'Microsoft Graph', permission: 'Mail.Read' }],
      delegatedPermissions: [
        { resource: 'Microsoft Graph', scopes: ['User.Read', 'offline_access'], consent: 'admin' },
      ],
    });
  });

  test('expiring lists expired and soon-expiring credentials, soonest first', async () => {
    const day = 86_400_000;
    const at = (days: number) => new Date(Date.now() + days * day + 3_600_000).toISOString();
    const { outcome } = await runCommand(
      cmd('apps expiring'),
      ['--days', '30'],
      [
        {
          path: '/applications',
          reply: page([
            {
              id: 'a',
              appId: 'app-a',
              displayName: 'Payroll Sync',
              passwordCredentials: [
                { keyId: 'k1', displayName: 'prod', hint: 'abc', endDateTime: at(10) },
                { keyId: 'k2', displayName: 'later', hint: 'def', endDateTime: at(200) },
              ],
              keyCredentials: [{ keyId: 'k3', displayName: 'CN=old', type: 'AsymmetricX509Cert', endDateTime: at(-5) }],
            },
          ]),
        },
      ],
    );
    expect(resultOf(outcome)).toMatchObject({
      truncated: false,
      items: [
        { keyId: 'k3', credentialType: 'certificate', daysLeft: -5, expired: true },
        { keyId: 'k1', credentialType: 'secret', daysLeft: 10, expired: false },
      ],
    });
  });
});

describe('conditional access', () => {
  const policy = {
    id: '77777777-7777-7777-7777-777777777777',
    displayName: 'Require MFA for admins',
    state: 'enabledForReportingButNotEnforced',
    conditions: {
      users: {
        includeUsers: [],
        excludeUsers: [USER_ID],
        includeGroups: [],
        excludeGroups: [],
        includeRoles: ['62e90394-69f5-4237-9190-012177145e10'],
        excludeRoles: [],
      },
      applications: { includeApplications: ['All'], excludeApplications: [] },
      clientAppTypes: ['all'],
      locations: null,
      platforms: null,
      signInRiskLevels: [],
      userRiskLevels: [],
    },
    grantControls: { operator: 'OR', builtInControls: ['mfa', 'compliantDevice'] },
    sessionControls: { signInFrequency: { isEnabled: true, value: 4, type: 'hours' } },
  };

  test('list summarizes state, scope and controls', async () => {
    const { outcome } = await runCommand(
      cmd('ca list'),
      [],
      [{ path: '/identity/conditionalAccess/policies', reply: page([policy]) }],
    );
    expect(resultOf(outcome)).toMatchObject({
      items: [
        {
          name: 'Require MFA for admins',
          state: 'report-only',
          users: { exclude: '1 specific', includeRoles: '1 specific' },
          apps: { include: 'All' },
          clientAppTypes: ['all'],
          grant: 'mfa OR compliantDevice',
          session: ['signInFrequency: 4 hours'],
        },
      ],
    });
  });

  test('get finds a policy by partial name and names excluded users and roles', async () => {
    const { outcome, calls } = await runCommand(
      cmd('ca get'),
      ['mfa for admins'],
      [
        {
          path: '/identity/conditionalAccess/policies',
          reply: page([policy, { ...policy, id: 'other', displayName: 'Block legacy auth' }]),
        },
        {
          method: 'POST',
          path: '/directoryObjects/getByIds',
          reply: {
            body: { value: [{ id: USER_ID, displayName: 'Break Glass', userPrincipalName: 'bg@contoso.com' }] },
          },
        },
        {
          path: '/roleManagement/directory/roleDefinitions',
          reply: page([
            { id: 'ga-def', templateId: '62e90394-69f5-4237-9190-012177145e10', displayName: 'Global Administrator' },
          ]),
        },
      ],
    );
    expect(calls.find((c) => c.method === 'POST')?.body).toEqual({ ids: [USER_ID], types: ['user', 'group'] });
    expect(resultOf(outcome)).toMatchObject({
      resolved: { excludeUsers: ['bg@contoso.com'], includeRoles: ['Global Administrator'], includeApps: ['All'] },
    });
  });
});

describe('identity protection', () => {
  test('risky-users filters by level and state', async () => {
    const { calls } = await runCommand(
      cmd('risky-users'),
      ['--level', 'high', '--state', 'atRisk'],
      [{ path: '/identityProtection/riskyUsers', reply: page([]) }],
    );
    expect(calls[0]?.query.$filter).toBe("riskLevel eq 'high' and riskState eq 'atRisk'");
  });

  test('confirm-compromised previews current risk, then posts the user id once confirmed', async () => {
    const routes = [
      { path: '/users/jane@contoso.com', reply: { body: jane } },
      {
        path: `/identityProtection/riskyUsers/${USER_ID}`,
        reply: { body: { id: USER_ID, riskLevel: 'medium', riskState: 'atRisk' } },
      },
      { method: 'POST', path: '/identityProtection/riskyUsers/confirmCompromised', reply: { status: 204 } },
    ];
    const first = await runCommand(cmd('risky-users confirm-compromised'), ['jane@contoso.com'], routes, {
      allowWrites: true,
    });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.method === 'POST')).toBe(false);
    if (first.outcome.kind !== 'confirm') return;
    expect(first.outcome.preview).toMatchObject({ risk: { riskLevel: 'medium', riskState: 'atRisk' } });
    const second = await runCommand(cmd('risky-users confirm-compromised'), ['jane@contoso.com'], routes, {
      allowWrites: true,
      confirm: first.outcome.token,
    });
    expect(second.calls.at(-1)).toMatchObject({ method: 'POST', body: { userIds: [USER_ID] } });
  });

  test('risk-detections shapes location and filters by user and level', async () => {
    const { outcome, calls } = await runCommand(
      cmd('risk-detections'),
      ['--user', 'jane@contoso.com', '--level', 'high'],
      [
        { path: '/users/jane@contoso.com', reply: { body: jane } },
        {
          path: '/identityProtection/riskDetections',
          reply: page([
            {
              id: 'd1',
              detectedDateTime: '2026-11-01T00:00:00Z',
              userPrincipalName: 'jane@contoso.com',
              riskEventType: 'unfamiliarFeatures',
              riskLevel: 'high',
              riskState: 'atRisk',
              ipAddress: '198.51.100.1',
              location: { city: 'Lagos', state: null, countryOrRegion: 'NG' },
            },
          ]),
        },
      ],
    );
    expect(calls.at(-1)?.query.$filter).toContain(`userId eq '${USER_ID}' and riskLevel eq 'high'`);
    expect(resultOf(outcome)).toMatchObject({
      items: [{ riskEventType: 'unfamiliarFeatures', location: 'Lagos, NG' }],
    });
  });
});

describe('entra devices', () => {
  const laptop = {
    id: DEVICE_ID,
    deviceId: DEVICE_DEVICEID,
    displayName: 'FIN-LAPTOP-01',
    operatingSystem: 'Windows',
    accountEnabled: true,
    trustType: 'AzureAd',
    approximateLastSignInDateTime: new Date(Date.now() - 100 * 86_400_000).toISOString(),
  };

  test('resolveEntraDevice accepts a deviceId and reports duplicate names as AMBIGUOUS', async () => {
    const { ctx } = mockContext([
      { path: `/devices/${DEVICE_DEVICEID}`, reply: graphError(404, 'Request_ResourceNotFound', 'nope') },
      {
        path: '/devices',
        reply: (req) =>
          req.query.$filter?.startsWith('deviceId')
            ? page([laptop])
            : page([laptop, { ...laptop, id: 'dup', deviceId: 'old' }]),
      },
    ]);
    expect((await resolveEntraDevice(ctx, DEVICE_DEVICEID)).id).toBe(DEVICE_ID);
    const error = (await resolveEntraDevice(ctx, 'FIN-LAPTOP-01').catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
  });

  test('list finds stale disabled Windows devices and derives join type and staleness', async () => {
    const { outcome, calls } = await runCommand(
      cmd('devices list'),
      ['--os', 'Windows', '--stale-days', '90', '--disabled'],
      [{ path: '/devices', reply: page([laptop]) }],
    );
    const filter = calls[0]?.query.$filter ?? '';
    expect(filter).toContain("startswith(operatingSystem, 'Windows')");
    expect(filter).toContain('approximateLastSignInDateTime le ');
    expect(filter).toContain('accountEnabled eq false');
    expect(resultOf(outcome)).toMatchObject({ items: [{ join: 'Entra joined', daysSinceSignIn: 100 }] });
  });

  test('disable previews the device, then patches accountEnabled false once confirmed', async () => {
    const routes = [
      { path: '/devices', reply: page([laptop]) },
      { method: 'PATCH', path: `/devices/${DEVICE_ID}`, reply: { status: 204 } },
    ];
    const first = await runCommand(cmd('devices disable'), ['FIN-LAPTOP-01'], routes, { allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.method === 'PATCH')).toBe(false);
    if (first.outcome.kind !== 'confirm') return;
    expect(first.outcome.preview).toMatchObject({ device: { displayName: 'FIN-LAPTOP-01', accountEnabled: true } });
    const second = await runCommand(cmd('devices disable'), ['FIN-LAPTOP-01'], routes, {
      allowWrites: true,
      confirm: first.outcome.token,
    });
    expect(second.calls.at(-1)).toMatchObject({ method: 'PATCH', body: { accountEnabled: false } });
  });
});

describe('tenant', () => {
  test('flags stale directory sync and lists domain authentication', async () => {
    const { outcome } = await runCommand(
      cmd('tenant'),
      [],
      [
        {
          path: '/organization',
          reply: page([
            {
              id: 'tenant-1',
              displayName: 'Contoso',
              verifiedDomains: [
                { name: 'contoso.onmicrosoft.com', isInitial: true, isDefault: false },
                { name: 'contoso.com', isInitial: false, isDefault: true },
              ],
              technicalNotificationMails: ['it@contoso.com'],
              onPremisesSyncEnabled: true,
              onPremisesLastSyncDateTime: new Date(Date.now() - 5 * 3_600_000).toISOString(),
            },
          ]),
        },
        {
          path: '/domains',
          reply: page([
            {
              id: 'contoso.com',
              isVerified: true,
              isDefault: true,
              authenticationType: 'Federated',
              supportedServices: ['Email'],
            },
          ]),
        },
      ],
    );
    expect(resultOf(outcome)).toMatchObject({
      defaultDomain: 'contoso.com',
      initialDomain: 'contoso.onmicrosoft.com',
      hoursSinceLastDirSync: 5,
      dirSyncStale: true,
      domains: [{ name: 'contoso.com', authentication: 'federated', isDefault: true }],
    });
  });
});
