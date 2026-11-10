import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands, getUserOverview, shapeSignin } from '../../src/domains/entra/users.ts';
import { commandById, graphError, type MockRoute, mockContext, page, runCommand } from '../helpers/graphMock.ts';

const cmd = (name: string) => commandById(commands, `entra users ${name}`);

const JANE = { id: 'u1', displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com', mail: 'jane@contoso.com' };
const janeRoute = (extra: Record<string, unknown> = {}): MockRoute => ({
  path: '/users/jane@contoso.com',
  reply: { body: { ...JANE, ...extra } },
});

const SKUS = [
  {
    skuId: 'sku-e3',
    skuPartNumber: 'SPE_E3',
    consumedUnits: 9,
    prepaidUnits: { enabled: 10 },
    servicePlans: [
      { servicePlanId: 'plan-teams', servicePlanName: 'TEAMS1' },
      { servicePlanId: 'plan-sway', servicePlanName: 'SWAY' },
    ],
  },
  { skuId: 'sku-p1', skuPartNumber: 'AAD_PREMIUM', consumedUnits: 0, prepaidUnits: { enabled: 5 }, servicePlans: [] },
];
const skuRoute: MockRoute = { path: '/subscribedSkus', reply: page(SKUS) };

async function confirmAndRun(name: string, argv: string[], routes: MockRoute[]) {
  const first = await runCommand(cmd(name), argv, routes, { allowWrites: true });
  expect(first.outcome.kind).toBe('confirm');
  expect(first.calls.every((c) => c.method === 'GET')).toBe(true);
  const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
  const second = await runCommand(cmd(name), argv, routes, { allowWrites: true, confirm: token });
  return { preview: first.outcome.kind === 'confirm' ? first.outcome.preview : undefined, ...second };
}

describe('users list', () => {
  test('combines filters, selects licenses and derives licensed', async () => {
    const { outcome, calls } = await runCommand(
      cmd('list'),
      ['--department', "O'Neil Sales", '--disabled', '--guests'],
      [
        {
          path: '/users',
          reply: page([
            { ...JANE, accountEnabled: false, assignedLicenses: [{ skuId: 'sku-e3' }] },
            { ...JANE, id: 'u2', assignedLicenses: [] },
          ]),
        },
      ],
    );
    expect(calls[0]?.query.$filter).toBe(
      "department eq 'O''Neil Sales' and accountEnabled eq false and userType eq 'Guest'",
    );
    expect(calls[0]?.query.$select).toContain('assignedLicenses');
    expect(calls[0]?.headers.ConsistencyLevel).toBeUndefined();
    const result = outcome.kind === 'result' ? (outcome.result as { items: { licensed: boolean }[] }) : undefined;
    expect(result?.items.map((u) => u.licensed)).toEqual([true, false]);
  });

  test('--unlicensed and --search use advanced queries', async () => {
    const { calls } = await runCommand(
      cmd('list'),
      ['--unlicensed', '--search', 'jane'],
      [{ path: '/users', reply: page([]) }],
    );
    expect(calls[0]?.query.$filter).toBe('assignedLicenses/$count eq 0');
    expect(calls[0]?.query.$search).toBe('"displayName:jane" OR "userPrincipalName:jane" OR "mail:jane"');
    expect(calls[0]?.query.$count).toBe('true');
    expect(calls[0]?.headers.ConsistencyLevel).toBe('eventual');
  });

  test('--inactive-days filters on sign-in activity and refuses other filters', async () => {
    const { outcome, calls } = await runCommand(
      cmd('list'),
      ['--inactive-days', '90'],
      [{ path: '/users', reply: page([{ ...JANE, signInActivity: { lastSignInDateTime: '2026-01-01T00:00:00Z' } }]) }],
    );
    expect(calls[0]?.query.$filter).toMatch(/^signInActivity\/lastSignInDateTime le \d{4}-/);
    expect(calls[0]?.query.$select).toContain('signInActivity');
    expect(outcome).toMatchObject({ result: { items: [{ lastSignIn: '2026-01-01T00:00:00Z' }] } });
    const error = (await runCommand(cmd('list'), ['--inactive-days', '90', '--department', 'Sales'], []).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(error.code).toBe('USAGE');
  });
});

describe('users get', () => {
  const routes = (signIn: MockRoute['reply']): MockRoute[] => [
    janeRoute({
      accountEnabled: true,
      onPremisesSyncEnabled: true,
      onPremisesLastSyncDateTime: '2026-11-09T00:00:00Z',
      lastPasswordChangeDateTime: new Date(Date.now() - 10 * 86_400_000).toISOString(),
      assignedLicenses: [
        { skuId: 'sku-e3', disabledPlans: [] },
        { skuId: 'sku-unknown', disabledPlans: [] },
      ],
    }),
    {
      path: '/users/u1/manager',
      reply: { body: { id: 'm1', displayName: 'Boss', userPrincipalName: 'boss@contoso.com' } },
    },
    { path: '/users/u1', reply: signIn },
    skuRoute,
  ];

  test('derives sync, password age, license names and last sign-in', async () => {
    const lastSuccess = new Date(Date.now() - 3 * 86_400_000).toISOString();
    const { ctx, calls } = mockContext(
      routes({
        body: { signInActivity: { lastSignInDateTime: lastSuccess, lastSuccessfulSignInDateTime: lastSuccess } },
      }),
    );
    const overview = await getUserOverview(ctx, 'jane@contoso.com');
    expect(overview).toMatchObject({
      id: 'u1',
      isSynced: true,
      onPremisesLastSyncDateTime: '2026-11-09T00:00:00Z',
      daysSincePasswordChange: 10,
      daysSinceLastSignIn: 3,
      licenses: ['SPE_E3', 'sku-unknown'],
      manager: { displayName: 'Boss', userPrincipalName: 'boss@contoso.com' },
    });
    expect(calls.find((c) => c.path === '/users/u1')?.query.$select).toBe('signInActivity');
  });

  test('missing AuditLog permission degrades only the sign-in section', async () => {
    const { ctx } = mockContext(routes(graphError(403, 'Authorization_RequestDenied', 'Insufficient privileges')));
    const overview = (await getUserOverview(ctx, 'jane@contoso.com')) as Record<string, unknown>;
    expect(overview.signIns).toMatchObject({ error: { code: 'FORBIDDEN' } });
    expect(overview.daysSinceLastSignIn).toBeUndefined();
    expect(overview.licenses).toEqual(['SPE_E3', 'sku-unknown']);
  });

  test('a display name matching two users is AMBIGUOUS', async () => {
    const error = (await runCommand(
      cmd('get'),
      ['Ja'],
      [
        {
          path: '/users',
          reply: page([JANE, { ...JANE, id: 'u2', displayName: 'Jane Roe', userPrincipalName: 'jroe@contoso.com' }]),
        },
      ],
    ).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
  });
});

describe('users groups', () => {
  test('splits groups by kind and directory roles', async () => {
    const { outcome } = await runCommand(
      cmd('groups'),
      ['jane@contoso.com'],
      [
        janeRoute(),
        {
          path: '/users/u1/transitiveMemberOf',
          reply: page([
            {
              '@odata.type': '#microsoft.graph.group',
              id: 'g1',
              displayName: 'Sales',
              groupTypes: ['Unified'],
              mail: 'sales@contoso.com',
            },
            {
              '@odata.type': '#microsoft.graph.group',
              id: 'g2',
              displayName: 'VPN',
              groupTypes: ['DynamicMembership'],
              securityEnabled: true,
              mailEnabled: false,
            },
            {
              '@odata.type': '#microsoft.graph.directoryRole',
              id: 'r1',
              displayName: 'Helpdesk Administrator',
              roleTemplateId: 't1',
            },
          ]),
        },
      ],
    );
    expect(outcome).toMatchObject({
      result: {
        groups: [
          { displayName: 'Sales', type: 'Microsoft 365', dynamic: false },
          { displayName: 'VPN', type: 'Security', dynamic: true, mail: null },
        ],
        directoryRoles: [{ displayName: 'Helpdesk Administrator' }],
      },
    });
  });
});

describe('users licenses', () => {
  test('explains direct vs group assignments, errors and disabled plans', async () => {
    const { outcome } = await runCommand(
      cmd('licenses'),
      ['jane@contoso.com'],
      [
        janeRoute({
          usageLocation: null,
          licenseAssignmentStates: [
            { skuId: 'sku-e3', assignedByGroup: null, state: 'Active', error: 'None', disabledPlans: ['plan-sway'] },
            { skuId: 'sku-p1', assignedByGroup: 'g9', state: 'Error', error: 'CountViolation', disabledPlans: [] },
          ],
        }),
        skuRoute,
        {
          path: '/users/u1/licenseDetails',
          reply: page([
            {
              skuId: 'sku-e3',
              servicePlans: [
                { servicePlanName: 'TEAMS1', provisioningStatus: 'PendingProvisioning' },
                { servicePlanName: 'SWAY', provisioningStatus: 'Disabled' },
              ],
            },
          ]),
        },
        { path: '/groups/g9', reply: { body: { id: 'g9', displayName: 'P1 Users' } } },
      ],
    );
    const result = outcome.kind === 'result' ? (outcome.result as Record<string, unknown>) : {};
    expect(result.assignments).toEqual([
      {
        sku: 'SPE_E3',
        skuId: 'sku-e3',
        assignedVia: 'direct',
        state: 'Active',
        error: null,
        disabledPlans: ['SWAY'],
        plansNotReady: [{ plan: 'TEAMS1', status: 'PendingProvisioning' }],
      },
      expect.objectContaining({
        sku: 'AAD_PREMIUM',
        assignedVia: 'group',
        group: { id: 'g9', displayName: 'P1 Users' },
        error: 'CountViolation',
        explanation: expect.stringContaining('No licenses left'),
      }),
    ]);
    expect(result.warnings).toEqual([expect.stringContaining('usageLocation'), expect.stringContaining('errors')]);
  });
});

describe('users auth-methods', () => {
  test('shapes methods by type and degrades the registration report', async () => {
    const { outcome } = await runCommand(
      cmd('auth-methods'),
      ['jane@contoso.com'],
      [
        janeRoute(),
        {
          path: '/users/u1/authentication/methods',
          reply: page([
            {
              '@odata.type': '#microsoft.graph.passwordAuthenticationMethod',
              id: '28c10230-6103-485e-b985-444c60001490',
            },
            {
              '@odata.type': '#microsoft.graph.phoneAuthenticationMethod',
              id: 'p1',
              phoneType: 'mobile',
              phoneNumber: '+1 555XXXX12',
              smsSignInState: 'notEnabled',
            },
            {
              '@odata.type': '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod',
              id: 'a1',
              displayName: 'Pixel 8',
              phoneAppVersion: '6.2',
            },
          ]),
        },
        {
          path: '/reports/authenticationMethods/userRegistrationDetails/u1',
          reply: graphError(403, 'Forbidden', 'no'),
        },
      ],
    );
    expect(outcome).toMatchObject({
      result: {
        methods: [
          { type: 'Password' },
          { id: 'p1', type: 'Phone', detail: 'mobile: +1 555XXXX12' },
          { id: 'a1', type: 'Microsoft Authenticator', detail: 'Pixel 8, app 6.2' },
        ],
        hasStrongMethod: true,
        registration: { error: { code: 'FORBIDDEN' } },
      },
    });
  });
});

describe('users devices', () => {
  test('merges owned and registered devices without duplicates', async () => {
    const pc = {
      '@odata.type': '#microsoft.graph.device',
      id: 'd1',
      displayName: 'PC-1',
      trustType: 'AzureAd',
      isCompliant: true,
    };
    const { outcome } = await runCommand(
      cmd('devices'),
      ['jane@contoso.com'],
      [
        janeRoute(),
        { path: '/users/u1/ownedDevices', reply: page([pc]) },
        {
          path: '/users/u1/registeredDevices',
          reply: page([pc, { ...pc, id: 'd2', displayName: 'Phone', trustType: 'Workplace' }]),
        },
      ],
    );
    expect(outcome).toMatchObject({
      result: {
        items: [
          { id: 'd1', joinType: 'Entra joined', relationship: ['owner', 'registered'] },
          { id: 'd2', joinType: 'Entra registered', relationship: ['registered'] },
        ],
      },
    });
  });
});

describe('users signins', () => {
  test('filters by user, window and failures; explains error codes', async () => {
    const { outcome, calls } = await runCommand(
      cmd('signins'),
      ['jane@contoso.com', '--failures-only', '--days', '3'],
      [
        janeRoute(),
        {
          path: '/auditLogs/signIns',
          reply: page([
            {
              id: 's1',
              createdDateTime: '2026-11-09T10:00:00Z',
              appDisplayName: 'Outlook',
              clientAppUsed: 'Mobile Apps and Desktop clients',
              ipAddress: '203.0.113.5',
              location: { city: 'Leeds', countryOrRegion: 'GB' },
              deviceDetail: { displayName: '', operatingSystem: 'iOS', browser: '', isCompliant: false },
              status: { errorCode: 53000, failureReason: 'Device is not in required device state' },
              conditionalAccessStatus: 'failure',
              authenticationRequirement: 'multiFactorAuthentication',
              riskLevelDuringSignIn: 'none',
            },
          ]),
        },
      ],
    );
    const filter = calls.find((c) => c.path === '/auditLogs/signIns')?.query.$filter ?? '';
    expect(filter).toStartWith("userId eq 'u1' and createdDateTime ge ");
    expect(filter).toEndWith(' and status/errorCode ne 0');
    expect(outcome).toMatchObject({
      result: {
        items: [
          {
            result: 'failure',
            errorCode: 53000,
            location: 'Leeds, GB',
            device: { os: 'iOS', isCompliant: false },
            explanation: { name: 'DeviceNotCompliant' },
            conditionalAccess: 'failure',
            mfaRequired: true,
            mfaSatisfied: false,
          },
        ],
      },
    });
  });

  test('shapeSignin links unknown codes to the Microsoft lookup and marks success', () => {
    expect(shapeSignin({ id: 'x', createdDateTime: 't', status: { errorCode: 999999 } }).explanation).toEqual({
      lookup: 'https://login.microsoftonline.com/error?code=999999',
    });
    const ok = shapeSignin({ id: 'y', createdDateTime: 't', status: { errorCode: 0 } });
    expect(ok.result).toBe('success');
    expect(ok.errorCode).toBeUndefined();
  });
});

describe('users create / update', () => {
  test('create generates a password, defaults the alias and sets the manager', async () => {
    const { outcome, calls } = await runCommand(
      cmd('create'),
      [
        '--display-name',
        'Bob Ray',
        '--upn',
        'bob.ray@contoso.com',
        '--usage-location',
        'gb',
        '--manager',
        'jane@contoso.com',
      ],
      [
        janeRoute(),
        {
          method: 'POST',
          path: '/users',
          reply: {
            status: 201,
            body: { id: 'u9', displayName: 'Bob Ray', userPrincipalName: 'bob.ray@contoso.com', mail: null },
          },
        },
        { method: 'PUT', path: '/users/u9/manager/$ref', reply: { status: 204 } },
      ],
      { allowWrites: true },
    );
    const post = calls.find((c) => c.method === 'POST');
    const body = post?.body as {
      mailNickname: string;
      usageLocation: string;
      passwordProfile: { password: string; forceChangePasswordNextSignIn: boolean };
    };
    expect(body.mailNickname).toBe('bob.ray');
    expect(body.usageLocation).toBe('GB');
    expect(body.passwordProfile.forceChangePasswordNextSignIn).toBe(true);
    expect(body.passwordProfile.password).toMatch(/^(?=.*[A-Z])(?=.*[a-z])(?=.*\d)(?=.*[^A-Za-z0-9]).{16}$/);
    expect(outcome).toMatchObject({
      result: { temporaryPassword: body.passwordProfile.password, manager: { id: 'u1' } },
    });
    expect(calls.find((c) => c.method === 'PUT')?.body).toEqual({
      '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/u1',
    });
  });

  test('update sends only provided fields and refuses an empty update', async () => {
    const { calls } = await runCommand(
      cmd('update'),
      ['jane@contoso.com', '--job-title', 'Lead', '--usage-location', 'us'],
      [janeRoute(), { method: 'PATCH', path: '/users/u1', reply: { status: 204 } }],
      { allowWrites: true },
    );
    expect(calls.find((c) => c.method === 'PATCH')?.body).toEqual({ jobTitle: 'Lead', usageLocation: 'US' });
    const error = (await runCommand(cmd('update'), ['jane@contoso.com'], [janeRoute()], { allowWrites: true }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(error.code).toBe('USAGE');
  });
});

describe('destructive user commands preview first', () => {
  test('disable previews current state, then patches accountEnabled', async () => {
    const { preview, outcome, calls } = await confirmAndRun(
      'disable',
      ['jane@contoso.com'],
      [
        janeRoute({ accountEnabled: true, onPremisesSyncEnabled: false }),
        { method: 'PATCH', path: '/users/u1', reply: { status: 204 } },
      ],
    );
    expect(preview).toMatchObject({
      accountEnabled: true,
      isSynced: false,
      action: expect.stringContaining('Block sign-in'),
    });
    expect(calls.find((c) => c.method === 'PATCH')?.body).toEqual({ accountEnabled: false });
    expect(outcome).toMatchObject({ result: { accountEnabled: false } });
  });

  test('reset-password returns the generated password it set', async () => {
    const { outcome, calls } = await confirmAndRun(
      'reset-password',
      ['jane@contoso.com'],
      [janeRoute(), { method: 'PATCH', path: '/users/u1', reply: { status: 204 } }],
    );
    const body = calls.find((c) => c.method === 'PATCH')?.body as {
      passwordProfile: { password: string; forceChangePasswordNextSignIn: boolean };
    };
    expect(body.passwordProfile.forceChangePasswordNextSignIn).toBe(true);
    expect(outcome).toMatchObject({ result: { temporaryPassword: body.passwordProfile.password } });
  });

  test('revoke-sessions and delete hit the right endpoints', async () => {
    const revoke = await confirmAndRun(
      'revoke-sessions',
      ['jane@contoso.com'],
      [janeRoute(), { method: 'POST', path: '/users/u1/revokeSignInSessions', reply: { body: { value: true } } }],
    );
    expect(revoke.calls.some((c) => c.method === 'POST' && c.path === '/users/u1/revokeSignInSessions')).toBe(true);
    const del = await confirmAndRun(
      'delete',
      ['jane@contoso.com'],
      [janeRoute(), { method: 'DELETE', path: '/users/u1', reply: { status: 204 } }],
    );
    expect(del.preview).toMatchObject({ action: expect.stringContaining('restored for 30 days') });
    expect(del.calls.some((c) => c.method === 'DELETE' && c.path === '/users/u1')).toBe(true);
  });

  test('tap posts lifetime and one-time flag and returns the pass', async () => {
    const { outcome, calls } = await confirmAndRun(
      'tap',
      ['jane@contoso.com', '--lifetime-minutes', '120', '--one-time'],
      [
        janeRoute(),
        {
          method: 'POST',
          path: '/users/u1/authentication/temporaryAccessPassMethods',
          reply: {
            status: 201,
            body: {
              id: 't1',
              temporaryAccessPass: '+drkzqAD',
              startDateTime: '2026-11-10T00:00:00Z',
              lifetimeInMinutes: 120,
              isUsableOnce: true,
              isUsable: true,
            },
          },
        },
      ],
    );
    expect(calls.find((c) => c.method === 'POST')?.body).toEqual({ lifetimeInMinutes: 120, isUsableOnce: true });
    expect(outcome).toMatchObject({
      result: { temporaryAccessPass: '+drkzqAD', expiresDateTime: '2026-11-10T02:00:00.000Z' },
    });
  });

  test('auth-method remove deletes through the type-specific endpoint and warns when no MFA remains', async () => {
    const { preview, calls } = await confirmAndRun(
      'auth-method remove',
      ['jane@contoso.com', 'a1'],
      [
        janeRoute(),
        {
          path: '/users/u1/authentication/methods',
          reply: page([
            { '@odata.type': '#microsoft.graph.passwordAuthenticationMethod', id: 'pw' },
            {
              '@odata.type': '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod',
              id: 'a1',
              displayName: 'Old phone',
            },
          ]),
        },
        { method: 'DELETE', path: '/users/u1/authentication/microsoftAuthenticatorMethods/a1', reply: { status: 204 } },
      ],
    );
    expect(preview).toMatchObject({
      method: { type: 'Microsoft Authenticator' },
      warning: expect.stringContaining('no MFA method'),
    });
    expect(
      calls.some(
        (c) => c.method === 'DELETE' && c.path === '/users/u1/authentication/microsoftAuthenticatorMethods/a1',
      ),
    ).toBe(true);
  });
});

describe('licenses', () => {
  test('license add resolves SKU and plan names, and requires a usage location', async () => {
    const routes = (usageLocation: string | null): MockRoute[] => [
      janeRoute({ usageLocation, assignedLicenses: [], licenseAssignmentStates: [] }),
      skuRoute,
      { method: 'POST', path: '/users/u1/assignLicense', reply: { body: JANE } },
    ];
    const argv = ['jane@contoso.com', '--sku', 'spe_e3', '--disable-plans', 'sway'];
    const blocked = (await runCommand(cmd('license add'), argv, routes(null), { allowWrites: true }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(blocked.code).toBe('USAGE');
    expect(blocked.hint).toContain('--usage-location');

    const dry = await runCommand(cmd('license add'), argv, routes(null), { dryRun: true });
    expect(dry.outcome).toMatchObject({
      result: { preview: { sku: { availableUnits: 1 }, warnings: [expect.stringContaining('usageLocation')] } },
    });

    const { calls } = await runCommand(cmd('license add'), argv, routes('US'), { allowWrites: true });
    expect(calls.find((c) => c.method === 'POST')?.body).toEqual({
      addLicenses: [{ skuId: 'sku-e3', disabledPlans: ['plan-sway'] }],
      removeLicenses: [],
    });
  });

  test('license remove refuses group-inherited licenses and removes direct ones', async () => {
    const routes = (assignedByGroup: string | null): MockRoute[] => [
      janeRoute({
        licenseAssignmentStates: [
          { skuId: 'sku-e3', assignedByGroup, state: 'Active', error: null, disabledPlans: [] },
        ],
      }),
      skuRoute,
      { method: 'POST', path: '/users/u1/assignLicense', reply: { body: JANE } },
    ];
    const inherited = (await runCommand(cmd('license remove'), ['jane@contoso.com', '--sku', 'SPE_E3'], routes('g1'), {
      allowWrites: true,
    }).catch((e: unknown) => e)) as CliError;
    expect(inherited.code).toBe('USAGE');
    expect(inherited.hint).toContain('g1');

    const { calls } = await confirmAndRun('license remove', ['jane@contoso.com', '--sku', 'SPE_E3'], routes(null));
    expect(calls.find((c) => c.method === 'POST')?.body).toEqual({ addLicenses: [], removeLicenses: ['sku-e3'] });
  });
});

describe('deleted users', () => {
  test('restore finds a deleted user by original UPN', async () => {
    const deleted = {
      id: 'u7',
      displayName: 'Old User',
      userPrincipalName: 'a1b2c3old@contoso.com',
      mail: null,
      deletedDateTime: '2026-11-01T00:00:00Z',
    };
    const { calls } = await runCommand(
      cmd('restore'),
      ['old@contoso.com'],
      [
        { path: '/directory/deletedItems/microsoft.graph.user', reply: page([deleted]) },
        {
          method: 'POST',
          path: '/directory/deletedItems/u7/restore',
          reply: { body: { ...deleted, userPrincipalName: 'old@contoso.com' } },
        },
      ],
      { allowWrites: true },
    );
    expect(calls.some((c) => c.method === 'POST' && c.path === '/directory/deletedItems/u7/restore')).toBe(true);
  });

  test('deleted lists days until purge', async () => {
    const { outcome } = await runCommand(
      cmd('deleted'),
      [],
      [
        {
          path: '/directory/deletedItems/microsoft.graph.user',
          reply: page([
            {
              id: 'u7',
              displayName: 'Old',
              userPrincipalName: 'x',
              deletedDateTime: new Date(Date.now() - 12 * 86_400_000).toISOString(),
            },
          ]),
        },
      ],
    );
    expect(outcome).toMatchObject({ result: { items: [{ id: 'u7', daysUntilPurge: 18 }] } });
  });
});
