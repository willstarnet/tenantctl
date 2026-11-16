import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands } from '../../src/domains/entra/index.ts';
import { cidrContains } from '../../src/domains/entra/policies.ts';
import { commandById, graphError, type MockRoute, mockContext, page, runCommand } from '../helpers/graphMock.ts';

const cmd = (id: string) => commandById(commands, `entra ${id}`);
const CA = '/identity/conditionalAccess/policies';
const LOCATIONS = '/identity/conditionalAccess/namedLocations';

const blockPolicy = {
  id: '11111111-0000-0000-0000-000000000001',
  displayName: 'CA010-AllUsers-DeviceCode-Block',
  state: 'enabledForReportingButNotEnforced',
  conditions: { users: { includeUsers: ['All'], excludeUsers: [] }, locations: { excludeLocations: ['loc-hq'] } },
  grantControls: { operator: 'OR', builtInControls: ['block'] },
};

describe('cidrContains', () => {
  test('matches IPv4 and compressed IPv6 inside a range, never across families', () => {
    expect(cidrContains('203.0.113.252/30', '203.0.113.255')).toBe(true);
    expect(cidrContains('203.0.113.252/30', '203.0.113.251')).toBe(false);
    expect(cidrContains('2001:db8:a0::/48', '2001:db8:a0:ffff::1')).toBe(true);
    expect(cidrContains('2001:db8:a0::/48', '2001:db8:a1::1')).toBe(false);
    expect(cidrContains('203.0.113.0/24', '::ffff:1')).toBe(false);
  });
});

describe('entra ca set-state', () => {
  const routes = (): MockRoute[] => [
    { path: CA, reply: page([blockPolicy]) },
    { method: 'PATCH', path: `${CA}/${blockPolicy.id}`, reply: { status: 204 } },
  ];

  test('previews the change with a warning for a block-everyone policy, then PATCHes only the state', async () => {
    const argv = ['CA010', '--state', 'on'];
    const first = await runCommand(cmd('ca set-state'), argv, routes(), { allowWrites: true });
    const preview = (first.outcome as { preview: { action: string; warnings: string[] } }).preview;
    expect(preview.action).toBe(
      'Change Conditional Access policy "CA010-AllUsers-DeviceCode-Block" from report-only to enabled',
    );
    expect(preview.warnings).toContain(
      'Blocks sign-in for every user it matches; make sure break-glass accounts are excluded.',
    );
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('ca set-state'), argv, routes(), { allowWrites: true, confirm: token });
    const patch = second.calls.find((c) => c.method === 'PATCH');
    expect(patch?.body).toEqual({ state: 'enabled' });
  });

  test('sends nothing when the policy is already in that state', async () => {
    const argv = ['CA010', '--state', 'report-only'];
    const first = await runCommand(cmd('ca set-state'), argv, routes(), { allowWrites: true });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('ca set-state'), argv, routes(), { allowWrites: true, confirm: token });
    expect(second.calls.some((c) => c.method === 'PATCH')).toBe(false);
    expect(second.outcome).toMatchObject({ kind: 'result', result: { status: 'unchanged' } });
  });
});

describe('entra ca locations', () => {
  const routes: MockRoute[] = [
    {
      path: LOCATIONS,
      reply: page([
        {
          '@odata.type': '#microsoft.graph.ipNamedLocation',
          id: 'loc-hq',
          displayName: 'HQ',
          isTrusted: true,
          ipRanges: [{ cidrAddress: '203.0.113.0/24' }, { cidrAddress: '2001:db8::/32' }],
        },
        {
          '@odata.type': '#microsoft.graph.countryNamedLocation',
          id: 'loc-c',
          displayName: 'Blocked countries',
          countriesAndRegions: ['KP'],
        },
      ]),
    },
    { path: CA, reply: page([blockPolicy]) },
  ];

  test('lists ranges and the policies that use each location', async () => {
    const { ctx } = mockContext(routes);
    const out = (await cmd('ca locations').run(ctx, {})) as { items: Record<string, unknown>[] };
    expect(out.items[0]).toMatchObject({
      name: 'HQ',
      type: 'ip',
      trusted: true,
      usedBy: [{ name: 'CA010-AllUsers-DeviceCode-Block', state: 'report-only', use: 'exclude' }],
    });
    expect(out.items[1]).toMatchObject({ name: 'Blocked countries', type: 'country', countries: ['KP'] });
  });

  test('--ip keeps only locations whose ranges contain the address', async () => {
    const { ctx } = mockContext(routes);
    const hit = (await cmd('ca locations').run(ctx, { ip: '2001:db8:5::9' })) as { items: Record<string, unknown>[] };
    expect(hit.items).toEqual([expect.objectContaining({ name: 'HQ', matchedRange: '2001:db8::/32' })]);
    const miss = (await cmd('ca locations').run(ctx, { ip: '198.51.100.7' })) as { items: unknown[] };
    expect(miss.items).toEqual([]);
    await expect(cmd('ca locations').run(ctx, { ip: 'not-an-ip' })).rejects.toMatchObject({
      code: 'USAGE',
    } satisfies Partial<CliError>);
  });
});

describe('entra policies', () => {
  test('flags settings that differ from the recommendations; a forbidden section degrades alone', async () => {
    const { ctx } = mockContext([
      { path: '/policies/identitySecurityDefaultsEnforcementPolicy', reply: { body: { isEnabled: false } } },
      {
        path: '/policies/authorizationPolicy',
        reply: {
          body: {
            allowInvitesFrom: 'everyone',
            allowedToSignUpEmailBasedSubscriptions: true,
            guestUserRoleId: '2af84b1e-32c8-42b7-82bc-daa82404023b',
            defaultUserRolePermissions: {
              permissionGrantPoliciesAssigned: [
                'ManagePermissionGrantsForSelf.microsoft-user-default-low',
                'ManagePermissionGrantsForOwnedResource.microsoft-dynamically-managed-permissions-for-team',
              ],
            },
          },
        },
      },
      { path: '/policies/adminConsentRequestPolicy', reply: { body: { isEnabled: false, reviewers: [] } } },
      {
        path: '/policies/authenticationMethodsPolicy',
        reply: graphError(403, 'accessDenied', 'Request Authorization failed'),
      },
    ]);
    const out = (await cmd('policies').run(ctx, {})) as Record<string, unknown> & { warnings: string[] };
    expect(out.authorization).toMatchObject({
      guestInvitesFrom: 'everyone',
      guestAccess: 'restricted',
      userConsent: ['microsoft-user-default-low'],
    });
    expect(out.authenticationMethods).toMatchObject({ error: { code: 'FORBIDDEN' } });
    expect(out.warnings).toEqual([
      expect.stringContaining('Users can consent to apps themselves (microsoft-user-default-low)'),
      'Anyone, including guests, can invite guests.',
      'Users can sign up for email-verified (self-service) subscriptions.',
      'The admin consent workflow is off: users cannot ask an admin for an app.',
    ]);
  });

  test('SMS and voice sign-in are flagged', async () => {
    const { ctx } = mockContext([
      { path: '/policies/identitySecurityDefaultsEnforcementPolicy', reply: { body: { isEnabled: false } } },
      { path: '/policies/authorizationPolicy', reply: { body: {} } },
      { path: '/policies/adminConsentRequestPolicy', reply: { body: { isEnabled: true, reviewers: [{}] } } },
      {
        path: '/policies/authenticationMethodsPolicy',
        reply: {
          body: {
            authenticationMethodConfigurations: [
              { id: 'Sms', state: 'enabled' },
              { id: 'Voice', state: 'disabled' },
              { id: 'Fido2', state: 'enabled' },
            ],
            systemCredentialPreferences: { state: 'enabled' },
          },
        },
      },
    ]);
    const out = (await cmd('policies').run(ctx, {})) as Record<string, unknown> & { warnings: string[] };
    expect(out.authenticationMethods).toMatchObject({
      methods: { Sms: 'enabled', Voice: 'disabled', Fido2: 'enabled' },
      systemPreferredMfa: 'enabled',
    });
    expect(out.warnings).toEqual(['Sms sign-in is enabled; these methods are phishable.']);
  });
});
