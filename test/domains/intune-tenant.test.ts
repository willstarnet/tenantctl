import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands } from '../../src/domains/intune/index.ts';
import { commandById, graphError, type MockRoute, mockContext, page, runCommand } from '../helpers/graphMock.ts';

const cmd = (id: string) => commandById(commands, `intune ${id}`);
const inDays = (days: number) => new Date(Date.now() + days * 86_400_000 + 3_600_000).toISOString();
const daysAgo = (days: number) => new Date(Date.now() - days * 86_400_000).toISOString();

const DEFENDER = 'fc780465-2017-40d4-a0c5-307022471b92';

function tenantRoutes(overrides: MockRoute[] = []): MockRoute[] {
  return [
    ...overrides,
    {
      path: '/deviceManagement/settings',
      reply: { body: { secureByDefault: false, deviceComplianceCheckinThresholdDays: 30 } },
    },
    {
      path: '/deviceManagement/applePushNotificationCertificate',
      reply: { body: { appleIdentifier: 'admin@contoso.com', expirationDateTime: inDays(9) } },
    },
    {
      path: '/deviceManagement/depOnboardingSettings',
      reply: page([
        {
          id: 'dep1',
          tokenName: 'Contoso ADE',
          appleIdentifier: 'it@contoso.com',
          tokenExpirationDateTime: inDays(200),
        },
      ]),
    },
    {
      path: '/deviceAppManagement/vppTokens',
      reply: page([
        {
          id: 'vpp1',
          organizationName: 'Contoso',
          appleId: 'apps@contoso.com',
          state: 'valid',
          expirationDateTime: inDays(-2),
        },
      ]),
    },
    {
      path: '/deviceManagement/androidManagedStoreAccountEnterpriseSettings',
      reply: { body: { bindStatus: 'boundAndValidated' } },
    },
    {
      path: '/deviceManagement/mobileThreatDefenseConnectors',
      reply: page([
        {
          id: DEFENDER,
          partnerState: 'enabled',
          androidEnabled: true,
          iosEnabled: true,
          windowsEnabled: true,
          macEnabled: false,
        },
      ]),
    },
  ];
}

describe('intune tenant', () => {
  test('counts days left on Apple tokens and turns what needs action into warnings', async () => {
    const { ctx } = mockContext(tenantRoutes());
    const out = (await cmd('tenant').run(ctx, { warnDays: 30 })) as {
      apns: { daysLeft: number; appleId: string };
      depTokens: { daysLeft: number }[];
      warnings: string[];
    };
    expect(out.apns).toMatchObject({ appleId: 'admin@contoso.com', daysLeft: 9 });
    expect(out.depTokens[0]?.daysLeft).toBe(200);
    expect(out.warnings).toEqual([
      expect.stringContaining('secureByDefault is off'),
      expect.stringMatching(/Apple MDM push certificate \(admin@contoso.com\) expires in 9 day\(s\).*SAME Apple ID/),
      expect.stringMatching(/Apple VPP token "Contoso" .* expired 2 day\(s\) ago/),
      expect.stringContaining('Defender for Endpoint connector does not feed device risk into compliance for macos'),
    ]);
  });

  test('a section Graph refuses becomes { error } and the rest still answers', async () => {
    const { ctx } = mockContext(
      tenantRoutes([{ path: '/deviceAppManagement/vppTokens', reply: graphError(403, 'Forbidden', 'Missing scope') }]),
    );
    const out = (await cmd('tenant').run(ctx, { warnDays: 30 })) as Record<string, unknown>;
    expect(out.vppTokens).toEqual({ error: { code: 'FORBIDDEN', message: expect.stringContaining('Missing scope') } });
    expect(out.apns).toMatchObject({ daysLeft: 9 });
  });
});

describe('intune cleanup-rules create', () => {
  const RULES = '/deviceManagement/managedDeviceCleanupRules';
  const devices = [
    { id: 'd1', deviceName: 'OLD-LAPTOP', operatingSystem: 'Windows', lastSyncDateTime: daysAgo(200) },
    { id: 'd2', deviceName: 'OLD-PHONE', operatingSystem: 'iOS', lastSyncDateTime: daysAgo(150) },
  ];
  const routes = (existing: unknown[] = []): MockRoute[] => [
    { path: RULES, reply: page(existing) },
    { method: 'POST', path: RULES, reply: { status: 201, body: { id: '00000000-0000-0000-0000-000000000000' } } },
    { path: '/deviceManagement/managedDevices', reply: page(devices) },
  ];

  test('previews the devices it would hide, waits for approval, then posts the rule', async () => {
    const argv = ['--days', '120', '--platform', 'windows'];
    const first = await runCommand(cmd('cleanup-rules create'), argv, routes(), { allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    const preview = (first.outcome as { preview: { action: string; affected: { count: number } } }).preview;
    expect(preview.affected.count).toBe(1);
    expect(preview.action).toContain('1 device record(s) visible now would be hidden');
    expect(first.calls.some((c) => c.method === 'POST')).toBe(false);

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('cleanup-rules create'), argv, routes(), { allowWrites: true, confirm: token });
    expect(second.calls.find((c) => c.method === 'POST')?.body).toEqual({
      '@odata.type': '#microsoft.graph.managedDeviceCleanupRule',
      displayName: 'Windows - 120 days inactive',
      description: '',
      deviceCleanupRulePlatformType: 'windows',
      deviceInactivityBeforeRetirementInDays: 120,
    });
  });

  test('refuses days outside 30-270 and a second rule for the same platform', async () => {
    const tooFew = runCommand(cmd('cleanup-rules create'), ['--days', '10'], routes(), {
      allowWrites: true,
      dryRun: true,
    });
    await expect(tooFew).rejects.toMatchObject({ code: 'USAGE' } satisfies Partial<CliError>);
    const existing = [
      {
        id: 'r1',
        displayName: 'All - 90',
        deviceCleanupRulePlatformType: 'all',
        deviceInactivityBeforeRetirementInDays: 90,
      },
    ];
    const twice = runCommand(cmd('cleanup-rules create'), ['--days', '90'], routes(existing), {
      allowWrites: true,
      dryRun: true,
    });
    await expect(twice).rejects.toMatchObject({ code: 'USAGE', message: expect.stringContaining('already exists') });
  });
});
