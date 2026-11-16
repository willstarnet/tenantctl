import { describe, expect, test } from 'bun:test';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { CliError } from '../../src/core/errors.ts';
import type { Context } from '../../src/core/types.ts';
import { commands, resolveManagedDevice } from '../../src/domains/intune/index.ts';
import {
  commandById,
  graphError,
  type MockRequest,
  mockContext,
  type MockRoute,
  page,
  runCommand,
  tempConfigDir,
} from '../helpers/graphMock.ts';

const DEVICES = '/deviceManagement/managedDevices';

const laptop = {
  id: '00000000-0000-0000-0000-000000000001',
  deviceName: 'DESKTOP-4F2K9',
  userId: 'u-1',
  userPrincipalName: 'jane@contoso.com',
  operatingSystem: 'Windows',
  osVersion: '10.0.22631.3007',
  serialNumber: '5CG1234XYZ',
  complianceState: 'noncompliant',
  managementState: 'managed',
  managedDeviceOwnerType: 'company',
  lastSyncDateTime: new Date(Date.now() - 3 * 86_400_000).toISOString(),
  azureADDeviceId: 'aad-dev-1',
};

const phone = {
  ...laptop,
  id: '00000000-0000-0000-0000-000000000002',
  deviceName: 'Jane iPhone',
  operatingSystem: 'iOS',
  serialNumber: 'F2LXYZ',
};

/** Answer device-list queries by which property the $filter targets. */
function deviceList(byFilter: Record<string, unknown[]>): MockRoute {
  return {
    path: DEVICES,
    reply: (req) => {
      const filter = req.query.$filter ?? '';
      const key = Object.keys(byFilter).find((k) => filter.includes(k));
      return page(key ? (byFilter[key] as unknown[]) : []);
    },
  };
}

describe('resolveManagedDevice', () => {
  test('resolves a serial number to the device id', async () => {
    const { ctx, calls } = mockContext([deviceList({ serialNumber: [laptop] })]);
    const device = await resolveManagedDevice(ctx, '5CG1234XYZ');
    expect(device.id).toBe('00000000-0000-0000-0000-000000000001');
    expect(calls.map((c) => c.query.$filter)).toEqual([
      "(deviceName eq '5CG1234XYZ')",
      "(serialNumber eq '5CG1234XYZ')",
    ]);
  });

  test('two devices with the same name are AMBIGUOUS, listing serial, user and last sync', async () => {
    const twin = {
      ...laptop,
      id: '00000000-0000-0000-0000-000000000009',
      serialNumber: 'OLD999',
      userPrincipalName: 'bob@contoso.com',
    };
    const { ctx } = mockContext([deviceList({ deviceName: [laptop, twin] })]);
    const error = (await resolveManagedDevice(ctx, 'desktop-4f2k9').catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
    expect(error.details).toEqual([
      expect.objectContaining({
        id: '00000000-0000-0000-0000-000000000001',
        serialNumber: '5CG1234XYZ',
        user: 'jane@contoso.com',
      }),
      expect.objectContaining({
        id: '00000000-0000-0000-0000-000000000009',
        serialNumber: 'OLD999',
        user: 'bob@contoso.com',
      }),
    ]);
    expect((error.details as { lastSyncDateTime: string }[])[0]?.lastSyncDateTime).toBe(laptop.lastSyncDateTime);
  });

  test('an Entra device id falls back to an azureADDeviceId lookup; an admin center URL works too', async () => {
    const entraId = '11111111-2222-3333-4444-555555555555';
    const { ctx, calls } = mockContext([deviceList({ azureADDeviceId: [{ ...laptop, azureADDeviceId: entraId }] })]);
    const url = `https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId/${entraId}`;
    expect((await resolveManagedDevice(ctx, url)).id).toBe('00000000-0000-0000-0000-000000000001');
    expect(calls[0]?.path).toBe(`${DEVICES}/${entraId}`);
    expect(calls[1]?.query.$filter).toBe(`(azureADDeviceId eq '${entraId}')`);
  });

  test('a partial name falls back to a prefix search; nothing found is NOT_FOUND', async () => {
    const { ctx } = mockContext([deviceList({ startswith: [laptop] })]);
    expect((await resolveManagedDevice(ctx, 'DESKTOP-4F')).id).toBe('00000000-0000-0000-0000-000000000001');
    const empty = mockContext([deviceList({})]);
    const error = (await resolveManagedDevice(empty.ctx, 'nope').catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('NOT_FOUND');
  });
});

describe('devices list', () => {
  const list = commandById(commands, 'intune devices list');

  test('combines filters server-side and shapes devices with derived fields', async () => {
    const { outcome, calls } = await runCommand(
      list,
      ['--os', 'windows', '--compliance', 'noncompliant', '--stale-days', '2'],
      [{ path: DEVICES, reply: page([laptop]) }],
    );
    const filter = calls[0]?.query.$filter ?? '';
    expect(filter).toContain("(operatingSystem eq 'Windows')");
    expect(filter).toContain("(complianceState eq 'noncompliant')");
    expect(filter).toMatch(/\(lastSyncDateTime lt \d{4}-\d\d-\d\dT/);
    expect(calls[0]?.query.$select).toContain('azureADDeviceId');
    const result = outcome.kind === 'result' ? (outcome.result as { items: Record<string, unknown>[] }) : null;
    expect(result?.items[0]).toMatchObject({
      id: '00000000-0000-0000-0000-000000000001',
      user: 'jane@contoso.com',
      os: 'Windows',
      complianceState: 'noncompliant',
      ownership: 'company',
      daysSinceLastSync: 3,
    });
  });

  test('--user filters on the primary user UPN, and matches userId when Graph rejects that filter', async () => {
    const jane = {
      path: '/users/jane@contoso.com',
      reply: { body: { id: 'u-1', userPrincipalName: 'jane@contoso.com' } },
    };
    const served = await runCommand(
      list,
      ['--user', 'jane@contoso.com'],
      [jane, { path: DEVICES, reply: page([laptop]) }],
    );
    expect(served.calls[1]?.query.$filter).toBe("(userPrincipalName eq 'jane@contoso.com')");

    const renamed = { ...phone, userPrincipalName: 'jane.old@contoso.com' };
    const other = { ...laptop, id: '00000000-0000-0000-0000-000000000003', userId: 'u-2' };
    const { outcome, calls } = await runCommand(
      list,
      ['--user', 'jane@contoso.com'],
      [
        jane,
        {
          path: DEVICES,
          reply: (req) =>
            req.query.$filter
              ? graphError(400, 'BadRequest', 'Unsupported parameter found in query')
              : page([renamed, other]),
        },
      ],
    );
    expect(calls.at(-1)?.query.$filter).toBeUndefined();
    const result = outcome.kind === 'result' ? (outcome.result as { items: { id: string }[] }) : null;
    expect(result?.items.map((d) => d.id)).toEqual([renamed.id]);
  });

  test('--search filters client-side, because Intune returns nothing for startswith inside an or', async () => {
    const { outcome, calls } = await runCommand(
      list,
      ['--search', 'f2lxyz', '--compliance', 'noncompliant'],
      [{ path: DEVICES, reply: page([laptop, phone]) }],
    );
    expect(calls.map((c) => c.query.$filter)).toEqual(["(complianceState eq 'noncompliant')"]);
    const result = outcome.kind === 'result' ? (outcome.result as { items: { id: string }[] }) : null;
    expect(result?.items.map((d) => d.id)).toEqual([phone.id]);
  });

  test('when Graph rejects an undocumented filter, it is applied client-side', async () => {
    const { outcome, calls } = await runCommand(
      list,
      ['--os', 'ios', '--compliance', 'noncompliant'],
      [
        {
          path: DEVICES,
          reply: (req) =>
            req.query.$filter?.includes('operatingSystem')
              ? graphError(400, 'BadRequest', 'Invalid filter clause')
              : page([laptop, phone]),
        },
      ],
    );
    expect(calls[1]?.query.$filter).toBe("(complianceState eq 'noncompliant')");
    const result = outcome.kind === 'result' ? (outcome.result as { items: { id: string }[] }) : null;
    expect(result?.items.map((d) => d.id)).toEqual(['00000000-0000-0000-0000-000000000002']);
  });
});

describe('devices compliance', () => {
  test('summarises the failing policy settings in one sentence', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'intune devices compliance'),
      ['00000000-0000-0000-0000-000000000001'],
      [
        { path: `${DEVICES}/00000000-0000-0000-0000-000000000001`, reply: { body: laptop } },
        {
          path: `${DEVICES}/00000000-0000-0000-0000-000000000001/deviceCompliancePolicyStates`,
          reply: page([
            {
              id: 'p1',
              displayName: 'Windows baseline',
              platformType: 'windows10AndLater',
              state: 'nonCompliant',
              settingCount: 3,
              settingStates: [
                {
                  setting: 'Windows10CompliancePolicy.BitLockerEnabled',
                  settingName: 'BitLockerEnabled',
                  state: 'nonCompliant',
                  currentValue: 'False',
                  errorCode: 0,
                },
                {
                  setting: 'Windows10CompliancePolicy.SecureBootEnabled',
                  settingName: 'SecureBootEnabled',
                  state: 'compliant',
                  currentValue: 'True',
                  errorCode: 0,
                },
              ],
            },
            {
              id: 'p2',
              displayName: 'Default Device Compliance Policy',
              state: 'nonCompliant',
              settingStates: [
                {
                  setting: 'DefaultDeviceCompliancePolicy.RequireRemainContact',
                  settingName: 'Is active',
                  state: 'nonCompliant',
                  errorCode: 0,
                },
              ],
            },
            { id: 'p3', displayName: 'Defender', state: 'compliant', settingStates: [] },
          ]),
        },
      ],
    );
    expect(calls.at(-1)?.query.$top).toBeUndefined();
    const result =
      outcome.kind === 'result'
        ? (outcome.result as { summary: string; policies: { policy: string; failingSettings: unknown[] }[] })
        : null;
    expect(result?.summary).toBe(
      'DESKTOP-4F2K9 is noncompliant: "Windows baseline" fails BitLockerEnabled (current value False); ' +
        '"Default Device Compliance Policy" fails Is active (the device has not checked in with Intune within the ' +
        'compliance validity period).',
    );
    expect(result?.policies[0]?.failingSettings).toHaveLength(1);
    expect(result?.policies[2]?.failingSettings).toEqual([]);
  });

  test('a noncompliant device with no failing setting suggests a sync', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'intune devices compliance'),
      ['00000000-0000-0000-0000-000000000001'],
      [
        { path: `${DEVICES}/00000000-0000-0000-0000-000000000001`, reply: { body: laptop } },
        {
          path: `${DEVICES}/00000000-0000-0000-0000-000000000001/deviceCompliancePolicyStates`,
          reply: page([{ id: 'p3', displayName: 'Defender', state: 'compliant' }]),
        },
      ],
    );
    const result = outcome.kind === 'result' ? (outcome.result as { summary: string }) : null;
    expect(result?.summary).toContain('Sync it and re-check');
  });
});

describe('remote actions', () => {
  const wipe = commandById(commands, 'intune devices wipe');
  const routes: MockRoute[] = [
    { path: `${DEVICES}/00000000-0000-0000-0000-000000000001`, reply: { body: laptop } },
    { method: 'POST', path: `${DEVICES}/00000000-0000-0000-0000-000000000001/wipe`, reply: { status: 204 } },
  ];

  test('wipe previews a loud factory-reset warning, then the confirmed run sends the wipe body', async () => {
    const first = await runCommand(wipe, ['00000000-0000-0000-0000-000000000001', '--keep-enrollment-data'], routes, {
      allowWrites: true,
    });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.method === 'POST')).toBe(false);
    const preview =
      first.outcome.kind === 'confirm'
        ? (first.outcome.preview as { action: string; device: Record<string, unknown> })
        : null;
    expect(preview?.action).toContain('FACTORY RESET DESKTOP-4F2K9');
    expect(preview?.action).toContain('serial 5CG1234XYZ');
    expect(preview?.action).toContain('user jane@contoso.com');
    expect(preview?.device).toMatchObject({ serialNumber: '5CG1234XYZ', daysSinceLastSync: 3 });

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(wipe, ['00000000-0000-0000-0000-000000000001', '--keep-enrollment-data'], routes, {
      allowWrites: true,
      confirm: token,
    });
    const post = second.calls.find((c) => c.method === 'POST');
    expect(post?.path).toBe(`${DEVICES}/00000000-0000-0000-0000-000000000001/wipe`);
    expect(post?.version).toBe('v1.0');
    expect(post?.body).toEqual({ keepEnrollmentData: true, keepUserData: false });
    expect(second.outcome).toMatchObject({ kind: 'result', result: { status: 'requested' } });
  });

  test('actions unsupported on the device OS are refused before anything is sent', async () => {
    const lock = await runCommand(
      commandById(commands, 'intune devices lock'),
      ['00000000-0000-0000-0000-000000000001'],
      routes,
      { allowWrites: true },
    ).catch((e: unknown) => e);
    expect((lock as CliError).code).toBe('USAGE');
    expect((lock as CliError).message).toContain('not supported on Windows');

    const iosRoutes: MockRoute[] = [
      { path: `${DEVICES}/00000000-0000-0000-0000-000000000002`, reply: { body: phone } },
    ];
    const scan = await runCommand(
      commandById(commands, 'intune devices defender-scan'),
      ['00000000-0000-0000-0000-000000000002'],
      iosRoutes,
      { allowWrites: true },
    ).catch((e: unknown) => e);
    expect((scan as CliError).code).toBe('USAGE');
    const wipeIos = await runCommand(wipe, ['00000000-0000-0000-0000-000000000002', '--keep-user-data'], iosRoutes, {
      allowWrites: true,
    }).catch((e: unknown) => e);
    expect((wipeIos as CliError).message).toContain('Windows devices only');
  });

  test('beta-only actions call the beta endpoint with the documented body', async () => {
    const { calls } = await runCommand(
      commandById(commands, 'intune devices rename'),
      ['00000000-0000-0000-0000-000000000001', '--name', 'SALES-LT-042'],
      [
        { path: `${DEVICES}/00000000-0000-0000-0000-000000000001`, reply: { body: laptop } },
        {
          method: 'POST',
          path: `${DEVICES}/00000000-0000-0000-0000-000000000001/setDeviceName`,
          reply: { status: 204 },
        },
      ],
      { allowWrites: true },
    );
    const post = calls.find((c) => c.method === 'POST');
    expect(post?.version).toBe('beta');
    expect(post?.body).toEqual({ deviceName: 'SALES-LT-042' });
  });

  test('defender-scan defaults to a quick scan', async () => {
    const { calls } = await runCommand(
      commandById(commands, 'intune devices defender-scan'),
      ['00000000-0000-0000-0000-000000000001'],
      [
        { path: `${DEVICES}/00000000-0000-0000-0000-000000000001`, reply: { body: laptop } },
        {
          method: 'POST',
          path: `${DEVICES}/00000000-0000-0000-0000-000000000001/windowsDefenderScan`,
          reply: { status: 204 },
        },
      ],
      { allowWrites: true },
    );
    expect(calls.find((c) => c.method === 'POST')?.body).toEqual({ quickScan: true });
  });
});

describe('secrets', () => {
  test('bitlocker-key filters by the Entra device id and fetches keys only with --reveal', async () => {
    const routes: MockRoute[] = [
      { path: `${DEVICES}/00000000-0000-0000-0000-000000000001`, reply: { body: laptop } },
      {
        path: '/informationProtection/bitlocker/recoveryKeys',
        reply: page([{ id: 'key-1', createdDateTime: '2024-01-01T00:00:00Z', volumeType: 'operatingSystemVolume' }]),
      },
      {
        path: '/informationProtection/bitlocker/recoveryKeys/key-1',
        reply: { body: { id: 'key-1', key: '111111-222222' } },
      },
    ];
    const cmd = commandById(commands, 'intune devices bitlocker-key');
    const hidden = await runCommand(cmd, ['00000000-0000-0000-0000-000000000001'], routes);
    expect(hidden.calls[1]?.query).toEqual({ $filter: "deviceId eq 'aad-dev-1'" });
    expect(hidden.calls).toHaveLength(2);
    const shown = await runCommand(cmd, ['00000000-0000-0000-0000-000000000001', '--reveal'], routes);
    expect(shown.calls[2]?.query.$select).toBe('key');
    const result = shown.outcome.kind === 'result' ? (shown.outcome.result as { keys: { key: string }[] }) : null;
    expect(result?.keys[0]?.key).toBe('111111-222222');
  });

  test('laps --reveal decodes the newest password', async () => {
    const utf16 = Buffer.from('N3w-Pa$$', 'utf16le').toString('base64');
    const utf8 = Buffer.from('0ld-pass', 'utf8').toString('base64');
    const { outcome, calls } = await runCommand(
      commandById(commands, 'intune devices laps'),
      ['00000000-0000-0000-0000-000000000001', '--reveal'],
      [
        { path: `${DEVICES}/00000000-0000-0000-0000-000000000001`, reply: { body: laptop } },
        {
          path: '/directory/deviceLocalCredentials/aad-dev-1',
          reply: {
            body: {
              id: 'aad-dev-1',
              lastBackupDateTime: '2024-05-02T00:00:00Z',
              credentials: [
                { accountName: 'Administrator', backupDateTime: '2024-04-01T00:00:00Z', passwordBase64: utf8 },
                { accountName: 'Administrator', backupDateTime: '2024-05-02T00:00:00Z', passwordBase64: utf16 },
              ],
            },
          },
        },
      ],
    );
    expect(calls[1]?.query.$select).toContain('credentials');
    const result =
      outcome.kind === 'result' ? (outcome.result as { credentials: { password: string; current: boolean }[] }) : null;
    expect(result?.credentials.map((c) => [c.password, c.current])).toEqual([
      ['N3w-Pa$$', true],
      ['0ld-pass', false],
    ]);
  });
});

describe('policies and reports', () => {
  test('config list merges classic profiles (v1.0) and settings catalog policies (beta) with a kind', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'intune config list'),
      [],
      [
        {
          path: '/deviceManagement/deviceConfigurations',
          reply: page([
            {
              '@odata.type': '#microsoft.graph.windows10GeneralConfiguration',
              id: 'c1',
              displayName: 'Restrictions',
              assignments: [{ id: 'a' }],
            },
          ]),
        },
        {
          path: '/deviceManagement/configurationPolicies',
          reply: page([{ id: 'sc1', name: 'Edge settings', platforms: 'windows10', isAssigned: false }]),
        },
      ],
    );
    expect(calls.find((c) => c.path === '/deviceManagement/configurationPolicies')?.version).toBe('beta');
    const result = outcome.kind === 'result' ? (outcome.result as { items: Record<string, unknown>[] }) : null;
    expect(result?.items).toEqual([
      expect.objectContaining({ id: 'c1', kind: 'profile', type: 'windows10GeneralConfiguration', assigned: true }),
      expect.objectContaining({ id: 'sc1', kind: 'settingsCatalog', assigned: false }),
    ]);
  });

  test('compliance status --failing-only keeps failing devices only', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'intune compliance status'),
      ['baseline', '--failing-only'],
      [
        {
          path: '/deviceManagement/deviceCompliancePolicies',
          reply: page([{ id: 'p1', displayName: 'Windows baseline' }]),
        },
        {
          path: '/deviceManagement/deviceCompliancePolicies/p1/deviceStatuses',
          reply: page([
            {
              id: 's1',
              deviceDisplayName: 'DESKTOP-4F2K9',
              userPrincipalName: 'jane@contoso.com',
              status: 'nonCompliant',
            },
            { id: 's2', deviceDisplayName: 'DESKTOP-OK', status: 'compliant' },
          ]),
        },
      ],
    );
    const result = outcome.kind === 'result' ? (outcome.result as { items: { device: string }[] }) : null;
    expect(result?.items.map((i) => i.device)).toEqual(['DESKTOP-4F2K9']);
  });

  test('apps status reports failed devices with hex error codes', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'intune apps status'),
      ['Zoom'],
      [
        { path: '/deviceAppManagement/mobileApps', reply: page([{ id: 'app1', displayName: 'Zoom Workplace' }]) },
        {
          path: '/deviceAppManagement/mobileApps/app1/installSummary',
          reply: { body: { installedDeviceCount: 10, failedDeviceCount: 1 } },
        },
        {
          path: '/deviceAppManagement/mobileApps/app1/deviceStatuses',
          reply: page([
            { deviceName: 'DESKTOP-4F2K9', installState: 'failed', errorCode: -2016281116 },
            { deviceName: 'DESKTOP-OK', installState: 'installed', errorCode: 0 },
          ]),
        },
      ],
    );
    const result =
      outcome.kind === 'result'
        ? (outcome.result as { summary: Record<string, number>; failures: { items: Record<string, unknown>[] } })
        : null;
    expect(result?.summary).toMatchObject({ installedDevices: 10, failedDevices: 1 });
    expect(result?.failures.items).toEqual([
      expect.objectContaining({ device: 'DESKTOP-4F2K9', errorCode: '0x87D1FDE4' }),
    ]);
  });

  test('report noncompliant batches policy states, groups devices by OS and ranks failing settings', async () => {
    const failing = {
      value: [
        {
          id: 'p1',
          displayName: 'Windows baseline',
          state: 'nonCompliant',
          settingStates: [{ settingName: 'BitLockerEnabled', state: 'nonCompliant', currentValue: 'False' }],
        },
      ],
    };
    const { outcome, calls } = await runCommand(
      commandById(commands, 'intune report noncompliant'),
      [],
      [
        {
          path: DEVICES,
          reply: page([
            laptop,
            { ...laptop, id: '00000000-0000-0000-0000-000000000003' },
            { ...phone, complianceState: 'inGracePeriod' },
          ]),
        },
        {
          method: 'POST',
          path: '/$batch',
          reply: (req) => ({
            body: {
              responses: (req.body as { requests: { id: string; url: string }[] }).requests.map((r) =>
                r.url.includes('000000000002')
                  ? { id: r.id, status: 403, body: { error: { code: 'Forbidden', message: 'no' } } }
                  : { id: r.id, status: 200, body: failing },
              ),
            },
          }),
        },
      ],
    );
    const batch = calls.filter((c) => c.path === '/$batch');
    expect(batch).toHaveLength(1);
    expect((batch[0]?.body as { requests: { url: string }[] }).requests.map((r) => r.url)).toEqual([
      `${DEVICES}/00000000-0000-0000-0000-000000000001/deviceCompliancePolicyStates`,
      `${DEVICES}/00000000-0000-0000-0000-000000000003/deviceCompliancePolicyStates`,
      `${DEVICES}/00000000-0000-0000-0000-000000000002/deviceCompliancePolicyStates`,
    ]);
    expect(calls[0]?.query.$filter).toBe("(complianceState eq 'noncompliant' or complianceState eq 'inGracePeriod')");
    const result = outcome.kind === 'result' ? (outcome.result as Record<string, unknown>) : {};
    expect(result.byOs).toEqual({ Windows: 2, iOS: 1 });
    expect(result.byState).toEqual({ noncompliant: 2, inGracePeriod: 1 });
    expect(result.byReason).toEqual([{ policy: 'Windows baseline', setting: 'BitLockerEnabled', devices: 2 }]);
    const devices = result.devices as { id: string; error?: { code: string } }[];
    expect(devices.find((d) => d.id === '00000000-0000-0000-0000-000000000002')?.error?.code).toBe('FORBIDDEN');
  });
});

describe('devices config-states', () => {
  const REPORT = '/deviceManagement/reports/getConfigurationPoliciesReportForDevice';
  const STATES = `${DEVICES}/${laptop.id}/deviceConfigurationStates`;
  const SCHEMA = [
    'PolicyId',
    'PolicyName',
    'PolicyBaseTypeName',
    'PolicyStatus',
    'UPN',
    'UserId',
    'PspdpuLastModifiedTimeUtc',
    'UnifiedPolicyType',
  ];
  const PROFILE = 'Microsoft.Management.Services.Api.DeviceConfiguration';
  const CATALOG = 'DeviceManagementConfigurationPolicy';
  const rows = [
    ['c1', 'Corporate Wi-Fi', PROFILE, 5, 'jane@contoso.com', 'u-1', '2026-11-01T10:00:00Z', 'Wifi'],
    ['c1', 'Corporate Wi-Fi', PROFILE, 2, '', '', '2026-11-02T10:00:00Z', 'Wifi'],
    ['s1', 'Edge settings', CATALOG, 6, '', '', '2026-11-03T10:00:00Z', 'SettingsCatalog'],
    ['e1', 'Defender AV', CATALOG, 2, '', '', '2026-11-03T10:00:00Z', 'MicrosoftDefenderAntivirus'],
    ['a1', 'Office ADMX', 'DeviceConfigurationAdmxPolicy', 1, '', '', null, null],
  ];

  /** Report pages of `size` rows, answered as a text body the way Intune streams reports. */
  function reportRoute(all: unknown[][], size: number): MockRoute {
    return {
      method: 'POST',
      path: REPORT,
      reply: (req) => {
        const { skip } = req.body as { skip: number };
        const values = all.slice(skip, skip + size);
        return {
          body: JSON.stringify({
            TotalRowCount: all.length,
            Schema: SCHEMA.map((Column) => ({ Column })),
            Values: values,
          }),
        };
      },
    };
  }

  const deviceRoute: MockRoute = { path: `${DEVICES}/${laptop.id}`, reply: { body: laptop } };
  const run = (ctx: Context) =>
    commandById(commands, 'intune devices config-states').run(ctx, { device: laptop.id }) as Promise<{
      counts: Record<string, number>;
      problems: Record<string, unknown>[];
      profiles: Record<string, unknown>[];
    }>;

  test('pages the portal report for the device and groups user contexts per policy', async () => {
    const { ctx, calls } = mockContext([
      deviceRoute,
      reportRoute(rows, 2),
      {
        path: STATES,
        reply: page([
          {
            id: 'c1',
            displayName: 'Corporate Wi-Fi',
            state: 'error',
            settingStates: [
              { settingName: 'SSID', state: 'error', errorDescription: 'Certificate not found' },
              { settingName: 'AutoConnect', state: 'compliant' },
            ],
          },
        ]),
      },
    ]);
    const result = await run(ctx);
    const reports = calls.filter((c) => c.path === REPORT);
    expect(reports.map((c) => (c.body as { skip: number }).skip)).toEqual([0, 2, 4]);
    expect(reports[0]?.version).toBe('beta');
    expect(reports[0]?.body).toMatchObject({
      filter:
        "((PolicyBaseTypeName eq 'Microsoft.Management.Services.Api.DeviceConfiguration') or " +
        "(PolicyBaseTypeName eq 'DeviceManagementConfigurationPolicy') or " +
        "(PolicyBaseTypeName eq 'DeviceConfigurationAdmxPolicy') or " +
        "(PolicyBaseTypeName eq 'Microsoft.Management.Services.Api.DeviceManagementIntent')) and " +
        `(IntuneDeviceId eq '${laptop.id}')`,
      top: 100,
      orderBy: ['PolicyName'],
    });
    expect(result.counts).toEqual({ error: 1, conflict: 1, succeeded: 1, notApplicable: 1 });
    expect(result.problems).toEqual([
      expect.objectContaining({
        policy: 'Corporate Wi-Fi',
        kind: 'profile',
        state: 'error',
        statusCode: 5,
        lastReported: '2026-11-02T10:00:00Z',
        users: [
          expect.objectContaining({ user: 'jane@contoso.com', state: 'error', statusCode: 5 }),
          expect.objectContaining({ user: null, state: 'succeeded', statusCode: 2 }),
        ],
        failingSettings: [expect.objectContaining({ setting: 'SSID', errorDescription: 'Certificate not found' })],
      }),
      expect.objectContaining({ policy: 'Edge settings', kind: 'settingsCatalog', state: 'conflict', statusCode: 6 }),
    ]);
    expect(result.problems[1]?.failingSettings).toBeUndefined();
    expect(result.profiles).toEqual([
      expect.objectContaining({ policy: 'Defender AV', kind: 'endpointSecurity', state: 'succeeded' }),
      expect.objectContaining({ policy: 'Office ADMX', kind: 'admx', state: 'notApplicable', statusCode: 1 }),
    ]);
  });

  test('only classic profiles in error or conflict trigger the per-setting call', async () => {
    const { ctx, calls } = mockContext([deviceRoute, reportRoute(rows.slice(2), 100)]);
    const result = await run(ctx);
    expect(result.problems.map((p) => p.policy)).toEqual(['Edge settings']);
    expect(calls.some((c) => c.path === STATES)).toBe(false);
  });

  test('a throttled per-setting call degrades to { error } on the classic problem', async () => {
    const { ctx } = mockContext([
      deviceRoute,
      reportRoute(rows.slice(0, 1), 100),
      { path: STATES, reply: graphError(429, 'TooManyRequests', 'Too many requests') },
    ]);
    const result = await run(ctx);
    expect(result.problems[0]?.failingSettings).toEqual({
      error: expect.objectContaining({ code: 'THROTTLED' }),
    });
  });

  test('an unreadable report is a GRAPH_ERROR, not a crash', async () => {
    const { ctx } = mockContext([deviceRoute, { method: 'POST', path: REPORT, reply: { body: '<html>oops' } }]);
    const error = (await run(ctx).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('GRAPH_ERROR');
  });
});

describe('config get / export / create', () => {
  const CONFIG = '/deviceManagement/deviceConfigurations';
  const CATALOG = '/deviceManagement/configurationPolicies';
  const GROUP = '11111111-1111-1111-1111-111111111111';
  const catalogPolicy = {
    '@odata.context':
      'https://graph.microsoft.com/beta/$metadata#deviceManagement/configurationPolicies(settings())/$entity',
    id: 'sc1',
    name: 'Screen lock',
    description: 'Lock after 15 minutes',
    platforms: 'macOS',
    technologies: 'mdm,appleRemoteManagement',
    roleScopeTagIds: ['0'],
    createdDateTime: '2026-11-01T00:00:00Z',
    lastModifiedDateTime: '2026-11-02T00:00:00Z',
    settingCount: 2,
    isAssigned: true,
    templateReference: { templateId: '', templateFamily: 'none', templateDisplayName: null },
    'settings@odata.context': 'https://graph.microsoft.com/beta/$metadata#settings',
    settings: [
      {
        id: '0',
        settingInstance: {
          '@odata.type': '#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance',
          settingDefinitionId: 'passcode_passcode',
          settingInstanceTemplateReference: null,
          groupSettingCollectionValue: [
            {
              settingValueTemplateReference: null,
              children: [
                {
                  '@odata.type': '#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance',
                  settingDefinitionId: 'passcode_maxinactivity',
                  simpleSettingValue: {
                    '@odata.type': '#microsoft.graph.deviceManagementConfigurationIntegerSettingValue',
                    value: 15,
                  },
                },
              ],
            },
          ],
        },
      },
      {
        id: '1',
        settingInstance: {
          '@odata.type': '#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance',
          settingDefinitionId: 'device_vendor_msft_lock',
          choiceSettingValue: {
            value: 'device_vendor_msft_lock_1',
            children: [
              {
                '@odata.type': '#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance',
                settingDefinitionId: 'device_vendor_msft_lock_message',
                simpleSettingCollectionValue: [{ value: 'Locked' }, { value: 'Ask IT' }],
              },
            ],
          },
        },
      },
    ],
  };
  const listRoutes: MockRoute[] = [
    { path: CONFIG, reply: page([{ id: 'c1', displayName: 'Corporate Wi-Fi' }]) },
    { path: CATALOG, reply: page([{ id: 'sc1', name: 'Screen lock' }]) },
  ];

  test('get flattens settings catalog values and names assignment groups', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'intune config get'),
      ['screen lock'],
      [
        ...listRoutes,
        { path: `${CATALOG}/sc1`, reply: { body: catalogPolicy } },
        {
          path: `${CATALOG}/sc1/assignments`,
          reply: page([
            { id: 'a1', target: { '@odata.type': '#microsoft.graph.groupAssignmentTarget', groupId: GROUP } },
          ]),
        },
        {
          method: 'POST',
          path: '/$batch',
          reply: { body: { responses: [{ id: '0', status: 200, body: { displayName: 'Mac Fleet' } }] } },
        },
      ],
    );
    expect(calls.find((c) => c.path === `${CATALOG}/sc1`)?.query.$expand).toBe('settings');
    const result = outcome.kind === 'result' ? (outcome.result as Record<string, unknown>) : {};
    expect(result).toMatchObject({
      id: 'sc1',
      name: 'Screen lock',
      kind: 'settingsCatalog',
      platforms: 'macOS',
      settingCount: 2,
      assignments: [{ target: 'group', group: 'Mac Fleet', groupId: GROUP }],
    });
    expect(result.templateReference).toBeUndefined();
    expect(result.settings).toEqual([
      { id: 'passcode_passcode', value: [[{ id: 'passcode_maxinactivity', value: 15 }]] },
      {
        id: 'device_vendor_msft_lock',
        value: 'device_vendor_msft_lock_1',
        choice: '1',
        children: [{ id: 'device_vendor_msft_lock_message', value: ['Locked', 'Ask IT'] }],
      },
    ]);
  });

  test('get shows a classic profile with only its configured settings', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'intune config get'),
      ['wi-fi'],
      [
        ...listRoutes,
        {
          path: `${CONFIG}/c1`,
          reply: {
            body: {
              '@odata.type': '#microsoft.graph.windowsWifiConfiguration',
              id: 'c1',
              displayName: 'Corporate Wi-Fi',
              version: 3,
              ssid: 'CONTOSO',
              connectAutomatically: true,
              proxySetting: 'notConfigured',
              preSharedKey: null,
              roleScopeTagIds: ['0'],
            },
          },
        },
        { path: `${CONFIG}/c1/assignments`, reply: page([]) },
      ],
    );
    const result = outcome.kind === 'result' ? (outcome.result as Record<string, unknown>) : {};
    expect(result).toMatchObject({ kind: 'profile', type: 'windowsWifiConfiguration', name: 'Corporate Wi-Fi' });
    expect(result.settings).toEqual({ ssid: 'CONTOSO', connectAutomatically: true });
  });

  test('export strips ids, read-only properties and annotations but keeps setting @odata.type', async () => {
    const out = join(tempConfigDir(), 'policy.json');
    const routes = [...listRoutes, { path: `${CATALOG}/sc1`, reply: { body: catalogPolicy } }];
    const { outcome } = await runCommand(commandById(commands, 'intune config export'), ['sc1'], routes);
    const body = (outcome.kind === 'result' ? outcome.result : {}) as Record<string, unknown>;
    expect(Object.keys(body).sort()).toEqual(
      ['description', 'name', 'platforms', 'roleScopeTagIds', 'settings', 'technologies'].sort(),
    );
    const text = JSON.stringify(body);
    expect(text).not.toContain('"id"');
    expect(text).not.toContain('@odata.context');
    expect(text).not.toContain('null');
    expect(body.settings).toEqual([
      {
        settingInstance: expect.objectContaining({
          '@odata.type': '#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance',
          settingDefinitionId: 'passcode_passcode',
        }),
      },
      expect.anything(),
    ]);

    const written = await runCommand(commandById(commands, 'intune config export'), ['sc1', '--out', out], routes);
    expect(written.outcome.kind === 'result' && written.outcome.result).toMatchObject({ file: out, settingCount: 2 });
    expect(JSON.parse(readFileSync(out, 'utf8'))).toEqual(body);

    const again = runCommand(commandById(commands, 'intune config export'), ['sc1', '--out', out], routes);
    await expect(again).rejects.toMatchObject({ code: 'USAGE', message: expect.stringContaining('already exists') });
    const forced = await runCommand(
      commandById(commands, 'intune config export'),
      ['sc1', '--out', out, '--force'],
      routes,
    );
    expect(forced.outcome.kind).toBe('result');
    const intoConfig = runCommand(
      commandById(commands, 'intune config export'),
      ['sc1', '--out', join(process.env.TENANTCTL_CONFIG_DIR ?? '', 'audit.jsonl'), '--force'],
      routes,
    );
    await expect(intoConfig).rejects.toMatchObject({
      code: 'USAGE',
      message: expect.stringContaining('config folder'),
    });
  });

  test('export refuses a classic profile', async () => {
    const error = (await runCommand(commandById(commands, 'intune config export'), ['c1'], listRoutes).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(error.code).toBe('USAGE');
  });

  describe('create', () => {
    const create = commandById(commands, 'intune config create');
    const file = (content: unknown) => {
      const path = join(tempConfigDir(), 'policy.json');
      writeFileSync(path, typeof content === 'string' ? content : JSON.stringify(content));
      return path;
    };
    const valid = {
      name: 'Screen lock',
      platforms: 'macOS',
      technologies: 'mdm',
      settings: [{ settingInstance: catalogPolicy.settings[1]?.settingInstance }],
    };
    const post: MockRoute = {
      method: 'POST',
      path: CATALOG,
      reply: { status: 201, body: { id: 'new1', name: 'Screen lock v2' } },
    };

    test.each([
      ['invalid JSON', '{ nope', 'not valid JSON'],
      ['an array', [valid], 'one JSON object'],
      ['no name', { ...valid, name: undefined }, 'missing name'],
      ['no platforms or technologies', { name: 'x', settings: valid.settings }, 'missing platforms, technologies'],
      ['empty settings', { ...valid, settings: [] }, 'no settings'],
      ['a setting without a definition id', { ...valid, settings: [{ settingInstance: {} }] }, 'settings[0]'],
    ])('refuses %s before calling Graph', async (_label, content, message) => {
      let posted = false;
      const route: MockRoute = {
        ...post,
        reply: () => {
          posted = true;
          return { status: 201, body: {} };
        },
      };
      const error = (await runCommand(create, ['--file', file(content)], [route], { allowWrites: true }).catch(
        (e: unknown) => e,
      )) as CliError;
      expect(error.code).toBe('USAGE');
      expect(error.message).toContain(message);
      expect(posted).toBe(false);
    });

    test('--dry-run previews name, platform and setting ids', async () => {
      const { outcome, calls } = await runCommand(create, ['--file', file(valid), '--name', 'Screen lock v2'], [post], {
        allowWrites: true,
        dryRun: true,
      });
      const preview = outcome.kind === 'result' ? (outcome.result as { preview: Record<string, unknown> }).preview : {};
      expect(preview).toMatchObject({
        name: 'Screen lock v2',
        platforms: 'macOS',
        technologies: 'mdm',
        settingCount: 1,
        settingIds: ['device_vendor_msft_lock'],
      });
      expect(String(preview.action)).toContain('"Screen lock v2"');
      expect(calls).toEqual([]);
    });

    test('posts the cleaned body to beta with --name applied and returns an assign hint', async () => {
      const { outcome, calls } = await runCommand(
        create,
        ['--file', file({ ...catalogPolicy, name: 'Old' }), '--name', 'Screen lock v2'],
        [post],
        { allowWrites: true },
      );
      const call = calls.find((c: MockRequest) => c.method === 'POST' && c.path === CATALOG);
      expect(call?.version).toBe('beta');
      const body = call?.body as Record<string, unknown>;
      expect(body).toMatchObject({
        name: 'Screen lock v2',
        platforms: 'macOS',
        technologies: 'mdm,appleRemoteManagement',
      });
      expect(body.id).toBeUndefined();
      expect(body.isAssigned).toBeUndefined();
      expect(body.settingCount).toBeUndefined();
      expect((body.settings as unknown[]).length).toBe(2);
      expect(outcome.kind === 'result' && outcome.result).toEqual({
        id: 'new1',
        name: 'Screen lock v2',
        settingCount: 2,
        next: 'tenantctl intune config assign new1 --group <group>',
      });
    });
  });
});

describe('compliance settings', () => {
  test('counts results per setting, most failing first', async () => {
    const COMPLIANCE = '/deviceManagement/deviceCompliancePolicies';
    const { outcome, calls } = await runCommand(
      commandById(commands, 'intune compliance settings'),
      ['baseline'],
      [
        { path: COMPLIANCE, reply: page([{ id: 'p1', displayName: 'Windows baseline' }]) },
        {
          path: `${COMPLIANCE}/p1/deviceSettingStateSummaries`,
          reply: page([
            { settingName: 'Windows10CompliancePolicy.SecureBootEnabled', compliantDeviceCount: 10 },
            {
              settingName: 'Windows10CompliancePolicy.BitLockerEnabled',
              compliantDeviceCount: 6,
              nonCompliantDeviceCount: 3,
              errorDeviceCount: 1,
            },
            {
              settingName: 'Windows10CompliancePolicy.OsMinimumVersion',
              nonCompliantDeviceCount: 1,
              conflictDeviceCount: 1,
            },
          ]),
        },
      ],
    );
    expect(calls.at(-1)?.version).toBe('v1.0');
    const result = outcome.kind === 'result' ? (outcome.result as { settings: Record<string, unknown>[] }) : null;
    expect(result?.settings.map((s) => [s.setting, s.failing])).toEqual([
      ['BitLockerEnabled', 4],
      ['OsMinimumVersion', 2],
      ['SecureBootEnabled', 0],
    ]);
    expect(result?.settings[0]).toMatchObject({ compliant: 6, nonCompliant: 3, error: 1, conflict: 0 });
  });
});
