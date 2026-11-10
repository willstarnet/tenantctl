import { describe, expect, test } from 'bun:test';
import { CliError } from '../../src/core/errors.ts';
import { commands, resolveManagedDevice } from '../../src/domains/intune/index.ts';
import { commandById, graphError, mockContext, type MockRoute, page, runCommand } from '../helpers/graphMock.ts';

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
