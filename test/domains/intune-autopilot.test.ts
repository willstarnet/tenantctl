import { describe, expect, test } from 'bun:test';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import type { CliError } from '../../src/core/errors.ts';
import { commands } from '../../src/domains/intune/index.ts';
import {
  commandById,
  type MockRequest,
  type MockRoute,
  page,
  runCommand,
  tempConfigDir,
} from '../helpers/graphMock.ts';

const AP = '/deviceManagement/windowsAutopilotDeviceIdentities';
const MANAGED = '/deviceManagement/managedDevices';
const PROFILES = '/deviceManagement/windowsAutopilotDeploymentProfiles';
const ZERO = '00000000-0000-0000-0000-000000000000';

const cmd = (id: string) => commandById(commands, `intune ${id}`);

const LAPTOP = {
  id: 'aaaaaaaa-0000-0000-0000-000000000001',
  serialNumber: '5CG1234ABC',
  manufacturer: 'Contoso',
  model: 'Book 14',
  groupTag: 'Sales',
  enrollmentState: 'enrolled',
  deploymentProfileAssignmentStatus: 'assignedInSync',
  userPrincipalName: 'jane@contoso.com',
  displayName: 'SALES-01',
  managedDeviceId: ZERO,
  azureAdDeviceId: 'eeeeeeee-0000-0000-0000-000000000001',
};
const SPARE = {
  id: 'aaaaaaaa-0000-0000-0000-000000000002',
  serialNumber: '5CG1234XYZ',
  model: 'Book 14',
  groupTag: '',
  enrollmentState: 'notContacted',
  userPrincipalName: '',
  managedDeviceId: ZERO,
};

/** Autopilot identities answering `contains(serialNumber,'x')` filters, plus GET by id. */
function identityRoutes(devices: Record<string, unknown>[] = [LAPTOP, SPARE]): MockRoute[] {
  return [
    {
      path: AP,
      reply: (req) => {
        const part = req.query.$filter?.match(/contains\(serialNumber,'(.*)'\)/)?.[1]?.toLowerCase();
        return page(part ? devices.filter((d) => String(d.serialNumber).toLowerCase().includes(part)) : devices);
      },
    },
    ...devices.map((d) => ({ path: `${AP}/${d.id}`, reply: { body: d } })),
  ];
}

/** No Intune record for any device. */
const noIntune: MockRoute = { path: MANAGED, reply: page([]) };

function batchRoute(answer: (url: string) => { status?: number; body: unknown }): MockRoute {
  return {
    method: 'POST',
    path: '/$batch',
    reply: (req: MockRequest) => {
      const requests = (req.body as { requests: { id: string; url: string }[] }).requests;
      return {
        body: {
          responses: requests.map((r) => {
            const { status, body } = answer(r.url);
            return { id: r.id, status: status ?? 200, body };
          }),
        },
      };
    },
  };
}

const result = <T>(outcome: { kind: string; result?: unknown }) => (outcome as { result: T }).result;

describe('intune autopilot get: resolution', () => {
  test('an exact serial wins over other partial matches, case-insensitively', async () => {
    const { outcome, calls } = await runCommand(cmd('autopilot get'), ['5cg1234abc'], [...identityRoutes(), noIntune]);
    expect(calls[0]?.query.$filter).toBe("contains(serialNumber,'5cg1234abc')");
    expect(result<{ device: { id: string; groupTag: string } }>(outcome).device).toMatchObject({
      id: LAPTOP.id,
      groupTag: 'Sales',
    });
  });

  test('an Intune device name resolves through its serial number', async () => {
    const routes: MockRoute[] = [
      ...identityRoutes(),
      {
        path: MANAGED,
        reply: (req) =>
          page(
            (req.query.$filter ?? '').includes("deviceName eq 'SALES-PC'")
              ? [{ id: 'md-1', deviceName: 'SALES-PC', serialNumber: '5CG1234XYZ', operatingSystem: 'Windows' }]
              : [],
          ),
      },
    ];
    const { outcome } = await runCommand(cmd('autopilot get'), ['SALES-PC'], routes);
    expect(result<{ device: { serialNumber: string } }>(outcome).device.serialNumber).toBe('5CG1234XYZ');
  });

  test('a partial serial matching several devices is AMBIGUOUS', async () => {
    const run = runCommand(cmd('autopilot get'), ['5CG1234'], identityRoutes());
    await expect(run).rejects.toMatchObject({ code: 'AMBIGUOUS' } satisfies Partial<CliError>);
  });

  test('shows the expanded deployment profiles and the linked Intune device', async () => {
    const linked = { ...LAPTOP, managedDeviceId: 'mmmmmmmm-0000-0000-0000-000000000001' };
    const profile = {
      '@odata.type': '#microsoft.graph.azureADWindowsAutopilotDeploymentProfile',
      id: 'p1',
      displayName: 'User driven',
    };
    const routes: MockRoute[] = [
      {
        path: `${AP}/${LAPTOP.id}`,
        reply: (req) => ({
          body: req.query.$expand
            ? { ...linked, deploymentProfile: profile, intendedDeploymentProfile: profile }
            : linked,
        }),
      },
      ...identityRoutes([linked]),
      {
        path: `${MANAGED}/${linked.managedDeviceId}`,
        reply: { body: { id: linked.managedDeviceId, deviceName: 'SALES-01', serialNumber: LAPTOP.serialNumber } },
      },
    ];
    const { outcome, calls } = await runCommand(cmd('autopilot get'), [LAPTOP.serialNumber], routes);
    expect(calls.find((c) => c.query.$expand)?.query.$expand).toBe('deploymentProfile,intendedDeploymentProfile');
    expect(result<Record<string, unknown>>(outcome)).toMatchObject({
      deploymentProfile: { id: 'p1', name: 'User driven', joinType: 'entraJoined' },
      intuneDevice: { id: linked.managedDeviceId, deviceName: 'SALES-01' },
    });
  });
});

describe('intune autopilot list', () => {
  test('--group-tag matches exactly and case-insensitively on the client', async () => {
    const devices = [LAPTOP, { ...SPARE, groupTag: 'Sales-East' }];
    const { outcome } = await runCommand(cmd('autopilot list'), ['--group-tag', 'sales'], identityRoutes(devices));
    expect(result<{ items: { id: string }[] }>(outcome).items.map((d) => d.id)).toEqual([LAPTOP.id]);
  });
});

describe('intune autopilot update', () => {
  const routes = (): MockRoute[] => [
    ...identityRoutes(),
    { method: 'POST', path: `${AP}/${LAPTOP.id}/updateDeviceProperties`, reply: { status: 204 } },
    { method: 'POST', path: `${AP}/${LAPTOP.id}/unassignUserFromDevice`, reply: { status: 204 } },
  ];

  test('dry run previews before/after values and sends nothing', async () => {
    const { outcome, calls } = await runCommand(
      cmd('autopilot update'),
      [LAPTOP.serialNumber, '--group-tag', 'Finance', '--clear-user'],
      routes(),
      { allowWrites: true, dryRun: true },
    );
    expect(calls.some((c) => c.method === 'POST')).toBe(false);
    const preview = result<{ preview: { action: string; changes: unknown[]; requests: { path: string }[] } }>(
      outcome,
    ).preview;
    expect(preview.changes).toEqual([
      { field: 'groupTag', before: 'Sales', after: 'Finance' },
      { field: 'assignedUser', before: 'jane@contoso.com', after: null },
    ]);
    expect(preview.action).toBe(
      'Change groupTag from Sales to Finance, assignedUser from jane@contoso.com to (none) on Autopilot device 5CG1234ABC',
    );
  });

  test('clears the group tag with an empty value and unassigns the user with the dedicated action', async () => {
    const { calls } = await runCommand(
      cmd('autopilot update'),
      [LAPTOP.serialNumber, '--clear-group-tag', '--clear-user'],
      routes(),
      { allowWrites: true },
    );
    const posts = calls.filter((c) => c.method === 'POST');
    expect(posts.map((c) => c.path)).toEqual([
      `${AP}/${LAPTOP.id}/unassignUserFromDevice`,
      `${AP}/${LAPTOP.id}/updateDeviceProperties`,
    ]);
    expect(posts[1]?.body).toEqual({ groupTag: '' });
  });

  test('sets the user with UPN and display name', async () => {
    const user = { id: 'u1', displayName: 'Sam Lee', userPrincipalName: 'sam@contoso.com', mail: null };
    const { calls } = await runCommand(
      cmd('autopilot update'),
      [LAPTOP.serialNumber, '--user', 'sam@contoso.com'],
      [...routes(), { path: '/users/sam@contoso.com', reply: { body: user } }],
      { allowWrites: true },
    );
    const post = calls.find((c) => c.path.endsWith('/updateDeviceProperties'));
    expect(post?.body).toEqual({ userPrincipalName: 'sam@contoso.com', addressableUserName: 'Sam Lee' });
  });

  test('refuses a run with nothing to change', async () => {
    const run = runCommand(cmd('autopilot update'), [LAPTOP.serialNumber], routes(), { allowWrites: true });
    await expect(run).rejects.toMatchObject({ code: 'USAGE' } satisfies Partial<CliError>);
  });
});

describe('intune autopilot delete', () => {
  const deleteRoutes = (extra: MockRoute[] = []): MockRoute[] => [
    ...extra,
    ...identityRoutes(),
    noIntune,
    { method: 'POST', path: `${AP}/${LAPTOP.id}/unassignUserFromDevice`, reply: { status: 204 } },
    { method: 'DELETE', path: `${AP}/${LAPTOP.id}`, reply: { status: 204 } },
  ];

  test('previews, waits for approval, then unassigns the user and deletes the identity', async () => {
    const first = await runCommand(cmd('autopilot delete'), [LAPTOP.serialNumber], deleteRoutes(), {
      allowWrites: true,
    });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.method !== 'GET')).toBe(false);
    const preview = first.outcome.kind === 'confirm' ? (first.outcome.preview as { action: string }) : null;
    expect(preview?.action).toBe(
      'Deregister 5CG1234ABC (Contoso Book 14, group tag Sales) from Windows Autopilot, unassigning ' +
        'jane@contoso.com first. It stops getting the Autopilot experience at its next reset; registering it ' +
        'again needs its hardware hash.',
    );
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';

    const denied = runCommand(cmd('autopilot delete'), [LAPTOP.serialNumber], deleteRoutes(), {
      allowWrites: true,
      confirm: token,
      approve: false,
    });
    await expect(denied).rejects.toMatchObject({ code: 'DENIED' } satisfies Partial<CliError>);

    const second = await runCommand(cmd('autopilot delete'), [LAPTOP.serialNumber], deleteRoutes(), {
      allowWrites: true,
      confirm: token,
    });
    expect(second.calls.filter((c) => c.method !== 'GET').map((c) => `${c.method} ${c.path}`)).toEqual([
      `POST ${AP}/${LAPTOP.id}/unassignUserFromDevice`,
      `DELETE ${AP}/${LAPTOP.id}`,
    ]);
  });

  test('refuses while an Intune device record with the same serial exists', async () => {
    const intune: MockRoute = {
      path: MANAGED,
      reply: page([{ id: 'md-9', deviceName: 'SALES-01', serialNumber: LAPTOP.serialNumber }]),
    };
    const run = runCommand(cmd('autopilot delete'), [LAPTOP.serialNumber], deleteRoutes([intune]), {
      allowWrites: true,
    });
    await expect(run).rejects.toMatchObject({
      code: 'USAGE',
      hint: 'Run `tenantctl intune devices delete md-9` first, then rerun this command.',
    });
  });
});

describe('intune autopilot import', () => {
  const HEADER = 'Device Serial Number,Windows Product ID,Hardware Hash,Group Tag,Assigned User';
  const csv = (lines: string[]) => {
    const path = join(tempConfigDir(), 'hashes.csv');
    writeFileSync(path, [HEADER, ...lines].join('\n'));
    return path;
  };
  const importRoute: MockRoute = {
    method: 'POST',
    path: '/deviceManagement/importedWindowsAutopilotDeviceIdentities/import',
    reply: (req) => ({
      body: {
        value: (
          req.body as { importedWindowsAutopilotDeviceIdentities: { serialNumber: string }[] }
        ).importedWindowsAutopilotDeviceIdentities.map((d, i) => ({
          id: `imp-${i}`,
          serialNumber: d.serialNumber,
          state: { deviceImportStatus: 'pending', deviceErrorCode: 0, deviceErrorName: '' },
        })),
      },
    }),
  };

  test('reports every bad row and sends nothing', async () => {
    const path = csv(['SN1,,QUJD,,', ',,QUJD,,', 'SN3,,not base64!,,', 'sn1,,QUJD,,', 'SN5,,QUJD,,jane']);
    const run = runCommand(cmd('autopilot import'), [path], [importRoute], { allowWrites: true });
    await expect(run).rejects.toMatchObject({
      code: 'USAGE',
      details: [
        { line: 3, serialNumber: null, problem: 'Device Serial Number is empty' },
        { line: 4, serialNumber: 'SN3', problem: 'Hardware Hash is not base64' },
        { line: 5, serialNumber: 'sn1', problem: 'serial sn1 also appears on line 2' },
        { line: 6, serialNumber: 'SN5', problem: 'Assigned User "jane" is not a UPN' },
      ],
    });
  });

  test('--group-tag overrides the CSV and the body uses the documented property names', async () => {
    const path = csv(['SN1,PK1,QUJD,Old,jane@contoso.com', 'SN2,,REVG,,']);
    const dry = await runCommand(cmd('autopilot import'), [path, '--group-tag', 'Sales'], [importRoute], {
      allowWrites: true,
      dryRun: true,
    });
    expect(result<{ preview: Record<string, unknown> }>(dry.outcome).preview).toMatchObject({
      action: 'Import 2 devices into Windows Autopilot',
      groupTags: { Sales: 2 },
    });
    expect(dry.calls).toHaveLength(0);

    const { outcome, calls } = await runCommand(
      cmd('autopilot import'),
      [path, '--group-tag', 'Sales'],
      [importRoute],
      {
        allowWrites: true,
      },
    );
    const body = calls[0]?.body as { importedWindowsAutopilotDeviceIdentities: Record<string, unknown>[] };
    expect(calls[0]?.version).toBe('v1.0');
    expect(body.importedWindowsAutopilotDeviceIdentities[0]).toMatchObject({
      '@odata.type': '#microsoft.graph.importedWindowsAutopilotDeviceIdentity',
      serialNumber: 'SN1',
      productKey: 'PK1',
      hardwareIdentifier: 'QUJD',
      groupTag: 'Sales',
      assignedUserPrincipalName: 'jane@contoso.com',
    });
    expect(result<Record<string, unknown>>(outcome)).toMatchObject({
      status: 'submitted',
      statuses: { pending: 2 },
      imported: [
        { id: 'imp-0', serialNumber: 'SN1', status: 'pending' },
        { id: 'imp-1', serialNumber: 'SN2', status: 'pending' },
      ],
    });
  });
});

describe('intune autopilot profiles assign / unassign', () => {
  const GROUP = { id: '11111111-1111-1111-1111-111111111111', displayName: 'Autopilot Devices', mail: null };
  const OTHER = { id: '22222222-2222-2222-2222-222222222222', displayName: 'Kiosks', mail: null };
  const profile = {
    '@odata.type': '#microsoft.graph.azureADWindowsAutopilotDeploymentProfile',
    id: 'p1',
    displayName: 'User driven',
    hardwareHashExtractionEnabled: true,
  };
  const assigned = (groupId: string, type = 'groupAssignmentTarget') => ({
    id: `p1_${groupId}`,
    target: { '@odata.type': `#microsoft.graph.${type}`, groupId },
  });
  const routes = (current: unknown[]): MockRoute[] => [
    { path: PROFILES, reply: page([profile]) },
    { path: `${PROFILES}/p1/assignments`, reply: page(current) },
    { method: 'POST', path: `${PROFILES}/p1/assignments`, reply: { status: 201, body: {} } },
    { method: 'DELETE', path: new RegExp(`^${PROFILES}/p1/assignments/`), reply: { status: 204 } },
    {
      path: '/groups',
      reply: (req) => page([GROUP, OTHER].filter((g) => (req.query.$filter ?? '').includes(`'${g.displayName}'`))),
    },
    { path: `/groups/${GROUP.id}/members`, reply: { body: { '@odata.count': 12, value: [] } } },
    { path: `/groups/${OTHER.id}/members`, reply: { body: { '@odata.count': 3, value: [] } } },
    batchRoute((url) => ({ body: [GROUP, OTHER].find((g) => url.startsWith(`/groups/${g.id}`)) ?? {} })),
  ];

  test('previews member count and the convert-all warning, then posts one assignment', async () => {
    const argv = ['User driven', '--group', 'Autopilot Devices'];
    const first = await runCommand(cmd('autopilot profiles assign'), argv, routes([assigned(OTHER.id)]), {
      allowWrites: true,
    });
    const preview = first.outcome.kind === 'confirm' ? (first.outcome.preview as Record<string, unknown>) : {};
    expect(preview.action).toBe('Assign Autopilot profile "User driven" to members of Autopilot Devices (12 members)');
    expect(String(preview.warning)).toContain('converts all targeted devices');
    expect(preview.result).toEqual([
      { target: 'group', group: 'Kiosks', groupId: OTHER.id },
      { target: 'group', group: 'Autopilot Devices', groupId: GROUP.id },
    ]);
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('autopilot profiles assign'), argv, routes([assigned(OTHER.id)]), {
      allowWrites: true,
      confirm: token,
    });
    const posts = second.calls.filter((c) => c.method === 'POST' && c.path === `${PROFILES}/p1/assignments`);
    expect(posts.map((c) => c.body)).toEqual([
      {
        '@odata.type': '#microsoft.graph.windowsAutopilotDeploymentProfileAssignment',
        target: { '@odata.type': '#microsoft.graph.groupAssignmentTarget', groupId: GROUP.id },
      },
    ]);
  });

  test('assigning an already-assigned group changes nothing', async () => {
    const argv = ['User driven', '--group', 'Autopilot Devices'];
    const first = await runCommand(cmd('autopilot profiles assign'), argv, routes([assigned(GROUP.id)]), {
      allowWrites: true,
    });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('autopilot profiles assign'), argv, routes([assigned(GROUP.id)]), {
      allowWrites: true,
      confirm: token,
    });
    expect(second.calls.some((c) => c.method === 'POST' && c.path.includes('/assignments'))).toBe(false);
    expect(result<Record<string, unknown>>(second.outcome)).toMatchObject({
      status: 'unchanged',
      action: 'Already in place: Assign Autopilot profile "User driven" to members of Autopilot Devices (12 members)',
    });
  });

  test('unassign --exclude deletes only the exclusion assignment', async () => {
    const current = [assigned(GROUP.id), assigned(OTHER.id, 'exclusionGroupAssignmentTarget')];
    const argv = ['User driven', '--exclude', 'Kiosks'];
    const first = await runCommand(cmd('autopilot profiles unassign'), argv, routes(current), { allowWrites: true });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('autopilot profiles unassign'), argv, routes(current), {
      allowWrites: true,
      confirm: token,
    });
    expect(second.calls.filter((c) => c.method === 'DELETE').map((c) => c.path)).toEqual([
      `${PROFILES}/p1/assignments/p1_${OTHER.id}`,
    ]);
  });
});

describe('intune autopilot events', () => {
  const day = 86_400_000;
  const iso = (daysBack: number) => new Date(Date.now() - daysBack * day).toISOString();
  const events = [
    {
      id: 'e1',
      eventDateTime: iso(1),
      deviceSerialNumber: '5CG1234ABC',
      managedDeviceName: 'SALES-01',
      userPrincipalName: 'jane@contoso.com',
      deploymentState: 'failure',
      deviceSetupStatus: 'failure',
      accountSetupStatus: 'notAttempted',
      enrollmentFailureDetails: 'App install timed out',
      deploymentTotalDuration: 'PT1H2M30S',
      deviceSetupDuration: 'PT58M',
      windowsAutopilotDeploymentProfileDisplayName: 'User driven',
    },
    { id: 'e2', eventDateTime: iso(2), deviceSerialNumber: 'SN2', deploymentState: 'success' },
    { id: 'e3', eventDateTime: iso(20), deviceSerialNumber: 'SN3', deploymentState: 'failure' },
  ];

  test('--failed-only --days keeps recent failures and shapes phases and durations', async () => {
    const route: MockRoute = { path: '/deviceManagement/autopilotEvents', reply: page(events) };
    const { outcome, calls } = await runCommand(cmd('autopilot events'), ['--failed-only', '--days', '7'], [route]);
    expect(calls[0]?.version).toBe('beta');
    expect(calls[0]?.query.$filter).toStartWith('eventDateTime ge ');
    expect(result<{ items: unknown[] }>(outcome).items).toEqual([
      {
        time: events[0]?.eventDateTime,
        device: 'SALES-01',
        serialNumber: '5CG1234ABC',
        user: 'jane@contoso.com',
        deploymentState: 'failure',
        enrollmentState: null,
        deviceSetupStatus: 'failure',
        accountSetupStatus: 'notAttempted',
        failedPhase: 'deviceSetup',
        failureDetails: 'App install timed out',
        totalMinutes: 62.5,
        deploymentMinutes: null,
        deviceSetupMinutes: 58,
        accountSetupMinutes: null,
        profile: 'User driven',
        enrollmentStatusPage: null,
        enrollmentType: null,
        osVersion: null,
        managedDeviceId: null,
      },
    ]);
  });

  test('a device argument matches the serial, newest first', async () => {
    const route: MockRoute = {
      path: '/deviceManagement/autopilotEvents',
      reply: page([events[2], { ...events[1], deviceSerialNumber: 'SN3' }]),
    };
    const { outcome } = await runCommand(cmd('autopilot events'), ['sn3'], [route]);
    expect(result<{ items: { time?: string }[] }>(outcome).items.map((e) => e.time)).toEqual([
      events[1]?.eventDateTime,
      events[2]?.eventDateTime,
    ]);
  });
});

describe('intune enrollment list', () => {
  const allUsers = { id: 'a0', target: { '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget' } };
  const configs = [
    {
      '@odata.type': '#microsoft.graph.windows10EnrollmentCompletionPageConfiguration',
      id: 'esp1',
      displayName: 'Corporate ESP',
      priority: 1,
      deviceEnrollmentConfigurationType: 'windows10EnrollmentCompletionPageConfiguration',
      showInstallationProgress: true,
      blockDeviceSetupRetryByUser: false,
      installProgressTimeoutInMinutes: 90,
      allowDeviceResetOnInstallFailure: true,
      allowDeviceUseOnInstallFailure: false,
      allowLogCollectionOnInstallFailure: true,
      customErrorMessage: 'Call the help desk',
      selectedMobileAppIds: ['app1'],
      trackInstallProgressForAutopilotOnly: true,
      assignments: [{ id: 'a1', target: { '@odata.type': '#microsoft.graph.groupAssignmentTarget', groupId: 'g1' } }],
    },
    {
      '@odata.type': '#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration',
      id: 'pr',
      displayName: 'All users and all devices',
      priority: 0,
      deviceEnrollmentConfigurationType: 'platformRestrictions',
      windowsRestriction: { platformBlocked: false, personalDeviceEnrollmentBlocked: true, osMinimumVersion: '10.0' },
      iosRestriction: { platformBlocked: true, personalDeviceEnrollmentBlocked: false },
      assignments: [allUsers],
    },
    {
      '@odata.type': '#microsoft.graph.deviceEnrollmentLimitConfiguration',
      id: 'lim',
      displayName: 'All users and all devices',
      priority: 0,
      deviceEnrollmentConfigurationType: 'limit',
      limit: 5,
      assignments: [allUsers],
    },
  ];
  const routes: MockRoute[] = [
    { path: '/deviceManagement/deviceEnrollmentConfigurations', reply: page(configs) },
    batchRoute((url) =>
      url.startsWith('/groups/g1')
        ? { body: { id: 'g1', displayName: 'Corporate Laptops' } }
        : { body: { id: 'app1', displayName: 'Company Portal' } },
    ),
  ];

  test('shapes ESP, platform restrictions and device limit settings with names', async () => {
    const { outcome, calls } = await runCommand(cmd('enrollment list'), [], routes);
    expect(calls[0]?.query.$expand).toBe('assignments');
    const items = result<{ items: Record<string, unknown>[] }>(outcome).items;
    expect(items.map((i) => i.type)).toEqual(['deviceLimit', 'enrollmentStatusPage', 'platformRestrictions']);
    expect(items[0]).toMatchObject({ settings: { limit: 5 }, assignments: [{ target: 'allUsers' }] });
    expect(items[1]).toMatchObject({
      name: 'Corporate ESP',
      settings: {
        showProgress: true,
        blockDeviceUseUntilInstalled: true,
        timeoutMinutes: 90,
        blockingApps: [{ id: 'app1', name: 'Company Portal' }],
        allowReset: true,
        allowUseOnFailure: false,
        customErrorMessage: 'Call the help desk',
      },
      assignments: [{ target: 'group', group: 'Corporate Laptops', groupId: 'g1' }],
    });
    expect(items[2]).toMatchObject({
      settings: {
        platforms: [
          { platform: 'windows', blocked: false, personalBlocked: true, osMinimumVersion: '10.0' },
          { platform: 'ios', blocked: true, personalBlocked: false },
        ],
      },
    });
  });

  test('--type keeps one kind of configuration', async () => {
    const { outcome } = await runCommand(cmd('enrollment list'), ['--type', 'deviceLimit'], routes);
    expect(result<{ items: { id: string }[] }>(outcome).items.map((i) => i.id)).toEqual(['lim']);
  });
});
