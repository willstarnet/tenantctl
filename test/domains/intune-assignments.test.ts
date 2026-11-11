import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands } from '../../src/domains/intune/index.ts';
import { commandById, type MockRequest, type MockRoute, page, runCommand } from '../helpers/graphMock.ts';

const APPS = '/deviceAppManagement/mobileApps';
const COMPLIANCE = '/deviceManagement/deviceCompliancePolicies';
const CONFIG = '/deviceManagement/deviceConfigurations';
const CATALOG = '/deviceManagement/configurationPolicies';

const SALES = { id: '11111111-1111-1111-1111-111111111111', displayName: 'Sales Laptops', mail: null };
const KIOSK = { id: '22222222-2222-2222-2222-222222222222', displayName: 'Kiosks', mail: null };
const PARENT = { id: '33333333-3333-3333-3333-333333333333', displayName: 'All Staff', mail: null };

const groupTarget = (groupId: string) => ({ '@odata.type': '#microsoft.graph.groupAssignmentTarget', groupId });
const exclusionTarget = (groupId: string) => ({
  '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget',
  groupId,
});
const allUsersTarget = { '@odata.type': '#microsoft.graph.allLicensedUsersAssignmentTarget' };

interface BatchCall {
  requests: { id: string; method: string; url: string }[];
}

function batchRequests(call: MockRequest | undefined): BatchCall['requests'] {
  const body = call?.body;
  if (body && typeof body === 'object' && 'requests' in body && Array.isArray(body.requests)) {
    return body.requests as BatchCall['requests'];
  }
  return [];
}

/** `$batch` route that answers each inner request by URL. */
function batchRoute(answer: (url: string) => { status?: number; body: unknown }): MockRoute {
  return {
    method: 'POST',
    path: '/$batch',
    reply: (req) => ({
      body: {
        responses: batchRequests(req).map((r) => {
          const { status, body } = answer(r.url);
          return { id: r.id, status: status ?? 200, body };
        }),
      },
    }),
  };
}

const groupRoutes: MockRoute[] = [
  {
    path: '/groups',
    reply: (req) => {
      const filter = req.query.$filter ?? '';
      return page([SALES, KIOSK].filter((g) => filter.includes(`'${g.displayName}'`)));
    },
  },
  { path: `/groups/${SALES.id}/members`, reply: { body: { '@odata.count': 42, value: [{ id: 'm1' }] } } },
  { path: `/groups/${KIOSK.id}/members`, reply: { body: { '@odata.count': 3, value: [{ id: 'm2' }] } } },
];

const namesBatch = batchRoute((url) => {
  const group = [SALES, KIOSK, PARENT].find((g) => url.startsWith(`/groups/${g.id}`));
  return group ? { body: group } : { status: 404, body: { error: { code: 'NotFound', message: 'gone' } } };
});

const zoom = { id: 'app1', displayName: 'Zoom Workplace' };
const existingAppAssignments = [
  { id: 'a-1', intent: 'available', target: allUsersTarget, settings: null },
  { id: 'a-2', intent: 'required', target: groupTarget(KIOSK.id), settings: null },
];

function appRoutes(assignments: unknown[] = existingAppAssignments): MockRoute[] {
  return [
    { path: APPS, reply: page([zoom]) },
    { path: `${APPS}/app1/assignments`, reply: page(assignments) },
    { method: 'POST', path: `${APPS}/app1/assign`, reply: { status: 204 } },
    ...groupRoutes,
    namesBatch,
  ];
}

const cmd = (id: string) => commandById(commands, `intune ${id}`);

describe('intune apps assign / unassign', () => {
  test('assign previews the member count and merged list, then sends existing assignments plus the new one', async () => {
    const argv = ['Zoom', '--intent', 'required', '--group', 'Sales Laptops'];
    const first = await runCommand(cmd('apps assign'), argv, appRoutes(), { allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.path.endsWith('/assign'))).toBe(false);
    const preview =
      first.outcome.kind === 'confirm'
        ? (first.outcome.preview as { action: string; current: unknown[]; result: Record<string, unknown>[] })
        : null;
    expect(preview?.action).toBe('Install Zoom Workplace (required) on members of Sales Laptops (42 members)');
    expect(preview?.current).toHaveLength(2);
    expect(preview?.result).toEqual([
      { target: 'allUsers', intent: 'available' },
      { target: 'group', group: 'Kiosks', groupId: KIOSK.id, intent: 'required' },
      { target: 'group', group: 'Sales Laptops', groupId: SALES.id, intent: 'required' },
    ]);

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('apps assign'), argv, appRoutes(), { allowWrites: true, confirm: token });
    const post = second.calls.find((c) => c.method === 'POST' && c.path === `${APPS}/app1/assign`);
    expect(post?.body).toEqual({
      mobileAppAssignments: [
        { '@odata.type': '#microsoft.graph.mobileAppAssignment', ...existingAppAssignments[0] },
        { '@odata.type': '#microsoft.graph.mobileAppAssignment', ...existingAppAssignments[1] },
        { '@odata.type': '#microsoft.graph.mobileAppAssignment', intent: 'required', target: groupTarget(SALES.id) },
      ],
    });
    expect(second.outcome).toMatchObject({
      kind: 'result',
      result: { status: 'assigned', next: 'tenantctl intune apps status "Zoom Workplace"' },
    });
  });

  test('available intent with --all-devices is refused before anything is read or sent', async () => {
    const run = runCommand(cmd('apps assign'), ['Zoom', '--intent', 'available', '--all-devices'], appRoutes(), {
      allowWrites: true,
      dryRun: true,
    });
    await expect(run).rejects.toMatchObject({ code: 'USAGE' } satisfies Partial<CliError>);
  });

  test('unassign removes only the target group and keeps every other assignment', async () => {
    const argv = ['Zoom', '--group', 'Kiosks'];
    const first = await runCommand(cmd('apps unassign'), argv, appRoutes(), { allowWrites: true });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const preview = first.outcome.kind === 'confirm' ? (first.outcome.preview as { action: string }) : null;
    expect(preview?.action).toBe(
      'Remove the required assignment of Zoom Workplace for members of Kiosks (3 members) (already installed copies stay installed)',
    );
    const second = await runCommand(cmd('apps unassign'), argv, appRoutes(), { allowWrites: true, confirm: token });
    const post = second.calls.find((c) => c.method === 'POST' && c.path === `${APPS}/app1/assign`);
    expect(post?.body).toEqual({
      mobileAppAssignments: [{ '@odata.type': '#microsoft.graph.mobileAppAssignment', ...existingAppAssignments[0] }],
    });
  });

  test('unassign of a group that is not assigned is NOT_FOUND', async () => {
    const run = runCommand(cmd('apps unassign'), ['Zoom', '--group', 'Sales Laptops'], appRoutes(), {
      allowWrites: true,
    });
    await expect(run).rejects.toMatchObject({ code: 'NOT_FOUND' } satisfies Partial<CliError>);
  });

  test('apps assignments names the target groups', async () => {
    const { outcome } = await runCommand(cmd('apps assignments'), ['Zoom'], appRoutes());
    expect(outcome).toMatchObject({
      kind: 'result',
      result: {
        app: { id: 'app1', name: 'Zoom Workplace' },
        assignments: [
          { target: 'allUsers', intent: 'available' },
          { target: 'group', group: 'Kiosks', intent: 'required' },
        ],
      },
    });
  });
});

describe('intune compliance / config assign', () => {
  test('compliance assign with --exclude appends an include and an exclusion to the existing list', async () => {
    const existing = [{ id: 'c-1', target: allUsersTarget }];
    const routes: MockRoute[] = [
      { path: COMPLIANCE, reply: page([{ id: 'p1', displayName: 'Windows baseline' }]) },
      { path: `${COMPLIANCE}/p1/assignments`, reply: page(existing) },
      { method: 'POST', path: `${COMPLIANCE}/p1/assign`, reply: page([]) },
      ...groupRoutes,
      namesBatch,
    ];
    const argv = ['baseline', '--group', 'Sales Laptops', '--exclude', 'Kiosks'];
    const first = await runCommand(cmd('compliance assign'), argv, routes, { allowWrites: true });
    const preview = first.outcome.kind === 'confirm' ? (first.outcome.preview as { action: string }) : null;
    expect(preview?.action).toBe(
      'Apply compliance policy "Windows baseline" to members of Sales Laptops (42 members), excluding members of Kiosks (3 members)',
    );
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('compliance assign'), argv, routes, { allowWrites: true, confirm: token });
    const post = second.calls.find((c) => c.method === 'POST' && c.path === `${COMPLIANCE}/p1/assign`);
    const type = '#microsoft.graph.deviceCompliancePolicyAssignment';
    expect(post?.body).toEqual({
      assignments: [
        { '@odata.type': type, ...existing[0] },
        { '@odata.type': type, target: groupTarget(SALES.id) },
        { '@odata.type': type, target: exclusionTarget(KIOSK.id) },
      ],
    });
  });

  test('config assign of a settings catalog policy uses the beta assign action', async () => {
    const routes: MockRoute[] = [
      { path: CONFIG, reply: page([{ id: 'c1', displayName: 'Corporate Wi-Fi' }]) },
      { path: CATALOG, reply: page([{ id: 'sc1', name: 'Edge settings' }]) },
      { path: `${CATALOG}/sc1/assignments`, reply: page([]) },
      { method: 'POST', path: `${CATALOG}/sc1/assign`, reply: page([]) },
      ...groupRoutes,
      namesBatch,
    ];
    const argv = ['Edge', '--group', 'Sales Laptops'];
    const first = await runCommand(cmd('config assign'), argv, routes, { allowWrites: true });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('config assign'), argv, routes, { allowWrites: true, confirm: token });
    const post = second.calls.find((c) => c.method === 'POST' && c.path === `${CATALOG}/sc1/assign`);
    expect(post?.version).toBe('beta');
    expect(post?.body).toEqual({
      assignments: [
        {
          '@odata.type': '#microsoft.graph.deviceManagementConfigurationPolicyAssignment',
          target: groupTarget(SALES.id),
        },
      ],
    });
  });
});

describe('intune assignments <group>', () => {
  test('scans apps and policies through $batch and reports direct, nested, excluded and all-users targeting', async () => {
    const assignmentsByUrl: Record<string, unknown[]> = {
      [`${APPS}/app1/assignments`]: [{ intent: 'required', target: groupTarget(SALES.id) }],
      [`${APPS}/app2/assignments`]: [{ intent: 'available', target: allUsersTarget }],
      [`${COMPLIANCE}/p1/assignments`]: [{ target: exclusionTarget(SALES.id) }],
      [`${CONFIG}/c1/assignments`]: [{ target: groupTarget(PARENT.id) }],
      [`${CATALOG}/sc1/assignments`]: [{ target: groupTarget(KIOSK.id) }],
    };
    const { outcome, calls } = await runCommand(
      cmd('assignments'),
      ['Sales Laptops'],
      [
        ...groupRoutes,
        { path: `/groups/${SALES.id}/transitiveMemberOf/microsoft.graph.group`, reply: page([PARENT]) },
        {
          path: APPS,
          reply: page([
            { id: 'app1', displayName: 'Zoom' },
            { id: 'app2', displayName: 'Company Portal' },
          ]),
        },
        { path: COMPLIANCE, reply: page([{ id: 'p1', displayName: 'Windows baseline' }]) },
        { path: CONFIG, reply: page([{ id: 'c1', displayName: 'Corporate Wi-Fi' }]) },
        { path: CATALOG, reply: page([{ id: 'sc1', name: 'Edge settings' }]) },
        batchRoute((url) => ({ body: { value: assignmentsByUrl[url] ?? [] } })),
      ],
    );
    const batches = calls.filter((c) => c.path === '/$batch');
    expect(batches.map((c) => c.version).sort()).toEqual(['beta', 'v1.0']);
    const v1 = batches.find((c) => c.version === 'v1.0');
    expect(batchRequests(v1).map((r) => r.url)).toEqual([
      `${APPS}/app1/assignments`,
      `${APPS}/app2/assignments`,
      `${COMPLIANCE}/p1/assignments`,
      `${CONFIG}/c1/assignments`,
    ]);
    expect(outcome).toMatchObject({
      kind: 'result',
      result: {
        group: { id: SALES.id, displayName: 'Sales Laptops', memberCount: 42 },
        included: [
          { kind: 'app', name: 'Zoom', intent: 'required' },
          { kind: 'config', name: 'Corporate Wi-Fi', via: 'All Staff' },
        ],
        excluded: [{ kind: 'compliance', name: 'Windows baseline' }],
        allUsers: [{ kind: 'app', name: 'Company Portal', intent: 'available' }],
        allDevices: [],
        truncated: false,
      },
    });
  });

  test('--max-items caps the scan and reports truncation', async () => {
    const { outcome } = await runCommand(
      cmd('assignments'),
      ['Sales Laptops', '--max-items', '1'],
      [
        ...groupRoutes,
        { path: `/groups/${SALES.id}/transitiveMemberOf/microsoft.graph.group`, reply: page([]) },
        { path: APPS, reply: page([{ id: 'app1', displayName: 'Zoom' }]) },
        { path: COMPLIANCE, reply: page([{ id: 'p1', displayName: 'Windows baseline' }]) },
        { path: CONFIG, reply: page([]) },
        { path: CATALOG, reply: page([]) },
        batchRoute(() => ({ body: { value: [] } })),
      ],
    );
    expect(outcome).toMatchObject({ kind: 'result', result: { truncated: true, scanned: { apps: 1 } } });
  });
});
