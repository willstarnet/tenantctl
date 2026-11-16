import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands } from '../../src/domains/intune/index.ts';
import { commandById, graphError, type MockRequest, type MockRoute, page, runCommand } from '../helpers/graphMock.ts';

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
      result: { status: 'assigned', next: "tenantctl intune apps status 'Zoom Workplace'" },
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

  test('read-only source fields are not sent back, and policy-set assignments are left alone', async () => {
    const direct = [{ id: 'a-2', intent: 'required', target: groupTarget(KIOSK.id), source: 'direct', sourceId: null }];
    const argv = ['Zoom', '--intent', 'required', '--group', 'Sales Laptops'];
    const first = await runCommand(cmd('apps assign'), argv, appRoutes(direct), { allowWrites: true });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('apps assign'), argv, appRoutes(direct), { allowWrites: true, confirm: token });
    const sent = second.calls.find((c) => c.path.endsWith('/assign'))?.body as { mobileAppAssignments: object[] };
    expect(sent.mobileAppAssignments[0]).not.toHaveProperty('source');
    expect(sent.mobileAppAssignments[0]).not.toHaveProperty('sourceId');

    const fromSet = [{ ...direct[0], source: 'policySets', sourceId: 'set-1' }];
    const refused = runCommand(cmd('apps assign'), argv, appRoutes(fromSet), { allowWrites: true, dryRun: true });
    await expect(refused).rejects.toMatchObject({ code: 'USAGE', message: expect.stringContaining('policy set') });
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
    expect(batches.map((c) => c.version)).toEqual(['beta']);
    expect(batchRequests(batches[0]).map((r) => r.url)).toEqual([
      `${APPS}/app1/assignments`,
      `${APPS}/app2/assignments`,
      `${COMPLIANCE}/p1/assignments`,
      `${CONFIG}/c1/assignments`,
      `${CATALOG}/sc1/assignments`,
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

const FILTERS = '/deviceManagement/assignmentFilters';
const CORP = {
  id: 'f0000000-0000-0000-0000-000000000001',
  displayName: 'Corporate devices',
  platform: 'windows10AndLater',
  rule: '(device.deviceOwnership -eq "Corporate")',
  assignmentFilterManagementType: 'devices',
};
const CORP_LAPTOPS = { ...CORP, id: 'f0000000-0000-0000-0000-000000000002', displayName: 'Corporate laptops' };
const MAC = { ...CORP, id: 'f0000000-0000-0000-0000-000000000003', displayName: 'Mac fleet', platform: 'macOS' };
const filtersRoute: MockRoute = { path: FILTERS, reply: page([CORP, CORP_LAPTOPS, MAC]) };
const filtered = (target: Record<string, unknown>, id: string | null, type: string) => ({
  ...target,
  deviceAndAppManagementAssignmentFilterId: id,
  deviceAndAppManagementAssignmentFilterType: type,
});
const COMPLIANCE_TYPE = '#microsoft.graph.deviceCompliancePolicyAssignment';

function complianceRoutes(existing: unknown[], assign: MockRoute['reply'] = page([])): MockRoute[] {
  return [
    {
      path: COMPLIANCE,
      reply: page([
        { '@odata.type': '#microsoft.graph.windows10CompliancePolicy', id: 'p1', displayName: 'Windows baseline' },
      ]),
    },
    { path: `${COMPLIANCE}/p1/assignments`, reply: page(existing) },
    { method: 'POST', path: `${COMPLIANCE}/p1/assign`, reply: assign },
    filtersRoute,
    ...groupRoutes,
    namesBatch,
  ];
}

/** Preview, then approve with the confirm token; returns the preview and the `/assign` request. */
async function previewAndApply(id: string, argv: string[], routes: MockRoute[]) {
  const first = await runCommand(cmd(id), argv, routes, { allowWrites: true });
  if (first.outcome.kind !== 'confirm') throw new Error(`expected a confirm outcome, got ${first.outcome.kind}`);
  expect(first.calls.some((c) => c.method === 'POST' && c.path.endsWith('/assign'))).toBe(false);
  const second = await runCommand(cmd(id), argv, routes, { allowWrites: true, confirm: first.outcome.token });
  const post = second.calls.find((c) => c.method === 'POST' && c.path.endsWith('/assign'));
  return { preview: first.outcome.preview as Record<string, unknown>, post, second };
}

describe('assign with assignment filters', () => {
  test('--filter puts the filter on the new include target and keeps existing assignments', async () => {
    const existing = [{ id: 'c-1', target: filtered(allUsersTarget, MAC.id, 'exclude') }];
    const { preview, post } = await previewAndApply(
      'compliance assign',
      ['baseline', '--group', 'Sales Laptops', '--filter', 'Corporate devices'],
      complianceRoutes(existing),
    );
    expect(preview.action).toBe(
      'Apply compliance policy "Windows baseline" to members of Sales Laptops (42 members) with filter "Corporate devices" (include)',
    );
    expect(preview.filter).toEqual({
      change: 'set',
      id: CORP.id,
      name: 'Corporate devices',
      mode: 'include',
      platform: 'windows10AndLater',
      managementType: 'devices',
      rule: CORP.rule,
    });
    expect(preview.result).toEqual([
      { target: 'allUsers', filter: 'Mac fleet', filterMode: 'exclude', filterId: MAC.id },
      {
        target: 'group',
        group: 'Sales Laptops',
        groupId: SALES.id,
        filter: 'Corporate devices',
        filterMode: 'include',
        filterId: CORP.id,
      },
    ]);
    expect(post?.version).toBe('beta');
    expect(post?.body).toEqual({
      assignments: [
        { '@odata.type': COMPLIANCE_TYPE, ...existing[0] },
        { '@odata.type': COMPLIANCE_TYPE, target: filtered(groupTarget(SALES.id), CORP.id, 'include') },
      ],
    });
  });

  test('a group already assigned without a filter gets the filter in place', async () => {
    const existing = [{ id: 'c-1', target: groupTarget(SALES.id) }];
    const { preview, post, second } = await previewAndApply(
      'compliance assign',
      ['baseline', '--group', 'Sales Laptops', '--filter', CORP.id, '--filter-mode', 'exclude'],
      complianceRoutes(existing),
    );
    expect(preview).toMatchObject({
      changed: true,
      action:
        'Add filter "Corporate devices" (exclude) to the assignment of compliance policy "Windows baseline" to members of Sales Laptops (42 members)',
    });
    expect(post?.body).toEqual({
      assignments: [
        { '@odata.type': COMPLIANCE_TYPE, id: 'c-1', target: filtered(groupTarget(SALES.id), CORP.id, 'exclude') },
      ],
    });
    expect(second.outcome).toMatchObject({ kind: 'result', result: { status: 'assigned' } });
  });

  test('the same filter already in place is reported unchanged and nothing is sent', async () => {
    const existing = [{ id: 'c-1', target: filtered(groupTarget(SALES.id), CORP.id, 'include') }];
    const { outcome, calls } = await runCommand(
      cmd('compliance assign'),
      ['baseline', '--group', 'Sales Laptops', '--filter', 'Corporate devices'],
      complianceRoutes(existing),
      { allowWrites: true, dryRun: true },
    );
    expect(outcome).toMatchObject({ kind: 'result', result: { dryRun: true, preview: { changed: false } } });
    expect(calls.some((c) => c.method === 'POST' && c.path.endsWith('/assign'))).toBe(false);
  });

  test('--clear-filter removes the filter from the existing assignment', async () => {
    const existing = [
      { id: 'c-1', target: filtered(groupTarget(SALES.id), CORP.id, 'include') },
      { id: 'c-2', target: filtered(allUsersTarget, MAC.id, 'exclude') },
    ];
    const { preview, post } = await previewAndApply(
      'compliance assign',
      ['baseline', '--group', 'Sales Laptops', '--clear-filter'],
      complianceRoutes(existing),
    );
    expect(preview.action).toBe(
      'Remove filter "Corporate devices" (include) from the assignment of compliance policy "Windows baseline" to members of Sales Laptops (42 members)',
    );
    expect(preview.filter).toEqual({
      change: 'clear',
      id: CORP.id,
      name: 'Corporate devices',
      mode: 'include',
      rule: CORP.rule,
    });
    expect(post?.body).toEqual({
      assignments: [
        { '@odata.type': COMPLIANCE_TYPE, id: 'c-1', target: filtered(groupTarget(SALES.id), null, 'none') },
        { '@odata.type': COMPLIANCE_TYPE, ...existing[1] },
      ],
    });
  });

  test('a filter on an --exclude group alone is refused before anything is read', async () => {
    // No routes: any Graph call would fail with NOT_FOUND instead of USAGE.
    const run = runCommand(
      cmd('compliance assign'),
      ['baseline', '--exclude', 'Kiosks', '--filter', 'Corporate devices'],
      [],
      { allowWrites: true, dryRun: true },
    );
    await expect(run).rejects.toMatchObject({ code: 'USAGE' } satisfies Partial<CliError>);
  });

  test('--filter-mode without --filter and --filter with --clear-filter are usage errors', async () => {
    for (const extra of [
      ['--filter-mode', 'exclude'],
      ['--filter', 'Corporate devices', '--clear-filter'],
    ]) {
      const run = runCommand(cmd('config assign'), ['Wi-Fi', '--group', 'Sales Laptops', ...extra], [], {
        allowWrites: true,
        dryRun: true,
      });
      await expect(run).rejects.toMatchObject({ code: 'USAGE' } satisfies Partial<CliError>);
    }
  });

  test('a filter name matching several filters is AMBIGUOUS', async () => {
    const run = runCommand(
      cmd('compliance assign'),
      ['baseline', '--group', 'Sales Laptops', '--filter', 'Corporate'],
      complianceRoutes([]),
      { allowWrites: true, dryRun: true },
    );
    await expect(run).rejects.toMatchObject({ code: 'AMBIGUOUS' } satisfies Partial<CliError>);
  });

  test('a filter for another platform is refused with a hint naming the right platform', async () => {
    const run = runCommand(
      cmd('compliance assign'),
      ['baseline', '--group', 'Sales Laptops', '--filter', 'Mac fleet'],
      complianceRoutes([]),
      { allowWrites: true, dryRun: true },
    );
    await expect(run).rejects.toMatchObject({
      code: 'USAGE',
      hint: 'Pick a windows filter: `tenantctl intune filters list --platform windows`.',
    } satisfies Partial<CliError>);
  });

  test('Graph refusing the filter on /assign comes back with a filter hint', async () => {
    const routes = complianceRoutes([], graphError(400, 'BadRequest', 'Invalid filter for this payload'));
    const argv = ['baseline', '--group', 'Sales Laptops', '--filter', 'Corporate devices'];
    const first = await runCommand(cmd('compliance assign'), argv, routes, { allowWrites: true });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const run = runCommand(cmd('compliance assign'), argv, routes, { allowWrites: true, confirm: token });
    await expect(run).rejects.toMatchObject({ status: 400 } satisfies Partial<CliError>);
    await run.catch((error: CliError) => expect(error.hint).toContain('filter "Corporate devices"'));
  });
});

describe('assignment filter display', () => {
  test('apps assignments names the filter', async () => {
    const assignments = [
      { id: 'a-1', intent: 'required', target: filtered(groupTarget(KIOSK.id), CORP.id, 'include') },
    ];
    const { outcome } = await runCommand(cmd('apps assignments'), ['Zoom'], [...appRoutes(assignments), filtersRoute]);
    expect(outcome).toMatchObject({
      kind: 'result',
      result: {
        assignments: [
          {
            target: 'group',
            group: 'Kiosks',
            intent: 'required',
            filter: 'Corporate devices',
            filterMode: 'include',
            filterId: CORP.id,
          },
        ],
      },
    });
  });

  test('when filters cannot be read, only the filter id and mode are shown', async () => {
    const assignments = [
      { id: 'a-1', intent: 'required', target: filtered(groupTarget(KIOSK.id), CORP.id, 'include') },
    ];
    const { outcome } = await runCommand(
      cmd('apps assignments'),
      ['Zoom'],
      [...appRoutes(assignments), { path: FILTERS, reply: graphError(403, 'Forbidden', 'Missing scope') }],
    );
    const result = outcome.kind === 'result' ? (outcome.result as { assignments: Record<string, unknown>[] }) : null;
    expect(result?.assignments[0]).toEqual({
      target: 'group',
      group: 'Kiosks',
      groupId: KIOSK.id,
      intent: 'required',
      filterMode: 'include',
      filterId: CORP.id,
    });
  });

  test('intune assignments <group> hits carry the filter name', async () => {
    const { outcome } = await runCommand(
      cmd('assignments'),
      ['Sales Laptops'],
      [
        ...groupRoutes,
        { path: `/groups/${SALES.id}/transitiveMemberOf/microsoft.graph.group`, reply: page([]) },
        { path: APPS, reply: page([]) },
        { path: COMPLIANCE, reply: page([{ id: 'p1', displayName: 'Windows baseline' }]) },
        { path: CONFIG, reply: page([]) },
        { path: CATALOG, reply: page([]) },
        filtersRoute,
        batchRoute(() => ({ body: { value: [{ target: filtered(groupTarget(SALES.id), CORP.id, 'include') }] } })),
      ],
    );
    expect(outcome).toMatchObject({
      kind: 'result',
      result: {
        included: [
          {
            kind: 'compliance',
            name: 'Windows baseline',
            filter: 'Corporate devices',
            filterMode: 'include',
            filterId: CORP.id,
          },
        ],
      },
    });
  });
});
