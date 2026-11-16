import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands } from '../../src/domains/intune/index.ts';
import { commandById, type MockRequest, type MockRoute, page, runCommand } from '../helpers/graphMock.ts';

const FILTERS = '/deviceManagement/assignmentFilters';
const APPS = '/deviceAppManagement/mobileApps';
const COMPLIANCE = '/deviceManagement/deviceCompliancePolicies';
const CONFIG = '/deviceManagement/deviceConfigurations';
const CATALOG = '/deviceManagement/configurationPolicies';

const CORP = {
  id: 'f0000000-0000-0000-0000-000000000001',
  displayName: 'Corporate devices',
  description: 'Corporate-owned Windows devices.',
  platform: 'windows10AndLater',
  rule: '(device.deviceOwnership -eq "Corporate")',
  assignmentFilterManagementType: 'devices',
  lastModifiedDateTime: '2026-01-05T10:00:00Z',
};
const CORP_LAPTOPS = { ...CORP, id: 'f0000000-0000-0000-0000-000000000002', displayName: 'Corporate laptops' };
const IPADS = {
  ...CORP,
  id: 'f0000000-0000-0000-0000-000000000003',
  displayName: 'Shared iPads',
  description: '',
  platform: 'iOS',
  rule: '(device.model -startsWith "iPad")',
};
const SALES = { id: '11111111-1111-1111-1111-111111111111', displayName: 'Sales Laptops' };

const cmd = (id: string) => commandById(commands, `intune ${id}`);
const filtersRoute: MockRoute = { path: FILTERS, reply: page([CORP, CORP_LAPTOPS, IPADS]) };

describe('intune filters list', () => {
  test('lists filters from beta with platform, management type and rule', async () => {
    const { outcome, calls } = await runCommand(cmd('filters list'), [], [filtersRoute]);
    const call = calls.find((c) => c.path === FILTERS);
    expect(call?.version).toBe('beta');
    expect(call?.query.$select).toContain('rule');
    expect(outcome).toMatchObject({
      kind: 'result',
      result: {
        items: [
          {
            id: CORP.id,
            name: 'Corporate devices',
            platform: 'windows10AndLater',
            managementType: 'devices',
            rule: CORP.rule,
            description: 'Corporate-owned Windows devices.',
            lastModifiedDateTime: '2026-01-05T10:00:00Z',
          },
          { name: 'Corporate laptops' },
          { name: 'Shared iPads', platform: 'iOS' },
        ],
        truncated: false,
      },
    });
  });

  test('--platform keeps only filters of that platform family', async () => {
    const { outcome } = await runCommand(cmd('filters list'), ['--platform', 'ios'], [filtersRoute]);
    expect(outcome).toMatchObject({ kind: 'result', result: { items: [{ name: 'Shared iPads' }], truncated: false } });
  });
});

interface BatchRequest {
  id: string;
  url: string;
}

function batchRequests(body: unknown): BatchRequest[] {
  return body && typeof body === 'object' && 'requests' in body && Array.isArray(body.requests) ? body.requests : [];
}

function batchRoute(answer: (url: string) => unknown): MockRoute {
  return {
    method: 'POST',
    path: '/$batch',
    reply: (req: MockRequest) => ({
      body: {
        responses: batchRequests(req.body).map((r) => ({
          id: r.id,
          status: 200,
          body: r.url.startsWith('/groups/') ? SALES : { value: answer(r.url) },
        })),
      },
    }),
  };
}

describe('intune filters get', () => {
  const groupTarget = (filterId: string | null, type: string) => ({
    '@odata.type': '#microsoft.graph.groupAssignmentTarget',
    groupId: SALES.id,
    deviceAndAppManagementAssignmentFilterId: filterId,
    deviceAndAppManagementAssignmentFilterType: type,
  });
  const scanRoutes: MockRoute[] = [
    filtersRoute,
    { path: APPS, reply: page([{ id: 'app1', displayName: 'Zoom' }]) },
    { path: COMPLIANCE, reply: page([{ id: 'p1', displayName: 'Windows baseline' }]) },
    { path: CONFIG, reply: page([{ id: 'c1', displayName: 'Corporate Wi-Fi' }]) },
    { path: CATALOG, reply: page([{ id: 'sc1', name: 'Edge settings' }]) },
    batchRoute((url) => {
      if (url === `${APPS}/app1/assignments`) return [{ intent: 'required', target: groupTarget(CORP.id, 'include') }];
      if (url === `${COMPLIANCE}/p1/assignments`) return [{ target: groupTarget(CORP.id, 'exclude') }];
      if (url === `${CONFIG}/c1/assignments`) return [{ target: groupTarget(CORP_LAPTOPS.id, 'include') }];
      return [{ target: groupTarget(null, 'none') }];
    }),
  ];

  test('shows the rule and every assignment using the filter, found through a beta $batch scan', async () => {
    const { outcome, calls } = await runCommand(cmd('filters get'), ['Corporate devices'], scanRoutes);
    const scan = calls.find((c) => c.path === '/$batch' && c.version === 'beta');
    expect(scan).toBeDefined();
    expect(outcome).toMatchObject({
      kind: 'result',
      result: {
        filter: { id: CORP.id, name: 'Corporate devices', platform: 'windows10AndLater', rule: CORP.rule },
        summary: 'Corporate devices is used by 2 assignment(s) on 2 item(s).',
        usedBy: [
          {
            kind: 'app',
            id: 'app1',
            name: 'Zoom',
            target: 'group',
            group: 'Sales Laptops',
            groupId: SALES.id,
            intent: 'required',
            filterMode: 'include',
          },
          { kind: 'compliance', id: 'p1', name: 'Windows baseline', target: 'group', filterMode: 'exclude' },
        ],
        scanned: { apps: 1, compliancePolicies: 1, configurationProfiles: 1, settingsCatalogPolicies: 1 },
        truncated: false,
      },
    });
  });

  test('a name matching several filters is AMBIGUOUS', async () => {
    const run = runCommand(cmd('filters get'), ['Corporate'], scanRoutes);
    await expect(run).rejects.toMatchObject({ code: 'AMBIGUOUS' } satisfies Partial<CliError>);
  });

  test('an unknown filter is NOT_FOUND', async () => {
    const run = runCommand(cmd('filters get'), ['Kiosk'], scanRoutes);
    await expect(run).rejects.toMatchObject({ code: 'NOT_FOUND' } satisfies Partial<CliError>);
  });
});
