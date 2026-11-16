import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands } from '../../src/domains/intune/index.ts';
import {
  commandById,
  graphError,
  type MockRequest,
  type MockRoute,
  mockContext,
  page,
  runCommand,
} from '../helpers/graphMock.ts';

const cmd = (id: string) => commandById(commands, `intune ${id}`);
const result = <T>(outcome: { kind: string; result?: unknown }) => (outcome as { result: T }).result;

const DM = '/deviceManagement';

/** $batch answering each request through `answer(path)`. */
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

const GROUP = { target: { '@odata.type': '#microsoft.graph.groupAssignmentTarget', groupId: 'g1' } };

describe('intune scripts list', () => {
  test('merges platform scripts and shell scripts, reporting a kind it cannot read', async () => {
    const routes: MockRoute[] = [
      {
        path: `${DM}/deviceManagementScripts`,
        reply: page([{ id: 'p1', displayName: 'Set time zone', runAsAccount: 'system', enforceSignatureCheck: true }]),
      },
      { path: `${DM}/deviceHealthScripts`, reply: graphError(403, 'Forbidden', 'Missing scope') },
      {
        path: `${DM}/deviceShellScripts`,
        reply: page([{ id: 's1', displayName: 'Install Rosetta', runAsAccount: 'user', executionFrequency: 'PT1H' }]),
      },
      batchRoute((url) => ({ body: { value: url.includes('/p1/') ? [GROUP] : [] } })),
    ];
    const { ctx } = mockContext(routes);
    const out = (await cmd('scripts list').run(ctx, { top: 50, all: false })) as {
      items: Record<string, unknown>[];
      unavailable?: unknown[];
    };
    expect(out.items).toEqual([
      expect.objectContaining({ kind: 'macosShell', name: 'Install Rosetta', runFrequency: 'PT1H', assigned: false }),
      expect.objectContaining({
        kind: 'platform',
        name: 'Set time zone',
        runAs: 'system',
        runFrequency: 'once',
        enforceSignatureCheck: true,
        assigned: true,
      }),
    ]);
    expect(out.unavailable).toEqual([
      { section: 'remediation', error: { code: 'FORBIDDEN', message: expect.stringContaining('Missing scope') } },
    ]);
  });

  test('a remediation shows its assignment schedule', async () => {
    const routes: MockRoute[] = [
      { path: `${DM}/deviceManagementScripts`, reply: page([]) },
      { path: `${DM}/deviceShellScripts`, reply: page([]) },
      { path: `${DM}/deviceHealthScripts`, reply: page([{ id: 'h1', displayName: 'Clear temp' }]) },
      batchRoute(() => ({
        body: {
          value: [
            {
              ...GROUP,
              runSchedule: {
                '@odata.type': '#microsoft.graph.deviceHealthScriptDailySchedule',
                interval: 1,
                time: '09:00:00',
              },
            },
          ],
        },
      })),
    ];
    const { outcome } = await runCommand(cmd('scripts list'), [], routes);
    expect(result<{ items: Record<string, unknown>[] }>(outcome).items[0]).toMatchObject({
      kind: 'remediation',
      runFrequency: 'daily, every 1, at 09:00:00',
    });
  });
});

describe('intune scripts status', () => {
  const states = [
    { id: 'd1', runState: 'success', resultMessage: 'ok', managedDevice: { id: 'm1', deviceName: 'SALES-01' } },
    {
      id: 'd2',
      runState: 'fail',
      errorCode: -1,
      errorDescription: 'x'.repeat(2000),
      lastStateUpdateDateTime: '2026-11-01T00:00:00Z',
      managedDevice: { id: 'm2', deviceName: 'SALES-02' },
    },
  ];
  const routes = (): MockRoute[] => [
    { path: `${DM}/deviceManagementScripts`, reply: page([{ id: 'p1', displayName: 'Set time zone' }]) },
    { path: `${DM}/deviceHealthScripts`, reply: page([]) },
    { path: `${DM}/deviceShellScripts`, reply: page([]) },
    {
      path: `${DM}/deviceManagementScripts/p1/runSummary`,
      reply: { body: { successDeviceCount: 1, errorDeviceCount: 1, successUserCount: 0, errorUserCount: 0 } },
    },
    { path: `${DM}/deviceManagementScripts/p1/deviceRunStates`, reply: page(states) },
  ];

  test('--failed-only keeps failing devices and truncates long errors', async () => {
    const { outcome, calls } = await runCommand(cmd('scripts status'), ['time zone', '--failed-only'], routes());
    const out = result<{ summary: unknown; devices: { deviceName: string; error: string }[] }>(outcome);
    expect(out.summary).toEqual({ success: 1, failed: 1, successUsers: 0, failedUsers: 0 });
    expect(out.devices.map((d) => d.deviceName)).toEqual(['SALES-02']);
    expect(out.devices[0]?.error.startsWith('x'.repeat(500))).toBe(true);
    expect(out.devices[0]?.error).toContain('(2000 chars)');
    expect(calls.find((c) => c.path.endsWith('/deviceRunStates'))?.query.$expand).toBe(
      'managedDevice($select=id,deviceName)',
    );
  });

  test('without --failed-only every device is listed', async () => {
    const { outcome } = await runCommand(cmd('scripts status'), ['p1'], routes());
    expect(result<{ devices: unknown[] }>(outcome).devices).toHaveLength(2);
  });
});

const RING = {
  '@odata.type': '#microsoft.graph.windowsUpdateForBusinessConfiguration',
  id: 'r1',
  displayName: 'Broad ring',
  qualityUpdatesDeferralPeriodInDays: 7,
  featureUpdatesDeferralPeriodInDays: 30,
  qualityUpdatesPaused: true,
  qualityUpdatesPauseExpiryDateTime: '2026-12-01T00:00:00Z',
  featureUpdatesPaused: false,
  featureUpdatesPauseExpiryDateTime: '0001-01-01T00:00:00Z',
  deadlineForQualityUpdatesInDays: 2,
  deadlineForFeatureUpdatesInDays: 7,
  deadlineGracePeriodInDays: 1,
};
const AUTOPATCH_RING = { ...RING, id: 'r2', displayName: 'Windows Autopatch Update Policy - Test' };
const FEATURE = {
  id: 'f1',
  displayName: 'Windows 11 rollout',
  featureUpdateVersion: 'Windows 11, version 24H2',
  rolloutSettings: { offerStartDateTimeInUTC: null, offerEndDateTimeInUTC: null, offerIntervalInDays: null },
};

function updateRoutes(): MockRoute[] {
  return [
    { path: `${DM}/deviceConfigurations`, reply: page([RING, AUTOPATCH_RING]) },
    { path: `${DM}/windowsFeatureUpdateProfiles`, reply: page([FEATURE]) },
    { path: `${DM}/windowsQualityUpdateProfiles`, reply: page([]) },
    { path: `${DM}/windowsDriverUpdateProfiles`, reply: graphError(403, 'Forbidden', 'nope') },
    batchRoute((url) =>
      url.startsWith('/groups/') ? { body: { id: 'g1', displayName: 'All laptops' } } : { body: { value: [GROUP] } },
    ),
  ];
}

describe('intune updates list', () => {
  test('filters rings by type, shapes deferrals and pauses, flags Autopatch and resolves group names', async () => {
    const { outcome, calls } = await runCommand(cmd('updates list'), [], updateRoutes());
    expect(calls.find((c) => c.path === `${DM}/deviceConfigurations`)?.query.$filter).toBe(
      "isof('microsoft.graph.windowsUpdateForBusinessConfiguration')",
    );
    const out = result<{ items: Record<string, unknown>[] }>(outcome);
    expect(out.items[0]).toMatchObject({
      kind: 'ring',
      name: 'Broad ring',
      managedByAutopatch: false,
      quality: { deferralDays: 7, deadlineDays: 2, paused: true, pauseExpiryDateTime: '2026-12-01T00:00:00Z' },
      feature: { deferralDays: 30, deadlineDays: 7, paused: false, pauseExpiryDateTime: null },
      deadlineGracePeriodDays: 1,
      assignments: [{ target: 'group', group: 'All laptops', groupId: 'g1' }],
    });
    expect(out.items[1]).toMatchObject({ name: AUTOPATCH_RING.displayName, managedByAutopatch: true });
    expect(out.items[2]).toMatchObject({
      kind: 'feature',
      targetVersion: FEATURE.featureUpdateVersion,
      rollout: 'immediate',
    });
    expect(out.items.map((i) => i.kind)).not.toContain('driver');
  });
});

describe('intune updates get', () => {
  test('a name matching a ring and a feature profile is AMBIGUOUS', async () => {
    const run = runCommand(cmd('updates get'), ['windows'], updateRoutes());
    await expect(run).rejects.toMatchObject({ code: 'AMBIGUOUS' } satisfies Partial<CliError>);
  });

  test('a unique name returns its settings and assignments', async () => {
    const { outcome } = await runCommand(cmd('updates get'), ['Windows 11 rollout'], updateRoutes());
    expect(result<{ policy: Record<string, unknown> }>(outcome).policy).toMatchObject({
      kind: 'feature',
      id: 'f1',
      managedByAutopatch: false,
      assignments: [{ group: 'All laptops' }],
    });
  });
});
