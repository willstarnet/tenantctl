import { describe, expect, test } from 'bun:test';
import { commands } from '../../src/domains/intune/index.ts';
import { type SyncComponent, waitForSync } from '../../src/domains/intune/syncStatus.ts';
import { commandById, mockContext, type MockRoute, runCommand } from '../helpers/graphMock.ts';

const DEVICES = '/deviceManagement/managedDevices';
const ID = '00000000-0000-0000-0000-000000000001';
const STATUS = `${DEVICES}/${ID}/getSyncStatus()`;

const laptop = {
  id: ID,
  deviceName: 'DESKTOP-4F2K9',
  userPrincipalName: 'jane@contoso.com',
  operatingSystem: 'Windows',
  osVersion: '10.0.26200.9974',
  serialNumber: '5CG1234XYZ',
  lastSyncDateTime: new Date().toISOString(),
};

const at = (offsetMs: number) => new Date(Date.now() + offsetMs).toISOString();
const stage = (name: string, status: string, offsetMs = 0, moreInfo: string | null = null): SyncComponent => ({
  name,
  status,
  reportedDateTime: at(offsetMs),
  moreInfo,
});

/** getSyncStatus answers with each list in turn, then keeps repeating the last one. */
function statusSequence(...replies: (() => SyncComponent[])[]): MockRoute {
  let call = 0;
  return {
    path: STATUS,
    reply: () => ({
      body: { deviceId: ID, components: (replies[Math.min(call++, replies.length - 1)] as () => SyncComponent[])() },
    }),
  };
}

const syncRoutes = (status: MockRoute): MockRoute[] => [
  { path: `${DEVICES}/${ID}`, reply: { body: laptop } },
  { method: 'POST', path: `${DEVICES}/${ID}/syncDevice`, reply: { status: 204 } },
  status,
];

describe('device sync status', () => {
  const sync = commandById(commands, 'intune devices sync');

  test('without a terminal or --wait-seconds, sync returns once sent and points to sync-status', async () => {
    const { outcome, calls } = await runCommand(sync, [ID], syncRoutes(statusSequence(() => [])), {
      allowWrites: true,
    });
    expect(calls.some((c) => c.path === STATUS)).toBe(false);
    expect(outcome).toMatchObject({ kind: 'result', result: { status: 'requested' } });
    const result = outcome.kind === 'result' ? (outcome.result as { next: string; sync?: unknown }) : null;
    expect(result?.sync).toBeUndefined();
    expect(result?.next).toContain('tenantctl intune devices sync-status DESKTOP-4F2K9');
  });

  test('--wait-seconds reports the stages; a compliance check still running does not hold up the result', async () => {
    const status = statusSequence(() => [
      stage('notifyingDevice', 'success'),
      stage('deviceConnecting', 'success'),
      stage('policies', 'partialSuccess', 0, '35/37'),
      stage('applications', 'success', 0, '12 offered'),
      stage('scripts', 'success', 0, 'No script updates'),
      stage('compliance', 'inProgress'),
    ]);
    const { outcome, calls } = await runCommand(sync, [ID, '--wait-seconds', '60'], syncRoutes(status), {
      allowWrites: true,
    });
    const statusCall = calls.find((c) => c.path === STATUS);
    expect(statusCall?.version).toBe('beta');
    const result = outcome.kind === 'result' ? (outcome.result as Record<string, any>) : {};
    expect(result.sync).toMatchObject({ state: 'partialSuccess', timedOut: false });
    expect(result.sync.stages).toContainEqual(expect.objectContaining({ stage: 'Policies', detail: '35/37' }));
    expect(result.next).toBeUndefined();
  });

  test('waiting ignores the previous sync until this one reports, and times out while the device is offline', async () => {
    const tenMinutesAgo = -600_000;
    const previous = () => [
      stage('notifyingDevice', 'success', tenMinutesAgo),
      stage('scripts', 'success', tenMinutesAgo),
    ];
    const notified = () => [stage('notifyingDevice', 'success')];
    const finished = () => [
      stage('notifyingDevice', 'success'),
      stage('deviceConnecting', 'success'),
      stage('policies', 'success', 0, '37/37'),
      stage('applications', 'success'),
      stage('scripts', 'success'),
    ];
    const options = { timeoutSeconds: 60, deviceName: 'DESKTOP-4F2K9', sleep: async () => {}, tickMs: 3000 };

    const done = mockContext([statusSequence(previous, notified, finished)]);
    const outcome = await waitForSync(done.ctx, ID, Date.now(), options);
    expect(outcome).toMatchObject({ state: 'succeeded', timedOut: false, waitedSeconds: 6 });
    expect(done.calls.filter((c) => c.path === STATUS)).toHaveLength(3);

    const offline = mockContext([statusSequence(previous, notified)]);
    const stuck = await waitForSync(offline.ctx, ID, Date.now(), { ...options, timeoutSeconds: 9 });
    expect(stuck).toMatchObject({ state: 'waitingForDevice', timedOut: true, waitedSeconds: 9 });
  });

  test('sync-status shows the latest sync with readable stage names', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'intune devices sync-status'),
      [ID],
      [
        { path: `${DEVICES}/${ID}`, reply: { body: laptop } },
        statusSequence(() => [stage('notifyingDevice', 'success')]),
      ],
    );
    expect(outcome).toMatchObject({
      kind: 'result',
      result: { state: 'waitingForDevice', stages: [{ stage: 'Notifying device', status: 'success' }] },
    });
  });
});
