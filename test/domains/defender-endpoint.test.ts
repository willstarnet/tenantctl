import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands, resolveMachine } from '../../src/domains/defender/index.ts';
import { commandById, type MockRoute, mockContext, page, runCommand } from '../helpers/graphMock.ts';

const cmd = (id: string) => commandById(commands, `defender ${id}`);

const laptop = {
  id: '1e5bc9d7e413ddd7902c2932e418702b84d0cc07',
  computerDnsName: 'laptop-7.contoso.com',
  osPlatform: 'Windows11',
  healthStatus: 'Active',
  riskScore: 'High',
  exposureLevel: 'Medium',
  lastSeen: new Date(Date.now() - 2 * 86_400_000).toISOString(),
  lastIpAddress: '10.0.0.7',
  aadDeviceId: '80fe8ff8-2624-418e-9591-41f0491218f9',
  machineTags: ['finance'],
  onboardingStatus: 'Onboarded',
};
const laptop70 = {
  ...laptop,
  id: 'aaaabbbbccccddddeeeeffff0000111122223333',
  computerDnsName: 'laptop-70.contoso.com',
};

/** GET /api/machines answering by which property the $filter targets. */
const machinesRoute: MockRoute = {
  path: '/api/machines',
  reply: (req) => {
    const filter = req.query.$filter ?? '';
    if (filter.includes('aadDeviceId')) return page(filter.includes(laptop.aadDeviceId) ? [laptop] : []);
    if (filter.includes("startswith(computerDnsName,'laptop-7')")) return page([laptop, laptop70]);
    return page([]);
  },
};

describe('resolveMachine', () => {
  test('a computer name picks the exact short-name match among prefix matches', async () => {
    const { ctx, calls } = mockContext([machinesRoute]);
    const machine = await resolveMachine(ctx, 'LAPTOP-7');
    expect(machine.id).toBe(laptop.id);
    expect(calls[0]?.path).toBe('/api/machines');
    expect(calls[0]?.query.$filter).toBe("startswith(computerDnsName,'laptop-7')");
  });

  test('an Entra device id resolves through aadDeviceId', async () => {
    const { ctx, calls } = mockContext([machinesRoute]);
    const machine = await resolveMachine(ctx, laptop.aadDeviceId);
    expect(machine.computerDnsName).toBe('laptop-7.contoso.com');
    expect(calls[0]?.query.$filter).toBe(`aadDeviceId eq '${laptop.aadDeviceId}'`);
  });

  test('a prefix shared by several machines is AMBIGUOUS', async () => {
    const { ctx } = mockContext([machinesRoute]);
    await expect(resolveMachine(ctx, 'laptop')).rejects.toMatchObject({
      code: 'NOT_FOUND',
    } satisfies Partial<CliError>);
    const twins = mockContext([{ path: '/api/machines', reply: page([laptop, laptop70]) }]);
    await expect(resolveMachine(twins.ctx, 'laptop')).rejects.toMatchObject({
      code: 'AMBIGUOUS',
    } satisfies Partial<CliError>);
  });
});

describe('defender machines', () => {
  test('list builds the MDE $filter and shapes machines', async () => {
    const { outcome, calls } = await runCommand(
      cmd('machines list'),
      ['--risk', 'high', '--health', 'active', '--os', 'Windows11', '--search', 'LAPTOP'],
      [{ path: '/api/machines', reply: page([laptop]) }],
    );
    expect(calls[0]?.query.$filter).toBe(
      "startswith(computerDnsName,'laptop') and riskScore eq 'High' and healthStatus eq 'Active' and osPlatform eq 'Windows11'",
    );
    expect(outcome).toMatchObject({
      kind: 'result',
      result: { items: [{ id: laptop.id, riskScore: 'High', daysSinceLastSeen: 2, machineTags: ['finance'] }] },
    });
  });

  test('get adds open alerts and recent actions', async () => {
    const { outcome, calls } = await runCommand(
      cmd('machines get'),
      ['laptop-7'],
      [
        machinesRoute,
        { path: `/api/machines/${laptop.id}/alerts/$count`, reply: { body: 4 } },
        {
          path: '/api/machineactions',
          reply: page([
            { id: 'old', type: 'RunAntiVirusScan', status: 'Succeeded', creationDateTimeUtc: '2026-11-01T00:00:00Z' },
            { id: 'new', type: 'Isolate', status: 'Pending', creationDateTimeUtc: '2026-11-09T00:00:00Z' },
          ]),
        },
      ],
    );
    expect(calls.find((c) => c.path === '/api/machineactions')?.query.$filter).toBe(`machineId eq '${laptop.id}'`);
    expect(calls.find((c) => c.path.endsWith('/alerts/$count'))?.query.$filter).toBe("status ne 'Resolved'");
    const result =
      outcome.kind === 'result' ? (outcome.result as { openAlerts: number; recentActions: { id: string }[] }) : null;
    expect(result?.openAlerts).toBe(4);
    expect(result?.recentActions.map((a) => a.id)).toEqual(['new', 'old']);
  });

  test('isolate previews and sends nothing, then the confirmed run posts Comment and IsolationType', async () => {
    const routes: MockRoute[] = [
      machinesRoute,
      {
        method: 'POST',
        path: `/api/machines/${laptop.id}/isolate`,
        reply: { status: 201, body: { id: 'act-1', type: 'Isolate', status: 'Pending' } },
      },
    ];
    const argv = ['laptop-7', '--selective', '--comment', 'Ransomware alert 1234'];
    const first = await runCommand(cmd('machines isolate'), argv, routes, { allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.method === 'POST')).toBe(false);
    const preview = first.outcome.kind === 'confirm' ? (first.outcome.preview as { action: string }) : null;
    expect(preview?.action).toBe('Isolate laptop-7.contoso.com from the network (Selective isolation)');

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(cmd('machines isolate'), argv, routes, { allowWrites: true, confirm: token });
    const post = second.calls.find((c) => c.method === 'POST');
    expect(post?.path).toBe(`/api/machines/${laptop.id}/isolate`);
    expect(post?.body).toEqual({ Comment: 'Ransomware alert 1234', IsolationType: 'Selective' });
    expect(second.outcome).toMatchObject({
      kind: 'result',
      result: {
        machineAction: { id: 'act-1', status: 'Pending' },
        next: 'm365 defender machines actions --machine laptop-7.contoso.com',
      },
    });
  });

  test('selective isolation is refused on Linux', async () => {
    const linux = { ...laptop, osPlatform: 'Linux' };
    const run = runCommand(
      cmd('machines isolate'),
      [laptop.aadDeviceId, '--selective', '--comment', 'x'],
      [{ path: '/api/machines', reply: page([linux]) }],
      { allowWrites: true },
    );
    await expect(run).rejects.toMatchObject({ code: 'USAGE' } satisfies Partial<CliError>);
  });

  test('scan posts ScanType Full with --full-scan', async () => {
    const { calls } = await runCommand(
      cmd('machines scan'),
      ['laptop-7', '--full-scan', '--comment', 'Alert follow-up'],
      [
        machinesRoute,
        {
          method: 'POST',
          path: `/api/machines/${laptop.id}/runAntiVirusScan`,
          reply: { status: 201, body: { id: 'act-2', status: 'Pending' } },
        },
      ],
      { allowWrites: true },
    );
    expect(calls.find((c) => c.method === 'POST')?.body).toEqual({ Comment: 'Alert follow-up', ScanType: 'Full' });
  });

  test('stop-quarantine-file rejects a non-SHA-1 hash', async () => {
    const run = runCommand(
      cmd('machines stop-quarantine-file'),
      ['laptop-7', '--sha1', 'abc', '--comment', 'x'],
      [machinesRoute],
      { allowWrites: true },
    );
    await expect(run).rejects.toMatchObject({ code: 'USAGE' } satisfies Partial<CliError>);
  });

  test('actions filters by machine and status', async () => {
    const { outcome, calls } = await runCommand(
      cmd('machines actions'),
      ['--machine', 'laptop-7', '--status', 'Pending'],
      [
        machinesRoute,
        {
          path: '/api/machineactions',
          reply: page([{ id: 'act-1', type: 'Isolate', status: 'Pending', computerDnsName: 'laptop-7.contoso.com' }]),
        },
      ],
    );
    expect(calls.find((c) => c.path === '/api/machineactions')?.query.$filter).toBe(
      `machineId eq '${laptop.id}' and status eq 'Pending'`,
    );
    expect(outcome).toMatchObject({
      kind: 'result',
      result: { items: [{ id: 'act-1', type: 'Isolate', status: 'Pending', machine: 'laptop-7.contoso.com' }] },
    });
  });

  test('package-url strips the quotes MDE wraps around the link', async () => {
    const { outcome } = await runCommand(
      cmd('machines package-url'),
      ['act-9'],
      [{ path: '/api/machineactions/act-9/GetPackageUri', reply: { body: { value: '"https://x.test/p.zip?t=1"' } } }],
    );
    expect(outcome).toMatchObject({ kind: 'result', result: { url: 'https://x.test/p.zip?t=1' } });
  });
});
