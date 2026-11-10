import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

/**
 * Runs the real CLI as a subprocess against a local fake Graph server: argv parsing, output envelope, exit codes,
 * write gating and the two-step confirm flow, end to end.
 */
const seen: { method: string; path: string; auth: string | null }[] = [];
let server: Bun.Server<undefined>;
const configDir = mkdtempSync(join(tmpdir(), 'm365-e2e-'));
// The agent-replayable token flow is what this file exercises; human approval (tty/dialog) is unit-tested.
writeFileSync(
  join(configDir, 'config.json'),
  JSON.stringify({ profiles: { default: { authMode: 'delegated', confirmMode: 'token' } } }),
);

beforeAll(() => {
  server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      seen.push({
        method: request.method,
        path: url.pathname + url.search,
        auth: request.headers.get('authorization'),
      });
      if (url.pathname === '/v1.0/users' && request.method === 'GET') {
        if (url.searchParams.get('$skiptoken')) return Response.json({ value: [{ id: '3', displayName: 'Cy' }] });
        return Response.json({
          value: [
            { id: '1', displayName: 'Ann', mail: 'ann@contoso.com' },
            { id: '2', displayName: 'Bo', mail: 'bo@contoso.com' },
          ],
          '@odata.nextLink': `${url.origin}/v1.0/users?$skiptoken=next`,
        });
      }
      if (url.pathname === '/v1.0/groups/g1' && request.method === 'DELETE') return new Response(null, { status: 204 });
      return Response.json(
        { error: { code: 'Request_ResourceNotFound', message: 'Resource does not exist.' } },
        { status: 404 },
      );
    },
  });
});

afterAll(() => server.stop(true));

async function cli(args: string[], env: Record<string, string> = {}) {
  const proc = Bun.spawn(['bun', join(import.meta.dir, '..', 'src', 'cli.ts'), ...args], {
    env: {
      PATH: process.env.PATH ?? '',
      HOME: configDir,
      M365_CONFIG_DIR: configDir,
      M365_ACCESS_TOKEN: 'e2e-token',
      M365_GRAPH_URL: `http://127.0.0.1:${server.port}`,
      ...env,
    },
    stdout: 'pipe',
    stderr: 'pipe',
  });
  const [stdout, stderr, code] = await Promise.all([
    new Response(proc.stdout).text(),
    new Response(proc.stderr).text(),
    proc.exited,
  ]);
  return { stdout, stderr, code };
}

describe('m365 CLI end to end', () => {
  test('JSON envelope with paging, field projection and bearer token', async () => {
    const { stdout, code } = await cli(['graph', 'get', '/users', '--top', '3', '--fields', 'displayName', '--json']);
    expect(code).toBe(0);
    expect(JSON.parse(stdout)).toEqual({
      ok: true,
      command: 'graph get',
      data: [{ displayName: 'Ann' }, { displayName: 'Bo' }, { displayName: 'Cy' }],
      meta: { count: 3, truncated: false },
    });
    expect(seen.find((r) => r.path.startsWith('/v1.0/users'))?.auth).toBe('Bearer e2e-token');
  });

  test('Graph 404 becomes NOT_FOUND with exit code 6', async () => {
    const { stdout, code } = await cli(['graph', 'get', '/nope', '--json']);
    expect(code).toBe(6);
    expect(JSON.parse(stdout).error.code).toBe('NOT_FOUND');
  });

  test('usage errors exit 2', async () => {
    expect((await cli(['graph', 'get', '--json'])).code).toBe(2);
    expect((await cli(['nosuch', '--json'])).code).toBe(2);
  });

  test('writes are refused until enabled, then need the previewed confirm token', async () => {
    const refused = await cli(['graph', 'delete', '/groups/g1', '--json']);
    expect(refused.code).toBe(7);
    expect(seen.some((r) => r.method === 'DELETE')).toBe(false);

    const preview = await cli(['graph', 'delete', '/groups/g1', '--json'], { M365_ALLOW_WRITES: '1' });
    expect(preview.code).toBe(3);
    const payload = JSON.parse(preview.stdout);
    expect(payload.preview).toEqual({
      method: 'DELETE',
      url: `http://127.0.0.1:${server.port}/v1.0/groups/g1`,
      body: null,
    });
    expect(payload.confirm.rerun).toBe(`m365 graph delete /groups/g1 --json --confirm ${payload.confirm.token}`);
    expect(seen.some((r) => r.method === 'DELETE')).toBe(false);

    const done = await cli(['graph', 'delete', '/groups/g1', '--json', '--confirm', payload.confirm.token], {
      M365_ALLOW_WRITES: '1',
    });
    expect(done.code).toBe(0);
    expect(seen.filter((r) => r.method === 'DELETE').map((r) => r.path)).toEqual(['/v1.0/groups/g1']);

    const audit = await cli(['config', 'audit-log', '--json']);
    expect(JSON.parse(audit.stdout).data.map((e: { outcome: string }) => e.outcome)).toEqual([
      'success',
      'confirmation_issued',
    ]);
  });

  test('table output for people', async () => {
    const { stdout, code } = await cli(['graph', 'get', '/users', '--top', '2', '-o', 'table']);
    expect(code).toBe(0);
    expect(stdout).toContain('displayName');
    expect(stdout).toContain('ann@contoso.com');
    expect(stdout).toContain('2 shown; More results exist');
  });
});
