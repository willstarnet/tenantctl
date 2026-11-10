import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { effectiveConfirmMode, weakens } from '../src/core/approval.ts';
import { loadConfig } from '../src/core/config.ts';
import type { CliError } from '../src/core/errors.ts';
import { auditLogPath } from '../src/core/safety.ts';
import { decryptText, encryptText } from '../src/core/secretStore.ts';
import { defineCommand } from '../src/core/types.ts';
import { COMMANDS } from '../src/registry.ts';
import { GraphClient } from '../src/core/graph.ts';
import { commandById, createGraphMock, graphError, mockContext, runCommand } from './helpers/graphMock.ts';

afterEach(() => {
  delete process.env.M365_CONFIRM_MODE;
});

describe('human approval mode', () => {
  test('the environment can make approval stricter but never weaker', () => {
    process.env.M365_CONFIRM_MODE = 'token';
    expect(effectiveConfirmMode({ authMode: 'delegated', confirmMode: 'tty' })).toBe('tty');
    process.env.M365_CONFIRM_MODE = 'dialog';
    expect(effectiveConfirmMode({ authMode: 'delegated', confirmMode: 'token' })).toBe('dialog');
  });

  test('moving to a weaker mode is detected', () => {
    expect(weakens({ authMode: 'delegated', confirmMode: 'dialog' }, 'token')).toBe(true);
    expect(weakens({ authMode: 'delegated', confirmMode: 'token' }, 'tty')).toBe(false);
  });

  let wiped = 0;
  const wipe = defineCommand({
    domain: 'test',
    name: 'wipe2',
    summary: 'wipe',
    args: [{ name: 'device', type: 'string', required: true, positional: true, description: '' }],
    risk: 'destructive',
    permissions: { application: ['X.ReadWrite.All'] },
    preview: async (_ctx, args) => ({ device: args.device, action: `Factory reset ${args.device}` }),
    run: async () => {
      wiped++;
      return { done: true };
    },
  });

  test('a valid confirm token is not enough when the person denies the change', async () => {
    const first = await runCommand(wipe, ['pc1'], [], { allowWrites: true });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const error = (await runCommand(wipe, ['pc1'], [], { allowWrites: true, confirm: token, approve: false }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(error.code).toBe('DENIED');
    expect(error.exitCode).toBe(10);
    expect(wiped).toBe(0);
    const outcomes = readFileSync(auditLogPath(), 'utf8')
      .trim()
      .split('\n')
      .map((l) => (JSON.parse(l) as { outcome: string }).outcome);
    expect(outcomes.at(-1)).toBe('denied');
  });

  test('weakening the approval mode or enabling writes needs the person; denial leaves the setting alone', async () => {
    const set = commandById(COMMANDS, 'config set');
    const strict = { confirmMode: 'tty' as const };
    const weaken = (await runCommand(set, ['confirmMode', 'token'], [], { approve: false, profile: strict }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(weaken.code).toBe('DENIED');
    const enable = (await runCommand(set, ['allowWrites', 'true'], [], { approve: false, profile: strict }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(enable.code).toBe('DENIED');
    expect(loadConfig().profiles.test?.allowWrites).toBeUndefined();
    await runCommand(set, ['confirmMode', 'dialog'], [], { approve: false, profile: strict });
    expect(loadConfig().profiles.test?.confirmMode).toBe('dialog');
  });
});

describe('token cache encryption', () => {
  test('round-trips, rejects tampering, and recognizes an old plaintext cache', () => {
    const sealed = encryptText('{"Account":{}}');
    expect(sealed).not.toContain('Account');
    expect(decryptText(sealed)).toEqual({ plain: '{"Account":{}}', legacy: false });
    const tampered = JSON.parse(sealed) as { data: string };
    tampered.data = Buffer.from('{"Account":{"x":1}}').toString('base64');
    expect(decryptText(JSON.stringify(tampered))).toBeNull();
    expect(decryptText('{"Account":{},"RefreshToken":{}}')).toEqual({
      plain: '{"Account":{},"RefreshToken":{}}',
      legacy: true,
    });
  });
});

describe('$batch', () => {
  test('chunks by 20, keeps order, retries throttled items and reports per-item errors', async () => {
    let round = 0;
    const { fetch: mockFetch, calls } = createGraphMock([
      {
        method: 'POST',
        path: '/$batch',
        reply: (req) => {
          const { requests } = req.body as { requests: { id: string; url: string }[] };
          round++;
          return {
            body: {
              responses: requests.map((r) => {
                if (r.url === '/users/u3' && round === 1)
                  return { id: r.id, status: 429, headers: { 'Retry-After': '1' } };
                if (r.url === '/users/u5')
                  return {
                    id: r.id,
                    status: 404,
                    body: { error: { code: 'Request_ResourceNotFound', message: 'gone' } },
                  };
                return { id: r.id, status: 200, body: { id: r.url.split('/').pop() } };
              }),
            },
          };
        },
      },
    ]);
    const graph = new GraphClient({
      baseUrl: 'https://graph.test',
      getToken: async () => 't',
      fetch: mockFetch,
      maxRetries: 2,
      sleep: async () => {},
    });
    const results = await graph.batch<{ id: string }>(Array.from({ length: 25 }, (_, i) => ({ path: `/users/u${i}` })));
    expect(
      calls.filter((c) => c.path === '/$batch').map((c) => (c.body as { requests: unknown[] }).requests.length),
    ).toEqual([20, 5, 1]);
    expect(results.map((r) => r.body?.id ?? null).slice(0, 5)).toEqual(['u0', 'u1', 'u2', 'u3', 'u4']);
    expect(results[5]?.error?.code).toBe('NOT_FOUND');
  });

  test('unversioned clients (Defender for Endpoint) have no version segment', () => {
    const { ctx } = mockContext([]);
    expect(ctx.defender.url('/machines', { query: { $top: 1 } })).toBe('https://defender.test/api/machines?$top=1');
    expect(ctx.graph.url('/users')).toBe('https://graph.test/v1.0/users');
  });

  test('a batch item failing is not a batch failure', async () => {
    const { ctx } = mockContext([
      {
        method: 'POST',
        path: '/$batch',
        reply: { body: { responses: [{ id: '0', status: 403, body: graphError(403, 'Forbidden', 'no').body }] } },
      },
    ]);
    const [only] = await ctx.graph.batch([{ path: '/users/x/messages' }]);
    expect(only?.error?.code).toBe('FORBIDDEN');
  });
});
