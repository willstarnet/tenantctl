import { describe, expect, test } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { parseArgs } from '../src/core/args.ts';
import { CliError } from '../src/core/errors.ts';
import { parseCsv } from '../src/core/graph.ts';
import { resolveUser } from '../src/core/resolve.ts';
import { auditLogPath, issueConfirmToken, verifyConfirmToken } from '../src/core/safety.ts';
import { defineCommand } from '../src/core/types.ts';
import { graphError, mockContext, page, runCommand } from './helpers/graphMock.ts';

describe('parseArgs', () => {
  const specs = [
    { name: 'user', type: 'string', positional: true, required: true, description: '' },
    { name: 'days', type: 'number', default: 7, description: '' },
    { name: 'failuresOnly', type: 'boolean', description: '' },
    { name: 'groups', type: 'string[]', description: '' },
    { name: 'state', type: 'string', choices: ['on', 'off'], description: '' },
  ] as const;

  test('positional, kebab flags, defaults, arrays', () => {
    expect(parseArgs(specs, ['jane@contoso.com', '--failures-only', '--groups', 'a,b', '--groups=c'])).toEqual({
      user: 'jane@contoso.com',
      days: 7,
      failuresOnly: true,
      groups: ['a', 'b', 'c'],
    });
  });

  test('flag form of a positional and --no- negation', () => {
    expect(parseArgs(specs, ['--user', 'x', '--no-failures-only', '--days=30'])).toMatchObject({
      user: 'x',
      days: 30,
      failuresOnly: false,
    });
  });

  test('rejects bad numbers, choices, unknown flags, missing required, extra positionals', () => {
    expect(() => parseArgs(specs, ['x', '--days', 'ten'])).toThrow('expects a number');
    expect(() => parseArgs(specs, ['x', '--state', 'maybe'])).toThrow('must be one of');
    expect(() => parseArgs(specs, ['x', '--nope'])).toThrow('Unknown option --nope');
    expect(() => parseArgs(specs, [])).toThrow('Missing required');
    expect(() => parseArgs(specs, ['x', 'y'])).toThrow('Unexpected argument');
  });
});

describe('GraphClient', () => {
  test('list pages through nextLink and stops at the limit', async () => {
    const { ctx, calls } = mockContext([
      {
        path: '/users',
        reply: (req) =>
          req.query.$skiptoken
            ? page([{ id: '3' }, { id: '4' }])
            : page([{ id: '1' }, { id: '2' }], 'https://graph.test/v1.0/users?$skiptoken=x'),
      },
    ]);
    const result = await ctx.graph.list('/users', { limit: 3 });
    expect(result.items.map((u) => u.id)).toEqual(['1', '2', '3']);
    expect(result.truncated).toBe(true);
    expect(calls[0]?.query.$top).toBe('3');
    const all = await ctx.graph.list('/users', { limit: Number.POSITIVE_INFINITY });
    expect(all).toEqual({ items: [{ id: '1' }, { id: '2' }, { id: '3' }, { id: '4' }], truncated: false });
  });

  test('maps Graph errors to CLI error codes', async () => {
    const { ctx } = mockContext([
      { path: '/users/x', reply: graphError(403, 'Authorization_RequestDenied', 'Insufficient privileges') },
    ]);
    const error = await ctx.graph.get('/users/x').catch((e: unknown) => e);
    expect(error).toBeInstanceOf(CliError);
    expect((error as CliError).code).toBe('FORBIDDEN');
    expect((error as CliError).requestId).toBe('test-request-id');
    expect(await ctx.graph.tryGet('/users/missing')).toBeNull();
  });

  test('parseCsv handles BOM, quotes, embedded commas and CRLF', () => {
    expect(parseCsv('\uFEFFName,Note\r\n"Doe, Jane","said ""hi"""\r\nBob,\r\n')).toEqual([
      { Name: 'Doe, Jane', Note: 'said "hi"' },
      { Name: 'Bob', Note: '' },
    ]);
  });
});

describe('resolveUser', () => {
  test('UPN goes straight to /users/{upn}', async () => {
    const { ctx, calls } = mockContext([
      { path: '/users/jane@contoso.com', reply: { body: { id: 'u1', displayName: 'Jane' } } },
    ]);
    expect((await resolveUser(ctx, 'jane@contoso.com')).id).toBe('u1');
    expect(calls).toHaveLength(1);
  });

  test('a display name with several matches is AMBIGUOUS with candidates, unless one matches exactly', async () => {
    const users = [
      { id: 'u1', displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com' },
      { id: 'u2', displayName: 'Jane Doe-Smith', userPrincipalName: 'jds@contoso.com' },
    ];
    const { ctx } = mockContext([{ path: '/users', reply: page(users) }]);
    expect((await resolveUser(ctx, 'jane doe')).id).toBe('u1');
    const error = (await resolveUser(ctx, 'Doe').catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
    expect(error.details).toHaveLength(2);
  });
});

describe('write gating and confirmation', () => {
  let runs = 0;
  const wipe = defineCommand({
    domain: 'test',
    name: 'wipe',
    summary: 'wipe',
    args: [{ name: 'device', type: 'string', required: true, positional: true, description: '' }],
    risk: 'destructive',
    permissions: { application: ['X.ReadWrite.All'] },
    preview: async (_ctx, args) => ({ device: args.device, action: 'wipe' }),
    run: async () => {
      runs++;
      return { done: true };
    },
  });

  test('writes disabled by default', async () => {
    const error = (await runCommand(wipe, ['pc1'], []).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('WRITES_DISABLED');
    expect(runs).toBe(0);
  });

  test('dry run previews even with writes disabled', async () => {
    const { outcome } = await runCommand(wipe, ['pc1'], [], { dryRun: true });
    expect(outcome).toEqual({
      kind: 'result',
      result: { dryRun: true, command: 'test wipe', risk: 'destructive', preview: { device: 'pc1', action: 'wipe' } },
    });
  });

  test('destructive needs a token bound to the same args; the token then executes it', async () => {
    const first = await runCommand(wipe, ['pc1'], [], { allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    expect(runs).toBe(0);
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const wrongTarget = (await runCommand(wipe, ['pc2'], [], { allowWrites: true, confirm: token }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(wrongTarget.code).toBe('CONFIRM_INVALID');
    const second = await runCommand(wipe, ['pc1'], [], { allowWrites: true, confirm: token });
    expect(second.outcome).toEqual({ kind: 'result', result: { done: true } });
    expect(runs).toBe(1);
    const log = readFileSync(auditLogPath(), 'utf8')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l) as { outcome: string });
    expect(log.map((e) => e.outcome)).toEqual(['confirmation_issued', 'success']);
  });

  test('expired tokens are refused', () => {
    const binding = { command: 'test wipe', profile: 'p', args: { device: 'pc1' } };
    const { token } = issueConfirmToken(binding, Date.now() - 11 * 60 * 1000);
    expect(() => verifyConfirmToken(token, binding)).toThrow('expired');
  });

  test('FORBIDDEN errors gain a permission hint for the command', async () => {
    const read = defineCommand({
      domain: 'test',
      name: 'read',
      summary: 'read',
      args: [],
      risk: 'read',
      permissions: { delegated: ['User.Read.All'], application: ['User.Read.All'] },
      run: (ctx) => ctx.graph.get('/users'),
    });
    const error = (await runCommand(read, [], [{ path: '/users', reply: graphError(403, 'Forbidden', 'no') }]).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(error.code).toBe('FORBIDDEN');
    expect(error.hint).toContain('User.Read.All');
    expect(existsSync(auditLogPath())).toBe(true);
  });
});
