import { describe, expect, test } from 'bun:test';
import { COMMANDS, commandId, DOMAINS, findCommand, GLOBAL_FLAGS } from '../src/registry.ts';

/** Commands that answer from local data and legitimately call no Graph endpoint. */
const OFFLINE: Record<string, true> = { 'entra signin-error': true, 'defender hunt templates': true };
const LOCAL_AREAS: Record<string, true> = { auth: true, config: true, tools: true, permissions: true, graph: true };
const PERMISSION_NAME = /^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$/;

describe('command registry', () => {
  test('ids are unique and every command resolves to itself', () => {
    const ids = COMMANDS.map(commandId);
    expect(ids.length).toBe(new Set(ids).size);
    for (const command of COMMANDS) expect(findCommand(commandId(command).split(' '))?.command).toBe(command);
  });

  test('every command belongs to a documented area', () => {
    const areas = new Set(DOMAINS.map((d) => d.name));
    expect(COMMANDS.filter((c) => !areas.has(c.domain)).map(commandId)).toEqual([]);
  });

  test('destructive commands preview before acting', () => {
    expect(COMMANDS.filter((c) => c.risk === 'destructive' && !c.preview).map(commandId)).toEqual([]);
  });

  test('args are camelCase and never shadow global flags', () => {
    const bad = COMMANDS.flatMap((c) =>
      c.args
        .filter((a) => GLOBAL_FLAGS.includes(a.name) || !/^[a-z][a-zA-Z0-9]*$/.test(a.name))
        .map((a) => `${commandId(c)} --${a.name}`),
    );
    expect(bad).toEqual([]);
  });

  test('tenant commands declare well-formed permissions for the API they call', () => {
    const missing = COMMANDS.filter(
      (c) =>
        !LOCAL_AREAS[c.domain] &&
        !OFFLINE[commandId(c)] &&
        !c.permissions.delegated?.length &&
        !c.permissions.application?.length &&
        !c.permissions.exchange &&
        !c.permissions.defender?.length,
    ).map(commandId);
    expect(missing).toEqual([]);
    const malformed = COMMANDS.flatMap((c) =>
      [
        ...(c.permissions.delegated ?? []),
        ...(c.permissions.application ?? []),
        ...(c.permissions.defender ?? []),
      ].filter((p) => !PERMISSION_NAME.test(p)),
    );
    expect(malformed).toEqual([]);
  });

  test('Graph permission lists hold only Graph scopes (login requests them all; one foreign scope fails it)', () => {
    // Exchange Online and Defender for Endpoint names; Graph has no scopes with these prefixes.
    const foreign =
      /^(Exchange\.Manage|Machine\.|Alert\.Read|Alert\.ReadWrite|Ti\.|Vulnerability\.|Software\.Read|Score\.Read|AdvancedQuery\.)/;
    const leaks = COMMANDS.flatMap((c) =>
      [...(c.permissions.delegated ?? []), ...(c.permissions.application ?? [])]
        .filter((p) => foreign.test(p))
        .map((p) => `${commandId(c)}: ${p}`),
    );
    expect(leaks).toEqual([]);
  });

  test('a required positional never follows an optional one', () => {
    const bad = COMMANDS.filter((c) => {
      const positionals = c.args.filter((a) => a.positional);
      return positionals.some((a, i) => a.required && positionals.slice(0, i).some((p) => !p.required));
    }).map(commandId);
    expect(bad).toEqual([]);
  });
});
