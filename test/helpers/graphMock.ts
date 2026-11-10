import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { Profile } from '../../src/core/config.ts';
import { CliError } from '../../src/core/errors.ts';
import { execute, type Outcome } from '../../src/core/execute.ts';
import type { ExoRunner, ExoStep, ExoStepResult } from '../../src/core/exo.ts';
import { GraphClient } from '../../src/core/graph.ts';
import type { AnyCommand, Context } from '../../src/core/types.ts';

/**
 * A fake Microsoft Graph for tests. Routes match on method + path without the version prefix
 * (`/users/abc`, not `/v1.0/users/abc`). Unmatched requests answer 404 with a Graph-shaped error, so a test
 * fails loudly when a command calls something unexpected.
 */
export interface MockRequest {
  method: string;
  /** Path without `/v1.0` or `/beta`, decoded. */
  path: string;
  version: string;
  /** Decoded query parameters, e.g. `{ $filter: "mail eq 'x'" }`. */
  query: Record<string, string>;
  body: unknown;
  headers: Record<string, string>;
}

export interface MockReply {
  status?: number;
  body?: unknown;
  headers?: Record<string, string>;
}

export interface MockRoute {
  method?: string;
  path: string | RegExp;
  /** Static reply, or a function of the request. Default: 200 with `{}`. */
  reply?: MockReply | ((request: MockRequest) => MockReply);
}

export interface GraphMock {
  fetch: typeof fetch;
  calls: MockRequest[];
}

export function graphError(status: number, code: string, message: string): MockReply {
  return { status, body: { error: { code, message, innerError: { 'request-id': 'test-request-id' } } } };
}

export function createGraphMock(routes: MockRoute[]): GraphMock {
  const calls: MockRequest[] = [];
  const fetchImpl = async (input: string | URL | Request, init?: RequestInit): Promise<Response> => {
    const url = new URL(typeof input === 'string' ? input : input instanceof URL ? input.href : input.url);
    const method = (init?.method ?? 'GET').toUpperCase();
    const match = url.pathname.match(/^\/(v1\.0|beta)(\/.*)$/);
    const request: MockRequest = {
      method,
      path: decodeURIComponent(match?.[2] ?? url.pathname),
      version: match?.[1] ?? '',
      query: Object.fromEntries([...url.searchParams.entries()]),
      body: typeof init?.body === 'string' ? safeJson(init.body) : undefined,
      headers: Object.fromEntries(Object.entries((init?.headers as Record<string, string> | undefined) ?? {})),
    };
    calls.push(request);
    const route = routes.find(
      (r) =>
        (r.method ?? 'GET').toUpperCase() === method &&
        (typeof r.path === 'string' ? r.path === request.path : r.path.test(request.path)),
    );
    const reply: MockReply = route
      ? typeof route.reply === 'function'
        ? route.reply(request)
        : (route.reply ?? {})
      : graphError(404, 'Request_ResourceNotFound', `No mock for ${method} ${request.path}`);
    const status = reply.status ?? 200;
    if (status === 204) return new Response(null, { status, headers: reply.headers });
    const isText = typeof reply.body === 'string';
    return new Response(isText ? (reply.body as string) : JSON.stringify(reply.body ?? {}), {
      status,
      headers: { 'content-type': isText ? 'text/plain' : 'application/json', ...reply.headers },
    });
  };
  return { fetch: fetchImpl as typeof fetch, calls };
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

/**
 * A fake Exchange PowerShell runner. `handler` answers each step (by cmdlet and params) with the objects the
 * cmdlet would return, or throws to simulate a cmdlet error. Every step is recorded in `steps`.
 */
export function fakeExo(
  handler: (step: ExoStep) => Record<string, unknown>[] = () => [],
): ExoRunner & { steps: ExoStep[] } {
  const steps: ExoStep[] = [];
  return {
    steps,
    async run(batch) {
      const results: ExoStepResult[] = [];
      for (const step of batch) {
        steps.push(step);
        try {
          results.push({ ok: true, data: handler(step) });
        } catch (error) {
          results.push({ ok: false, data: [], error: error instanceof Error ? error.message : String(error) });
          if (!step.continueOnError) break;
        }
      }
      return results;
    },
  };
}

/** A Context whose Graph and Defender clients talk to the mock, with no real auth or retries. */
export function mockContext(
  routes: MockRoute[],
  options: { profile?: Partial<Profile>; full?: boolean; exo?: ExoRunner; approve?: boolean } = {},
): { ctx: Context; calls: MockRequest[]; logs: string[] } {
  const mock = createGraphMock(routes);
  const logs: string[] = [];
  const client = (baseUrl: string, versioned: boolean) =>
    new GraphClient({
      baseUrl,
      getToken: async () => 'test-token',
      fetch: mock.fetch,
      maxRetries: 0,
      sleep: async () => {},
      versioned,
    });
  const ctx: Context = {
    graph: client('https://graph.test', true),
    defender: client('https://defender.test/api', false),
    exo: options.exo ?? fakeExo(),
    profileName: 'test',
    profile: { authMode: 'delegated', tenantId: 'tenant-test', ...options.profile },
    full: options.full ?? false,
    log: (message) => logs.push(message),
    async requireApproval() {
      if (options.approve === false) throw new CliError('DENIED', 'The change was not approved');
    },
  };
  return { ctx, calls: mock.calls, logs };
}

/** Fresh, empty config directory for tests that touch config, confirm keys or the audit log. */
export function tempConfigDir(): string {
  return mkdtempSync(join(tmpdir(), 'm365-test-'));
}

/** Collection response body. */
export function page(value: unknown[], nextLink?: string): MockReply {
  return { body: { value, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) } };
}

/**
 * Run a command exactly as the CLI would (arg parsing, write gating, confirm flow, audit) against the mock.
 * Uses the static token set by test/setup.ts.
 */
export async function runCommand(
  command: AnyCommand,
  argv: string[],
  routes: MockRoute[],
  options: {
    allowWrites?: boolean;
    confirm?: string;
    dryRun?: boolean;
    full?: boolean;
    exo?: ExoRunner;
    /** Human approval answer for destructive commands (default: approved). */
    approve?: boolean;
    profile?: Partial<Profile>;
  } = {},
): Promise<{ outcome: Outcome; calls: MockRequest[] }> {
  const mock = createGraphMock(routes);
  const outcome = await execute(
    command,
    command.name ? `${command.domain} ${command.name}` : command.domain,
    argv,
    { full: options.full ?? false, debug: false, dryRun: options.dryRun ?? false, confirm: options.confirm },
    {
      profileName: 'test',
      profile: {
        authMode: 'delegated',
        tenantId: 'tenant-test',
        allowWrites: options.allowWrites ?? false,
        ...options.profile,
      },
      fetch: mock.fetch,
      log: () => {},
      approver: async () => options.approve ?? true,
      exo: options.exo ?? fakeExo(),
    },
  );
  return { outcome, calls: mock.calls };
}

/** Look up a command by its full id, failing the test if it does not exist. */
export function commandById(commands: AnyCommand[], id: string): AnyCommand {
  const found = commands.find((c) => (c.name ? `${c.domain} ${c.name}` : c.domain) === id);
  if (!found) throw new Error(`No command "${id}"`);
  return found;
}
