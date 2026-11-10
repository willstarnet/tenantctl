import { CliError } from '../core/errors.ts';
import { listLimit, LIST_ARGS } from '../core/listing.ts';
import { type AnyCommand, defineCommand, type ListResult } from '../core/types.ts';

const VERSION_ARG = { name: 'beta', type: 'boolean', description: 'Call the beta endpoint instead of v1.0.' } as const;
const ADVANCED_ARG = {
  name: 'advanced',
  type: 'boolean',
  description: 'Send ConsistencyLevel: eventual ($search, $count and advanced filters on directory objects).',
} as const;
const PATH_ARG = {
  name: 'path',
  type: 'string',
  required: true,
  positional: true,
  description: 'Graph path after the version, with any query string, e.g. "/users?$select=id,displayName".',
} as const;

function parseBody(raw: string | undefined): unknown {
  if (raw === undefined) return undefined;
  try {
    return JSON.parse(raw) as unknown;
  } catch (error) {
    throw new CliError('USAGE', `--body is not valid JSON: ${(error as Error).message}`);
  }
}

function rawWrite(method: 'POST' | 'PATCH' | 'PUT' | 'DELETE'): AnyCommand {
  const withBody = method !== 'DELETE';
  return defineCommand({
    domain: 'graph',
    name: method.toLowerCase(),
    summary: `Raw ${method} to any Graph path`,
    description:
      'Escape hatch for operations no dedicated command covers. Always treated as destructive: it previews, ' +
      'then needs --confirm. Prefer a dedicated command when one exists.',
    examples: [
      `m365 graph ${method.toLowerCase()} /users/jane@contoso.com${withBody ? ` --body '{"jobTitle":"Manager"}'` : ''}`,
    ],
    args: [PATH_ARG, { name: 'body', type: 'string', description: 'JSON request body.' }, VERSION_ARG],
    risk: 'destructive',
    permissions: {},
    async preview(ctx, args) {
      return {
        method,
        url: ctx.graph.url(args.path, { version: args.beta ? 'beta' : 'v1.0' }),
        body: parseBody(args.body) ?? null,
      };
    },
    async run(ctx, args) {
      const result = await ctx.graph.request(method, args.path, {
        version: args.beta ? 'beta' : 'v1.0',
        body: parseBody(args.body),
      });
      return result ?? { status: 'done' };
    },
  });
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'graph',
    name: 'get',
    summary: 'Raw GET of any Graph path; collections are paged like list commands',
    description:
      'Use when no dedicated command covers what you need. Results come back unshaped. Needs whatever ' +
      'permission the endpoint documents.',
    examples: [
      'm365 graph get "/users?$filter=department eq \'Sales\'&$select=displayName,mail"',
      'm365 graph get /deviceManagement/managedDevices --beta --top 5',
    ],
    args: [PATH_ARG, VERSION_ARG, ADVANCED_ARG, ...LIST_ARGS],
    risk: 'read',
    permissions: {},
    async run(ctx, args) {
      const options = { version: args.beta ? ('beta' as const) : ('v1.0' as const), advanced: args.advanced };
      const first = await ctx.graph.get<Record<string, unknown> | string>(args.path, options);
      if (typeof first !== 'object' || first === null || !Array.isArray(first.value)) return first;
      const limit = listLimit(args);
      const items = (first.value as unknown[]).slice(0, limit);
      let next = first['@odata.nextLink'] as string | undefined;
      while (next && items.length < limit) {
        const page = await ctx.graph.get<{ value?: unknown[]; '@odata.nextLink'?: string }>(next, options);
        items.push(...(page.value ?? []).slice(0, limit - items.length));
        next = page['@odata.nextLink'];
      }
      return { items, truncated: Boolean(next) || (first.value as unknown[]).length > limit } satisfies ListResult;
    },
  }),
  rawWrite('POST'),
  rawWrite('PATCH'),
  rawWrite('PUT'),
  rawWrite('DELETE'),
];
