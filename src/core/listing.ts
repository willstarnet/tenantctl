/** Standard paging args for every list command. Spread into `args` and pass `listLimit(args)` to `graph.list`. */
export const LIST_ARGS = [
  { name: 'top', type: 'number', default: 50, description: 'Maximum results to return.' },
  { name: 'all', type: 'boolean', description: 'Return every result (pages through everything; can be slow).' },
] as const;

export function listLimit(args: { top: number; all: boolean }): number {
  return args.all ? Number.POSITIVE_INFINITY : args.top;
}
