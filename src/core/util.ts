import { CliError } from './errors.ts';

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function isGuid(value: string): boolean {
  return GUID.test(value);
}

/** ISO timestamp `days` days before now, for `ge` filters on date-time properties. */
export function daysAgo(days: number, now = Date.now()): string {
  return new Date(now - days * 86_400_000).toISOString();
}

/** Whole days since an ISO timestamp, or null when absent/unparseable. */
export function daysSince(iso: string | null | undefined, now = Date.now()): number | null {
  if (!iso) return null;
  const time = Date.parse(iso);
  if (Number.isNaN(time) || time <= 0) return null;
  return Math.floor((now - time) / 86_400_000);
}

export type SettledError = { error: { code: string; message: string } };
export type Settled<T> = { [K in keyof T]: Awaited<T[K]> | SettledError };

/** True for a section `settle()` turned into `{ error }` (and nothing else). */
export function isSettledError(value: unknown): value is SettledError {
  return typeof value === 'object' && value !== null && 'error' in value && Object.keys(value).length === 1;
}

/**
 * Run independent lookups concurrently; a failing section becomes `{ error }` instead of failing the whole
 * composite. Troubleshooting views stay useful when the caller lacks one permission.
 */
export async function settle<T extends Record<string, Promise<unknown>>>(tasks: T): Promise<Settled<T>> {
  const keys = Object.keys(tasks);
  const results = await Promise.allSettled(Object.values(tasks));
  const out: Record<string, unknown> = {};
  results.forEach((result, index) => {
    const key = keys[index] as string;
    if (result.status === 'fulfilled') {
      out[key] = result.value;
    } else {
      const reason = result.reason as unknown;
      out[key] = {
        error: {
          code: reason instanceof CliError ? reason.code : 'INTERNAL',
          message: reason instanceof Error ? reason.message : String(reason),
        },
      };
    }
  });
  return out as Settled<T>;
}

/** Build a compact object, dropping undefined values (keeps null: null means "Graph says empty"). */
export function compact<T extends Record<string, unknown>>(value: T): Partial<T> {
  const out: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value)) if (entry !== undefined) out[key] = entry;
  return out as Partial<T>;
}

export function bytesToGb(bytes: number | null | undefined): number | null {
  if (typeof bytes !== 'number') return null;
  return Math.round((bytes / 1024 ** 3) * 100) / 100;
}
