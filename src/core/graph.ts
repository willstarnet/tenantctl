import { CliError, type ErrorCode } from './errors.ts';
import type { ListResult } from './types.ts';

export type ApiVersion = 'v1.0' | 'beta';
export type QueryValue = string | number | boolean | undefined;

export interface RequestOptions {
  /** OData and plain query parameters; undefined values are dropped. Keys keep their `$`. */
  query?: Record<string, QueryValue>;
  version?: ApiVersion;
  headers?: Record<string, string>;
  /**
   * Sends `ConsistencyLevel: eventual`. Required on directory objects (users, groups, devices, apps) for
   * `$search`, `$count`, and advanced filters (`endsWith`, `ne`, `not`, filtering on collections).
   */
  advanced?: boolean;
  body?: unknown;
}

export interface ListOptions extends RequestOptions {
  /** Maximum items to return; `Infinity` pages through everything. Default 50. */
  limit?: number;
  /** `$top` per page, capped by `limit`. `false` for endpoints that reject `$top`. Default 100. */
  pageSize?: number | false;
}

export interface GraphClientOptions {
  baseUrl: string;
  getToken: () => Promise<string>;
  fetch?: typeof fetch;
  debug?: (message: string) => void;
  maxRetries?: number;
  sleep?: (ms: number) => Promise<void>;
  /** False for OData APIs without a version segment (Defender for Endpoint: `<base>/api/machines`). */
  versioned?: boolean;
}

/** One request inside a `$batch`; `path` is relative to the version (`/users/x/mailFolders/inbox/messageRules`). */
export interface BatchRequest {
  method?: string;
  path: string;
  query?: Record<string, QueryValue>;
  headers?: Record<string, string>;
  body?: unknown;
}

/** A batched response: the Graph status and body, or a mapped error the caller can report per item. */
export interface BatchResult<T = unknown> {
  status: number;
  body: T;
  error?: CliError;
}

/** Graph accepts at most 20 requests per `$batch`. */
const BATCH_SIZE = 20;

interface GraphErrorBody {
  error?: { code?: string; message?: string; innerError?: { 'request-id'?: string; 'client-request-id'?: string } };
}

const RETRYABLE = new Set([429, 502, 503, 504]);
const MAX_RETRY_WAIT_MS = 30_000;

/** Encode one URL path segment (UPNs may contain `#`, `'`, `/` in guest accounts). */
export function seg(value: string | number): string {
  return encodeURIComponent(String(value));
}

/** Quote a string literal for an OData `$filter`. */
export function odata(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

function buildQuery(query: Record<string, QueryValue> | undefined): string {
  if (!query) return '';
  const parts: string[] = [];
  for (const [key, value] of Object.entries(query)) {
    if (value === undefined) continue;
    parts.push(`${key}=${encodeURIComponent(String(value))}`);
  }
  return parts.length ? `?${parts.join('&')}` : '';
}

function errorCodeFor(status: number): ErrorCode {
  if (status === 401) return 'AUTH_FAILED';
  if (status === 403) return 'FORBIDDEN';
  if (status === 404) return 'NOT_FOUND';
  if (status === 429) return 'THROTTLED';
  return 'GRAPH_ERROR';
}

function hintFor(status: number, graphCode: string | undefined): string | undefined {
  if (status === 401)
    return 'The token was rejected. Run `tenantctl auth status`, then `tenantctl auth login` if it expired.';
  if (status === 429) return 'Graph is throttling this tenant/app. Wait a minute and retry with a smaller --top.';
  if (graphCode === 'Request_UnsupportedQuery' || graphCode === 'BadRequest') {
    return 'Graph rejected the query; check filter syntax and values.';
  }
  return undefined;
}

/** One error shape for direct and batched failures, so hints and exit codes match either way. */
function graphError(
  status: number,
  label: string,
  message: string,
  parsed: GraphErrorBody,
  headerRequestId: string | null,
): CliError {
  const graphCode = parsed.error?.code;
  return new CliError(errorCodeFor(status), `${label}: ${message}`, {
    status,
    graphCode,
    requestId: parsed.error?.innerError?.['request-id'] ?? headerRequestId ?? undefined,
    hint: hintFor(status, graphCode),
  });
}

export class GraphClient {
  private readonly baseUrl: string;
  private readonly getToken: () => Promise<string>;
  private readonly fetchImpl: typeof fetch;
  private readonly debug?: (message: string) => void;
  private readonly maxRetries: number;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly versioned: boolean;

  constructor(options: GraphClientOptions) {
    this.baseUrl = options.baseUrl.replace(/\/+$/, '');
    this.getToken = options.getToken;
    this.fetchImpl = options.fetch ?? fetch;
    this.debug = options.debug;
    this.maxRetries = options.maxRetries ?? 4;
    this.sleep = options.sleep ?? ((ms) => Bun.sleep(ms));
    this.versioned = options.versioned ?? true;
  }

  url(path: string, options: RequestOptions = {}): string {
    if (/^https?:\/\//.test(path)) return path;
    const version = options.version ?? 'v1.0';
    const query = buildQuery(options.query);
    const joined = path.includes('?') && query ? `${path}&${query.slice(1)}` : `${path}${query}`;
    const relative = joined.startsWith('/') ? joined : `/${joined}`;
    return this.versioned ? `${this.baseUrl}/${version}${relative}` : `${this.baseUrl}${relative}`;
  }

  private async send(method: string, url: string, options: RequestOptions, redirect: 'follow' | 'manual' = 'follow') {
    const headers: Record<string, string> = {
      Authorization: `Bearer ${await this.getToken()}`,
      Accept: 'application/json',
      ...(options.advanced ? { ConsistencyLevel: 'eventual' } : {}),
      ...options.headers,
    };
    let body: string | undefined;
    if (options.body !== undefined) {
      headers['Content-Type'] ??= 'application/json';
      body = typeof options.body === 'string' ? options.body : JSON.stringify(options.body);
    }
    for (let attempt = 0; ; attempt++) {
      const started = Date.now();
      let response: Response;
      try {
        response = await this.fetchImpl(url, { method, headers, body, redirect });
      } catch (error) {
        if (attempt >= this.maxRetries) {
          throw new CliError('GRAPH_ERROR', `Network error calling Graph: ${(error as Error).message}`);
        }
        await this.sleep(1000 * 2 ** attempt);
        continue;
      }
      this.debug?.(`${method} ${url} -> ${response.status} (${Date.now() - started} ms)`);
      if (RETRYABLE.has(response.status) && attempt < this.maxRetries) {
        const retryAfter = Number(response.headers.get('retry-after'));
        const wait = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter * 1000 : 1000 * 2 ** attempt;
        await response.body?.cancel();
        await this.sleep(Math.min(wait, MAX_RETRY_WAIT_MS));
        continue;
      }
      return response;
    }
  }

  private async fail(method: string, url: string, response: Response): Promise<never> {
    const text = await response.text();
    let parsed: GraphErrorBody = {};
    try {
      parsed = JSON.parse(text) as GraphErrorBody;
    } catch {
      // Non-JSON error bodies (gateway pages) fall through to the status text.
    }
    const message = parsed.error?.message || text.slice(0, 300) || response.statusText;
    const path = url.replace(this.baseUrl, '').split('?')[0];
    throw graphError(response.status, `${method} ${path}`, message, parsed, response.headers.get('request-id'));
  }

  async request<T = unknown>(method: string, path: string, options: RequestOptions = {}): Promise<T> {
    return (await this.requestWithHeaders<T>(method, path, options)).body;
  }

  /**
   * Like `request`, but also returns status and headers: async operations (team creation, archive) answer 202
   * with the new resource and operation ids only in `Location` / `Content-Location`.
   */
  async requestWithHeaders<T = unknown>(
    method: string,
    path: string,
    options: RequestOptions = {},
  ): Promise<{ status: number; headers: Headers; body: T }> {
    const url = this.url(path, options);
    const response = await this.send(method, url, options);
    if (!response.ok) return this.fail(method, url, response);
    const text = response.status === 204 ? '' : await response.text();
    const type = response.headers.get('content-type') ?? '';
    const body = (text ? (type.includes('json') ? JSON.parse(text) : text) : undefined) as T;
    return { status: response.status, headers: response.headers, body };
  }

  get<T = unknown>(path: string, options?: RequestOptions): Promise<T> {
    return this.request<T>('GET', path, options);
  }

  post<T = unknown>(path: string, body?: unknown, options: RequestOptions = {}): Promise<T> {
    return this.request<T>('POST', path, { ...options, body });
  }

  patch<T = unknown>(path: string, body: unknown, options: RequestOptions = {}): Promise<T> {
    return this.request<T>('PATCH', path, { ...options, body });
  }

  put<T = unknown>(path: string, body: unknown, options: RequestOptions = {}): Promise<T> {
    return this.request<T>('PUT', path, { ...options, body });
  }

  delete<T = unknown>(path: string, options?: RequestOptions): Promise<T> {
    return this.request<T>('DELETE', path, options);
  }

  /** GET that returns null on 404 instead of throwing. */
  async tryGet<T = unknown>(path: string, options?: RequestOptions): Promise<T | null> {
    try {
      return await this.get<T>(path, options);
    } catch (error) {
      if (error instanceof CliError && error.code === 'NOT_FOUND') return null;
      throw error;
    }
  }

  /**
   * Send independent requests through `$batch`, 20 per call, results in input order. Items Graph throttles
   * (429/502/503/504) are retried in later rounds after the longest Retry-After. A failed item carries `error`
   * instead of failing the batch, so composite commands can report per mailbox/device.
   */
  async batch<T = unknown>(requests: BatchRequest[], version: ApiVersion = 'v1.0'): Promise<BatchResult<T>[]> {
    const results: BatchResult<T>[] = new Array(requests.length);
    let pending = requests.map((_, index) => index);
    for (let attempt = 0; pending.length; attempt++) {
      const retry: number[] = [];
      let wait = 0;
      for (let start = 0; start < pending.length; start += BATCH_SIZE) {
        const chunk = pending.slice(start, start + BATCH_SIZE);
        const payload = {
          requests: chunk.map((index) => {
            const request = requests[index] as BatchRequest;
            const relative = request.path.startsWith('/') ? request.path : `/${request.path}`;
            const hasBody = request.body !== undefined;
            return {
              id: String(index),
              method: request.method ?? 'GET',
              url: `${relative}${buildQuery(request.query)}`,
              headers: { ...(hasBody ? { 'Content-Type': 'application/json' } : {}), ...request.headers },
              ...(hasBody ? { body: request.body } : {}),
            };
          }),
        };
        const response = await this.post<{
          responses?: { id: string; status: number; headers?: Record<string, string>; body?: unknown }[];
        }>('/$batch', payload, { version });
        for (const item of response?.responses ?? []) {
          const index = Number(item.id);
          if (RETRYABLE.has(item.status) && attempt < this.maxRetries) {
            retry.push(index);
            const after = Number(item.headers?.['Retry-After'] ?? item.headers?.['retry-after']);
            wait = Math.max(wait, Number.isFinite(after) && after > 0 ? after * 1000 : 1000 * 2 ** attempt);
            continue;
          }
          const body = item.body as T;
          if (item.status < 400) {
            results[index] = { status: item.status, body };
            continue;
          }
          const parsed = (item.body ?? {}) as GraphErrorBody;
          const label = `${requests[index]?.method ?? 'GET'} ${requests[index]?.path ?? ''}`;
          const message = parsed.error?.message || `HTTP ${item.status}`;
          results[index] = { status: item.status, body, error: graphError(item.status, label, message, parsed, null) };
        }
      }
      pending = retry;
      if (pending.length) await this.sleep(Math.min(wait, MAX_RETRY_WAIT_MS));
    }
    return results;
  }

  /** Follow `@odata.nextLink` until `limit` items are collected. */
  async list<T = Record<string, unknown>>(path: string, options: ListOptions = {}): Promise<ListResult<T>> {
    const limit = options.limit ?? 50;
    const pageSize = options.pageSize ?? 100;
    const query = { ...options.query };
    if (pageSize !== false && query.$top === undefined) query.$top = Math.min(limit, pageSize);
    let next: string | undefined = this.url(path, { ...options, query });
    const items: T[] = [];
    while (next) {
      const page: { value?: T[]; '@odata.nextLink'?: string } = await this.get(next, options);
      for (const item of page.value ?? []) {
        if (items.length >= limit) return { items, truncated: true };
        items.push(item);
      }
      next = page['@odata.nextLink'];
      if (next && items.length >= limit) return { items, truncated: true };
    }
    return { items, truncated: false };
  }

  /**
   * Usage reports (`/reports/get...`) answer with a redirect to a pre-authenticated CSV download. Follow it
   * without forwarding the bearer token and parse the CSV into objects keyed by header.
   */
  async report(path: string, options: RequestOptions = {}): Promise<Record<string, string>[]> {
    const url = this.url(path, options);
    const response = await this.send('GET', url, { ...options, headers: { Accept: 'text/csv' } }, 'manual');
    let csv: string;
    if (response.status >= 300 && response.status < 400) {
      const location = response.headers.get('location');
      if (!location) throw new CliError('GRAPH_ERROR', `Report ${path} redirected without a location`);
      const download = await this.fetchImpl(location);
      if (!download.ok) throw new CliError('GRAPH_ERROR', `Report download failed: ${download.status}`);
      csv = await download.text();
    } else if (response.ok) {
      csv = await response.text();
    } else {
      return this.fail('GET', url, response);
    }
    return parseCsv(csv);
  }
}

/** RFC 4180 CSV -> objects. Strips the UTF-8 BOM Graph reports start with. */
export function parseCsv(text: string): Record<string, string>[] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  const input = text.replace(/^\uFEFF/, '');
  for (let i = 0; i < input.length; i++) {
    const char = input[i];
    if (quoted) {
      if (char === '"' && input[i + 1] === '"') {
        field += '"';
        i++;
      } else if (char === '"') {
        quoted = false;
      } else {
        field += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ',') {
      row.push(field);
      field = '';
    } else if (char === '\n' || char === '\r') {
      if (char === '\r' && input[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += char;
    }
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  const [header, ...data] = rows.filter((r) => r.some((cell) => cell !== ''));
  if (!header) return [];
  return data.map((cells) => Object.fromEntries(header.map((name, index) => [name, cells[index] ?? ''])));
}
