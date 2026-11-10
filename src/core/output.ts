import type { CliError } from './errors.ts';
import { isListResult } from './types.ts';

export type OutputFormat = 'json' | 'table' | 'csv';

export interface OutputOptions {
  format: OutputFormat;
  /** Keep only these (dotted) fields on each item / the result object. */
  fields?: string[];
}

function getPath(value: unknown, path: string): unknown {
  let current = value;
  for (const part of path.split('.')) {
    if (current === null || typeof current !== 'object') return undefined;
    current = (current as Record<string, unknown>)[part];
  }
  return current;
}

function project(value: unknown, fields: string[] | undefined): unknown {
  if (!fields?.length || value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.map((item) => project(item, fields));
  return Object.fromEntries(fields.map((field) => [field, getPath(value, field)]));
}

export interface Envelope {
  ok: true;
  command: string;
  data: unknown;
  meta?: { count: number; truncated: boolean; hint?: string };
}

export function envelope(command: string, result: unknown, fields?: string[]): Envelope {
  if (isListResult(result)) {
    return {
      ok: true,
      command,
      data: project(result.items, fields),
      meta: {
        count: result.items.length,
        truncated: result.truncated,
        ...(result.truncated ? { hint: 'More results exist: raise --top or pass --all.' } : {}),
      },
    };
  }
  return { ok: true, command, data: project(result, fields) };
}

function cell(value: unknown, width = 60): string {
  if (value === null || value === undefined) return '';
  let text = typeof value === 'object' ? JSON.stringify(value) : String(value);
  text = text.replace(/\s+/g, ' ');
  return text.length > width ? `${text.slice(0, width - 1)}…` : text;
}

function table(rows: Record<string, unknown>[], fields?: string[]): string {
  if (!rows.length) return '(none)';
  const columns =
    fields ??
    [...new Set(rows.slice(0, 20).flatMap((row) => Object.keys(row)))]
      .filter((key) => rows.some((row) => row[key] === null || typeof row[key] !== 'object'))
      .slice(0, 8);
  const body = rows.map((row) => columns.map((column) => cell(getPath(row, column))));
  const widths = columns.map((column, index) => Math.max(column.length, ...body.map((line) => line[index]!.length)));
  const line = (cells: string[]) =>
    cells
      .map((text, index) => text.padEnd(widths[index]!))
      .join('  ')
      .trimEnd();
  return [line(columns), line(widths.map((width) => '-'.repeat(width))), ...body.map(line)].join('\n');
}

function isRowArray(value: unknown): value is Record<string, unknown>[] {
  return (
    Array.isArray(value) && value.every((item) => item !== null && typeof item === 'object' && !Array.isArray(item))
  );
}

function describe(value: unknown, fields?: string[], indent = ''): string {
  if (isRowArray(value)) return table(value, fields);
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return `${indent}${cell(value, 200)}`;
  const scalars: string[] = [];
  const sections: string[] = [];
  const entries = Object.entries(value as Record<string, unknown>);
  const keyWidth = Math.max(0, ...entries.map(([key]) => key.length));
  for (const [key, entry] of entries) {
    if (isRowArray(entry) && entry.length) {
      sections.push(`\n${indent}${key} (${entry.length})\n${table(entry)}`);
    } else if (entry && typeof entry === 'object' && !Array.isArray(entry) && Object.keys(entry).length > 3) {
      sections.push(`\n${indent}${key}\n${describe(entry, undefined, `${indent}  `)}`);
    } else {
      scalars.push(`${indent}${key.padEnd(keyWidth)}  ${cell(entry, 120)}`);
    }
  }
  return [...scalars, ...sections].join('\n');
}

function csvCell(value: unknown): string {
  if (value === null || value === undefined) return '';
  const text = typeof value === 'object' ? JSON.stringify(value) : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

/**
 * Spreadsheet export: one row per list item (or one row for a single object). Columns are `--fields` when
 * given, else every key in first-seen order; nested values are JSON in their cell.
 */
function csv(data: unknown, fields: string[] | undefined): string {
  const rows = Array.isArray(data) ? data : [data];
  const objects = rows.map((row) =>
    row !== null && typeof row === 'object' ? (row as Record<string, unknown>) : { value: row },
  );
  const columns = fields?.length ? fields : [...new Set(objects.flatMap((row) => Object.keys(row)))];
  const lines = [columns.map(csvCell).join(',')];
  for (const row of objects) lines.push(columns.map((column) => csvCell(getPath(row, column))).join(','));
  return lines.join('\r\n');
}

export function renderSuccess(command: string, result: unknown, options: OutputOptions): string {
  const env = envelope(command, result, options.fields);
  if (options.format === 'json') return JSON.stringify(env, null, 2);
  if (options.format === 'csv') return csv(env.data, options.fields);
  const text = describe(env.data, options.fields);
  if (env.meta?.truncated) return `${text}\n\n${env.meta.count} shown; ${env.meta.hint}`;
  return env.meta ? `${text}\n\n${env.meta.count} result${env.meta.count === 1 ? '' : 's'}` : text;
}

export function renderError(error: CliError, format: OutputFormat): string {
  if (format === 'json') return JSON.stringify({ ok: false, error: error.toJSON() }, null, 2);
  const lines = [`Error [${error.code}]: ${error.message}`];
  if (error.hint) lines.push(`Hint: ${error.hint}`);
  if (error.details !== undefined) lines.push(describe(error.details));
  if (error.requestId) lines.push(`Graph request id: ${error.requestId}`);
  return lines.join('\n');
}

export function renderConfirmation(payload: Record<string, unknown>, format: OutputFormat): string {
  if (format === 'json') return JSON.stringify(payload, null, 2);
  const confirm = payload.confirm as { token: string; expiresAt: string; rerun: string };
  return [
    `Confirmation required for: ${payload.command as string}`,
    '',
    describe(payload.preview),
    '',
    `To proceed (valid until ${confirm.expiresAt}):`,
    `  ${confirm.rerun}`,
  ].join('\n');
}
