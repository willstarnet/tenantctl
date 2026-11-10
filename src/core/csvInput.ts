import { existsSync, readFileSync } from 'node:fs';
import { CliError } from './errors.ts';
import { parseCsv } from './graph.ts';

/** Bulk commands refuse files beyond this, so a wrong file can't fan out into thousands of changes. */
export const MAX_CSV_ROWS = 500;

/** One CSV row with case- and spacing-insensitive column lookup (`Display Name` = `displayName`). */
export interface CsvRow {
  /** 1-based line number of the row in the file (header is line 1), for error messages. */
  line: number;
  get(...names: string[]): string | undefined;
}

const normalize = (name: string) => name.toLowerCase().replace(/[\s_-]/g, '');

export function readCsvRows(path: string, required: string[] = []): CsvRow[] {
  if (!existsSync(path)) throw new CliError('USAGE', `CSV file not found: ${path}`);
  const records = parseCsv(readFileSync(path, 'utf8'));
  if (!records.length) throw new CliError('USAGE', `${path} has no data rows`);
  if (records.length > MAX_CSV_ROWS) {
    throw new CliError('USAGE', `${path} has ${records.length} rows; the limit is ${MAX_CSV_ROWS} per run`, {
      hint: 'Split the file and run each part separately.',
    });
  }
  const headers = Object.keys(records[0] as Record<string, string>);
  const byName: Record<string, string> = {};
  for (const header of headers) byName[normalize(header)] = header;
  const missing = required.filter((name) => !byName[normalize(name)]);
  if (missing.length) {
    throw new CliError('USAGE', `${path} is missing column${missing.length > 1 ? 's' : ''}: ${missing.join(', ')}`, {
      hint: `Columns found: ${headers.join(', ')}`,
    });
  }
  return records.map((record, index) => ({
    line: index + 2,
    get(...names: string[]) {
      for (const name of names) {
        const header = byName[normalize(name)];
        const value = header ? record[header]?.trim() : undefined;
        if (value) return value;
      }
      return undefined;
    },
  }));
}
