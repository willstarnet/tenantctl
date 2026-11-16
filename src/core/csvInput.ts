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

/**
 * CSV text in the encoding the file declares. Windows PowerShell 5.1 `Out-File` (used by Get-WindowsAutopilotInfo
 * and many export scripts) writes UTF-16LE with a BOM; everything else is read as UTF-8.
 */
function readText(path: string): string {
  const bytes = readFileSync(path);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return bytes.subarray(2).toString('utf16le');
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    // UTF-16BE: swap byte pairs into little-endian order.
    const swapped = Buffer.from(bytes.subarray(2));
    swapped.swap16();
    return swapped.toString('utf16le');
  }
  return bytes.toString('utf8');
}

export function readCsvRows(path: string, required: string[] = []): CsvRow[] {
  if (!existsSync(path)) throw new CliError('USAGE', `CSV file not found: ${path}`);
  const records = parseCsv(readText(path));
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
