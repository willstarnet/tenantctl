import { createHmac, randomBytes, timingSafeEqual } from 'node:crypto';
import { appendFileSync, chmodSync, existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { configDir, ensureConfigDir, writePrivateFile } from './config.ts';
import { CliError } from './errors.ts';
import type { ArgSpec } from './types.ts';

/** How long a confirmation token from a preview stays valid. */
export const CONFIRM_TTL_MS = 10 * 60 * 1000;

export interface ConfirmBinding {
  command: string;
  profile: string;
  tenant?: string;
  args: Record<string, unknown>;
}

function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`;
  if (value && typeof value === 'object') {
    const entries = Object.entries(value as Record<string, unknown>)
      .filter(([, entry]) => entry !== undefined)
      .sort(([a], [b]) => a.localeCompare(b));
    return `{${entries.map(([key, entry]) => `${JSON.stringify(key)}:${canonical(entry)}`).join(',')}}`;
  }
  return JSON.stringify(value);
}

function confirmKey(): Buffer {
  const path = join(configDir(), 'confirm.key');
  if (existsSync(path)) return Buffer.from(readFileSync(path, 'utf8').trim(), 'hex');
  const key = randomBytes(32);
  writePrivateFile(path, key.toString('hex'));
  return key;
}

function sign(expires: number, binding: ConfirmBinding): string {
  return createHmac('sha256', confirmKey())
    .update(`${expires}.${canonical(binding)}`)
    .digest('base64url')
    .slice(0, 22);
}

/**
 * A destructive command runs only with a token minted by its own preview, bound to the exact command, args,
 * profile and tenant. Changing any argument (a different device, say) invalidates it.
 */
export function issueConfirmToken(binding: ConfirmBinding, now = Date.now()): { token: string; expiresAt: string } {
  const expires = Math.floor((now + CONFIRM_TTL_MS) / 1000);
  return { token: `${expires}.${sign(expires, binding)}`, expiresAt: new Date(expires * 1000).toISOString() };
}

export function verifyConfirmToken(token: string, binding: ConfirmBinding, now = Date.now()): void {
  const [expiresRaw, signature] = token.split('.');
  const expires = Number(expiresRaw);
  if (!signature || !Number.isInteger(expires)) {
    throw new CliError('CONFIRM_INVALID', 'Malformed confirmation token', {
      hint: 'Run the command without --confirm to get a fresh preview and token.',
    });
  }
  if (expires * 1000 < now) {
    throw new CliError('CONFIRM_INVALID', 'Confirmation token expired', {
      hint: 'Run the command without --confirm to get a fresh preview and token.',
    });
  }
  const expected = Buffer.from(sign(expires, binding));
  const actual = Buffer.from(signature);
  if (expected.length !== actual.length || !timingSafeEqual(expected, actual)) {
    throw new CliError('CONFIRM_INVALID', 'Confirmation token does not match this command and arguments', {
      hint: 'Tokens are bound to the exact command, arguments and profile. Re-run without --confirm to preview.',
    });
  }
}

/** Args with secret values replaced, for audit logs and rerun hints. */
export function redactArgs(specs: readonly ArgSpec[], args: Record<string, unknown>): Record<string, unknown> {
  const out = { ...args };
  for (const spec of specs) if (spec.secret && out[spec.name] !== undefined) out[spec.name] = '[redacted]';
  return out;
}

export interface AuditEntry {
  command: string;
  profile: string;
  tenant?: string;
  actor?: string;
  risk: string;
  args: Record<string, unknown>;
  outcome: 'success' | 'error' | 'confirmation_issued' | 'denied';
  error?: string;
}

export function auditLogPath(): string {
  return join(configDir(), 'audit.jsonl');
}

/** Append-only local record of every write attempt. Failure to log never blocks the command. */
export function appendAudit(entry: AuditEntry): void {
  try {
    ensureConfigDir();
    const path = auditLogPath();
    appendFileSync(path, `${JSON.stringify({ time: new Date().toISOString(), ...entry })}\n`, { mode: 0o600 });
    if (process.platform !== 'win32') chmodSync(path, 0o600);
  } catch {
    // Read-only home directories (containers) still get to run commands.
  }
}
