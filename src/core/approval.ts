import { closeSync, openSync, readSync, writeSync } from 'node:fs';
import type { ConfirmMode, Profile } from './config.ts';
import { CliError } from './errors.ts';

/**
 * Destructive changes need a person, not just the agent that previewed them. A confirm token alone is visible to
 * the agent, which could replay it; `tty` and `dialog` modes ask the human on a channel the agent's shell cannot
 * answer (the controlling terminal, an OS dialog). This raises the bar against confused or prompt-injected
 * agents; it is not a sandbox against a deliberately malicious process running as the same OS user, which could
 * edit the config file directly.
 */
export type ConcreteMode = Exclude<ConfirmMode, 'auto'>;
export type Approver = (summary: string) => Promise<boolean>;

const STRENGTH: Record<ConcreteMode, number> = { token: 0, tty: 1, dialog: 2 };
const DIALOG_TITLE = 'tenantctl: approve change';
/** A pending approval is abandoned after this; the agent can re-run to ask again. */
const DIALOG_TIMEOUT_SECONDS = 300;

function hasDesktop(): boolean {
  const env = process.env;
  if (env.SSH_CONNECTION || env.SSH_TTY) return false;
  if (process.platform === 'darwin' || process.platform === 'win32') return true;
  return Boolean((env.DISPLAY || env.WAYLAND_DISPLAY) && (Bun.which('zenity') || Bun.which('kdialog')));
}

function concrete(mode: ConfirmMode | undefined): ConcreteMode {
  if (mode === 'token' || mode === 'tty' || mode === 'dialog') return mode;
  return hasDesktop() ? 'dialog' : 'tty';
}

/**
 * The profile setting decides; the environment may only make it stricter, so an agent exporting
 * `TENANTCTL_CONFIRM_MODE=token` cannot switch human approval off.
 */
export function effectiveConfirmMode(profile: Profile): ConcreteMode {
  const configured = concrete(profile.confirmMode);
  const fromEnv = process.env.TENANTCTL_CONFIRM_MODE;
  if (fromEnv === 'tty' || fromEnv === 'dialog') return STRENGTH[fromEnv] > STRENGTH[configured] ? fromEnv : configured;
  return configured;
}

/** Whether moving to `next` lowers protection below what is in effect now (needs a human to approve). */
export function weakens(profile: Profile, next: ConfirmMode | undefined): boolean {
  return STRENGTH[concrete(next)] < STRENGTH[effectiveConfirmMode(profile)];
}

function dialogApprove(summary: string): boolean {
  const text = `${summary}\n\nApprove only if you asked for this change.`;
  if (process.platform === 'darwin') {
    const osascript = Bun.which('osascript');
    if (!osascript) throw unavailable('osascript not found');
    // The message travels as an argument, never inside the AppleScript source, so it cannot inject script.
    const result = Bun.spawnSync(
      [
        osascript,
        '-e',
        'on run argv',
        '-e',
        `display dialog (item 1 of argv) with title "${DIALOG_TITLE}" buttons {"Deny", "Approve"} default button "Deny" cancel button "Deny" with icon caution giving up after ${DIALOG_TIMEOUT_SECONDS}`,
        '-e',
        'end run',
        text,
      ],
      { stdout: 'pipe', stderr: 'pipe' },
    );
    const out = result.stdout.toString();
    return result.exitCode === 0 && out.includes('button returned:Approve') && !out.includes('gave up:true');
  }
  if (process.platform === 'win32') {
    const shell = Bun.which('pwsh') ?? Bun.which('powershell');
    if (!shell) throw unavailable('PowerShell not found');
    const result = Bun.spawnSync(
      [
        shell,
        '-NoProfile',
        '-Command',
        `Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show($env:TENANTCTL_APPROVAL_TEXT, '${DIALOG_TITLE}', 'YesNo', 'Warning', 'No').ToString()`,
      ],
      { stdout: 'pipe', stderr: 'pipe', env: { ...process.env, TENANTCTL_APPROVAL_TEXT: text } },
    );
    return result.exitCode === 0 && result.stdout.toString().trim() === 'Yes';
  }
  const zenity = Bun.which('zenity');
  if (zenity) {
    const result = Bun.spawnSync(
      [
        zenity,
        '--question',
        '--no-markup',
        `--title=${DIALOG_TITLE}`,
        `--text=${text}`,
        '--ok-label=Approve',
        '--cancel-label=Deny',
        '--default-cancel',
        `--timeout=${DIALOG_TIMEOUT_SECONDS}`,
      ],
      { stdout: 'ignore', stderr: 'ignore' },
    );
    return result.exitCode === 0;
  }
  const kdialog = Bun.which('kdialog');
  if (!kdialog) throw unavailable('no zenity or kdialog for a desktop dialog');
  const result = Bun.spawnSync(
    [kdialog, '--title', DIALOG_TITLE, '--warningyesno', text, '--yes-label', 'Approve', '--no-label', 'Deny'],
    {
      stdout: 'ignore',
      stderr: 'ignore',
    },
  );
  return result.exitCode === 0;
}

/** Ask on the controlling terminal; the agent's piped stdin/stdout never reach it. */
function ttyApprove(summary: string): boolean {
  let fd: number;
  try {
    fd = openSync(process.platform === 'win32' ? 'CON' : '/dev/tty', 'r+');
  } catch {
    throw unavailable('no controlling terminal');
  }
  try {
    writeSync(fd, `\n${DIALOG_TITLE}\n\n${summary}\n\nApprove this change? [y/N] `);
    const buffer = Buffer.alloc(256);
    let answer = '';
    while (!answer.includes('\n')) {
      const read = readSync(fd, buffer, 0, buffer.length, null);
      if (read <= 0) break;
      answer += buffer.toString('utf8', 0, read);
    }
    const approved = /^y(es)?$/i.test(answer.trim());
    writeSync(fd, approved ? 'Approved.\n' : 'Cancelled.\n');
    return approved;
  } finally {
    closeSync(fd);
  }
}

function unavailable(reason: string): CliError {
  return new CliError('DENIED', `Cannot ask for human approval: ${reason}`, {
    hint:
      'Run the command yourself in a terminal (tty mode needs one), or pick another approval method in your ' +
      'own terminal: `tenantctl config set confirmMode tty|dialog|token`.',
  });
}

/** The approver for a mode; `token` approves nothing extra (the confirm token was already verified). */
export function approverFor(mode: ConcreteMode): Approver {
  if (mode === 'token') return async () => true;
  if (mode === 'tty') return async (summary) => ttyApprove(summary);
  return async (summary) => dialogApprove(summary);
}

/** Bounds that keep a dialog on screen: detail lines, characters per value, items per list. */
const DETAIL_LINES = 30;
const VALUE_CHARS = 160;
const LIST_ITEMS = 3;
const INDENT = '  ';
/** Shown in the header, or (targetId) an internal binding already visible in the target's own fields. */
const HEADER_KEYS: Record<string, true> = { action: true, warnings: true, targetId: true };
const ACRONYMS: Record<string, string> = {
  api: 'API',
  id: 'ID',
  ids: 'IDs',
  imei: 'IMEI',
  ip: 'IP',
  mfa: 'MFA',
  os: 'OS',
  sku: 'SKU',
  smtp: 'SMTP',
  upn: 'UPN',
  url: 'URL',
};
const ISO_DATE = /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})(?::\d{2}(?:\.\d+)?)?(Z|[+-]\d{2}:\d{2})$/;

type Scalar = string | number | boolean;
type Fields = Record<string, unknown>;
type Request = { method: string; path: string; version?: unknown; body?: unknown };

const isFields = (value: unknown): value is Fields =>
  typeof value === 'object' && value !== null && !Array.isArray(value);
const isScalar = (value: unknown): value is Scalar =>
  typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';
const isRequest = (value: unknown): value is Request =>
  isFields(value) && typeof value.method === 'string' && typeof value.path === 'string';

function isEmpty(value: unknown): boolean {
  if (value === null || value === undefined || value === '') return true;
  if (Array.isArray(value)) return value.every(isEmpty);
  if (isFields(value)) return Object.values(value).every(isEmpty);
  return false;
}

/** `lastSyncDateTime` → `Last sync`, `userPrincipalName` → `User principal name`, `osVersion` → `OS version`. */
function label(key: string): string {
  const words = (key.replace(/DateTime$/, '') || key)
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .split(/[\s_-]+/)
    .filter(Boolean)
    .map((word) => word.toLowerCase())
    .map((word) => (Object.hasOwn(ACRONYMS, word) ? ACRONYMS[word] : word));
  const first = words[0] ?? key;
  words[0] = first.charAt(0).toUpperCase() + first.slice(1);
  return words.join(' ');
}

function scalar(value: Scalar): string {
  if (typeof value === 'boolean') return value ? 'yes' : 'no';
  if (typeof value === 'number') return String(value);
  const date = ISO_DATE.exec(value);
  if (date) return `${date[1]} ${date[2]} ${date[3] === 'Z' ? 'UTC' : date[3]}`;
  const flat = value.replace(/\s+/g, ' ').trim();
  return flat.length > VALUE_CHARS ? `${flat.slice(0, VALUE_CHARS - 1)}…` : flat;
}

function renderFields(fields: Fields, depth: number, out: string[]): void {
  for (const [key, value] of Object.entries(fields)) renderField(label(key), value, depth, out);
}

function renderField(name: string, value: unknown, depth: number, out: string[]): void {
  if (isEmpty(value)) return;
  const pad = INDENT.repeat(depth);
  if (isRequest(value)) {
    const version = isScalar(value.version) ? ` (${scalar(value.version)})` : '';
    out.push(`${pad}${name}: ${value.method} ${value.path}${version}`);
    renderField('Body', value.body, depth + 1, out);
  } else if (Array.isArray(value)) {
    renderList(
      name,
      value.filter((item) => !isEmpty(item)),
      depth,
      out,
    );
  } else if (isFields(value)) {
    out.push(`${pad}${name}:`);
    renderFields(value, depth + 1, out);
  } else {
    out.push(`${pad}${name}: ${isScalar(value) ? scalar(value) : String(value)}`);
  }
}

/** Short scalar lists stay on one line; anything else becomes bullets, capped at LIST_ITEMS. */
function renderList(name: string, items: unknown[], depth: number, out: string[]): void {
  const pad = INDENT.repeat(depth);
  if (items.every(isScalar)) {
    const inline = items.map(scalar).join(', ');
    if (inline.length <= VALUE_CHARS) {
      out.push(`${pad}${name}: ${inline}`);
      return;
    }
  }
  out.push(`${pad}${name} (${items.length}):`);
  for (const item of items.slice(0, LIST_ITEMS)) {
    const lines: string[] = [];
    if (isFields(item) && !isRequest(item)) renderFields(item, depth + 2, lines);
    else renderField('', item, depth + 2, lines);
    // The first line of each item carries the bullet; `- ` is as wide as INDENT, so later lines stay aligned.
    const [first = '', ...rest] = lines;
    out.push(`${pad}${INDENT}- ${first.trimStart().replace(/^: /, '')}`, ...rest);
  }
  if (items.length > LIST_ITEMS) out.push(`${pad}${INDENT}… and ${items.length - LIST_ITEMS} more`);
}

/**
 * Readable plain-text summary of a preview for a dialog or terminal prompt. The action, target tenant and every
 * warning are shown in full first; the remaining detail is laid out as labelled, indented fields and bounded so
 * dialogs stay on screen.
 */
export function approvalSummary(commandId: string, profileLabel: string, preview: unknown): string {
  const fields = isFields(preview) ? preview : {};
  const lines: string[] = [];
  if (typeof fields.action === 'string') lines.push(fields.action, '');
  lines.push(`Tenant: ${profileLabel}`, `Command: tenantctl ${commandId}`);

  const warnings = Array.isArray(fields.warnings)
    ? fields.warnings.filter((w): w is string => typeof w === 'string')
    : [];
  if (warnings.length) lines.push('', ...warnings.map((w) => `Warning: ${w}`));

  const detail: string[] = [];
  if (isFields(preview)) {
    renderFields(
      Object.fromEntries(Object.entries(preview).filter(([key]) => !Object.hasOwn(HEADER_KEYS, key))),
      0,
      detail,
    );
  } else {
    renderField('Details', preview, 0, detail);
  }
  if (detail.length > DETAIL_LINES) {
    const hidden = detail.length - (DETAIL_LINES - 1);
    detail.splice(DETAIL_LINES - 1, Infinity, `… ${hidden} more lines; add --dry-run to see the full preview`);
  }
  if (detail.length) lines.push('', ...detail);
  return lines.join('\n');
}
