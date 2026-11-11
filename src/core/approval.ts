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
    writeSync(fd, `\n${summary}\n\nType APPROVE to make this change (anything else cancels): `);
    const buffer = Buffer.alloc(256);
    let answer = '';
    while (!answer.includes('\n')) {
      const read = readSync(fd, buffer, 0, buffer.length, null);
      if (read <= 0) break;
      answer += buffer.toString('utf8', 0, read);
    }
    const approved = answer.trim() === 'APPROVE';
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

/** Plain-text summary of a preview for a dialog or terminal prompt, bounded so dialogs stay readable. */
export function approvalSummary(commandId: string, profileLabel: string, preview: unknown): string {
  const lines = [`tenantctl ${commandId}`, `Tenant/profile: ${profileLabel}`];
  if (preview && typeof preview === 'object') {
    const action = 'action' in preview && typeof preview.action === 'string' ? preview.action : null;
    if (action) lines.push('', action);
    const detail = JSON.stringify(preview, null, 1).replace(/\n\s*/g, ' ');
    lines.push('', detail.length > 700 ? `${detail.slice(0, 699)}…` : detail);
  }
  return lines.join('\n');
}
