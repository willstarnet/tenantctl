import { CliError, type ErrorCode } from './errors.ts';

/**
 * Exchange Online admin features that Graph lacks (mailbox permissions, shared mailboxes, mailbox forwarding,
 * distribution lists, quarantine, holds) run through the ExchangeOnlineManagement PowerShell module, connected
 * with an access token this CLI already holds (`Connect-ExchangeOnline -AccessToken`, module 3.1+).
 *
 * The script text is fixed. Cmdlet names come from code; the token, organization and every parameter value
 * travel in environment variables as JSON and are splatted, so user input never becomes PowerShell source.
 */
export interface ExoStep {
  /** Verb-Noun cmdlet name; always a constant in the calling command. */
  cmdlet: string;
  params?: Record<string, unknown>;
  /** Properties to return (strongly recommended: Exchange objects carry hundreds). */
  select?: string[];
  /** Keep running later steps if this one fails (each failure is reported on its own result). */
  continueOnError?: boolean;
}

export interface ExoStepResult {
  ok: boolean;
  data: Record<string, unknown>[];
  error?: string;
}

export interface ExoRunner {
  run(steps: ExoStep[]): Promise<ExoStepResult[]>;
}

export interface ExoConnection {
  token: string;
  /** App-only: the tenant's *.onmicrosoft.com domain. */
  organization?: string;
  /** Delegated: the signed-in admin's UPN. */
  userPrincipalName?: string;
}

const CMDLET = /^[A-Z][A-Za-z]+-[A-Z][A-Za-z]+$/;
const BEGIN = '<<<M365-EXO-JSON>>>';
const END = '<<<M365-EXO-END>>>';
const TIMEOUT_MS = 240_000;

const SCRIPT = `
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
function Out-Result($value) { [Console]::Out.Write('${BEGIN}' + (ConvertTo-Json -InputObject $value -Depth 8 -Compress) + '${END}') }
function ConvertTo-Plain($object, $properties) {
  if ($null -eq $properties -or $properties.Count -eq 0) { $properties = @($object.PSObject.Properties.Name) }
  $plain = [ordered]@{}
  foreach ($name in $properties) {
    $value = $object.$name
    if ($null -eq $value -or $value -is [string] -or $value -is [bool] -or $value -is [int] -or $value -is [long] -or $value -is [double]) { $plain[$name] = $value }
    elseif ($value -is [datetime]) { $plain[$name] = $value.ToUniversalTime().ToString('o') }
    elseif ($value -is [System.Collections.IEnumerable]) { $plain[$name] = @($value | ForEach-Object { "$_" }) }
    else { $plain[$name] = "$value" }
  }
  $plain
}
try { Import-Module ExchangeOnlineManagement -MinimumVersion 3.1.0 -ErrorAction Stop }
catch { Out-Result @{ ok = $false; stage = 'module'; message = $_.Exception.Message }; exit 0 }
$steps = @($env:M365_EXO_STEPS | ConvertFrom-Json -AsHashtable)
$connect = @{ AccessToken = $env:M365_EXO_TOKEN; ShowBanner = $false; CommandName = @($steps | ForEach-Object { $_.cmdlet } | Select-Object -Unique) }
if ($env:M365_EXO_ORG) { $connect.Organization = $env:M365_EXO_ORG } else { $connect.UserPrincipalName = $env:M365_EXO_UPN }
try { Connect-ExchangeOnline @connect | Out-Null }
catch { Out-Result @{ ok = $false; stage = 'connect'; message = $_.Exception.Message }; exit 0 }
$results = [System.Collections.ArrayList]::new()
foreach ($step in $steps) {
  $params = if ($step.params) { $step.params } else { @{} }
  try {
    $output = & $step.cmdlet @params
    $data = @($output | Where-Object { $null -ne $_ } | ForEach-Object { ConvertTo-Plain $_ $step.select })
    [void]$results.Add(@{ ok = $true; data = $data })
  } catch {
    [void]$results.Add(@{ ok = $false; data = @(); error = $_.Exception.Message })
    if (-not $step.continueOnError) { break }
  }
}
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
Out-Result @{ ok = $true; results = @($results) }
`;

/** pwsh -EncodedCommand takes base64 of UTF-16LE. */
const ENCODED_SCRIPT = Buffer.from(SCRIPT, 'utf16le').toString('base64');

interface ScriptOutput {
  ok: boolean;
  stage?: 'module' | 'connect';
  message?: string;
  results?: ExoStepResult[];
}

function pwshPath(): string {
  const found = process.env.M365_PWSH || Bun.which('pwsh');
  if (!found) {
    throw new CliError('CONFIG', 'Exchange admin commands need PowerShell 7 (pwsh)', {
      hint:
        'Install PowerShell 7 (macOS: `brew install powershell`), then ' +
        '`pwsh -c "Install-Module ExchangeOnlineManagement -Scope CurrentUser"`.',
    });
  }
  return found;
}

/** Map an Exchange error message to the CLI's error codes so agents branch the same way as for Graph. */
export function exoErrorCode(message: string): ErrorCode {
  if (/couldn't be found|couldn't find|isn't found|not found|doesn't exist/i.test(message)) return 'NOT_FOUND';
  if (/access denied|isn't within your current|not authorized|permission|role assignment/i.test(message))
    return 'FORBIDDEN';
  if (/throttl|too many requests|micro delay/i.test(message)) return 'THROTTLED';
  return 'GRAPH_ERROR';
}

export function createExoRunner(connect: () => Promise<ExoConnection>, debug?: (message: string) => void): ExoRunner {
  return {
    async run(steps) {
      for (const step of steps) {
        if (!CMDLET.test(step.cmdlet)) throw new CliError('INTERNAL', `Invalid cmdlet name "${step.cmdlet}"`);
      }
      const pwsh = pwshPath();
      const connection = await connect();
      const started = Date.now();
      // `-Command -` silently drops multi-line function definitions; the constant script goes as -EncodedCommand.
      const child = Bun.spawn([pwsh, '-NoProfile', '-NonInteractive', '-EncodedCommand', ENCODED_SCRIPT], {
        stdin: 'ignore',
        stdout: 'pipe',
        stderr: 'pipe',
        env: {
          ...process.env,
          M365_EXO_TOKEN: connection.token,
          M365_EXO_ORG: connection.organization ?? '',
          M365_EXO_UPN: connection.userPrincipalName ?? '',
          M365_EXO_STEPS: JSON.stringify(steps),
        },
      });
      const timer = setTimeout(() => child.kill(), TIMEOUT_MS);
      const [stdout, stderr] = await Promise.all([
        new Response(child.stdout).text(),
        new Response(child.stderr).text(),
      ]);
      await child.exited;
      clearTimeout(timer);
      debug?.(`pwsh ${steps.map((s) => s.cmdlet).join(', ')} -> exit ${child.exitCode} (${Date.now() - started} ms)`);
      const begin = stdout.lastIndexOf(BEGIN);
      const end = stdout.lastIndexOf(END);
      if (begin < 0 || end < begin) {
        throw new CliError(
          'GRAPH_ERROR',
          `Exchange PowerShell produced no result${child.signalCode ? ' (timed out)' : ''}`,
          {
            hint: (stderr || stdout).trim().split('\n').slice(-3).join(' ') || undefined,
          },
        );
      }
      const output = JSON.parse(stdout.slice(begin + BEGIN.length, end)) as ScriptOutput;
      if (!output.ok) {
        const message = output.message ?? 'unknown error';
        if (output.stage === 'module') {
          throw new CliError('CONFIG', `ExchangeOnlineManagement module unavailable: ${message}`, {
            hint: 'Run `pwsh -c "Install-Module ExchangeOnlineManagement -Scope CurrentUser"` (version 3.1 or later).',
          });
        }
        throw new CliError(
          exoErrorCode(message) === 'FORBIDDEN' ? 'FORBIDDEN' : 'AUTH_FAILED',
          `Connecting to Exchange Online failed: ${message}`,
          {
            hint:
              'Delegated: run `m365 auth login --service exchange` with an Exchange admin account. App-only: the app ' +
              'needs the Exchange.ManageAsApp application permission (Office 365 Exchange Online) and an Exchange ' +
              'Administrator role assignment.',
          },
        );
      }
      return output.results ?? [];
    },
  };
}

/** Run steps and fail on the first failed step, as most commands want. */
export async function exoRun(runner: ExoRunner, steps: ExoStep[]): Promise<Record<string, unknown>[][]> {
  const results = await runner.run(steps);
  return steps.map((step, index) => {
    const result = results[index];
    if (!result) throw new CliError('GRAPH_ERROR', `${step.cmdlet} did not run (an earlier step failed)`);
    if (!result.ok) throw new CliError(exoErrorCode(result.error ?? ''), `${step.cmdlet}: ${result.error ?? 'failed'}`);
    return result.data;
  });
}
