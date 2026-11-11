#!/usr/bin/env bun
import pkg from '../package.json' with { type: 'json' };
import { kebab } from './core/args.ts';
import { resolveProfile, writesEnabled } from './core/config.ts';
import { CliError, EXIT, toCliError } from './core/errors.ts';
import { execute, type GlobalOptions } from './core/execute.ts';
import { renderConfirmation, renderError, renderSuccess, type OutputFormat } from './core/output.ts';
import type { AnyCommand } from './core/types.ts';
import { commandId, DOMAINS, domainCommands, findCommand, usage } from './registry.ts';

interface ParsedGlobals extends GlobalOptions {
  output: OutputFormat;
  fields?: string[];
  help: boolean;
  version: boolean;
  rest: string[];
}

function defaultFormat(): OutputFormat {
  const env = process.env.TENANTCTL_OUTPUT;
  if (env === 'json' || env === 'table' || env === 'csv') return env;
  return process.stdout.isTTY ? 'table' : 'json';
}

/** Pull CLI-owned flags out of argv wherever they appear; everything else belongs to the command. */
function parseGlobals(argv: string[]): ParsedGlobals {
  const out: ParsedGlobals = {
    output: defaultFormat(),
    full: false,
    debug: false,
    dryRun: false,
    help: false,
    version: false,
    rest: [],
  };
  const takeValue = (index: number, flag: string): string => {
    const value = argv[index + 1];
    if (value === undefined) throw new CliError('USAGE', `${flag} needs a value`);
    return value;
  };
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i] as string;
    if (token === '--') {
      out.rest.push(...argv.slice(i));
      break;
    }
    const [flag, inline] = token.startsWith('--') ? (token.split(/=(.*)/s, 2) as [string, string?]) : [token];
    const value = () => {
      if (inline !== undefined) return inline;
      const next = takeValue(i, flag);
      i++;
      return next;
    };
    switch (flag) {
      case '--output':
      case '-o': {
        const format = value();
        if (format !== 'json' && format !== 'table' && format !== 'csv') {
          throw new CliError('USAGE', '--output must be json, table or csv');
        }
        out.output = format;
        break;
      }
      case '--json':
        out.output = 'json';
        break;
      case '--csv':
        out.output = 'csv';
        break;
      case '--fields':
        out.fields = value()
          .split(',')
          .map((field) => field.trim())
          .filter(Boolean);
        break;
      case '--full':
        out.full = true;
        break;
      case '--profile':
        out.profile = value();
        break;
      case '--confirm':
        out.confirm = value();
        break;
      case '--dry-run':
        out.dryRun = true;
        break;
      case '--debug':
        out.debug = true;
        break;
      case '--help':
      case '-h':
        out.help = true;
        break;
      case '--version':
      case '-V':
        out.version = true;
        break;
      default:
        out.rest.push(token);
    }
  }
  return out;
}

function riskTag(command: AnyCommand): string {
  if (command.risk === 'destructive') return ' [destructive]';
  if (command.risk === 'write') return ' [write]';
  return '';
}

function rootHelp(): string {
  const width = Math.max(...DOMAINS.map((d) => d.name.length));
  return [
    `tenantctl ${pkg.version} - Microsoft 365 administration over Microsoft Graph`,
    '',
    'Usage: tenantctl <area> <command> [args] [options]',
    '',
    'Areas:',
    ...DOMAINS.map((d) => `  ${d.name.padEnd(width)}  ${d.summary}`),
    '',
    'Global options:',
    '  -o, --output json|table|csv  Output format (default: table on a terminal, json otherwise)',
    '  --json / --csv           Same as --output json / --output csv (csv: lists as a spreadsheet)',
    '  --fields a,b.c           Keep only these fields in the output',
    '  --full                   Return raw Graph objects instead of compact summaries',
    '  --profile <name>         Use a configured profile (tenant/credentials)',
    '  --dry-run                Preview a change without making it',
    '  --confirm <token>        confirmMode token only: execute a destructive change previewed earlier',
    '  --debug                  Log every Graph request to stderr',
    '  -h, --help               Help for any area or command',
    '',
    'Exit codes: 0 ok, 1 error, 2 usage, 3 confirmation required (token mode), 4 sign-in required, 5 permission denied,',
    '            6 not found, 7 writes disabled, 8 ambiguous name, 9 throttled, 10 not approved by a person',
    '',
    'Start with `tenantctl auth login`, then e.g. `tenantctl troubleshoot user jane@contoso.com`.',
    'Agents: `tenantctl tools` prints every command with its arguments as JSON.',
  ].join('\n');
}

function domainHelp(domain: string): string {
  const info = DOMAINS.find((d) => d.name === domain);
  const all = domainCommands(domain);
  const showWrites = writesEnabled(resolveProfile().profile);
  const shown = all.filter((c) => showWrites || c.risk === 'read');
  const hidden = all.length - shown.length;
  const width = Math.max(...shown.map((c) => commandId(c).length));
  const lines = [`tenantctl ${domain} - ${info?.summary ?? ''}`, '', 'Commands:'];
  for (const command of shown)
    lines.push(`  ${commandId(command).padEnd(width)}  ${command.summary}${riskTag(command)}`);
  if (hidden) {
    lines.push(
      '',
      `${hidden} write command${hidden === 1 ? '' : 's'} hidden: writes are disabled (TENANTCTL_ALLOW_WRITES=1).`,
    );
  }
  lines.push('', `Run \`tenantctl ${domain} <command> --help\` for details.`);
  return lines.join('\n');
}

function commandHelp(command: AnyCommand): string {
  const lines = [`${usage(command)}`, '', command.summary + riskTag(command)];
  if (command.description) lines.push('', command.description);
  if (command.args.length) {
    lines.push('', 'Arguments:');
    const label = (a: AnyCommand['args'][number]) => (a.positional ? `<${kebab(a.name)}>` : `--${kebab(a.name)}`);
    const width = Math.max(...command.args.map((a) => label(a).length));
    for (const arg of command.args) {
      const extras = [
        arg.required ? 'required' : '',
        arg.choices ? `one of: ${arg.choices.join('|')}` : '',
        arg.default !== undefined ? `default: ${String(arg.default)}` : '',
        arg.positional ? `or --${kebab(arg.name)}` : '',
      ].filter(Boolean);
      lines.push(`  ${label(arg).padEnd(width)}  ${arg.description}${extras.length ? ` (${extras.join('; ')})` : ''}`);
    }
  }
  if (command.examples?.length) lines.push('', 'Examples:', ...command.examples.map((e) => `  ${e}`));
  const { delegated, application, roles, exchange, defender } = command.permissions;
  if (delegated?.length || application?.length || exchange || defender?.length) {
    lines.push('', 'Permissions:');
    if (delegated?.length) lines.push(`  delegated:   ${delegated.join(', ')}`);
    if (application?.length) lines.push(`  application: ${application.join(', ')}`);
    if (exchange) lines.push('  exchange:    Exchange Online PowerShell (tenantctl auth login --service exchange)');
    if (defender?.length) lines.push(`  defender:    ${defender.join(', ')} (WindowsDefenderATP)`);
    if (roles?.length) lines.push(`  admin role:  ${roles.join(' or ')}`);
  } else if (command.domain === 'graph') {
    lines.push('', 'Permissions: whatever the called endpoint documents.');
  }
  if (command.risk === 'destructive') {
    lines.push(
      '',
      'Asks you to approve the preview (OS dialog or terminal) before making the change. With confirmMode token it runs in two steps instead: the first run prints a --confirm token; re-run with it to execute.',
    );
  }
  return lines.join('\n');
}

function shellQuote(token: string): string {
  return /^[\w@%+=:,./-]+$/.test(token) ? token : `'${token.replace(/'/g, `'\\''`)}'`;
}

/** The exact command to re-run with the token, with secret values masked. */
function rerunHint(argv: string[], command: AnyCommand, token: string): string {
  const secretFlags = new Set(command.args.filter((a) => a.secret).map((a) => `--${kebab(a.name)}`));
  const parts: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const current = argv[i] as string;
    if (current === '--confirm') {
      i++;
      continue;
    }
    if (current.startsWith('--confirm=')) continue;
    const flag = current.split('=')[0] as string;
    if (secretFlags.has(flag)) {
      parts.push(flag, '<same value>');
      if (!current.includes('=')) i++;
      continue;
    }
    parts.push(shellQuote(current));
  }
  return ['tenantctl', ...parts, '--confirm', token].join(' ');
}

async function main(argv: string[]): Promise<number> {
  let format = defaultFormat();
  try {
    const globals = parseGlobals(argv);
    format = globals.output;
    if (globals.version) {
      process.stdout.write(`${pkg.version}\n`);
      return EXIT.ok;
    }
    const words: string[] = [];
    for (const token of globals.rest) {
      if (token.startsWith('-')) break;
      words.push(token);
    }
    if (!words.length || words[0] === 'help') {
      const target = words.slice(words[0] === 'help' ? 1 : 0);
      const found = target.length ? findCommand(target) : undefined;
      if (found) process.stdout.write(`${commandHelp(found.command)}\n`);
      else if (target[0] && DOMAINS.some((d) => d.name === target[0]))
        process.stdout.write(`${domainHelp(target[0])}\n`);
      else process.stdout.write(`${rootHelp()}\n`);
      return words.length || globals.help ? EXIT.ok : EXIT.usage;
    }
    const found = findCommand(words);
    if (!found) {
      if (DOMAINS.some((d) => d.name === words[0])) {
        const unknown = words.length > 1;
        (unknown ? process.stderr : process.stdout).write(
          `${unknown ? `Unknown command: tenantctl ${words.join(' ')}\n\n` : ''}${domainHelp(words[0] as string)}\n`,
        );
        return unknown ? EXIT.usage : EXIT.ok;
      }
      throw new CliError('USAGE', `Unknown area "${words[0]}"`, {
        hint: `Areas: ${DOMAINS.map((d) => d.name).join(', ')}. Run \`tenantctl --help\`.`,
      });
    }
    const { command, consumed } = found;
    const id = commandId(command);
    if (globals.help) {
      process.stdout.write(`${commandHelp(command)}\n`);
      return EXIT.ok;
    }
    const { name: profileName, profile } = resolveProfile(globals.profile);
    const outcome = await execute(command, id, globals.rest.slice(consumed), globals, { profileName, profile });
    if (outcome.kind === 'confirm') {
      const payload = {
        ok: false,
        confirmationRequired: true,
        command: id,
        preview: outcome.preview,
        confirm: { token: outcome.token, expiresAt: outcome.expiresAt, rerun: rerunHint(argv, command, outcome.token) },
        instructions:
          'Nothing has changed yet. Show this preview to the user and get explicit approval before re-running ' +
          'with --confirm.',
      };
      process.stdout.write(`${renderConfirmation(payload, format)}\n`);
      return EXIT.confirmationRequired;
    }
    process.stdout.write(`${renderSuccess(id, outcome.result, { format, fields: globals.fields })}\n`);
    return EXIT.ok;
  } catch (error) {
    const cliError = toCliError(error);
    if (format === 'json') process.stdout.write(`${renderError(cliError, format)}\n`);
    else process.stderr.write(`${renderError(cliError, format)}\n`);
    if (cliError.code === 'INTERNAL' && process.env.TENANTCTL_DEBUG && error instanceof Error) {
      process.stderr.write(`${error.stack}\n`);
    }
    return cliError.exitCode;
  }
}

process.exitCode = await main(process.argv.slice(2));
