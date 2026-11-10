import { kebab } from './core/args.ts';
import type { AnyCommand } from './core/types.ts';
import { commands as defender } from './domains/defender/index.ts';
import { commands as entra } from './domains/entra/index.ts';
import { commands as exchange } from './domains/exchange/index.ts';
import { commands as graph } from './domains/graph.ts';
import { commands as health } from './domains/health/index.ts';
import { commands as purview } from './domains/purview/index.ts';
import { commands as intune } from './domains/intune/index.ts';
import { commands as sharepoint } from './domains/sharepoint/index.ts';
import { commands as system } from './domains/system.ts';
import { commands as teams } from './domains/teams/index.ts';
import { commands as troubleshoot } from './domains/troubleshoot/index.ts';

export interface DomainInfo {
  name: string;
  summary: string;
}

/** Display order for help and the tool catalog. */
export const DOMAINS: DomainInfo[] = [
  { name: 'troubleshoot', summary: 'One-shot diagnostics: a user, a device, sign-in failures, mail delivery' },
  {
    name: 'entra',
    summary:
      'Users, guests, groups, licenses, sign-ins, MFA, roles, apps and consent, Conditional Access, risk, reviews',
  },
  { name: 'intune', summary: 'Managed devices, compliance, configuration, apps, Autopilot, remote actions' },
  {
    name: 'defender',
    summary: 'Incidents, alerts, hunting, Secure Score, device response (isolate, scan, investigate)',
  },
  {
    name: 'exchange',
    summary: 'Mailboxes, rules, forwarding, permissions, shared mailboxes, DLs, quarantine, holds, trace, rooms',
  },
  { name: 'sharepoint', summary: 'Sites, libraries, files, sharing and permissions, OneDrive' },
  { name: 'purview', summary: 'Unified audit log search (who did what, across services), eDiscovery cases and holds' },
  { name: 'teams', summary: 'Teams, channels, membership, presence, call quality' },
  { name: 'health', summary: 'Microsoft 365 service health and message center' },
  { name: 'graph', summary: 'Raw Microsoft Graph requests (escape hatch)' },
  { name: 'auth', summary: 'Sign in, sign out, show the current identity and permissions' },
  { name: 'config', summary: 'Profiles, settings, local audit log' },
  { name: 'tools', summary: 'Machine-readable catalog of every command (for agents)' },
  { name: 'permissions', summary: 'Graph permissions the app registration needs' },
];

export const COMMANDS: AnyCommand[] = [
  ...troubleshoot,
  ...entra,
  ...intune,
  ...defender,
  ...exchange,
  ...sharepoint,
  ...purview,
  ...teams,
  ...health,
  ...graph,
  ...system,
];

export function commandId(command: AnyCommand): string {
  return command.name ? `${command.domain} ${command.name}` : command.domain;
}

const BY_ID = new Map(COMMANDS.map((command) => [commandId(command), command]));

/** Longest command id matching the leading words. */
export function findCommand(words: string[]): { command: AnyCommand; consumed: number } | undefined {
  for (let count = Math.min(words.length, 5); count >= 1; count--) {
    const command = BY_ID.get(words.slice(0, count).join(' '));
    if (command) return { command, consumed: count };
  }
  return undefined;
}

export function domainCommands(domain: string): AnyCommand[] {
  return COMMANDS.filter((command) => command.domain === domain);
}

/** Flag names the CLI owns; commands must not declare args with these names. */
export const GLOBAL_FLAGS = ['output', 'fields', 'full', 'profile', 'confirm', 'dryRun', 'debug', 'help', 'version'];

export function usage(command: AnyCommand): string {
  const parts = [`m365 ${commandId(command)}`];
  for (const arg of command.args.filter((a) => a.positional)) {
    parts.push(arg.required ? `<${kebab(arg.name)}>` : `[${kebab(arg.name)}]`);
  }
  for (const arg of command.args.filter((a) => !a.positional && a.required)) {
    parts.push(`--${kebab(arg.name)} <${arg.type === 'string[]' ? 'a,b' : arg.type}>`);
  }
  if (command.args.some((a) => !a.positional && !a.required)) parts.push('[options]');
  return parts.join(' ');
}
