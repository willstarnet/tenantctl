import { parseArgs } from './args.ts';
import { approvalSummary, approverFor, type Approver, effectiveConfirmMode } from './approval.ts';
import { Authenticator, decodeClaims } from './auth.ts';
import { DEFAULT_GRAPH_URL, type Profile, readSecrets, writesEnabled } from './config.ts';
import { CliError } from './errors.ts';
import { createExoRunner, type ExoRunner } from './exo.ts';
import { GraphClient } from './graph.ts';
import { appendAudit, issueConfirmToken, redactArgs, verifyConfirmToken, type ConfirmBinding } from './safety.ts';
import type { AnyCommand, Context } from './types.ts';

const DEFENDER_API = 'https://api.security.microsoft.com/api';

export interface GlobalOptions {
  profile?: string;
  full: boolean;
  debug: boolean;
  dryRun: boolean;
  confirm?: string;
}

export interface RunEnvironment {
  profileName: string;
  profile: Profile;
  /** Overridable for tests. */
  fetch?: typeof fetch;
  log?: (message: string) => void;
  /** Human approval; defaults to the profile's confirm mode. Tests inject a stub. */
  approver?: Approver;
  /** Exchange PowerShell runner; tests inject a fake. */
  exo?: ExoRunner;
}

export type Outcome =
  | { kind: 'result'; result: unknown }
  | { kind: 'confirm'; preview: unknown; token: string; expiresAt: string };

/** App-only Exchange connections need the tenant's *.onmicrosoft.com domain; delegated ones the admin's UPN. */
async function exoConnection(auth: Authenticator, graph: GraphClient, profile: Profile) {
  const token = await auth.getToken('exchange');
  if (profile.authMode === 'app') {
    const org = await graph.get<{ value?: { verifiedDomains?: { name: string; isInitial?: boolean }[] }[] }>(
      '/organization',
      {
        query: { $select: 'verifiedDomains' },
      },
    );
    const initial = org.value?.[0]?.verifiedDomains?.find((d) => d.isInitial)?.name;
    if (!initial) throw new CliError('CONFIG', "Could not find the tenant's .onmicrosoft.com domain for Exchange");
    return { token, organization: initial };
  }
  const claims = decodeClaims(token);
  return { token, userPrincipalName: claims.upn ?? claims.preferred_username ?? profile.exchangeAccount?.username };
}

export function createContext(env: RunEnvironment, options: Pick<GlobalOptions, 'full' | 'debug'>): Context {
  const log = env.log ?? ((message: string) => process.stderr.write(`${message}\n`));
  const auth = new Authenticator(env.profileName, env.profile, readSecrets());
  const debug = options.debug ? log : undefined;
  const graph = new GraphClient({
    baseUrl: env.profile.graphUrl ?? DEFAULT_GRAPH_URL,
    getToken: () => auth.getToken('graph'),
    fetch: env.fetch,
    debug,
  });
  const defender = new GraphClient({
    baseUrl: DEFENDER_API,
    getToken: () => auth.getToken('defender'),
    fetch: env.fetch,
    debug,
    versioned: false,
  });
  const exo = env.exo ?? createExoRunner(() => exoConnection(auth, graph, env.profile), debug);
  const approver = env.approver ?? approverFor(effectiveConfirmMode(env.profile));
  return {
    graph,
    defender,
    exo,
    profileName: env.profileName,
    profile: env.profile,
    full: options.full,
    log,
    async requireApproval(summary) {
      if (!(await approver(summary))) {
        throw new CliError('DENIED', 'The change was not approved', {
          hint: 'Nothing was changed. Ask the user whether they still want it before trying again.',
        });
      }
    },
  };
}

/** Commands list every permission they need together, so the hint names all of them, not alternatives. */
function permissionHint(command: AnyCommand, profile: Profile): string {
  const { delegated, application, roles } = command.permissions;
  const parts: string[] = [];
  if (profile.authMode === 'app' && application?.length) {
    parts.push(`The app registration needs application permission ${application.join(' and ')} with admin consent`);
  } else if (delegated?.length) {
    // Re-running the default login re-requests the full scope set; `--scopes` would replace it.
    const login = command.risk === 'read' ? 'tenantctl auth login' : 'tenantctl auth login --include-writes';
    parts.push(`Needs delegated scope ${delegated.join(' and ')}: run \`${login}\` (an admin may need to consent)`);
  }
  if (command.permissions.exchange) {
    parts.push(
      profile.authMode === 'app'
        ? 'For the Exchange part: Office 365 Exchange Online Exchange.ManageAsApp and an Exchange role on the app'
        : 'For the Exchange part: an Exchange admin role, signed in with `tenantctl auth login --service exchange`',
    );
  }
  if (command.permissions.defender?.length) {
    parts.push(`Defender for Endpoint (WindowsDefenderATP) permission ${command.permissions.defender.join(' and ')}`);
  }
  if (roles?.length && profile.authMode !== 'app') parts.push(`and an admin role such as ${roles.join(' or ')}`);
  return parts.length ? `${parts.join('. ')}.` : 'The signed-in identity lacks permission for this operation.';
}

/** A bare Graph 403 gains the command's permission needs; a hint the command wrote itself is more specific. */
function enrich(error: unknown, command: AnyCommand, profile: Profile): unknown {
  if (error instanceof CliError && error.code === 'FORBIDDEN' && !error.hint) {
    return new CliError('FORBIDDEN', error.message, {
      status: error.status,
      graphCode: error.graphCode,
      requestId: error.requestId,
      hint: permissionHint(command, profile),
    });
  }
  return error;
}

/**
 * Parse, gate and run one command.
 *
 * Reads run directly. Writes need writes enabled for the profile. Destructive commands also need a person: in
 * `dialog` and `tty` confirm modes the person approves the current preview on a channel the agent cannot answer,
 * in one run. In `token` mode (no person available) they run in two steps: without `--confirm` they return a
 * preview and a token bound to the exact arguments; with a valid token they execute. `--dry-run` previews any
 * write without needing writes enabled.
 */
export async function execute(
  command: AnyCommand,
  commandId: string,
  tokens: string[],
  options: GlobalOptions,
  env: RunEnvironment,
): Promise<Outcome> {
  const args = parseArgs(command.args, tokens);
  const ctx = createContext(env, options);
  const preview = async () => (command.preview ? command.preview(ctx, args) : { command: commandId, args });

  try {
    if (command.risk === 'read') return { kind: 'result', result: await command.run(ctx, args) };

    if (options.dryRun) {
      return {
        kind: 'result',
        result: { dryRun: true, command: commandId, risk: command.risk, preview: await preview() },
      };
    }
    if (!writesEnabled(env.profile)) {
      throw new CliError('WRITES_DISABLED', `"${commandId}" changes the tenant and writes are disabled`, {
        hint:
          'Writes are off by default. Ask the user before enabling them: export TENANTCTL_ALLOW_WRITES=1 for this ' +
          'session, or `tenantctl config set allowWrites true` for the profile. Use --dry-run to preview meanwhile.',
      });
    }

    const audit = {
      command: commandId,
      profile: env.profileName,
      tenant: env.profile.tenantId ?? env.profile.account?.tenantId,
      actor: env.profile.authMode === 'app' ? `app:${env.profile.clientId}` : env.profile.account?.username,
      risk: command.risk,
      args: redactArgs(command.args, args),
    };

    if (command.risk === 'destructive') {
      if (effectiveConfirmMode(env.profile) === 'token') {
        const binding: ConfirmBinding = { command: commandId, profile: env.profileName, tenant: audit.tenant, args };
        if (!options.confirm) {
          const shown = await preview();
          const { token, expiresAt } = issueConfirmToken(binding);
          appendAudit({ ...audit, outcome: 'confirmation_issued' });
          return { kind: 'confirm', preview: shown, token, expiresAt };
        }
        verifyConfirmToken(options.confirm, binding);
      }
      const label = `${env.profileName}${audit.tenant ? ` (${audit.tenant})` : ''}`;
      const summary = approvalSummary(commandId, label, await preview());
      try {
        await ctx.requireApproval(summary);
      } catch (error) {
        if (error instanceof CliError && error.code === 'DENIED') appendAudit({ ...audit, outcome: 'denied' });
        throw error;
      }
    }

    try {
      const result = await command.run(ctx, args);
      appendAudit({ ...audit, outcome: 'success' });
      return { kind: 'result', result };
    } catch (error) {
      appendAudit({ ...audit, outcome: 'error', error: error instanceof Error ? error.message : String(error) });
      throw error;
    }
  } catch (error) {
    throw enrich(error, command, env.profile);
  }
}
