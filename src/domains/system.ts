import { existsSync, readFileSync } from 'node:fs';
import { effectiveConfirmMode, weakens } from '../core/approval.ts';
import { kebab } from '../core/args.ts';
import { Authenticator, decodeClaims, grantedPermissions, type Resource } from '../core/auth.ts';
import {
  configDir,
  DEFAULT_PUBLIC_CLIENT_ID,
  loadConfig,
  type Profile,
  readSecrets,
  saveConfig,
  updateStoredProfile,
  writesEnabled,
} from '../core/config.ts';
import { CliError } from '../core/errors.ts';
import { auditLogPath } from '../core/safety.ts';
import { cacheKey } from '../core/secretStore.ts';
import { type AnyCommand, type ArgSpec, type Context, defineCommand, type ListResult } from '../core/types.ts';
import { COMMANDS, commandId } from '../registry.ts';

/** Directory role template ids (the `wids` claim) for the roles an M365 admin typically holds. */
const ROLE_NAMES: Record<string, string> = {
  '62e90394-69f5-4237-9190-012177145e10': 'Global Administrator',
  'f2ef992c-3afb-46b9-b7cf-a126ee74c451': 'Global Reader',
  'fe930be7-5e62-47db-91af-98c3a49a38b1': 'User Administrator',
  '729827e3-9c14-49f7-bb1b-9608f156bbb8': 'Helpdesk Administrator',
  'c4e39bd9-1100-46d3-8c65-fb160da0071f': 'Authentication Administrator',
  '7be44c8a-adaf-4e2a-84d6-ab2649e08a13': 'Privileged Authentication Administrator',
  'e8611ab8-c189-46e8-94e1-60213ab1f814': 'Privileged Role Administrator',
  'fdd7a751-b60b-444a-984c-02652fe8fa1c': 'Groups Administrator',
  '4d6ac14f-3453-41d0-bef9-a3e0c569773a': 'License Administrator',
  'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9': 'Conditional Access Administrator',
  '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3': 'Application Administrator',
  '158c047a-c907-4556-b7ef-446551a6b5f7': 'Cloud Application Administrator',
  '3a2c62db-5318-420d-8d74-23affee5d9d5': 'Intune Administrator',
  '7698a772-787b-4ac8-901f-60d6b08affd2': 'Cloud Device Administrator',
  '194ae4cb-b126-40b2-bd5b-6091b380977d': 'Security Administrator',
  '5f2222b1-57c3-48ba-8ad5-d4759f1fde6f': 'Security Operator',
  '5d6b6bb7-de71-4623-b4af-96380a352509': 'Security Reader',
  '29232cdf-9323-42fd-ade2-1d097af3e4de': 'Exchange Administrator',
  'f28a1f50-f6e7-4571-818b-6a12f2af6b6c': 'SharePoint Administrator',
  '69091246-20e8-4a56-aa4d-066075b2a7a8': 'Teams Administrator',
  '4a5d8f65-41da-4de4-8968-e035b65339cf': 'Reports Reader',
  'f023fd81-a637-4b56-95fd-791ac0226033': 'Service Support Administrator',
};

/** Areas with no fixed permission set: local config, and the raw Graph escape hatch. */
const NON_TENANT_DOMAINS: Record<string, true> = {
  auth: true,
  config: true,
  tools: true,
  permissions: true,
  graph: true,
};

/** Commands that call Microsoft 365 with a declared permission set. */
function tenantCommands(includeWrites: boolean, domain?: string): AnyCommand[] {
  return COMMANDS.filter(
    (c) => !NON_TENANT_DOMAINS[c.domain] && (!domain || c.domain === domain) && (includeWrites || c.risk === 'read'),
  );
}

function unionPermissions(commands: AnyCommand[], kind: 'delegated' | 'application'): string[] {
  return [...new Set(commands.flatMap((c) => c.permissions[kind] ?? []))].sort();
}

function authenticator(ctx: Context): Authenticator {
  return new Authenticator(ctx.profileName, ctx.profile, readSecrets());
}

async function status(ctx: Context) {
  const auth = authenticator(ctx);
  const claims = decodeClaims(await auth.getToken());
  return {
    profile: ctx.profileName,
    source: auth.source,
    tenantId: claims.tid ?? ctx.profile.tenantId,
    identity:
      auth.source === 'app'
        ? `${claims.app_displayname ?? 'app'} (${claims.appid ?? ctx.profile.clientId})`
        : (claims.upn ?? claims.preferred_username ?? claims.unique_name ?? ctx.profile.account?.username),
    adminRoles: (claims.wids ?? []).map((id) => ROLE_NAMES[id]).filter((name): name is string => Boolean(name)),
    permissions: grantedPermissions(claims).sort(),
    expiresAt: claims.exp ? new Date(claims.exp * 1000).toISOString() : null,
    writesEnabled: writesEnabled(ctx.profile),
    destructiveApproval: effectiveConfirmMode(ctx.profile),
    tokenCacheProtection: cacheKey().store,
    exchangeSignedInAs: ctx.profile.exchangeAccount?.username ?? null,
  };
}

const PROFILE_KEYS = [
  'authMode',
  'tenantId',
  'clientId',
  'certificatePath',
  'certificateThumbprint',
  'allowWrites',
  'confirmMode',
  'graphUrl',
  'authorityHost',
] as const;
type ProfileKey = (typeof PROFILE_KEYS)[number];

function parseProfileValue(key: ProfileKey, raw: string): Profile[ProfileKey] {
  if (key === 'allowWrites') {
    if (!/^(true|false)$/i.test(raw)) throw new CliError('USAGE', 'allowWrites must be true or false');
    return raw.toLowerCase() === 'true';
  }
  if (key === 'authMode' && raw !== 'delegated' && raw !== 'app') {
    throw new CliError('USAGE', 'authMode must be delegated or app');
  }
  if (key === 'confirmMode' && !['token', 'tty', 'dialog', 'auto'].includes(raw)) {
    throw new CliError('USAGE', 'confirmMode must be token, tty, dialog or auto');
  }
  return raw as Profile[ProfileKey];
}

/**
 * Settings that lower human oversight need the human: weakening the approval method, and turning writes on
 * when a human-approval mode is in effect. Otherwise an agent could switch off the very check meant for it.
 */
async function guardSetting(ctx: Context, key: ProfileKey, next: Profile[ProfileKey] | undefined): Promise<void> {
  if (key === 'confirmMode' && weakens(ctx.profile, next as Profile['confirmMode'])) {
    await ctx.requireApproval(
      `m365 config set confirmMode ${String(next ?? 'auto')}\nProfile: ${ctx.profileName}\n\n` +
        `Lowers approval of destructive changes from "${effectiveConfirmMode(ctx.profile)}" to "${String(next)}".`,
    );
  }
  if (
    key === 'allowWrites' &&
    next === true &&
    effectiveConfirmMode(ctx.profile) !== 'token' &&
    !ctx.profile.allowWrites
  ) {
    await ctx.requireApproval(
      `m365 config set allowWrites true\nProfile: ${ctx.profileName}\n\nLets this CLI change the tenant.`,
    );
  }
}

function jsonSchemaFor(args: readonly ArgSpec[]) {
  const properties: Record<string, unknown> = {};
  for (const arg of args) {
    const type =
      arg.type === 'string[]'
        ? { type: 'array', items: { type: 'string' } }
        : { type: arg.type === 'number' ? 'number' : arg.type };
    properties[arg.name] = {
      ...type,
      description: arg.description,
      ...(arg.choices ? { enum: arg.choices } : {}),
      ...(arg.default !== undefined ? { default: arg.default } : {}),
    };
  }
  return { type: 'object', properties, required: args.filter((a) => a.required).map((a) => a.name) };
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'auth',
    name: 'login',
    summary: 'Sign in (device code for delegated profiles; verifies credentials for app-only profiles)',
    description:
      'Delegated: prints a code and URL on stderr; the user opens the URL, enters the code and signs in with an ' +
      'admin account. The command waits until they finish. Relay the stderr message to the user verbatim. ' +
      'By default it requests every read permission the CLI uses (plus write permissions when writes are ' +
      'enabled or --include-writes is passed). Without a profile it uses the Microsoft Graph Command Line Tools ' +
      'app, so no app registration is needed.',
    examples: ['m365 auth login', 'm365 auth login --tenant contoso.onmicrosoft.com --include-writes'],
    args: [
      {
        name: 'service',
        type: 'string',
        choices: ['graph', 'exchange', 'defender'],
        default: 'graph',
        description:
          'What to sign in to: graph (everything except the items below), exchange (Exchange admin commands that ' +
          'run through PowerShell), defender (Defender for Endpoint device actions).',
      },
      {
        name: 'scopes',
        type: 'string[]',
        description: 'Graph only: request exactly these delegated scopes instead of the default set.',
      },
      {
        name: 'includeWrites',
        type: 'boolean',
        description: 'Graph only: also request the scopes write commands need.',
      },
      {
        name: 'tenant',
        type: 'string',
        description: "Tenant id or domain to sign in to (default: the account's home tenant).",
      },
    ],
    risk: 'read',
    permissions: {},
    async run(ctx, args) {
      const profile: Profile = { ...ctx.profile, ...(args.tenant ? { tenantId: args.tenant } : {}) };
      const auth = new Authenticator(ctx.profileName, profile, readSecrets());
      const service = args.service as Resource;
      if (service === 'graph' && auth.source === 'static-token') {
        throw new CliError('USAGE', 'M365_ACCESS_TOKEN is set; unset it to sign in with a profile');
      }
      const scopes =
        args.scopes ?? unionPermissions(tenantCommands(args.includeWrites || writesEnabled(profile)), 'delegated');
      const result = await auth.login(service, scopes, (message) => process.stderr.write(`\n${message}\n\n`));
      if (profile.authMode === 'app') {
        const claims = decodeClaims(result.accessToken);
        return service === 'graph'
          ? { verified: true, ...(await status({ ...ctx, profile })) }
          : { verified: true, service, roles: claims.roles ?? [] };
      }
      const account = result.account!;
      const signedIn = { homeAccountId: account.homeAccountId, username: account.username, tenantId: account.tenantId };
      if (service === 'exchange') {
        updateStoredProfile(ctx.profileName, { exchangeAccount: signedIn });
        return { profile: ctx.profileName, service, signedInAs: account.username, tenantId: account.tenantId };
      }
      if (service === 'defender') {
        return { profile: ctx.profileName, service, signedInAs: account.username, scopes: result.scopes };
      }
      const granted = result.scopes.filter(
        (s) => !['openid', 'profile', 'offline_access', 'email'].includes(s.toLowerCase()),
      );
      updateStoredProfile(ctx.profileName, {
        authMode: 'delegated',
        tenantId: args.tenant ?? ctx.profile.tenantId,
        account: signedIn,
        scopes: granted,
      });
      const missing = scopes.filter((s) => !granted.some((g) => g.toLowerCase() === s.toLowerCase()));
      return {
        profile: ctx.profileName,
        signedInAs: account.username,
        tenantId: account.tenantId,
        grantedScopes: granted.length,
        ...(missing.length
          ? { notGranted: missing, note: 'Commands needing these scopes will fail until an admin consents.' }
          : {}),
      };
    },
  }),
  defineCommand({
    domain: 'auth',
    name: 'status',
    summary: 'Show who is signed in, the tenant, admin roles, granted permissions and whether writes are enabled',
    description: 'Run this first when a command fails with AUTH_* or FORBIDDEN, to see what the token actually grants.',
    args: [],
    risk: 'read',
    permissions: {},
    run: (ctx) => status(ctx),
  }),
  defineCommand({
    domain: 'auth',
    name: 'logout',
    summary: 'Delete the cached tokens for the profile',
    args: [],
    risk: 'read',
    permissions: {},
    async run(ctx) {
      const removed = authenticator(ctx).logout();
      if (ctx.profile.account || ctx.profile.exchangeAccount) {
        updateStoredProfile(ctx.profileName, { account: undefined, exchangeAccount: undefined, scopes: undefined });
      }
      return {
        profile: ctx.profileName,
        signedOut: removed || Boolean(ctx.profile.account || ctx.profile.exchangeAccount),
      };
    },
  }),
  defineCommand({
    domain: 'config',
    name: 'init',
    summary: 'Create or update a profile (tenant, app registration, auth mode)',
    description:
      'Delegated profiles need nothing else; `m365 auth login` works without this. App-only profiles need a ' +
      'tenant, the app registration client id, and either --cert (PEM with private key and certificate) or the ' +
      'M365_CLIENT_SECRET environment variable. Secrets are never written to the config file.',
    examples: [
      'm365 config init --mode app --tenant contoso.onmicrosoft.com --client-id 00000000-... --cert ~/certs/m365.pem',
      'm365 --profile fabrikam config init --tenant fabrikam.onmicrosoft.com',
    ],
    args: [
      { name: 'mode', type: 'string', choices: ['delegated', 'app'], description: 'Auth mode (default: delegated).' },
      { name: 'tenant', type: 'string', description: 'Tenant id or primary domain.' },
      { name: 'clientId', type: 'string', description: 'App registration (client) id.' },
      {
        name: 'cert',
        type: 'string',
        description: 'App-only: path to a PEM containing the private key and certificate.',
      },
      {
        name: 'thumbprint',
        type: 'string',
        description: 'App-only: SHA-256 thumbprint, if the PEM has no certificate.',
      },
      { name: 'default', type: 'boolean', description: 'Make this the default profile.' },
    ],
    risk: 'read',
    permissions: {},
    async run(ctx, args) {
      const mode = args.mode ?? loadConfig().profiles[ctx.profileName]?.authMode ?? 'delegated';
      const profile = updateStoredProfile(ctx.profileName, {
        authMode: mode,
        ...(args.tenant ? { tenantId: args.tenant } : {}),
        ...(args.clientId ? { clientId: args.clientId } : {}),
        ...(args.cert ? { certificatePath: args.cert } : {}),
        ...(args.thumbprint ? { certificateThumbprint: args.thumbprint } : {}),
      });
      if (args.default) {
        const config = loadConfig();
        config.defaultProfile = ctx.profileName;
        saveConfig(config);
      }
      const next =
        mode === 'app'
          ? 'Run `m365 auth login` to verify the credentials, then `m365 permissions --mode app` for what to grant.'
          : 'Run `m365 auth login`.';
      return {
        profile: ctx.profileName,
        saved: { ...profile, clientId: profile.clientId ?? `${DEFAULT_PUBLIC_CLIENT_ID} (default)` },
        next,
      };
    },
  }),
  defineCommand({
    domain: 'config',
    name: 'show',
    summary: 'Show the active profile with environment overrides applied (never shows secrets)',
    args: [],
    risk: 'read',
    permissions: {},
    async run(ctx) {
      const secrets = readSecrets();
      const config = loadConfig();
      return {
        profile: ctx.profileName,
        configDir: configDir(),
        defaultProfile: config.defaultProfile ?? null,
        profiles: Object.keys(config.profiles),
        settings: { ...ctx.profile, scopes: ctx.profile.scopes ? `${ctx.profile.scopes.length} scopes` : undefined },
        writesEnabled: writesEnabled(ctx.profile),
        env: {
          clientSecretSet: Boolean(secrets.clientSecret),
          accessTokenSet: Boolean(secrets.accessToken),
          overrides: Object.keys(process.env).filter(
            (key) => key.startsWith('M365_') && !/SECRET|TOKEN|PASSPHRASE/.test(key),
          ),
        },
      };
    },
  }),
  defineCommand({
    domain: 'config',
    name: 'set',
    summary: 'Set one profile setting',
    examples: ['m365 config set allowWrites true', 'm365 config set tenantId contoso.onmicrosoft.com'],
    args: [
      {
        name: 'key',
        type: 'string',
        required: true,
        positional: true,
        choices: PROFILE_KEYS,
        description: 'Setting name.',
      },
      { name: 'value', type: 'string', required: true, positional: true, description: 'New value.' },
    ],
    risk: 'read',
    permissions: {},
    async run(ctx, args) {
      const value = parseProfileValue(args.key, args.value);
      await guardSetting(ctx, args.key, value);
      updateStoredProfile(ctx.profileName, { [args.key]: value });
      return { profile: ctx.profileName, [args.key]: value };
    },
  }),
  defineCommand({
    domain: 'config',
    name: 'unset',
    summary: 'Remove one profile setting',
    args: [
      {
        name: 'key',
        type: 'string',
        required: true,
        positional: true,
        choices: PROFILE_KEYS,
        description: 'Setting name.',
      },
    ],
    risk: 'read',
    permissions: {},
    async run(ctx, args) {
      await guardSetting(ctx, args.key, undefined);
      updateStoredProfile(ctx.profileName, { [args.key]: undefined });
      return { profile: ctx.profileName, unset: args.key };
    },
  }),
  defineCommand({
    domain: 'config',
    name: 'use',
    summary: 'Make a profile the default',
    args: [{ name: 'name', type: 'string', required: true, positional: true, description: 'Profile name.' }],
    risk: 'read',
    permissions: {},
    async run(_ctx, args) {
      const config = loadConfig();
      if (!config.profiles[args.name]) {
        throw new CliError('NOT_FOUND', `No profile "${args.name}"`, {
          hint: `Profiles: ${Object.keys(config.profiles).join(', ') || 'none'}`,
        });
      }
      config.defaultProfile = args.name;
      saveConfig(config);
      return { defaultProfile: args.name };
    },
  }),
  defineCommand({
    domain: 'config',
    name: 'audit-log',
    summary: 'Show the local log of write attempts made with this CLI (newest first)',
    args: [{ name: 'top', type: 'number', default: 20, description: 'Entries to show.' }],
    risk: 'read',
    permissions: {},
    async run(_ctx, args): Promise<ListResult> {
      const path = auditLogPath();
      if (!existsSync(path)) return { items: [], truncated: false };
      const lines = readFileSync(path, 'utf8').trim().split('\n').filter(Boolean).reverse();
      return {
        items: lines.slice(0, args.top).map((line) => JSON.parse(line) as unknown),
        truncated: lines.length > args.top,
      };
    },
  }),
  defineCommand({
    domain: 'tools',
    name: '',
    summary: 'Machine-readable catalog of commands: arguments (JSON Schema), risk, permissions',
    description:
      'Write commands are omitted unless writes are enabled or --include-writes is passed. Each entry has the ' +
      'exact argv prefix in `command`; pass args as --kebab-case flags.',
    args: [
      { name: 'domain', type: 'string', positional: true, description: 'Only this area.' },
      { name: 'includeWrites', type: 'boolean', description: 'Include write commands even when writes are disabled.' },
    ],
    risk: 'read',
    permissions: {},
    async run(ctx, args) {
      const include = args.includeWrites || writesEnabled(ctx.profile);
      const selected = COMMANDS.filter(
        (c) => (!args.domain || c.domain === args.domain) && (include || c.risk === 'read'),
      );
      return {
        usage: 'm365 <command> [--flag value ...] --output json',
        writesEnabled: writesEnabled(ctx.profile),
        confirmFlow:
          'Destructive commands exit 3 with a preview and confirm.rerun; show the preview to the user and only re-run with --confirm after they approve.',
        commands: selected.map((c) => ({
          command: commandId(c),
          summary: c.summary,
          ...(c.description ? { description: c.description } : {}),
          risk: c.risk,
          flags: c.args.map(
            (a) => `${a.positional ? `<${kebab(a.name)}> | ` : ''}--${kebab(a.name)}${a.required ? ' (required)' : ''}`,
          ),
          inputSchema: jsonSchemaFor(c.args),
          ...(c.examples?.length ? { examples: c.examples } : {}),
          permissions: c.permissions,
        })),
      };
    },
  }),
  defineCommand({
    domain: 'permissions',
    name: '',
    summary: 'List the Graph permissions to grant the app registration (or consent to) for this CLI',
    examples: ['m365 permissions --mode app', 'm365 permissions --mode delegated --include-writes --by-command'],
    args: [
      {
        name: 'mode',
        type: 'string',
        choices: ['delegated', 'application'],
        default: 'application',
        description: 'Permission type.',
      },
      { name: 'includeWrites', type: 'boolean', description: 'Include permissions for write commands.' },
      { name: 'domain', type: 'string', description: 'Only this area.' },
      { name: 'byCommand', type: 'boolean', description: 'Show which commands need each permission.' },
    ],
    risk: 'read',
    permissions: {},
    async run(_ctx, args) {
      const selected = tenantCommands(args.includeWrites, args.domain);
      const kind = args.mode === 'delegated' ? 'delegated' : 'application';
      const permissions = unionPermissions(selected, kind);
      const defender = [...new Set(selected.flatMap((c) => c.permissions.defender ?? []))].sort();
      const usesExchange = selected.some((c) => c.permissions.exchange);
      const otherApis = {
        ...(usesExchange
          ? {
              exchangeOnline:
                kind === 'application'
                  ? 'Office 365 Exchange Online: Exchange.ManageAsApp, plus the Exchange Administrator role on the app'
                  : 'Exchange admin role; sign in with `m365 auth login --service exchange`',
            }
          : {}),
        ...(defender.length ? { defenderForEndpoint: { api: 'WindowsDefenderATP', permissions: defender } } : {}),
      };
      if (!args.byCommand) {
        return { mode: kind, includeWrites: args.includeWrites, count: permissions.length, permissions, ...otherApis };
      }
      return {
        mode: kind,
        permissions: permissions.map((permission) => ({
          permission,
          commands: selected.filter((c) => c.permissions[kind]?.includes(permission)).map(commandId),
        })),
        ...otherApis,
      };
    },
  }),
];
