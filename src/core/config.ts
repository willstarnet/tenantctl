import { chmodSync, existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { CliError } from './errors.ts';

export type AuthMode = 'delegated' | 'app';

/**
 * Microsoft Graph Command Line Tools: Microsoft's first-party public client (the one Microsoft Graph PowerShell
 * uses). Lets delegated sign-in work without registering an app; a tenant may block it, in which case the profile
 * sets its own `clientId`.
 */
export const DEFAULT_PUBLIC_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
export const DEFAULT_GRAPH_URL = 'https://graph.microsoft.com';
export const DEFAULT_AUTHORITY_HOST = 'https://login.microsoftonline.com';

export interface SignedInAccount {
  homeAccountId: string;
  username: string;
  tenantId: string;
}

export interface Profile {
  authMode: AuthMode;
  tenantId?: string;
  clientId?: string;
  /** App-only: PEM file holding the private key (and ideally the certificate, so the thumbprint is derived). */
  certificatePath?: string;
  /** App-only: SHA-256 thumbprint, hex. Only needed when `certificatePath` has no certificate block. */
  certificateThumbprint?: string;
  allowWrites?: boolean;
  /** Delegated: scopes consented at the last `auth login`. */
  scopes?: string[];
  /** Delegated: the account from the last `auth login`. */
  account?: SignedInAccount;
  /** Delegated: the account from the last `auth login --service exchange` (a separate Exchange sign-in). */
  exchangeAccount?: SignedInAccount;
  /**
   * How a person approves destructive changes: `tty` (typed at the terminal) or `dialog` (an OS dialog) in the
   * same run; `token` (no person: the agent re-runs with --confirm); `auto` (dialog when a desktop session exists,
   * else tty). Default auto.
   */
  confirmMode?: ConfirmMode;
  graphUrl?: string;
  authorityHost?: string;
}

export type ConfirmMode = 'token' | 'tty' | 'dialog' | 'auto';

export interface ConfigFile {
  defaultProfile?: string;
  profiles: Record<string, Profile>;
}

/** Settings that may come from the environment and must never be written to disk. */
export interface Secrets {
  clientSecret?: string;
  certificatePassphrase?: string;
  /** A ready bearer token; bypasses MSAL entirely (CI, tests, `az account get-access-token`). */
  accessToken?: string;
  /** Ready tokens for the other resources (Exchange Online, Defender for Endpoint), same purpose. */
  exchangeAccessToken?: string;
  defenderAccessToken?: string;
}

let migrated = false;

export function configDir(): string {
  const env = process.env;
  if (env.TENANTCTL_CONFIG_DIR) return env.TENANTCTL_CONFIG_DIR;
  const base =
    process.platform === 'win32' && env.APPDATA ? env.APPDATA : env.XDG_CONFIG_HOME || join(homedir(), '.config');
  const dir = join(base, 'tenantctl');
  // One-time move from the folder the CLI used when it was called `m365`, so profiles and sign-ins carry over.
  if (!migrated) {
    migrated = true;
    const legacy = join(base, 'm365');
    if (!existsSync(dir) && existsSync(join(legacy, 'config.json'))) renameSync(legacy, dir);
  }
  return dir;
}

export function ensureConfigDir(): string {
  const dir = configDir();
  if (!existsSync(dir)) mkdirSync(dir, { recursive: true, mode: 0o700 });
  return dir;
}

/** Write a file only the current user can read, atomically. */
export function writePrivateFile(path: string, content: string): void {
  ensureConfigDir();
  const tmp = `${path}.${process.pid}.tmp`;
  writeFileSync(tmp, content, { mode: 0o600 });
  renameSync(tmp, path);
  if (process.platform !== 'win32') chmodSync(path, 0o600);
}

export function loadConfig(): ConfigFile {
  const path = join(configDir(), 'config.json');
  if (!existsSync(path)) return { profiles: {} };
  try {
    const parsed = JSON.parse(readFileSync(path, 'utf8')) as Partial<ConfigFile>;
    return { defaultProfile: parsed.defaultProfile, profiles: parsed.profiles ?? {} };
  } catch (error) {
    throw new CliError('CONFIG', `Cannot parse ${path}: ${(error as Error).message}`, {
      hint: 'Fix or delete the file, then run `tenantctl config init`.',
    });
  }
}

export function saveConfig(config: ConfigFile): void {
  writePrivateFile(join(configDir(), 'config.json'), `${JSON.stringify(config, null, 2)}\n`);
}

export function activeProfileName(explicit?: string): string {
  return explicit || process.env.TENANTCTL_PROFILE || loadConfig().defaultProfile || 'default';
}

function envBool(value: string | undefined): boolean | undefined {
  if (value === undefined || value === '') return undefined;
  if (/^(1|true|yes|on)$/i.test(value)) return true;
  if (/^(0|false|no|off)$/i.test(value)) return false;
  throw new CliError('CONFIG', `Expected a boolean, got "${value}"`);
}

/** The stored profile with environment overrides applied. Missing profile => delegated with defaults. */
export function resolveProfile(explicit?: string): { name: string; profile: Profile } {
  const name = activeProfileName(explicit);
  const stored = loadConfig().profiles[name];
  const env = process.env;
  const mode = env.TENANTCTL_AUTH_MODE as AuthMode | undefined;
  if (mode && mode !== 'delegated' && mode !== 'app') {
    throw new CliError('CONFIG', `TENANTCTL_AUTH_MODE must be "delegated" or "app", got "${mode}"`);
  }
  const profile: Profile = {
    ...(stored ?? { authMode: 'delegated' }),
    ...(mode ? { authMode: mode } : {}),
    ...(env.TENANTCTL_TENANT_ID ? { tenantId: env.TENANTCTL_TENANT_ID } : {}),
    ...(env.TENANTCTL_CLIENT_ID ? { clientId: env.TENANTCTL_CLIENT_ID } : {}),
    ...(env.TENANTCTL_CERT_PATH ? { certificatePath: env.TENANTCTL_CERT_PATH } : {}),
    ...(env.TENANTCTL_CERT_THUMBPRINT ? { certificateThumbprint: env.TENANTCTL_CERT_THUMBPRINT } : {}),
    ...(env.TENANTCTL_GRAPH_URL ? { graphUrl: env.TENANTCTL_GRAPH_URL } : {}),
    ...(env.TENANTCTL_AUTHORITY_HOST ? { authorityHost: env.TENANTCTL_AUTHORITY_HOST } : {}),
  };
  const allowWrites = envBool(env.TENANTCTL_ALLOW_WRITES);
  if (allowWrites !== undefined) profile.allowWrites = allowWrites;
  return { name, profile };
}

export function readSecrets(): Secrets {
  const env = process.env;
  return {
    clientSecret: env.TENANTCTL_CLIENT_SECRET || undefined,
    certificatePassphrase: env.TENANTCTL_CERT_PASSPHRASE || undefined,
    accessToken: env.TENANTCTL_ACCESS_TOKEN || undefined,
    exchangeAccessToken: env.TENANTCTL_EXCHANGE_ACCESS_TOKEN || undefined,
    defenderAccessToken: env.TENANTCTL_DEFENDER_ACCESS_TOKEN || undefined,
  };
}

export function writesEnabled(profile: Profile): boolean {
  return profile.allowWrites === true;
}

/** Persist changes to a stored profile, leaving environment overrides out of the file. */
export function updateStoredProfile(name: string, patch: Partial<Profile>): Profile {
  const config = loadConfig();
  const current = config.profiles[name] ?? { authMode: 'delegated' };
  const next = { ...current, ...patch };
  for (const key of Object.keys(next) as (keyof Profile)[]) {
    if (next[key] === undefined) delete next[key];
  }
  config.profiles[name] = next;
  config.defaultProfile ??= name;
  saveConfig(config);
  return next;
}
