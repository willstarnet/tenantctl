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
   * How a person approves destructive changes: `token` (the agent re-runs with --confirm), `tty` (typed at the
   * terminal), `dialog` (an OS dialog), `auto` (dialog when a desktop session exists, else tty). Default auto.
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

export function configDir(): string {
  const env = process.env;
  if (env.M365_CONFIG_DIR) return env.M365_CONFIG_DIR;
  if (process.platform === 'win32' && env.APPDATA) return join(env.APPDATA, 'm365');
  return join(env.XDG_CONFIG_HOME || join(homedir(), '.config'), 'm365');
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
      hint: 'Fix or delete the file, then run `m365 config init`.',
    });
  }
}

export function saveConfig(config: ConfigFile): void {
  writePrivateFile(join(configDir(), 'config.json'), `${JSON.stringify(config, null, 2)}\n`);
}

export function activeProfileName(explicit?: string): string {
  return explicit || process.env.M365_PROFILE || loadConfig().defaultProfile || 'default';
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
  const mode = env.M365_AUTH_MODE as AuthMode | undefined;
  if (mode && mode !== 'delegated' && mode !== 'app') {
    throw new CliError('CONFIG', `M365_AUTH_MODE must be "delegated" or "app", got "${mode}"`);
  }
  const profile: Profile = {
    ...(stored ?? { authMode: 'delegated' }),
    ...(mode ? { authMode: mode } : {}),
    ...(env.M365_TENANT_ID ? { tenantId: env.M365_TENANT_ID } : {}),
    ...(env.M365_CLIENT_ID ? { clientId: env.M365_CLIENT_ID } : {}),
    ...(env.M365_CERT_PATH ? { certificatePath: env.M365_CERT_PATH } : {}),
    ...(env.M365_CERT_THUMBPRINT ? { certificateThumbprint: env.M365_CERT_THUMBPRINT } : {}),
    ...(env.M365_GRAPH_URL ? { graphUrl: env.M365_GRAPH_URL } : {}),
    ...(env.M365_AUTHORITY_HOST ? { authorityHost: env.M365_AUTHORITY_HOST } : {}),
  };
  const allowWrites = envBool(env.M365_ALLOW_WRITES);
  if (allowWrites !== undefined) profile.allowWrites = allowWrites;
  return { name, profile };
}

export function readSecrets(): Secrets {
  const env = process.env;
  return {
    clientSecret: env.M365_CLIENT_SECRET || undefined,
    certificatePassphrase: env.M365_CERT_PASSPHRASE || undefined,
    accessToken: env.M365_ACCESS_TOKEN || undefined,
    exchangeAccessToken: env.M365_EXCHANGE_ACCESS_TOKEN || undefined,
    defenderAccessToken: env.M365_DEFENDER_ACCESS_TOKEN || undefined,
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
