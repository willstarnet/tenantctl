import {
  type AccountInfo,
  type AuthenticationResult,
  ConfidentialClientApplication,
  type Configuration,
  type ICachePlugin,
  InteractionRequiredAuthError,
  LogLevel,
  PublicClientApplication,
} from '@azure/msal-node';
import { createPrivateKey, X509Certificate } from 'node:crypto';
import { existsSync, readFileSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import {
  configDir,
  DEFAULT_AUTHORITY_HOST,
  DEFAULT_GRAPH_URL,
  DEFAULT_PUBLIC_CLIENT_ID,
  type Profile,
  type Secrets,
  writePrivateFile,
} from './config.ts';
import { CliError } from './errors.ts';
import { decryptText, encryptText } from './secretStore.ts';

export interface TokenClaims {
  tid?: string;
  oid?: string;
  upn?: string;
  unique_name?: string;
  preferred_username?: string;
  appid?: string;
  app_displayname?: string;
  scp?: string;
  roles?: string[];
  wids?: string[];
  exp?: number;
}

/** Decode (not verify) a JWT payload; we only read our own token to report who/what it grants. */
export function decodeClaims(token: string): TokenClaims {
  const payload = token.split('.')[1];
  if (!payload) return {};
  try {
    return JSON.parse(Buffer.from(payload, 'base64url').toString('utf8')) as TokenClaims;
  } catch {
    return {};
  }
}

/** Granted Graph permissions in a token: delegated `scp` or application `roles`. */
export function grantedPermissions(claims: TokenClaims): string[] {
  if (claims.scp) return claims.scp.split(' ').filter(Boolean);
  return claims.roles ?? [];
}

/** MSAL cache persisted encrypted; an older plaintext cache is re-saved encrypted on first read. */
function fileCachePlugin(path: string): ICachePlugin {
  return {
    async beforeCacheAccess(context) {
      if (!existsSync(path)) return;
      const decrypted = decryptText(readFileSync(path, 'utf8'));
      if (!decrypted) return;
      context.tokenCache.deserialize(decrypted.plain);
      if (decrypted.legacy) writePrivateFile(path, encryptText(decrypted.plain));
    },
    async afterCacheAccess(context) {
      if (context.cacheHasChanged) writePrivateFile(path, encryptText(context.tokenCache.serialize()));
    },
  };
}

function loadCertificate(profile: Profile, secrets: Secrets): { thumbprintSha256: string; privateKey: string } {
  const path = profile.certificatePath as string;
  if (!existsSync(path)) throw new CliError('CONFIG', `Certificate file not found: ${path}`);
  const pem = readFileSync(path, 'utf8');
  let privateKey: string;
  try {
    privateKey = createPrivateKey({ key: pem, format: 'pem', passphrase: secrets.certificatePassphrase })
      .export({ format: 'pem', type: 'pkcs8' })
      .toString();
  } catch (error) {
    throw new CliError('CONFIG', `Cannot read the private key in ${path}: ${(error as Error).message}`, {
      hint: 'The PEM must contain the private key. Set M365_CERT_PASSPHRASE if it is encrypted.',
    });
  }
  let thumbprint = profile.certificateThumbprint;
  const certBlock = pem.match(/-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/)?.[0];
  if (!thumbprint && certBlock) thumbprint = new X509Certificate(certBlock).fingerprint256;
  if (!thumbprint) {
    throw new CliError('CONFIG', `No certificate in ${path} to derive the thumbprint from`, {
      hint: 'Append the certificate PEM to the key file, or set certificateThumbprint (SHA-256).',
    });
  }
  return { thumbprintSha256: thumbprint.replace(/[^0-9a-f]/gi, '').toLowerCase(), privateKey };
}

function msalError(error: unknown, loginHint: string): CliError {
  if (error instanceof CliError) return error;
  if (error instanceof InteractionRequiredAuthError) {
    return new CliError('AUTH_REQUIRED', `Sign-in required: ${error.errorMessage || error.message}`, {
      hint: `Run \`${loginHint}\`${loginHint === 'm365 auth login' ? ' (add `--include-writes` if write commands need scopes you have not consented to)' : ''}.`,
    });
  }
  const message = error instanceof Error ? error.message : String(error);
  const aadsts = message.match(/AADSTS\d+/)?.[0];
  return new CliError('AUTH_FAILED', message.split('\n')[0] ?? message, {
    graphCode: aadsts,
    hint: aadsts ? `Look up ${aadsts} with \`m365 entra signin-error ${aadsts.slice(6)}\`.` : undefined,
  });
}

/** Resources the CLI calls. Each has its own token audience; Exchange also needs its own public client. */
export type Resource = 'graph' | 'exchange' | 'defender';

/** First-party public client of the Exchange Online PowerShell module, pre-authorized for Exchange Online. */
export const EXO_PUBLIC_CLIENT_ID = 'fb78d390-0c51-40cd-8e17-fdbfab77341b';
export const EXCHANGE_RESOURCE = 'https://outlook.office365.com';
/** Defender for Endpoint still issues tokens for the legacy resource even though calls go to api.security. */
export const DEFENDER_RESOURCE = 'https://api.securitycenter.microsoft.com';

interface ResourceSpec {
  appScope: string;
  clientId: string;
  delegatedScopes: string[];
  cacheSuffix: string;
  account: 'account' | 'exchangeAccount';
  loginHint: string;
}

/**
 * Acquires tokens for one profile. Per resource, three sources in precedence order: a static token from the
 * environment, app-only client credentials, delegated device code.
 */
export class Authenticator {
  readonly profileName: string;
  readonly profile: Profile;
  private readonly secrets: Secrets;
  private readonly tokens: Partial<Record<Resource, string>> = {};

  constructor(profileName: string, profile: Profile, secrets: Secrets) {
    this.profileName = profileName;
    this.profile = profile;
    this.secrets = secrets;
  }

  get source(): 'static-token' | 'app' | 'delegated' {
    if (this.secrets.accessToken) return 'static-token';
    return this.profile.authMode;
  }

  private staticToken(resource: Resource): string | undefined {
    if (resource === 'exchange') return this.secrets.exchangeAccessToken;
    if (resource === 'defender') return this.secrets.defenderAccessToken;
    return this.secrets.accessToken;
  }

  private spec(resource: Resource): ResourceSpec {
    const graphClient = this.profile.clientId ?? DEFAULT_PUBLIC_CLIENT_ID;
    if (resource === 'exchange') {
      return {
        appScope: `${EXCHANGE_RESOURCE}/.default`,
        clientId: EXO_PUBLIC_CLIENT_ID,
        delegatedScopes: [`${EXCHANGE_RESOURCE}/.default`],
        cacheSuffix: '-exchange',
        account: 'exchangeAccount',
        loginHint: 'm365 auth login --service exchange',
      };
    }
    if (resource === 'defender') {
      // Microsoft's own public clients cannot be granted Defender for Endpoint scopes (AADSTS65002).
      if (this.profile.authMode === 'delegated' && graphClient === DEFAULT_PUBLIC_CLIENT_ID) {
        throw new CliError('CONFIG', 'Defender for Endpoint actions need your own app registration', {
          hint:
            'Use an app-only profile with WindowsDefenderATP application permissions (Machine.Isolate, ' +
            'Machine.Scan, ...), or register a public client with those delegated permissions and run ' +
            '`m365 config set clientId <app-id>` then `m365 auth login --service defender`.',
        });
      }
      return {
        appScope: `${DEFENDER_RESOURCE}/.default`,
        clientId: graphClient,
        delegatedScopes: [`${DEFENDER_RESOURCE}/.default`],
        cacheSuffix: '',
        account: 'account',
        loginHint: 'm365 auth login --service defender',
      };
    }
    return {
      appScope: `${(this.profile.graphUrl ?? DEFAULT_GRAPH_URL).replace(/\/+$/, '')}/.default`,
      clientId: graphClient,
      delegatedScopes: this.profile.scopes ?? ['User.Read'],
      cacheSuffix: '',
      account: 'account',
      loginHint: 'm365 auth login',
    };
  }

  private cachePath(suffix: string): string {
    return join(configDir(), `msal-cache-${this.profileName}${suffix}.json`);
  }

  private baseConfig(tenant: string, clientId: string, cacheSuffix: string): Configuration {
    return {
      auth: { clientId, authority: `${this.profile.authorityHost ?? DEFAULT_AUTHORITY_HOST}/${tenant}` },
      cache: { cachePlugin: fileCachePlugin(this.cachePath(cacheSuffix)) },
      system: { loggerOptions: { logLevel: LogLevel.Error, piiLoggingEnabled: false, loggerCallback: () => {} } },
    };
  }

  private confidentialClient(): ConfidentialClientApplication {
    const { tenantId, clientId, certificatePath } = this.profile;
    if (!tenantId || !clientId) {
      throw new CliError('CONFIG', 'App-only auth needs a tenant id and client id', {
        hint: 'Run `m365 config init --mode app --tenant <id> --client-id <id> [--cert <pem>]`.',
      });
    }
    const config = this.baseConfig(tenantId, clientId, '');
    if (certificatePath) {
      config.auth.clientCertificate = loadCertificate(this.profile, this.secrets);
    } else if (this.secrets.clientSecret) {
      config.auth.clientSecret = this.secrets.clientSecret;
    } else {
      throw new CliError('CONFIG', 'App-only auth needs a certificate or a client secret', {
        hint: 'Set certificatePath in the profile (preferred) or export M365_CLIENT_SECRET.',
      });
    }
    return new ConfidentialClientApplication(config);
  }

  private publicClient(spec: ResourceSpec): PublicClientApplication {
    const tenant = this.profile.tenantId ?? this.profile[spec.account]?.tenantId ?? 'organizations';
    return new PublicClientApplication(this.baseConfig(tenant, spec.clientId, spec.cacheSuffix));
  }

  private async signedInAccount(client: PublicClientApplication, spec: ResourceSpec): Promise<AccountInfo> {
    const cache = client.getTokenCache();
    const homeId = this.profile[spec.account]?.homeAccountId;
    const account = homeId ? await cache.getAccountByHomeId(homeId) : (await cache.getAllAccounts())[0];
    if (!account) {
      throw new CliError(
        'AUTH_REQUIRED',
        `Profile "${this.profileName}" is not signed in${spec.account === 'exchangeAccount' ? ' to Exchange Online' : ''}`,
        {
          hint: `Run \`${spec.loginHint}\` and complete the device code prompt.`,
        },
      );
    }
    return account;
  }

  async getToken(resource: Resource = 'graph'): Promise<string> {
    const fixed = this.staticToken(resource);
    if (fixed) return fixed;
    const cached = this.tokens[resource];
    if (cached) return cached;
    const spec = this.spec(resource);
    try {
      let result: AuthenticationResult | null;
      if (this.profile.authMode === 'app') {
        result = await this.confidentialClient().acquireTokenByClientCredential({ scopes: [spec.appScope] });
      } else {
        const client = this.publicClient(spec);
        const account = await this.signedInAccount(client, spec);
        result = await client.acquireTokenSilent({ account, scopes: spec.delegatedScopes });
      }
      if (!result?.accessToken) throw new CliError('AUTH_FAILED', 'Token request returned no access token');
      this.tokens[resource] = result.accessToken;
      return result.accessToken;
    } catch (error) {
      throw msalError(error, spec.loginHint);
    }
  }

  /**
   * Delegated: run the device code flow, printing the code prompt to stderr. App-only: acquire a token to prove
   * the credentials work. Graph takes the scopes to consent; other resources use their fixed scope.
   */
  async login(
    resource: Resource,
    graphScopes: string[],
    prompt: (message: string) => void,
  ): Promise<AuthenticationResult> {
    delete this.tokens[resource];
    const spec = this.spec(resource);
    try {
      if (this.profile.authMode === 'app') {
        const result = await this.confidentialClient().acquireTokenByClientCredential({ scopes: [spec.appScope] });
        if (!result) throw new CliError('AUTH_FAILED', 'Token request returned nothing');
        return result;
      }
      const result = await this.publicClient(spec).acquireTokenByDeviceCode({
        scopes: resource === 'graph' ? graphScopes : spec.delegatedScopes,
        deviceCodeCallback: (response) => prompt(response.message),
      });
      if (!result?.account) throw new CliError('AUTH_FAILED', 'Device code sign-in returned no account');
      return result;
    } catch (error) {
      throw msalError(error, spec.loginHint);
    }
  }

  /** Remove every cached token for the profile (Graph/Defender and Exchange caches). */
  logout(): boolean {
    for (const key of Object.keys(this.tokens) as Resource[]) delete this.tokens[key];
    let removed = false;
    for (const suffix of ['', '-exchange']) {
      const path = this.cachePath(suffix);
      if (existsSync(path)) {
        rmSync(path);
        removed = true;
      }
    }
    return removed;
  }
}
