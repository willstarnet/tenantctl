import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { configDir, writePrivateFile } from './config.ts';

/**
 * Token caches hold refresh tokens, so they are encrypted at rest with AES-256-GCM. The key lives in the OS
 * secret store (macOS Keychain, libsecret, Windows DPAPI); only when none is usable does it fall back to a
 * 0600 key file next to the cache, which `m365 auth status` reports.
 */
export type KeyStore = 'keychain' | 'libsecret' | 'dpapi' | 'file';

const SERVICE = 'm365-cli';
const ACCOUNT = 'token-cache-key';
const ENVELOPE_VERSION = 1;

interface Envelope {
  v: number;
  alg: 'aes-256-gcm';
  iv: string;
  tag: string;
  data: string;
}

function run(cmd: string[], stdin?: string): { ok: boolean; out: string } {
  const which = Bun.which(cmd[0] as string);
  if (!which) return { ok: false, out: '' };
  const result = Bun.spawnSync([which, ...cmd.slice(1)], {
    stdin: stdin === undefined ? 'ignore' : Buffer.from(stdin),
    stdout: 'pipe',
    stderr: 'ignore',
  });
  return { ok: result.exitCode === 0, out: result.stdout.toString().trim() };
}

function keychainKey(): Buffer | null {
  const found = run(['security', 'find-generic-password', '-s', SERVICE, '-a', ACCOUNT, '-w']);
  if (found.ok && /^[0-9a-f]{64}$/.test(found.out)) return Buffer.from(found.out, 'hex');
  const key = randomBytes(32);
  const added = run([
    'security',
    'add-generic-password',
    '-s',
    SERVICE,
    '-a',
    ACCOUNT,
    '-w',
    key.toString('hex'),
    '-U',
  ]);
  return added.ok ? key : null;
}

function libsecretKey(): Buffer | null {
  const found = run(['secret-tool', 'lookup', 'service', SERVICE, 'account', ACCOUNT]);
  if (found.ok && /^[0-9a-f]{64}$/.test(found.out)) return Buffer.from(found.out, 'hex');
  const key = randomBytes(32);
  const stored = run(
    ['secret-tool', 'store', '--label', 'm365 CLI token cache key', 'service', SERVICE, 'account', ACCOUNT],
    key.toString('hex'),
  );
  return stored.ok ? key : null;
}

/** DPAPI ties the protected key to the Windows user account; the protected blob itself can sit in a file. */
function dpapiKey(): Buffer | null {
  const path = join(configDir(), 'cache.key.dpapi');
  const shell = Bun.which('pwsh') ? 'pwsh' : 'powershell';
  const prelude = 'Add-Type -AssemblyName System.Security;';
  if (existsSync(path)) {
    const unprotected = run(
      [
        shell,
        '-NoProfile',
        '-Command',
        `${prelude}[Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String([Console]::In.ReadToEnd()), $null, 'CurrentUser'))`,
      ],
      readFileSync(path, 'utf8'),
    );
    if (unprotected.ok) return Buffer.from(unprotected.out, 'base64');
  }
  const key = randomBytes(32);
  const protectedKey = run(
    [
      shell,
      '-NoProfile',
      '-Command',
      `${prelude}[Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect([Convert]::FromBase64String([Console]::In.ReadToEnd()), $null, 'CurrentUser'))`,
    ],
    key.toString('base64'),
  );
  if (!protectedKey.ok) return null;
  writePrivateFile(path, protectedKey.out);
  return key;
}

function fileKey(): Buffer {
  const path = join(configDir(), 'cache.key');
  if (existsSync(path)) return Buffer.from(readFileSync(path, 'utf8').trim(), 'hex');
  const key = randomBytes(32);
  writePrivateFile(path, key.toString('hex'));
  return key;
}

let resolved: { key: Buffer; store: KeyStore } | undefined;

/** The cache key and where it is kept. `M365_TOKEN_CACHE_KEYSTORE=file` forces the file store (CI, containers). */
export function cacheKey(): { key: Buffer; store: KeyStore } {
  if (resolved) return resolved;
  const forced = process.env.M365_TOKEN_CACHE_KEYSTORE;
  if (forced !== 'file') {
    const attempts: [KeyStore, () => Buffer | null][] =
      process.platform === 'darwin'
        ? [['keychain', keychainKey]]
        : process.platform === 'win32'
          ? [['dpapi', dpapiKey]]
          : [['libsecret', libsecretKey]];
    for (const [store, load] of attempts) {
      const key = load();
      if (key) return (resolved = { key, store });
    }
  }
  return (resolved = { key: fileKey(), store: 'file' });
}

export function encryptText(plain: string): string {
  const { key } = cacheKey();
  const iv = randomBytes(12);
  const cipher = createCipheriv('aes-256-gcm', key, iv);
  const data = Buffer.concat([cipher.update(plain, 'utf8'), cipher.final()]);
  const envelope: Envelope = {
    v: ENVELOPE_VERSION,
    alg: 'aes-256-gcm',
    iv: iv.toString('base64'),
    tag: cipher.getAuthTag().toString('base64'),
    data: data.toString('base64'),
  };
  return JSON.stringify(envelope);
}

function isEnvelope(value: unknown): value is Envelope {
  return (
    typeof value === 'object' &&
    value !== null &&
    'alg' in value &&
    value.alg === 'aes-256-gcm' &&
    'data' in value &&
    typeof value.data === 'string'
  );
}

/**
 * Decrypt an envelope. Returns `{ plain, legacy: true }` for an older plaintext cache so the caller re-saves it
 * encrypted, and null when the key no longer matches (the user simply signs in again).
 */
export function decryptText(stored: string): { plain: string; legacy: boolean } | null {
  let parsed: unknown;
  try {
    parsed = JSON.parse(stored);
  } catch {
    return null;
  }
  if (!isEnvelope(parsed)) return { plain: stored, legacy: true };
  try {
    const decipher = createDecipheriv('aes-256-gcm', cacheKey().key, Buffer.from(parsed.iv, 'base64'));
    decipher.setAuthTag(Buffer.from(parsed.tag, 'base64'));
    const plain = Buffer.concat([decipher.update(Buffer.from(parsed.data, 'base64')), decipher.final()]).toString(
      'utf8',
    );
    return { plain, legacy: false };
  } catch {
    return null;
  }
}
