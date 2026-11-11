import { dlopen, type Pointer, ptr, toArrayBuffer } from 'bun:ffi';
import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { configDir, writePrivateFile } from './config.ts';

/**
 * Token caches hold refresh tokens, so they are encrypted at rest with AES-256-GCM. The key lives in the OS
 * secret store (macOS Keychain, libsecret, Windows DPAPI); only when none is usable does it fall back to a
 * 0600 key file next to the cache, which `tenantctl auth status` reports.
 */
export type KeyStore = 'keychain' | 'libsecret' | 'dpapi' | 'file';

const SERVICE = 'tenantctl';
/** Keychain/libsecret service before the CLI was renamed from `m365`; its key moves to SERVICE on first use. */
const LEGACY_SERVICE = 'm365-cli';
const HEX_KEY = /^[0-9a-f]{64}$/;
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

function keychainLookup(service: string): Buffer | null {
  const found = run(['security', 'find-generic-password', '-s', service, '-a', ACCOUNT, '-w']);
  return found.ok && HEX_KEY.test(found.out) ? Buffer.from(found.out, 'hex') : null;
}

function keychainKey(): Buffer | null {
  const found = keychainLookup(SERVICE);
  if (found) return found;
  const legacy = keychainLookup(LEGACY_SERVICE);
  const key = legacy ?? randomBytes(32);
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
  if (!added.ok) return null;
  if (legacy) run(['security', 'delete-generic-password', '-s', LEGACY_SERVICE, '-a', ACCOUNT]);
  return key;
}

function libsecretLookup(service: string): Buffer | null {
  const found = run(['secret-tool', 'lookup', 'service', service, 'account', ACCOUNT]);
  return found.ok && HEX_KEY.test(found.out) ? Buffer.from(found.out, 'hex') : null;
}

function libsecretKey(): Buffer | null {
  const found = libsecretLookup(SERVICE);
  if (found) return found;
  const legacy = libsecretLookup(LEGACY_SERVICE);
  const key = legacy ?? randomBytes(32);
  const stored = run(
    ['secret-tool', 'store', '--label', 'tenantctl token cache key', 'service', SERVICE, 'account', ACCOUNT],
    key.toString('hex'),
  );
  if (!stored.ok) return null;
  if (legacy) run(['secret-tool', 'clear', 'service', LEGACY_SERVICE, 'account', ACCOUNT]);
  return key;
}

/** CRYPTPROTECT_UI_FORBIDDEN: fail instead of showing a prompt (there is no one to answer it). */
const CRYPTPROTECT_UI_FORBIDDEN = 0x1;

/**
 * CryptProtectData / CryptUnprotectData for the current Windows user, called in-process. Spawning PowerShell for
 * this cost a PowerShell start on every command. Returns null when the call fails.
 */
export function dpapi(operation: 'protect' | 'unprotect', input: Buffer): Buffer | null {
  const signature = { args: ['ptr', 'ptr', 'ptr', 'ptr', 'ptr', 'u32', 'ptr'], returns: 'i32' } as const;
  const crypt32 = dlopen('crypt32.dll', { CryptProtectData: signature, CryptUnprotectData: signature });
  const kernel32 = dlopen('kernel32.dll', { LocalFree: { args: ['ptr'], returns: 'ptr' } });
  try {
    // DATA_BLOB on 64-bit Windows: DWORD cbData, 4 bytes of padding, BYTE* pbData.
    const blobIn = Buffer.alloc(16);
    blobIn.writeUInt32LE(input.length, 0);
    blobIn.writeBigUInt64LE(BigInt(ptr(input)), 8);
    const blobOut = Buffer.alloc(16);
    const call = operation === 'protect' ? crypt32.symbols.CryptProtectData : crypt32.symbols.CryptUnprotectData;
    if (!call(ptr(blobIn), null, null, null, null, CRYPTPROTECT_UI_FORBIDDEN, ptr(blobOut))) return null;
    const length = blobOut.readUInt32LE(0);
    // pbData is an address Windows allocated; bun:ffi represents native addresses as its branded Pointer number.
    const data = Number(blobOut.readBigUInt64LE(8)) as Pointer;
    // Copy out before LocalFree releases the memory Windows allocated for the result.
    const result = Buffer.from(new Uint8Array(toArrayBuffer(data, 0, length)));
    kernel32.symbols.LocalFree(data);
    return result;
  } finally {
    crypt32.close();
    kernel32.close();
  }
}

/** DPAPI ties the protected key to the Windows user account; the protected blob itself can sit in a file. */
function dpapiKey(): Buffer | null {
  const path = join(configDir(), 'cache.key.dpapi');
  if (existsSync(path)) {
    const key = dpapi('unprotect', Buffer.from(readFileSync(path, 'utf8').trim(), 'base64'));
    if (key) return key;
  }
  const key = randomBytes(32);
  const sealed = dpapi('protect', key);
  if (!sealed) return null;
  writePrivateFile(path, sealed.toString('base64'));
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

/** The cache key and where it is kept. `TENANTCTL_TOKEN_CACHE_KEYSTORE=file` forces the file store (CI, containers). */
export function cacheKey(): { key: Buffer; store: KeyStore } {
  if (resolved) return resolved;
  const forced = process.env.TENANTCTL_TOKEN_CACHE_KEYSTORE;
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
