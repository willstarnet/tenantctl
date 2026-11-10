import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// Tests never touch the real config dir or a real tenant.
process.env.M365_CONFIG_DIR = mkdtempSync(join(tmpdir(), 'm365-test-'));
process.env.M365_ACCESS_TOKEN = 'test-token';
delete process.env.M365_ALLOW_WRITES;
delete process.env.M365_PROFILE;
process.env.M365_TOKEN_CACHE_KEYSTORE = 'file';
process.env.M365_EXCHANGE_ACCESS_TOKEN = 'test-exchange-token';
process.env.M365_DEFENDER_ACCESS_TOKEN = 'test-defender-token';
delete process.env.M365_CONFIRM_MODE;
