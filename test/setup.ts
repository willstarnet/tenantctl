import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// Tests never touch the real config dir or a real tenant.
process.env.TENANTCTL_CONFIG_DIR = mkdtempSync(join(tmpdir(), 'tenantctl-test-'));
process.env.TENANTCTL_ACCESS_TOKEN = 'test-token';
delete process.env.TENANTCTL_ALLOW_WRITES;
delete process.env.TENANTCTL_PROFILE;
process.env.TENANTCTL_TOKEN_CACHE_KEYSTORE = 'file';
process.env.TENANTCTL_EXCHANGE_ACCESS_TOKEN = 'test-exchange-token';
process.env.TENANTCTL_DEFENDER_ACCESS_TOKEN = 'test-defender-token';
delete process.env.TENANTCTL_CONFIRM_MODE;
