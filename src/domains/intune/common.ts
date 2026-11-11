import { CliError } from '../../core/errors.ts';
import { type ApiVersion, odata, type QueryValue, seg } from '../../core/graph.ts';
import { pickOne } from '../../core/resolve.ts';
import type { Context, ListResult } from '../../core/types.ts';
import { compact, daysSince, isGuid } from '../../core/util.ts';

export const DEVICES_PATH = '/deviceManagement/managedDevices';

export const READ_DEVICES = 'DeviceManagementManagedDevices.Read.All';
export const READ_CONFIG = 'DeviceManagementConfiguration.Read.All';
export const PRIVILEGED_OPS = 'DeviceManagementManagedDevices.PrivilegedOperations.All';
export const WRITE_DEVICES = 'DeviceManagementManagedDevices.ReadWrite.All';
export const INTUNE_READ_ROLES = ['Intune Administrator', 'Global Reader'];
export const INTUNE_ADMIN_ROLES = ['Intune Administrator'];

/** The managedDevice properties every compact view needs; list responses omit some unless selected. */
export const DEVICE_SELECT = [
  'id',
  'deviceName',
  'userId',
  'userPrincipalName',
  'userDisplayName',
  'operatingSystem',
  'osVersion',
  'model',
  'manufacturer',
  'serialNumber',
  'complianceState',
  'managementState',
  'managedDeviceOwnerType',
  'enrolledDateTime',
  'lastSyncDateTime',
  'isEncrypted',
  'azureADDeviceId',
].join(',');

export interface ManagedDevice {
  id: string;
  deviceName: string | null;
  userId?: string | null;
  userPrincipalName?: string | null;
  userDisplayName?: string | null;
  operatingSystem?: string | null;
  osVersion?: string | null;
  model?: string | null;
  manufacturer?: string | null;
  serialNumber?: string | null;
  complianceState?: string | null;
  managementState?: string | null;
  managedDeviceOwnerType?: string | null;
  enrolledDateTime?: string | null;
  lastSyncDateTime?: string | null;
  isEncrypted?: boolean | null;
  azureADDeviceId?: string | null;
  [key: string]: unknown;
}

export type Platform = 'windows' | 'ios' | 'android' | 'macos' | 'linux' | 'other';

/** Intune reports free-text OS names (`Windows`, `iOS`, `iPadOS`, `AndroidForWork`, `macOS`, `Linux`). */
export function platformOf(device: Pick<ManagedDevice, 'operatingSystem'>): Platform {
  const os = (device.operatingSystem ?? '').toLowerCase();
  if (os.startsWith('windows')) return 'windows';
  if (os === 'ios' || os === 'ipados') return 'ios';
  if (os.startsWith('android')) return 'android';
  if (os.startsWith('mac')) return 'macos';
  if (os.startsWith('linux')) return 'linux';
  return 'other';
}

/** Intune reports "never synced" as 0001-01-01; treat it as null. */
function realDate(iso: string | null | undefined): string | null {
  return iso && !iso.startsWith('0001-') ? iso : null;
}

export function shapeManagedDevice(device: ManagedDevice) {
  return compact({
    id: device.id,
    deviceName: device.deviceName,
    user: device.userPrincipalName || null,
    os: device.operatingSystem ?? null,
    osVersion: device.osVersion ?? null,
    model: device.model ?? null,
    manufacturer: device.manufacturer ?? null,
    serialNumber: device.serialNumber || null,
    complianceState: device.complianceState ?? null,
    managementState: device.managementState,
    ownership: device.managedDeviceOwnerType ?? null,
    enrolledDateTime: realDate(device.enrolledDateTime),
    lastSyncDateTime: realDate(device.lastSyncDateTime),
    daysSinceLastSync: daysSince(realDate(device.lastSyncDateTime)),
    isEncrypted: device.isEncrypted ?? null,
    azureADDeviceId: device.azureADDeviceId || null,
  });
}

/** Header every preview and action result carries, so the agent can confirm it hit the right machine. */
export function deviceSummary(device: ManagedDevice) {
  return {
    id: device.id,
    deviceName: device.deviceName,
    user: device.userPrincipalName || null,
    os: [device.operatingSystem, device.osVersion].filter(Boolean).join(' ') || null,
    serialNumber: device.serialNumber || null,
    lastSyncDateTime: realDate(device.lastSyncDateTime),
    daysSinceLastSync: daysSince(realDate(device.lastSyncDateTime)),
    complianceState: device.complianceState ?? null,
    managementState: device.managementState ?? null,
  };
}

/** One $filter clause, plus the equivalent client-side test for when Graph rejects it. */
export interface DeviceClause {
  filter: string;
  /** Listed as filterable in the managedDevice docs (complianceState, lastSyncDateTime, enrolledDateTime). */
  documented: boolean;
  /** Never sent to Graph: Intune accepts the filter but returns wrong results, so only `test` applies. */
  clientOnly?: boolean;
  test: (device: ManagedDevice) => boolean;
}

/**
 * Devices whose primary user is `user`. Intune rejects `userId eq` with 400, and `/users/{id}/managedDevices`
 * returned none of one user's two devices in a live check of 59 users (2026-11-10), so match the primary user's UPN,
 * which Intune filters case-insensitively but does not document. If Graph rejects it, the scan matches on userId.
 */
export function primaryUserClause(user: { id: string; userPrincipalName: string }): DeviceClause {
  return {
    filter: `userPrincipalName eq ${odata(user.userPrincipalName)}`,
    documented: false,
    test: (device) => device.userId === user.id,
  };
}

/**
 * Page through a collection keeping only items that pass `keep`, stopping at `limit` matches. Used when Graph
 * cannot filter server-side; cost is proportional to the tenant's collection size.
 */
export async function scan<T>(
  ctx: Context,
  path: string,
  query: Record<string, QueryValue>,
  keep: (item: T) => boolean,
  limit: number,
  version: ApiVersion = 'v1.0',
): Promise<ListResult<T>> {
  let next: string | undefined = ctx.graph.url(path, { version, query: { ...query, $top: 100 } });
  const items: T[] = [];
  while (next) {
    const page: { value?: T[]; '@odata.nextLink'?: string } = await ctx.graph.get(next);
    for (const item of page.value ?? []) {
      if (!keep(item)) continue;
      if (items.length >= limit) return { items, truncated: true };
      items.push(item);
    }
    next = page['@odata.nextLink'];
  }
  return { items, truncated: false };
}

const joinClauses = (clauses: DeviceClause[]) =>
  clauses.length ? clauses.map((c) => `(${c.filter})`).join(' and ') : undefined;

/**
 * List managed devices with server-side filtering. The managedDevice docs only promise $filter on a few
 * properties; others (operatingSystem, ownership, userPrincipalName, deviceName, serialNumber) work in practice but
 * are undocumented, so on a 400 the undocumented clauses move client-side instead of failing the command.
 */
export async function listDevices(
  ctx: Context,
  clauses: DeviceClause[],
  limit: number,
  select = DEVICE_SELECT,
): Promise<ListResult<ManagedDevice>> {
  const local = clauses.filter((c) => c.clientOnly);
  const remote = clauses.filter((c) => !c.clientOnly);
  const load = (server: DeviceClause[], client: DeviceClause[]) => {
    const query = { $select: select, $filter: joinClauses(server) };
    if (!client.length) return ctx.graph.list<ManagedDevice>(DEVICES_PATH, { query, limit });
    return scan<ManagedDevice>(ctx, DEVICES_PATH, query, (device) => client.every((c) => c.test(device)), limit);
  };
  try {
    return await load(remote, local);
  } catch (error) {
    const moved = remote.filter((c) => !c.documented);
    if (!(error instanceof CliError && error.status === 400) || !moved.length) throw error;
    ctx.log(`Graph rejected the device filter; applying ${moved.map((c) => c.filter).join(' and ')} client-side.`);
    return load(
      remote.filter((c) => c.documented),
      [...moved, ...local],
    );
  }
}

const same = (a: string | null | undefined, b: string) => (a ?? '').toLowerCase() === b.toLowerCase();

function describeDevice(device: ManagedDevice) {
  return {
    id: device.id,
    deviceName: device.deviceName,
    serialNumber: device.serialNumber || null,
    user: device.userPrincipalName || null,
    os: device.operatingSystem ?? null,
    lastSyncDateTime: realDate(device.lastSyncDateTime),
  };
}

/**
 * Find one Intune managed device from what a person would say: Intune device id, Entra device id
 * (azureADDeviceId), an Intune admin center URL containing it, device name (case-insensitive) or serial number.
 * Several devices share a name surprisingly often (re-enrolment, re-imaging), so collisions are AMBIGUOUS.
 */
export async function resolveManagedDevice(ctx: Context, ref: string, select = DEVICE_SELECT): Promise<ManagedDevice> {
  let value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A device is required');
  const fromUrl = value.match(/mdmDeviceId\/([0-9a-f-]{36})/i);
  if (fromUrl?.[1]) value = fromUrl[1];
  if (isGuid(value)) {
    const direct = await ctx.graph.tryGet<ManagedDevice>(`${DEVICES_PATH}/${seg(value)}`, {
      query: { $select: select },
    });
    if (direct) return direct;
    const byEntraId = await listDevices(
      ctx,
      [
        {
          filter: `azureADDeviceId eq ${odata(value)}`,
          documented: false,
          test: (d) => same(d.azureADDeviceId, value),
        },
      ],
      10,
      select,
    );
    if (!byEntraId.items.length) {
      throw new CliError('NOT_FOUND', `No Intune device with id or Entra device id ${value}`, {
        hint: 'The device may not be enrolled in Intune. Check Entra with `tenantctl entra devices get` if available.',
      });
    }
    return pickOne(byEntraId.items, value, 'device', (d) => [d.azureADDeviceId], describeDevice);
  }
  const [byName, bySerial] = await Promise.all([
    listDevices(
      ctx,
      [{ filter: `deviceName eq ${odata(value)}`, documented: false, test: (d) => same(d.deviceName, value) }],
      25,
      select,
    ),
    listDevices(
      ctx,
      [{ filter: `serialNumber eq ${odata(value)}`, documented: false, test: (d) => same(d.serialNumber, value) }],
      25,
      select,
    ),
  ]);
  const seen: Record<string, ManagedDevice> = {};
  for (const device of [...byName.items, ...bySerial.items]) seen[device.id] = device;
  let candidates = Object.values(seen);
  if (!candidates.length) {
    const fuzzy = await listDevices(
      ctx,
      [
        {
          filter: `startswith(deviceName,${odata(value)})`,
          documented: false,
          test: (d) => (d.deviceName ?? '').toLowerCase().startsWith(value.toLowerCase()),
        },
      ],
      25,
      select,
    );
    candidates = fuzzy.items;
  }
  return pickOne(candidates, value, 'device', (d) => [d.deviceName, d.serialNumber], describeDevice);
}

/** Run `fn` over `items` with at most `limit` in flight, so composite views do not trip Intune throttling. */
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let index = 0;
  const worker = async () => {
    while (index < items.length) {
      const current = index++;
      results[current] = await fn(items[current] as T);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

/** Turn a rejected promise into `{ error }` so one failing item does not sink a composite. */
export async function attempt<T>(promise: Promise<T>): Promise<T | { error: { code: string; message: string } }> {
  try {
    return await promise;
  } catch (error) {
    return {
      error: {
        code: error instanceof CliError ? error.code : 'INTERNAL',
        message: error instanceof Error ? error.message : String(error),
      },
    };
  }
}

/** `#microsoft.graph.win32LobApp` -> `win32LobApp`. */
export function odataType(item: Record<string, unknown>): string | null {
  const type = item['@odata.type'];
  return typeof type === 'string' ? type.replace(/^#microsoft\.graph\./, '') : null;
}

/** Find a policy/profile/app by id or display name among a small catalog list (Intune lists rarely exceed hundreds). */
export function pickByName<T extends { id: string }>(
  items: T[],
  ref: string,
  kind: string,
  name: (item: T) => string | null | undefined,
): T {
  const value = ref.trim();
  const byId = items.find((item) => item.id === value);
  if (byId) return byId;
  const wanted = value.toLowerCase();
  const matches = items.filter((item) => (name(item) ?? '').toLowerCase().includes(wanted));
  return pickOne(
    matches,
    value,
    kind,
    (item) => [name(item)],
    (item) => ({ id: item.id, name: name(item) }),
  );
}
