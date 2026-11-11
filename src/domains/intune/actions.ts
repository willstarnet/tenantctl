import { CliError } from '../../core/errors.ts';
import { type ApiVersion, seg } from '../../core/graph.ts';
import { type AnyCommand, type ArgSpec, type Context, defineCommand, type Risk } from '../../core/types.ts';
import {
  DEVICES_PATH,
  deviceSummary,
  INTUNE_ADMIN_ROLES,
  type ManagedDevice,
  type Platform,
  platformOf,
  PRIVILEGED_OPS,
  READ_DEVICES,
  resolveManagedDevice,
  WRITE_DEVICES,
} from './common.ts';

const DEVICE_ARG: ArgSpec = {
  name: 'device',
  type: 'string',
  required: true,
  positional: true,
  description: 'Device name, serial number, Intune device id or Entra device id.',
};

const PLATFORM_NAMES: Record<Platform, string> = {
  windows: 'Windows',
  ios: 'iOS/iPadOS',
  android: 'Android',
  macos: 'macOS',
  linux: 'Linux',
  other: 'this platform',
};

type Args = Record<string, unknown>;

interface ActionSpec {
  name: string;
  summary: string;
  description: string;
  examples: string[];
  risk: Exclude<Risk, 'read'>;
  /** Bound action segment after the device, e.g. `syncDevice`; omitted for DELETE of the device itself. */
  action?: string;
  version?: ApiVersion;
  permission: string;
  roles?: string[];
  /** Platforms Intune supports the action on; others are refused before anything is sent. */
  platforms?: Platform[];
  args?: ArgSpec[];
  body?: (args: Args, device: ManagedDevice) => unknown;
  /** Plain-language description of what will happen to this device. */
  describe: (device: ManagedDevice, args: Args) => string;
  /** Extra argument checks that depend on the device. */
  validate?: (device: ManagedDevice, args: Args) => void;
}

function checkPlatform(spec: ActionSpec, device: ManagedDevice): void {
  if (!spec.platforms) return;
  const platform = platformOf(device);
  if (spec.platforms.includes(platform)) return;
  throw new CliError(
    'USAGE',
    `"devices ${spec.name}" is not supported on ${device.operatingSystem ?? PLATFORM_NAMES[platform]} (${device.deviceName ?? device.id})`,
    { hint: `Intune supports it on ${spec.platforms.map((p) => PLATFORM_NAMES[p]).join(', ')} only.` },
  );
}

async function target(ctx: Context, spec: ActionSpec, args: Args): Promise<ManagedDevice> {
  const device = await resolveManagedDevice(ctx, String(args.device));
  checkPlatform(spec, device);
  spec.validate?.(device, args);
  return device;
}

/**
 * Remote device actions share one shape: resolve the device, refuse unsupported platforms, preview, then POST
 * the bound action. Every action previews (destructive ones must; writes use it for --dry-run).
 */
function deviceAction(spec: ActionSpec): AnyCommand {
  const requestPath = (device: ManagedDevice) =>
    spec.action ? `${DEVICES_PATH}/${seg(device.id)}/${spec.action}` : `${DEVICES_PATH}/${seg(device.id)}`;
  const permissions = spec.permission === PRIVILEGED_OPS ? [READ_DEVICES, PRIVILEGED_OPS] : [spec.permission];
  return defineCommand({
    domain: 'intune',
    name: `devices ${spec.name}`,
    summary: spec.summary,
    description: spec.description,
    examples: spec.examples,
    args: [DEVICE_ARG, ...(spec.args ?? [])] as readonly ArgSpec[],
    risk: spec.risk,
    permissions: { delegated: permissions, application: permissions, roles: spec.roles ?? INTUNE_ADMIN_ROLES },
    async preview(ctx, args) {
      const device = await target(ctx, spec, args);
      return {
        action: spec.describe(device, args),
        device: deviceSummary(device),
        request: {
          method: spec.action ? 'POST' : 'DELETE',
          path: requestPath(device),
          version: spec.version ?? 'v1.0',
          body: spec.body?.(args, device),
        },
      };
    },
    async run(ctx, args) {
      const device = await target(ctx, spec, args);
      const options = { version: spec.version };
      const response = spec.action
        ? await ctx.graph.post(requestPath(device), spec.body?.(args, device), options)
        : await ctx.graph.delete(requestPath(device), options);
      return {
        status: spec.action ? 'requested' : 'deleted',
        action: spec.describe(device, args),
        device: deviceSummary(device),
        ...(response !== undefined ? { response } : {}),
        next: spec.action
          ? `The device acts on this at its next check-in. Track it with \`tenantctl intune devices actions ${device.deviceName ?? device.id}\`.`
          : undefined,
      };
    },
  });
}

const name = (device: ManagedDevice) => device.deviceName ?? device.id;

function who(device: ManagedDevice): string {
  const parts = [
    device.userPrincipalName ? `user ${device.userPrincipalName}` : 'no primary user',
    [device.operatingSystem, device.osVersion].filter(Boolean).join(' ') || 'unknown OS',
    device.serialNumber ? `serial ${device.serialNumber}` : 'no serial',
    device.lastSyncDateTime ? `last sync ${device.lastSyncDateTime}` : 'never synced',
  ];
  return parts.join(', ');
}

export const commands: AnyCommand[] = [
  deviceAction({
    name: 'sync',
    summary: 'Ask a device to check in with Intune now',
    description:
      'Use after changing a policy, assignment or app so the device picks it up without waiting for the next ' +
      'scheduled check-in (about every 8 hours). The device must be online; check lastSyncDateTime with ' +
      '`devices get` a few minutes later.',
    examples: ['tenantctl intune devices sync DESKTOP-4F2K9'],
    risk: 'write',
    action: 'syncDevice',
    permission: PRIVILEGED_OPS,
    describe: (d) => `Send a sync (check-in) request to ${name(d)}.`,
  }),
  deviceAction({
    name: 'restart',
    summary: 'Reboot a device immediately',
    description:
      'Reboots the device as soon as it receives the command, without asking the user; unsaved work is lost. ' +
      'Windows, macOS, supervised iOS/iPadOS and Android Enterprise dedicated/fully managed devices.',
    examples: ['tenantctl intune devices restart DESKTOP-4F2K9'],
    risk: 'destructive',
    action: 'rebootNow',
    permission: PRIVILEGED_OPS,
    platforms: ['windows', 'macos', 'ios', 'android'],
    describe: (d) => `Reboot ${name(d)} immediately (${who(d)}). Unsaved work on the device is lost.`,
  }),
  deviceAction({
    name: 'lock',
    summary: 'Remotely lock a phone, tablet or Mac',
    description:
      'Locks the screen; the user unlocks with their passcode. Use for a lost or stolen mobile device before ' +
      'deciding to wipe. Not supported on Windows (use `devices restart` or disable the user instead).',
    examples: ['tenantctl intune devices lock "Jane iPhone"'],
    risk: 'destructive',
    action: 'remoteLock',
    permission: PRIVILEGED_OPS,
    platforms: ['ios', 'android', 'macos'],
    describe: (d) => `Lock the screen of ${name(d)} now (${who(d)}).`,
  }),
  deviceAction({
    name: 'locate',
    summary: 'Ask a device to report its location',
    description:
      'Supported on Windows 10/11, supervised iOS/iPadOS in Lost Mode and Android Enterprise dedicated/fully ' +
      'managed devices; location services must be on. The location appears in the Intune admin center once the ' +
      'device responds; `devices actions` shows whether it completed.',
    examples: ['tenantctl intune devices locate "Jane iPhone"'],
    risk: 'write',
    action: 'locateDevice',
    permission: WRITE_DEVICES,
    platforms: ['windows', 'ios', 'android'],
    describe: (d) => `Request the current location of ${name(d)}.`,
  }),
  deviceAction({
    name: 'defender-scan',
    summary: 'Start a Microsoft Defender Antivirus scan on a Windows device',
    description:
      'Quick scan by default; --full scans every file and can take hours. Results show in `devices get` ' +
      '(defender.lastQuickScanDateTime / lastFullScanDateTime) after the device reports back.',
    examples: [
      'tenantctl intune devices defender-scan DESKTOP-4F2K9',
      'tenantctl intune devices defender-scan DESKTOP-4F2K9 --full',
    ],
    risk: 'write',
    action: 'windowsDefenderScan',
    permission: PRIVILEGED_OPS,
    platforms: ['windows'],
    args: [{ name: 'fullScan', type: 'boolean', description: 'Full scan instead of a quick scan.' }],
    body: (args) => ({ quickScan: !args.fullScan }),
    describe: (d, args) => `Start a Defender ${args.fullScan ? 'full' : 'quick'} scan on ${name(d)}.`,
  }),
  deviceAction({
    name: 'defender-update',
    summary: 'Update Microsoft Defender Antivirus signatures on a Windows device',
    description: 'Use when `devices get` shows defender.signatureUpdateOverdue. Windows only.',
    examples: ['tenantctl intune devices defender-update DESKTOP-4F2K9'],
    risk: 'write',
    action: 'windowsDefenderUpdateSignatures',
    permission: PRIVILEGED_OPS,
    platforms: ['windows'],
    describe: (d) => `Update Defender security intelligence on ${name(d)}.`,
  }),
  deviceAction({
    name: 'rotate-bitlocker',
    summary: 'Rotate the BitLocker recovery key of a Windows device',
    description:
      'Use after a recovery key was shared with a user, or when `devices bitlocker-key` finds no key in Entra. ' +
      'The device generates a new key and escrows it to Entra on its next check-in. Requires Entra joined or ' +
      'hybrid joined Windows with key rotation enabled in the BitLocker policy. (beta API)',
    examples: ['tenantctl intune devices rotate-bitlocker DESKTOP-4F2K9'],
    risk: 'write',
    action: 'rotateBitLockerKeys',
    version: 'beta',
    permission: WRITE_DEVICES,
    platforms: ['windows'],
    describe: (d) => `Rotate the BitLocker recovery key on ${name(d)}; the old key stops working.`,
  }),
  deviceAction({
    name: 'rotate-laps',
    summary: 'Rotate the Windows LAPS local admin password now',
    description:
      'Use after the LAPS password was disclosed. The device sets a new password and backs it up on its next ' +
      'check-in; read it with `devices laps --reveal`. Requires a Windows LAPS policy. (beta API)',
    examples: ['tenantctl intune devices rotate-laps DESKTOP-4F2K9'],
    risk: 'write',
    action: 'rotateLocalAdminPassword',
    version: 'beta',
    permission: PRIVILEGED_OPS,
    platforms: ['windows'],
    describe: (d) => `Rotate the local administrator password on ${name(d)}.`,
  }),
  deviceAction({
    name: 'rename',
    summary: 'Rename a device (sets the hostname on the device)',
    description:
      'Changes the device name on the device itself, not just the Intune record. Windows needs a restart to take ' +
      'effect; iOS/iPadOS must be supervised. Windows names: max 15 characters, letters, digits and hyphens. ' +
      '(beta API)',
    examples: ['tenantctl intune devices rename DESKTOP-4F2K9 --name SALES-LT-042'],
    risk: 'write',
    action: 'setDeviceName',
    version: 'beta',
    permission: PRIVILEGED_OPS,
    platforms: ['windows', 'ios', 'macos'],
    args: [{ name: 'name', type: 'string', required: true, description: 'New device name.' }],
    body: (args) => ({ deviceName: String(args.name) }),
    validate: (d, args) => {
      const value = String(args.name);
      if (platformOf(d) === 'windows' && !/^(?![0-9]+$)[A-Za-z0-9-]{1,15}$/.test(value)) {
        throw new CliError('USAGE', `"${value}" is not a valid Windows device name`, {
          hint: 'Use 1-15 letters, digits or hyphens, not all digits.',
        });
      }
    },
    describe: (d, args) => `Rename ${name(d)} to ${String(args.name)}.`,
  }),
  deviceAction({
    name: 'collect-logs',
    summary: 'Collect diagnostic logs from a Windows device',
    description:
      'Asks the device to upload its Intune diagnostic bundle (event logs, registry, MDM diagnostics). Download ' +
      'it from the device\'s "Device diagnostics" page in the Intune admin center once collected. Windows 10 ' +
      '1909+ only; the response includes the request id and status. (beta API)',
    examples: ['tenantctl intune devices collect-logs DESKTOP-4F2K9'],
    risk: 'write',
    action: 'createDeviceLogCollectionRequest',
    version: 'beta',
    permission: WRITE_DEVICES,
    platforms: ['windows'],
    body: () => ({ templateType: { templateType: 'predefined' } }),
    describe: (d) => `Collect diagnostic logs from ${name(d)}.`,
  }),
  deviceAction({
    name: 'retire',
    summary: 'Retire a device: remove company data, apps and profiles; keep personal data',
    description:
      'Use when an employee leaves with a personal (BYOD) device or a device is being repurposed. Removes ' +
      'managed apps, company data, Wi-Fi/VPN/email profiles and unenrolls the device; personal data stays. ' +
      'The Intune record is removed after the device checks in (or after 180 days if it never does).',
    examples: ['tenantctl intune devices retire "Jane iPhone"'],
    risk: 'destructive',
    action: 'retire',
    permission: PRIVILEGED_OPS,
    describe: (d) =>
      `Retire ${name(d)} (${who(d)}): remove company apps, data and profiles and unenroll it from Intune. ` +
      'Personal data is kept.',
  }),
  deviceAction({
    name: 'wipe',
    summary: 'Factory reset a device, erasing all data',
    description:
      'Use for a lost/stolen corporate device or before reissuing one. Erases everything on the device and ' +
      'cannot be undone. --keep-enrollment-data (Windows) keeps the device enrolled / Autopilot-provisioned; ' +
      '--keep-user-data (Windows) keeps user files but removes apps and settings. Macs with Apple silicon or T2 ' +
      'may need --macos-unlock-code (6 digits) to unlock after the wipe. Consider `devices retire` for personal ' +
      'devices.',
    examples: [
      'tenantctl intune devices wipe DESKTOP-4F2K9',
      'tenantctl intune devices wipe DESKTOP-4F2K9 --keep-enrollment-data',
    ],
    risk: 'destructive',
    action: 'wipe',
    permission: PRIVILEGED_OPS,
    platforms: ['windows', 'ios', 'android', 'macos'],
    args: [
      { name: 'keepEnrollmentData', type: 'boolean', description: 'Windows: keep enrollment/provisioning data.' },
      { name: 'keepUserData', type: 'boolean', description: 'Windows: keep user files.' },
      { name: 'macosUnlockCode', type: 'string', secret: true, description: 'macOS: six-digit recovery PIN.' },
    ],
    validate: (d, args) => {
      if ((args.keepEnrollmentData || args.keepUserData) && platformOf(d) !== 'windows') {
        throw new CliError('USAGE', '--keep-enrollment-data and --keep-user-data apply to Windows devices only');
      }
      if (args.macosUnlockCode !== undefined) {
        if (platformOf(d) !== 'macos') throw new CliError('USAGE', '--macos-unlock-code applies to macOS devices only');
        if (!/^\d{6}$/.test(String(args.macosUnlockCode)))
          throw new CliError('USAGE', '--macos-unlock-code must be 6 digits');
      }
    },
    body: (args) => ({
      keepEnrollmentData: Boolean(args.keepEnrollmentData),
      keepUserData: Boolean(args.keepUserData),
      ...(args.macosUnlockCode !== undefined ? { macOsUnlockCode: String(args.macosUnlockCode) } : {}),
    }),
    describe: (d, args) =>
      `FACTORY RESET ${name(d)}: ${args.keepUserData ? 'remove all apps and settings (user files kept)' : 'ERASE ALL DATA'} ` +
      `on this device (${who(d)}). This cannot be undone.` +
      (args.keepEnrollmentData ? ' Enrollment data is kept so it re-enrolls automatically.' : ''),
  }),
  deviceAction({
    name: 'fresh-start',
    summary: 'Reinstall Windows and remove OEM bloatware (Fresh Start)',
    description:
      'Removes apps that were not installed with Windows (including OEM software and Win32 apps) and updates ' +
      "Windows. With --keep-user-data the user's files stay; otherwise everything except the OS is removed. " +
      'Entra joined devices stay joined. Windows only.',
    examples: ['tenantctl intune devices fresh-start DESKTOP-4F2K9 --keep-user-data'],
    risk: 'destructive',
    action: 'cleanWindowsDevice',
    permission: PRIVILEGED_OPS,
    platforms: ['windows'],
    args: [{ name: 'keepUserData', type: 'boolean', description: 'Keep user files.' }],
    body: (args) => ({ keepUserData: Boolean(args.keepUserData) }),
    describe: (d, args) =>
      `Fresh Start ${name(d)} (${who(d)}): remove all installed apps and reinstall Windows; user files are ` +
      `${args.keepUserData ? 'kept' : 'DELETED'}.`,
  }),
  deviceAction({
    name: 'delete',
    summary: 'Delete the Intune record of a device',
    description:
      'Removes the device from Intune without touching the device; use for stale records of devices that are ' +
      'already gone. If the device is still in use, its management certificate stops working and it drops out ' +
      'of compliance. Use `devices retire` or `devices wipe` to remove data first.',
    examples: ['tenantctl intune devices delete OLD-LAPTOP-01'],
    risk: 'destructive',
    permission: WRITE_DEVICES,
    describe: (d) => `Delete the Intune record of ${name(d)} (${who(d)}). The device itself is not changed.`,
  }),
  deviceAction({
    name: 'reset-passcode',
    summary: 'Remove or reset the passcode of a phone or tablet',
    description:
      'For a user locked out of a managed mobile device. Android Enterprise work profile / fully managed: a new ' +
      'temporary passcode is generated (shown in the Intune admin center). iOS/iPadOS: the passcode is removed. ' +
      'Not supported on Windows or macOS.',
    examples: ['tenantctl intune devices reset-passcode "Jane Pixel"'],
    risk: 'destructive',
    action: 'resetPasscode',
    permission: PRIVILEGED_OPS,
    platforms: ['android', 'ios'],
    describe: (d) => `Reset or remove the passcode on ${name(d)} (${who(d)}).`,
  }),
];
