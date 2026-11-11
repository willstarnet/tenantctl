import { odata, seg } from '../../core/graph.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { bytesToGb, isSettledError, settle } from '../../core/util.ts';
import { commands as entraCommands } from '../entra/index.ts';
import { type EntraDevice, shapeEntraDevice } from '../entra/directory.ts';
import { DEVICES_PATH, deviceSummary } from '../intune/common.ts';
import {
  complianceSummary,
  devicePolicyStates,
  commands as intuneCommands,
  type ManagedDevice,
  type PolicyState,
  type PolicyStateView,
  platformOf,
  resolveManagedDevice,
  shapePolicyStates,
} from '../intune/index.ts';
import { type Finding, permissionsOf, sortFindings, unavailable, verdict } from './findings.ts';
import { analyzeSignins, fetchSignins } from './signins.ts';

const SOURCES = [...entraCommands, ...intuneCommands];
const STALE_SYNC_DAYS = 7;
const VERY_STALE_SYNC_DAYS = 30;
const LOW_STORAGE_PERCENT = 10;
/** Policy states that mean "not working on this device". */
const BAD_STATES: Record<string, true> = { nonCompliant: true, error: true, conflict: true };
/** Intune reports a device with no Entra registration as an all-zero GUID. */
const NO_ENTRA_ID = /^0{8}-0{4}-0{4}-0{4}-0{12}$/;

interface DeviceDetail {
  freeStorageSpaceInBytes?: number | null;
  totalStorageSpaceInBytes?: number | null;
  jailBroken?: string | null;
  deviceActionResults?: { actionName?: string | null; actionState?: string | null; startDateTime?: string | null }[];
}

async function complianceStates(ctx: Context, deviceId: string): Promise<PolicyStateView[]> {
  return shapePolicyStates(await devicePolicyStates(ctx, deviceId));
}

async function configStates(ctx: Context, deviceId: string): Promise<PolicyStateView[]> {
  const { items } = await ctx.graph.list<PolicyState>(`${DEVICES_PATH}/${seg(deviceId)}/deviceConfigurationStates`, {
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  return shapePolicyStates(items);
}

async function entraObject(ctx: Context, entraDeviceId: string): Promise<EntraDevice | null> {
  const { items } = await ctx.graph.list<EntraDevice>('/devices', {
    query: {
      $filter: `deviceId eq ${odata(entraDeviceId)}`,
      $select:
        'id,deviceId,displayName,operatingSystem,operatingSystemVersion,accountEnabled,trustType,isCompliant,isManaged,approximateLastSignInDateTime,registrationDateTime',
    },
    limit: 1,
  });
  return items[0] ?? null;
}

async function deviceDetail(ctx: Context, id: string): Promise<DeviceDetail> {
  return ctx.graph.get<DeviceDetail>(`${DEVICES_PATH}/${seg(id)}`, {
    query: { $select: 'id,freeStorageSpaceInBytes,totalStorageSpaceInBytes,jailBroken,deviceActionResults' },
  });
}

/** Failed sign-ins of the primary user that came from this device (matched on the Entra device id). */
async function deviceSigninFailures(ctx: Context, device: ManagedDevice, days: number) {
  if (!device.userId) return null;
  const { items, truncated } = await fetchSignins(ctx, device.userId, days);
  const fromDevice = items.filter(
    (s) => s.deviceDetail?.deviceId && s.deviceDetail.deviceId === device.azureADDeviceId,
  );
  return analyzeSignins(fromDevice, days, truncated);
}

function syncFindings(device: ManagedDevice, name: string): Finding[] {
  const days = deviceSummary(device).daysSinceLastSync;
  if (days === null || days <= STALE_SYNC_DAYS) return [];
  return [
    {
      severity: days > VERY_STALE_SYNC_DAYS ? 'high' : 'medium',
      area: 'sync',
      issue: `Has not checked in with Intune for ${days} days`,
      detail: 'Compliance and policies reflect the last check-in. The device must be on, online and signed in to sync.',
      next: `tenantctl intune devices sync "${name}"`,
    },
  ];
}

function complianceFindings(device: ManagedDevice, policies: PolicyStateView[], name: string): Finding[] {
  if (device.complianceState === 'compliant') return [];
  const failing = policies.filter((p) => BAD_STATES[p.state ?? '']);
  return [
    {
      severity: device.complianceState === 'inGracePeriod' ? 'medium' : 'high',
      area: 'compliance',
      issue: `Device is ${device.complianceState ?? 'not compliant'}`,
      detail: complianceSummary(device, policies),
      next: failing.length
        ? `tenantctl intune devices compliance "${name}"`
        : `tenantctl intune devices sync "${name}"`,
    },
  ];
}

function entraFindings(device: ManagedDevice, entra: EntraDevice | null | undefined, name: string): Finding[] {
  if (!device.azureADDeviceId || NO_ENTRA_ID.test(device.azureADDeviceId)) {
    return [
      {
        severity: 'medium',
        area: 'entra',
        issue: 'Not registered in Entra ID',
        detail: 'Conditional Access cannot see the device, so "require compliant device" policies block it.',
      },
    ];
  }
  if (entra === undefined) return [];
  if (entra === null) {
    return [
      {
        severity: 'medium',
        area: 'entra',
        issue: 'Its Entra ID device object no longer exists',
        detail:
          'The device was deleted from Entra ID; it must re-register (rejoin) before Conditional Access trusts it.',
      },
    ];
  }
  const findings: Finding[] = [];
  if (entra.accountEnabled === false) {
    findings.push({
      severity: 'high',
      area: 'entra',
      issue: 'Device is disabled in Entra ID',
      detail: 'Sign-ins from it fail with 135011.',
      next: `tenantctl entra devices enable ${entra.id}`,
    });
  }
  if (device.complianceState === 'compliant' && entra.isCompliant === false) {
    findings.push({
      severity: 'medium',
      area: 'entra',
      issue: 'Intune says compliant but Entra ID still has it non-compliant',
      detail: 'The compliance result has not reached Entra ID yet; a sync usually resolves it.',
      next: `tenantctl intune devices sync "${name}"`,
    });
  }
  return findings;
}

function hardwareFindings(device: ManagedDevice, detail: DeviceDetail | null, name: string): Finding[] {
  const findings: Finding[] = [];
  const platform = platformOf(device);
  if (device.isEncrypted === false && (platform === 'windows' || platform === 'macos')) {
    findings.push({
      severity: 'medium',
      area: 'security',
      issue: 'Disk is not encrypted',
      detail: 'Most compliance policies require BitLocker/FileVault.',
      next: `tenantctl intune devices compliance "${name}"`,
    });
  }
  if (detail?.jailBroken === 'True') {
    findings.push({ severity: 'high', area: 'security', issue: 'Device reports as jailbroken/rooted' });
  }
  const total = detail?.totalStorageSpaceInBytes ?? 0;
  const free = detail?.freeStorageSpaceInBytes ?? 0;
  if (total > 0 && (free / total) * 100 < LOW_STORAGE_PERCENT) {
    findings.push({
      severity: 'low',
      area: 'storage',
      issue: `Only ${bytesToGb(free)} GB of ${bytesToGb(total)} GB free`,
      detail: 'Low storage blocks OS and app updates.',
    });
  }
  for (const action of detail?.deviceActionResults ?? []) {
    if (action.actionState !== 'pending') continue;
    findings.push({
      severity: ['wipe', 'retire', 'delete', 'fullScan'].includes(action.actionName ?? '') ? 'high' : 'info',
      area: 'actions',
      issue: `Remote action "${action.actionName}" is pending since ${action.startDateTime ?? 'unknown'}`,
      next: `tenantctl intune devices actions "${name}"`,
    });
  }
  return findings;
}

function configFindings(policies: PolicyStateView[], name: string): Finding[] {
  const broken = policies.filter((p) => BAD_STATES[p.state ?? '']);
  if (!broken.length) return [];
  return [
    {
      severity: 'medium',
      area: 'configuration',
      issue: `${broken.length} configuration profile${broken.length === 1 ? '' : 's'} in error or conflict`,
      detail: broken.map((p) => `${p.policy} (${p.state})`).join(', '),
      next: `tenantctl intune devices config-states "${name}"`,
    },
  ];
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'troubleshoot',
    name: 'device',
    summary: 'Why a device is non-compliant, blocked or misbehaving: compliance, sync, Entra state, profiles, sign-ins',
    description:
      'Start here for "my laptop/phone can\'t access email/Teams", "device not compliant" or "is this device ' +
      'healthy". Accepts a device name, serial number, Intune device id or Entra device id. Combines Intune ' +
      'compliance (with the failing settings in one sentence), last check-in, the Entra ID device object (disabled, ' +
      'missing, compliance not yet reported), configuration profile errors, encryption, storage, pending remote ' +
      "actions, and the primary user's failed sign-ins from this device. `findings` is ranked, each with a next " +
      'command; `unavailable` lists what could not be checked.',
    examples: ['tenantctl troubleshoot device LAPTOP-0042', 'tenantctl troubleshoot device C02XK1ABJHD5'],
    args: [
      {
        name: 'device',
        type: 'string',
        required: true,
        positional: true,
        description: 'Device name, serial number or id.',
      },
      { name: 'days', type: 'number', default: 7, description: 'Sign-in history to check for this device, in days.' },
    ],
    risk: 'read',
    permissions: permissionsOf(
      SOURCES,
      [
        'intune devices get',
        'intune devices compliance',
        'intune devices config-states',
        'intune devices actions',
        'entra devices get',
        'entra users signins',
      ],
      ['Intune Administrator', 'Global Reader'],
    ),
    async run(ctx, args) {
      const device = await resolveManagedDevice(ctx, args.device);
      const name = device.deviceName ?? device.id;
      const hasEntraId = Boolean(device.azureADDeviceId) && !NO_ENTRA_ID.test(device.azureADDeviceId ?? '');
      const sections = await settle({
        compliance: complianceStates(ctx, device.id),
        configuration: configStates(ctx, device.id),
        detail: deviceDetail(ctx, device.id),
        entra: hasEntraId ? entraObject(ctx, device.azureADDeviceId as string) : Promise.resolve(undefined),
        signins: deviceSigninFailures(ctx, device, args.days),
      });
      const compliance = isSettledError(sections.compliance) ? [] : sections.compliance;
      const entra = isSettledError(sections.entra) ? undefined : sections.entra;
      const findings = sortFindings([
        ...complianceFindings(device, compliance, name),
        ...syncFindings(device, name),
        ...entraFindings(device, entra, name),
        ...(isSettledError(sections.configuration) ? [] : configFindings(sections.configuration, name)),
        ...hardwareFindings(device, isSettledError(sections.detail) ? null : sections.detail, name),
        ...(isSettledError(sections.signins) || !sections.signins
          ? []
          : sections.signins.failureGroups
              .filter((g) => !g.benign)
              .map((g): Finding => ({
                severity: 'medium',
                area: 'sign-in',
                issue:
                  `${g.count} failed sign-in${g.count === 1 ? '' : 's'} from this device: ${g.errorCode} ${g.name ?? ''}`.trim(),
                detail: [g.meaning, g.failedPolicies.length ? `Blocked by: ${g.failedPolicies.join('; ')}` : null]
                  .filter(Boolean)
                  .join(' '),
                next: `tenantctl entra signin-error ${g.errorCode}`,
              }))),
      ]);
      const missing = unavailable(sections);
      return {
        device: deviceSummary(device),
        verdict: verdict(
          name,
          findings,
          missing.map((m) => m.section),
        ),
        findings,
        unavailable: missing,
        compliance: sections.compliance,
        configuration: sections.configuration,
        entra: entra ? shapeEntraDevice(entra) : (sections.entra ?? null),
        signins: sections.signins,
      };
    },
  }),
];
