import { seg } from '../../core/graph.ts';
import { resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { isSettledError, settle } from '../../core/util.ts';
import { commands as entraCommands } from '../entra/index.ts';
import { type GraphUser, getAuthMethodsSummary, getUserOverview } from '../entra/users.ts';
import {
  getInboxRules,
  getMailboxForwarding,
  isInternal,
  listTenantDomains,
  type MailboxForwarding,
  shapeAutoReply,
  type ShapedInboxRule,
  commands as exchangeCommands,
} from '../exchange/index.ts';
import type { AutomaticRepliesSetting } from '../exchange/mailbox.ts';
import { DEVICE_SELECT, listDevices, primaryUserClause } from '../intune/common.ts';
import { commands as intuneCommands, type ManagedDevice, shapeManagedDevice } from '../intune/index.ts';
import { type Finding, permissionsOf, sortFindings, unavailable, verdict } from './findings.ts';
import { analyzeSignins, fetchSignins, signinFindings } from './signins.ts';

const SOURCES = [...entraCommands, ...intuneCommands, ...exchangeCommands];
const STALE_DEVICE_DAYS = 30;

const USER_ARG = {
  name: 'user',
  type: 'string',
  required: true,
  positional: true,
  description: 'User UPN, email, display name, object id, or "me".',
} as const;
const DAYS_ARG = {
  name: 'days',
  type: 'number',
  default: 7,
  description: 'Sign-in history to analyze, in days (max 30).',
} as const;

const USER_SELECT =
  'id,displayName,userPrincipalName,mail,accountEnabled,userType,onPremisesSyncEnabled,assignedLicenses,licenseAssignmentStates';

interface RiskyUser {
  riskLevel?: string | null;
  riskState?: string | null;
  riskDetail?: string | null;
  riskLastUpdatedDateTime?: string | null;
}

function accountFindings(user: GraphUser, skuNames: Record<string, string>): Finding[] {
  const upn = user.userPrincipalName;
  const findings: Finding[] = [];
  if (user.accountEnabled === false) {
    findings.push({
      severity: 'high',
      area: 'account',
      issue: 'Account is disabled',
      detail: 'Every sign-in fails with 50057 until it is re-enabled.',
      next: `tenantctl entra users enable ${upn}`,
    });
  }
  for (const state of user.licenseAssignmentStates ?? []) {
    if (!state.error || state.error === 'None') continue;
    findings.push({
      severity: 'high',
      area: 'licenses',
      issue: `License ${skuNames[state.skuId] ?? state.skuId} failed to assign: ${state.error}`,
      detail: state.assignedByGroup ? `Assigned through group ${state.assignedByGroup}.` : 'Assigned directly.',
      next: `tenantctl entra users licenses ${upn}`,
    });
  }
  if (!user.assignedLicenses?.length && user.userType !== 'Guest') {
    findings.push({
      severity: 'medium',
      area: 'licenses',
      issue: 'No licenses assigned',
      detail:
        'Without a license the user has no mailbox, Teams or Office apps. Pick a SKU with spare units from the license list.',
      next: 'tenantctl entra licenses',
    });
  }
  if (user.onPremisesSyncEnabled) {
    findings.push({
      severity: 'info',
      area: 'account',
      issue: 'Synced from on-premises Active Directory',
      detail: 'Password, enable/disable and profile changes must be made in AD, or the next sync reverts them.',
    });
  }
  return findings;
}

function riskFindings(risk: RiskyUser | null, upn: string): Finding[] {
  if (!risk?.riskState || !['atRisk', 'confirmedCompromised'].includes(risk.riskState)) return [];
  const high = risk.riskState === 'confirmedCompromised' || risk.riskLevel === 'high' || risk.riskLevel === 'medium';
  return [
    {
      severity: high ? 'high' : 'medium',
      area: 'risk',
      issue: `Entra ID Protection rates the user ${risk.riskLevel ?? 'at'} risk (${risk.riskState})`,
      detail: `Risk detail: ${risk.riskDetail ?? 'none'}; updated ${risk.riskLastUpdatedDateTime ?? 'unknown'}. Risk policies may block sign-in or force a password change.`,
      next: `tenantctl entra risk-detections --user ${upn}`,
    },
  ];
}

/** Reads only the two facts it needs, so it accepts both the compact and `--full` auth-methods summaries. */
function mfaFindings(mfa: { hasStrongMethod?: boolean; registration?: unknown }, upn: string): Finding[] {
  const registration = mfa.registration;
  const unregistered =
    typeof registration === 'object' && registration !== null && 'isMfaRegistered' in registration
      ? registration.isMfaRegistered === false
      : false;
  const none = mfa.hasStrongMethod === false || unregistered;
  if (!none) return [];
  return [
    {
      severity: 'medium',
      area: 'mfa',
      issue: 'No MFA method registered',
      detail:
        'Sign-ins that require MFA prompt for registration, or fail (50079/53004) if registration is blocked. A ' +
        'Temporary Access Pass lets the user register a method.',
      next: `tenantctl entra users tap ${upn}`,
    },
  ];
}

function deviceFindings(devices: ManagedDevice[]): Finding[] {
  const findings: Finding[] = [];
  for (const device of devices) {
    const shaped = shapeManagedDevice(device);
    const name = device.deviceName ?? device.id;
    if (device.complianceState === 'noncompliant') {
      findings.push({
        severity: 'medium',
        area: 'devices',
        issue: `${name} is not compliant`,
        detail: 'Conditional Access policies that require a compliant device block it (53000).',
        next: `tenantctl troubleshoot device "${name}"`,
      });
    } else if ((shaped.daysSinceLastSync ?? 0) > STALE_DEVICE_DAYS) {
      findings.push({
        severity: 'low',
        area: 'devices',
        issue: `${name} has not checked in with Intune for ${shaped.daysSinceLastSync} days`,
        next: `tenantctl troubleshoot device "${name}"`,
      });
    }
  }
  return findings;
}

function mailboxFindings(
  autoReply: { autoReplyOn: boolean | null } | null,
  rules: ShapedInboxRule[] | null,
  upn: string,
): Finding[] {
  const findings: Finding[] = [];
  for (const rule of rules ?? []) {
    if (!rule.enabled) continue;
    if (rule.suspicious || rule.forwardsExternally) {
      findings.push({
        severity: 'high',
        area: 'mailbox',
        issue: `Inbox rule "${rule.name ?? rule.id}" ${rule.forwardsExternally ? 'forwards mail outside the organization' : 'looks suspicious'}`,
        detail: `${rule.reason ?? ''} Attackers who take over a mailbox often add such rules.`.trim(),
        next: `tenantctl exchange rules list ${upn}`,
      });
    } else if (rule.deletes || rule.movesToFolder) {
      findings.push({
        severity: 'info',
        area: 'mailbox',
        issue: `Inbox rule "${rule.name ?? rule.id}" ${rule.deletes ? 'deletes' : `moves mail to ${rule.movesToFolder}`}`,
        detail: 'Explains "missing" mail that was in fact delivered.',
        next: `tenantctl exchange rules list ${upn}`,
      });
    }
  }
  if (autoReply?.autoReplyOn) {
    findings.push({
      severity: 'info',
      area: 'mailbox',
      issue: 'Automatic replies (out of office) are on',
      next: `tenantctl exchange auto-reply get ${upn}`,
    });
  }
  return findings;
}

interface ForwardingCheck extends MailboxForwarding {
  /** null when tenant domains could not be read. */
  forwardsExternally: boolean | null;
}

/**
 * Mailbox-level forwarding (set in OWA or by an admin) is invisible to Graph and a favourite persistence trick
 * after an account takeover, so it is checked through the Exchange PowerShell bridge.
 */
async function mailboxForwarding(ctx: Context, upn: string): Promise<ForwardingCheck> {
  const [forwarding, domains] = await Promise.all([
    getMailboxForwarding(ctx, upn),
    listTenantDomains(ctx).catch(() => null),
  ]);
  const target = forwarding.forwardingSmtpAddress;
  return { ...forwarding, forwardsExternally: target ? (domains ? !isInternal(target, domains) : null) : false };
}

function forwardingFindings(forwarding: ForwardingCheck | null, upn: string): Finding[] {
  const target = forwarding?.forwardingSmtpAddress ?? forwarding?.forwardingAddress;
  if (!forwarding || !target) return [];
  const copy = forwarding.deliverToMailboxAndForward ? 'keeps a copy' : 'does not keep a copy';
  return [
    {
      severity: forwarding.forwardsExternally === false ? 'info' : 'high',
      area: 'mailbox',
      issue: `Mailbox forwards all mail to ${target}${forwarding.forwardsExternally ? ' (outside the organization)' : ''}`,
      detail: `Set on the mailbox itself, not as an inbox rule; it ${copy}. Confirm the user or an admin set this on purpose.`,
      next: `tenantctl exchange forwarding ${upn}`,
    },
  ];
}

async function loadSkuNames(ctx: Context): Promise<Record<string, string>> {
  const { items } = await ctx.graph.list<{ skuId: string; skuPartNumber: string }>('/subscribedSkus', {
    query: { $select: 'skuId,skuPartNumber' },
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  return Object.fromEntries(items.map((s) => [s.skuId, s.skuPartNumber]));
}

async function mailboxRules(ctx: Context, userId: string) {
  const domains = await listTenantDomains(ctx).catch(() => null);
  return getInboxRules(ctx, userId, domains);
}

async function loadAutoReply(ctx: Context, userId: string) {
  const settings = await ctx.graph.get<{ automaticRepliesSetting?: AutomaticRepliesSetting }>(
    `/users/${seg(userId)}/mailboxSettings`,
  );
  return shapeAutoReply(settings.automaticRepliesSetting);
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'troubleshoot',
    name: 'user',
    summary:
      'Everything that can stop a user working: account, licenses, MFA, risk, sign-in failures, devices, mailbox',
    description:
      'Start here for any "X can\'t sign in / get email / use Teams" request. Checks run in parallel and each ' +
      'degrades on its own: `unavailable` lists sections that could not be checked (usually a missing permission ' +
      "or licence, e.g. sign-in logs need Entra ID P1, risk needs P2, another user's mailbox needs app-only auth, " +
      'mailbox-level forwarding needs `tenantctl auth login --service exchange` and pwsh). `findings` is ranked high > ' +
      'medium > low > info, each with the command to investigate or fix it; `verdict` is a one-line answer to ' +
      'relay to the user. Sign-in failures are grouped by error code and marked "still failing" when no successful ' +
      'sign-in followed them. Mailbox forwarding (inbox rules and the mailbox-level setting) outside the company is ' +
      'flagged high: it is the usual persistence trick after an account takeover.',
    examples: ['tenantctl troubleshoot user jane@contoso.com', 'tenantctl troubleshoot user "Jane Doe" --days 14'],
    args: [USER_ARG, DAYS_ARG],
    risk: 'read',
    permissions: permissionsOf(
      SOURCES,
      [
        'entra users get',
        'entra users licenses',
        'entra users auth-methods',
        'entra users signins',
        'entra risky-users',
        'intune devices list',
        'exchange auto-reply get',
        'exchange rules list',
        'exchange forwarding',
      ],
      ['Global Reader', 'Security Reader', 'Helpdesk Administrator'],
    ),
    async run(ctx, args) {
      const user = (await resolveUser(ctx, args.user, USER_SELECT)) as GraphUser;
      const upn = user.userPrincipalName;
      const sections = await settle({
        overview: getUserOverview(ctx, user.id),
        mfa: getAuthMethodsSummary(ctx, user.id),
        signins: fetchSignins(ctx, user.id, args.days),
        risk: ctx.graph.tryGet<RiskyUser>(`/identityProtection/riskyUsers/${seg(user.id)}`),
        devices: listDevices(ctx, [primaryUserClause(user)], 25, DEVICE_SELECT),
        skuNames: loadSkuNames(ctx),
        autoReply: loadAutoReply(ctx, user.id),
        inboxRules: mailboxRules(ctx, user.id),
        mailboxForwarding: mailboxForwarding(ctx, upn),
      });

      const signins = isSettledError(sections.signins)
        ? sections.signins
        : analyzeSignins(sections.signins.items, args.days, sections.signins.truncated);
      const devices = isSettledError(sections.devices) ? sections.devices : sections.devices.items;
      const findings = sortFindings([
        ...accountFindings(user, isSettledError(sections.skuNames) ? {} : sections.skuNames),
        ...(isSettledError(sections.risk) ? [] : riskFindings(sections.risk, upn)),
        ...(isSettledError(sections.mfa) ? [] : mfaFindings(sections.mfa, upn)),
        ...(isSettledError(signins) ? [] : signinFindings(signins, upn)),
        ...(isSettledError(devices) ? [] : deviceFindings(devices)),
        ...mailboxFindings(
          isSettledError(sections.autoReply) ? null : sections.autoReply,
          isSettledError(sections.inboxRules) ? null : sections.inboxRules,
          upn,
        ),
        ...forwardingFindings(isSettledError(sections.mailboxForwarding) ? null : sections.mailboxForwarding, upn),
      ]);
      if (!isSettledError(signins) && signins.total === 0) {
        findings.push({ severity: 'info', area: 'sign-in', issue: `No sign-ins in the last ${args.days} days` });
      }
      const missing = unavailable(sections);
      return {
        user: { id: user.id, displayName: user.displayName, userPrincipalName: upn },
        verdict: verdict(
          upn,
          findings,
          missing.map((m) => m.section),
        ),
        findings,
        unavailable: missing,
        account: sections.overview,
        mfa: sections.mfa,
        risk: sections.risk,
        signins,
        devices: isSettledError(devices) ? devices : devices.map(shapeManagedDevice),
        mailbox: {
          autoReply: sections.autoReply,
          inboxRules: sections.inboxRules,
          forwarding: sections.mailboxForwarding,
        },
      };
    },
  }),
  defineCommand({
    domain: 'troubleshoot',
    name: 'signin',
    summary:
      'Why a user cannot sign in: failures grouped by error code, the Conditional Access policy that blocked them',
    description:
      'Deeper than `troubleshoot user` for sign-in problems. Each failure group carries the error explained, the ' +
      'apps/IPs/locations/devices involved, the Conditional Access policies whose result was "failure" with the ' +
      'controls they required (MFA, compliant device...), and the latest correlation id for a Microsoft support ' +
      'case. `latest` is the most recent attempt; compare `lastSuccess` to tell "still failing" from "fixed". ' +
      'Requires Entra ID P1 for sign-in logs.',
    examples: [
      'tenantctl troubleshoot signin jane@contoso.com',
      'tenantctl troubleshoot signin jane --app Outlook --days 2',
    ],
    args: [
      USER_ARG,
      DAYS_ARG,
      {
        name: 'app',
        type: 'string',
        description: 'Only sign-ins to apps whose name starts with this, e.g. "Microsoft Teams".',
      },
    ],
    risk: 'read',
    permissions: permissionsOf(
      SOURCES,
      ['entra users get', 'entra users signins', 'entra users auth-methods'],
      ['Reports Reader', 'Security Reader', 'Global Reader'],
    ),
    async run(ctx, args) {
      const user = (await resolveUser(ctx, args.user, USER_SELECT)) as GraphUser;
      const upn = user.userPrincipalName;
      const sections = await settle({
        signins: fetchSignins(ctx, user.id, args.days, args.app),
        mfa: getAuthMethodsSummary(ctx, user.id),
      });
      if (isSettledError(sections.signins)) {
        return {
          user: upn,
          verdict: `Could not read sign-in logs: ${sections.signins.error.message}`,
          unavailable: unavailable(sections),
        };
      }
      const analysis = analyzeSignins(sections.signins.items, args.days, sections.signins.truncated);
      const findings = sortFindings([
        ...accountFindings(user, {}).filter((f) => f.area === 'account'),
        ...(isSettledError(sections.mfa) ? [] : mfaFindings(sections.mfa, upn)),
        ...signinFindings(analysis, upn),
      ]);
      if (analysis.total === 0) {
        findings.push({
          severity: 'info',
          area: 'sign-in',
          issue: `No sign-in attempts${args.app ? ` to ${args.app}` : ''} in the last ${args.days} days`,
          detail:
            'The attempt never reached Entra ID: wrong username, wrong tenant, or a client that never got that far. Logs can lag a few minutes.',
        });
      }
      const missing = unavailable(sections);
      return {
        user: {
          id: user.id,
          displayName: user.displayName,
          userPrincipalName: upn,
          accountEnabled: user.accountEnabled,
        },
        verdict: verdict(
          upn,
          findings,
          missing.map((m) => m.section),
        ),
        findings,
        unavailable: missing,
        ...analysis,
        mfa: sections.mfa,
      };
    },
  }),
];
