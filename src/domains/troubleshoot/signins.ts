import { odata } from '../../core/graph.ts';
import type { Context } from '../../core/types.ts';
import { signinErrorInfo, signinErrorLookupUrl } from '../../core/signinErrors.ts';
import { daysAgo } from '../../core/util.ts';
import { type SignIn, shapeSignin } from '../entra/users.ts';
import type { Finding } from './findings.ts';

/** The sign-in fields troubleshooting reads beyond `shapeSignin`: which Conditional Access policy failed and why. */
export interface RawSignIn extends SignIn {
  appliedConditionalAccessPolicies?: {
    displayName?: string | null;
    result?: string | null;
    enforcedGrantControls?: string[] | null;
  }[];
}

export interface FailureGroup {
  errorCode: number;
  name: string | null;
  meaning: string | null;
  action: string | null;
  benign: boolean;
  count: number;
  firstSeen: string;
  lastSeen: string;
  apps: string[];
  ipAddresses: string[];
  locations: string[];
  devices: string[];
  /** Conditional Access policies that returned `failure`, with the controls they demanded. */
  failedPolicies: string[];
  failureReason: string | null;
  /** Correlation id of the latest failure: quote it in a Microsoft support case. */
  latestCorrelationId: string | null;
  lookup?: string;
}

export interface SigninAnalysis {
  window: string;
  total: number;
  failures: number;
  lastSuccess: string | null;
  /** The most recent attempt, so the agent can tell "still failing" from "failed earlier, fine now". */
  latest: Record<string, unknown> | null;
  failureGroups: FailureGroup[];
  truncated: boolean;
}

/** Sign-ins are capped per call; enough to cover a week of a normal user's attempts. */
const MAX_SIGNINS = 300;

/** Device-related codes: the device, not the user, is the problem. */
const DEVICE_CODES: Record<number, true> = { 53000: true, 53001: true, 50097: true, 135011: true };
/** Conditional Access codes: the per-policy breakdown in `troubleshoot signin` says which policy and why. */
const CA_CODES: Record<number, true> = { 53003: true, 50131: true, 530032: true, 53002: true };

export async function fetchSignins(ctx: Context, userId: string, days: number, app?: string) {
  const filters = [`userId eq ${odata(userId)}`, `createdDateTime ge ${daysAgo(days)}`];
  if (app) filters.push(`startswith(appDisplayName, ${odata(app)})`);
  return ctx.graph.list<RawSignIn>('/auditLogs/signIns', {
    query: { $filter: filters.join(' and ') },
    limit: MAX_SIGNINS,
  });
}

function distinct(values: (string | null | undefined)[], max = 5): string[] {
  return [...new Set(values.filter((v): v is string => Boolean(v)))].slice(0, max);
}

/** Group failed sign-ins by error code, newest group first, each code explained. */
export function analyzeSignins(signins: RawSignIn[], days: number, truncated: boolean): SigninAnalysis {
  const sorted = [...signins].sort((a, b) => b.createdDateTime.localeCompare(a.createdDateTime));
  const byCode = new Map<number, RawSignIn[]>();
  for (const signin of sorted) {
    const code = signin.status?.errorCode ?? 0;
    if (code === 0) continue;
    byCode.set(code, [...(byCode.get(code) ?? []), signin]);
  }
  const failureGroups = [...byCode.entries()].map(([code, list]): FailureGroup => {
    const info = signinErrorInfo(code);
    const newest = list[0] as RawSignIn;
    return {
      errorCode: code,
      name: info?.name ?? null,
      meaning: info?.meaning ?? null,
      action: info?.action ?? null,
      benign: info?.benign ?? false,
      count: list.length,
      firstSeen: (list.at(-1) as RawSignIn).createdDateTime,
      lastSeen: newest.createdDateTime,
      apps: distinct(list.map((s) => s.appDisplayName)),
      ipAddresses: distinct(list.map((s) => s.ipAddress)),
      locations: distinct(list.map((s) => [s.location?.city, s.location?.countryOrRegion].filter(Boolean).join(', '))),
      devices: distinct(list.map((s) => s.deviceDetail?.displayName)),
      failedPolicies: distinct(
        list.flatMap((s) =>
          (s.appliedConditionalAccessPolicies ?? [])
            .filter((p) => p.result === 'failure')
            .map(
              (p) =>
                `${p.displayName ?? 'unnamed policy'}${p.enforcedGrantControls?.length ? ` (requires ${p.enforcedGrantControls.join(', ')})` : ''}`,
            ),
        ),
        10,
      ),
      failureReason: newest.status?.failureReason ?? null,
      latestCorrelationId: newest.correlationId ?? null,
      ...(info ? {} : { lookup: signinErrorLookupUrl(code) }),
    };
  });
  failureGroups.sort((a, b) => b.lastSeen.localeCompare(a.lastSeen));
  const latest = sorted[0];
  return {
    window: `last ${days} day${days === 1 ? '' : 's'}`,
    total: sorted.length,
    failures: sorted.length - sorted.filter((s) => (s.status?.errorCode ?? 0) === 0).length,
    lastSuccess: sorted.find((s) => (s.status?.errorCode ?? 0) === 0)?.createdDateTime ?? null,
    latest: latest ? shapeSignin(latest) : null,
    failureGroups,
    truncated,
  };
}

/**
 * Turn failure groups into findings. A code whose latest failure is newer than the last success is "still
 * failing" (high); one that has since been followed by a success is history (low). Benign interrupts are skipped.
 */
export function signinFindings(analysis: SigninAnalysis, upn: string): Finding[] {
  const findings: Finding[] = [];
  for (const group of analysis.failureGroups) {
    if (group.benign) continue;
    const ongoing = !analysis.lastSuccess || group.lastSeen > analysis.lastSuccess;
    const spray = group.errorCode === 50126 && group.ipAddresses.length >= 3 && group.count >= 10;
    let next = `tenantctl entra signin-error ${group.errorCode}`;
    if (DEVICE_CODES[group.errorCode] && group.devices[0]) next = `tenantctl troubleshoot device "${group.devices[0]}"`;
    else if (CA_CODES[group.errorCode]) next = `tenantctl troubleshoot signin ${upn}`;
    else if (spray) next = `tenantctl entra risk-detections --user ${upn}`;
    findings.push({
      severity: spray || ongoing ? 'high' : 'low',
      area: 'sign-in',
      issue:
        `${group.count} sign-in failure${group.count === 1 ? '' : 's'} with ${group.errorCode} ${group.name ?? ''}`.trim() +
        (ongoing ? ' (still failing)' : ' (succeeded since)'),
      detail: [
        group.meaning,
        group.failedPolicies.length ? `Blocked by: ${group.failedPolicies.join('; ')}` : null,
        spray ? `Failures from ${group.ipAddresses.length}+ IPs: possible password spray.` : null,
        group.action ? `Fix: ${group.action}` : null,
      ]
        .filter(Boolean)
        .join(' '),
      next,
    });
  }
  return findings;
}
