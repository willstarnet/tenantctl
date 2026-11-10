import { odata } from '../../core/graph.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { daysSince, isSettledError, type Settled, settle } from '../../core/util.ts';
import { commands as defenderCommands } from '../defender/index.ts';
import { commands as entraCommands } from '../entra/index.ts';
import { commands as healthCommands, type HealthSummary, summarizeHealth } from '../health/index.ts';
import { commands as intuneCommands } from '../intune/index.ts';
import { type Finding, permissionsOf, sortFindings, unavailable, verdict } from './findings.ts';

const SOURCES = [...entraCommands, ...intuneCommands, ...defenderCommands, ...healthCommands];
/** Entra Connect syncs every 30 minutes; three hours without a sync means it has stopped. */
const DIRSYNC_STALE_HOURS = 3;
const SECURE_SCORE_LOW_PERCENT = 50;

interface ExpiringCredential {
  app: string;
  credentialType: string;
  daysLeft: number;
  expired: boolean;
}

interface SkuRow {
  skuPartNumber: string;
  available: number;
  enabled: number;
  warning: number;
  suspended: number;
  appliesTo: string | null;
  lowOnLicenses: boolean;
}

interface ComplianceSummary {
  compliantDeviceCount?: number;
  nonCompliantDeviceCount?: number;
  inGracePeriodCount?: number;
  errorDeviceCount?: number;
  conflictDeviceCount?: number;
  unknownDeviceCount?: number;
}

interface DirectorySync {
  tenant: string | null;
  enabled: boolean;
  lastSync: string | null;
  hoursSinceLastSync: number | null;
}

interface RiskSummary {
  count: number;
  truncated: boolean;
  users: string[];
}

interface IncidentSummary {
  count: number;
  truncated: boolean;
  incidents: { id: string; name: string | null; created: string | null }[];
}

interface ScoreSummary {
  percent: number;
  date: string | null;
}

function sourceCommand(id: string): AnyCommand {
  const found = SOURCES.find((c) => `${c.domain} ${c.name}` === id);
  if (!found) throw new Error(`troubleshoot tenant: unknown source command "${id}"`);
  return found;
}

/** Reuse the dedicated commands so this check and `entra apps expiring` / `entra licenses` never disagree. */
const appsExpiring = sourceCommand('entra apps expiring');
const licenses = sourceCommand('entra licenses');

async function expiringCredentials(ctx: Context): Promise<ExpiringCredential[]> {
  const result = (await appsExpiring.run(ctx, { days: 30, top: 200, all: false })) as ListResult<ExpiringCredential>;
  return result.items;
}

async function skuRows(ctx: Context): Promise<SkuRow[]> {
  const result = (await licenses.run(ctx, {})) as ListResult<SkuRow> | SkuRow[];
  return Array.isArray(result) ? result : result.items;
}

async function serviceHealth(ctx: Context) {
  const { items } = await ctx.graph.list<{ id: string; service: string; status: string }>(
    '/admin/serviceAnnouncement/healthOverviews',
    { limit: Number.POSITIVE_INFINITY, pageSize: false },
  );
  return summarizeHealth(items);
}

async function directorySync(ctx: Context): Promise<DirectorySync> {
  const { items } = await ctx.graph.list<{
    displayName: string;
    onPremisesSyncEnabled?: boolean | null;
    onPremisesLastSyncDateTime?: string | null;
  }>('/organization', {
    query: { $select: 'displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime' },
    limit: 1,
    pageSize: false,
  });
  const org = items[0];
  const last = org?.onPremisesLastSyncDateTime ?? null;
  return {
    tenant: org?.displayName ?? null,
    enabled: org?.onPremisesSyncEnabled === true,
    lastSync: last,
    hoursSinceLastSync: last ? Math.floor((Date.now() - Date.parse(last)) / 3_600_000) : null,
  };
}

async function deviceCompliance(ctx: Context): Promise<ComplianceSummary> {
  return ctx.graph.get<ComplianceSummary>('/deviceManagement/deviceCompliancePolicyDeviceStateSummary');
}

async function highRiskUsers(ctx: Context): Promise<RiskSummary> {
  const { items, truncated } = await ctx.graph.list<{ userPrincipalName?: string; riskLevel?: string }>(
    '/identityProtection/riskyUsers',
    { query: { $filter: `riskState eq 'atRisk' and riskLevel eq 'high'` }, limit: 50 },
  );
  return { count: items.length, truncated, users: items.map((u) => u.userPrincipalName ?? 'unknown').slice(0, 10) };
}

async function activeHighIncidents(ctx: Context): Promise<IncidentSummary> {
  const { items, truncated } = await ctx.graph.list<{ id: string; displayName?: string; createdDateTime?: string }>(
    '/security/incidents',
    {
      query: {
        $filter: `status eq ${odata('active')} and severity eq ${odata('high')}`,
        $select: 'id,displayName,createdDateTime',
      },
      limit: 50,
    },
  );
  return {
    count: items.length,
    truncated,
    incidents: items
      .slice(0, 10)
      .map((i) => ({ id: i.id, name: i.displayName ?? null, created: i.createdDateTime ?? null })),
  };
}

async function secureScore(ctx: Context): Promise<ScoreSummary | null> {
  const { items } = await ctx.graph.list<{ currentScore?: number; maxScore?: number; createdDateTime?: string }>(
    '/security/secureScores',
    { limit: 1 },
  );
  const latest = items[0];
  if (!latest?.maxScore) return null;
  return {
    percent: Math.round(((latest.currentScore ?? 0) / latest.maxScore) * 100),
    date: latest.createdDateTime ?? null,
  };
}

type Sections = Settled<{
  serviceHealth: HealthSummary;
  directorySync: DirectorySync;
  licenses: SkuRow[];
  appCredentials: ExpiringCredential[];
  deviceCompliance: ComplianceSummary;
  riskyUsers: RiskSummary;
  incidents: IncidentSummary;
  secureScore: ScoreSummary | null;
}>;

function loadSections(ctx: Context): Promise<Sections> {
  return settle({
    serviceHealth: serviceHealth(ctx),
    directorySync: directorySync(ctx),
    licenses: skuRows(ctx),
    appCredentials: expiringCredentials(ctx),
    deviceCompliance: deviceCompliance(ctx),
    riskyUsers: highRiskUsers(ctx),
    incidents: activeHighIncidents(ctx),
    secureScore: secureScore(ctx),
  });
}

function tenantFindings(s: Sections): Finding[] {
  const findings: Finding[] = [];
  if (!isSettledError(s.serviceHealth)) {
    for (const d of s.serviceHealth.degraded) {
      findings.push({
        severity: 'high',
        area: 'service health',
        issue: `Microsoft reports an issue with ${d.service} (${d.status})`,
        next: `m365 health issues --service "${d.service}" --active`,
      });
    }
  }
  if (
    !isSettledError(s.directorySync) &&
    s.directorySync.enabled &&
    (s.directorySync.hoursSinceLastSync ?? 0) > DIRSYNC_STALE_HOURS
  ) {
    findings.push({
      severity: 'high',
      area: 'directory sync',
      issue: `Entra Connect has not synced for ${s.directorySync.hoursSinceLastSync} hours`,
      detail:
        'New users, password changes and group changes from on-premises AD are not reaching Microsoft 365. Check the Entra Connect server.',
      next: 'm365 entra tenant',
    });
  }
  if (!isSettledError(s.licenses)) {
    for (const sku of s.licenses) {
      if (sku.appliesTo !== 'User' || sku.enabled === 0) continue;
      if (sku.suspended > 0 || sku.warning > 0) {
        findings.push({
          severity: 'high',
          area: 'licenses',
          issue: `${sku.skuPartNumber}: ${sku.warning} units in renewal grace period, ${sku.suspended} suspended`,
          detail: 'Renew the subscription before the units are disabled.',
          next: 'm365 entra licenses',
        });
      } else if (sku.lowOnLicenses) {
        findings.push({
          severity: sku.available <= 0 ? 'medium' : 'low',
          area: 'licenses',
          issue: `${sku.skuPartNumber}: ${sku.available} of ${sku.enabled} licenses left`,
          detail: 'New hires cannot be licensed once this runs out.',
          next: 'm365 entra licenses',
        });
      }
    }
  }
  if (!isSettledError(s.appCredentials) && s.appCredentials.length) {
    const expired = s.appCredentials.filter((c) => c.expired);
    const soon = s.appCredentials.filter((c) => !c.expired);
    if (expired.length) {
      findings.push({
        severity: 'medium',
        area: 'app credentials',
        issue: `${expired.length} app secret/certificate${expired.length === 1 ? ' has' : 's have'} expired`,
        detail: `Apps: ${[...new Set(expired.map((c) => c.app))].slice(0, 5).join(', ')}. Integrations using them fail to authenticate (7000222).`,
        next: 'm365 entra apps expiring',
      });
    }
    if (soon.length) {
      findings.push({
        severity: 'medium',
        area: 'app credentials',
        issue: `${soon.length} app secret/certificate${soon.length === 1 ? '' : 's'} expire within 30 days`,
        detail: `Soonest: ${soon[0]?.app} in ${soon[0]?.daysLeft} days.`,
        next: 'm365 entra apps expiring',
      });
    }
  }
  if (!isSettledError(s.deviceCompliance)) {
    const c = s.deviceCompliance;
    const bad = (c.nonCompliantDeviceCount ?? 0) + (c.errorDeviceCount ?? 0) + (c.conflictDeviceCount ?? 0);
    const total = bad + (c.compliantDeviceCount ?? 0) + (c.inGracePeriodCount ?? 0);
    if (bad > 0) {
      findings.push({
        severity: total && bad / total > 0.1 ? 'medium' : 'low',
        area: 'devices',
        issue: `${bad} of ${total} devices are non-compliant or in error`,
        next: 'm365 intune report noncompliant',
      });
    }
  }
  if (!isSettledError(s.riskyUsers) && s.riskyUsers.count) {
    findings.push({
      severity: 'high',
      area: 'identity risk',
      issue: `${s.riskyUsers.count}${s.riskyUsers.truncated ? '+' : ''} user${s.riskyUsers.count === 1 ? ' is' : 's are'} at high risk`,
      detail: `e.g. ${s.riskyUsers.users.slice(0, 5).join(', ')}`,
      next: 'm365 entra risky-users --level high --state atRisk',
    });
  }
  if (!isSettledError(s.incidents) && s.incidents.count) {
    findings.push({
      severity: 'high',
      area: 'security',
      issue: `${s.incidents.count}${s.incidents.truncated ? '+' : ''} active high-severity Defender incident${s.incidents.count === 1 ? '' : 's'}`,
      detail: s.incidents.incidents
        .map((i) => i.name)
        .filter(Boolean)
        .slice(0, 3)
        .join('; '),
      next: 'm365 defender incidents list --status active --severity high',
    });
  }
  if (!isSettledError(s.secureScore) && s.secureScore && s.secureScore.percent < SECURE_SCORE_LOW_PERCENT) {
    findings.push({
      severity: 'low',
      area: 'security',
      issue: `Secure Score is ${s.secureScore.percent}%`,
      next: 'm365 defender secure-score',
    });
  }
  return findings;
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'troubleshoot',
    name: 'tenant',
    summary:
      'Tenant health check: Microsoft outages, directory sync, licenses, expiring app secrets, devices, risk, incidents',
    description:
      'Use for "is something wrong with our Microsoft 365?", "why are several users affected?", or a daily/weekly ' +
      'admin check. Each area is checked in parallel and degrades independently; `unavailable` lists what could ' +
      'not be checked (e.g. risk needs Entra ID P2, incidents need Defender XDR). `findings` is ranked with the ' +
      'command to drill into each.',
    examples: ['m365 troubleshoot tenant'],
    args: [],
    risk: 'read',
    permissions: permissionsOf(
      SOURCES,
      [
        'health overview',
        'entra tenant',
        'entra licenses',
        'entra apps expiring',
        'intune compliance list',
        'entra risky-users',
        'defender incidents list',
        'defender secure-score',
      ],
      ['Global Reader'],
    ),
    async run(ctx) {
      const sections = await loadSections(ctx);
      const findings = sortFindings(tenantFindings(sections));
      const missing = unavailable(sections);
      return {
        verdict: verdict(
          'the tenant',
          findings,
          missing.map((m) => m.section),
        ),
        findings,
        unavailable: missing,
        ...sections,
        checkedAt: new Date().toISOString(),
        directorySyncAgeDays: isSettledError(sections.directorySync)
          ? undefined
          : daysSince(sections.directorySync.lastSync),
      };
    },
  }),
];
