import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { pickOne } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact, isGuid, settle } from '../../core/util.ts';
import { resolveServicePrincipal, roleNames, shapeServicePrincipal } from './directory.ts';

/**
 * Governance and tenant-relationship views: app consent review and the illicit-consent response (revoke, disable),
 * access reviews, GDAP partner relationships and subscription lifecycle.
 */

const ALL = Number.POSITIVE_INFINITY;
const DAY_MS = 86_400_000;
const NIL_GUID = '00000000-0000-0000-0000-000000000000';

function daysUntil(iso: string | null | undefined, now = Date.now()): number | null {
  if (!iso) return null;
  const time = Date.parse(iso);
  return Number.isNaN(time) ? null : Math.floor((time - now) / DAY_MS);
}

// ---------------------------------------------------------------------------------------------------------------
// App consents
// ---------------------------------------------------------------------------------------------------------------

interface GraphPermissionGrant {
  id: string;
  clientId: string;
  resourceId: string;
  scope?: string | null;
  consentType: string;
  principalId?: string | null;
}

interface GraphAppRoleAssignment {
  id: string;
  appRoleId: string;
  resourceId: string;
  resourceDisplayName?: string | null;
  createdDateTime?: string | null;
}

interface ResourceServicePrincipal {
  id: string;
  displayName: string;
  appRoles?: { id: string; value?: string | null; displayName?: string | null }[];
}

export interface DelegatedGrant {
  grantId: string;
  resource: string;
  /** admin = consented for all users (AllPrincipals); user = one user's own consent (Principal). */
  consent: 'admin' | 'user';
  principal?: { id: string; displayName?: string | null; userPrincipalName?: string | null };
  scopes: string[];
}

export interface ApplicationPermission {
  assignmentId: string;
  resource: string;
  permission: string;
  permissionName?: string;
  grantedDateTime: string | null;
}

type ResourceLookup = (id: string) => Promise<ResourceServicePrincipal | null>;

/** Resource service principals (Microsoft Graph, SharePoint...) fetched once per id. */
function resourceLookup(ctx: Context): ResourceLookup {
  const cache = new Map<string, Promise<ResourceServicePrincipal | null>>();
  return (id) => {
    let pending = cache.get(id);
    if (!pending) {
      pending = ctx.graph
        .get<ResourceServicePrincipal>(`/servicePrincipals/${seg(id)}`, {
          query: { $select: 'id,displayName,appRoles' },
        })
        .catch(() => null);
      cache.set(id, pending);
    }
    return pending;
  };
}

async function principalNames(
  ctx: Context,
  ids: string[],
): Promise<Record<string, { displayName?: string | null; userPrincipalName?: string | null }>> {
  const names: Record<string, { displayName?: string | null; userPrincipalName?: string | null }> = {};
  for (let start = 0; start < ids.length; start += 1000) {
    const found = await ctx.graph.post<{
      value?: { id: string; displayName?: string | null; userPrincipalName?: string | null }[];
    }>('/directoryObjects/getByIds', { ids: ids.slice(start, start + 1000), types: ['user'] });
    for (const o of found?.value ?? [])
      names[o.id] = { displayName: o.displayName, userPrincipalName: o.userPrincipalName };
  }
  return names;
}

/** Delegated permission grants the app holds (`oauth2PermissionGrants` where clientId is the app's SP). */
export async function loadDelegatedGrants(
  ctx: Context,
  spId: string,
  resource: ResourceLookup = resourceLookup(ctx),
): Promise<DelegatedGrant[]> {
  const { items } = await ctx.graph.list<GraphPermissionGrant>('/oauth2PermissionGrants', {
    query: { $filter: `clientId eq ${odata(spId)}` },
    limit: ALL,
    pageSize: false,
  });
  const principalIds = [...new Set(items.map((g) => g.principalId).filter((id): id is string => Boolean(id)))];
  const principals = principalIds.length ? await principalNames(ctx, principalIds) : {};
  return Promise.all(
    items.map(async (g) => {
      const admin = g.consentType === 'AllPrincipals';
      const principalId = g.principalId ?? undefined;
      return {
        grantId: g.id,
        resource: (await resource(g.resourceId))?.displayName ?? g.resourceId,
        consent: admin ? 'admin' : 'user',
        ...(admin || !principalId ? {} : { principal: { id: principalId, ...principals[principalId] } }),
        scopes: (g.scope ?? '').split(' ').filter(Boolean),
      } satisfies DelegatedGrant;
    }),
  );
}

/** Application (app-only) permissions: app role assignments on the app's SP, named from the resource's appRoles. */
export async function loadApplicationPermissions(
  ctx: Context,
  spId: string,
  resource: ResourceLookup = resourceLookup(ctx),
): Promise<ApplicationPermission[]> {
  const { items } = await ctx.graph.list<GraphAppRoleAssignment>(`/servicePrincipals/${seg(spId)}/appRoleAssignments`, {
    limit: ALL,
    pageSize: false,
  });
  return Promise.all(
    items.map(async (a) => {
      const res = await resource(a.resourceId);
      const role = res?.appRoles?.find((r) => r.id === a.appRoleId);
      return compact({
        assignmentId: a.id,
        resource: a.resourceDisplayName ?? res?.displayName ?? a.resourceId,
        permission: a.appRoleId === NIL_GUID ? 'default access' : (role?.value ?? a.appRoleId),
        permissionName: role?.displayName ?? undefined,
        grantedDateTime: a.createdDateTime ?? null,
      }) as ApplicationPermission;
    }),
  );
}

function describeGrant(g: DelegatedGrant): string {
  const who = g.consent === 'admin' ? 'all users' : (g.principal?.userPrincipalName ?? g.principal?.id ?? 'one user');
  return `${g.resource}: ${g.scopes.join(' ') || '(no scopes)'} [${who}]`;
}

const appArg = {
  name: 'app',
  type: 'string',
  required: true,
  positional: true,
  description: 'App name, appId or object id.',
} as const;

async function planRevoke(
  ctx: Context,
  args: {
    app: string;
    grant?: string[];
    allDelegated: boolean;
    appRoleAssignment?: string[];
    allApplication: boolean;
  },
) {
  const wantsDelegated = args.allDelegated || Boolean(args.grant?.length);
  const wantsApplication = args.allApplication || Boolean(args.appRoleAssignment?.length);
  if (!wantsDelegated && !wantsApplication) {
    throw new CliError(
      'USAGE',
      'Say what to revoke: --grant, --all-delegated, --app-role-assignment or --all-application',
      {
        hint: `List the app's grants and their ids first: m365 entra apps consents "${args.app}"`,
      },
    );
  }
  const sp = await resolveServicePrincipal(ctx, args.app);
  const resource = resourceLookup(ctx);
  const [grants, appPermissions] = await Promise.all([
    wantsDelegated ? loadDelegatedGrants(ctx, sp.id, resource) : Promise.resolve([]),
    wantsApplication ? loadApplicationPermissions(ctx, sp.id, resource) : Promise.resolve([]),
  ]);
  const pick = <T>(items: T[], ids: string[] | undefined, all: boolean, idOf: (item: T) => string, kind: string) => {
    if (all) return items;
    const wanted = ids ?? [];
    const unknown = wanted.filter((id) => !items.some((item) => idOf(item) === id));
    if (unknown.length) {
      throw new CliError('NOT_FOUND', `"${sp.displayName}" has no ${kind} with id ${unknown.join(', ')}`, {
        hint: `Valid ids: ${items.map(idOf).join(', ') || '(none)'}. See m365 entra apps consents "${sp.id}".`,
      });
    }
    return items.filter((item) => wanted.includes(idOf(item)));
  };
  const delegated = pick(grants, args.grant, args.allDelegated, (g) => g.grantId, 'delegated grant');
  const application = pick(
    appPermissions,
    args.appRoleAssignment,
    args.allApplication,
    (p) => p.assignmentId,
    'app role assignment',
  );
  if (!delegated.length && !application.length) {
    throw new CliError('NOT_FOUND', `"${sp.displayName}" holds no consents of the selected kind; nothing to revoke`, {
      hint: `m365 entra apps consents "${sp.id}" shows what it holds.`,
    });
  }
  return { sp, delegated, application };
}

const consentCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'apps consents',
    summary: 'Every delegated grant and application permission an app holds, with ids for revoking',
    description:
      'Use for consent reviews and illicit-consent-grant investigations. `delegated` lists OAuth2 permission grants: ' +
      'consent "admin" covers all users, "user" is one user\'s own consent (principal shown). `application` lists ' +
      'app-only permissions (no user needed; the riskiest kind, e.g. Mail.ReadWrite). Ids (`grantId`, ' +
      '`assignmentId`) feed `apps revoke-consent`. A section shows `{ error }` when a permission is missing.',
    examples: [
      'm365 entra apps consents "Contoso Mail Sync"',
      'm365 entra apps consents 1b2c3d4e-0000-0000-0000-000000000000',
    ],
    args: [appArg],
    risk: 'read',
    permissions: {
      delegated: ['Application.Read.All', 'Directory.Read.All'],
      application: ['Application.Read.All', 'Directory.Read.All'],
    },
    async run(ctx, args) {
      const sp = await resolveServicePrincipal(ctx, args.app);
      const resource = resourceLookup(ctx);
      const sections = await settle({
        delegated: loadDelegatedGrants(ctx, sp.id, resource),
        application: loadApplicationPermissions(ctx, sp.id, resource),
      });
      const delegated = Array.isArray(sections.delegated) ? sections.delegated : null;
      const application = Array.isArray(sections.application) ? sections.application : null;
      const total = (delegated?.length ?? 0) + (application?.length ?? 0);
      return compact({
        app: ctx.full ? sp : shapeServicePrincipal(sp),
        summary: compact({
          adminConsentedGrants: delegated?.filter((g) => g.consent === 'admin').length,
          userConsentedGrants: delegated?.filter((g) => g.consent === 'user').length,
          applicationPermissions: application?.length,
        }),
        delegated: sections.delegated,
        application: sections.application,
        next: total
          ? `To remove everything: m365 entra apps revoke-consent "${sp.id}" --all-delegated --all-application`
          : undefined,
      });
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'apps revoke-consent',
    summary: 'Revoke delegated grants and/or application permissions from an app (illicit consent response)',
    description:
      'Illicit-consent-grant playbook: 1) `apps consents <app>` to see what it holds; 2) this command to revoke ' +
      '(--all-delegated --all-application, or pick ids with --grant / --app-role-assignment); 3) `apps disable <app>` ' +
      'so it cannot get new tokens; 4) review what it did: `m365 entra signins --app <appId> --days 30` and ' +
      '`m365 entra audit --category ApplicationManagement`. Access tokens already issued stay valid until they ' +
      'expire (about an hour); revoke sessions of affected users if needed. Users can consent again unless the app ' +
      'is disabled or user consent is restricted. Preview lists every grant and permission that will be removed.',
    examples: [
      'm365 entra apps revoke-consent "Contoso Mail Sync" --all-delegated --all-application',
      'm365 entra apps revoke-consent "Contoso Mail Sync" --grant <grantId>',
    ],
    args: [
      appArg,
      { name: 'grant', type: 'string[]', description: 'Delegated grant id(s) (grantId from `apps consents`).' },
      { name: 'allDelegated', type: 'boolean', description: 'Revoke every delegated grant of the app.' },
      {
        name: 'appRoleAssignment',
        type: 'string[]',
        description: 'Application permission assignment id(s) (assignmentId from `apps consents`).',
      },
      { name: 'allApplication', type: 'boolean', description: 'Revoke every application permission of the app.' },
    ],
    risk: 'destructive',
    permissions: {
      delegated: [
        'Application.Read.All',
        'Directory.Read.All',
        'DelegatedPermissionGrant.ReadWrite.All',
        'AppRoleAssignment.ReadWrite.All',
      ],
      application: [
        'Application.Read.All',
        'Directory.Read.All',
        'DelegatedPermissionGrant.ReadWrite.All',
        'AppRoleAssignment.ReadWrite.All',
      ],
      roles: ['Cloud Application Administrator', 'Privileged Role Administrator'],
    },
    async preview(ctx, args) {
      const { sp, delegated, application } = await planRevoke(ctx, args);
      const parts = [
        delegated.length ? `${delegated.length} delegated grant(s) (${delegated.map(describeGrant).join('; ')})` : '',
        application.length
          ? `${application.length} application permission(s) (${application.map((p) => `${p.resource}: ${p.permission}`).join(', ')})`
          : '',
      ].filter(Boolean);
      return {
        action: `Revoke from app "${sp.displayName}": ${parts.join(' and ')}`,
        app: shapeServicePrincipal(sp),
        delegated,
        application,
      };
    },
    async run(ctx, args) {
      const { sp, delegated, application } = await planRevoke(ctx, args);
      const results = await ctx.graph.batch([
        ...delegated.map((g) => ({ method: 'DELETE', path: `/oauth2PermissionGrants/${seg(g.grantId)}` })),
        ...application.map((p) => ({
          method: 'DELETE',
          path: `/servicePrincipals/${seg(sp.id)}/appRoleAssignments/${seg(p.assignmentId)}`,
        })),
      ]);
      const status = (index: number) => {
        const r = results[index];
        if (!r?.error) return { status: 'revoked' };
        if (r.error.code === 'NOT_FOUND') return { status: 'alreadyRemoved' };
        return { status: 'failed', error: r.error.message };
      };
      return {
        app: { id: sp.id, appId: sp.appId, displayName: sp.displayName },
        delegated: delegated.map((g, i) => ({
          grantId: g.grantId,
          resource: g.resource,
          scopes: g.scopes,
          ...status(i),
        })),
        application: application.map((p, i) => ({
          assignmentId: p.assignmentId,
          resource: p.resource,
          permission: p.permission,
          ...status(delegated.length + i),
        })),
        next: [
          `m365 entra apps disable "${sp.id}"`,
          `m365 entra signins --app ${sp.appId} --days 30`,
          'm365 entra audit --category ApplicationManagement --days 30',
        ],
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'apps disable',
    summary: 'Stop an enterprise app from signing anyone in or getting tokens',
    description:
      'Sets the service principal accountEnabled false: no user can sign in to it and it gets no new tokens, ' +
      'app-only included. Consents and assignments are kept, so `apps enable` undoes it. Use after ' +
      '`apps revoke-consent` for a malicious app, or to pause a misbehaving integration. Tokens already issued ' +
      'stay valid until they expire. Needs Application.ReadWrite.All (Application.ReadWrite.OwnedBy only covers ' +
      "apps the caller owns, so it can't disable third-party apps).",
    examples: ['m365 entra apps disable "Contoso Mail Sync"'],
    args: [appArg],
    risk: 'destructive',
    permissions: {
      delegated: ['Application.ReadWrite.All'],
      application: ['Application.ReadWrite.All'],
      roles: ['Cloud Application Administrator', 'Application Administrator'],
    },
    async preview(ctx, args) {
      const sp = await resolveServicePrincipal(ctx, args.app);
      const app = shapeServicePrincipal(sp);
      return compact({
        action: `Disable enterprise app "${sp.displayName}" (appId ${sp.appId}): nobody can sign in to it`,
        app,
        note: sp.accountEnabled === false ? 'The app is already disabled.' : undefined,
        warning: app.isMicrosoftApp
          ? 'This is a Microsoft first-party app; disabling it can break Microsoft 365 features.'
          : undefined,
      });
    },
    async run(ctx, args) {
      const sp = await resolveServicePrincipal(ctx, args.app);
      await ctx.graph.patch(`/servicePrincipals/${seg(sp.id)}`, { accountEnabled: false });
      return {
        app: { ...shapeServicePrincipal(sp), accountEnabled: false },
        status: 'disabled',
        next: `Review its activity: m365 entra signins --app ${sp.appId} --days 30`,
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'apps enable',
    summary: 'Re-enable a disabled enterprise app',
    description: 'Sets the service principal accountEnabled true; users and the app can get tokens again.',
    examples: ['m365 entra apps enable "Contoso Mail Sync"'],
    args: [appArg],
    risk: 'write',
    permissions: {
      delegated: ['Application.ReadWrite.All'],
      application: ['Application.ReadWrite.All'],
      roles: ['Cloud Application Administrator', 'Application Administrator'],
    },
    async preview(ctx, args) {
      const sp = await resolveServicePrincipal(ctx, args.app);
      return compact({
        action: `Enable enterprise app "${sp.displayName}" (appId ${sp.appId})`,
        app: shapeServicePrincipal(sp),
        note: sp.accountEnabled === true ? 'The app is already enabled.' : undefined,
      });
    },
    async run(ctx, args) {
      const sp = await resolveServicePrincipal(ctx, args.app);
      await ctx.graph.patch(`/servicePrincipals/${seg(sp.id)}`, { accountEnabled: true });
      return { app: { ...shapeServicePrincipal(sp), accountEnabled: true }, status: 'enabled' };
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Access reviews
// ---------------------------------------------------------------------------------------------------------------

interface ReviewScope {
  '@odata.type'?: string;
  query?: string | null;
  principalScopes?: ReviewScope[];
  resourceScopes?: ReviewScope[];
}

interface ReviewerScope {
  query?: string | null;
  queryRoot?: string | null;
}

interface ReviewDefinition {
  id: string;
  displayName: string;
  status?: string | null;
  createdDateTime?: string | null;
  createdBy?: { displayName?: string | null; userPrincipalName?: string | null } | null;
  scope?: ReviewScope | null;
  instanceEnumerationScope?: ReviewScope | null;
  reviewers?: ReviewerScope[];
  settings?: {
    instanceDurationInDays?: number | null;
    autoApplyDecisionsEnabled?: boolean | null;
    defaultDecisionEnabled?: boolean | null;
    defaultDecision?: string | null;
    recurrence?: {
      pattern?: { type?: string | null; interval?: number | null } | null;
      range?: { startDate?: string | null; endDate?: string | null } | null;
    } | null;
  } | null;
}

interface ReviewInstance {
  id: string;
  status?: string | null;
  startDateTime?: string | null;
  endDateTime?: string | null;
}

interface ReviewDecision {
  id: string;
  decision?: string | null;
  recommendation?: string | null;
  justification?: string | null;
  reviewedDateTime?: string | null;
  applyResult?: string | null;
  reviewedBy?: { id?: string | null; displayName?: string | null; userPrincipalName?: string | null } | null;
  principal?: {
    '@odata.type'?: string;
    id?: string | null;
    displayName?: string | null;
    userPrincipalName?: string | null;
  } | null;
  resource?: { displayName?: string | null; type?: string | null } | null;
}

const DEFINITIONS = '/identityGovernance/accessReviews/definitions';
const DEFINITION_SELECT =
  'id,displayName,status,createdDateTime,createdBy,scope,instanceEnumerationScope,reviewers,settings';
const GUID_IN = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

/** One line describing whose access a review covers, from its scope query. */
export function reviewScopeSummary(scope: ReviewScope | null | undefined, perGroup?: ReviewScope | null): string {
  if (!scope) return 'unknown';
  if (scope.principalScopes || scope.resourceScopes) {
    const resources = (scope.resourceScopes ?? []).map((s) => reviewScopeSummary(s)).join(', ');
    return `access packages / resources: ${resources || 'unknown'}`;
  }
  const query = scope.query ?? '';
  const guests = /userType eq 'Guest'/i.test(query) ? ' (guests only)' : '';
  const id = GUID_IN.exec(query)?.[0];
  if (perGroup?.query) {
    const groups = /Unified/.test(perGroup.query) ? 'every Microsoft 365 group' : `groups matching ${perGroup.query}`;
    return `members of ${groups}${guests}`;
  }
  if (/roleDefinitions|roleAssignmentSchedule|roleEligibilitySchedule/i.test(query)) {
    return `Entra role${id ? ` ${id}` : ''} assignments${guests}`;
  }
  if (/^\/groups\//i.test(query)) return `members of group ${id ?? query}${guests}`;
  if (/^\/servicePrincipals\//i.test(query)) return `users assigned to app ${id ?? query}${guests}`;
  return `${query}${guests}`;
}

function reviewerSummary(reviewers: ReviewerScope[] | undefined): string[] {
  if (!reviewers?.length) return ['self-review'];
  return reviewers.map((r) => {
    const query = r.query ?? '';
    if (/manager/i.test(query)) return 'manager';
    if (/owners/i.test(query)) return 'group owners';
    const id = GUID_IN.exec(query)?.[0];
    if (/^\/users\//i.test(query)) return `user ${id ?? query}`;
    if (/^\/groups\//i.test(query)) return `members of group ${id ?? query}`;
    return query || 'unknown';
  });
}

function recurrenceSummary(def: ReviewDefinition): string {
  const pattern = def.settings?.recurrence?.pattern;
  if (!pattern?.type) return 'one-time';
  const unit = pattern.type === 'weekly' ? 'week' : pattern.type.toLowerCase().includes('month') ? 'month' : 'year';
  const n = pattern.interval ?? 1;
  if (unit === 'month' && n === 3) return 'quarterly';
  return n === 1 ? `every ${unit}` : `every ${n} ${unit}s`;
}

function shapeInstance(i: ReviewInstance | undefined, now = Date.now()) {
  if (!i) return null;
  return compact({
    id: i.id,
    status: i.status ?? null,
    startDateTime: i.startDateTime ?? null,
    endDateTime: i.endDateTime ?? null,
    daysLeft: i.status === 'InProgress' ? (daysUntil(i.endDateTime, now) ?? undefined) : undefined,
  });
}

/** Current (in progress) instance, else the most recent one that has started. */
function latestInstance(instances: ReviewInstance[], now = Date.now()): ReviewInstance | undefined {
  const started = instances
    .filter((i) => !i.startDateTime || Date.parse(i.startDateTime) <= now)
    .sort((a, b) => Date.parse(b.startDateTime ?? '') - Date.parse(a.startDateTime ?? ''));
  return started.find((i) => i.status === 'InProgress') ?? started[0];
}

function nextInstance(instances: ReviewInstance[], now = Date.now()): ReviewInstance | undefined {
  return instances
    .filter((i) => i.startDateTime && Date.parse(i.startDateTime) > now)
    .sort((a, b) => Date.parse(a.startDateTime ?? '') - Date.parse(b.startDateTime ?? ''))[0];
}

async function resolveReview(ctx: Context, ref: string): Promise<ReviewDefinition> {
  const value = ref.trim();
  if (isGuid(value)) {
    const def = await ctx.graph.tryGet<ReviewDefinition>(`${DEFINITIONS}/${seg(value)}`, {
      query: { $select: 'id,displayName,status' },
    });
    if (!def)
      throw new CliError('NOT_FOUND', `No access review with id ${value}`, { hint: 'm365 entra access-reviews list' });
    return def;
  }
  const { items } = await ctx.graph.list<ReviewDefinition>(DEFINITIONS, {
    query: { $select: 'id,displayName,status' },
    limit: ALL,
  });
  const wanted = value.toLowerCase();
  return pickOne(
    items.filter((d) => d.displayName.toLowerCase().includes(wanted)),
    value,
    'access review',
    (d) => [d.displayName],
    (d) => ({ id: d.id, displayName: d.displayName, status: d.status }),
  );
}

const ACCESS_REVIEW_NOTE =
  'Access reviews need Microsoft Entra ID P2 or Entra ID Governance licenses; without them the list is empty or ' +
  'Graph answers 403.';

const accessReviewCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'access-reviews list',
    summary: 'Access review series: what they cover, who reviews, how often, current instance',
    description:
      `Use for "are we reviewing guest access / admin roles?". ${ACCESS_REVIEW_NOTE} \`scope\` summarizes whose ` +
      'access is reviewed (ids are groups, apps or roles; resolve with `entra groups get` / `entra apps get`). ' +
      '`currentInstance` is the running or most recent instance (`daysLeft` while in progress); `nextInstance` ' +
      'appears when Graph has already scheduled one. Next: `access-reviews decisions <id>`.',
    examples: ['m365 entra access-reviews list', 'm365 entra access-reviews list --active'],
    args: [{ name: 'active', type: 'boolean', description: 'Only reviews in progress now.' }, ...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: ['AccessReview.Read.All'], application: ['AccessReview.Read.All'] },
    async run(ctx, args): Promise<ListResult> {
      const result = await ctx.graph.list<ReviewDefinition>(DEFINITIONS, {
        query: { $select: DEFINITION_SELECT, $filter: args.active ? "status eq 'InProgress'" : undefined },
        limit: listLimit(args),
      });
      if (ctx.full) return result;
      const instances = await ctx.graph.batch<{ value?: ReviewInstance[] }>(
        result.items.map((d) => ({
          path: `${DEFINITIONS}/${seg(d.id)}/instances`,
          query: { $select: 'id,status,startDateTime,endDateTime' },
        })),
      );
      const now = Date.now();
      return {
        items: result.items.map((d, index) => {
          const batch = instances[index];
          const list = batch?.body?.value ?? [];
          return compact({
            id: d.id,
            displayName: d.displayName,
            status: d.status ?? null,
            scope: reviewScopeSummary(d.scope, d.instanceEnumerationScope),
            scopeQuery: d.scope?.query ?? undefined,
            reviewers: reviewerSummary(d.reviewers),
            recurrence: recurrenceSummary(d),
            durationDays: d.settings?.instanceDurationInDays ?? undefined,
            autoApply: d.settings?.autoApplyDecisionsEnabled ?? undefined,
            defaultDecision: d.settings?.defaultDecisionEnabled ? (d.settings.defaultDecision ?? undefined) : undefined,
            currentInstance: batch?.error
              ? { error: batch.error.message }
              : shapeInstance(latestInstance(list, now), now),
            nextInstance: shapeInstance(nextInstance(list, now), now) ?? undefined,
            createdBy: d.createdBy?.userPrincipalName ?? d.createdBy?.displayName ?? undefined,
          });
        }),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'access-reviews decisions',
    summary: 'Decisions in the current (or latest) instance of an access review',
    description:
      `Use for "who still hasn't been reviewed?" or "who was denied?". ${ACCESS_REVIEW_NOTE} Picks the in-progress ` +
      'instance, else the most recent one (--instance to choose). `decision` is Approve, Deny, DontKnow or ' +
      "NotReviewed; `recommendation` is Graph's suggestion from sign-in activity. --pending shows only NotReviewed. " +
      'Group-wide reviews of every Microsoft 365 group have one instance per group; pass --instance for others.',
    examples: [
      'm365 entra access-reviews decisions "Quarterly guest review" --pending',
      'm365 entra access-reviews decisions <reviewId> --all',
    ],
    args: [
      { name: 'reviewId', type: 'string', required: true, positional: true, description: 'Review id or name.' },
      { name: 'pending', type: 'boolean', description: 'Only decisions not made yet (NotReviewed).' },
      { name: 'instance', type: 'string', description: 'Instance id (default: current or latest).' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['AccessReview.Read.All'], application: ['AccessReview.Read.All'] },
    async run(ctx, args) {
      const review = await resolveReview(ctx, args.reviewId);
      const instancesPath = `${DEFINITIONS}/${seg(review.id)}/instances`;
      let instance: ReviewInstance | undefined;
      if (args.instance) {
        instance = await ctx.graph.get<ReviewInstance>(`${instancesPath}/${seg(args.instance)}`, {
          query: { $select: 'id,status,startDateTime,endDateTime' },
        });
      } else {
        const { items } = await ctx.graph.list<ReviewInstance>(instancesPath, {
          query: { $select: 'id,status,startDateTime,endDateTime' },
          limit: ALL,
        });
        instance = latestInstance(items);
      }
      if (!instance) {
        throw new CliError('NOT_FOUND', `Access review "${review.displayName}" has no started instance yet`, {
          hint: 'm365 entra access-reviews list shows when it starts.',
        });
      }
      const result = await ctx.graph.list<ReviewDecision>(`${instancesPath}/${seg(instance.id)}/decisions`, {
        query: { $filter: args.pending ? "decision eq 'NotReviewed'" : undefined },
        limit: listLimit(args),
      });
      const decisions = ctx.full
        ? result.items
        : result.items.map((d) => {
            const reviewed = d.reviewedBy?.id && d.reviewedBy.id !== NIL_GUID;
            return compact({
              principal: d.principal?.displayName ?? d.principal?.id ?? null,
              principalUpn: d.principal?.userPrincipalName || undefined,
              principalType:
                d.principal?.['@odata.type']?.replace('#microsoft.graph.', '').replace('Identity', '') ?? undefined,
              resource: d.resource?.displayName ?? undefined,
              decision: d.decision ?? null,
              recommendation: d.recommendation ?? null,
              reviewedBy: reviewed ? d.reviewedBy?.userPrincipalName || d.reviewedBy?.displayName || null : null,
              reviewedDateTime: reviewed ? (d.reviewedDateTime ?? null) : undefined,
              justification: d.justification || undefined,
              applyResult: d.applyResult && d.applyResult !== 'New' ? d.applyResult : undefined,
            });
          });
      const counts: Record<string, number> = {};
      for (const d of result.items) counts[d.decision ?? 'unknown'] = (counts[d.decision ?? 'unknown'] ?? 0) + 1;
      return compact({
        review: { id: review.id, displayName: review.displayName, status: review.status ?? null },
        instance: shapeInstance(instance),
        counts,
        decisions,
        truncated: result.truncated || undefined,
      });
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Partners (GDAP) and subscriptions
// ---------------------------------------------------------------------------------------------------------------

interface DelegatedAdminRelationship {
  id: string;
  displayName: string;
  status?: string | null;
  duration?: string | null;
  autoExtendDuration?: string | null;
  activatedDateTime?: string | null;
  endDateTime?: string | null;
  customer?: { displayName?: string | null; tenantId?: string | null } | null;
  accessDetails?: { unifiedRoles?: { roleDefinitionId: string }[] } | null;
}

/** ISO 8601 day/year durations as used by GDAP (`P730D`, `P2Y`, `PT0S`) to whole days; null when not parseable. */
function durationDays(value: string | null | undefined): number | null {
  const match = /^P(?:(\d+)Y)?(?:(\d+)M)?(?:(\d+)D)?(?:T.*)?$/.exec(value ?? '');
  if (!match) return null;
  return Number(match[1] ?? 0) * 365 + Number(match[2] ?? 0) * 30 + Number(match[3] ?? 0);
}

export function shapeRelationship(r: DelegatedAdminRelationship, names: Record<string, string>, now = Date.now()) {
  const autoExtend = (durationDays(r.autoExtendDuration) ?? 0) > 0;
  const daysLeft = daysUntil(r.endDateTime, now);
  const live = r.status === 'active' || r.status === 'expiring';
  return compact({
    id: r.id,
    displayName: r.displayName,
    customer: r.customer?.displayName ?? null,
    customerTenantId: r.customer?.tenantId ?? null,
    status: r.status ?? null,
    roles: (r.accessDetails?.unifiedRoles ?? []).map((u) => names[u.roleDefinitionId] ?? u.roleDefinitionId),
    durationDays: durationDays(r.duration) ?? undefined,
    activatedDateTime: r.activatedDateTime ?? null,
    endDateTime: r.endDateTime ?? null,
    daysLeft: live ? daysLeft : undefined,
    autoExtend,
    expiringSoon: live && !autoExtend && daysLeft !== null && daysLeft < 30,
  });
}

interface CompanySubscription {
  id: string;
  skuId?: string | null;
  skuPartNumber?: string | null;
  status?: string | null;
  totalLicenses?: number | null;
  isTrial?: boolean | null;
  createdDateTime?: string | null;
  nextLifecycleDateTime?: string | null;
  ownerType?: string | null;
}

/** Subscriptions within this many days of their next lifecycle step are flagged `renewalSoon`. */
export const RENEWAL_WINDOW_DAYS = 30;

export function shapeSubscription(s: CompanySubscription, now = Date.now()) {
  const days = daysUntil(s.nextLifecycleDateTime, now);
  return compact({
    skuPartNumber: s.skuPartNumber ?? null,
    skuId: s.skuId ?? undefined,
    status: s.status ?? null,
    totalLicenses: s.totalLicenses ?? null,
    isTrial: s.isTrial ?? false,
    createdDateTime: s.createdDateTime ?? null,
    nextLifecycleDateTime: s.nextLifecycleDateTime ?? null,
    daysToNextLifecycle: days,
    renewalSoon: s.status !== 'Deleted' && days !== null && days < RENEWAL_WINDOW_DAYS,
    ownerType: s.ownerType ?? undefined,
    id: s.id,
  });
}

const tenantRelationshipCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'partners',
    summary: 'GDAP relationships this partner tenant holds with customers: roles, expiry, auto-extend',
    description:
      'For MSPs (run in the partner tenant): which customer relationships exist, which Entra roles each grants, ' +
      'and when they expire. `daysLeft` counts to endDateTime; `expiringSoon` is true for live relationships ' +
      'ending within 30 days without auto-extend (renew in Partner Center). This Graph API covers the partner side ' +
      'only: in a customer tenant it returns nothing or 403, and Graph has no API listing which partners hold ' +
      'admin access to us (see Microsoft 365 admin center > Settings > Partner relationships). Role names need ' +
      'RoleManagement.Read.Directory; without it role ids are shown.',
    examples: ['m365 entra partners', 'm365 entra partners --active --all'],
    args: [{ name: 'active', type: 'boolean', description: 'Only active relationships.' }, ...LIST_ARGS],
    risk: 'read',
    permissions: {
      delegated: ['DelegatedAdminRelationship.Read.All', 'RoleManagement.Read.Directory'],
      application: ['DelegatedAdminRelationship.Read.All', 'RoleManagement.Read.Directory'],
    },
    async run(ctx, args): Promise<ListResult> {
      const relationships = ctx.graph
        .list<DelegatedAdminRelationship>('/tenantRelationships/delegatedAdminRelationships', {
          query: { $filter: args.active ? "status eq 'active'" : undefined },
          limit: listLimit(args),
          pageSize: 300,
        })
        .catch((error: unknown) => {
          // Graph answers 401 (not 403) with a valid token when the tenant is not a Microsoft partner tenant.
          if (error instanceof CliError && error.status === 401) {
            throw new CliError('FORBIDDEN', 'No GDAP relationships are visible from this tenant', {
              status: 401,
              hint:
                'GDAP relationships are listed only in a Microsoft partner tenant, for an admin with a partner role ' +
                '(e.g. Admin Agent). A customer tenant sees its partners in the Microsoft 365 admin center > ' +
                'Settings > Partner relationships.',
            });
          }
          throw error;
        });
      const [result, names] = await Promise.all([
        relationships,
        roleNames(ctx).catch((error: unknown) => {
          ctx.log(
            `Role names unavailable (${error instanceof Error ? error.message : String(error)}); showing role ids.`,
          );
          return {} as Record<string, string>;
        }),
      ]);
      if (ctx.full) return result;
      const now = Date.now();
      return { items: result.items.map((r) => shapeRelationship(r, names, now)), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'entra',
    name: 'subscriptions',
    summary: 'Commercial subscriptions: seats, trial, status and days until the next lifecycle step',
    description:
      'Use for renewals and "why are licenses about to stop working?". `status` Enabled -> Warning (expired, grace) ' +
      '-> Suspended -> Deleted; `nextLifecycleDateTime` is when it moves to the next state if not renewed, ' +
      `\`daysToNextLifecycle\` counts to it and \`renewalSoon\` is true under ${RENEWAL_WINDOW_DAYS} days. Auto-renewing ` +
      'subscriptions also show a date (the renewal). Seat usage per SKU: `m365 entra licenses`.',
    examples: ['m365 entra subscriptions', 'm365 entra subscriptions --renewal-soon'],
    args: [
      {
        name: 'renewalSoon',
        type: 'boolean',
        description: `Only subscriptions within ${RENEWAL_WINDOW_DAYS} days of their next step.`,
      },
    ],
    risk: 'read',
    permissions: { delegated: ['Organization.Read.All'], application: ['Organization.Read.All'] },
    async run(ctx, args): Promise<ListResult> {
      const result = await ctx.graph.list<CompanySubscription>('/directory/subscriptions', {
        limit: ALL,
        pageSize: false,
      });
      if (ctx.full) return result;
      const now = Date.now();
      const items = result.items
        .map((s) => shapeSubscription(s, now))
        .filter((s) => !args.renewalSoon || s.renewalSoon)
        .sort((a, b) => (a.daysToNextLifecycle ?? Infinity) - (b.daysToNextLifecycle ?? Infinity));
      return { items, truncated: false };
    },
  }),
];

export const commands: AnyCommand[] = [...consentCommands, ...accessReviewCommands, ...tenantRelationshipCommands];
