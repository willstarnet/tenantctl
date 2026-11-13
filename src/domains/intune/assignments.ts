import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { type ResolvedGroup, resolveGroup } from '../../core/resolve.ts';
import { type AnyCommand, type ArgSpec, type Context, defineCommand, type ListResult } from '../../core/types.ts';
import { compact, isSettledError, settle } from '../../core/util.ts';
import {
  APPS_PATH,
  attempt,
  COMPLIANCE_PATH,
  CONFIG_PATH,
  INTUNE_ADMIN_ROLES,
  INTUNE_READ_ROLES,
  odataType,
  type Policy,
  pickByName,
  READ_APPS,
  READ_CONFIG,
  READ_GROUPS,
  resolveCatalogItem,
  SETTINGS_CATALOG_PATH,
} from './common.ts';

/**
 * Intune `/assign` actions REPLACE the whole assignment list of an app or policy. Every assign/unassign here
 * reads the current assignments first, merges the change, and sends the full resulting list, so existing
 * assignments survive.
 *
 * Assignments are read and written through beta for every kind: the v1.0 assignment targets have no
 * `deviceAndAppManagementAssignmentFilterId/Type` (v1.0 deviceAndAppManagementAssignmentTarget has no
 * properties), so a v1.0 read-merge-write would silently strip the assignment filter from every existing
 * assignment. A live check (2026-11) confirmed v1.0 omits filters that beta returns.
 */

const WRITE_APPS = 'DeviceManagementApps.ReadWrite.All';
const WRITE_CONFIG = 'DeviceManagementConfiguration.ReadWrite.All';
const ALL = Number.POSITIVE_INFINITY;
export const DEFAULT_MAX_ITEMS = 500;

export type AssignableKind = 'app' | 'compliance' | 'config' | 'settingsCatalog';

interface KindSpec {
  path: string;
  /** `@odata.type` of one assignment in the `/assign` body. */
  assignmentType: string;
  /** Body property of the `/assign` action. */
  bodyKey: 'mobileAppAssignments' | 'assignments';
  label: string;
}

const KINDS: Record<AssignableKind, KindSpec> = {
  app: {
    path: APPS_PATH,
    assignmentType: '#microsoft.graph.mobileAppAssignment',
    bodyKey: 'mobileAppAssignments',
    label: 'app',
  },
  compliance: {
    path: COMPLIANCE_PATH,
    assignmentType: '#microsoft.graph.deviceCompliancePolicyAssignment',
    bodyKey: 'assignments',
    label: 'compliance policy',
  },
  config: {
    path: CONFIG_PATH,
    assignmentType: '#microsoft.graph.deviceConfigurationAssignment',
    bodyKey: 'assignments',
    label: 'configuration profile',
  },
  settingsCatalog: {
    path: SETTINGS_CATALOG_PATH,
    assignmentType: '#microsoft.graph.deviceManagementConfigurationPolicyAssignment',
    bodyKey: 'assignments',
    label: 'settings catalog policy',
  },
};

const TARGET_TYPES = {
  group: '#microsoft.graph.groupAssignmentTarget',
  exclude: '#microsoft.graph.exclusionGroupAssignmentTarget',
  allUsers: '#microsoft.graph.allLicensedUsersAssignmentTarget',
  allDevices: '#microsoft.graph.allDevicesAssignmentTarget',
} as const;

type TargetKind = keyof typeof TARGET_TYPES;

interface Target {
  '@odata.type'?: string;
  groupId?: string | null;
  deviceAndAppManagementAssignmentFilterId?: string | null;
  deviceAndAppManagementAssignmentFilterType?: string | null;
  [key: string]: unknown;
}

export interface Assignment {
  id?: string;
  intent?: string | null;
  target?: Target | null;
  [key: string]: unknown;
}

export interface Subject {
  kind: AssignableKind;
  id: string;
  name: string;
  /** App/policy platform as Graph names it (`windows10CompliancePolicy`, `win32LobApp`, `windows10`), when known. */
  platform?: string | null;
}

interface ResolvedTarget {
  kind: TargetKind;
  target: Target;
  group?: ResolvedGroup;
  memberCount?: number | null;
}

type Args = Record<string, unknown>;

/** `#microsoft.graph.groupAssignmentTarget` (or without `#`) -> `group`; unknown kinds keep their type name. */
export function targetKind(target: Target | null | undefined): string {
  const type = String(target?.['@odata.type'] ?? '').replace(/^#?microsoft\.graph\./, '');
  for (const [kind, full] of Object.entries(TARGET_TYPES)) {
    if (full.endsWith(`.${type}`)) return kind;
  }
  return type || 'unknown';
}

function sameTarget(a: Target | null | undefined, b: Target): boolean {
  if (targetKind(a) !== targetKind(b)) return false;
  return (a?.groupId ?? null) === (b.groupId ?? null);
}

/** Drop read-only OData annotations from an assignment Graph returned, keeping its type. */
function clean(spec: KindSpec, assignment: Assignment): Assignment {
  const out: Assignment = { '@odata.type': spec.assignmentType };
  for (const [key, value] of Object.entries(assignment)) {
    if (key.startsWith('@odata.') && key !== '@odata.type') continue;
    out[key] = value;
  }
  return out;
}

export async function listAssignments(ctx: Context, subject: Subject): Promise<Assignment[]> {
  const spec = KINDS[subject.kind];
  const { items } = await ctx.graph.list<Assignment>(`${spec.path}/${seg(subject.id)}/assignments`, {
    version: 'beta',
    limit: ALL,
    pageSize: false,
  });
  return items;
}

// ---------------------------------------------------------------------------------------------------------------
// Assignment filters (beta only: deviceManagement/assignmentFilters has no v1.0 equivalent)

export const FILTERS_PATH = '/deviceManagement/assignmentFilters';
const FILTER_SELECT = 'id,displayName,description,platform,rule,assignmentFilterManagementType,lastModifiedDateTime';
export type FilterMode = 'include' | 'exclude';
export const FILTER_MODES = ['include', 'exclude'] as const;

export interface AssignmentFilter {
  id: string;
  displayName?: string | null;
  description?: string | null;
  platform?: string | null;
  rule?: string | null;
  assignmentFilterManagementType?: string | null;
  lastModifiedDateTime?: string | null;
  [key: string]: unknown;
}

/** Every assignment filter (tenants have a handful); `select: false` returns the full objects. */
export async function listFilters(
  ctx: Context,
  options: { limit?: number; select?: boolean } = {},
): Promise<ListResult<AssignmentFilter>> {
  return ctx.graph.list<AssignmentFilter>(FILTERS_PATH, {
    version: 'beta',
    query: options.select === false ? {} : { $select: FILTER_SELECT },
    limit: options.limit ?? ALL,
  });
}

export function resolveFilter(filters: AssignmentFilter[], ref: string): AssignmentFilter {
  return pickByName(filters, ref, 'assignment filter', (f) => f.displayName);
}

const FAMILIES: [RegExp, string][] = [
  [/^(windows|win32|winget|microsoftstore|officesuite|editionupgrade|sharedpc)/i, 'windows'],
  [/^(ios|managedios)/i, 'ios'],
  [/^macos/i, 'macos'],
  [/^(android|aosp|managedandroid)/i, 'android'],
  [/^linux/i, 'linux'],
];

/**
 * Coarse platform (windows, ios, macos, android, linux) of a filter platform (`windows10AndLater`), a policy or
 * app `@odata.type` (`windows10CompliancePolicy`, `win32LobApp`) or a settings catalog `platforms` value. Null when
 * unknown or mixed, in which case no platform check is made.
 */
export function platformFamily(value: string | null | undefined): string | null {
  if (!value || value.includes(',')) return null;
  return FAMILIES.find(([pattern]) => pattern.test(value))?.[1] ?? null;
}

interface FilterRef {
  id: string;
  mode: string;
}

function filterOf(target: Target | null | undefined): FilterRef | null {
  const id = target?.deviceAndAppManagementAssignmentFilterId;
  const mode = target?.deviceAndAppManagementAssignmentFilterType;
  return id && mode && mode !== 'none' ? { id, mode } : null;
}

/** The target with `filter` set, or with no filter (`none`) when `filter` is null. */
function withFilter(target: Target, filter: FilterRef | null): Target {
  return {
    ...target,
    deviceAndAppManagementAssignmentFilterId: filter?.id ?? null,
    deviceAndAppManagementAssignmentFilterType: filter?.mode ?? 'none',
  };
}

type FilterNames = Record<string, string | null>;

const namesOf = (filters: AssignmentFilter[]): FilterNames =>
  Object.fromEntries(filters.map((f) => [f.id, f.displayName ?? null]));

/** Filter names for display; one list call, only when an assignment carries a filter. A failure leaves ids only. */
async function filterNames(ctx: Context, assignments: Assignment[]): Promise<FilterNames | undefined> {
  if (!assignments.some((a) => filterOf(a.target))) return {};
  const loaded = await attempt(listFilters(ctx));
  if ('error' in loaded) {
    ctx.log(`Assignment filter names unavailable: ${loaded.error.message}`);
    return undefined;
  }
  return namesOf(loaded.items);
}

/** `filter`, `filterMode`, `filterId` of a target; `filter` (the name) is omitted when names are unavailable. */
function filterFields(target: Target | null | undefined, filters: FilterNames | undefined) {
  const filter = filterOf(target);
  if (!filter) return {};
  return compact({
    filter: filters ? (filters[filter.id] ?? null) : undefined,
    filterMode: filter.mode,
    filterId: filter.id,
  });
}

// ---------------------------------------------------------------------------------------------------------------

/** Members of a group via `$count`; null when Graph will not count them. */
export async function memberCount(ctx: Context, groupId: string): Promise<number | null> {
  try {
    const r = await ctx.graph.get<{ '@odata.count'?: number }>(`/groups/${seg(groupId)}/members`, {
      query: { $count: true, $top: 1, $select: 'id' },
      advanced: true,
    });
    return r['@odata.count'] ?? null;
  } catch {
    return null;
  }
}

/**
 * Display names of the groups the assignments target, fetched through `$batch`; `known` names skip the lookup.
 * Unknown or unreadable groups map to null.
 */
export async function assignmentGroupNames(
  ctx: Context,
  assignments: Assignment[],
  known: Record<string, string> = {},
) {
  const names: Record<string, string | null> = { ...known };
  const ids = assignments.flatMap((a) => (typeof a.target?.groupId === 'string' ? [a.target.groupId] : []));
  const missing = [...new Set(ids)].filter((id) => !(id in names));
  if (!missing.length) return names;
  const results = await ctx.graph.batch<{ displayName?: string }>(
    missing.map((id) => ({ path: `/groups/${seg(id)}`, query: { $select: 'id,displayName' } })),
  );
  missing.forEach((id, index) => {
    const result = results[index];
    names[id] = result && !result.error ? (result.body?.displayName ?? null) : null;
  });
  return names;
}

/**
 * Compact assignment. `filters` maps filter ids to names: pass it to show `filter` by name; without it only
 * `filterMode` and `filterId` are shown.
 */
export function shapeAssignment(
  assignment: Assignment,
  names: Record<string, string | null>,
  filters?: Record<string, string | null>,
) {
  const target = assignment.target ?? {};
  const groupId = typeof target.groupId === 'string' ? target.groupId : undefined;
  return compact({
    target: targetKind(target),
    group: groupId ? (names[groupId] ?? null) : undefined,
    groupId,
    intent: assignment.intent ?? undefined,
    ...filterFields(target, filters),
  });
}

/** Shape assignment lists with group and filter names, looked up once for all of them. */
export async function shapeAssignmentLists(
  ctx: Context,
  lists: Assignment[][],
  targets: ResolvedTarget[] = [],
  filters?: AssignmentFilter[],
) {
  const known: Record<string, string> = {};
  for (const t of targets) if (t.group) known[t.group.id] = t.group.displayName;
  const [names, filterMap] = await Promise.all([
    assignmentGroupNames(ctx, lists.flat(), known),
    filters ? namesOf(filters) : filterNames(ctx, lists.flat()),
  ]);
  return lists.map((list) => list.map((a) => shapeAssignment(a, names, filterMap)));
}

async function resolveApp(ctx: Context, ref: string): Promise<Subject> {
  const app = await resolveCatalogItem(ctx, APPS_PATH, ref, 'app');
  return { kind: 'app', id: app.id, name: app.displayName ?? app.id, platform: odataType(app) };
}

async function resolveCompliance(ctx: Context, ref: string): Promise<Subject> {
  const policy = await resolveCatalogItem(ctx, COMPLIANCE_PATH, ref, 'compliance policy');
  return { kind: 'compliance', id: policy.id, name: policy.displayName ?? policy.id, platform: odataType(policy) };
}

async function listSettingsCatalog(ctx: Context): Promise<Subject[]> {
  const { items } = await ctx.graph.list<Policy>(SETTINGS_CATALOG_PATH, {
    version: 'beta',
    query: { $select: 'id,name,platforms' },
    limit: ALL,
  });
  return items.map((p) => ({
    kind: 'settingsCatalog',
    id: p.id,
    name: p.name ?? p.id,
    platform: typeof p.platforms === 'string' ? p.platforms : null,
  }));
}

async function listClassicProfiles(ctx: Context): Promise<Subject[]> {
  const { items } = await ctx.graph.list<Policy>(CONFIG_PATH, { query: { $select: 'id,displayName' }, limit: ALL });
  return items.map((p) => ({ kind: 'config', id: p.id, name: p.displayName ?? p.id, platform: odataType(p) }));
}

/** A configuration profile by name or id: classic profiles (v1.0) and settings catalog policies (beta) together. */
export async function resolveConfigProfile(ctx: Context, ref: string): Promise<Subject> {
  const [classic, catalog] = await Promise.all([listClassicProfiles(ctx), attempt(listSettingsCatalog(ctx))]);
  if ('error' in catalog) ctx.log(`Settings catalog policies unavailable: ${catalog.error.message}`);
  const items = [...classic, ...('error' in catalog ? [] : catalog)];
  return pickByName(items, ref, 'configuration profile', (item) => item.name);
}

async function groupTarget(ctx: Context, ref: string, kind: 'group' | 'exclude'): Promise<ResolvedTarget> {
  const group = await resolveGroup(ctx, ref);
  return {
    kind,
    group,
    memberCount: await memberCount(ctx, group.id),
    target: { '@odata.type': TARGET_TYPES[kind], groupId: group.id },
  };
}

const allTarget = (kind: 'allUsers' | 'allDevices'): ResolvedTarget => ({
  kind,
  target: { '@odata.type': TARGET_TYPES[kind] },
});

async function resolveTargets(ctx: Context, args: Args, mode: 'assign' | 'unassign'): Promise<ResolvedTarget[]> {
  const include = [
    args.group ? 'group' : null,
    args.allUsers ? 'allUsers' : null,
    args.allDevices ? 'allDevices' : null,
  ].filter(Boolean);
  const chosen = include.length + (args.exclude ? 1 : 0);
  if (include.length > 1 || chosen === 0 || (mode === 'unassign' && chosen > 1)) {
    throw new CliError(
      'USAGE',
      mode === 'assign'
        ? 'Pass one of --group, --all-users or --all-devices (optionally with --exclude <group>)'
        : 'Pass exactly one of --group, --exclude, --all-users or --all-devices',
    );
  }
  const targets: ResolvedTarget[] = [];
  if (typeof args.group === 'string') targets.push(await groupTarget(ctx, args.group, 'group'));
  if (args.allUsers) targets.push(allTarget('allUsers'));
  if (args.allDevices) targets.push(allTarget('allDevices'));
  if (typeof args.exclude === 'string') targets.push(await groupTarget(ctx, args.exclude, 'exclude'));
  return targets;
}

function describeTarget(target: ResolvedTarget): string {
  if (target.kind === 'allUsers') return 'all licensed users';
  if (target.kind === 'allDevices') return 'all devices';
  const name = target.group?.displayName ?? target.target.groupId ?? 'unknown group';
  return typeof target.memberCount === 'number'
    ? `members of ${name} (${target.memberCount} members)`
    : `members of ${name}`;
}

function describeSubject(subject: Subject): string {
  return subject.kind === 'app' ? subject.name : `${KINDS[subject.kind].label} "${subject.name}"`;
}

/** `filter "Corporate devices" (include)`; falls back to the id when the name is unknown. */
function describeFilter(filter: FilterRef, filters: FilterNames | undefined): string {
  const name = filters?.[filter.id];
  return `filter ${name ? `"${name}"` : filter.id} (${filter.mode})`;
}

function assignAction(
  subject: Subject,
  targets: ResolvedTarget[],
  intent: string | undefined,
  filterPhrase = '',
): string {
  const include = targets.find((t) => t.kind !== 'exclude');
  const exclude = targets.find((t) => t.kind === 'exclude');
  const what = describeSubject(subject);
  if (!include && exclude) {
    return `Exclude ${describeTarget(exclude)} from ${what}${intent ? ` (${intent})` : ''}`;
  }
  const to = include ? describeTarget(include) : '';
  let sentence: string;
  if (subject.kind !== 'app') sentence = `Apply ${what} to ${to}`;
  else if (intent === 'available') sentence = `Make ${what} available in Company Portal to ${to}`;
  else if (intent === 'uninstall') sentence = `Uninstall ${what} from ${to}`;
  else sentence = `Install ${what} (required) on ${to}`;
  sentence += filterPhrase;
  return exclude ? `${sentence}, excluding ${describeTarget(exclude)}` : sentence;
}

/** Sentence for an existing assignment whose only change is its filter. */
function filterChangeAction(
  subject: Subject,
  targets: ResolvedTarget[],
  before: FilterRef | null,
  after: FilterRef | null,
  filters: FilterNames | undefined,
): string {
  const include = targets.find((t) => t.kind !== 'exclude') as ResolvedTarget;
  const exclude = targets.find((t) => t.kind === 'exclude');
  const of = `the assignment of ${describeSubject(subject)} to ${describeTarget(include)}`;
  let sentence: string;
  if (after && before) {
    sentence = `Change the filter on ${of} from ${describeFilter(before, filters)} to ${describeFilter(after, filters)}`;
  } else if (after) sentence = `Add ${describeFilter(after, filters)} to ${of}`;
  else sentence = `Remove ${describeFilter(before as FilterRef, filters)} from ${of}`;
  return exclude ? `${sentence}, and exclude ${describeTarget(exclude)}` : sentence;
}

function unassignAction(subject: Subject, target: ResolvedTarget, removed: Assignment[]): string {
  const what = describeSubject(subject);
  if (target.kind === 'exclude') return `Stop excluding ${describeTarget(target)} from ${what}`;
  const intents = [...new Set(removed.map((a) => a.intent).filter(Boolean))].join('/');
  const sentence = `Remove the ${intents ? `${intents} ` : ''}assignment of ${what} for ${describeTarget(target)}`;
  return subject.kind === 'app' && intents.includes('required')
    ? `${sentence} (already installed copies stay installed)`
    : sentence;
}

/** What `--filter` / `--clear-filter` ask for on the included target. */
type FilterRequest = { set: AssignmentFilter; mode: FilterMode } | { clear: true } | null;

interface Plan {
  subject: Subject;
  targets: ResolvedTarget[];
  current: Assignment[];
  next: Assignment[];
  changed: boolean;
  action: string;
  filterRequest?: FilterRequest;
  /** Filter on the included target before and after, when the plan touches filters. */
  filterBefore?: FilterRef | null;
  /** Every filter, when the plan had to read them (reused for display). */
  filters?: AssignmentFilter[];
}

/** Refuse filter arguments that cannot work, before anything is read. */
function checkFilterArgs(args: Args) {
  const hasInclude = Boolean(args.group || args.allUsers || args.allDevices);
  if (args.filterMode && !args.filter) {
    throw new CliError('USAGE', '--filter-mode needs --filter', {
      hint: 'Pass --filter <name|id> with --filter-mode, or drop --filter-mode (the default is include).',
    });
  }
  if (args.clearFilter && args.filter) {
    throw new CliError('USAGE', 'Pass either --filter or --clear-filter, not both', {
      hint: '--filter replaces any filter already on the assignment.',
    });
  }
  if ((args.filter || args.clearFilter) && !hasInclude) {
    throw new CliError('USAGE', 'Assignment filters apply to included targets, not to an --exclude group', {
      hint:
        'Pass --group, --all-users or --all-devices with --filter. To keep matching devices out, use ' +
        '--filter-mode exclude instead of an excluded group.',
    });
  }
}

/** Refuse a filter whose platform or management type cannot apply to the app/policy, when that is known cheaply. */
function checkFilterFits(subject: Subject, filter: AssignmentFilter) {
  const name = filter.displayName ?? filter.id;
  if (subject.kind !== 'app' && filter.assignmentFilterManagementType === 'apps') {
    throw new CliError(
      'USAGE',
      `Filter "${name}" is a managed-apps filter; ${describeSubject(subject)} needs a device filter`,
      {
        hint: 'List device filters with `tenantctl intune filters list` (managementType devices).',
      },
    );
  }
  const wanted = platformFamily(subject.platform);
  const offered = platformFamily(filter.platform);
  if (wanted && offered && wanted !== offered) {
    throw new CliError(
      'USAGE',
      `Filter "${name}" is for ${filter.platform} but ${describeSubject(subject)} is a ${wanted} ${KINDS[subject.kind].label}`,
      { hint: `Pick a ${wanted} filter: \`tenantctl intune filters list --platform ${wanted}\`.` },
    );
  }
}

async function planAssign(ctx: Context, subject: Subject, args: Args): Promise<Plan> {
  const spec = KINDS[subject.kind];
  const intent = subject.kind === 'app' ? String(args.intent) : undefined;
  if (intent === 'available' && args.allDevices) {
    throw new CliError('USAGE', 'Available apps can only be targeted at users, not --all-devices', {
      hint: 'Use --all-users or a user group for available apps, or --intent required for devices.',
    });
  }
  const wantsFilter = typeof args.filter === 'string';
  const [targets, current, loaded] = await Promise.all([
    resolveTargets(ctx, args, 'assign'),
    listAssignments(ctx, subject),
    // --filter must resolve; for --clear-filter the names are only for the action sentence.
    wantsFilter ? listFilters(ctx) : args.clearFilter ? attempt(listFilters(ctx)) : undefined,
  ]);
  const filters = loaded && !('error' in loaded) ? loaded.items : undefined;
  let filterRequest: FilterRequest = null;
  if (wantsFilter) {
    const filter = resolveFilter(filters ?? [], String(args.filter));
    checkFilterFits(subject, filter);
    filterRequest = { set: filter, mode: args.filterMode === 'exclude' ? 'exclude' : 'include' };
  } else if (args.clearFilter) filterRequest = { clear: true };
  const desired: FilterRef | null | undefined = !filterRequest
    ? undefined
    : 'set' in filterRequest
      ? { id: filterRequest.set.id, mode: filterRequest.mode }
      : null;

  const next = current.map((a) => clean(spec, a));
  let changed = false;
  let filterBefore: FilterRef | null = null;
  let onlyFilterChanged = false;
  for (const target of targets) {
    const filtered = target.kind !== 'exclude' && desired !== undefined;
    const index = next.findIndex((a) => sameTarget(a.target, target.target));
    const existing = index >= 0 ? next[index] : undefined;
    if (!existing) {
      const newTarget = filtered && desired ? withFilter(target.target, desired) : target.target;
      next.push({ '@odata.type': spec.assignmentType, ...(intent ? { intent } : {}), target: newTarget });
      changed = true;
      continue;
    }
    let updated = existing;
    const intentChanged = Boolean(intent && existing.intent !== intent);
    // Intune allows one intent per group per app: switching intent replaces the existing assignment.
    if (intentChanged) updated = { ...updated, intent };
    let filterChanged = false;
    if (filtered) {
      const before = filterOf(existing.target);
      filterBefore = before;
      filterChanged = desired ? before?.id !== desired.id || before?.mode !== desired.mode : before !== null;
      if (filterChanged)
        updated = { ...updated, target: withFilter(existing.target ?? target.target, desired ?? null) };
    }
    if (intentChanged || filterChanged) {
      next[index] = updated;
      changed = true;
    }
    if (filterChanged && !intentChanged) onlyFilterChanged = true;
  }

  const names = filters ? namesOf(filters) : undefined;
  let action: string;
  if (onlyFilterChanged) {
    action = filterChangeAction(subject, targets, filterBefore, desired ?? null, names);
  } else {
    let phrase = '';
    if (desired) phrase = ` with ${describeFilter(desired, names)}`;
    else if (desired === null && filterBefore) phrase = `, removing ${describeFilter(filterBefore, names)}`;
    action = assignAction(subject, targets, intent, phrase);
  }
  return { subject, targets, current, next, changed, action, filterRequest, filterBefore, filters };
}

async function planUnassign(ctx: Context, subject: Subject, args: Args): Promise<Plan> {
  const spec = KINDS[subject.kind];
  const intent = subject.kind === 'app' && typeof args.intent === 'string' ? args.intent : undefined;
  const [targets, current] = await Promise.all([resolveTargets(ctx, args, 'unassign'), listAssignments(ctx, subject)]);
  const target = targets[0] as ResolvedTarget;
  const matches = (a: Assignment) => sameTarget(a.target, target.target) && (!intent || a.intent === intent);
  const removed = current.filter(matches);
  if (!removed.length) {
    throw new CliError('NOT_FOUND', `${describeSubject(subject)} is not assigned to ${describeTarget(target)}`, {
      hint:
        subject.kind === 'app'
          ? `List its assignments with \`tenantctl intune apps assignments ${subject.id}\`.`
          : `Check what targets the group with \`tenantctl intune assignments ${target.group?.id ?? '<group>'}\`.`,
    });
  }
  const next = current.filter((a) => !matches(a)).map((a) => clean(spec, a));
  return { subject, targets, current, next, changed: true, action: unassignAction(subject, target, removed) };
}

function describeFilterRequest(plan: Plan) {
  const request = plan.filterRequest;
  if (!request) return undefined;
  if ('set' in request) {
    return compact({
      change: 'set',
      id: request.set.id,
      name: request.set.displayName ?? null,
      mode: request.mode,
      platform: request.set.platform ?? null,
      managementType: request.set.assignmentFilterManagementType ?? undefined,
      rule: request.set.rule ?? null,
    });
  }
  const before = plan.filterBefore;
  const filter = before ? plan.filters?.find((f) => f.id === before.id) : undefined;
  return compact({
    change: 'clear',
    id: before?.id ?? null,
    name: before ? (filter?.displayName ?? null) : undefined,
    mode: before?.mode,
    rule: filter ? (filter.rule ?? null) : undefined,
  });
}

async function describePlan(ctx: Context, plan: Plan) {
  const spec = KINDS[plan.subject.kind];
  const [current, result] = await shapeAssignmentLists(ctx, [plan.current, plan.next], plan.targets, plan.filters);
  const request = plan.filterRequest;
  return {
    action: plan.changed ? plan.action : `Already in place: ${plan.action}`,
    [plan.subject.kind === 'app' ? 'app' : 'policy']: {
      id: plan.subject.id,
      name: plan.subject.name,
      kind: plan.subject.kind,
    },
    targets: plan.targets.map((t) =>
      compact({
        target: t.kind,
        group: t.group?.displayName,
        groupId: t.group?.id,
        memberCount: t.group ? (t.memberCount ?? null) : undefined,
        filter: request && 'set' in request && t.kind !== 'exclude' ? (request.set.displayName ?? null) : undefined,
        filterMode: request && 'set' in request && t.kind !== 'exclude' ? request.mode : undefined,
      }),
    ),
    filter: describeFilterRequest(plan),
    changed: plan.changed,
    current,
    result,
    request: plan.changed
      ? { method: 'POST', path: `${spec.path}/${plan.subject.id}/assign`, version: 'beta' }
      : undefined,
  };
}

async function applyPlan(ctx: Context, plan: Plan, mode: 'assign' | 'unassign', next: string) {
  const spec = KINDS[plan.subject.kind];
  const view = await describePlan(ctx, plan);
  if (!plan.changed) return { status: 'unchanged', ...view, request: undefined };
  try {
    await ctx.graph.post(
      `${spec.path}/${seg(plan.subject.id)}/assign`,
      { [spec.bodyKey]: plan.next },
      { version: 'beta' },
    );
  } catch (error) {
    const request = plan.filterRequest;
    if (!(error instanceof CliError) || error.status !== 400 || !request || !('set' in request)) throw error;
    const name = request.set.displayName ?? request.set.id;
    throw new CliError(error.code, error.message, {
      status: error.status,
      graphCode: error.graphCode,
      requestId: error.requestId,
      details: error.details,
      hint:
        `Intune may not accept filter "${name}" (${request.set.platform ?? 'unknown platform'}, ` +
        `${request.set.assignmentFilterManagementType ?? 'devices'}) on this ${KINDS[plan.subject.kind].label}. ` +
        'Pick a filter for the same platform with `tenantctl intune filters list`.',
    });
  }
  return {
    status: mode === 'assign' ? 'assigned' : 'unassigned',
    action: view.action,
    ...(plan.subject.kind === 'app' ? { app: view.app } : { policy: view.policy }),
    targets: view.targets,
    ...(view.filter ? { filter: view.filter } : {}),
    assignments: view.result,
    next,
  };
}

const TARGET_ARGS: ArgSpec[] = [
  { name: 'group', type: 'string', description: 'Target group: name, mail or id.' },
  { name: 'allUsers', type: 'boolean', description: 'Target all licensed users.' },
  { name: 'allDevices', type: 'boolean', description: 'Target all Intune-managed devices.' },
];

const FILTER_ARGS: ArgSpec[] = [
  {
    name: 'filter',
    type: 'string',
    description: 'Assignment filter (name or id) for the included target; replaces any filter it already has.',
  },
  {
    name: 'filterMode',
    type: 'string',
    choices: FILTER_MODES,
    description: 'include (only matching devices get it, the default) or exclude (matching devices do not).',
  },
  {
    name: 'clearFilter',
    type: 'boolean',
    description: 'Remove the filter from the included target, so every member gets it.',
  },
];

const INTENTS = ['required', 'available', 'uninstall'] as const;

interface AreaSpec {
  area: 'apps' | 'compliance' | 'config';
  noun: string;
  ref: string;
  resolve: (ctx: Context, ref: string) => Promise<Subject>;
  permission: string;
  /** Extra read scopes beyond `permission` (filters and their names on app commands). */
  extra: string[];
  next: (subject: Subject) => string;
  description: string;
  examples: { assign: string[]; unassign: string[] };
}

const quoted = (value: string) => (/[\s"']/.test(value) ? `"${value.replace(/"/g, '\\"')}"` : value);

const AREAS: AreaSpec[] = [
  {
    area: 'apps',
    noun: 'app',
    ref: 'App name or id.',
    resolve: resolveApp,
    permission: WRITE_APPS,
    extra: [READ_CONFIG],
    next: (s) => `tenantctl intune apps status ${quoted(s.name)}`,
    description:
      'intent required installs the app on targeted devices/users, available offers it in Company Portal ' +
      '(users only: refused with --all-devices, and Intune ignores available for device groups), uninstall ' +
      'removes it. A group already assigned with another intent is switched to the new one. Assignment settings ' +
      '(delivery optimisation, notifications) take Intune defaults; adjust them in the admin center.',
    examples: {
      assign: [
        'tenantctl intune apps assign "Zoom Workplace" --intent required --group "Sales Laptops"',
        'tenantctl intune apps assign "Company Portal" --intent available --all-users --exclude "Kiosk Users"',
        'tenantctl intune apps assign "Zoom Workplace" --intent required --group "All Staff" --filter "Corporate devices"',
      ],
      unassign: ['tenantctl intune apps unassign "Zoom Workplace" --group "Sales Laptops"'],
    },
  },
  {
    area: 'compliance',
    noun: 'compliance policy',
    ref: 'Compliance policy name or id.',
    resolve: resolveCompliance,
    permission: WRITE_CONFIG,
    extra: [],
    next: (s) => `tenantctl intune compliance status ${quoted(s.name)}`,
    description:
      'Newly targeted devices are evaluated at their next check-in; non-compliant ones can lose access through ' +
      'Conditional Access after the grace period.',
    examples: {
      assign: [
        'tenantctl intune compliance assign "Windows baseline" --group "All Windows" --exclude "Kiosks"',
        'tenantctl intune compliance assign "Windows baseline" --group "All Windows" --clear-filter',
      ],
      unassign: ['tenantctl intune compliance unassign "Windows baseline" --exclude "Kiosks"'],
    },
  },
  {
    area: 'config',
    noun: 'configuration profile',
    ref: 'Profile name or id (classic profile or settings catalog policy).',
    resolve: resolveConfigProfile,
    permission: WRITE_CONFIG,
    extra: [],
    next: (s) =>
      s.kind === 'config' ? `tenantctl intune config status ${quoted(s.name)}` : `tenantctl intune assignments <group>`,
    description:
      'Works for classic profiles (deviceConfigurations) and settings catalog / endpoint security policies ' +
      '(configurationPolicies; Microsoft has not published a v1.0 equivalent).',
    examples: {
      assign: [
        'tenantctl intune config assign "Corporate Wi-Fi" --group "Sales Laptops"',
        'tenantctl intune config assign "Edge settings" --all-devices --filter "Corporate devices" --filter-mode include',
      ],
      unassign: ['tenantctl intune config unassign "Edge settings" --all-devices'],
    },
  },
];

const FILTER_DESCRIPTION =
  '--filter adds an assignment filter (default --filter-mode include) to the included target; if that target is ' +
  'already assigned, its filter is replaced in place. --clear-filter removes the filter from the included ' +
  'target (assigning it unfiltered if new). Without either, a new assignment has no filter and an existing ' +
  "one keeps its filter. Excluded groups cannot carry a filter. A filter whose platform differs from the item's " +
  'is refused. Assignments are read and written through the beta API because v1.0 does not return or accept ' +
  'filters (a v1.0 rewrite would drop them).';

function assignCommand(spec: AreaSpec): AnyCommand {
  const isApp = spec.area === 'apps';
  const intentArg: ArgSpec = {
    name: 'intent',
    type: 'string',
    required: true,
    choices: INTENTS,
    description: 'required (install), available (Company Portal, users only) or uninstall.',
  };
  const args: ArgSpec[] = [
    { name: 'item', type: 'string', required: true, positional: true, description: spec.ref },
    ...(isApp ? [intentArg] : []),
    ...TARGET_ARGS,
    { name: 'exclude', type: 'string', description: 'Exclude this group (name, mail or id).' },
    ...FILTER_ARGS,
  ];
  const plan = async (ctx: Context, args: Args) => {
    checkFilterArgs(args);
    return planAssign(ctx, await spec.resolve(ctx, String(args.item)), args);
  };
  const scopes = [spec.permission, ...spec.extra, READ_GROUPS];
  return defineCommand({
    domain: 'intune',
    name: `${spec.area} assign`,
    summary: `Assign a ${spec.noun} to a group, all users or all devices (keeps existing assignments)`,
    description:
      `Adds one target to the ${spec.noun}'s assignments; existing assignments are read first and kept, ` +
      `because Intune's assign action replaces the whole list. ${spec.description} ${FILTER_DESCRIPTION} The ` +
      'preview shows the target group with its member count, the filter with its rule, the current assignments ' +
      'and the resulting list.',
    examples: spec.examples.assign,
    args: args as readonly ArgSpec[],
    risk: 'destructive',
    permissions: { delegated: scopes, application: scopes, roles: INTUNE_ADMIN_ROLES },
    async preview(ctx, args) {
      return describePlan(ctx, await plan(ctx, args as unknown as Args));
    },
    async run(ctx, args) {
      const resolved = await plan(ctx, args as unknown as Args);
      return applyPlan(ctx, resolved, 'assign', spec.next(resolved.subject));
    },
  });
}

function unassignCommand(spec: AreaSpec): AnyCommand {
  const isApp = spec.area === 'apps';
  const args: ArgSpec[] = [
    { name: 'item', type: 'string', required: true, positional: true, description: spec.ref },
    ...TARGET_ARGS,
    { name: 'exclude', type: 'string', description: 'Remove the exclusion of this group instead.' },
    ...(isApp
      ? [
          {
            name: 'intent',
            type: 'string',
            choices: INTENTS,
            description: 'Only remove the assignment with this intent.',
          } satisfies ArgSpec,
        ]
      : []),
  ];
  const plan = async (ctx: Context, args: Args) => planUnassign(ctx, await spec.resolve(ctx, String(args.item)), args);
  const scopes = [spec.permission, ...spec.extra, READ_GROUPS];
  return defineCommand({
    domain: 'intune',
    name: `${spec.area} unassign`,
    summary: `Remove one target from a ${spec.noun}'s assignments (keeps the others)`,
    description:
      `Removes only the assignment for the given group (or all users / all devices / the exclusion); the ` +
      `other assignments are re-sent unchanged, filters included. Removing a required app assignment does not ` +
      'uninstall the app; assign it with --intent uninstall for that. To drop only a filter, use assign ' +
      '--clear-filter. Uses the beta assign action so other assignments keep their filters.',
    examples: spec.examples.unassign,
    args: args as readonly ArgSpec[],
    risk: 'destructive',
    permissions: { delegated: scopes, application: scopes, roles: INTUNE_ADMIN_ROLES },
    async preview(ctx, args) {
      return describePlan(ctx, await plan(ctx, args as unknown as Args));
    },
    async run(ctx, args) {
      const resolved = await plan(ctx, args as unknown as Args);
      return applyPlan(ctx, resolved, 'unassign', spec.next(resolved.subject));
    },
  });
}

// ---------------------------------------------------------------------------------------------------------------
// Tenant-wide assignment scan (shared by `intune assignments <group>` and `intune filters get`)

type SectionError = { code: string; message: string };

export interface AssignmentScan {
  /** Items whose assignments were read (capped by maxItems). */
  scanned: Subject[];
  results: { item: Subject; assignments: Assignment[] }[];
  errors: { kind: AssignableKind; id: string; name: string; error: SectionError }[];
  /** Catalog sections Graph refused (e.g. no permission for apps). */
  unavailable: { section: string; error: SectionError }[];
  /** More items existed than maxItems. */
  truncated: boolean;
}

async function listApps(ctx: Context): Promise<Subject[]> {
  const { items } = await ctx.graph.list<Policy>(APPS_PATH, { query: { $select: 'id,displayName' }, limit: ALL });
  return items.map((p) => ({ kind: 'app', id: p.id, name: p.displayName ?? p.id, platform: odataType(p) }));
}

async function listCompliance(ctx: Context): Promise<Subject[]> {
  const { items } = await ctx.graph.list<Policy>(COMPLIANCE_PATH, {
    query: { $select: 'id,displayName' },
    limit: ALL,
  });
  return items.map((p) => ({ kind: 'compliance', id: p.id, name: p.displayName ?? p.id, platform: odataType(p) }));
}

/**
 * Read the assignments of every app, compliance policy, classic configuration profile and settings catalog policy
 * (up to `maxItems`) through one beta `$batch` run.
 */
export async function scanAssignments(ctx: Context, maxItems: number): Promise<AssignmentScan> {
  const sections = await settle({
    app: listApps(ctx),
    compliance: listCompliance(ctx),
    config: listClassicProfiles(ctx),
    settingsCatalog: listSettingsCatalog(ctx),
  });
  const unavailable: AssignmentScan['unavailable'] = [];
  const all: Subject[] = [];
  for (const kind of ['app', 'compliance', 'config', 'settingsCatalog'] as const) {
    const section = sections[kind];
    if (isSettledError(section)) unavailable.push({ section: kind, ...section });
    else all.push(...section);
  }
  const scanned = all.slice(0, Math.max(0, maxItems));
  const responses = await ctx.graph.batch<{ value?: Assignment[] }>(
    scanned.map((item) => ({ path: `${KINDS[item.kind].path}/${seg(item.id)}/assignments` })),
    'beta',
  );
  const results: AssignmentScan['results'] = [];
  const errors: AssignmentScan['errors'] = [];
  scanned.forEach((item, index) => {
    const response = responses[index];
    if (!response || response.error) {
      errors.push({
        kind: item.kind,
        id: item.id,
        name: item.name,
        error: { code: response?.error?.code ?? 'INTERNAL', message: response?.error?.message ?? 'No batch response' },
      });
      return;
    }
    results.push({ item, assignments: response.body?.value ?? [] });
  });
  return { scanned, results, errors, unavailable, truncated: all.length > scanned.length };
}

export function scanCounts(scan: AssignmentScan) {
  const count = (kind: AssignableKind) => scan.scanned.filter((c) => c.kind === kind).length;
  return {
    apps: count('app'),
    compliancePolicies: count('compliance'),
    configurationProfiles: count('config'),
    settingsCatalogPolicies: count('settingsCatalog'),
  };
}

interface GroupHit {
  kind: AssignableKind;
  id: string;
  name: string;
  intent?: string;
  via?: string;
  filter?: string | null;
  filterMode?: string;
  filterId?: string;
}

async function parentGroups(ctx: Context, groupId: string): Promise<Record<string, string>> {
  const { items } = await ctx.graph.list<{ id: string; displayName?: string | null; '@odata.type'?: string }>(
    `/groups/${seg(groupId)}/transitiveMemberOf/microsoft.graph.group`,
    { query: { $select: 'id,displayName' }, limit: ALL },
  );
  return Object.fromEntries(items.map((g) => [g.id, g.displayName ?? g.id]));
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'apps assignments',
    summary: 'Who an app is assigned to: groups, all users / all devices, intent, filters and exclusions',
    description:
      'Use before changing an app deployment or to answer "who gets Zoom?". target is group, exclude, allUsers ' +
      'or allDevices; intent is required, available, uninstall or availableWithoutEnrollment. filter, ' +
      'filterMode (include/exclude) and filterId show an assignment filter when one is set. Reads the beta ' +
      'assignments, because v1.0 omits filters.',
    examples: ['tenantctl intune apps assignments "Zoom Workplace"'],
    args: [{ name: 'app', type: 'string', required: true, positional: true, description: 'App name or id.' }],
    risk: 'read',
    permissions: {
      delegated: [READ_APPS, READ_CONFIG, READ_GROUPS],
      application: [READ_APPS, READ_CONFIG, READ_GROUPS],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const app = await resolveApp(ctx, args.app);
      const assignments = await listAssignments(ctx, app);
      if (ctx.full) return { app, assignments };
      const [shaped] = await shapeAssignmentLists(ctx, [assignments]);
      return { app: { id: app.id, name: app.name }, assignments: shaped };
    },
  }),
  ...AREAS.flatMap((area) => [assignCommand(area), unassignCommand(area)]),
  defineCommand({
    domain: 'intune',
    name: 'assignments',
    summary: 'Which apps, compliance policies and configuration profiles target a group',
    description:
      'Answers "what will a device or user in group X get?". Scans every app, compliance policy, classic ' +
      'configuration profile and settings catalog policy and reads their assignments through beta $batch ' +
      '(v1.0 omits assignment filters). included = targets this group or a group it is nested in (via names ' +
      'the parent); excluded = this group is excluded; allUsers/allDevices = applies to everyone regardless of ' +
      'group. filter/filterMode show an assignment filter narrowing a hit. --max-items caps how many ' +
      'apps/policies are scanned (truncated=true when hit). Sections Graph refused are listed in unavailable.',
    examples: [
      'tenantctl intune assignments "Sales Laptops"',
      'tenantctl intune assignments "All Staff" --max-items 2000',
    ],
    args: [
      { name: 'group', type: 'string', required: true, positional: true, description: 'Group name, mail or id.' },
      {
        name: 'maxItems',
        type: 'number',
        default: DEFAULT_MAX_ITEMS,
        description: 'Maximum apps + policies to scan.',
      },
    ],
    risk: 'read',
    permissions: {
      delegated: [READ_APPS, READ_CONFIG, READ_GROUPS],
      application: [READ_APPS, READ_CONFIG, READ_GROUPS],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const group = await resolveGroup(ctx, args.group);
      const [sections, scan] = await Promise.all([
        settle({ memberCount: memberCount(ctx, group.id), parents: parentGroups(ctx, group.id) }),
        scanAssignments(ctx, args.maxItems),
      ]);
      const unavailable: AssignmentScan['unavailable'] = [];
      const parents = isSettledError(sections.parents) ? {} : sections.parents;
      if (isSettledError(sections.parents)) unavailable.push({ section: 'parentGroups', ...sections.parents });
      unavailable.push(...scan.unavailable);
      const relevant = scan.results.flatMap((r) => r.assignments);
      const filters = await filterNames(ctx, relevant);
      const included: GroupHit[] = [];
      const excluded: GroupHit[] = [];
      const allUsers: GroupHit[] = [];
      const allDevices: GroupHit[] = [];
      for (const { item, assignments } of scan.results) {
        for (const assignment of assignments) {
          const target = assignment.target ?? {};
          const kind = targetKind(target);
          const hit = compact({
            kind: item.kind,
            id: item.id,
            name: item.name,
            intent: assignment.intent ?? undefined,
            ...filterFields(target, filters),
          }) as GroupHit;
          const groupId = typeof target.groupId === 'string' ? target.groupId : null;
          if (kind === 'allUsers') allUsers.push(hit);
          else if (kind === 'allDevices') allDevices.push(hit);
          else if (groupId === group.id) (kind === 'exclude' ? excluded : included).push(hit);
          else if (groupId && groupId in parents) {
            (kind === 'exclude' ? excluded : included).push({ ...hit, via: parents[groupId] });
          }
        }
      }
      const memberTotal = isSettledError(sections.memberCount) ? null : sections.memberCount;
      return {
        group: { id: group.id, displayName: group.displayName, memberCount: memberTotal },
        summary:
          `${group.displayName} is targeted by ${included.length} item(s)` +
          `${excluded.length ? ` and excluded from ${excluded.length}` : ''}; ` +
          `${allUsers.length + allDevices.length} more apply to all users or all devices.`,
        included,
        excluded,
        allUsers,
        allDevices,
        scanned: scanCounts(scan),
        truncated: scan.truncated,
        ...(scan.errors.length ? { errors: scan.errors } : {}),
        ...(unavailable.length ? { unavailable } : {}),
      };
    },
  }),
];
