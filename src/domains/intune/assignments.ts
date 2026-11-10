import { CliError } from '../../core/errors.ts';
import { type ApiVersion, seg } from '../../core/graph.ts';
import { type ResolvedGroup, resolveGroup } from '../../core/resolve.ts';
import { type AnyCommand, type ArgSpec, type Context, defineCommand } from '../../core/types.ts';
import { compact, isSettledError, settle } from '../../core/util.ts';
import { attempt, INTUNE_ADMIN_ROLES, INTUNE_READ_ROLES, pickByName, READ_CONFIG } from './common.ts';
import {
  APPS_PATH,
  COMPLIANCE_PATH,
  CONFIG_PATH,
  type Policy,
  READ_APPS,
  resolveCatalogItem,
  SETTINGS_CATALOG_PATH,
} from './policies.ts';

/**
 * Intune `/assign` actions REPLACE the whole assignment list of an app or policy. Every assign/unassign here
 * reads the current assignments first, merges the change, and sends the full resulting list, so existing
 * assignments survive.
 */

const WRITE_APPS = 'DeviceManagementApps.ReadWrite.All';
const WRITE_CONFIG = 'DeviceManagementConfiguration.ReadWrite.All';
const READ_GROUPS = 'GroupMember.Read.All';
const ALL = Number.POSITIVE_INFINITY;

export type AssignableKind = 'app' | 'compliance' | 'config' | 'settingsCatalog';

interface KindSpec {
  path: string;
  version: ApiVersion;
  /** `@odata.type` of one assignment in the `/assign` body. */
  assignmentType: string;
  /** Body property of the `/assign` action. */
  bodyKey: 'mobileAppAssignments' | 'assignments';
  label: string;
}

const KINDS: Record<AssignableKind, KindSpec> = {
  app: {
    path: APPS_PATH,
    version: 'v1.0',
    assignmentType: '#microsoft.graph.mobileAppAssignment',
    bodyKey: 'mobileAppAssignments',
    label: 'app',
  },
  compliance: {
    path: COMPLIANCE_PATH,
    version: 'v1.0',
    assignmentType: '#microsoft.graph.deviceCompliancePolicyAssignment',
    bodyKey: 'assignments',
    label: 'compliance policy',
  },
  config: {
    path: CONFIG_PATH,
    version: 'v1.0',
    assignmentType: '#microsoft.graph.deviceConfigurationAssignment',
    bodyKey: 'assignments',
    label: 'configuration profile',
  },
  settingsCatalog: {
    path: SETTINGS_CATALOG_PATH,
    version: 'beta',
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
}

interface ResolvedTarget {
  kind: TargetKind;
  target: Target;
  group?: ResolvedGroup;
  memberCount?: number | null;
}

type Args = Record<string, unknown>;

/** `#microsoft.graph.groupAssignmentTarget` (or without `#`) -> `group`; unknown kinds keep their type name. */
function targetKind(target: Target | null | undefined): string {
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
    version: spec.version,
    limit: ALL,
    pageSize: false,
  });
  return items;
}

async function memberCount(ctx: Context, groupId: string): Promise<number | null> {
  try {
    const count = await ctx.graph.get<string | number>(`/groups/${seg(groupId)}/members/$count`, { advanced: true });
    const value = Number(count);
    return Number.isFinite(value) ? value : null;
  } catch {
    return null;
  }
}

/** Display names for group ids, fetched through `$batch`; unknown or unreadable groups map to null. */
async function groupNames(ctx: Context, ids: string[], known: Record<string, string> = {}) {
  const names: Record<string, string | null> = { ...known };
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

export function shapeAssignment(assignment: Assignment, names: Record<string, string | null>) {
  const target = assignment.target ?? {};
  const groupId = typeof target.groupId === 'string' ? target.groupId : undefined;
  const filterType = target.deviceAndAppManagementAssignmentFilterType;
  return compact({
    target: targetKind(target),
    group: groupId ? (names[groupId] ?? null) : undefined,
    groupId,
    intent: assignment.intent ?? undefined,
    filterId: (filterType && filterType !== 'none' && target.deviceAndAppManagementAssignmentFilterId) || undefined,
    filterType: filterType && filterType !== 'none' ? filterType : undefined,
  });
}

async function shapeAll(ctx: Context, lists: Assignment[][], targets: ResolvedTarget[] = []) {
  const known: Record<string, string> = {};
  for (const t of targets) if (t.group) known[t.group.id] = t.group.displayName;
  const ids = lists.flat().flatMap((a) => (typeof a.target?.groupId === 'string' ? [a.target.groupId] : []));
  const names = await groupNames(ctx, ids, known);
  return lists.map((list) => list.map((a) => shapeAssignment(a, names)));
}

async function resolveApp(ctx: Context, ref: string): Promise<Subject> {
  const app = await resolveCatalogItem(ctx, APPS_PATH, ref, 'app');
  return { kind: 'app', id: app.id, name: app.displayName ?? app.id };
}

async function resolveCompliance(ctx: Context, ref: string): Promise<Subject> {
  const policy = await resolveCatalogItem(ctx, COMPLIANCE_PATH, ref, 'compliance policy');
  return { kind: 'compliance', id: policy.id, name: policy.displayName ?? policy.id };
}

async function listSettingsCatalog(ctx: Context): Promise<Subject[]> {
  const { items } = await ctx.graph.list<Policy>(SETTINGS_CATALOG_PATH, {
    version: 'beta',
    query: { $select: 'id,name' },
    limit: ALL,
  });
  return items.map((p) => ({ kind: 'settingsCatalog', id: p.id, name: p.name ?? p.id }));
}

async function listClassicProfiles(ctx: Context): Promise<Subject[]> {
  const { items } = await ctx.graph.list<Policy>(CONFIG_PATH, { query: { $select: 'id,displayName' }, limit: ALL });
  return items.map((p) => ({ kind: 'config', id: p.id, name: p.displayName ?? p.id }));
}

/** A configuration profile by name or id: classic profiles (v1.0) and settings catalog policies (beta) together. */
async function resolveProfile(ctx: Context, ref: string): Promise<Subject> {
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

function assignAction(subject: Subject, targets: ResolvedTarget[], intent: string | undefined): string {
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
  return exclude ? `${sentence}, excluding ${describeTarget(exclude)}` : sentence;
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

interface Plan {
  subject: Subject;
  targets: ResolvedTarget[];
  current: Assignment[];
  next: Assignment[];
  changed: boolean;
  action: string;
}

async function planAssign(ctx: Context, subject: Subject, args: Args): Promise<Plan> {
  const spec = KINDS[subject.kind];
  const intent = subject.kind === 'app' ? String(args.intent) : undefined;
  if (intent === 'available' && args.allDevices) {
    throw new CliError('USAGE', 'Available apps can only be targeted at users, not --all-devices', {
      hint: 'Use --all-users or a user group for available apps, or --intent required for devices.',
    });
  }
  const [targets, current] = await Promise.all([resolveTargets(ctx, args, 'assign'), listAssignments(ctx, subject)]);
  const next = current.map((a) => clean(spec, a));
  let changed = false;
  for (const target of targets) {
    const index = next.findIndex((a) => sameTarget(a.target, target.target));
    const existing = index >= 0 ? next[index] : undefined;
    if (!existing) {
      next.push({ '@odata.type': spec.assignmentType, ...(intent ? { intent } : {}), target: target.target });
      changed = true;
    } else if (intent && existing.intent !== intent) {
      // Intune allows one intent per group per app: switching intent replaces the existing assignment.
      next[index] = { ...existing, intent };
      changed = true;
    }
  }
  return { subject, targets, current, next, changed, action: assignAction(subject, targets, intent) };
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
          ? `List its assignments with \`m365 intune apps assignments ${subject.id}\`.`
          : `Check what targets the group with \`m365 intune assignments ${target.group?.id ?? '<group>'}\`.`,
    });
  }
  const next = current.filter((a) => !matches(a)).map((a) => clean(spec, a));
  return { subject, targets, current, next, changed: true, action: unassignAction(subject, target, removed) };
}

async function describePlan(ctx: Context, plan: Plan) {
  const spec = KINDS[plan.subject.kind];
  const [current, result] = await shapeAll(ctx, [plan.current, plan.next], plan.targets);
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
      }),
    ),
    changed: plan.changed,
    current,
    result,
    request: plan.changed
      ? { method: 'POST', path: `${spec.path}/${plan.subject.id}/assign`, version: spec.version }
      : undefined,
  };
}

async function applyPlan(ctx: Context, plan: Plan, mode: 'assign' | 'unassign', next: string) {
  const spec = KINDS[plan.subject.kind];
  const view = await describePlan(ctx, plan);
  if (!plan.changed) return { status: 'unchanged', ...view, request: undefined };
  await ctx.graph.post(
    `${spec.path}/${seg(plan.subject.id)}/assign`,
    { [spec.bodyKey]: plan.next },
    { version: spec.version },
  );
  return {
    status: mode === 'assign' ? 'assigned' : 'unassigned',
    action: view.action,
    ...(plan.subject.kind === 'app' ? { app: view.app } : { policy: view.policy }),
    targets: view.targets,
    assignments: view.result,
    next,
  };
}

const TARGET_ARGS: ArgSpec[] = [
  { name: 'group', type: 'string', description: 'Target group: name, mail or id.' },
  { name: 'allUsers', type: 'boolean', description: 'Target all licensed users.' },
  { name: 'allDevices', type: 'boolean', description: 'Target all Intune-managed devices.' },
];

const INTENTS = ['required', 'available', 'uninstall'] as const;

interface AreaSpec {
  area: 'apps' | 'compliance' | 'config';
  noun: string;
  ref: string;
  resolve: (ctx: Context, ref: string) => Promise<Subject>;
  permission: string;
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
    next: (s) => `m365 intune apps status ${quoted(s.name)}`,
    description:
      'intent required installs the app on targeted devices/users, available offers it in Company Portal ' +
      '(users only: refused with --all-devices, and Intune ignores available for device groups), uninstall ' +
      'removes it. A group already assigned with another intent is switched to the new one. Assignment settings ' +
      '(delivery optimisation, notifications) take Intune defaults; adjust them in the admin center.',
    examples: {
      assign: [
        'm365 intune apps assign "Zoom Workplace" --intent required --group "Sales Laptops"',
        'm365 intune apps assign "Company Portal" --intent available --all-users --exclude "Kiosk Users"',
      ],
      unassign: ['m365 intune apps unassign "Zoom Workplace" --group "Sales Laptops"'],
    },
  },
  {
    area: 'compliance',
    noun: 'compliance policy',
    ref: 'Compliance policy name or id.',
    resolve: resolveCompliance,
    permission: WRITE_CONFIG,
    next: (s) => `m365 intune compliance status ${quoted(s.name)}`,
    description:
      'Newly targeted devices are evaluated at their next check-in; non-compliant ones can lose access through ' +
      'Conditional Access after the grace period.',
    examples: {
      assign: ['m365 intune compliance assign "Windows baseline" --group "All Windows" --exclude "Kiosks"'],
      unassign: ['m365 intune compliance unassign "Windows baseline" --exclude "Kiosks"'],
    },
  },
  {
    area: 'config',
    noun: 'configuration profile',
    ref: 'Profile name or id (classic profile or settings catalog policy).',
    resolve: resolveProfile,
    permission: WRITE_CONFIG,
    next: (s) =>
      s.kind === 'config' ? `m365 intune config status ${quoted(s.name)}` : `m365 intune assignments <group>`,
    description:
      'Works for classic profiles (v1.0 deviceConfigurations) and settings catalog / endpoint security policies ' +
      '(beta configurationPolicies; Microsoft has not published a v1.0 equivalent).',
    examples: {
      assign: ['m365 intune config assign "Corporate Wi-Fi" --group "Sales Laptops"'],
      unassign: ['m365 intune config unassign "Edge settings" --all-devices'],
    },
  },
];

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
  ];
  const plan = async (ctx: Context, args: Args) => planAssign(ctx, await spec.resolve(ctx, String(args.item)), args);
  return defineCommand({
    domain: 'intune',
    name: `${spec.area} assign`,
    summary: `Assign a ${spec.noun} to a group, all users or all devices (keeps existing assignments)`,
    description:
      `Adds one target to the ${spec.noun}'s assignments; existing assignments are read first and kept, ` +
      `because Intune's assign action replaces the whole list. ${spec.description} The preview shows the ` +
      'target group with its member count, the current assignments and the resulting list.',
    examples: spec.examples.assign,
    args: args as readonly ArgSpec[],
    risk: 'destructive',
    permissions: {
      delegated: [spec.permission, READ_GROUPS],
      application: [spec.permission, READ_GROUPS],
      roles: INTUNE_ADMIN_ROLES,
    },
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
  return defineCommand({
    domain: 'intune',
    name: `${spec.area} unassign`,
    summary: `Remove one target from a ${spec.noun}'s assignments (keeps the others)`,
    description:
      `Removes only the assignment for the given group (or all users / all devices / the exclusion); the ` +
      `other assignments are re-sent unchanged. Removing a required app assignment does not uninstall the app; ` +
      'assign it with --intent uninstall for that.',
    examples: spec.examples.unassign,
    args: args as readonly ArgSpec[],
    risk: 'destructive',
    permissions: {
      delegated: [spec.permission, READ_GROUPS],
      application: [spec.permission, READ_GROUPS],
      roles: INTUNE_ADMIN_ROLES,
    },
    async preview(ctx, args) {
      return describePlan(ctx, await plan(ctx, args as unknown as Args));
    },
    async run(ctx, args) {
      const resolved = await plan(ctx, args as unknown as Args);
      return applyPlan(ctx, resolved, 'unassign', spec.next(resolved.subject));
    },
  });
}

interface GroupHit {
  kind: AssignableKind;
  id: string;
  name: string;
  intent?: string;
  via?: string;
  filterType?: string;
}

const DEFAULT_MAX_ITEMS = 500;

async function parentGroups(ctx: Context, groupId: string): Promise<Record<string, string>> {
  const { items } = await ctx.graph.list<{ id: string; displayName?: string | null; '@odata.type'?: string }>(
    `/groups/${seg(groupId)}/transitiveMemberOf/microsoft.graph.group`,
    { query: { $select: 'id,displayName' }, limit: ALL },
  );
  return Object.fromEntries(items.map((g) => [g.id, g.displayName ?? g.id]));
}

async function listApps(ctx: Context): Promise<Subject[]> {
  const { items } = await ctx.graph.list<Policy>(APPS_PATH, { query: { $select: 'id,displayName' }, limit: ALL });
  return items.map((p) => ({ kind: 'app', id: p.id, name: p.displayName ?? p.id }));
}

async function listCompliance(ctx: Context): Promise<Subject[]> {
  const { items } = await ctx.graph.list<Policy>(COMPLIANCE_PATH, {
    query: { $select: 'id,displayName' },
    limit: ALL,
  });
  return items.map((p) => ({ kind: 'compliance', id: p.id, name: p.displayName ?? p.id }));
}

async function assignmentsByBatch(ctx: Context, items: Subject[], version: ApiVersion) {
  return ctx.graph.batch<{ value?: Assignment[] }>(
    items.map((item) => ({ path: `${KINDS[item.kind].path}/${seg(item.id)}/assignments` })),
    version,
  );
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'apps assignments',
    summary: 'Who an app is assigned to: groups, all users / all devices, intent and exclusions',
    description:
      'Use before changing an app deployment or to answer "who gets Zoom?". target is group, exclude, allUsers ' +
      'or allDevices; intent is required, available, uninstall or availableWithoutEnrollment. filterType shows ' +
      'an assignment filter (include/exclude) when one is set.',
    examples: ['m365 intune apps assignments "Zoom Workplace"'],
    args: [{ name: 'app', type: 'string', required: true, positional: true, description: 'App name or id.' }],
    risk: 'read',
    permissions: {
      delegated: [READ_APPS, READ_GROUPS],
      application: [READ_APPS, READ_GROUPS],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const app = await resolveApp(ctx, args.app);
      const assignments = await listAssignments(ctx, app);
      if (ctx.full) return { app, assignments };
      const [shaped] = await shapeAll(ctx, [assignments]);
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
      'configuration profile and settings catalog policy (beta) and reads their assignments through $batch. ' +
      'included = targets this group or a group it is nested in (via names the parent); excluded = this group ' +
      'is excluded; allUsers/allDevices = applies to everyone regardless of group. --max-items caps how many ' +
      'apps/policies are scanned (truncated=true when hit). Sections Graph refused are listed in unavailable.',
    examples: ['m365 intune assignments "Sales Laptops"', 'm365 intune assignments "All Staff" --max-items 2000'],
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
      const sections = await settle({
        memberCount: memberCount(ctx, group.id),
        parents: parentGroups(ctx, group.id),
        app: listApps(ctx),
        compliance: listCompliance(ctx),
        config: listClassicProfiles(ctx),
        settingsCatalog: listSettingsCatalog(ctx),
      });
      const unavailable: { section: string; error: { code: string; message: string } }[] = [];
      const parents = isSettledError(sections.parents) ? {} : sections.parents;
      if (isSettledError(sections.parents)) unavailable.push({ section: 'parentGroups', ...sections.parents });
      const all: Subject[] = [];
      for (const kind of ['app', 'compliance', 'config', 'settingsCatalog'] as const) {
        const section = sections[kind];
        if (isSettledError(section)) unavailable.push({ section: kind, ...section });
        else all.push(...section);
      }
      const max = Math.max(0, args.maxItems);
      const scanned = all.slice(0, max);
      const v1 = scanned.filter((c) => KINDS[c.kind].version === 'v1.0');
      const beta = scanned.filter((c) => KINDS[c.kind].version === 'beta');
      const [v1Results, betaResults] = await Promise.all([
        assignmentsByBatch(ctx, v1, 'v1.0'),
        assignmentsByBatch(ctx, beta, 'beta'),
      ]);
      const included: GroupHit[] = [];
      const excluded: GroupHit[] = [];
      const allUsers: GroupHit[] = [];
      const allDevices: GroupHit[] = [];
      const errors: { kind: AssignableKind; id: string; name: string; error: { code: string; message: string } }[] = [];
      const visit = (item: Subject, result: (typeof v1Results)[number] | undefined) => {
        if (!result || result.error) {
          errors.push({
            kind: item.kind,
            id: item.id,
            name: item.name,
            error: { code: result?.error?.code ?? 'INTERNAL', message: result?.error?.message ?? 'No batch response' },
          });
          return;
        }
        for (const assignment of result.body?.value ?? []) {
          const target = assignment.target ?? {};
          const kind = targetKind(target);
          const filterType = target.deviceAndAppManagementAssignmentFilterType;
          const hit = compact({
            kind: item.kind,
            id: item.id,
            name: item.name,
            intent: assignment.intent ?? undefined,
            filterType: filterType && filterType !== 'none' ? filterType : undefined,
          }) as GroupHit;
          const groupId = typeof target.groupId === 'string' ? target.groupId : null;
          if (kind === 'allUsers') allUsers.push(hit);
          else if (kind === 'allDevices') allDevices.push(hit);
          else if (groupId === group.id) (kind === 'exclude' ? excluded : included).push(hit);
          else if (groupId && groupId in parents) {
            (kind === 'exclude' ? excluded : included).push({ ...hit, via: parents[groupId] });
          }
        }
      };
      v1.forEach((item, index) => visit(item, v1Results[index]));
      beta.forEach((item, index) => visit(item, betaResults[index]));
      const count = (kind: AssignableKind) => scanned.filter((c) => c.kind === kind).length;
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
        scanned: {
          apps: count('app'),
          compliancePolicies: count('compliance'),
          configurationProfiles: count('config'),
          settingsCatalogPolicies: count('settingsCatalog'),
        },
        truncated: all.length > scanned.length,
        ...(errors.length ? { errors } : {}),
        ...(unavailable.length ? { unavailable } : {}),
      };
    },
  }),
];
