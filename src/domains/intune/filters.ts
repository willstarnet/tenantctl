import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, defineCommand } from '../../core/types.ts';
import { compact } from '../../core/util.ts';
import {
  type AssignmentFilter,
  assignmentGroupNames,
  DEFAULT_MAX_ITEMS,
  listFilters,
  platformFamily,
  resolveFilter,
  scanAssignments,
  scanCounts,
  targetKind,
} from './assignments.ts';
import { INTUNE_READ_ROLES, READ_APPS, READ_CONFIG, READ_GROUPS } from './common.ts';

const PLATFORMS = ['windows', 'ios', 'macos', 'android'] as const;

function shapeFilter(filter: AssignmentFilter) {
  return compact({
    id: filter.id,
    name: filter.displayName ?? null,
    platform: filter.platform ?? null,
    managementType: filter.assignmentFilterManagementType ?? null,
    rule: filter.rule ?? null,
    description: filter.description || undefined,
    lastModifiedDateTime: filter.lastModifiedDateTime ?? null,
  });
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'intune',
    name: 'filters list',
    summary: 'List Intune assignment filters with platform, management type and rule',
    description:
      'Assignment filters narrow an app or policy assignment to matching devices (e.g. corporate-owned only). ' +
      'managementType is devices (enrolled devices) or apps (managed apps / app protection). --platform matches ' +
      'the platform family client-side (windows covers windows10AndLater, windows81AndLater...). Use the name ' +
      'or id with `intune apps|compliance|config assign --filter`. (beta API: assignment filters have no v1.0 ' +
      'equivalent.)',
    examples: ['tenantctl intune filters list', 'tenantctl intune filters list --platform windows'],
    args: [
      { name: 'platform', type: 'string', choices: PLATFORMS, description: 'Only filters for this platform.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: [READ_CONFIG], application: [READ_CONFIG], roles: INTUNE_READ_ROLES },
    async run(ctx, args) {
      const limit = listLimit(args);
      const platform = args.platform;
      const result = await listFilters(ctx, { select: !ctx.full, limit: platform ? undefined : limit });
      const matching = platform ? result.items.filter((f) => platformFamily(f.platform) === platform) : result.items;
      const items = matching.slice(0, limit);
      const truncated = result.truncated || matching.length > items.length;
      return { items: ctx.full ? items : items.map(shapeFilter), truncated };
    },
  }),
  defineCommand({
    domain: 'intune',
    name: 'filters get',
    summary: 'Show one assignment filter: rule, platform and every assignment that uses it',
    description:
      'Accepts the filter name (or part of it) or id. usedBy lists each app/policy assignment carrying the ' +
      'filter: target (group, allUsers, allDevices), group name and filterMode (include/exclude). The documented ' +
      '`payloads` property of the filter came back empty in a live check even for filters in use, so usage is ' +
      'found by scanning every app, compliance policy, configuration profile and settings catalog policy through ' +
      'beta $batch, like `intune assignments`. --max-items caps the scan (truncated=true when hit). (beta API.)',
    examples: ['tenantctl intune filters get "Corporate devices"'],
    args: [
      { name: 'filter', type: 'string', required: true, positional: true, description: 'Filter name or id.' },
      {
        name: 'maxItems',
        type: 'number',
        default: DEFAULT_MAX_ITEMS,
        description: 'Maximum apps + policies to scan.',
      },
    ],
    risk: 'read',
    permissions: {
      delegated: [READ_CONFIG, READ_APPS, READ_GROUPS],
      application: [READ_CONFIG, READ_APPS, READ_GROUPS],
      roles: INTUNE_READ_ROLES,
    },
    async run(ctx, args) {
      const { items } = await listFilters(ctx, { select: !ctx.full });
      const filter = resolveFilter(items, args.filter);
      const scan = await scanAssignments(ctx, args.maxItems);
      const uses = scan.results.flatMap(({ item, assignments }) =>
        assignments
          .filter(
            (a) =>
              a.target?.deviceAndAppManagementAssignmentFilterId === filter.id &&
              a.target.deviceAndAppManagementAssignmentFilterType !== 'none',
          )
          .map((assignment) => ({ item, assignment })),
      );
      const names = await assignmentGroupNames(
        ctx,
        uses.map(({ assignment }) => assignment),
      );
      const usedBy = uses.map(({ item, assignment }) => {
        const groupId = typeof assignment.target?.groupId === 'string' ? assignment.target.groupId : undefined;
        return compact({
          kind: item.kind,
          id: item.id,
          name: item.name,
          target: targetKind(assignment.target),
          group: groupId ? (names[groupId] ?? null) : undefined,
          groupId,
          intent: assignment.intent ?? undefined,
          filterMode: assignment.target?.deviceAndAppManagementAssignmentFilterType ?? null,
        });
      });
      return {
        filter: ctx.full ? filter : shapeFilter(filter),
        summary: `${filter.displayName ?? filter.id} is used by ${usedBy.length} assignment(s) on ${new Set(usedBy.map((u) => u.id)).size} item(s).`,
        usedBy,
        scanned: scanCounts(scan),
        truncated: scan.truncated,
        ...(scan.errors.length ? { errors: scan.errors } : {}),
        ...(scan.unavailable.length ? { unavailable: scan.unavailable } : {}),
      };
    },
  }),
];
