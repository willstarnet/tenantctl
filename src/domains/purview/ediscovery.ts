import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { pickOne, type ResolvedUser, resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { compact, isGuid } from '../../core/util.ts';

/**
 * Microsoft Purview eDiscovery (Premium) cases through Graph v1.0 /security/cases/ediscoveryCases. Holds are
 * placed per custodian: add the person as a custodian, attach their mailbox + OneDrive as a userSource, then
 * call the custodian `applyHold` action. `removeHold` lifts it while keeping the custodian in the case.
 */

const CASES = '/security/cases/ediscoveryCases';

const READ = {
  delegated: ['eDiscovery.Read.All'],
  application: ['eDiscovery.Read.All'],
  roles: ['eDiscovery Manager'],
} as const;
// The Graph permission tables list eDiscovery.Read.All as least privileged even for these POSTs; changes to
// cases need eDiscovery.ReadWrite.All in practice, so that is what writes declare.
const WRITE = {
  delegated: ['eDiscovery.ReadWrite.All'],
  application: ['eDiscovery.ReadWrite.All'],
  roles: ['eDiscovery Manager'],
} as const;

const APP_ONLY_NOTE =
  "App-only use needs eDiscovery.ReadWrite.All plus the app's service principal added to the Purview " +
  'eDiscoveryManager role group (and eDiscovery Administrator via Add-eDiscoveryCaseAdmin, or membership of each ' +
  'case). Delegated callers must be eDiscovery Manager/Administrator and a member of the case.';

export interface EdiscoveryCase {
  id: string;
  displayName: string;
  description?: string | null;
  status?: string;
  externalId?: string | null;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: { user?: { displayName?: string; userPrincipalName?: string } };
}

export interface Custodian {
  id: string;
  displayName?: string;
  email?: string;
  status?: string;
  holdStatus?: string;
  createdDateTime?: string;
  releasedDateTime?: string | null;
}

interface UserSource {
  id: string;
  email?: string;
  includedSources?: string;
  holdStatus?: string;
}

export function shapeCase(c: EdiscoveryCase) {
  return {
    id: c.id,
    name: c.displayName,
    status: c.status ?? null,
    description: c.description || null,
    externalId: c.externalId || null,
    createdDateTime: c.createdDateTime ?? null,
    createdBy: c.createdBy?.user?.userPrincipalName ?? c.createdBy?.user?.displayName ?? null,
  };
}

export function shapeCustodian(c: Custodian) {
  return {
    id: c.id,
    displayName: c.displayName ?? null,
    email: c.email ?? null,
    status: c.status ?? null,
    holdStatus: c.holdStatus ?? null,
    createdDateTime: c.createdDateTime ?? null,
    releasedDateTime: c.releasedDateTime ?? null,
  };
}

/** A case by id or name; several name matches are AMBIGUOUS. */
export async function resolveCase(ctx: Context, ref: string): Promise<EdiscoveryCase> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'An eDiscovery case is required');
  if (isGuid(value)) {
    const found = await ctx.graph.tryGet<EdiscoveryCase>(`${CASES}/${seg(value)}`);
    if (!found) throw new CliError('NOT_FOUND', `No eDiscovery case with id ${value}`);
    return found;
  }
  const { items } = await ctx.graph.list<EdiscoveryCase>(CASES, { limit: Number.POSITIVE_INFINITY, pageSize: false });
  const wanted = value.toLowerCase();
  const matches = items.filter((c) => c.displayName?.toLowerCase().includes(wanted));
  return pickOne(matches, value, 'eDiscovery case', (c) => [c.displayName], shapeCase);
}

async function listCustodians(ctx: Context, caseId: string): Promise<Custodian[]> {
  const { items } = await ctx.graph.list<Custodian>(`${CASES}/${seg(caseId)}/custodians`, {
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  return items;
}

function findCustodian(custodians: Custodian[], user: ResolvedUser): Custodian | undefined {
  const addresses = [user.mail, user.userPrincipalName].filter(Boolean).map((a) => (a as string).toLowerCase());
  return custodians.find((c) => c.email && addresses.includes(c.email.toLowerCase()));
}

interface HoldPlan {
  kase: EdiscoveryCase;
  targets: { user: ResolvedUser; custodian: Custodian | undefined }[];
}

async function planHold(ctx: Context, caseRef: string, userRefs: string[]): Promise<HoldPlan> {
  const kase = await resolveCase(ctx, caseRef);
  const [custodians, users] = await Promise.all([
    listCustodians(ctx, kase.id),
    Promise.all(userRefs.map((ref) => resolveUser(ctx, ref))),
  ]);
  return { kase, targets: users.map((user) => ({ user, custodian: findCustodian(custodians, user) })) };
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'purview',
    name: 'ediscovery list',
    summary: 'List eDiscovery (Premium) cases with status',
    description: `Cases the caller can see. Next: \`ediscovery custodians <case>\` for who is on hold. ${APP_ONLY_NOTE}`,
    examples: ['m365 purview ediscovery list'],
    args: [...LIST_ARGS],
    risk: 'read',
    permissions: READ,
    async run(ctx, args) {
      const { items, truncated } = await ctx.graph.list<EdiscoveryCase>(CASES, {
        limit: listLimit(args),
        pageSize: false,
      });
      return { items: ctx.full ? items : items.map(shapeCase), truncated };
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'ediscovery create',
    summary: 'Create an eDiscovery (Premium) case',
    description: `Creates an empty case; add people with \`ediscovery hold <case> <user...>\`. ${APP_ONLY_NOTE}`,
    examples: ['m365 purview ediscovery create --name "HR-2026-014 Bob Smith" --description "Departure review"'],
    args: [
      { name: 'name', type: 'string', required: true, description: 'Case name (unique in the tenant).' },
      { name: 'description', type: 'string', description: 'Case description.' },
      { name: 'externalId', type: 'string', description: 'Your own case/ticket number.' },
    ],
    risk: 'write',
    permissions: WRITE,
    async preview(_ctx, args) {
      return {
        action: `Create eDiscovery case "${args.name}"`,
        name: args.name,
        description: args.description ?? null,
      };
    },
    async run(ctx, args) {
      const created = await ctx.graph.post<EdiscoveryCase>(
        CASES,
        compact({ displayName: args.name, description: args.description, externalId: args.externalId }),
      );
      return ctx.full ? created : shapeCase(created);
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'ediscovery custodians',
    summary: 'People in an eDiscovery case and whether their data is on hold',
    description:
      'holdStatus: notApplied, applying, applied, removing, partial (some sources failed). status released = ' +
      `removed from the case. ${APP_ONLY_NOTE}`,
    examples: ['m365 purview ediscovery custodians "HR-2026-014"'],
    args: [{ name: 'case', type: 'string', required: true, positional: true, description: 'Case name or id.' }],
    risk: 'read',
    permissions: READ,
    async run(ctx, args) {
      const kase = await resolveCase(ctx, args.case);
      const custodians = await listCustodians(ctx, kase.id);
      return { items: ctx.full ? custodians : custodians.map(shapeCustodian), truncated: false };
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'ediscovery hold',
    summary: "Preserve people's mailbox and OneDrive under an eDiscovery case hold",
    description:
      'For each person: adds them as a custodian (if not already), attaches their mailbox and OneDrive as data ' +
      'sources, then applies the hold. Held content cannot be permanently deleted by the user or by retention ' +
      'until the hold is released, and it counts against their storage. Applying takes minutes to hours: check ' +
      `\`ediscovery custodians <case>\` until holdStatus is applied. ${APP_ONLY_NOTE}`,
    examples: ['m365 purview ediscovery hold "HR-2026-014" bob@contoso.com jane@contoso.com'],
    args: [
      { name: 'case', type: 'string', required: true, positional: true, description: 'Case name or id.' },
      {
        name: 'user',
        type: 'string[]',
        required: true,
        positional: true,
        description: 'People to hold: UPN, email, name or id.',
      },
    ],
    risk: 'destructive',
    permissions: {
      ...WRITE,
      delegated: [...WRITE.delegated, 'User.ReadBasic.All'],
      application: [...WRITE.application, 'User.Read.All'],
    },
    async preview(ctx, args) {
      const plan = await planHold(ctx, args.case, args.user);
      const names = plan.targets.map((t) => t.user.userPrincipalName).join(', ');
      return {
        action: `Place the mailbox and OneDrive of ${names} on hold in eDiscovery case "${plan.kase.displayName}"`,
        case: shapeCase(plan.kase),
        people: plan.targets.map(({ user, custodian }) => ({
          user: user.userPrincipalName,
          displayName: user.displayName,
          alreadyCustodian: Boolean(custodian),
          currentHoldStatus: custodian?.holdStatus ?? null,
        })),
      };
    },
    async run(ctx, args) {
      const plan = await planHold(ctx, args.case, args.user);
      const base = `${CASES}/${seg(plan.kase.id)}/custodians`;
      const held: { user: string; custodianId: string; addedToCase: boolean }[] = [];
      for (const { user, custodian } of plan.targets) {
        const email = user.mail ?? user.userPrincipalName;
        const target = custodian ?? (await ctx.graph.post<Custodian>(base, { email }));
        const sources = custodian
          ? (await ctx.graph.list<UserSource>(`${base}/${seg(target.id)}/userSources`, { pageSize: false })).items
          : [];
        if (!sources.length) {
          await ctx.graph.post(`${base}/${seg(target.id)}/userSources`, { email, includedSources: 'mailbox,site' });
        }
        held.push({ user: user.userPrincipalName, custodianId: target.id, addedToCase: !custodian });
      }
      await ctx.graph.post(`${base}/applyHold`, { ids: held.map((h) => h.custodianId) });
      return {
        case: plan.kase.displayName,
        holdStatus: 'applying',
        people: held,
        next: `m365 purview ediscovery custodians ${plan.kase.id}`,
      };
    },
  }),
  defineCommand({
    domain: 'purview',
    name: 'ediscovery release',
    summary: "Remove the eDiscovery hold from a person's mailbox and OneDrive",
    description:
      'Calls removeHold for the custodian; they stay in the case (status active) so their data can still be ' +
      'searched, but held content becomes deletable again (retention policies and user deletes apply). Other ' +
      `cases or retention holds on the same person are unaffected. ${APP_ONLY_NOTE}`,
    examples: ['m365 purview ediscovery release "HR-2026-014" bob@contoso.com'],
    args: [
      { name: 'case', type: 'string', required: true, positional: true, description: 'Case name or id.' },
      {
        name: 'user',
        type: 'string',
        required: true,
        positional: true,
        description: 'Custodian: UPN, email, name or id.',
      },
    ],
    risk: 'destructive',
    permissions: {
      ...WRITE,
      delegated: [...WRITE.delegated, 'User.ReadBasic.All'],
      application: [...WRITE.application, 'User.Read.All'],
    },
    async preview(ctx, args) {
      const plan = await planHold(ctx, args.case, [args.user]);
      const { user, custodian } = plan.targets[0] as HoldPlan['targets'][number];
      if (!custodian) {
        throw new CliError('NOT_FOUND', `${user.userPrincipalName} is not a custodian in "${plan.kase.displayName}"`, {
          hint: `List custodians with \`m365 purview ediscovery custodians ${plan.kase.id}\`.`,
        });
      }
      return {
        action: `Remove the eDiscovery hold on ${user.userPrincipalName}'s mailbox and OneDrive in case "${plan.kase.displayName}"`,
        case: shapeCase(plan.kase),
        custodian: shapeCustodian(custodian),
      };
    },
    async run(ctx, args) {
      const plan = await planHold(ctx, args.case, [args.user]);
      const { user, custodian } = plan.targets[0] as HoldPlan['targets'][number];
      if (!custodian)
        throw new CliError('NOT_FOUND', `${user.userPrincipalName} is not a custodian in "${plan.kase.displayName}"`);
      await ctx.graph.post(`${CASES}/${seg(plan.kase.id)}/custodians/${seg(custodian.id)}/removeHold`);
      return {
        case: plan.kase.displayName,
        user: user.userPrincipalName,
        custodianId: custodian.id,
        holdStatus: 'removing',
        next: `m365 purview ediscovery custodians ${plan.kase.id}`,
      };
    },
  }),
];
