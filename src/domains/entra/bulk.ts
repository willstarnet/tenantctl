import { readCsvRows } from '../../core/csvInput.ts';
import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { type ResolvedGroup, resolveGroup, resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { compact, isGuid } from '../../core/util.ts';
import { type ResolvedMember, resolveMember } from './directory.ts';
import { generatePassword } from './users.ts';

/**
 * Guest invitations and CSV-driven bulk changes (create users, add group members). Bulk commands report one
 * result per CSV row and never stop at the first failed row; follow-up calls go through `$batch`.
 */

const UPN_SHAPE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
/** Graph answers this when a member reference is added twice. */
const ALREADY_MEMBER = /already exist/i;

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** Run `fn` over `items` with at most `limit` in flight; results in input order. */
async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const index = next++;
      results[index] = await fn(items[index] as T);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

/** Resolve each distinct reference once; failures become `{ error }` so every row can report its own problem. */
async function resolveEach<R>(
  refs: Iterable<string>,
  resolve: (ref: string) => Promise<R>,
): Promise<Map<string, R | { error: string }>> {
  const unique = [...new Set(refs)];
  const resolved = await mapLimit(unique, 5, (ref) =>
    resolve(ref).then(
      (value): R | { error: string } => value,
      (error: unknown) => ({ error: errorText(error) }),
    ),
  );
  return new Map(unique.map((ref, index) => [ref.toLowerCase(), resolved[index] as R | { error: string }]));
}

function isFailure(value: unknown): value is { error: string } {
  return typeof value === 'object' && value !== null && 'error' in value;
}

const memberRef = (ctx: Context, id: string) => ({ '@odata.id': ctx.graph.url(`/directoryObjects/${seg(id)}`) });

async function verifiedDomains(ctx: Context): Promise<string[]> {
  const org = await ctx.graph.get<{ value?: { verifiedDomains?: { name: string }[] }[] }>('/organization', {
    query: { $select: 'verifiedDomains' },
  });
  return (org.value ?? []).flatMap((o) => o.verifiedDomains ?? []).map((d) => d.name.toLowerCase());
}

// ---------------------------------------------------------------------------------------------------------------
// Guest invite
// ---------------------------------------------------------------------------------------------------------------

interface Invitation {
  id: string;
  status?: string | null;
  inviteRedeemUrl?: string | null;
  invitedUserEmailAddress?: string | null;
  invitedUser?: { id: string; userPrincipalName?: string | null } | null;
}

async function planInvite(ctx: Context, args: { email: string; redirectUrl: string; group?: string }) {
  const email = args.email.trim();
  if (!UPN_SHAPE.test(email)) throw new CliError('USAGE', `"${email}" is not an email address`);
  if (!/^https?:\/\//i.test(args.redirectUrl)) {
    throw new CliError('USAGE', `--redirect-url must be an http(s) URL, got "${args.redirectUrl}"`);
  }
  const domain = (email.split('@')[1] as string).toLowerCase();
  const [domains, group] = await Promise.all([
    verifiedDomains(ctx),
    args.group ? resolveGroup(ctx, args.group) : Promise.resolve(undefined),
  ]);
  if (domains.includes(domain)) {
    throw new CliError(
      'USAGE',
      `${domain} is one of this tenant's verified domains, so ${email} is an internal user, not a guest`,
      {
        hint: `Find them with \`m365 entra users get ${email}\`, or create a member account with \`m365 entra users create --upn ${email} --display-name "..."\`.`,
      },
    );
  }
  return { email, group };
}

const inviteCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'users invite',
    summary: 'Invite an external person as a B2B guest (optionally straight into a group)',
    description:
      'Creates the guest account right away and returns its id, UPN (name_domain#EXT#@tenant) and ' +
      '`inviteRedeemUrl`. The invitation email is sent unless --no-email; then share the redeem URL yourself. ' +
      '`status` PendingAcceptance means the guest has not redeemed yet. Inviting an existing guest re-sends the ' +
      "invitation. Refuses addresses in the tenant's own verified domains (those are internal users). --group adds " +
      'the guest to a security or Microsoft 365 group; if that fails the guest still exists and the result says so. ' +
      'Fails when external collaboration settings block the domain or restrict who can invite.',
    examples: [
      'm365 entra users invite partner@fabrikam.com --name "Pat Partner" --group "Project X Guests"',
      'm365 entra users invite contractor@example.org --no-email',
    ],
    args: [
      { name: 'email', type: 'string', required: true, positional: true, description: "Guest's email address." },
      { name: 'name', type: 'string', description: 'Display name for the guest.' },
      { name: 'message', type: 'string', description: 'Custom text added to the invitation email.' },
      { name: 'noEmail', type: 'boolean', description: 'Do not send the invitation email (share inviteRedeemUrl).' },
      {
        name: 'redirectUrl',
        type: 'string',
        default: 'https://myapps.microsoft.com',
        description: 'Where the guest lands after redeeming.',
      },
      { name: 'group', type: 'string', description: 'Group to add the guest to (name, mail or id).' },
    ],
    risk: 'write',
    permissions: {
      delegated: ['User.Invite.All', 'User.Read', 'GroupMember.ReadWrite.All'],
      application: ['User.Invite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All'],
      roles: ['Guest Inviter'],
    },
    async preview(ctx, args) {
      const { email, group } = await planInvite(ctx, args);
      return compact({
        action: `Invite ${email} as a guest${group ? ` and add them to "${group.displayName}"` : ''}`,
        invitedUserEmailAddress: email,
        invitedUserDisplayName: args.name,
        sendInvitationMessage: !args.noEmail,
        inviteRedirectUrl: args.redirectUrl,
        group,
      });
    },
    async run(ctx, args) {
      const { email, group } = await planInvite(ctx, args);
      const invitation = await ctx.graph.post<Invitation>(
        '/invitations',
        compact({
          invitedUserEmailAddress: email,
          inviteRedirectUrl: args.redirectUrl,
          sendInvitationMessage: !args.noEmail,
          invitedUserDisplayName: args.name,
          invitedUserMessageInfo: args.message ? { customizedMessageBody: args.message } : undefined,
        }),
      );
      const userId = invitation.invitedUser?.id;
      let groupResult: unknown;
      if (group && userId) {
        groupResult = await ctx.graph.post(`/groups/${seg(group.id)}/members/$ref`, memberRef(ctx, userId)).then(
          () => ({ id: group.id, displayName: group.displayName, status: 'added' }),
          (error: unknown) =>
            ALREADY_MEMBER.test(errorText(error))
              ? { id: group.id, displayName: group.displayName, status: 'alreadyMember' }
              : {
                  id: group.id,
                  displayName: group.displayName,
                  status: 'failed',
                  error: errorText(error),
                  hint: `Guest created; retry with \`m365 entra groups add-member "${group.displayName}" ${userId}\`.`,
                },
        );
      }
      return compact({
        id: userId ?? null,
        userPrincipalName: invitation.invitedUser?.userPrincipalName ?? null,
        email,
        status: invitation.status ?? null,
        inviteRedeemUrl: invitation.inviteRedeemUrl ?? null,
        emailSent: !args.noEmail,
        group: groupResult,
      });
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Bulk user create
// ---------------------------------------------------------------------------------------------------------------

interface SkuInfo {
  skuId: string;
  skuPartNumber: string;
  consumedUnits: number;
  prepaidUnits?: { enabled?: number };
}

interface PlannedUser {
  line: number;
  displayName?: string;
  userPrincipalName?: string;
  mailNickname?: string;
  password?: string;
  usageLocation?: string;
  department?: string;
  jobTitle?: string;
  manager?: { id?: string; inFile?: string; label: string };
  license?: SkuInfo;
  groups: ResolvedGroup[];
  /** Existing account with this UPN: the row is skipped. */
  existingId?: string;
  errors: string[];
  warnings: string[];
}

export interface BulkCreateResult {
  line: number;
  userPrincipalName: string | null;
  status: 'created' | 'skipped' | 'failed';
  id?: string;
  temporaryPassword?: string;
  error?: string;
  followUpErrors?: string[];
}

async function loadSkus(ctx: Context): Promise<SkuInfo[]> {
  const { items } = await ctx.graph.list<SkuInfo>('/subscribedSkus', {
    query: { $select: 'skuId,skuPartNumber,consumedUnits,prepaidUnits' },
    limit: Number.POSITIVE_INFINITY,
    pageSize: false,
  });
  return items;
}

/** Read and validate every row without changing anything; `run` acts only on rows without errors. */
async function planBulkCreate(ctx: Context, path: string): Promise<PlannedUser[]> {
  const rows = readCsvRows(path, ['DisplayName', 'UserPrincipalName']);
  const firstLine = new Map<string, number>();
  const planned = rows.map((row) => {
    const plan: PlannedUser = {
      line: row.line,
      displayName: row.get('DisplayName'),
      userPrincipalName: row.get('UserPrincipalName', 'UPN'),
      mailNickname: row.get('MailNickname'),
      password: row.get('Password'),
      usageLocation: row.get('UsageLocation')?.toUpperCase(),
      department: row.get('Department'),
      jobTitle: row.get('JobTitle'),
      groups: [],
      errors: [],
      warnings: [],
    };
    const refs = {
      manager: row.get('Manager'),
      license: row.get('License', 'Sku'),
      groups: (row.get('Groups') ?? '')
        .split(';')
        .map((g) => g.trim())
        .filter(Boolean),
    };
    if (!plan.displayName) plan.errors.push('DisplayName is empty');
    const upn = plan.userPrincipalName;
    if (!upn) plan.errors.push('UserPrincipalName is empty');
    else if (!UPN_SHAPE.test(upn)) plan.errors.push(`UserPrincipalName "${upn}" is not name@domain`);
    else {
      const seen = firstLine.get(upn.toLowerCase());
      if (seen) plan.errors.push(`duplicate UserPrincipalName (also on line ${seen})`);
      else firstLine.set(upn.toLowerCase(), row.line);
    }
    if (plan.usageLocation && !/^[A-Z]{2}$/.test(plan.usageLocation)) {
      plan.errors.push(`UsageLocation "${plan.usageLocation}" is not a two-letter country code`);
    }
    if (refs.license && !plan.usageLocation) plan.errors.push('License needs a UsageLocation');
    return { plan, refs };
  });

  const checkable = planned.filter((p) => p.plan.userPrincipalName && UPN_SHAPE.test(p.plan.userPrincipalName));
  const fileUpns = new Set(checkable.map((p) => (p.plan.userPrincipalName as string).toLowerCase()));
  const managerRefs = planned.flatMap((p) =>
    p.refs.manager && !fileUpns.has(p.refs.manager.toLowerCase()) ? [p.refs.manager] : [],
  );
  const [existing, skus, managers, groups] = await Promise.all([
    ctx.graph.batch<{ id: string }>(
      checkable.map((p) => ({
        path: `/users/${seg(p.plan.userPrincipalName as string)}`,
        query: { $select: 'id,userPrincipalName' },
      })),
    ),
    planned.some((p) => p.refs.license) ? loadSkus(ctx) : Promise.resolve([]),
    resolveEach(managerRefs, (ref) => resolveUser(ctx, ref)),
    resolveEach(
      planned.flatMap((p) => p.refs.groups),
      (ref) => resolveGroup(ctx, ref),
    ),
  ]);

  const existingIds = new Map<string, string>();
  checkable.forEach(({ plan }, index) => {
    const result = existing[index];
    if (!result) return;
    if (!result.error) {
      plan.existingId = result.body.id;
      existingIds.set((plan.userPrincipalName as string).toLowerCase(), result.body.id);
    } else if (result.error.code !== 'NOT_FOUND') {
      plan.errors.push(`could not check whether the user exists: ${result.error.message}`);
    }
  });

  const validInFile = new Set(
    planned.filter((p) => !p.plan.errors.length).map((p) => (p.plan.userPrincipalName as string).toLowerCase()),
  );
  const seats = new Map<string, number>();
  for (const { plan, refs } of planned) {
    if (refs.manager) {
      const key = refs.manager.toLowerCase();
      if (fileUpns.has(key)) {
        const existingId = existingIds.get(key);
        if (existingId) plan.manager = { id: existingId, label: refs.manager };
        else if (validInFile.has(key)) plan.manager = { inFile: key, label: refs.manager };
        else plan.errors.push(`Manager ${refs.manager} is a row in this file that has errors`);
      } else {
        const manager = managers.get(key);
        if (!manager || isFailure(manager))
          plan.errors.push(`Manager "${refs.manager}": ${manager?.error ?? 'not found'}`);
        else plan.manager = { id: manager.id, label: manager.userPrincipalName };
      }
    }
    if (refs.license) {
      const wanted = refs.license.toLowerCase();
      const sku = skus.find((s) => s.skuPartNumber.toLowerCase() === wanted || s.skuId.toLowerCase() === wanted);
      if (!sku) {
        plan.errors.push(
          `License "${refs.license}" is not a SKU in this tenant (have: ${skus.map((s) => s.skuPartNumber).join(', ') || 'none'})`,
        );
      } else {
        plan.license = sku;
      }
    }
    for (const ref of refs.groups) {
      const group = groups.get(ref.toLowerCase());
      if (!group || isFailure(group)) plan.errors.push(`Group "${ref}": ${group?.error ?? 'not found'}`);
      else plan.groups.push(group);
    }
    if (plan.license && !plan.errors.length && !plan.existingId) {
      seats.set(plan.license.skuId, (seats.get(plan.license.skuId) ?? 0) + 1);
    }
  }
  for (const sku of skus) {
    const needed = seats.get(sku.skuId) ?? 0;
    const free = (sku.prepaidUnits?.enabled ?? 0) - sku.consumedUnits;
    if (needed > free) {
      for (const { plan } of planned) {
        if (plan.license?.skuId === sku.skuId && !plan.existingId) {
          plan.warnings.push(`${sku.skuPartNumber}: ${needed} needed, ${free} seats free; some assignments will fail`);
        }
      }
    }
  }
  return planned.map((p) => p.plan);
}

function rowStatus(plan: PlannedUser): 'valid' | 'skipped' | 'invalid' {
  if (plan.errors.length) return 'invalid';
  return plan.existingId ? 'skipped' : 'valid';
}

const bulkUserCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'users bulk-create',
    summary: 'Create many cloud users from a CSV, with manager, license and groups per row',
    description:
      'CSV columns: DisplayName, UserPrincipalName (required); optional MailNickname, Password, UsageLocation, ' +
      'Department, JobTitle, Manager (UPN/name, may be another row in the file), License (SKU part number, e.g. ' +
      'SPE_E3; needs UsageLocation), Groups (semicolon-separated names or ids). Max 500 rows. Always run with ' +
      '--dry-run first: it validates every row (missing fields, duplicate UPNs, existing accounts, unknown SKU, ' +
      'manager or group, seat shortfall) and creates nothing. Rows whose UPN already exists are skipped; rows with ' +
      'errors fail; the rest are created. Generated passwords are returned once per row as temporaryPassword ' +
      '(passwords from the file are never echoed); users must change them at first sign-in. A row whose manager, ' +
      'license or group step fails is still created and lists followUpErrors.',
    examples: [
      'm365 --dry-run entra users bulk-create new-starters.csv',
      'm365 entra users bulk-create new-starters.csv',
    ],
    args: [{ name: 'file', type: 'string', required: true, positional: true, description: 'Path to the CSV file.' }],
    risk: 'write',
    permissions: {
      delegated: ['User.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'GroupMember.ReadWrite.All'],
      application: ['User.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'GroupMember.ReadWrite.All'],
      roles: ['User Administrator'],
    },
    async preview(ctx, args) {
      const plans = await planBulkCreate(ctx, args.file);
      const rows = plans.map((p) =>
        compact({
          line: p.line,
          userPrincipalName: p.userPrincipalName ?? null,
          displayName: p.displayName,
          status: rowStatus(p),
          license: p.license?.skuPartNumber,
          manager: p.manager?.label,
          groups: p.groups.length ? p.groups.map((g) => g.displayName) : undefined,
          password: p.errors.length || p.existingId ? undefined : p.password ? 'from file' : 'generated when run',
          errors: p.errors.length ? p.errors : undefined,
          warnings: p.warnings.length ? p.warnings : undefined,
          existingId: p.existingId,
        }),
      );
      const count = (status: string) => rows.filter((r) => r.status === status).length;
      const summary = {
        rows: rows.length,
        valid: count('valid'),
        skipped: count('skipped'),
        invalid: count('invalid'),
      };
      return {
        action: `Create ${summary.valid} user(s) from ${args.file} (${summary.skipped} already exist, ${summary.invalid} with errors)`,
        summary,
        rows,
      };
    },
    async run(ctx, args) {
      const plans = await planBulkCreate(ctx, args.file);
      const results = new Map<PlannedUser, BulkCreateResult>();
      const toCreate: { plan: PlannedUser; password: string }[] = [];
      for (const plan of plans) {
        const userPrincipalName = plan.userPrincipalName ?? null;
        if (plan.errors.length) {
          results.set(plan, { line: plan.line, userPrincipalName, status: 'failed', error: plan.errors.join('; ') });
        } else if (plan.existingId) {
          results.set(plan, {
            line: plan.line,
            userPrincipalName,
            status: 'skipped',
            id: plan.existingId,
            error: 'already exists',
          });
        } else {
          toCreate.push({ plan, password: plan.password ?? generatePassword() });
        }
      }

      const created = await ctx.graph.batch<{ id: string }>(
        toCreate.map(({ plan, password }) => ({
          method: 'POST',
          path: '/users',
          body: compact({
            accountEnabled: true,
            displayName: plan.displayName,
            userPrincipalName: plan.userPrincipalName,
            // mailNickname rejects characters such as @ ( ) [ ] ; : < > , and spaces.
            mailNickname:
              plan.mailNickname ??
              (plan.userPrincipalName?.split('@')[0] as string).replace(/[^A-Za-z0-9._\-'!#^~]/g, ''),
            passwordProfile: { password, forceChangePasswordNextSignIn: true },
            usageLocation: plan.usageLocation,
            department: plan.department,
            jobTitle: plan.jobTitle,
          }),
        })),
      );
      const newIds = new Map<string, string>();
      toCreate.forEach(({ plan, password }, index) => {
        const response = created[index];
        const userPrincipalName = plan.userPrincipalName as string;
        if (!response || response.error) {
          results.set(plan, {
            line: plan.line,
            userPrincipalName,
            status: 'failed',
            error: response?.error?.message ?? 'no response',
          });
          return;
        }
        newIds.set(userPrincipalName.toLowerCase(), response.body.id);
        results.set(plan, {
          line: plan.line,
          userPrincipalName,
          status: 'created',
          id: response.body.id,
          ...(plan.password ? {} : { temporaryPassword: password }),
        });
      });

      // Manager, license and group steps for every created user, in one batch run.
      const followUps: { plan: PlannedUser; step: string; method: string; path: string; body: unknown }[] = [];
      for (const { plan } of toCreate) {
        const userId = newIds.get((plan.userPrincipalName as string).toLowerCase());
        if (!userId) continue;
        const managerId = plan.manager?.id ?? (plan.manager?.inFile ? newIds.get(plan.manager.inFile) : undefined);
        if (plan.manager && !managerId) {
          const result = results.get(plan) as BulkCreateResult;
          result.followUpErrors = [...(result.followUpErrors ?? []), `manager: ${plan.manager.label} was not created`];
        }
        if (managerId) {
          followUps.push({
            plan,
            step: `manager ${plan.manager?.label}`,
            method: 'PUT',
            path: `/users/${seg(userId)}/manager/$ref`,
            body: memberRef(ctx, managerId),
          });
        }
        if (plan.license) {
          followUps.push({
            plan,
            step: `license ${plan.license.skuPartNumber}`,
            method: 'POST',
            path: `/users/${seg(userId)}/assignLicense`,
            body: { addLicenses: [{ skuId: plan.license.skuId, disabledPlans: [] }], removeLicenses: [] },
          });
        }
        for (const group of plan.groups) {
          followUps.push({
            plan,
            step: `group ${group.displayName}`,
            method: 'POST',
            path: `/groups/${seg(group.id)}/members/$ref`,
            body: memberRef(ctx, userId),
          });
        }
      }
      const followUpResults = await ctx.graph.batch(
        followUps.map(({ method, path, body }) => ({ method, path, body })),
      );
      followUps.forEach(({ plan, step }, index) => {
        const error = followUpResults[index]?.error;
        if (!error) return;
        const result = results.get(plan) as BulkCreateResult;
        result.followUpErrors = [...(result.followUpErrors ?? []), `${step}: ${error.message}`];
      });

      const ordered = plans.map((plan) => results.get(plan) as BulkCreateResult);
      const count = (status: BulkCreateResult['status']) => ordered.filter((r) => r.status === status).length;
      return {
        summary: { created: count('created'), skipped: count('skipped'), failed: count('failed') },
        results: ordered,
        note: ordered.some((r) => r.temporaryPassword)
          ? 'temporaryPassword values are shown once; hand them to each user securely.'
          : undefined,
      };
    },
  }),
];

// ---------------------------------------------------------------------------------------------------------------
// Bulk group membership
// ---------------------------------------------------------------------------------------------------------------

interface MemberRow {
  line: number;
  value: string | null;
  member?: ResolvedMember;
  error?: string;
}

async function planAddMembers(ctx: Context, args: { group: string; fromCsv: string; column: string }) {
  const rows = readCsvRows(args.fromCsv, [args.column]);
  const group = await resolveGroup(ctx, args.group);
  const planned: MemberRow[] = rows.map((row) => {
    const value = row.get(args.column) ?? null;
    return value ? { line: row.line, value } : { line: row.line, value, error: `${args.column} is empty` };
  });
  // UPNs and ids resolve in one $batch; anything else (names, prefixes, devices, groups) goes through resolveMember.
  const direct = planned.filter((r) => r.value && (r.value.includes('@') || isGuid(r.value)) && !r.value.includes(':'));
  const found = await ctx.graph.batch<{ id: string; displayName: string; userPrincipalName: string }>(
    direct.map((r) => ({
      path: `/users/${seg(r.value as string)}`,
      query: { $select: 'id,displayName,userPrincipalName' },
    })),
  );
  direct.forEach((row, index) => {
    const result = found[index];
    if (result && !result.error) {
      row.member = {
        id: result.body.id,
        type: 'user',
        displayName: result.body.displayName,
        detail: result.body.userPrincipalName,
      };
    }
  });
  const rest = planned.filter((r) => r.value && !r.member);
  const resolved = await resolveEach(
    rest.map((r) => r.value as string),
    (ref) => resolveMember(ctx, ref),
  );
  for (const row of rest) {
    const member = resolved.get((row.value as string).toLowerCase());
    if (!member || isFailure(member)) row.error = member?.error ?? 'not found';
    else row.member = member;
  }
  return { group, rows: planned };
}

const bulkGroupCommands: AnyCommand[] = [
  defineCommand({
    domain: 'entra',
    name: 'groups add-members',
    summary: 'Add every member listed in a CSV column to a group, reporting each row',
    description:
      'Reads --column (default UserPrincipalName) from the CSV; values may be UPNs, emails, names or ids, and ' +
      'user:/group:/device: prefixes work as in `groups add-member`. Adds through Graph $batch (20 per request). ' +
      'Each row is added, alreadyMember or failed (unresolvable or rejected); one bad row never stops the rest. ' +
      'Run with --dry-run first to see unresolved rows. Distribution and mail-enabled security groups are managed ' +
      'in Exchange Online. Max 500 rows.',
    examples: [
      'm365 --dry-run entra groups add-members "Finance Team" --from-csv finance.csv',
      'm365 entra groups add-members "Project X" --from-csv people.csv --column Email',
    ],
    args: [
      {
        name: 'group',
        type: 'string',
        required: true,
        positional: true,
        description: 'Group name, mail address or id.',
      },
      { name: 'fromCsv', type: 'string', required: true, description: 'CSV file with one member per row.' },
      { name: 'column', type: 'string', default: 'UserPrincipalName', description: 'Column holding the member.' },
    ],
    risk: 'write',
    permissions: {
      delegated: ['GroupMember.ReadWrite.All', 'User.Read.All'],
      application: ['GroupMember.ReadWrite.All', 'User.Read.All'],
      roles: ['Groups Administrator'],
    },
    async preview(ctx, args) {
      const { group, rows } = await planAddMembers(ctx, args);
      const members = rows.filter((r) => r.member);
      return {
        action: `Add ${members.length} member(s) from ${args.fromCsv} to group "${group.displayName}"`,
        group,
        summary: { rows: rows.length, resolved: members.length, unresolved: rows.length - members.length },
        members: members.map((r) => ({ line: r.line, ...r.member })),
        unresolved: rows.filter((r) => !r.member).map((r) => ({ line: r.line, value: r.value, error: r.error })),
      };
    },
    async run(ctx, args) {
      const { group, rows } = await planAddMembers(ctx, args);
      const ids = [...new Set(rows.flatMap((r) => (r.member ? [r.member.id] : [])))];
      const responses = await ctx.graph.batch(
        ids.map((id) => ({ method: 'POST', path: `/groups/${seg(group.id)}/members/$ref`, body: memberRef(ctx, id) })),
      );
      const outcome = new Map(ids.map((id, index) => [id, responses[index]]));
      const results = rows.map((row) => {
        if (!row.member) return { line: row.line, member: row.value, status: 'failed', error: row.error };
        const response = outcome.get(row.member.id);
        const base = {
          line: row.line,
          member: row.value,
          id: row.member.id,
          type: row.member.type,
          displayName: row.member.displayName,
        };
        if (!response?.error) return { ...base, status: 'added' };
        if (response.status === 400 && ALREADY_MEMBER.test(response.error.message))
          return { ...base, status: 'alreadyMember' };
        return { ...base, status: 'failed', error: response.error.message };
      });
      const count = (status: string) => results.filter((r) => r.status === status).length;
      return {
        group: { id: group.id, displayName: group.displayName },
        summary: { added: count('added'), alreadyMember: count('alreadyMember'), failed: count('failed') },
        results,
      };
    },
  }),
];

export const commands: AnyCommand[] = [...inviteCommands, ...bulkUserCommands, ...bulkGroupCommands];
