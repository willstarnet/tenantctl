import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { type ResolvedUser, resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { bytesToGb, compact, daysSince } from '../../core/util.ts';
import type { Permission } from './files.ts';
import { type Drive, type IdentitySet, shapeQuota, who } from './resolve.ts';

const DRIVE_SELECT = 'id,driveType,webUrl,owner,quota,createdDateTime,lastModifiedDateTime';

const HANDOVER_NOTE =
  'Delegated callers can only change sharing on OneDrives they can already open (being SharePoint ' +
  'Administrator is not enough); use an app-only profile with Files.ReadWrite.All for handovers.';

/** The owner's OneDrive, or NOT_FOUND explaining why it may be gone. */
async function ownerDrive(ctx: Context, ownerRef: string): Promise<{ owner: ResolvedUser; drive: Drive }> {
  const owner = await resolveUser(ctx, ownerRef);
  const drive = await ctx.graph.tryGet<Drive>(`/users/${seg(owner.id)}/drive`, { query: { $select: DRIVE_SELECT } });
  if (!drive) {
    throw new CliError('NOT_FOUND', `${owner.userPrincipalName} has no OneDrive`, {
      hint:
        'It was never provisioned, or the user is deleted and the OneDrive is past its retention period. A ' +
        "deleted user's OneDrive is reachable here only while the user is restorable; afterwards use the " +
        'Microsoft 365 admin center (Users > Deleted users) or SharePoint admin center.',
    });
  }
  return { owner, drive };
}

function identities(permission: Permission): IdentitySet[] {
  return [
    ...(permission.grantedToV2 ? [permission.grantedToV2] : permission.grantedTo ? [permission.grantedTo] : []),
    ...(permission.grantedToIdentitiesV2 ?? permission.grantedToIdentities ?? []),
  ];
}

/** True when the identity set is this user (by object id, email or the UPN inside a SharePoint claims login). */
function isUser(set: IdentitySet, user: ResolvedUser): boolean {
  const names = [user.userPrincipalName, user.mail].filter(Boolean).map((n) => (n as string).toLowerCase());
  return [set.user, set.siteUser].some((identity) => {
    if (!identity) return false;
    if (identity.id && identity.id.toLowerCase() === user.id.toLowerCase()) return true;
    const email = identity.email?.toLowerCase();
    const login = identity.loginName?.split('|').pop()?.toLowerCase();
    return names.some((n) => n === email || n === login);
  });
}

interface RevokePlan {
  owner: ResolvedUser;
  grantee: ResolvedUser;
  drive: Drive;
  remove: Permission[];
  /** Sharing links that also include other people: deleting them would cut those people off too. */
  shared: Permission[];
}

async function planRevoke(ctx: Context, ownerRef: string, granteeRef: string): Promise<RevokePlan> {
  const [{ owner, drive }, grantee] = await Promise.all([ownerDrive(ctx, ownerRef), resolveUser(ctx, granteeRef)]);
  const { items } = await ctx.graph.list<Permission>(`/drives/${seg(drive.id)}/items/root/permissions`, {
    pageSize: false,
    limit: Number.POSITIVE_INFINITY,
  });
  const remove: Permission[] = [];
  const shared: Permission[] = [];
  for (const permission of items) {
    if (permission.inheritedFrom || permission.roles?.includes('owner')) continue;
    const sets = identities(permission);
    const invited = permission.invitation?.email?.toLowerCase();
    const matches =
      sets.some((set) => isUser(set, grantee)) ||
      (invited !== undefined && [grantee.userPrincipalName, grantee.mail].some((n) => n?.toLowerCase() === invited));
    if (!matches) continue;
    if (sets.length > 1 && sets.some((set) => !isUser(set, grantee))) shared.push(permission);
    else remove.push(permission);
  }
  return { owner, grantee, drive, remove, shared };
}

function describePermission(permission: Permission) {
  return compact({
    id: permission.id,
    roles: permission.roles ?? [],
    kind: permission.link ? 'link' : 'direct',
    linkScope: permission.link?.scope,
  });
}

const PERIOD_ARG = {
  name: 'period',
  type: 'string',
  choices: ['D7', 'D30', 'D90', 'D180'],
  default: 'D7',
  description: 'Activity window the report aggregates over.',
} as const;
const NEAR_QUOTA_ARG = {
  name: 'nearQuota',
  type: 'boolean',
  description: 'Only accounts/sites using 90% or more of their allocated storage.',
} as const;

const CONCEALED_HINT =
  'Usage reports are showing concealed (hashed) names. A Global Administrator can turn off "Display concealed ' +
  'user, group, and site names in all reports" in Microsoft 365 admin center > Settings > Org settings > Reports.';

const REPORT_NOTE =
  'Usage reports lag 24-48 hours. If the tenant conceals names in reports, owners and URLs appear hashed and ' +
  'filters cannot match them (turn it off in Microsoft 365 admin center > Settings > Org settings > Reports). ' +
  'Delegated callers need a reports role such as Reports Reader.';

function storage(row: Record<string, string>) {
  const used = Number(row['Storage Used (Byte)']);
  const allocated = Number(row['Storage Allocated (Byte)']);
  const lastActivity = row['Last Activity Date'] || null;
  return {
    usedGb: bytesToGb(used),
    allocatedGb: bytesToGb(allocated),
    percentUsed: allocated > 0 ? Math.round((used / allocated) * 1000) / 10 : null,
    fileCount: Number(row['File Count']) || 0,
    activeFileCount: Number(row['Active File Count']) || 0,
    lastActivityDate: lastActivity,
    daysSinceActivity: daysSince(lastActivity),
    isDeleted: row['Is Deleted'] === 'True',
  };
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'sharepoint',
    name: 'onedrive get',
    summary: "Show a user's OneDrive: URL, storage used/remaining, quota state",
    description:
      'Answers "is Jane\'s OneDrive full", "where is Bob\'s OneDrive". state: normal, nearing (<10% left), ' +
      'critical (<1% left), exceeded (uploads blocked). recycleBinGb counts toward usage. NOT_FOUND usually means ' +
      'the OneDrive was never provisioned (user never opened it) or the user has no SharePoint/OneDrive license. ' +
      'Next: `sharepoint onedrive usage --user <upn>` for file counts and last activity.',
    examples: ['m365 sharepoint onedrive get jane@contoso.com'],
    args: [{ name: 'user', type: 'string', required: true, positional: true, description: 'UPN, email, name or id.' }],
    risk: 'read',
    permissions: {
      delegated: ['Files.Read.All', 'User.ReadBasic.All'],
      application: ['Files.Read.All', 'User.Read.All'],
    },
    async run(ctx, args) {
      const user = await resolveUser(ctx, args.user);
      const drive = await ctx.graph.tryGet<Drive>(`/users/${user.id}/drive`, {
        query: { $select: 'id,driveType,webUrl,owner,quota,createdDateTime,lastModifiedDateTime' },
      });
      if (!drive) {
        throw new CliError('NOT_FOUND', `${user.userPrincipalName} has no OneDrive`, {
          hint:
            'It is created the first time the user opens OneDrive and needs a license with SharePoint/OneDrive. ' +
            'Check licenses with the entra user commands.',
        });
      }
      if (ctx.full) return drive;
      return {
        user: user.userPrincipalName,
        driveId: drive.id,
        webUrl: drive.webUrl ?? null,
        owner: who(drive.owner),
        ...shapeQuota(drive.quota),
        createdDateTime: drive.createdDateTime ?? null,
        lastModifiedDateTime: drive.lastModifiedDateTime ?? null,
        daysSinceModified: daysSince(drive.lastModifiedDateTime),
      };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'onedrive usage',
    summary: 'OneDrive storage and activity per account from the usage report',
    description:
      'Tenant-wide report: storage used/allocated, file counts and last activity per OneDrive. Use --near-quota ' +
      'to find accounts at 90%+ (sorted fullest first), --user for one person. ' +
      `For a live single-user view use \`onedrive get\`. ${REPORT_NOTE}`,
    examples: ['m365 sharepoint onedrive usage --near-quota', 'm365 sharepoint onedrive usage --user jane@contoso.com'],
    args: [
      { name: 'user', type: 'string', description: 'Only this owner (UPN or display name).' },
      NEAR_QUOTA_ARG,
      PERIOD_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['Reports.Read.All'], application: ['Reports.Read.All'], roles: ['Reports Reader'] },
    async run(ctx, args) {
      const rows = await ctx.graph.report(`/reports/getOneDriveUsageAccountDetail(period='${args.period}')`);
      const concealed = rows.length > 0 && !rows.some((r) => r['Owner Principal Name']?.includes('@'));
      if (concealed && args.user)
        throw new CliError('CONFIG', 'Cannot filter by user: report names are concealed', { hint: CONCEALED_HINT });
      if (concealed) ctx.log(CONCEALED_HINT);
      const wanted = args.user?.toLowerCase();
      let items = rows
        .filter(
          (r) =>
            !wanted ||
            r['Owner Principal Name']?.toLowerCase() === wanted ||
            r['Owner Display Name']?.toLowerCase() === wanted,
        )
        .map((r) => ({
          owner: r['Owner Display Name'] || null,
          ownerPrincipalName: r['Owner Principal Name'] || null,
          url: r['Site URL'] || null,
          ...storage(r),
        }));
      if (args.nearQuota) {
        items = items
          .filter((i) => (i.percentUsed ?? 0) >= 90)
          .sort((a, b) => (b.percentUsed ?? 0) - (a.percentUsed ?? 0));
      }
      const limit = listLimit(args);
      return { items: items.slice(0, limit), truncated: items.length > limit };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'sites usage',
    summary: 'SharePoint site storage, file counts and last activity from the usage report',
    description:
      'Tenant-wide report sorted by storage used: find full sites (--near-quota), inactive sites ' +
      '(daysSinceActivity) or one site (--site, matched against the URL). Excludes OneDrive (see `onedrive ' +
      `usage\`). ${REPORT_NOTE}`,
    examples: ['m365 sharepoint sites usage --near-quota', 'm365 sharepoint sites usage --site /sites/finance'],
    args: [
      { name: 'site', type: 'string', description: 'Only sites whose URL contains this text.' },
      NEAR_QUOTA_ARG,
      PERIOD_ARG,
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['Reports.Read.All'], application: ['Reports.Read.All'], roles: ['Reports Reader'] },
    async run(ctx, args) {
      const rows = await ctx.graph.report(`/reports/getSharePointSiteUsageDetail(period='${args.period}')`);
      const concealed = rows.length > 0 && !rows.some((r) => r['Site URL']?.startsWith('http'));
      if (concealed && args.site)
        throw new CliError('CONFIG', 'Cannot filter by site: report names are concealed', { hint: CONCEALED_HINT });
      if (concealed) ctx.log(CONCEALED_HINT);
      const wanted = args.site?.toLowerCase();
      let items = rows
        .filter((r) => !wanted || r['Site URL']?.toLowerCase().includes(wanted))
        .map((r) => ({
          url: r['Site URL'] || null,
          siteId: r['Site Id'] || null,
          owner: r['Owner Display Name'] || null,
          ownerPrincipalName: r['Owner Principal Name'] || null,
          template: r['Root Web Template'] || null,
          pageViews: Number(r['Page View Count']) || 0,
          ...storage(r),
        }))
        .sort((a, b) => (b.usedGb ?? 0) - (a.usedGb ?? 0));
      if (args.nearQuota) items = items.filter((i) => (i.percentUsed ?? 0) >= 90);
      const limit = listLimit(args);
      return { items: items.slice(0, limit), truncated: items.length > limit };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'onedrive grant',
    summary: "Give someone access to another user's whole OneDrive (departing-employee handover)",
    description:
      "Shares the owner's OneDrive root folder, i.e. all of their files, with the grantee (sign-in required) and " +
      'returns the drive URL to send them. It does NOT make the grantee a site collection admin of the OneDrive ' +
      '(for that use SharePoint admin center or `Set-SPOUser -IsSiteCollectionAdmin` / PnP): folders or files ' +
      'inside that have their own (broken-inheritance) permissions are not covered, and the grantee cannot ' +
      "manage OneDrive settings. The drive must still exist: a deleted user's OneDrive is retained only for the " +
      'OneDrive ' +
      'retention period (30 days by default after deletion). --notify emails the grantee a sharing link; ' +
      '--expires-days is honoured by SharePoint only for that link-based invitation, so pair it with --notify. ' +
      `Undo with \`onedrive revoke <owner> <grantee>\`. ${HANDOVER_NOTE}`,
    examples: [
      'm365 sharepoint onedrive grant bob@contoso.com manager@contoso.com',
      'm365 sharepoint onedrive grant bob@contoso.com manager@contoso.com --role write --notify --expires-days 30',
    ],
    args: [
      {
        name: 'owner',
        type: 'string',
        required: true,
        positional: true,
        description: 'OneDrive owner: UPN, email, name or id.',
      },
      {
        name: 'grantee',
        type: 'string',
        required: true,
        positional: true,
        description: 'Who gets access: UPN, email, name or id.',
      },
      {
        name: 'role',
        type: 'string',
        choices: ['read', 'write'],
        default: 'read',
        description: 'read = view/download; write = also edit/delete.',
      },
      { name: 'notify', type: 'boolean', description: 'Email the grantee a sharing invitation.' },
      { name: 'expiresDays', type: 'number', description: 'Access expires after N days (only with --notify).' },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['Files.ReadWrite.All', 'User.ReadBasic.All'],
      application: ['Files.ReadWrite.All', 'User.Read.All'],
    },
    async preview(ctx, args) {
      const [{ owner, drive }, grantee] = await Promise.all([
        ownerDrive(ctx, args.owner),
        resolveUser(ctx, args.grantee),
      ]);
      const quota = shapeQuota(drive.quota);
      return {
        action:
          `Give ${grantee.userPrincipalName} ${args.role} access to all files in ${owner.userPrincipalName}'s ` +
          `OneDrive (${drive.webUrl ?? drive.id})${args.notify ? ' and email them an invitation' : ''}`,
        owner: owner.userPrincipalName,
        driveUrl: drive.webUrl ?? null,
        usedGb: quota?.usedGb ?? null,
        grantee: grantee.userPrincipalName,
        role: args.role,
        notify: args.notify,
        expiresDays: args.expiresDays ?? null,
      };
    },
    async run(ctx, args) {
      const [{ owner, drive }, grantee] = await Promise.all([
        ownerDrive(ctx, args.owner),
        resolveUser(ctx, args.grantee),
      ]);
      const expires =
        args.expiresDays && args.expiresDays > 0
          ? new Date(Date.now() + args.expiresDays * 86_400_000).toISOString()
          : undefined;
      const response = await ctx.graph.post<{ value?: Permission[] }>(
        `/drives/${seg(drive.id)}/items/root/invite`,
        compact({
          recipients: [{ email: grantee.mail ?? grantee.userPrincipalName }],
          requireSignIn: true,
          sendInvitation: args.notify,
          roles: [args.role],
          expirationDateTime: expires,
        }),
      );
      const permission = response.value?.[0];
      if (ctx.full) return response;
      return {
        owner: owner.userPrincipalName,
        grantee: grantee.userPrincipalName,
        permissionId: permission?.id ?? null,
        roles: permission?.roles ?? [args.role],
        driveUrl: drive.webUrl ?? null,
        expires: permission?.expirationDateTime ?? null,
        next: `Send ${grantee.userPrincipalName} the driveUrl; undo with: m365 sharepoint onedrive revoke ${owner.userPrincipalName} ${grantee.userPrincipalName}`,
      };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'onedrive revoke',
    summary: "Remove someone's access to another user's OneDrive root (undo onedrive grant)",
    description:
      "Deletes the grantee's own permissions on the OneDrive root folder (direct grants and links only they " +
      'hold). Links that also include other people are listed under `sharedLinks` and left alone; remove them ' +
      'with `files unshare`. Access the grantee has to individual files shared separately is not touched (see ' +
      `\`files permissions\`). ${HANDOVER_NOTE}`,
    examples: ['m365 sharepoint onedrive revoke bob@contoso.com manager@contoso.com'],
    args: [
      {
        name: 'owner',
        type: 'string',
        required: true,
        positional: true,
        description: 'OneDrive owner: UPN, email, name or id.',
      },
      { name: 'grantee', type: 'string', required: true, positional: true, description: 'Whose access to remove.' },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['Files.ReadWrite.All', 'User.ReadBasic.All'],
      application: ['Files.ReadWrite.All', 'User.Read.All'],
    },
    async preview(ctx, args) {
      const plan = await planRevoke(ctx, args.owner, args.grantee);
      return {
        action: plan.remove.length
          ? `Remove ${plan.remove.length} permission(s) giving ${plan.grantee.userPrincipalName} access to ` +
            `${plan.owner.userPrincipalName}'s OneDrive`
          : `Nothing to remove: ${plan.grantee.userPrincipalName} has no direct access to ${plan.owner.userPrincipalName}'s OneDrive root`,
        owner: plan.owner.userPrincipalName,
        driveUrl: plan.drive.webUrl ?? null,
        grantee: plan.grantee.userPrincipalName,
        remove: plan.remove.map(describePermission),
        sharedLinks: plan.shared.map(describePermission),
      };
    },
    async run(ctx, args) {
      const plan = await planRevoke(ctx, args.owner, args.grantee);
      for (const permission of plan.remove) {
        await ctx.graph.delete(`/drives/${seg(plan.drive.id)}/items/root/permissions/${seg(permission.id)}`);
      }
      return {
        owner: plan.owner.userPrincipalName,
        grantee: plan.grantee.userPrincipalName,
        removed: plan.remove.map((p) => p.id),
        sharedLinks: plan.shared.map(describePermission),
      };
    },
  }),
];
