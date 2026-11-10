import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { settle } from '../../core/util.ts';
import {
  type DriveItem,
  type Identity,
  type IdentitySet,
  itemPath,
  type ResolvedDriveItem,
  resolveDriveItem,
  resolveSite,
  who,
} from './resolve.ts';

const REF_ARG = {
  name: 'ref',
  type: 'string',
  required: true,
  positional: true,
  description: 'File/folder URL or sharing link, driveId:itemId, or a library path when --site is given.',
} as const;
const SITE_OPT = {
  name: 'site',
  type: 'string',
  description: 'Site whose default library contains <ref> as a path, e.g. "Reports/Q1.xlsx".',
} as const;

const REF_NOTE =
  'URLs are resolved through the /shares API, which needs Files.ReadWrite.All or Sites.ReadWrite.All even ' +
  'for reads; --site <site> with a path or driveId:itemId avoids that. Delegated sign-ins only see files the ' +
  'signed-in user can open (SharePoint Administrator does not grant file access); use app-only for tenant-wide.';

export interface Permission {
  id: string;
  roles?: string[];
  link?: { scope?: string; type?: string; webUrl?: string; preventsDownload?: boolean };
  inheritedFrom?: { path?: string; id?: string; driveId?: string } | null;
  expirationDateTime?: string | null;
  hasPassword?: boolean;
  grantedToV2?: IdentitySet;
  grantedToIdentitiesV2?: IdentitySet[];
  grantedTo?: IdentitySet;
  grantedToIdentities?: IdentitySet[];
  invitation?: { email?: string };
}

export interface Grantee {
  type: 'user' | 'group' | 'sharePointGroup' | 'application' | 'device';
  displayName: string | null;
  email: string | null;
  external?: boolean;
}

interface DriveItemVersion {
  id: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: IdentitySet;
  size?: number;
}

interface SearchResponse {
  value?: {
    hitsContainers?: { total?: number; moreResultsAvailable?: boolean; hits?: { resource?: DriveItem }[] }[];
  }[];
}

// SharePoint stores "no expiry" as DateTime.MinValue.
const NO_EXPIRY = /^0001-01-01/;

function toGrantee(set: IdentitySet): Grantee | null {
  // A SharePoint principal is usually reported twice (user + siteUser); the first present identity wins.
  const pairs: [Grantee['type'], Identity | undefined][] = [
    ['user', set.user],
    ['group', set.group],
    ['sharePointGroup', set.siteGroup],
    ['user', set.siteUser],
    ['application', set.application],
    ['device', set.device],
  ];
  const found = pairs.find(([, identity]) => identity);
  if (!found) return null;
  const [type, identity] = found as [Grantee['type'], Identity];
  // loginName is a claims string such as "i:0#.f|membership|bob@contoso.com".
  const email = identity.email ?? (identity.loginName?.includes('@') ? identity.loginName.split('|').pop() : undefined);
  return { type, displayName: identity.displayName ?? null, email: email ?? null };
}

function isExternal(email: string | null, tenantDomains: readonly string[] | undefined): boolean | undefined {
  if (!email) return undefined;
  if (email.toLowerCase().includes('#ext#')) return true;
  if (!tenantDomains) return undefined;
  const domain = email.split('@')[1]?.toLowerCase();
  return domain ? !tenantDomains.includes(domain) : undefined;
}

/**
 * One permission on a driveItem as an admin reads it: link (scope anonymous = anyone with the link,
 * organization = anyone in the tenant, users = specific people, existingAccess = no new access), direct grant,
 * or inherited from a parent folder/library. With `tenantDomains` (verified domains, lowercase) each grantee
 * gets `external`.
 */
export function shapePermission(permission: Permission, tenantDomains?: readonly string[]) {
  const sets = [
    ...(permission.grantedToV2 ? [permission.grantedToV2] : permission.grantedTo ? [permission.grantedTo] : []),
    ...(permission.grantedToIdentitiesV2 ?? permission.grantedToIdentities ?? []),
  ];
  const grantedTo: Grantee[] = [];
  for (const set of sets) {
    const grantee = toGrantee(set);
    if (grantee) grantedTo.push(grantee);
  }
  const invited = permission.invitation?.email;
  if (invited && !grantedTo.some((g) => g.email?.toLowerCase() === invited.toLowerCase())) {
    grantedTo.push({ type: 'user', displayName: null, email: invited });
  }
  for (const grantee of grantedTo) {
    const external = isExternal(grantee.email, tenantDomains);
    if (external !== undefined) grantee.external = external;
  }
  const inherited = Boolean(permission.inheritedFrom);
  const expires = permission.expirationDateTime;
  return {
    id: permission.id,
    kind: inherited ? 'inherited' : permission.link ? 'link' : 'direct',
    roles: permission.roles ?? [],
    linkScope: permission.link?.scope ?? null,
    linkType: permission.link?.type ?? null,
    linkUrl: permission.link?.webUrl ?? null,
    preventsDownload: permission.link?.preventsDownload,
    inheritedFrom: inherited ? (permission.inheritedFrom?.path ?? permission.inheritedFrom?.id ?? null) : null,
    expires: expires && !NO_EXPIRY.test(expires) ? expires : null,
    hasPassword: permission.hasPassword,
    grantedTo,
  };
}

function shapeItem(item: DriveItem) {
  return {
    id: item.id,
    driveId: item.parentReference?.driveId ?? null,
    name: item.name,
    path: itemPath(item),
    webUrl: item.webUrl ?? null,
    size: item.size ?? null,
    lastModifiedBy: who(item.lastModifiedBy),
    lastModifiedDateTime: item.lastModifiedDateTime ?? null,
  };
}

function itemUrl(target: ResolvedDriveItem): string {
  return `/drives/${seg(target.driveId)}/items/${seg(target.itemId)}`;
}

async function listPermissions(ctx: Context, target: ResolvedDriveItem): Promise<Permission[]> {
  const { items } = await ctx.graph.list<Permission>(`${itemUrl(target)}/permissions`, {
    pageSize: false,
    limit: Number.POSITIVE_INFINITY,
  });
  return items;
}

/** Tenant-wide search uses app-only `region`; only multi-geo tenants publish it, so fall back to asking. */
async function searchRegion(ctx: Context, region: string | undefined): Promise<string | undefined> {
  if (region || ctx.profile.authMode !== 'app') return region;
  const root = await ctx.graph.get<{ siteCollection?: { dataLocationCode?: string } }>('/sites/root', {
    query: { $select: 'siteCollection' },
  });
  const code = root.siteCollection?.dataLocationCode;
  if (!code) {
    throw new CliError('USAGE', 'App-only tenant-wide search needs the tenant region', {
      hint: 'Pass --region with your SharePoint data location, e.g. NAM, EUR, APC, GBR, AUS, CAN, JPN, IND.',
    });
  }
  return code;
}

interface UnshareArgs {
  ref: string;
  permissionId?: string;
  allLinks: boolean;
  site?: string;
}

async function planUnshare(ctx: Context, args: UnshareArgs) {
  if (Boolean(args.permissionId) === args.allLinks) {
    throw new CliError('USAGE', 'Pass either a permission id or --all-links', {
      hint: 'Find permission ids with `m365 sharepoint files permissions <ref>`.',
    });
  }
  const target = await resolveDriveItem(ctx, args.ref, { site: args.site });
  const permissions = await listPermissions(ctx, target);
  if (args.permissionId) {
    const permission = permissions.find((p) => p.id === args.permissionId);
    if (!permission) {
      throw new CliError('NOT_FOUND', `No permission ${args.permissionId} on "${target.item.name}"`, {
        hint: 'List current permissions with `m365 sharepoint files permissions <ref>`.',
      });
    }
    if (permission.inheritedFrom) {
      throw new CliError(
        'USAGE',
        `Permission ${permission.id} is inherited from ${permission.inheritedFrom.path ?? 'a parent'}`,
        {
          hint:
            'Inherited access can only be removed where it is defined: run `files permissions` on the parent ' +
            'folder/library and unshare there.',
        },
      );
    }
    return { target, remove: [permission], skippedInherited: [] as Permission[] };
  }
  const links = permissions.filter((p) => p.link);
  return {
    target,
    remove: links.filter((p) => !p.inheritedFrom),
    skippedInherited: links.filter((p) => p.inheritedFrom),
  };
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'sharepoint',
    name: 'files search',
    summary: 'Find files by name or content in a site, or across SharePoint/OneDrive',
    description:
      "With --site, searches that site's default library. Without it, uses Microsoft Search across all sites " +
      'and OneDrives the caller can see (app-only: shared content in one geo region, --region or derived from ' +
      'the tenant). Use the returned driveId:id as <ref> for `files get`, `files permissions`, `files versions`.',
    examples: ['m365 sharepoint files search "budget 2026"', 'm365 sharepoint files search invoice --site finance'],
    args: [
      { name: 'query', type: 'string', required: true, positional: true, description: 'Search text.' },
      SITE_OPT,
      { name: 'region', type: 'string', description: 'App-only tenant-wide search region (NAM, EUR, ...).' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['Sites.Read.All'], application: ['Sites.Read.All'] },
    async run(ctx, args) {
      const limit = listLimit(args);
      if (args.site) {
        const site = await resolveSite(ctx, args.site, 'id');
        const result = await ctx.graph.list<DriveItem>(
          `/sites/${site.id}/drive/root/search(q=${encodeURIComponent(odata(args.query))})`,
          { limit },
        );
        return { items: ctx.full ? result.items : result.items.map(shapeItem), truncated: result.truncated };
      }
      const region = await searchRegion(ctx, args.region);
      const items: DriveItem[] = [];
      let more = true;
      while (more && items.length < limit) {
        const size = Math.min(limit - items.length, 500);
        const response = await ctx.graph.post<SearchResponse>('/search/query', {
          requests: [
            { entityTypes: ['driveItem'], query: { queryString: args.query }, from: items.length, size, region },
          ],
        });
        const container = response.value?.[0]?.hitsContainers?.[0];
        const hits = (container?.hits ?? []).flatMap((h) => (h.resource ? [h.resource] : []));
        items.push(...hits);
        more = Boolean(container?.moreResultsAvailable) && hits.length > 0;
      }
      return { items: ctx.full ? items : items.map(shapeItem), truncated: more };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'files get',
    summary: 'Show a file or folder: size, type, who created/changed it, and whether it is shared',
    description:
      `sharedScope (anonymous/organization/users) appears when the item is shared. ${REF_NOTE} ` +
      'Next: `files permissions` for exactly who has access.',
    examples: [
      'm365 sharepoint files get "https://contoso.sharepoint.com/:x:/s/finance/EbC..."',
      'm365 sharepoint files get "Reports/Q1.xlsx" --site finance',
    ],
    args: [REF_ARG, SITE_OPT],
    risk: 'read',
    permissions: { delegated: ['Sites.Read.All'], application: ['Sites.Read.All'] },
    async run(ctx, args) {
      const { item } = await resolveDriveItem(ctx, args.ref, { site: args.site });
      if (ctx.full) return item;
      return {
        ...shapeItem(item),
        type: item.folder ? 'folder' : item.package ? (item.package.type ?? 'package') : 'file',
        mimeType: item.file?.mimeType ?? null,
        childCount: item.folder?.childCount,
        sizeMb: typeof item.size === 'number' ? Math.round((item.size / 1024 ** 2) * 100) / 100 : null,
        createdBy: who(item.createdBy),
        createdDateTime: item.createdDateTime ?? null,
        siteId: item.parentReference?.siteId ?? null,
        shared: Boolean(item.shared),
        sharedScope: item.shared?.scope ?? null,
        sharedBy: who(item.shared?.sharedBy),
        sharedDateTime: item.shared?.sharedDateTime ?? null,
      };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'files permissions',
    summary: 'Who can access a file or folder: sharing links, direct grants, inherited access, externals',
    description:
      'The answer to "who can see this file". kind: link (linkScope anonymous = anyone with the URL, no sign-in; ' +
      'organization = anyone in the tenant with the URL; users = only named people; existingAccess = no new ' +
      'access), direct (granted to a person/group), inherited (from a parent folder or the library; ' +
      'inheritedFrom shows where). grantedTo lists people/groups with email and external (outside the ' +
      "tenant's verified domains or a guest account). anonymousLinks and externalUsers summarize the risk. " +
      `Permission ids feed \`files unshare\`. ${REF_NOTE}`,
    examples: ['m365 sharepoint files permissions "https://contoso.sharepoint.com/:w:/s/hr/EfX..."'],
    args: [REF_ARG, SITE_OPT],
    risk: 'read',
    permissions: {
      delegated: ['Sites.Read.All', 'Domain.Read.All'],
      application: ['Sites.Read.All', 'Domain.Read.All'],
    },
    async run(ctx, args) {
      const target = await resolveDriveItem(ctx, args.ref, { site: args.site });
      const [raw, parts] = await Promise.all([
        listPermissions(ctx, target),
        settle({
          domains: ctx.graph.get<{ value?: { id: string; isVerified?: boolean }[] }>('/domains', {
            query: { $select: 'id,isVerified' },
          }),
        }),
      ]);
      if (ctx.full) return { item: target.item, permissions: raw };
      const tenantDomains =
        'error' in parts.domains
          ? undefined
          : (parts.domains.value ?? []).filter((d) => d.isVerified).map((d) => d.id.toLowerCase());
      const permissions = raw.map((p) => shapePermission(p, tenantDomains));
      const externalUsers = [
        ...new Set(
          permissions.flatMap((p) => p.grantedTo.filter((g) => g.external).map((g) => g.email ?? g.displayName ?? '')),
        ),
      ];
      return {
        item: {
          name: target.item.name,
          path: itemPath(target.item),
          webUrl: target.item.webUrl ?? null,
          ref: `${target.driveId}:${target.itemId}`,
        },
        summary: {
          total: permissions.length,
          links: permissions.filter((p) => p.linkScope).length,
          direct: permissions.filter((p) => p.kind === 'direct').length,
          inherited: permissions.filter((p) => p.kind === 'inherited').length,
        },
        anonymousLinks: permissions
          .filter((p) => p.linkScope === 'anonymous')
          .map((p) => ({ id: p.id, linkType: p.linkType, expires: p.expires, inherited: p.kind === 'inherited' })),
        externalUsers,
        ...('error' in parts.domains
          ? { externalCheck: `guest accounts only; could not read verified domains: ${parts.domains.error.message}` }
          : {}),
        permissions,
      };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'files unshare',
    summary: 'Remove a sharing link or direct grant from a file, or every sharing link with --all-links',
    description:
      'Removes access immediately: anyone relying on the link/grant loses it. Pass a permission id from ' +
      '`files permissions`, or --all-links to delete every non-inherited sharing link. Inherited permissions ' +
      'cannot be removed on the item; unshare them on the parent named in inheritedFrom. The preview lists ' +
      `exactly what will be removed. ${REF_NOTE}`,
    examples: [
      'm365 sharepoint files unshare "https://contoso.sharepoint.com/:x:/s/finance/EbC..." aTowIy5mfG1lbWJl',
      'm365 sharepoint files unshare "Reports/Q1.xlsx" --site finance --all-links',
    ],
    args: [
      REF_ARG,
      { name: 'permissionId', type: 'string', positional: true, description: 'Permission id to remove.' },
      { name: 'allLinks', type: 'boolean', description: 'Remove every sharing link on the item.' },
      SITE_OPT,
    ],
    risk: 'destructive',
    permissions: { delegated: ['Sites.ReadWrite.All'], application: ['Sites.ReadWrite.All'] },
    async preview(ctx, args) {
      const plan = await planUnshare(ctx, args);
      return {
        item: {
          name: plan.target.item.name,
          path: itemPath(plan.target.item),
          webUrl: plan.target.item.webUrl ?? null,
        },
        remove: plan.remove.map((p) => shapePermission(p)),
        skippedInherited: plan.skippedInherited.map((p) => ({
          id: p.id,
          inheritedFrom: p.inheritedFrom?.path ?? null,
        })),
        action: plan.remove.length
          ? `Delete ${plan.remove.length} permission(s) from "${plan.target.item.name}"; people using them lose access.`
          : `Nothing to remove: "${plan.target.item.name}" has no removable sharing links.`,
      };
    },
    async run(ctx, args) {
      const plan = await planUnshare(ctx, args);
      for (const permission of plan.remove) {
        await ctx.graph.delete(`${itemUrl(plan.target)}/permissions/${seg(permission.id)}`);
      }
      return {
        item: plan.target.item.name,
        removed: plan.remove.map((p) => p.id),
        skippedInherited: plan.skippedInherited.map((p) => p.id),
        next: 'Verify with `m365 sharepoint files permissions <ref>`.',
      };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'files versions',
    summary: 'List the version history of a file, newest first',
    description: `Version ids feed \`files restore-version\`. The first entry (current: true) is the live version. ${REF_NOTE}`,
    examples: ['m365 sharepoint files versions "Reports/Q1.xlsx" --site finance'],
    args: [REF_ARG, SITE_OPT, ...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: ['Sites.Read.All'], application: ['Sites.Read.All'] },
    async run(ctx, args) {
      const target = await resolveDriveItem(ctx, args.ref, { site: args.site });
      if (target.item.folder)
        throw new CliError('USAGE', `"${target.item.name}" is a folder; only files have versions`);
      const result = await ctx.graph.list<DriveItemVersion>(`${itemUrl(target)}/versions`, {
        pageSize: false,
        limit: listLimit(args),
      });
      if (ctx.full) return result;
      const sorted = [...result.items].sort((a, b) =>
        (b.lastModifiedDateTime ?? '').localeCompare(a.lastModifiedDateTime ?? ''),
      );
      return {
        items: sorted.map((v, index) => ({
          id: v.id,
          current: index === 0 && !result.truncated,
          lastModifiedDateTime: v.lastModifiedDateTime ?? null,
          lastModifiedBy: who(v.lastModifiedBy),
          size: v.size ?? null,
        })),
        truncated: result.truncated,
      };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'files restore-version',
    summary: 'Make an earlier version of a file the current version',
    description:
      'Restoring adds a new current version with the old content; history is kept, so it can be undone by ' +
      `restoring again. Get version ids from \`files versions\`. ${REF_NOTE}`,
    examples: ['m365 sharepoint files restore-version "Reports/Q1.xlsx" 3.0 --site finance'],
    args: [
      REF_ARG,
      { name: 'versionId', type: 'string', required: true, positional: true, description: 'Version id, e.g. 3.0.' },
      SITE_OPT,
    ],
    risk: 'write',
    permissions: { delegated: ['Sites.ReadWrite.All'], application: ['Sites.ReadWrite.All'] },
    async run(ctx, args) {
      const target = await resolveDriveItem(ctx, args.ref, { site: args.site });
      await ctx.graph.post(`${itemUrl(target)}/versions/${seg(args.versionId)}/restoreVersion`);
      return {
        item: target.item.name,
        restoredVersion: args.versionId,
        next: 'Check with `m365 sharepoint files versions <ref>`.',
      };
    },
  }),
];
