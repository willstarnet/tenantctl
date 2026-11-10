import { CliError } from '../../core/errors.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';
import { settle } from '../../core/util.ts';
import {
  type Drive,
  type Identity,
  type IdentitySet,
  resolveSite,
  SITE_SELECT,
  type Site,
  shapeQuota,
  shapeSite,
  who,
} from './resolve.ts';

const SITE_ARG = {
  name: 'site',
  type: 'string',
  required: true,
  positional: true,
  description: 'Site URL (any page or file URL under it works), Graph site id, or site name.',
} as const;

const DRIVE_SELECT = 'id,name,description,driveType,webUrl,createdDateTime,lastModifiedDateTime,owner,quota';

interface SitePermission {
  id: string;
  roles?: string[];
  grantedToIdentitiesV2?: IdentitySet[];
  grantedToIdentities?: IdentitySet[];
}

interface DirectoryObject {
  id: string;
  displayName?: string;
  userPrincipalName?: string;
  mail?: string;
}

function shapeDrive(drive: Drive) {
  return {
    id: drive.id,
    name: drive.name ?? null,
    driveType: drive.driveType ?? null,
    webUrl: drive.webUrl ?? null,
    owner: who(drive.owner),
    quota: shapeQuota(drive.quota),
    createdDateTime: drive.createdDateTime ?? null,
    lastModifiedDateTime: drive.lastModifiedDateTime ?? null,
  };
}

/**
 * Graph v1.0 has no groupId on site; a group-connected site's default library is owned by its Microsoft 365
 * group, so the library owner tells us.
 */
function connectedGroup(drive: Drive | null): Identity | null {
  return drive?.owner?.group?.id ? drive.owner.group : null;
}

async function memberCount(ctx: Context, groupId: string, relation: 'owners' | 'members'): Promise<number | null> {
  const result = await ctx.graph.get<{ '@odata.count'?: number }>(`/groups/${groupId}/${relation}`, {
    query: { $count: true, $top: 1, $select: 'id' },
    advanced: true,
  });
  return result['@odata.count'] ?? null;
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'sharepoint',
    name: 'sites list',
    summary: 'Find SharePoint sites by keyword, or list every site (app-only)',
    description:
      'With --search, matches site titles/URLs (/sites?search=). Without it, app-only profiles list every site ' +
      'in the tenant (/sites/getAllSites, application permission only); delegated sign-ins must pass --search. ' +
      'OneDrive personal sites show isPersonalSite: true. Next: `sharepoint sites get <webUrl>`.',
    examples: ['m365 sharepoint sites list --search finance', 'm365 sharepoint sites list --all'],
    args: [
      { name: 'search', type: 'string', description: 'Keyword matched against site name/title/URL.' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { delegated: ['Sites.Read.All'], application: ['Sites.Read.All'] },
    async run(ctx, args) {
      let path = '/sites';
      const query: Record<string, string> = { $select: SITE_SELECT };
      if (args.search) {
        query.search = args.search;
      } else if (ctx.profile.authMode === 'app') {
        path = '/sites/getAllSites';
      } else {
        throw new CliError('USAGE', 'Listing every site needs app-only auth', {
          hint: 'Pass --search <keyword>, or use an app-only profile (Sites.Read.All application permission).',
        });
      }
      const result = await ctx.graph.list<Site>(path, { query, limit: listLimit(args) });
      return { items: ctx.full ? result.items : result.items.map(shapeSite), truncated: result.truncated };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'sites get',
    summary: 'Show a site: storage quota, library count, and its Microsoft 365 group and owners',
    description:
      'Answers "how full is this site", "who owns it", "is it a Teams/M365 group site". storage is the default ' +
      'document library quota (state normal/nearing <10% left/critical <1% left/exceeded). groupConnected is ' +
      "derived from the default library being owned by an M365 group; owners are that group's owners. Classic " +
      'site collection admins are not exposed by Graph. Next: `sharepoint sites permissions`, `sites drives`.',
    examples: ['m365 sharepoint sites get https://contoso.sharepoint.com/sites/finance'],
    args: [SITE_ARG],
    risk: 'read',
    permissions: {
      delegated: ['Sites.Read.All', 'GroupMember.Read.All'],
      application: ['Sites.Read.All', 'GroupMember.Read.All'],
    },
    async run(ctx, args) {
      const site = await resolveSite(ctx, args.site);
      const parts = await settle({
        drive: ctx.graph.get<Drive>(`/sites/${site.id}/drive`, { query: { $select: DRIVE_SELECT } }),
        drives: ctx.graph.list<Drive>(`/sites/${site.id}/drives`, {
          query: { $select: 'id' },
          limit: Number.POSITIVE_INFINITY,
        }),
      });
      const drive = 'error' in parts.drive ? null : parts.drive;
      const group = connectedGroup(drive);
      const owners = group?.id
        ? (
            await settle({
              owners: ctx.graph.list<DirectoryObject>(`/groups/${group.id}/owners`, {
                query: { $select: 'id,displayName,userPrincipalName,mail' },
                limit: 100,
              }),
            })
          ).owners
        : null;
      if (ctx.full) return { site, ...parts, owners };
      return {
        ...shapeSite(site),
        description: site.description ?? null,
        storage: 'error' in parts.drive ? parts.drive : { library: parts.drive.name, ...shapeQuota(parts.drive.quota) },
        libraryCount: 'error' in parts.drives ? parts.drives : parts.drives.items.length,
        groupConnected: drive ? Boolean(group) : null,
        group: group
          ? {
              id: group.id,
              displayName: group.displayName ?? null,
              email: group.email ?? null,
              owners:
                owners && 'error' in owners
                  ? owners
                  : (owners?.items ?? []).map((o) => ({
                      displayName: o.displayName ?? null,
                      userPrincipalName: o.userPrincipalName ?? o.mail ?? null,
                    })),
            }
          : null,
      };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'sites drives',
    summary: 'List the document libraries of a site with their storage use',
    description:
      'Each library (drive) with owner and quota. Libraries of one site share the site storage quota, so ' +
      'percentUsed is the same site-wide figure. Use a library id with `files` commands as driveId:itemId.',
    examples: ['m365 sharepoint sites drives https://contoso.sharepoint.com/sites/finance'],
    args: [SITE_ARG, ...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: ['Sites.Read.All'], application: ['Sites.Read.All'] },
    async run(ctx, args) {
      const site = await resolveSite(ctx, args.site, 'id');
      const result = await ctx.graph.list<Drive>(`/sites/${site.id}/drives`, {
        query: { $select: DRIVE_SELECT },
        limit: listLimit(args),
      });
      return { items: ctx.full ? result.items : result.items.map(shapeDrive), truncated: result.truncated };
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'sites permissions',
    summary: 'Show which apps have Sites.Selected access to a site, and its M365 group owner/member counts',
    description:
      'appPermissions lists applications granted access to this specific site (Sites.Selected grants), with ' +
      'roles such as read or write. For group-connected sites, site access follows the Microsoft 365 ' +
      'group: owners are site owners, members are site members (counts shown; list them with `entra groups ' +
      'members`). Classic SharePoint groups (Owners/Members/Visitors), site collection admins and per-user site ' +
      'grants are not exposed by Graph v1.0; use SharePoint Online Management Shell or PnP PowerShell. To see ' +
      'who can open a particular file use `sharepoint files permissions`. Reading app grants needs ' +
      'Sites.FullControl.All.',
    examples: ['m365 sharepoint sites permissions https://contoso.sharepoint.com/sites/finance'],
    args: [SITE_ARG],
    risk: 'read',
    permissions: {
      delegated: ['Sites.FullControl.All', 'GroupMember.Read.All'],
      application: ['Sites.FullControl.All', 'GroupMember.Read.All'],
    },
    async run(ctx, args) {
      const site = await resolveSite(ctx, args.site, 'id,displayName,webUrl');
      const parts = await settle({
        appPermissions: ctx.graph.list<SitePermission>(`/sites/${site.id}/permissions`, {
          pageSize: false,
          limit: Number.POSITIVE_INFINITY,
        }),
        drive: ctx.graph.get<Drive>(`/sites/${site.id}/drive`, { query: { $select: 'id,owner' } }),
      });
      const group = 'error' in parts.drive ? null : connectedGroup(parts.drive);
      const counts = group?.id
        ? await settle({ owners: memberCount(ctx, group.id, 'owners'), members: memberCount(ctx, group.id, 'members') })
        : null;
      if (ctx.full) return { site, ...parts, groupCounts: counts };
      return {
        site: site.displayName ?? null,
        webUrl: site.webUrl,
        appPermissions:
          'error' in parts.appPermissions
            ? parts.appPermissions
            : parts.appPermissions.items.map((p) => ({
                id: p.id,
                roles: p.roles ?? [],
                apps: (p.grantedToIdentitiesV2 ?? p.grantedToIdentities ?? []).map((set) => ({
                  displayName: set.application?.displayName ?? null,
                  appId: set.application?.id ?? null,
                })),
              })),
        groupConnected: 'error' in parts.drive ? parts.drive : Boolean(group),
        group: group
          ? { id: group.id, displayName: group.displayName ?? null, owners: counts?.owners, members: counts?.members }
          : null,
      };
    },
  }),
];
