import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { pickOne } from '../../core/resolve.ts';
import type { Context } from '../../core/types.ts';
import { bytesToGb } from '../../core/util.ts';

/**
 * Sites and files are named by URL in conversation ("the Finance site", a link pasted from Teams), so every
 * SharePoint/OneDrive command resolves targets here.
 */

export interface Identity {
  id?: string;
  displayName?: string;
  email?: string;
  loginName?: string;
}

/** identitySet / sharePointIdentitySet. */
export interface IdentitySet {
  user?: Identity;
  group?: Identity;
  application?: Identity;
  device?: Identity;
  siteUser?: Identity;
  siteGroup?: Identity;
}

export interface Quota {
  used?: number;
  total?: number;
  remaining?: number;
  deleted?: number;
  state?: string;
}

export interface Site {
  id: string;
  name?: string;
  displayName?: string;
  description?: string;
  webUrl: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  isPersonalSite?: boolean;
}

export interface Drive {
  id: string;
  name?: string;
  driveType?: string;
  webUrl?: string;
  description?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  owner?: IdentitySet;
  quota?: Quota;
}

export interface DriveItem {
  id: string;
  name: string;
  webUrl?: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: IdentitySet;
  lastModifiedBy?: IdentitySet;
  parentReference?: { driveId?: string; driveType?: string; id?: string; path?: string; siteId?: string };
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  package?: { type?: string };
  shared?: { scope?: string; owner?: IdentitySet; sharedBy?: IdentitySet; sharedDateTime?: string };
}

export interface ResolvedDriveItem {
  driveId: string;
  itemId: string;
  item: DriveItem;
}

export const SITE_SELECT = 'id,name,displayName,description,webUrl,createdDateTime,lastModifiedDateTime,isPersonalSite';

// Graph site ids are `{hostname},{site-collection-guid},{web-guid}`.
const SITE_ID = /^[a-z0-9.-]+,[0-9a-f-]{36},[0-9a-f-]{36}$/i;
const DRIVE_ITEM_ID = /^([^\s:/]+):([^\s:/]+)$/;

/** Display name of whoever is in an identity set (user first, then group/app/SharePoint principals). */
export function who(set: IdentitySet | undefined | null): string | null {
  if (!set) return null;
  const identity = set.user ?? set.group ?? set.siteGroup ?? set.siteUser ?? set.application ?? set.device;
  return identity?.displayName ?? identity?.email ?? null;
}

export function shapeQuota(quota: Quota | undefined | null) {
  if (!quota) return null;
  const percentUsed =
    typeof quota.used === 'number' && quota.total ? Math.round((quota.used / quota.total) * 1000) / 10 : null;
  return {
    usedGb: bytesToGb(quota.used),
    totalGb: bytesToGb(quota.total),
    remainingGb: bytesToGb(quota.remaining),
    recycleBinGb: bytesToGb(quota.deleted),
    percentUsed,
    state: quota.state ?? null,
  };
}

export function shapeSite(site: Site) {
  return {
    id: site.id,
    name: site.name ?? null,
    displayName: site.displayName ?? null,
    webUrl: site.webUrl,
    createdDateTime: site.createdDateTime ?? null,
    lastModifiedDateTime: site.lastModifiedDateTime ?? null,
    isPersonalSite: site.isPersonalSite ?? false,
  };
}

/** Folder path of an item inside its library: parentReference.path is `/drives/{id}/root:/Folder`. */
export function itemPath(item: DriveItem): string | null {
  const parent = item.parentReference?.path;
  if (parent === undefined) return null;
  const folder = parent.includes('root:') ? parent.slice(parent.indexOf('root:') + 5) : parent;
  return `${folder.replace(/\/$/, '')}/${item.name}`;
}

/**
 * A site URL (any page or file URL under the site works), a Graph site id, or a name to search for.
 * URLs resolve to the site collection under /sites/, /teams/ or /personal/; subsites need their site id.
 */
export async function resolveSite(ctx: Context, ref: string, select = SITE_SELECT): Promise<Site> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A site is required');
  const query = { $select: select };
  if (/^https?:\/\//i.test(value)) {
    let url: URL;
    try {
      url = new URL(value);
    } catch {
      throw new CliError('USAGE', `"${value}" is not a valid URL`);
    }
    const segments = decodeURIComponent(url.pathname).split('/').filter(Boolean);
    const managed = segments[0]?.toLowerCase();
    const sitePath =
      segments.length >= 2 && (managed === 'sites' || managed === 'teams' || managed === 'personal')
        ? segments.slice(0, 2)
        : [];
    const path = sitePath.length
      ? `/sites/${seg(url.hostname)}:/${sitePath.map(seg).join('/')}`
      : `/sites/${seg(url.hostname)}`;
    const site = await ctx.graph.tryGet<Site>(path, { query });
    if (!site) {
      throw new CliError('NOT_FOUND', `No site at ${url.origin}/${sitePath.join('/')}`, {
        hint: 'Check the URL, or find the site with `m365 sharepoint sites list --search <name>`.',
      });
    }
    return site;
  }
  if (SITE_ID.test(value)) {
    const site = await ctx.graph.tryGet<Site>(`/sites/${value}`, { query });
    if (!site) throw new CliError('NOT_FOUND', `No site with id ${value}`);
    return site;
  }
  const { items } = await ctx.graph.list<Site>('/sites', { query: { search: value, $select: select }, limit: 25 });
  return pickOne(
    items.filter((s) => !s.isPersonalSite),
    value,
    'site',
    (s) => [s.displayName, s.name, s.webUrl],
    (s) => ({ id: s.id, displayName: s.displayName, webUrl: s.webUrl }),
  );
}

/** The `/shares` id for a URL: `u!` + unpadded base64url of the UTF-8 URL. */
export function encodeSharingUrl(url: string): string {
  return `u!${Buffer.from(url, 'utf8').toString('base64url')}`;
}

/**
 * A sharing link or any file/folder URL (resolved through /shares), `driveId:itemId`, or with `site` a path
 * inside that site's default document library ("Shared Documents" is the library root, so pass `Reports/q1.xlsx`).
 */
export async function resolveDriveItem(
  ctx: Context,
  ref: string,
  options: { site?: string } = {},
): Promise<ResolvedDriveItem> {
  const value = ref.trim();
  if (!value) throw new CliError('USAGE', 'A file or folder is required');
  let item: DriveItem | null;
  if (/^https?:\/\//i.test(value)) {
    try {
      item = await ctx.graph.tryGet<DriveItem>(`/shares/${encodeSharingUrl(value)}/driveItem`);
    } catch (error) {
      if (error instanceof CliError && error.code === 'FORBIDDEN') {
        throw new CliError('FORBIDDEN', error.message, {
          status: error.status,
          requestId: error.requestId,
          hint:
            'Resolving a URL uses the /shares API, which needs Files.ReadWrite.All (or Sites.ReadWrite.All). ' +
            'Without it, pass --site <site> with the path inside the library, or driveId:itemId.',
        });
      }
      throw error;
    }
    if (!item) {
      throw new CliError('NOT_FOUND', 'The link does not resolve to a file or folder', {
        hint:
          'The link may be expired or removed, or (delegated sign-in) the signed-in user cannot open it. ' +
          'Try --site <site> with the file path, or app-only auth.',
      });
    }
  } else if (options.site) {
    const site = await resolveSite(ctx, options.site, 'id,webUrl');
    const path = value.split('/').filter(Boolean).map(seg).join('/');
    item = path
      ? await ctx.graph.tryGet<DriveItem>(`/sites/${site.id}/drive/root:/${path}`)
      : await ctx.graph.get<DriveItem>(`/sites/${site.id}/drive/root`);
    if (!item) {
      throw new CliError('NOT_FOUND', `No item "${value}" in the default library of ${site.webUrl}`, {
        hint: 'Paths are relative to the library root, e.g. "Reports/Q1.xlsx". Find it with `m365 sharepoint files search`.',
      });
    }
  } else {
    const match = DRIVE_ITEM_ID.exec(value);
    if (!match) {
      throw new CliError('USAGE', `Cannot tell which file "${value}" is`, {
        hint: 'Pass a file/sharing URL, driveId:itemId (from `files search`), or a library path with --site <site>.',
      });
    }
    const [, driveId = '', itemId = ''] = match;
    item = await ctx.graph.tryGet<DriveItem>(`/drives/${seg(driveId)}/items/${seg(itemId)}`);
    if (!item) throw new CliError('NOT_FOUND', `No item ${itemId} in drive ${driveId}`);
  }
  const driveId = item.parentReference?.driveId;
  if (!driveId) throw new CliError('GRAPH_ERROR', `Graph returned item ${item.id} without its drive id`);
  return { driveId, itemId: item.id, item };
}
