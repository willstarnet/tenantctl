import { describe, expect, test } from 'bun:test';
import { CliError } from '../../src/core/errors.ts';
import { commands, resolveDriveItem, resolveSite, shapePermission } from '../../src/domains/sharepoint/index.ts';
import { commandById, type MockRoute, mockContext, page, runCommand } from '../helpers/graphMock.ts';

const SITE_ID = 'contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222';
const SITE = {
  id: SITE_ID,
  displayName: 'Finance',
  name: 'finance',
  webUrl: 'https://contoso.sharepoint.com/sites/finance',
};
const ITEM = {
  id: 'item1',
  name: 'Q1.xlsx',
  webUrl: 'https://contoso.sharepoint.com/sites/finance/Shared%20Documents/Reports/Q1.xlsx',
  size: 2048,
  parentReference: { driveId: 'b!drive', path: '/drives/b!drive/root:/Reports' },
};

/** The /shares encoding exactly as the Graph docs describe it: base64, strip '=', '/'->'_', '+'->'-', prefix u!. */
function documentedShareId(url: string): string {
  const base64 = btoa(String.fromCharCode(...new TextEncoder().encode(url)));
  return `u!${base64.replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-')}`;
}

describe('resolveSite', () => {
  test('a page or file URL resolves to its site collection by server-relative path', async () => {
    const { ctx, calls } = mockContext([
      { path: '/sites/contoso.sharepoint.com:/sites/finance', reply: { body: SITE } },
    ]);
    const site = await resolveSite(
      ctx,
      'https://contoso.sharepoint.com/sites/finance/Shared%20Documents/Forms/AllItems.aspx',
    );
    expect(site.id).toBe(SITE_ID);
    expect(calls).toHaveLength(1);
  });

  test('the tenant root URL resolves by hostname alone', async () => {
    const { ctx, calls } = mockContext([{ path: '/sites/contoso.sharepoint.com', reply: { body: { id: 'root' } } }]);
    expect((await resolveSite(ctx, 'https://contoso.sharepoint.com/')).id).toBe('root');
    expect(calls[0]?.path).toBe('/sites/contoso.sharepoint.com');
  });

  test('a Graph site id is fetched directly', async () => {
    const { ctx, calls } = mockContext([{ path: `/sites/${SITE_ID}`, reply: { body: SITE } }]);
    await resolveSite(ctx, SITE_ID);
    expect(calls[0]?.path).toBe(`/sites/${SITE_ID}`);
  });

  test('names search /sites; an exact title wins, otherwise several matches are AMBIGUOUS', async () => {
    const sites = [SITE, { ...SITE, id: 'x', displayName: 'Finance Archive', name: 'finance-archive' }];
    const { ctx, calls } = mockContext([{ path: '/sites', reply: page(sites) }]);
    expect((await resolveSite(ctx, 'finance')).id).toBe(SITE_ID);
    expect(calls[0]?.query.search).toBe('finance');
    const error = (await resolveSite(ctx, 'fin').catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
    expect(error.details).toHaveLength(2);
  });
});

describe('resolveDriveItem', () => {
  test('a sharing link is looked up via /shares with the documented u! base64url encoding', async () => {
    const link = 'https://contoso.sharepoint.com/:x:/s/finance/EbCdeF?e=Ab1+/x>>';
    const shareId = documentedShareId(link);
    expect(shareId).not.toMatch(/[=+/]/);
    const { ctx, calls } = mockContext([{ path: `/shares/${shareId}/driveItem`, reply: { body: ITEM } }]);
    const resolved = await resolveDriveItem(ctx, link);
    expect(resolved).toMatchObject({ driveId: 'b!drive', itemId: 'item1' });
    expect(calls[0]?.headers.Prefer).toBeUndefined();
  });

  test('a path with --site reads from the site default library', async () => {
    const { ctx, calls } = mockContext([
      { path: '/sites', reply: page([SITE]) },
      { path: `/sites/${SITE_ID}/drive/root:/Reports/Q1 final.xlsx`, reply: { body: ITEM } },
    ]);
    await resolveDriveItem(ctx, '/Reports/Q1 final.xlsx', { site: 'Finance' });
    expect(calls.at(-1)?.path).toBe(`/sites/${SITE_ID}/drive/root:/Reports/Q1 final.xlsx`);
  });

  test('driveId:itemId goes straight to the drive; anything else is a usage error', async () => {
    const { ctx, calls } = mockContext([{ path: '/drives/b!drive/items/item1', reply: { body: ITEM } }]);
    expect((await resolveDriveItem(ctx, 'b!drive:item1')).itemId).toBe('item1');
    expect(calls).toHaveLength(1);
    const error = (await resolveDriveItem(ctx, 'Reports/Q1.xlsx').catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('USAGE');
  });
});

describe('shapePermission', () => {
  const domains = ['contoso.com', 'contoso.onmicrosoft.com'];

  test('anonymous link: kind, scope/type, MinValue expiry means none', () => {
    const shaped = shapePermission(
      {
        id: 'p1',
        roles: ['write'],
        link: { scope: 'anonymous', type: 'edit', webUrl: 'https://contoso.sharepoint.com/:x:/s/f/abc' },
        expirationDateTime: '0001-01-01T00:00:00Z',
        hasPassword: false,
      },
      domains,
    );
    expect(shaped).toMatchObject({
      kind: 'link',
      linkScope: 'anonymous',
      linkType: 'edit',
      expires: null,
      roles: ['write'],
    });
  });

  test('specific-people link flags recipients outside verified domains and guest accounts as external', () => {
    const shaped = shapePermission(
      {
        id: 'p2',
        roles: ['read'],
        link: { scope: 'users', type: 'view' },
        expirationDateTime: '2026-12-01T00:00:00Z',
        grantedToIdentitiesV2: [
          { user: { displayName: 'Jane', email: 'jane@contoso.com' }, siteUser: { displayName: 'Jane' } },
          { siteUser: { displayName: 'Vendor', loginName: 'i:0#.f|membership|bob@fabrikam.com' } },
          { user: { displayName: 'Guest', email: 'amy_gmail.com#ext#@contoso.onmicrosoft.com' } },
        ],
      },
      domains,
    );
    expect(shaped.expires).toBe('2026-12-01T00:00:00Z');
    expect(shaped.grantedTo).toEqual([
      { type: 'user', displayName: 'Jane', email: 'jane@contoso.com', external: false },
      { type: 'user', displayName: 'Vendor', email: 'bob@fabrikam.com', external: true },
      { type: 'user', displayName: 'Guest', email: 'amy_gmail.com#ext#@contoso.onmicrosoft.com', external: true },
    ]);
  });

  test('inherited SharePoint group grant', () => {
    const shaped = shapePermission({
      id: 'p3',
      roles: ['owner'],
      grantedToV2: { siteGroup: { displayName: 'Finance Owners' } },
      inheritedFrom: { path: '/drives/b!drive/root:' },
    });
    expect(shaped).toMatchObject({
      kind: 'inherited',
      inheritedFrom: '/drives/b!drive/root:',
      grantedTo: [{ type: 'sharePointGroup', displayName: 'Finance Owners', email: null }],
    });
  });
});

const PERMISSIONS = [
  { id: 'anon', roles: ['read'], link: { scope: 'anonymous', type: 'view' } },
  { id: 'org', roles: ['write'], link: { scope: 'organization', type: 'edit' } },
  {
    id: 'ext',
    roles: ['read'],
    grantedToV2: { user: { displayName: 'Bob', email: 'bob@fabrikam.com' } },
  },
  {
    id: 'inh',
    roles: ['read'],
    link: { scope: 'organization', type: 'view' },
    inheritedFrom: { path: '/drives/b!drive/root:/Reports' },
  },
];

function itemRoutes(onDelete?: MockRoute['reply']): MockRoute[] {
  return [
    { path: '/drives/b!drive/items/item1', reply: { body: ITEM } },
    { path: '/drives/b!drive/items/item1/permissions', reply: page(PERMISSIONS) },
    {
      path: '/domains',
      reply: page([
        { id: 'contoso.com', isVerified: true },
        { id: 'fabrikam.com', isVerified: false },
      ]),
    },
    { method: 'DELETE', path: /^\/drives\/b!drive\/items\/item1\/permissions\//, reply: onDelete ?? { status: 204 } },
  ];
}

describe('files permissions', () => {
  test('summarizes anonymous links and external users', async () => {
    const command = commandById(commands, 'sharepoint files permissions');
    const { outcome } = await runCommand(command, ['b!drive:item1'], itemRoutes());
    expect(outcome).toMatchObject({
      kind: 'result',
      result: {
        summary: { total: 4, links: 3, direct: 1, inherited: 1 },
        anonymousLinks: [{ id: 'anon', linkType: 'view', expires: null, inherited: false }],
        externalUsers: ['bob@fabrikam.com'],
        item: { path: '/Reports/Q1.xlsx', ref: 'b!drive:item1' },
      },
    });
  });
});

describe('files unshare', () => {
  const command = commandById(commands, 'sharepoint files unshare');

  test('previews before deleting, then the confirmed run deletes that permission', async () => {
    const first = await runCommand(command, ['b!drive:item1', 'ext'], itemRoutes(), { allowWrites: true });
    expect(first.outcome).toMatchObject({ kind: 'confirm', preview: { remove: [{ id: 'ext' }] } });
    expect(first.calls.some((c) => c.method === 'DELETE')).toBe(false);
    expect(JSON.stringify(first.outcome)).toContain('Delete 1 permission(s) from \\"Q1.xlsx\\"');

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(command, ['b!drive:item1', 'ext'], itemRoutes(), {
      allowWrites: true,
      confirm: token,
    });
    const deletes = second.calls.filter((c) => c.method === 'DELETE');
    expect(deletes.map((c) => c.path)).toEqual(['/drives/b!drive/items/item1/permissions/ext']);
  });

  test('refuses an inherited permission with a hint, without deleting', async () => {
    const error = (await runCommand(command, ['b!drive:item1', 'inh'], itemRoutes(), { allowWrites: true }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(error.code).toBe('USAGE');
    expect(error.message).toContain('/drives/b!drive/root:/Reports');
    expect(error.hint).toContain('parent');
  });

  test('--all-links removes every non-inherited link and reports skipped inherited ones', async () => {
    const first = await runCommand(command, ['b!drive:item1', '--all-links'], itemRoutes(), { allowWrites: true });
    expect(first.outcome).toMatchObject({
      kind: 'confirm',
      preview: { remove: [{ id: 'anon' }, { id: 'org' }], skippedInherited: [{ id: 'inh' }] },
    });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(command, ['b!drive:item1', '--all-links'], itemRoutes(), {
      allowWrites: true,
      confirm: token,
    });
    expect(second.calls.filter((c) => c.method === 'DELETE').map((c) => c.path.split('/').pop())).toEqual([
      'anon',
      'org',
    ]);
  });

  test('needs exactly one of a permission id or --all-links', async () => {
    const error = (await runCommand(command, ['b!drive:item1'], itemRoutes(), { dryRun: true }).catch(
      (e: unknown) => e,
    )) as CliError;
    expect(error.code).toBe('USAGE');
  });
});

describe('sites', () => {
  test('sites get: quota, library count, group connection and owners', async () => {
    const command = commandById(commands, 'sharepoint sites get');
    const { outcome } = await runCommand(
      command,
      ['https://contoso.sharepoint.com/sites/finance'],
      [
        { path: '/sites/contoso.sharepoint.com:/sites/finance', reply: { body: SITE } },
        {
          path: `/sites/${SITE_ID}/drive`,
          reply: {
            body: {
              id: 'b!drive',
              name: 'Documents',
              owner: { group: { id: 'g1', displayName: 'Finance', email: 'finance@contoso.com' } },
              quota: {
                used: 5 * 1024 ** 3,
                total: 10 * 1024 ** 3,
                remaining: 5 * 1024 ** 3,
                deleted: 0,
                state: 'normal',
              },
            },
          },
        },
        { path: `/sites/${SITE_ID}/drives`, reply: page([{ id: 'b!drive' }, { id: 'b!other' }]) },
        {
          path: '/groups/g1/owners',
          reply: page([{ id: 'u1', displayName: 'Jane', userPrincipalName: 'jane@contoso.com' }]),
        },
      ],
    );
    expect(outcome).toMatchObject({
      result: {
        displayName: 'Finance',
        storage: { library: 'Documents', usedGb: 5, totalGb: 10, percentUsed: 50, state: 'normal' },
        libraryCount: 2,
        groupConnected: true,
        group: { id: 'g1', owners: [{ displayName: 'Jane', userPrincipalName: 'jane@contoso.com' }] },
      },
    });
  });

  test('sites list: delegated needs --search; app-only lists all via getAllSites', async () => {
    const command = commandById(commands, 'sharepoint sites list');
    const error = (await runCommand(command, [], []).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('USAGE');
    const { ctx, calls } = mockContext([{ path: '/sites/getAllSites', reply: page([SITE]) }], {
      profile: { authMode: 'app' },
    });
    expect(await command.run(ctx, { top: 50, all: false })).toMatchObject({ items: [{ webUrl: SITE.webUrl }] });
    expect(calls[0]?.path).toBe('/sites/getAllSites');
  });
});

describe('files search', () => {
  test('tenant-wide search posts a driveItem search request and shapes hits', async () => {
    const command = commandById(commands, 'sharepoint files search');
    const { outcome, calls } = await runCommand(
      command,
      ['budget', '--top', '5'],
      [
        {
          method: 'POST',
          path: '/search/query',
          reply: {
            body: { value: [{ hitsContainers: [{ moreResultsAvailable: false, hits: [{ resource: ITEM }] }] }] },
          },
        },
      ],
    );
    expect(calls[0]?.body).toEqual({
      requests: [{ entityTypes: ['driveItem'], query: { queryString: 'budget' }, from: 0, size: 5 }],
    });
    expect(outcome).toMatchObject({
      result: { items: [{ driveId: 'b!drive', id: 'item1', path: '/Reports/Q1.xlsx' }], truncated: false },
    });
  });

  test('site search uses the library search function with a quoted query', async () => {
    const command = commandById(commands, 'sharepoint files search');
    const { calls } = await runCommand(
      command,
      ["o'neil", '--site', SITE_ID],
      [
        { path: `/sites/${SITE_ID}`, reply: { body: SITE } },
        { path: `/sites/${SITE_ID}/drive/root/search(q='o''neil')`, reply: page([ITEM]) },
      ],
    );
    expect(calls.at(-1)?.path).toBe(`/sites/${SITE_ID}/drive/root/search(q='o''neil')`);
  });
});

describe('onedrive', () => {
  test('onedrive get shapes quota for the resolved user', async () => {
    const command = commandById(commands, 'sharepoint onedrive get');
    const { outcome } = await runCommand(
      command,
      ['jane@contoso.com'],
      [
        { path: '/users/jane@contoso.com', reply: { body: { id: 'u1', userPrincipalName: 'jane@contoso.com' } } },
        {
          path: '/users/u1/drive',
          reply: {
            body: {
              id: 'd1',
              webUrl: 'https://contoso-my.sharepoint.com/personal/jane',
              quota: { used: 1024 ** 3 * 0.95, total: 1024 ** 3, state: 'nearing' },
            },
          },
        },
      ],
    );
    expect(outcome).toMatchObject({ result: { user: 'jane@contoso.com', percentUsed: 95, state: 'nearing' } });
  });

  const csv =
    'Report Refresh Date,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Storage Used (Byte),Storage Allocated (Byte),Owner Principal Name,Report Period\n' +
    '2026-11-08,https://contoso-my.sharepoint.com/personal/jane,Jane,False,2026-11-07,100,5,950,1000,jane@contoso.com,7\n' +
    '2026-11-08,https://contoso-my.sharepoint.com/personal/bob,Bob,False,2026-11-01,10,0,100,1000,bob@contoso.com,7\n';

  test('onedrive usage --near-quota keeps accounts at 90%+', async () => {
    const command = commandById(commands, 'sharepoint onedrive usage');
    const { outcome, calls } = await runCommand(
      command,
      ['--near-quota'],
      [{ path: "/reports/getOneDriveUsageAccountDetail(period='D7')", reply: { body: csv } }],
    );
    expect(calls[0]?.path).toBe("/reports/getOneDriveUsageAccountDetail(period='D7')");
    expect(outcome).toMatchObject({ result: { items: [{ ownerPrincipalName: 'jane@contoso.com', percentUsed: 95 }] } });
  });

  test('filtering by user on a concealed report explains how to unconceal', async () => {
    const command = commandById(commands, 'sharepoint onedrive usage');
    const hashed = csv.replace(/jane@contoso.com|bob@contoso.com/g, 'A1B2C3D4E5');
    const error = (await runCommand(
      command,
      ['--user', 'jane@contoso.com'],
      [{ path: "/reports/getOneDriveUsageAccountDetail(period='D7')", reply: { body: hashed } }],
    ).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('CONFIG');
    expect(error.hint).toContain('concealed');
  });
});

describe('onedrive grant / revoke', () => {
  const bob = { id: 'u-bob', displayName: 'Bob', userPrincipalName: 'bob@contoso.com', mail: 'bob@contoso.com' };
  const mgr = { id: 'u-mgr', displayName: 'Manager', userPrincipalName: 'mgr@contoso.com', mail: 'mgr@contoso.com' };
  const drive = {
    id: 'b!bob',
    webUrl: 'https://contoso-my.sharepoint.com/personal/bob_contoso_com/Documents',
    quota: { used: 2 * 1024 ** 3, total: 1024 ** 4 },
  };
  const routes = (): MockRoute[] => [
    { path: '/users/bob@contoso.com', reply: { body: bob } },
    { path: '/users/mgr@contoso.com', reply: { body: mgr } },
    { path: '/users/u-bob/drive', reply: { body: drive } },
    {
      method: 'POST',
      path: '/drives/b!bob/items/root/invite',
      reply: { body: { value: [{ id: 'perm-1', roles: ['write'], grantedToV2: { user: { id: 'u-mgr' } } }] } },
    },
    {
      path: '/drives/b!bob/items/root/permissions',
      reply: page([
        { id: 'owner', roles: ['owner'], grantedToV2: { user: { id: 'u-bob' } } },
        {
          id: 'direct',
          roles: ['write'],
          grantedToV2: { siteUser: { loginName: 'i:0#.f|membership|mgr@contoso.com' } },
        },
        {
          id: 'shared-link',
          roles: ['read'],
          link: { scope: 'users', type: 'view' },
          grantedToIdentitiesV2: [{ user: { id: 'u-mgr' } }, { user: { email: 'eve@contoso.com' } }],
        },
        { id: 'other', roles: ['read'], grantedToV2: { user: { email: 'eve@contoso.com' } } },
      ]),
    },
    { method: 'DELETE', path: '/drives/b!bob/items/root/permissions/direct', reply: { status: 204 } },
  ];

  test('grant previews owner, drive and grantee, then invites on the root with sign-in required', async () => {
    const command = commandById(commands, 'sharepoint onedrive grant');
    const argv = ['bob@contoso.com', 'mgr@contoso.com', '--role', 'write'];
    const first = await runCommand(command, argv, routes(), { allowWrites: true });
    expect(first.outcome).toMatchObject({
      kind: 'confirm',
      preview: {
        owner: 'bob@contoso.com',
        driveUrl: drive.webUrl,
        usedGb: 2,
        grantee: 'mgr@contoso.com',
        role: 'write',
      },
    });
    expect(JSON.stringify(first.outcome)).toContain("write access to all files in bob@contoso.com's OneDrive");
    expect(first.calls.some((c) => c.method === 'POST')).toBe(false);

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(command, argv, routes(), { allowWrites: true, confirm: token });
    expect(second.calls.find((c) => c.method === 'POST')?.body).toEqual({
      recipients: [{ email: 'mgr@contoso.com' }],
      requireSignIn: true,
      sendInvitation: false,
      roles: ['write'],
    });
    expect(second.outcome).toMatchObject({ result: { permissionId: 'perm-1', driveUrl: drive.webUrl } });
  });

  test('grant --notify --expires-days sends the invitation with an expiry', async () => {
    const command = commandById(commands, 'sharepoint onedrive grant');
    const argv = ['bob@contoso.com', 'mgr@contoso.com', '--notify', '--expires-days', '30'];
    const first = await runCommand(command, argv, routes(), { allowWrites: true });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(command, argv, routes(), { allowWrites: true, confirm: token });
    const body = second.calls.find((c) => c.method === 'POST')?.body as Record<string, unknown>;
    expect(body).toMatchObject({ sendInvitation: true, roles: ['read'] });
    const days = (Date.parse(String(body.expirationDateTime)) - Date.now()) / 86_400_000;
    expect(Math.round(days)).toBe(30);
  });

  test('grant explains a missing OneDrive', async () => {
    const command = commandById(commands, 'sharepoint onedrive grant');
    const error = (await runCommand(
      command,
      ['bob@contoso.com', 'mgr@contoso.com'],
      routes().filter((r) => r.path !== '/users/u-bob/drive'),
      { allowWrites: true },
    ).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('NOT_FOUND');
    expect(error.hint).toContain('retention');
  });

  test("revoke deletes only the grantee's own root permissions and leaves shared links", async () => {
    const command = commandById(commands, 'sharepoint onedrive revoke');
    const first = await runCommand(command, ['bob@contoso.com', 'mgr@contoso.com'], routes(), { allowWrites: true });
    expect(first.outcome).toMatchObject({
      kind: 'confirm',
      preview: { remove: [{ id: 'direct' }], sharedLinks: [{ id: 'shared-link' }] },
    });
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(command, ['bob@contoso.com', 'mgr@contoso.com'], routes(), {
      allowWrites: true,
      confirm: token,
    });
    expect(second.calls.filter((c) => c.method === 'DELETE').map((c) => c.path)).toEqual([
      '/drives/b!bob/items/root/permissions/direct',
    ]);
  });
});

describe('settings set', () => {
  test('previews current vs new sharing, then PATCHes sharingCapability', async () => {
    const command = commandById(commands, 'sharepoint settings set');
    const routes: MockRoute[] = [
      { path: '/admin/sharepoint/settings', reply: { body: { sharingCapability: 'externalUserAndGuestSharing' } } },
      {
        method: 'PATCH',
        path: '/admin/sharepoint/settings',
        reply: { body: { sharingCapability: 'externalUserSharingOnly' } },
      },
    ];
    const first = await runCommand(command, ['--sharing', 'externalUserSharingOnly'], routes, { allowWrites: true });
    expect(first.outcome).toMatchObject({
      kind: 'confirm',
      preview: { current: 'externalUserAndGuestSharing', new: 'externalUserSharingOnly' },
    });
    expect(first.calls.some((c) => c.method === 'PATCH')).toBe(false);
    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(command, ['--sharing', 'externalUserSharingOnly'], routes, {
      allowWrites: true,
      confirm: token,
    });
    expect(second.calls.find((c) => c.method === 'PATCH')?.body).toEqual({
      sharingCapability: 'externalUserSharingOnly',
    });
  });
});
