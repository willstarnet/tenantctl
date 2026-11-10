# SharePoint and OneDrive

`m365 sharepoint ...` covers sites, document libraries, files and their sharing, OneDrive storage, usage reports
and tenant sharing settings. OneDrive commands live here as `sharepoint onedrive ...`.

Sites are accepted as a URL (any page, library or file URL under the site), a Graph site id
(`host,guid,guid`) or a name. Files are accepted as a file/folder URL or sharing link, `driveId:itemId` (from
`files search`), or a path inside a site's default library with `--site` (`"Reports/Q1.xlsx" --site finance`).

## Common tasks

| Admin intent                                                  | Command                                                                                        |
| ------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- |
| Find a site                                                   | `m365 sharepoint sites list --search finance`                                                  |
| List every site in the tenant (app-only)                      | `m365 sharepoint sites list --all`                                                             |
| How full is this site / who owns it / is it a Teams site      | `m365 sharepoint sites get https://contoso.sharepoint.com/sites/finance`                       |
| Which document libraries does a site have                     | `m365 sharepoint sites drives <site>`                                                          |
| Which apps can access this site; group owner/member counts    | `m365 sharepoint sites permissions <site>`                                                     |
| Which sites are nearly full or inactive                       | `m365 sharepoint sites usage --near-quota`                                                     |
| Find a file                                                   | `m365 sharepoint files search "budget 2026"` / `... --site finance`                            |
| Details of a file (size, who changed it, is it shared)        | `m365 sharepoint files get <url>`                                                              |
| Who can see this file? Is it shared externally / anonymously? | `m365 sharepoint files permissions <url>`                                                      |
| Remove a sharing link or someone's access                     | `m365 sharepoint files unshare <url> <permissionId>`                                           |
| Kill every sharing link on a file                             | `m365 sharepoint files unshare <url> --all-links`                                              |
| Someone overwrote a file; roll it back                        | `m365 sharepoint files versions <url>` then `files restore-version <url> <versionId>`          |
| Is Jane's OneDrive full / where is it                         | `m365 sharepoint onedrive get jane@contoso.com`                                                |
| Which OneDrives are near quota                                | `m365 sharepoint onedrive usage --near-quota`                                                  |
| Departing employee: give the manager access to their OneDrive | `m365 sharepoint onedrive grant bob@contoso.com manager@contoso.com [--role write] [--notify]` |
| Take that access away again                                   | `m365 sharepoint onedrive revoke bob@contoso.com manager@contoso.com`                          |
| Who deleted a file on this site / OneDrive                    | `m365 purview audit file-deleted --site <url> --file <name>` (see purview.md)                  |
| Can people share externally? Anonymous links allowed?         | `m365 sharepoint settings`                                                                     |
| Restrict external sharing tenant-wide                         | `m365 sharepoint settings set --sharing externalUserSharingOnly`                               |

Reading `files permissions`: `kind` is `link`, `direct` or `inherited` (from a parent folder or the library;
`inheritedFrom` says where). Link `linkScope` is `anonymous` (anyone with the URL, no sign-in),
`organization` (anyone in the tenant), `users` (named people) or `existingAccess`. `anonymousLinks` and
`externalUsers` (outside the tenant's verified domains, or guest `#ext#` accounts) summarize exposure. Inherited
permissions cannot be removed on the item: unshare them on the parent.

## Permissions, roles and licensing

| Commands                                             | Delegated                                                                     | Application                        | Notes                                                                                                                |
| ---------------------------------------------------- | ----------------------------------------------------------------------------- | ---------------------------------- | -------------------------------------------------------------------------------------------------------------------- |
| `sites list/get/drives`, `files search/get/versions` | Sites.Read.All                                                                | Sites.Read.All                     | `sites get` adds GroupMember.Read.All for group owners. `sites list` without `--search` is app-only (`getAllSites`). |
| `files permissions`                                  | Sites.Read.All, Domain.Read.All                                               | same                               | Domain.Read.All is for external-user detection; without it only guest accounts are flagged.                          |
| `files unshare`, `files restore-version`             | Sites.ReadWrite.All                                                           | Sites.ReadWrite.All                |                                                                                                                      |
| File refs given as URLs                              | Files.ReadWrite.All or Sites.ReadWrite.All                                    | same                               | The `/shares` API needs write scope even to read. Use `--site` + path or `driveId:itemId` to stay read-only.         |
| `sites permissions`                                  | Sites.FullControl.All, GroupMember.Read.All                                   | same                               | Reading Sites.Selected app grants needs full control.                                                                |
| `onedrive get`                                       | Files.Read.All, User.ReadBasic.All                                            | Files.Read.All, User.Read.All      |                                                                                                                      |
| `onedrive grant`, `onedrive revoke`                  | Files.ReadWrite.All, User.ReadBasic.All                                       | Files.ReadWrite.All, User.Read.All | Destructive. Delegated callers must already be able to open the owner's OneDrive; use app-only for handovers.        |
| `onedrive usage`, `sites usage`                      | Reports.Read.All + Reports Reader role                                        | Reports.Read.All                   | Reports lag 24-48 h.                                                                                                 |
| `settings`                                           | SharePointTenantSettings.Read.All + SharePoint Administrator or Global Reader | same                               |                                                                                                                      |
| `settings set`                                       | SharePointTenantSettings.ReadWrite.All + SharePoint Administrator             | same                               | Destructive: tenant-wide.                                                                                            |

- Delegated sign-ins only see sites and files the signed-in person can open; the SharePoint Administrator role
  does not grant file access. Use an app-only profile for tenant-wide file and sharing investigations.
- App-only tenant-wide `files search` searches shared content in one geo region; pass `--region` (e.g. `NAM`,
  `EUR`) when the tenant does not publish a data location code.
- Usage reports show hashed names when "Display concealed user, group, and site names in all reports" is on
  (Microsoft 365 admin center > Settings > Org settings > Reports); `--user`/`--site` filters fail until it is off.
- OneDrive/SharePoint needs a SharePoint Online plan (included in Microsoft 365 Business/E plans). A OneDrive
  only exists after the user first opens it.

## Not available via Graph

| Task                                                                                                                   | Use instead                                                                                                                                                                                                                                                                  |
| ---------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| List or change site collection administrators                                                                          | SharePoint Online Management Shell (`Get-SPOUser -Site ... \| ? IsSiteAdmin`, `Set-SPOUser -IsSiteCollectionAdmin`) or PnP PowerShell (`Get-PnPSiteCollectionAdmin`)                                                                                                         |
| Classic SharePoint groups (Owners/Members/Visitors) and their members, per-user site permissions                       | PnP PowerShell (`Get-PnPGroup`, `Get-PnPGroupMember`, `Set-PnPWebPermission`) or SharePoint REST                                                                                                                                                                             |
| Per-site external sharing setting, site lock state, site storage quota changes                                         | `Set-SPOSite -SharingCapability / -LockState / -StorageQuota`                                                                                                                                                                                                                |
| Create non-group sites (communication/classic team sites), delete or restore sites                                     | `New-SPOSite`/`New-PnPSite`, `Remove-SPOSite`, `Restore-SPODeletedSite`. Graph has site create and archive only in beta (not implemented here); group-connected sites are created by creating a Microsoft 365 group.                                                         |
| Tenant-wide "what is shared externally" reports                                                                        | SharePoint Advanced Management data access governance reports or PnP sharing reports. Sharing activity is in the audit log (`m365 purview audit search --record-type sharePointSharingOperation`); Graph only answers current state per item (`files permissions`).          |
| Retention policies and non-eDiscovery holds on sites/OneDrive                                                          | Security & Compliance PowerShell (`Get-RetentionCompliancePolicy`). Preserving one person's OneDrive + mailbox is covered by `m365 purview ediscovery hold <case> <user>`. Retention labels are in Graph under `/security/labels/retentionLabels`, not covered by this area. |
| Make someone site collection admin / secondary owner of a OneDrive (full control, incl. items with unique permissions) | `Set-SPOUser -IsSiteCollectionAdmin` on the OneDrive URL, or the Microsoft 365 admin center. Read/write access to all files for a manager is covered by `onedrive grant` (shares the root folder).                                                                           |
| Restore files from the site/OneDrive recycle bin                                                                       | PnP PowerShell (`Restore-PnPRecycleBinItem`); Graph's driveItem restore works only for OneDrive Personal                                                                                                                                                                     |
| Hub sites, site designs/scripts, sensitivity labels on sites                                                           | SharePoint Online Management Shell / PnP PowerShell                                                                                                                                                                                                                          |
