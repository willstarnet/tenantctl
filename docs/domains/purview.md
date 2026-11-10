# Microsoft Purview (`m365 purview`)

`m365 purview ...` searches the unified audit log (who did what, across Exchange, SharePoint/OneDrive, Entra,
Teams and more) through the Graph Audit Search API, and manages eDiscovery (Premium) cases, custodians and holds
through Graph v1.0 `/security/cases/ediscoveryCases`.

## Common tasks

| Admin intent                                               | Command                                                                                                       |
| ---------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| Who deleted this file / what was deleted on this site      | `m365 purview audit file-deleted --site https://contoso.sharepoint.com/sites/finance --file Q1.xlsx`          |
| What was deleted in Bob's OneDrive this month              | `m365 purview audit file-deleted --site https://contoso-my.sharepoint.com/personal/bob_contoso_com --days 30` |
| Who has been reading the CEO's mailbox (not the CEO)       | `m365 purview audit mailbox-access ceo@contoso.com --non-owner --days 14`                                     |
| What did Bob do in someone else's mailbox                  | `m365 purview audit mailbox-access shared@contoso.com --accessed-by bob@contoso.com`                          |
| Who changed this mailbox / what did this admin change      | `m365 purview audit admin-changes --user admin@contoso.com` / `... --operation Set-Mailbox`                   |
| What did Jane do in the last two days (any service)        | `m365 purview audit search --user jane@contoso.com --days 2`                                                  |
| Activity from a suspicious IP                              | `m365 purview audit search --ip 203.0.113.5 --days 30`                                                        |
| Specific operations in a date range                        | `m365 purview audit search --operation FileDownloaded,FileAccessed --start 2026-11-01 --end 2026-11-05`       |
| A search outlived its wait; get the records later          | `m365 purview audit results <queryId>`                                                                        |
| Which audit searches exist / are still running             | `m365 purview audit queries`                                                                                  |
| List eDiscovery cases                                      | `m365 purview ediscovery list`                                                                                |
| Open a case for an investigation                           | `m365 purview ediscovery create --name "HR-2026-014" --description "Departure review"`                        |
| Preserve a departing/suspected user's mailbox and OneDrive | `m365 purview ediscovery hold "HR-2026-014" bob@contoso.com`                                                  |
| Is the hold applied yet / who is held in this case         | `m365 purview ediscovery custodians "HR-2026-014"`                                                            |
| Lift the hold on one person                                | `m365 purview ediscovery release "HR-2026-014" bob@contoso.com`                                               |

## Reading the output

- Audit searches are asynchronous: the CLI creates a query, polls it every ~5 s for up to `--wait-seconds`
  (default 120) and pages the records. If Purview is still working, the result is
  `{ queryId, status, next }`; run `next` later. Nothing is lost and no tenant setting changes, so all audit
  commands are `read`.
- Records: `time`, `operation`, `service`, `user` (who acted), `ip`, `objectId` (file URL, mailbox item or
  changed object), `auditLogRecordType`, and `details` with the useful audit data: `file`/`site`/`folder` for
  SharePoint/OneDrive; `mailbox`, `logonType` (owner, delegate, admin), `folders`, `subjects`, `accessType`
  (Bind = one item opened, Sync = folder synced by a client), `client` for mailbox activity; cmdlet
  `parameters` for Exchange admin; `targets`/`modified` for Entra. `--full` returns the raw records.
- `--keyword` matches free text inside the workload-specific audit data, not operation names or users; use
  `--operation`/`--user` for those. `--user` is the person who performed the action. Mailbox audit records are
  filed under the actor, which is why `mailbox-access` matches the owner via the audit data.
- Presets send the site URL/file name or mailbox UPN as the keyword and then keep only records whose
  `SiteUrl`/`SourceFileName` or `MailboxOwnerUPN` match exactly, so `--top` counts records before that filter.
- Operation names used by presets: `file-deleted` = FileDeleted, FileDeletedFirstStageRecycleBin,
  FileDeletedSecondStageRecycleBin and the Folder\* equivalents (SecondStage = purged from the site collection
  recycle bin); `mailbox-access` = MailItemsAccessed, FolderBind, SendAs, SendOnBehalf, MoveToDeletedItems,
  SoftDelete, HardDelete, UpdateInboxRules, UpdateFolderPermissions; `admin-changes` = record types
  exchangeAdmin and azureActiveDirectory.
- eDiscovery `holdStatus`: notApplied, applying, applied, removing, partial (some sources failed). Holds take
  minutes to hours to apply; held content cannot be purged by the user or by retention policies until released.

## Licensing, permissions and roles

| Commands                         | Delegated / application permission                                                   | Purview role (delegated)                                        |
| -------------------------------- | ------------------------------------------------------------------------------------ | --------------------------------------------------------------- |
| `audit search/results/queries`   | AuditLogsQuery.Read.All                                                              | Audit Reader (View-Only Audit Logs) or Compliance Administrator |
| `audit file-deleted`             | AuditLogsQuery-SharePoint.Read.All, AuditLogsQuery-OneDrive.Read.All, Sites.Read.All | same                                                            |
| `audit mailbox-access`           | AuditLogsQuery-Exchange.Read.All (+ User.ReadBasic.All / User.Read.All)              | same                                                            |
| `audit admin-changes`            | AuditLogsQuery-Exchange.Read.All, AuditLogsQuery-Entra.Read.All                      | same                                                            |
| `ediscovery list/custodians`     | eDiscovery.Read.All                                                                  | eDiscovery Manager, member of the case                          |
| `ediscovery create/hold/release` | eDiscovery.ReadWrite.All (+ User.ReadBasic.All / User.Read.All)                      | eDiscovery Manager, member of the case                          |

- **Audit (Standard)** (Microsoft 365 E3/Business Premium and most business plans) is on by default and keeps
  records for **180 days** (90 days for records created before 17 October 2023). **Audit (Premium)** (E5,
  E5 Compliance, or the Audit Premium add-on) keeps Entra, Exchange, SharePoint and OneDrive records for
  **one year** by default, supports custom audit retention policies, and up to 10 years with the 10-Year Audit
  Log Retention add-on. A search older than the retention period simply returns no records.
- **MailItemsAccessed** is logged by default for owner, delegate and admin access. It was Premium-only until
  Microsoft's 2024 expanded logging (message center MC711333) brought it, `Send` and `SearchQueryInitiated*`
  to Audit (Standard); the `SensitivityLabel` property on MailItemsAccessed still needs Audit (Premium)
  licenses on the audited users. If a mailbox shows no MailItemsAccessed records, check that mailbox auditing
  is on (`Get-Mailbox <upn> | fl AuditEnabled`) and that the tenant has not customized the audited actions.
- The per-service `AuditLogsQuery-*` scopes only return that service's records; general `audit search`
  therefore declares the cross-service `AuditLogsQuery.Read.All`. The Graph page for reading a single audit
  query lists ThreatIntelligence.Read.All; in practice the audit scopes above are used for the whole flow.
- **eDiscovery (Premium)** (E5 / E5 Compliance / eDiscovery and Audit add-on) is what the Graph eDiscovery API
  manages; custodians and custodian holds are Premium features. eDiscovery (Standard) cases are managed in the
  Purview portal / Security & Compliance PowerShell (`New-ComplianceCase`, `New-CaseHoldPolicy`).
- **App-only eDiscovery** works but needs more than Graph consent: grant eDiscovery.ReadWrite.All, add the app's
  service principal object id to the `eDiscoveryManager` role group (`Add-RoleGroupMember`) and, for access to
  all cases, make it an eDiscovery Administrator (`Add-eDiscoveryCaseAdmin`). Delegated callers must be
  eDiscovery Manager/Administrator and a member of the case.
- The Graph permission tables list eDiscovery.Read.All as least privileged even for creating cases and
  applying holds; the CLI declares eDiscovery.ReadWrite.All for those writes.
- Audit searches by restricted (administrative-unit scoped) admins only return records for users in their units.

## Not available via Graph

| Task                                                                                        | Use instead                                                                                                                     |
| ------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| Turn unified audit logging on/off, mailbox audit settings (`AuditEnabled`, audited actions) | Exchange Online PowerShell `Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled`, `Set-Mailbox -AuditOwner/-AuditDelegate` |
| Audit log retention policies                                                                | Security & Compliance PowerShell `New-UnifiedAuditLogRetentionPolicy` or the Purview portal                                     |
| Exporting very large audit result sets / continuous ingestion                               | Office 365 Management Activity API or Microsoft Sentinel connector                                                              |
| eDiscovery (Standard) cases and case holds                                                  | Security & Compliance PowerShell `New-ComplianceCase`, `New-CaseHoldPolicy`, `New-CaseHoldRule`                                 |
| eDiscovery searches, review sets and exports                                                | Graph `/security/cases/ediscoveryCases/{id}/searches` and `reviewSets` (not implemented here) or the Purview portal             |
| Litigation hold on a whole mailbox (no case)                                                | Exchange Online PowerShell `Set-Mailbox -LitigationHoldEnabled $true`                                                           |
| Retention policies/labels, DLP policies, sensitivity labels, Insider Risk                   | Purview portal / Security & Compliance PowerShell (`Get-RetentionCompliancePolicy`, `Get-DlpCompliancePolicy`)                  |
| Cancel or delete an audit search                                                            | Purview portal (Audit > search jobs); the Graph API has no delete for audit queries                                             |
