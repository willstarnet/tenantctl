# Exchange Online (`m365 exchange`)

Mailboxes are users: every `<user>` accepts a UPN, email/alias, display name, object id or `me`.

## Common tasks

| Admin intent                                                | Command                                                                                                                                                           |
| ----------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| "What kind of mailbox is this / is out-of-office on?"       | `m365 exchange mailbox jane@contoso.com`                                                                                                                          |
| Check someone's out-of-office                               | `m365 exchange auto-reply get jane@contoso.com`                                                                                                                   |
| Turn on out-of-office for a leaver                          | `m365 exchange auto-reply set jane@contoso.com --on --internal-message "Jane has left; contact bob@contoso.com" --external-audience all --external-message "..."` |
| Schedule out-of-office                                      | `m365 exchange auto-reply set jane@contoso.com --scheduled --start 2026-12-20T17:00 --end 2027-01-05T08:00`                                                       |
| Turn out-of-office off                                      | `m365 exchange auto-reply set jane@contoso.com --off`                                                                                                             |
| "Is Jane's mail being forwarded anywhere?"                  | `m365 exchange forwarding jane@contoso.com`                                                                                                                       |
| Look for malicious inbox rules (compromised account)        | `m365 exchange rules list jane@contoso.com`                                                                                                                       |
| Sweep the tenant / a department for forwarding or BEC rules | `m365 exchange rules audit` / `m365 exchange rules audit --group "Finance"`                                                                                       |
| Remove / pause a bad rule                                   | `m365 exchange rules delete jane@contoso.com "<rule name or id>"` / `m365 exchange rules disable ...`                                                             |
| "My email to X never arrived"                               | `m365 exchange trace --sender jane@contoso.com --recipient x@partner.com --days 3`, then `m365 exchange trace details <id> --recipient x@partner.com`             |
| "I'm not getting mail from vendor.com"                      | `m365 exchange trace --recipient jane@contoso.com --status filteredAsSpam` (also try `quarantined`, `failed`)                                                     |
| Find a message in a mailbox                                 | `m365 exchange messages search jane@contoso.com --from billing@vendor.com --since 7`                                                                              |
| Purge a phishing message from a mailbox                     | `m365 exchange messages delete jane@contoso.com "<internet-message-id>" [--hard]`                                                                                 |
| Who can see/edit a calendar                                 | `m365 exchange calendar permissions jane@contoso.com`                                                                                                             |
| What's on someone's / a room's calendar                     | `m365 exchange calendar events "Room 4.01" --days 3`                                                                                                              |
| Whose mailbox is nearly full?                               | `m365 exchange usage --near-quota` / `m365 exchange usage --user jane`                                                                                            |
| Forward a leaver's mail to their manager                    | `m365 exchange forwarding set jane@contoso.com --to bob@contoso.com --keep-copy`                                                                                  |
| Stop mailbox-level forwarding                               | `m365 exchange forwarding clear jane@contoso.com`                                                                                                                 |
| Who has access to this mailbox?                             | `m365 exchange permissions list finance@contoso.com`                                                                                                              |
| Give / remove FullAccess, Send As, Send on Behalf           | `m365 exchange permissions add finance@contoso.com bob@contoso.com --access full` / `permissions remove ... --access sendas`                                      |
| Create a shared mailbox for a team                          | `m365 exchange shared-mailbox create --name "Finance" --email finance@contoso.com --members jane@contoso.com,bob@contoso.com`                                     |
| Convert a leaver's mailbox to shared                        | `m365 exchange mailbox convert jane@contoso.com --to shared`                                                                                                      |
| Distribution lists: find, members, add/remove, create       | `m365 exchange dl list --search finance`, `dl members <dl>`, `dl add-member <dl> jane bob`, `dl remove-member <dl> bob`, `dl create --name ... --email ...`       |
| "Mail from X went to quarantine"                            | `m365 exchange quarantine list --recipient jane@contoso.com --days 14`, then `quarantine release "<identity>"` or `quarantine delete "<identity>"`                |
| Litigation hold on/off                                      | `m365 exchange hold jane@contoso.com --on --duration-days 2555 --note "Case 42"` / `--off`                                                                        |
| Enable the online archive                                   | `m365 exchange archive enable jane@contoso.com`                                                                                                                   |
| Is a mail flow rule interfering?                            | `m365 exchange transport-rules`                                                                                                                                   |
| Mailbox type, aliases, size, last logon, forwarding         | `m365 exchange mailbox details jane@contoso.com`                                                                                                                  |
| Find a free meeting room                                    | `m365 exchange rooms free --start 2026-11-12T14:00:00Z --end 2026-11-12T15:00:00Z --capacity 6` (`rooms list --building HQ`)                                      |

## Reading the output

- `rules list` flags per rule: `forwards`, `forwardsExternally` (target domain is not a verified tenant domain;
  `null` when `/domains` can't be read), `deletes`, `movesToFolder`, `marksRead`, and `suspicious` + `reason`.
  Suspicious means external forwarding, or delete/move to a rarely opened folder (RSS Feeds, Archive,
  Conversation History, ...) combined with mark-as-read, payment/security keywords, or a meaningless rule name:
  the classic business email compromise pattern. Next: check sign-ins, reset credentials, delete the rule.
- `trace` returns one row per recipient; `status` is delivered, failed, pending, expanded, quarantined,
  filteredAsSpam. `trace details` lists the hop events; `finalEvent` usually explains the outcome.
- `usage` `percentUsed` is storage used against the Prohibit Send quota.

## Prerequisites

- **Other people's mailboxes need app-only auth.** Delegated Graph tokens reach only the signed-in user's
  mailbox and mailboxes shared with them. Mailbox, rules, messages and calendar commands work tenant-wide with an
  app registration holding the application permissions (`MailboxSettings.Read[Write]`, `Mail.ReadBasic.All`,
  `Mail.ReadWrite`, `Calendars.ReadBasic`). Exchange
  [application access policies / RBAC for Applications](https://learn.microsoft.com/exchange/permissions-exo/application-rbac)
  can scope that app to specific mailboxes.
- **Message trace** (`trace`, `trace details`): `ExchangeMessageTrace.Read.All`, and a one-time service principal
  for appId `8bd644d1-64a1-4d4b-ae52-2e0cbf64e373` (`New-MgServicePrincipal -AppId 8bd644d1-64a1-4d4b-ae52-2e0cbf64e373`).
  Window: at most 10 days per query, within the last 90 days. Throttled at 100 requests / 5 minutes.
- **Usage report**: `Reports.Read.All`; delegated callers also need a reports-capable role (Reports Reader, Global
  Reader, Exchange Administrator). If "Display concealed user, group, and site names in all reports" is on (admin
  center > Settings > Org settings > Reports), names are hashed.
- **Rules audit / forwarding** read `/domains` (`Domain.Read.All`) to tell internal from external targets; the
  audit's default scope lists licensed users (`User.Read.All`), `--group` needs `GroupMember.Read.All`.
- Exchange Online licence on the target mailbox; mailbox-less users fail with `MailboxNotEnabledForRESTAPI`.

## Exchange Online PowerShell bridge

`forwarding` (mailbox level), `forwarding set|clear`, `permissions`, `shared-mailbox create`, `mailbox convert`,
`mailbox details`, `dl`, `quarantine`, `hold`, `archive enable` and `transport-rules` run Exchange cmdlets through
the Exchange Online PowerShell bridge (needs pwsh 7 + the ExchangeOnlineManagement module 3.1+). Permissions are
Exchange RBAC: delegated `Exchange.Manage` with the Exchange Administrator role, or app-only
`Exchange.ManageAsApp` with an Exchange Administrator role assignment. Quarantine actions need Security
Administrator (or the Quarantine Administrator role group). Mailbox and group arguments accept email, alias,
UPN or display name; Exchange resolves them and an ambiguous name fails with `AMBIGUOUS`. If the bridge is
unavailable, `exchange forwarding` still returns inbox-rule forwarding and reports `mailboxForwarding.error`.

Rooms (`rooms list`, `rooms free`) use Graph Places (`Place.Read.All`) and getSchedule (`Calendars.ReadBasic`).

## Not covered by this CLI

Use [Exchange Online PowerShell](https://learn.microsoft.com/powershell/exchange/exchange-online-powershell)
(`Connect-ExchangeOnline`) or the Exchange admin center for these:

| Task                                                                                    | Exchange Online PowerShell                                                                                   |
| --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| Create room/equipment mailboxes, booking policies                                       | `New-Mailbox -Room`, `Set-CalendarProcessing`                                                                |
| Create or change mail flow (transport) rules (listing: `m365 exchange transport-rules`) | `New-/Set-TransportRule`                                                                                     |
| Inbound/outbound connectors                                                             | `Get-InboundConnector`, `Get-OutboundConnector`                                                              |
| eDiscovery / retention holds (litigation hold: `m365 exchange hold`)                    | Purview / `m365 graph` security APIs                                                                         |
| Auto-expanding archive (archive: `m365 exchange archive enable`)                        | `Enable-Mailbox -AutoExpandingArchive`                                                                       |
| Per-folder sizes, mailbox quotas changes                                                | `Get-MailboxFolderStatistics`, `Set-Mailbox -ProhibitSendQuota` (Graph v1.0 mailFolder has item counts only) |
| Change hidden-from-GAL or aliases (read: `m365 exchange mailbox details`)               | `Set-Mailbox -HiddenFromAddressListsEnabled`, `Set-Mailbox -EmailAddresses @{add=...}`                       |
| Message trace older than 90 days                                                        | `Start-HistoricalSearch`                                                                                     |
| Tenant-wide search and purge of a message                                               | Purview content search + `New-ComplianceSearchAction -Purge`, or Defender Threat Explorer                    |
| Outlook on the web / POP / IMAP / SMTP AUTH per mailbox                                 | `Set-CASMailbox`                                                                                             |
| Hidden inbox rules (some attacker-created rules don't appear in Graph or Outlook)       | `Get-InboxRule -Mailbox jane -IncludeHidden`                                                                 |
