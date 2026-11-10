# Defender (`m365 defender`)

Security incidents and alerts from every Microsoft Defender product, advanced hunting (KQL), Microsoft Secure Score
and email threat submissions, all through the Microsoft Graph security API (`/security/...`). Device inventory and
response actions (`machines ...`) go to the Defender for Endpoint API (`https://api.security.microsoft.com/api`).

## Common tasks

| You want to...                                            | Command                                                                                                                             |
| --------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| See what is open right now                                | `m365 defender incidents list --status active`                                                                                      |
| Find high-severity incidents nobody owns                  | `m365 defender incidents list --severity high --unassigned`                                                                         |
| Explain an incident (alerts, users, devices, IPs, emails) | `m365 defender incidents get <id>`                                                                                                  |
| Take ownership / start working an incident                | `m365 defender incidents update <id> --assign-to me@contoso.com --status inProgress`                                                |
| Close an incident as a false positive                     | `m365 defender incidents update <id> --status resolved --classification falsePositive --determination notMalicious --comment "..."` |
| List new alerts from one product                          | `m365 defender alerts list --status new --source microsoftDefenderForOffice365`                                                     |
| List alerts for an ATT&CK tactic                          | `m365 defender alerts list --category CredentialAccess --days 7`                                                                    |
| Inspect one alert and its evidence                        | `m365 defender alerts get <id>`                                                                                                     |
| "Has anything flagged Jane?"                              | `m365 defender alerts for-user jane@contoso.com --days 30`                                                                          |
| "Is this laptop involved in any alerts?"                  | `m365 defender alerts for-device DESKTOP-123`                                                                                       |
| See ready-made hunting queries                            | `m365 defender hunt templates`                                                                                                      |
| Who received mail from a sender                           | `m365 defender hunt --template emails-from-sender --param sender=bad@evil.example`                                                  |
| Find a phishing email by subject                          | `m365 defender hunt --template email-by-subject --param "subject=Invoice overdue"`                                                  |
| Did a user click a malicious link                         | `m365 defender hunt --template url-clicks-for-user --param user=jane@contoso.com`                                                   |
| Where a user signed in from                               | `m365 defender hunt --template signins-for-user --param user=jane@contoso.com --timespan P7D`                                       |
| What ran / connected / logged on to a device              | `m365 defender hunt --template device-processes --param device=DESKTOP-123` (also `device-network`, `device-logons`)                |
| How widespread a file is                                  | `m365 defender hunt --template file-hash-prevalence --param sha256=<hash>`                                                          |
| Suspicious inbox rules for a mailbox                      | `m365 defender hunt --template inbox-rule-changes --param user=jane@contoso.com`                                                    |
| Run your own KQL                                          | `m365 defender hunt "DeviceInfo \| summarize arg_max(Timestamp, *) by DeviceName \| take 20"`                                       |
| "How secure are we / what should we fix first?"           | `m365 defender secure-score --top-actions 10`                                                                                       |
| Report a phishing / wrongly junked email to Microsoft     | `m365 defender threat-submit email jane@contoso.com <messageId> --category phishing`                                                |
| Defender's view of devices (risk, health, exposure)       | `m365 defender machines list --risk high` / `m365 defender machines list --health inactive --stale-days 30`                         |
| One machine: open alerts, recent response actions         | `m365 defender machines get LAPTOP-7` (name, Defender machine id or Entra device id)                                                |
| Cut a compromised laptop off the network                  | `m365 defender machines isolate LAPTOP-7 --comment "Ransomware alert 1234" [--selective]`                                           |
| Reconnect it after remediation                            | `m365 defender machines release LAPTOP-7 --comment "Cleared by SOC"`                                                                |
| Run an antivirus scan                                     | `m365 defender machines scan LAPTOP-7 [--full-scan] --comment "Alert follow-up"`                                                    |
| Collect forensics and download them                       | `m365 defender machines collect-package LAPTOP-7 --comment "..."`, then `m365 defender machines package-url <actionId>`             |
| Only allow Microsoft-signed apps to run / undo            | `m365 defender machines restrict LAPTOP-7 --comment "..."` / `m365 defender machines unrestrict LAPTOP-7 --comment "..."`           |
| Kill and quarantine a malicious file                      | `m365 defender machines stop-quarantine-file LAPTOP-7 --sha1 <sha1> --comment "..."`                                                |
| Did the isolation / scan go through?                      | `m365 defender machines actions --machine LAPTOP-7` (`--status Pending` for everything in flight)                                   |

`--param` values are inserted as escaped KQL string literals, so quotes in names or subjects are safe. `--timespan`
accepts ISO 8601 (`P7D`, `PT12H`, `2026-11-01T00:00:00Z/P7D`) or shorthand (`7d`, `12h`); the default is 30 days.

`machines` response actions need `--comment` (recorded in the Defender action center) and return the machine action
id and status; `isolate`, `restrict` and `stop-quarantine-file` are destructive (preview + confirm token). Actions a
platform does not support (selective isolation or restrict on non-Windows, for example) are refused before anything
is sent.

## Licensing and roles

| Command                                                                                                             | Licence                                                                                                                                   | Graph permission (least privilege)                                                                                                                                                          | Admin role (delegated)                                        |
| ------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------- |
| `incidents list/get`                                                                                                | Microsoft Defender XDR (any Defender workload, e.g. Microsoft 365 E5 / E5 Security, Defender for Endpoint P2, Defender for Office 365 P2) | SecurityIncident.Read.All                                                                                                                                                                   | Security Reader                                               |
| `incidents update`                                                                                                  | as above                                                                                                                                  | SecurityIncident.ReadWrite.All                                                                                                                                                              | Security Operator                                             |
| `alerts list/get`                                                                                                   | as above                                                                                                                                  | SecurityAlert.Read.All                                                                                                                                                                      | Security Reader                                               |
| `alerts update`                                                                                                     | as above                                                                                                                                  | SecurityAlert.ReadWrite.All                                                                                                                                                                 | Security Operator                                             |
| `hunt`, `alerts for-user/for-device`                                                                                | Defender XDR; each table needs its product (see `hunt templates` → `requires`)                                                            | ThreatHunting.Read.All (+ SecurityAlert.Read.All, User.Read.All for for-user)                                                                                                               | Security Reader                                               |
| `secure-score`                                                                                                      | Any Microsoft 365 / Entra subscription                                                                                                    | SecurityEvents.Read.All                                                                                                                                                                     | Security Reader or Global Reader                              |
| `threat-submit email` (beta API)                                                                                    | Exchange Online Protection or Defender for Office 365                                                                                     | ThreatSubmission.ReadWrite (delegated) / ThreatSubmission.ReadWrite.All (app), + User.Read.All                                                                                              | Security Administrator                                        |
| `machines list/get/actions`                                                                                         | Defender for Endpoint P1/P2 or Defender for Business                                                                                      | Defender for Endpoint API: Machine.Read.All (get also Machine.ReadWrite.All, Alert.ReadWrite.All; actions Machine.ReadWrite.All)                                                            | Security Reader                                               |
| `machines isolate/release`, `scan`, `collect-package`, `restrict/unrestrict`, `stop-quarantine-file`, `package-url` | as above                                                                                                                                  | Machine.Isolate, Machine.Scan, Machine.CollectForensics, Machine.RestrictExecution, Machine.StopAndQuarantine (package-url: Machine.ReadWrite.All app / Machine.CollectForensics delegated) | Security Operator (Defender RBAC: Active remediation actions) |

Advanced hunting limits (Graph `runHuntingQuery`): data covers the last 30 days, a query returns at most 100,000 rows
and 50 MB, runs time out after 3 minutes, and each tenant has a CPU quota (HTTP 429 when exceeded; it resets on the
next 15-minute cycle). `signins-for-user` reads `EntraIdSignInEvents`, which needs Microsoft Entra ID P2.

`machines` commands call the Defender for Endpoint API, whose permissions are `WindowsDefenderATP` permissions, not
Graph scopes. Delegated use needs your own app registration with those delegated permissions; app-only use needs
them as application permissions with admin consent. Defender rate-limits actions to 100 calls per minute.

## Not available via Graph

| Task                                                                         | Use instead                                                                                                                                                                                   |
| ---------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Live response sessions (run scripts, get files on a device)                  | Defender for Endpoint API (`/api/machines/{id}/runliveresponse`) or the Defender portal                                                                                                       |
| Manage the Tenant Allow/Block List (senders, URLs, files, spoofed senders)   | Exchange Online PowerShell (`New-TenantAllowBlockListItems`, `Get-TenantAllowBlockListItems`) or the Defender portal. Graph only adds an allow entry as a side effect of an email submission. |
| Release, preview or delete quarantined email                                 | Exchange Online PowerShell (`Get-QuarantineMessage`, `Release-QuarantineMessage`) or the Defender portal                                                                                      |
| Soft/hard delete a malicious email from mailboxes (purge)                    | Defender portal Threat Explorer / Take action, or `New-ComplianceSearchAction -Purge` in Security & Compliance PowerShell                                                                     |
| Merge incidents or move alerts between incidents                             | Graph has this (`incident: mergeIncidents`, `alert: moveAlerts`) but this CLI does not wrap it; use `m365 graph post`.                                                                        |
| Create custom detection rules                                                | Defender portal (Graph `security/rules/detectionRules` is beta and not wrapped here)                                                                                                          |
| Defender for Endpoint exposure score, vulnerabilities (TVM)                  | Defender for Endpoint API, or hunting tables `DeviceInfo`, `DeviceTvmSoftwareVulnerabilities` via `m365 defender hunt`; per-machine `exposureLevel` is in `m365 defender machines list`       |
| Configure anti-phishing / anti-spam / Safe Links / Safe Attachments policies | Exchange Online PowerShell (`Set-AntiPhishPolicy`, `Set-SafeLinksPolicy`, ...) or the Defender portal                                                                                         |
| Mark a Secure Score action as resolved through third party / risk accepted   | Graph supports PATCH on `secureScoreControlProfiles` but this CLI does not wrap it; use the Defender portal.                                                                                  |
