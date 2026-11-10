# Intune (`m365 intune`)

Managed devices, compliance, configuration profiles, apps, Windows Autopilot and remote device actions, through
the Microsoft Graph Intune API (`/deviceManagement/...`, `/deviceAppManagement/...`).

Every `<device>` accepts a device name (case-insensitive), serial number, Intune device id, Entra device id
(`azureADDeviceId`) or an Intune admin center URL containing `mdmDeviceId/<id>`. Names shared by several devices
(re-enrolled or re-imaged machines) return `AMBIGUOUS` with id, serial, user, OS and last sync for each; pass the
id of the right one.

## Common tasks

| Admin intent                                                              | Command                                                                                                                                         |
| ------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| "What devices does Jane have?"                                            | `m365 intune devices list --user jane@contoso.com`                                                                                              |
| Find non-compliant Windows devices                                        | `m365 intune devices list --os windows --compliance noncompliant`                                                                               |
| Find devices that stopped checking in                                     | `m365 intune devices list --stale-days 30 --all`                                                                                                |
| Find a device by serial or name prefix                                    | `m365 intune devices list --search 5CG1234`                                                                                                     |
| Everything about one device (storage, encryption, Defender, primary user) | `m365 intune devices get DESKTOP-4F2K9`                                                                                                         |
| "Why is my laptop non-compliant / blocked by Conditional Access?"         | `m365 intune devices compliance DESKTOP-4F2K9`                                                                                                  |
| "Why didn't the Wi-Fi/VPN/BitLocker profile apply?"                       | `m365 intune devices config-states DESKTOP-4F2K9`                                                                                               |
| "Is Chrome installed / which version?"                                    | `m365 intune devices apps DESKTOP-4F2K9 --search chrome`                                                                                        |
| User stuck at the BitLocker recovery screen                               | `m365 intune devices bitlocker-key DESKTOP-4F2K9 --reveal`                                                                                      |
| Local admin password (Windows LAPS)                                       | `m365 intune devices laps DESKTOP-4F2K9 --reveal`                                                                                               |
| Did the wipe / sync / lock go through?                                    | `m365 intune devices actions DESKTOP-4F2K9`                                                                                                     |
| Push a policy change to a device now                                      | `m365 intune devices sync DESKTOP-4F2K9`                                                                                                        |
| Reboot a device                                                           | `m365 intune devices restart DESKTOP-4F2K9`                                                                                                     |
| Lost phone: lock it / find it                                             | `m365 intune devices lock "Jane iPhone"` / `m365 intune devices locate "Jane iPhone"`                                                           |
| Run a Defender scan / update signatures                                   | `m365 intune devices defender-scan DESKTOP-4F2K9 [--full-scan]` / `m365 intune devices defender-update DESKTOP-4F2K9`                           |
| Rotate a disclosed BitLocker key / LAPS password                          | `m365 intune devices rotate-bitlocker DESKTOP-4F2K9` / `m365 intune devices rotate-laps DESKTOP-4F2K9`                                          |
| Rename a device                                                           | `m365 intune devices rename DESKTOP-4F2K9 --name SALES-LT-042`                                                                                  |
| Collect diagnostic logs                                                   | `m365 intune devices collect-logs DESKTOP-4F2K9`                                                                                                |
| Leaver with a personal phone: remove company data only                    | `m365 intune devices retire "Jane iPhone"`                                                                                                      |
| Lost/stolen corporate laptop: factory reset                               | `m365 intune devices wipe DESKTOP-4F2K9`                                                                                                        |
| Reset a laptop but keep it enrolled / keep user files                     | `m365 intune devices wipe DESKTOP-4F2K9 --keep-enrollment-data` / `m365 intune devices fresh-start DESKTOP-4F2K9 --keep-user-data`              |
| Remove a stale device record                                              | `m365 intune devices delete OLD-LAPTOP-01`                                                                                                      |
| User locked out of their phone                                            | `m365 intune devices reset-passcode "Jane Pixel"`                                                                                               |
| Which compliance policies exist and who fails them?                       | `m365 intune compliance list`                                                                                                                   |
| Devices failing one compliance policy                                     | `m365 intune compliance status "Windows baseline" --failing-only`                                                                               |
| Which configuration profiles exist?                                       | `m365 intune config list`                                                                                                                       |
| Devices where a profile failed                                            | `m365 intune config status "Corporate Wi-Fi" --failing-only`                                                                                    |
| Find an app                                                               | `m365 intune apps list --search zoom`                                                                                                           |
| "Did the app install everywhere? Where did it fail?"                      | `m365 intune apps status "Zoom Workplace"`                                                                                                      |
| Autopilot: is this new laptop registered and profiled?                    | `m365 intune autopilot list --serial 5CG1234`                                                                                                   |
| "How many devices are non-compliant and why?"                             | `m365 intune report noncompliant`                                                                                                               |
| "Who gets Zoom?"                                                          | `m365 intune apps assignments "Zoom Workplace"`                                                                                                 |
| Deploy an app to a group / offer it in Company Portal                     | `m365 intune apps assign "Zoom Workplace" --intent required --group "Sales Laptops"` / `--intent available --all-users`                         |
| Uninstall an app from a group                                             | `m365 intune apps assign "Zoom Workplace" --intent uninstall --group "Sales Laptops"`                                                           |
| Stop deploying an app to a group                                          | `m365 intune apps unassign "Zoom Workplace" --group "Sales Laptops"`                                                                            |
| Apply a compliance policy to a group (with an exclusion)                  | `m365 intune compliance assign "Windows baseline" --group "All Windows" --exclude "Kiosks"`                                                     |
| Assign / unassign a configuration profile (classic or settings catalog)   | `m365 intune config assign "Corporate Wi-Fi" --group "Sales Laptops"` / `m365 intune config unassign "Corporate Wi-Fi" --group "Sales Laptops"` |
| "What will a device in group X get?"                                      | `m365 intune assignments "Sales Laptops"`                                                                                                       |

Destructive actions (`restart`, `lock`, `retire`, `wipe`, `fresh-start`, `delete`, `reset-passcode`) first return
a preview naming the device, user, OS, serial and last sync, plus a confirm token. Actions a platform does not
support (e.g. `lock` on Windows, `defender-scan` on iOS) are refused with a `USAGE` error before anything is sent.

`assign` / `unassign` (apps, compliance, config) are destructive: Intune's assign action replaces the whole
assignment list, so the CLI reads the current assignments, adds or removes only the requested target, and sends the
merged list. The preview shows the target group with its member count, the current assignments and the resulting
list. Targets: `--group`, `--all-users`, `--all-devices`, plus `--exclude <group>`. Available apps are user-targeted
only, so `--intent available --all-devices` is refused. Removing a required app assignment does not uninstall it.
`intune assignments <group>` scans every app and policy (up to `--max-items`, default 500, via `$batch`) and also
reports assignments to groups the given group is nested in (`via`).

## Licensing and roles

- **Intune Plan 1** (in Microsoft 365 E3/E5/Business Premium/EMS) for managed devices. Remote help, advanced
  analytics and Endpoint Privilege Management are add-ons, not covered here.
- **Delegated:** `Intune Administrator` for everything; `Global Reader` (or an Intune RBAC read-only role) for
  read commands. BitLocker keys need `Cloud Device Administrator`, `Helpdesk Administrator`, `Intune
  Administrator`, `Security Reader` or `Global Reader`; revealing LAPS passwords needs `Cloud Device Administrator`
  or `Intune Administrator`.
- **Graph permissions:** `DeviceManagementManagedDevices.Read.All` (devices), `DeviceManagementConfiguration.Read.All`
  (compliance/config), `DeviceManagementApps.Read.All` (apps), `DeviceManagementServiceConfig.Read.All` (Autopilot),
  `DeviceManagementManagedDevices.PrivilegedOperations.All` (wipe, retire, sync, lock, ...),
  `DeviceManagementManagedDevices.ReadWrite.All` (delete, locate, rotate BitLocker, collect logs),
  `BitlockerKey.Read.All` and `DeviceLocalCredential.Read.All` for the secrets. Assignments use
  `DeviceManagementApps.ReadWrite.All` (apps) or `DeviceManagementConfiguration.ReadWrite.All` (compliance/config),
  plus `GroupMember.Read.All` to resolve groups and count members. `m365 permissions` lists them per command.
- Reading a BitLocker key or LAPS password is written to the Entra audit log.
- `devices get` Defender fields need Windows devices with Microsoft Defender Antivirus; Intune shows Intune's view
  of AV health. Defender for Endpoint detail (alerts, exposure) lives in `m365 defender`.

## Beta APIs used

`devices apps` (detectedApps), the `autopilotEnrolled`/`joinType` fields of `devices get`, `devices rotate-bitlocker`,
`devices rotate-laps`, `devices rename`, `devices collect-logs`, settings catalog policies in `config list`, `config
assign/unassign` and `intune assignments` for settings catalog policies, `apps status` and `autopilot list`
(deployment profile status). Beta behaviour can change without notice.

## Filtering notes

The managedDevice docs only promise `$filter` on `complianceState`, `lastSyncDateTime` and `enrolledDateTime`.
`devices list` also sends filters on user, OS, ownership, name and serial that Intune accepts in practice; if Graph
rejects them, the CLI retries with only the documented filters and applies the rest client-side (slower on large
tenants; a note is printed to stderr). The user filter is `userPrincipalName eq`: Intune rejects `userId eq`, and
`/users/{id}/managedDevices` omits some users' devices. `apps list --search`, `autopilot list --user`,
`--failing-only` and `devices list --search` filter client-side; the last because Intune returns an empty page for
`startswith()` inside an `or`.

## Not available via Graph

| Task                                                                   | Use instead                                                                                                             |
| ---------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| Per-device status of settings catalog / endpoint security policies     | Intune admin center (Devices > Configuration > policy > Device status); `devices config-states` covers classic profiles |
| Download collected diagnostic logs                                     | Intune admin center > device > Device diagnostics                                                                       |
| Show the location returned by `devices locate`                         | Intune admin center > device > Locate device                                                                            |
| Activation Lock bypass, Lost Mode, play sound, shared iPad user logout | Intune admin center (beta-only or not exposed in v1.0; not wrapped here)                                                |
| Remote help / remote control sessions                                  | Intune Remote Help app or admin center                                                                                  |
| Run proactive remediation scripts on demand                            | Intune admin center (Devices > Scripts and remediations)                                                                |
| Import Autopilot hardware hashes from a CSV with group tag in bulk     | Intune admin center or `Get-WindowsAutopilotInfo -Online`                                                               |
| Turn BitLocker on / suspend it on a device                             | Endpoint security disk encryption policy in Intune; `manage-bde` on the device                                          |
| App protection (MAM) wipe of a single app                              | Intune admin center > Apps > App selective wipe                                                                         |
| Co-managed device workloads, Configuration Manager actions             | Configuration Manager console                                                                                           |
