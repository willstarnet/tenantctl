# Changelog

All notable changes are listed here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed

- `intune apps status` failed for every app: Microsoft removed the `installSummary` and `deviceStatuses` endpoints.
  It now reads the Intune app install reports, the same data as the admin center, and lists every targeted device
  with its install state (failed first), detail and error code. `--failed-only` keeps failed rows; `summary` is null
  until Intune has results for a newly assigned app. The `failures` field is replaced by `devices`.

## [0.2.1] - 2026-11-16

### Added

- Windows Autopilot from the command line:
  - `intune autopilot get`: a device found by serial, name or id, with its deployment profile and Intune record
  - `intune autopilot update`: set or clear the group tag, assigned user and display name
  - `intune autopilot delete`: deregister a device; refused while the device still has an Intune record
  - `intune autopilot import` and `intune autopilot imports`: import hardware hashes from a
    `Get-WindowsAutopilotInfo` CSV and track the import
  - `intune autopilot sync` and `intune autopilot sync-status`
  - `intune autopilot profiles list|get|assign|unassign`: deployment profiles and who gets them
  - `intune autopilot events`: deployment history with failures and phase durations
  - `intune autopilot list` gained `--group-tag` and `--enrollment-state`
- `intune enrollment list` and `intune enrollment get`: enrollment status pages, platform restrictions,
  device limits and Windows Hello for Business, with priorities and assignments.
- `entra devices delete`: deletes an Entra device after approval. The preview counts the BitLocker recovery keys
  that go with it and warns when the device is still in Intune or Autopilot.
- `entra devices settings`: who can join and register devices, MFA, device quota, local administrators and the
  Entra LAPS switch.
- `intune filters list` and `intune filters get`. `get` shows a filter's rule and every app and policy that uses
  it.
- `--filter`, `--filter-mode` and `--clear-filter` on `intune apps|compliance|config assign`: add, change or
  remove an assignment filter. Assignments now show the filter by name.
- `intune config get`: one configuration profile with its settings and assignments, for reading a policy back
  after a portal change.
- `intune config export` and `intune config create --file`: copy a settings catalog or endpoint security policy
  to JSON, edit it, and create it as a new policy.
- `intune compliance settings`: how many devices pass or fail each setting of a compliance policy.
- `intune tenant`: Apple push certificate, ADE and VPP tokens with days left, Managed Google Play, Mobile Threat
  Defense connectors and compliance settings, with warnings for anything expiring or off.
- `intune cleanup-rules list` and `intune cleanup-rules create` (the preview counts the devices it would hide).
- `intune rbac`: Intune roles with their assignments, scope tags and Multi Admin Approval policies.
- `intune scripts list` and `intune scripts status`: platform scripts, Remediations and macOS shell scripts,
  with per-device failures.
- `intune updates list` and `intune updates get`: update rings and feature, quality and driver update profiles,
  with the ones Windows Autopatch manages marked.
- `entra ca set-state`: turn a Conditional Access policy on, off or to report-only.
- `entra ca locations`: named locations with their ranges and the policies that use them; `--ip` checks
  whether an address is in one.
- `entra policies`: security defaults, user consent, guest invites, self-service sign-up, the admin consent
  workflow and authentication methods, with warnings where they differ from Microsoft's recommendations.

### Fixed

- `permissions` example and the `config init --mode app` hint suggested `--mode app`; the option is
  `--mode application`.
- Choice options accept any letter case: `intune devices list --os Windows` no longer fails.
- `intune apps|compliance|config assign` and `unassign` removed the assignment filter from every assignment
  they re-sent. They now read and write assignments through the beta API, which keeps the filters.
- `intune devices config-states` and `troubleshoot device` no longer hit Intune throttling, and now include
  settings catalog, administrative template and endpoint security policies. They read the per-device report
  the Intune admin center uses.
- `autopilot import` reads the UTF-16 CSV that Get-WindowsAutopilotInfo writes on Windows PowerShell 5.1.
- `entra devices delete` now detects Autopilot registrations, and points to `intune devices delete` then
  `intune autopilot delete` as Microsoft's order.
- Assign and unassign no longer send back the read-only `source` fields, refuse items whose assignments come
  from a policy set, and treat a configuration assignment in `remove` state as an exclusion.

### Security

- `next` commands and hints shell-quote tenant data, so a device, group or app name containing `$(...)` or
  backticks cannot run anything when an agent runs the suggestion.
- Destructive commands act only on the object the person approved: if the name resolves to another object by
  the time the change runs (or a confirm token is reused after that), nothing changes.
- `intune autopilot update` and `intune autopilot import` need approval like other destructive changes; an
  import is refused if the CSV changed after its preview. Changes need the whole serial, not part of it.
- `intune config export --out` never overwrites a file without `--force`, never writes through a symlink or into
  the tenantctl config folder, and creates the file readable only by you.
- Approval dialogs show every warning in full before the details are cut short.

### Changed

- Assignment output names the filter (`filter`) and calls its mode `filterMode` (was `filterType`).
- List results report the collection size as `meta.total` when Graph returns it; a truncated list says
  "50 of 62 shown" instead of only "more results exist". `entra groups members` always asks for the count.

### Removed

- The Scoop bucket. Windows installs use the PowerShell install script.

## [0.2.0] - 2026-11-10

### Changed

- Renamed from `m365` to `tenantctl`: the command, `TENANTCTL_*` environment variables (were `M365_*`), the
  config folder and the skill (`skills/tenantctl`). The first run moves an existing `m365` config folder and
  token-cache key, so profiles and sign-ins carry over.
- Destructive commands ask for approval in the same run: the dialog or terminal prompt shows the preview and the
  change runs once you approve. The two-step `--confirm` token remains only for `confirmMode token`.
- Windows reads the token-cache key through DPAPI in-process instead of starting PowerShell on every command.
- `entra groups get` and `intune assignments` return group member counts again.

### Added

- Prebuilt binaries for macOS (arm64, x64), Linux (x64, arm64, musl, baseline x64) and Windows (x64, arm64,
  baseline x64), install scripts for Linux and Windows, and a Homebrew tap.
- The Exchange "PowerShell 7 missing" hint names the install command for your OS.

## 0.1.0 - 2026-11-10 (not released)

### Added

- First version as `m365`: Entra ID, Intune, Defender, Exchange Online, SharePoint/OneDrive, Purview audit, Teams
  and service health commands, `troubleshoot` diagnostics, write gating with human approval, and an audit log.

[Unreleased]: https://github.com/willstarnet/tenantctl/compare/v0.2.1...HEAD
[0.2.1]: https://github.com/willstarnet/tenantctl/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/willstarnet/tenantctl/releases/tag/v0.2.0
