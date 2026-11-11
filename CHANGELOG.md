# Changelog

All notable changes are listed here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
  baseline x64), install scripts for macOS/Linux and Windows, a Homebrew tap and a Scoop bucket.
- The Exchange "PowerShell 7 missing" hint names the install command for your OS.

## 0.1.0 - 2026-11-10 (not released)

### Added

- First version as `m365`: Entra ID, Intune, Defender, Exchange Online, SharePoint/OneDrive, Purview audit, Teams
  and service health commands, `troubleshoot` diagnostics, write gating with human approval, and an audit log.

[Unreleased]: https://github.com/willstarnet/tenantctl/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/willstarnet/tenantctl/releases/tag/v0.2.0
