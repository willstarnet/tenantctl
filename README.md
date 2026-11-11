# tenantctl

[![CI](https://github.com/willstarnet/tenantctl/actions/workflows/ci.yml/badge.svg)](https://github.com/willstarnet/tenantctl/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Microsoft 365 administration from the command line, built for AI agents and the people working with them.
`tenantctl` covers Entra ID, Intune, Defender, Exchange Online, SharePoint/OneDrive, Purview audit, Teams and
service health through Microsoft Graph.

```sh
tenantctl troubleshoot user jane@contoso.com        # why can't Jane sign in / get mail / use Teams?
tenantctl troubleshoot device LAPTOP-0042           # why is this laptop non-compliant?
tenantctl entra users signins jane --failures-only  # what failed, with each error code explained
tenantctl intune devices wipe "Bob's iPhone"        # shows what it will erase and waits for your approval
tenantctl exchange rules audit                      # find mailboxes forwarding mail outside the company
tenantctl health overview                           # is it us, or is Microsoft down?
```

- **Names, not ids.** Users by email or display name, devices by name or serial, sites by URL. When a name matches
  several things you get the candidates, never a silent guess.
- **Compact, explained output.** JSON for agents, tables for people, CSV for spreadsheets.
- **Safe by default.** Read-only until you turn writes on. Destructive changes wait for you to approve them in a
  dialog the agent can't click, and every change attempt is logged.

## Install

macOS / Linux:

```sh
curl -fsSL https://raw.githubusercontent.com/willstarnet/tenantctl/main/scripts/install.sh | sh
```

Windows (PowerShell):

```powershell
irm https://raw.githubusercontent.com/willstarnet/tenantctl/main/scripts/install.ps1 | iex
```

Or with a package manager:

```sh
brew install willstarnet/tap/tenantctl
scoop bucket add willstarnet https://github.com/willstarnet/scoop-bucket; scoop install willstarnet/tenantctl
```

Builds are published for macOS (arm64, x64), Linux (x64, arm64, glibc and musl) and Windows (x64, arm64) on the
[releases page](https://github.com/willstarnet/tenantctl/releases). The install scripts verify the download
against `SHA256SUMS` and need no admin rights. To build from source, see [CONTRIBUTING.md](CONTRIBUTING.md).

## Sign in

```sh
tenantctl auth login
```

Open the URL it prints, enter the code and sign in with your admin account. Commands then run as you, within
your admin roles. No app registration is needed: it uses Microsoft's "Microsoft Graph Command Line Tools" app,
the one Microsoft Graph PowerShell uses. Add `--include-writes` to also consent to the write permissions.

For automation, sign in as an app instead: register an app with the permissions from
`tenantctl permissions --mode application`, then

```sh
tenantctl config init --mode app --tenant contoso.onmicrosoft.com --client-id <app-id> --cert ~/secure/app.pem
```

Some features use other Microsoft APIs and need one more sign-in:

- **Exchange admin** (mailbox permissions, forwarding, distribution lists, quarantine, holds): PowerShell 7 with the
  ExchangeOnlineManagement module, then `tenantctl auth login --service exchange`.
- **Defender for Endpoint device actions** (isolate, scan): your own app registration with WindowsDefenderATP
  permissions, then `tenantctl auth login --service defender`.

Several tenants: `tenantctl --profile <name> config init --tenant <domain>`, then `tenantctl config use <name>`.
`tenantctl auth status` shows who you are signed in as, your roles and permissions, and whether writes are on.

## Making changes

| Risk        | Examples                                                        | What it takes                        |
| ----------- | --------------------------------------------------------------- | ------------------------------------ |
| read        | list, get, sign-ins, compliance, message trace, audit search    | nothing                              |
| write       | sync a device, add a group member, set an auto-reply            | writes enabled                       |
| destructive | wipe, retire, disable a user, reset a password, delete, unshare | writes enabled **and your approval** |

- Writes are off by default. Turn them on with `tenantctl config set allowWrites true` (or
  `TENANTCTL_ALLOW_WRITES=1` for one session). `--dry-run` previews any change without making it.
- A destructive command shows you the target, its current state and what will happen, in a desktop dialog or a
  terminal prompt. It runs only if you approve; otherwise it exits with code 10 and nothing changes.
- An agent can't switch this off: environment variables can only make approval stricter, and lowering it needs
  your approval too. For unattended automation, `confirmMode token` replaces the person with a two-step confirm
  token.
- Every change attempt is written to `audit.jsonl` in the config folder: `tenantctl config audit-log`.

## Using it with an AI agent

Give the agent [skills/tenantctl/SKILL.md](skills/tenantctl/SKILL.md). It covers discovering commands, the approval
rules and step-by-step playbooks. For Claude Code:

```sh
mkdir -p ~/.claude/skills/tenantctl && curl -fsSL -o ~/.claude/skills/tenantctl/SKILL.md \
  https://raw.githubusercontent.com/willstarnet/tenantctl/main/skills/tenantctl/SKILL.md
```

Agents can also discover everything themselves with `tenantctl tools --json` (every command with a JSON Schema for
its arguments) and `tenantctl <area> --help`.

## Commands

| Area           | For                                                                                                  |
| -------------- | ---------------------------------------------------------------------------------------------------- |
| `troubleshoot` | one-shot diagnostics for a user, sign-in, device, mail delivery or the whole tenant                  |
| `entra`        | users, guests, groups, licenses, sign-ins, MFA, roles, apps and consent, Conditional Access, risk    |
| `intune`       | devices and remote actions, compliance, configuration, apps, assignments, Autopilot                  |
| `defender`     | incidents, alerts, advanced hunting, Secure Score, device response                                   |
| `exchange`     | mailboxes, rules, forwarding, permissions, shared mailboxes, lists, quarantine, holds, message trace |
| `sharepoint`   | sites, files, sharing and permissions, OneDrive                                                      |
| `purview`      | unified audit log search, eDiscovery cases and holds                                                 |
| `teams`        | teams, channels, membership, presence, call quality                                                  |
| `health`       | Microsoft 365 service health and message center                                                      |
| `graph`        | raw Microsoft Graph requests                                                                         |

The full list of every command, with arguments, examples and permissions, is in
[docs/COMMANDS.md](docs/COMMANDS.md). The permissions each command needs are summarized in
[docs/PERMISSIONS.md](docs/PERMISSIONS.md).

## Output and exit codes

`--output json|table|csv` (JSON when not on a terminal), `--fields a,b.c` to trim, `--full` for raw Graph objects.

```json
{ "ok": true, "command": "intune devices list", "data": [ ... ], "meta": { "count": 50, "truncated": true } }
{ "ok": false, "error": { "code": "FORBIDDEN", "message": "...", "hint": "..." } }
```

Exit codes: 0 ok, 1 error, 2 usage, 3 confirmation required (token mode), 4 sign-in required, 5 permission denied,
6 not found, 7 writes disabled, 8 ambiguous name, 9 throttled, 10 not approved.

## Configuration

Settings live in `~/.config/tenantctl/` (`%APPDATA%\tenantctl` on Windows) and are set with
`tenantctl config set <key> <value>`. Token caches there are encrypted; the key is kept in the macOS Keychain,
libsecret on Linux or DPAPI on Windows.

| Setting                      | Environment variable                                   | Notes                                       |
| ---------------------------- | ------------------------------------------------------ | ------------------------------------------- |
| `authMode`                   | `TENANTCTL_AUTH_MODE`                                  | `delegated` (default) or `app`              |
| `tenantId` / `clientId`      | `TENANTCTL_TENANT_ID` / `TENANTCTL_CLIENT_ID`          |                                             |
| `certificatePath`            | `TENANTCTL_CERT_PATH`                                  | app-only PEM with key and certificate       |
| `allowWrites`                | `TENANTCTL_ALLOW_WRITES`                               |                                             |
| `confirmMode`                | `TENANTCTL_CONFIRM_MODE` (stricter only)               | `auto` (default), `dialog`, `tty`, `token`  |
| `graphUrl` / `authorityHost` | `TENANTCTL_GRAPH_URL` / `TENANTCTL_AUTHORITY_HOST`     | national clouds                             |
| -                            | `TENANTCTL_CLIENT_SECRET`, `TENANTCTL_CERT_PASSPHRASE` | secrets, never written to disk              |
| -                            | `TENANTCTL_ACCESS_TOKEN`                               | use a ready token (e.g. from the Azure CLI) |
| -                            | `TENANTCTL_PROFILE`, `TENANTCTL_CONFIG_DIR`            |                                             |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Report security issues privately as [SECURITY.md](SECURITY.md) describes.
Changes are listed in [CHANGELOG.md](CHANGELOG.md).

## License

[MIT](LICENSE)
