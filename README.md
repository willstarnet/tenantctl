# tenantctl - Microsoft 365 administration for AI agents

[![CI](https://github.com/willstarnet/tenantctl/actions/workflows/ci.yml/badge.svg)](https://github.com/willstarnet/tenantctl/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

`tenantctl` is a command-line tool that lets an AI assistant (Claude, Copilot, omp, or any agent that can run shell
commands) manage and troubleshoot a Microsoft 365 tenant through Microsoft Graph. People can use it too.

It covers **Entra ID, Intune, Defender, Exchange Online, SharePoint/OneDrive, Teams** and **service health**,
and it is built for the questions admins actually ask:

```sh
tenantctl troubleshoot user jane@contoso.com        # why can't Jane sign in / get mail / use Teams?
tenantctl troubleshoot device LAPTOP-0042           # why is this laptop non-compliant?
tenantctl entra users signins jane --failures-only  # what failed, with each error code explained
tenantctl intune devices wipe "Bob's iPhone"        # shows you what it will erase and waits for your approval
tenantctl exchange rules audit                      # find mailboxes forwarding mail outside the company
tenantctl sharepoint files permissions <file-url>   # who can see this file?
tenantctl health overview                           # is it us, or is Microsoft down?
```

What makes it agent-friendly:

- **Names, not ids.** Users by email, UPN or display name; devices by name or serial; sites by URL; teams by
  name. Several matches return an `AMBIGUOUS` error listing the candidates, so the agent can ask which one.
- **Compact, explained output.** Each command returns the fields that answer the question plus derived facts
  (`daysSinceLastSync`, `forwardsExternally`, the meaning of sign-in error 53000). `--full` gives raw Graph.
- **Machine-readable everything.** JSON output when not on a terminal, a stable `{ ok, data, meta }`
  envelope, distinct exit codes, and `tenantctl tools` prints every command with a JSON Schema for its arguments.
- **Safe by default.** Read-only until writes are enabled. Destructive actions wait for you to approve them in a
  dialog the agent can't click. Every change attempt is logged locally.

## Install

Prebuilt binaries for macOS (arm64/x64), Linux (x64/arm64, glibc and musl/Alpine) and Windows (x64/arm64) are
attached to each [GitHub release](https://github.com/willstarnet/tenantctl/releases). The install scripts pick the
right one, check it against `SHA256SUMS`, and put it in your user's folders (no admin rights needed).

macOS / Linux, into `~/.local/bin`:

```sh
curl -fsSL https://raw.githubusercontent.com/willstarnet/tenantctl/main/scripts/install.sh | sh
# a specific version: curl -fsSL …/install.sh | TENANTCTL_VERSION=0.2.0 sh
```

Windows (PowerShell), into `%LOCALAPPDATA%\tenantctl` and onto your Path:

```powershell
irm https://raw.githubusercontent.com/willstarnet/tenantctl/main/scripts/install.ps1 | iex
```

With a package manager: `brew install willstarnet/tap/tenantctl`, or
`scoop bucket add willstarnet https://github.com/willstarnet/scoop-bucket && scoop install willstarnet/tenantctl`.

### From source

You need [Bun](https://bun.sh) only to build. The binary it produces runs on its own.

```sh
git clone https://github.com/willstarnet/tenantctl.git ~/git/tenantctl && cd ~/git/tenantctl
bun install
bun run install:local      # builds for this machine, installs to ~/.local/bin/tenantctl
# If `tenantctl` is "command not found", ~/.local/bin is not on PATH; the installer prints the line to add, e.g. zsh:
#   echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.zshrc && source ~/.zshrc
tenantctl --version
```

`bun run build` cross-compiles every target into `dist/`; `bun run build --archive` also writes the release archives
and `SHA256SUMS`. Pushing a `v*` tag runs the same build in GitHub Actions and publishes the release. For
development, `bun link` makes `tenantctl` run straight from source.

Upgrading from `m365` (this tool's old name): the first `tenantctl` run moves `~/.config/m365` (Windows:
`%APPDATA%\m365`) to `tenantctl` and moves the token-cache key in the Keychain/libsecret, so profiles and sign-ins
carry over. Environment variables are now `TENANTCTL_*` (were `M365_*`).

## Sign in

### Option 1: as yourself (delegated, quickest)

```sh
tenantctl auth login
```

It prints a code and a URL. Open the URL, enter the code, and sign in with your admin account. There's no app
registration to create: it uses Microsoft's own "Microsoft Graph Command Line Tools" app, the same one Microsoft
Graph PowerShell uses. Every action then runs as you, within your admin roles and Conditional Access policies,
and shows up under your name in Entra audit logs.

The first sign-in asks for consent to the permissions the CLI's read commands use. Add `--include-writes` to
also request the write permissions. A Global Administrator can consent for the whole organization. If your
tenant blocks that app, register your own public client app and run
`tenantctl config init --client-id <app-id> --tenant <tenant>`.

### Option 2: as an app (app-only, for automation)

1. In Entra ID, go to **App registrations > New registration**.
2. Under **Certificates & secrets**, upload a certificate. Use a client secret only if you have to.
3. Under **API permissions**, add the **Application** permissions from `tenantctl permissions --mode application`
   (add `--include-writes` if this app may make changes), then click **Grant admin consent**.
4. Configure the profile:

   ```sh
   tenantctl config init --mode app --tenant contoso.onmicrosoft.com --client-id <app-id> --cert ~/secure/tenantctl.pem
   tenantctl auth login      # checks the credentials
   ```

The PEM file must contain the private key. Include the certificate in it as well, so the thumbprint can be
worked out for you. To use a client secret instead, export `TENANTCTL_CLIENT_SECRET`. Secrets are never written to
the config file.

### Exchange admin and Defender for Endpoint

Some admin features have no Graph API. Two kinds of commands use other Microsoft APIs:

- **Exchange admin commands** run through Exchange Online PowerShell, for example mailbox permissions, shared
  mailboxes, mailbox forwarding, distribution lists, quarantine and holds. They need PowerShell 7 (macOS: `brew install
  powershell`; Windows: `winget install Microsoft.PowerShell`; Linux: Microsoft's package repo) and the module (`pwsh -c "Install-Module ExchangeOnlineManagement -Scope CurrentUser"`). Exchange
  needs its own sign-in:

  ```sh
  tenantctl auth login --service exchange
  ```

  For app-only profiles, give the app the **Exchange.ManageAsApp** application permission (Office 365 Exchange
  Online) and assign it the Exchange Administrator role.

- **Defender for Endpoint device actions** (isolate, antivirus scan, investigation package) need your own app
  registration with WindowsDefenderATP permissions, either app-only or as a public client. Microsoft's built-in
  client can't be granted them. For a public client, set it with `tenantctl config set clientId <app-id>`, then run
  `tenantctl auth login --service defender`.

### Several tenants

```sh
tenantctl --profile fabrikam config init --tenant fabrikam.onmicrosoft.com
tenantctl --profile fabrikam auth login
tenantctl config use fabrikam        # make it the default
```

`tenantctl auth status` shows who you are signed in as, your admin roles, the permissions you were granted, whether
writes are on, how destructive changes are approved, and where the token-cache key is kept.

## Making changes safely

| Risk        | Examples                                                                                | What it takes                        |
| ----------- | --------------------------------------------------------------------------------------- | ------------------------------------ |
| read        | list, get, sign-ins, compliance, trace, audit search                                    | nothing                              |
| write       | sync a device, add a group member, set an auto-reply                                    | writes enabled                       |
| destructive | wipe, retire, disable a user, reset a password, revoke sessions, delete, remove sharing | writes enabled **and your approval** |

- **Writes are off by default.** Turn them on per session with `export TENANTCTL_ALLOW_WRITES=1`, or for a profile
  with `tenantctl config set allowWrites true`.
- **`--dry-run`** shows what any write would do, even while writes are off.
- **Destructive commands need you, not just the agent.** The command resolves the target and shows you what it
  will do and to what (current state included) in a macOS, Windows or Linux desktop dialog, or a prompt on your
  terminal where you type `APPROVE`. The agent's shell can't answer either. It runs only if you approve; if you
  deny or don't answer within 5 minutes, it exits with code 10 and nothing changes. Use `--dry-run` to see the
  preview without being asked.
- **Choose how you approve:**

  | Mode             | How you approve                                                                                                                                                                                       |
  | ---------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
  | `auto` (default) | a dialog when you have a desktop session, otherwise the terminal prompt                                                                                                                               |
  | `dialog`         | always a desktop dialog                                                                                                                                                                               |
  | `tty`            | always the terminal prompt                                                                                                                                                                            |
  | `token`          | no person: two steps, the first run prints a preview and a confirm token tied to that exact command, profile and tenant (10 minutes); re-run with `--confirm <token>`. Only for unattended automation |

  Set it with `tenantctl config set confirmMode <mode>`.

- **Agents can't switch this off.** `TENANTCTL_CONFIRM_MODE` in the environment can only make approval stricter.
  Lowering the mode, or turning on `allowWrites` while approval is on, needs your approval too.
- **What this protects against:** a confused or prompt-injected agent. It is not a sandbox against malware
  running under your account.
- **Audit log.** Every write attempt is appended to `~/.config/tenantctl/audit.jsonl`: the time, who made it, the
  command, the arguments (with secrets redacted) and the outcome (`success`, `error`, `confirmation_issued`,
  `denied`). View it with `tenantctl config audit-log`.

## Using it with an AI agent

Give the agent [`skills/tenantctl/SKILL.md`](skills/tenantctl/SKILL.md). It is written as an agent skill (Claude Code /
omp / Copilot instructions) and covers how to discover commands, the approval protocol, and step-by-step
playbooks for common requests. For Claude Code:

```sh
mkdir -p ~/.claude/skills && ln -s ~/git/tenantctl/skills/tenantctl ~/.claude/skills/tenantctl
```

The agent can discover commands itself:

```sh
tenantctl tools --json                # every command, args as JSON Schema, risk, permissions
tenantctl intune --help               # one area
tenantctl intune devices wipe --help  # one command
```

## Output and exit codes

`--output json|table|csv`. The default is a table on a terminal and JSON otherwise. `--csv` exports lists as a
spreadsheet, for example `tenantctl intune devices list --compliance noncompliant --all --csv > devices.csv`.
`--fields a,b.c` trims the output and sets the CSV column order. `--full` returns the raw Graph objects.

```json
{ "ok": true, "command": "intune devices list", "data": [ ... ], "meta": { "count": 50, "truncated": true, "hint": "..." } }
{ "ok": false, "error": { "code": "FORBIDDEN", "message": "...", "hint": "The app registration needs ..." } }
```

| Exit | Meaning                                                        |
| ---- | -------------------------------------------------------------- |
| 0    | ok                                                             |
| 1    | error (Graph or internal)                                      |
| 2    | usage error                                                    |
| 3    | confirmation required (destructive preview)                    |
| 4    | sign-in required or failed                                     |
| 5    | permission denied (the hint names the missing permission)      |
| 6    | not found                                                      |
| 7    | writes disabled                                                |
| 8    | ambiguous name (`details` lists the candidates)                |
| 9    | throttled                                                      |
| 10   | not approved by a person (denied, timed out, or no way to ask) |

## Configuration reference

| Setting                      | Env override                                                         | Notes                                                                          |
| ---------------------------- | -------------------------------------------------------------------- | ------------------------------------------------------------------------------ |
| `authMode`                   | `TENANTCTL_AUTH_MODE`                                                | `delegated` (default) or `app`                                                 |
| `tenantId`                   | `TENANTCTL_TENANT_ID`                                                | tenant id or domain                                                            |
| `clientId`                   | `TENANTCTL_CLIENT_ID`                                                | default for delegated: Microsoft Graph Command Line Tools                      |
| `certificatePath`            | `TENANTCTL_CERT_PATH`                                                | app-only PEM (key + certificate); passphrase via `TENANTCTL_CERT_PASSPHRASE`   |
| `certificateThumbprint`      | `TENANTCTL_CERT_THUMBPRINT`                                          | only if the PEM has no certificate                                             |
| `allowWrites`                | `TENANTCTL_ALLOW_WRITES`                                             | `true`/`1` to enable writes                                                    |
| `confirmMode`                | `TENANTCTL_CONFIRM_MODE` (stricter only)                             | `auto` (default), `dialog`, `tty`, `token`                                     |
| `graphUrl` / `authorityHost` | `TENANTCTL_GRAPH_URL` / `TENANTCTL_AUTHORITY_HOST`                   | national clouds                                                                |
| -                            | `TENANTCTL_CLIENT_SECRET`                                            | app-only secret (env only)                                                     |
| -                            | `TENANTCTL_ACCESS_TOKEN`                                             | use a ready token, e.g. `az account get-access-token --resource-type ms-graph` |
| -                            | `TENANTCTL_EXCHANGE_ACCESS_TOKEN`, `TENANTCTL_DEFENDER_ACCESS_TOKEN` | ready tokens for Exchange Online / Defender for Endpoint                       |
| -                            | `TENANTCTL_PWSH`                                                     | path to `pwsh` if it isn't on PATH                                             |
| -                            | `TENANTCTL_TOKEN_CACHE_KEYSTORE=file`                                | keep the cache key in a file (headless Linux, containers)                      |
| -                            | `TENANTCTL_PROFILE`, `TENANTCTL_CONFIG_DIR`, `TENANTCTL_OUTPUT`      | profile, config location, default output                                       |

Config lives in `~/.config/tenantctl/` (`%APPDATA%\tenantctl` on Windows). That folder holds `config.json`, the token
caches (`msal-cache-<profile>[-exchange].json`), the confirm-token key and the audit log, each readable only by
your user account. The token caches are **encrypted** (AES-256-GCM). The key is kept in the macOS Keychain,
libsecret on Linux, or protected with DPAPI on Windows; only when none of those is available does it fall back to
a `cache.key` file, and `tenantctl auth status` shows which.

## Checking it against your tenant

`bun run smoke:live` runs every read-only command against your signed-in tenant with real sample targets. It also
exercises the query filters Microsoft doesn't document. It reports each command as `ok`, `unavailable` (licence
or permission), `broken` (a bug in this tool) or `skipped` (no sample data), and it never runs a write.

## Reference

- [docs/COMMANDS.md](docs/COMMANDS.md): every command (generated)
- [docs/PERMISSIONS.md](docs/PERMISSIONS.md): Graph permissions per command (generated)
- `docs/domains/*.md`: per area, a map from tasks to commands, licensing prerequisites, and what Graph can't do
  (and which tool can)
- [AGENTS.md](AGENTS.md): how to add commands
- [CONTRIBUTING.md](CONTRIBUTING.md), [SECURITY.md](SECURITY.md), [SUPPORT.md](SUPPORT.md), [CHANGELOG.md](CHANGELOG.md)

## License

[MIT](LICENSE).
