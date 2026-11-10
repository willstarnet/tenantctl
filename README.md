# m365 - Microsoft 365 administration for AI agents

`m365` is a command-line tool that lets an AI assistant (Claude, Copilot, omp, or any agent that can run shell
commands) manage and troubleshoot a Microsoft 365 tenant through Microsoft Graph. People can use it too.

It covers **Entra ID, Intune, Defender, Exchange Online, SharePoint/OneDrive, Teams** and **service health**,
and it is built for the questions admins actually ask:

```sh
m365 troubleshoot user jane@contoso.com        # why can't Jane sign in / get mail / use Teams?
m365 troubleshoot device LAPTOP-0042           # why is this laptop non-compliant?
m365 entra users signins jane --failures-only  # what failed, with each error code explained
m365 intune devices wipe "Bob's iPhone"        # previews first, then needs a confirm token
m365 exchange rules audit                      # find mailboxes forwarding mail outside the company
m365 sharepoint files permissions <file-url>   # who can see this file?
m365 health overview                           # is it us, or is Microsoft down?
```

What makes it agent-friendly:

- **Names, not ids.** Users by email, UPN or display name; devices by name or serial; sites by URL; teams by
  name. Several matches return an `AMBIGUOUS` error listing the candidates, so the agent can ask which one.
- **Compact, explained output.** Each command returns the fields that answer the question plus derived facts
  (`daysSinceLastSync`, `forwardsExternally`, the meaning of sign-in error 53000). `--full` gives raw Graph.
- **Machine-readable everything.** JSON output when not on a terminal, a stable `{ ok, data, meta }`
  envelope, distinct exit codes, and `m365 tools` prints every command with a JSON Schema for its arguments.
- **Safe by default.** Read-only until writes are enabled. Destructive actions need a confirm token from a
  preview. Every change attempt is logged locally.

## Install

You need [Bun](https://bun.sh) only to build. The binary it produces runs on its own.

```sh
git clone <this repo> ~/git/M365Management && cd ~/git/M365Management
bun install
bun run install:local      # builds for this machine, installs to ~/.local/bin/m365
# If `m365` is "command not found", ~/.local/bin is not on PATH; the installer prints the line to add, e.g. zsh:
#   echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.zshrc && source ~/.zshrc
m365 --version
```

`bun run build` cross-compiles for macOS, Linux and Windows (x64/arm64) into `dist/`. For development, `bun link`
makes `m365` run straight from source.

## Sign in

### Option 1: as yourself (delegated, quickest)

```sh
m365 auth login
```

It prints a code and a URL. Open the URL, enter the code, and sign in with your admin account. There's no app
registration to create: it uses Microsoft's own "Microsoft Graph Command Line Tools" app, the same one Microsoft
Graph PowerShell uses. Every action then runs as you, within your admin roles and Conditional Access policies,
and shows up under your name in Entra audit logs.

The first sign-in asks for consent to the permissions the CLI's read commands use. Add `--include-writes` to
also request the write permissions. A Global Administrator can consent for the whole organization. If your
tenant blocks that app, register your own public client app and run
`m365 config init --client-id <app-id> --tenant <tenant>`.

### Option 2: as an app (app-only, for automation)

1. In Entra ID, go to **App registrations > New registration**.
2. Under **Certificates & secrets**, upload a certificate. Use a client secret only if you have to.
3. Under **API permissions**, add the **Application** permissions from `m365 permissions --mode application`
   (add `--include-writes` if this app may make changes), then click **Grant admin consent**.
4. Configure the profile:

   ```sh
   m365 config init --mode app --tenant contoso.onmicrosoft.com --client-id <app-id> --cert ~/secure/m365.pem
   m365 auth login      # checks the credentials
   ```

The PEM file must contain the private key. Include the certificate in it as well, so the thumbprint can be
worked out for you. To use a client secret instead, export `M365_CLIENT_SECRET`. Secrets are never written to
the config file.

### Exchange admin and Defender for Endpoint

Some admin features have no Graph API. Two kinds of commands use other Microsoft APIs:

- **Exchange admin commands** run through Exchange Online PowerShell, for example mailbox permissions, shared
  mailboxes, mailbox forwarding, distribution lists, quarantine and holds. They need PowerShell 7 (`brew install
  powershell`) and the module (`pwsh -c "Install-Module ExchangeOnlineManagement -Scope CurrentUser"`). Exchange
  needs its own sign-in:

  ```sh
  m365 auth login --service exchange
  ```

  For app-only profiles, give the app the **Exchange.ManageAsApp** application permission (Office 365 Exchange
  Online) and assign it the Exchange Administrator role.

- **Defender for Endpoint device actions** (isolate, antivirus scan, investigation package) need your own app
  registration with WindowsDefenderATP permissions, either app-only or as a public client. Microsoft's built-in
  client can't be granted them. For a public client, set it with `m365 config set clientId <app-id>`, then run
  `m365 auth login --service defender`.

### Several tenants

```sh
m365 --profile fabrikam config init --tenant fabrikam.onmicrosoft.com
m365 --profile fabrikam auth login
m365 config use fabrikam        # make it the default
```

`m365 auth status` shows who you are signed in as, your admin roles, the permissions you were granted, whether
writes are on, how destructive changes are approved, and where the token-cache key is kept.

## Making changes safely

| Risk        | Examples                                                                                | What it takes                                          |
| ----------- | --------------------------------------------------------------------------------------- | ------------------------------------------------------ |
| read        | list, get, sign-ins, compliance, trace, audit search                                    | nothing                                                |
| write       | sync a device, add a group member, set an auto-reply                                    | writes enabled                                         |
| destructive | wipe, retire, disable a user, reset a password, revoke sessions, delete, remove sharing | writes enabled, a confirm token, **and your approval** |

- **Writes are off by default.** Turn them on per session with `export M365_ALLOW_WRITES=1`, or for a profile
  with `m365 config set allowWrites true`.
- **`--dry-run`** shows what any write would do, even while writes are off.
- **Destructive commands need you, not just the agent.**
  1. The first run changes nothing. It resolves the target, shows its current state, and exits with code 3,
     returning a confirm token. The token is tied to that exact command, its arguments, the profile and the
     tenant, and it expires after 10 minutes.
  2. When the agent re-runs with `--confirm <token>`, `m365` asks **you** directly. It shows a macOS, Windows
     or Linux desktop dialog, or a prompt on your terminal where you type `APPROVE`. The agent's shell can't
     answer either. If you deny, it exits with code 10 and nothing changes.
- **Choose how you approve:**

  | Mode             | How you approve                                                         |
  | ---------------- | ----------------------------------------------------------------------- |
  | `auto` (default) | a dialog when you have a desktop session, otherwise the terminal prompt |
  | `dialog`         | always a desktop dialog                                                 |
  | `tty`            | always the terminal prompt                                              |
  | `token`          | the confirm token alone (only for unattended automation)                |

  Set it with `m365 config set confirmMode <mode>`.

- **Agents can't switch this off.** `M365_CONFIRM_MODE` in the environment can only make approval stricter.
  Lowering the mode, or turning on `allowWrites` while approval is on, needs your approval too.
- **What this protects against:** a confused or prompt-injected agent. It is not a sandbox against malware
  running under your account.
- **Audit log.** Every write attempt is appended to `~/.config/m365/audit.jsonl`: the time, who made it, the
  command, the arguments (with secrets redacted) and the outcome (`success`, `error`, `confirmation_issued`,
  `denied`). View it with `m365 config audit-log`.

## Using it with an AI agent

Give the agent [`skills/m365/SKILL.md`](skills/m365/SKILL.md). It is written as an agent skill (Claude Code /
omp / Copilot instructions) and covers how to discover commands, the confirm protocol, and step-by-step
playbooks for common requests. For Claude Code:

```sh
mkdir -p ~/.claude/skills && ln -s ~/git/M365Management/skills/m365 ~/.claude/skills/m365
```

The agent can discover commands itself:

```sh
m365 tools --json                # every command, args as JSON Schema, risk, permissions
m365 intune --help               # one area
m365 intune devices wipe --help  # one command
```

## Output and exit codes

`--output json|table|csv`. The default is a table on a terminal and JSON otherwise. `--csv` exports lists as a
spreadsheet, for example `m365 intune devices list --compliance noncompliant --all --csv > devices.csv`.
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

| Setting                      | Env override                                               | Notes                                                                          |
| ---------------------------- | ---------------------------------------------------------- | ------------------------------------------------------------------------------ |
| `authMode`                   | `M365_AUTH_MODE`                                           | `delegated` (default) or `app`                                                 |
| `tenantId`                   | `M365_TENANT_ID`                                           | tenant id or domain                                                            |
| `clientId`                   | `M365_CLIENT_ID`                                           | default for delegated: Microsoft Graph Command Line Tools                      |
| `certificatePath`            | `M365_CERT_PATH`                                           | app-only PEM (key + certificate); passphrase via `M365_CERT_PASSPHRASE`        |
| `certificateThumbprint`      | `M365_CERT_THUMBPRINT`                                     | only if the PEM has no certificate                                             |
| `allowWrites`                | `M365_ALLOW_WRITES`                                        | `true`/`1` to enable writes                                                    |
| `confirmMode`                | `M365_CONFIRM_MODE` (stricter only)                        | `auto` (default), `dialog`, `tty`, `token`                                     |
| `graphUrl` / `authorityHost` | `M365_GRAPH_URL` / `M365_AUTHORITY_HOST`                   | national clouds                                                                |
| -                            | `M365_CLIENT_SECRET`                                       | app-only secret (env only)                                                     |
| -                            | `M365_ACCESS_TOKEN`                                        | use a ready token, e.g. `az account get-access-token --resource-type ms-graph` |
| -                            | `M365_EXCHANGE_ACCESS_TOKEN`, `M365_DEFENDER_ACCESS_TOKEN` | ready tokens for Exchange Online / Defender for Endpoint                       |
| -                            | `M365_PWSH`                                                | path to `pwsh` if it isn't on PATH                                             |
| -                            | `M365_TOKEN_CACHE_KEYSTORE=file`                           | keep the cache key in a file (headless Linux, containers)                      |
| -                            | `M365_PROFILE`, `M365_CONFIG_DIR`, `M365_OUTPUT`           | profile, config location, default output                                       |

Config lives in `~/.config/m365/` (`%APPDATA%\m365` on Windows). That folder holds `config.json`, the token
caches (`msal-cache-<profile>[-exchange].json`), the confirm-token key and the audit log, each readable only by
your user account. The token caches are **encrypted** (AES-256-GCM). The key is kept in the macOS Keychain,
libsecret on Linux, or protected with DPAPI on Windows; only when none of those is available does it fall back to
a `cache.key` file, and `m365 auth status` shows which.

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
