# M365Management - contributor guide

`m365` is a Bun + TypeScript CLI that lets an AI agent (or a person) administer a Microsoft 365 tenant through
Microsoft Graph: Entra ID, Intune, Defender, Exchange Online, SharePoint/OneDrive, Teams and service health.

The point of the project is to make a conversation like "Jane can't sign in" or "wipe the laptop Bob lost" turn
into one or two obvious commands with compact, explainable output. Every design choice serves that.

## Commands

| Task            | Command                                                                          |
| --------------- | -------------------------------------------------------------------------------- |
| Install deps    | `bun install`                                                                    |
| Run from source | `bun src/cli.ts <area> <command> ...`                                            |
| Typecheck       | `bun run typecheck`                                                              |
| Tests           | `bun test` (one file: `bun test test/domains/intune.test.ts`)                    |
| Regenerate docs | `bun run docs` (writes `docs/COMMANDS.md`, `docs/PERMISSIONS.md`)                |
| Build binaries  | `bun run build` (all platforms into `dist/`)                                     |
| Install locally | `bun run install:local` (builds for this machine, copies to `~/.local/bin/m365`) |
| Format          | `bun run format`                                                                 |

## Layout

```
src/cli.ts                 argv -> global flags -> command lookup -> execute -> render
src/registry.ts            every command, area summaries, lookup
src/core/types.ts          defineCommand, ArgSpec, Context (graph, defender, exo, requireApproval), ListResult, Risk
src/core/graph.ts          GraphClient: retries, paging, errors, CSV reports, batch() ($batch, 20/call, per-item
                           errors), requestWithHeaders (202 Location), unversioned mode; seg(), odata()
src/core/exo.ts            Exchange Online PowerShell bridge: exoRun(ctx.exo, steps) / ctx.exo.run (fixed script,
                           params as JSON env, -AccessToken connect)
src/core/resolve.ts        resolveUser / resolveGroup / pickOne (names -> ids, AMBIGUOUS errors)
src/core/execute.ts        write gating, confirm tokens, human approval, audit, permission hints; builds Context
src/core/approval.ts       confirm modes (token/tty/dialog/auto), env-can-only-strengthen, OS dialog + /dev/tty
src/core/auth.ts           MSAL per resource (graph, exchange, defender): device code, client credentials, static
src/core/secretStore.ts    AES-256-GCM token cache; key in Keychain / libsecret / DPAPI / file fallback
src/core/config.ts         profiles (~/.config/m365/config.json), env overrides
src/core/csvInput.ts       readCsvRows() for bulk commands (500-row cap, tolerant column names)
src/core/listing.ts        LIST_ARGS (--top/--all) and listLimit()
src/core/util.ts           settle() + isSettledError(), daysAgo(), daysSince(), compact(), isGuid()
src/core/signinErrors.ts   AADSTS / sign-in error code catalog
src/domains/<area>/        commands for one area; index.ts exports `commands` and reusable helpers
src/domains/troubleshoot/  composite diagnostics; findings.ts (Finding, verdict, permissionsOf), signins.ts
test/helpers/graphMock.ts  mockContext(), runCommand(), fakeExo(), page(), graphError()
scripts/smoke-live.ts      read-only run of every read command against a real tenant (`bun run smoke:live`)
skills/m365/SKILL.md       how an AI agent should use the CLI
docs/domains/<area>.md     task -> command mapping, licensing needs, Graph gaps
```

## Writing a command

```ts
export const devicesSync = defineCommand({
  domain: 'intune',
  name: 'devices sync', // `m365 intune devices sync`
  summary: 'Ask a device to check in with Intune now',
  description: 'When to use it and how to read the result, for agents.',
  examples: ['m365 intune devices sync DESKTOP-123'],
  args: [
    { name: 'device', type: 'string', required: true, positional: true, description: 'Device name, serial or id.' },
  ],
  risk: 'write',
  permissions: {
    delegated: ['DeviceManagementManagedDevices.PrivilegedOperations.All'],
    application: ['DeviceManagementManagedDevices.PrivilegedOperations.All'],
    roles: ['Intune Administrator'],
  },
  async preview(ctx, args) {
    /* resolve the target, return what will happen */
  },
  async run(ctx, args) {
    /* do it, return a compact result */
  },
});
```

Rules:

1. **Accept what a person would say.** Users by UPN, email, display name or id (`resolveUser`); groups by name,
   mail or id (`resolveGroup`); devices by name, serial or id; sites by URL; teams by name. Use `pickOne` so
   several matches become an `AMBIGUOUS` error listing candidates, never a silent first pick.
2. **Compact output by default.** Return the fields an admin needs to answer the question, with friendly names
   and derived facts (`daysSinceLastSync`, `isMfaRegistered`, `forwardsExternally`). Return the raw Graph object
   when `ctx.full` is true. Never dump `@odata.*` noise.
3. **Lists return `ListResult`** (`{ items, truncated }`) via `ctx.graph.list(path, { limit: listLimit(args) })`
   and spread `...LIST_ARGS` into `args`. Send `$select` so payloads stay small.
4. **Risk is honest.** `read` changes nothing. `write` is a low-impact or easily reversed change (sync, add a
   member, enable). `destructive` is disruptive, security-sensitive or irreversible (wipe, retire, delete,
   disable, password reset, revoke sessions, remove access, grant mailbox/file access, reboot) and must implement
   `preview`, which resolves the target, shows its current state, and carries a plain-language `action` string:
   the human approval dialog shows it. The executor handles gating, confirm tokens, human approval and audit;
   commands never check `allowWrites` themselves.
5. **Permissions are least privilege** and taken from the "Permissions" table of the endpoint's Microsoft Graph
   doc page, for both delegated and application. Add `roles` when the delegated caller also needs an admin role.
   Endpoints that do not support one type leave it out. These lists generate `m365 permissions` and the login
   scope set, so they must be accurate.
6. **Verify every endpoint against current docs** (`https://learn.microsoft.com/graph/api/<page>` or the raw
   Markdown in `microsoftgraph/microsoft-graph-docs-contrib/api-reference/{v1.0,beta}/api/`). Prefer v1.0; use
   beta only when v1.0 lacks the feature, and say so in the description.
7. **Encode user input**: `seg()` for path segments, `odata()` for filter literals. Never interpolate raw input.
8. **Composite views use `settle()`** so one forbidden section degrades to `{ error }` instead of failing; narrow
   with `isSettledError()`. Pass `settle()` named async loaders, not inline `.then(...)` chains: the
   `Promise<unknown>` constraint widens a `.then` result to `unknown`. Troubleshoot views return `verdict`,
   ranked `findings` and `unavailable`; a finding's `next` is a runnable command with no prose, because agents
   run it verbatim. They declare permissions with `permissionsOf()` over the commands they compose.
9. **Arg names** are camelCase (`--kebab-case` on the CLI) and must not be `output`, `fields`, `full`, `profile`,
   `confirm`, `dryRun`, `debug`, `help` or `version` (the global parser would swallow them;
   `test/registry.test.ts` enforces this). Mark passwords and other secrets `secret: true`. A `string[]` whose
   values may contain commas sets `noSplit: true` (repeat the flag instead).
10. **Tests** use `test/helpers/graphMock.ts` and cover behaviour a user would notice: the right Graph call and
    filter, target resolution, shaping/derived fields, and that destructive commands preview before acting.
    No real tenant, no network; `fakeExo()` stands in for Exchange PowerShell, and `runCommand(..., { approve })`
    for the human. Tests must never open a dialog or prompt: the e2e config pins `confirmMode: token`.
11. **Other APIs.** Defender for Endpoint goes through `ctx.defender` (unversioned paths such as `/machines`).
    Exchange admin features with no Graph API go through `ctx.exo`; batch related cmdlets into one session (each
    one costs seconds), always `select` properties, and keep cmdlet names constant. Both need extra sign-in
    (`auth login --service exchange|defender`); declare their RBAC/API permissions rather than Graph scopes.
12. **Many items, one call.** A command that makes one Graph request per user/device/policy uses
    `ctx.graph.batch()`. Bulk commands read files with `readCsvRows()` and report a result per row.

## Errors and exit codes

Throw `CliError(code, message, { hint })` with a hint that tells the agent what to do next. Exit codes: 0 ok,
1 error, 2 usage, 3 confirmation required, 4 sign-in required, 5 permission denied, 6 not found, 7 writes
disabled, 8 ambiguous, 9 throttled, 10 not approved by a person.
