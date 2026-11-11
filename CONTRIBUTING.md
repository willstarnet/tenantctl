# Contributing

Thanks for helping with tenantctl. Small, tested pull requests are the easiest to review. Everyone taking part
follows the [Code of Conduct](CODE_OF_CONDUCT.md). Report security problems privately as
[SECURITY.md](SECURITY.md) describes, never in a public issue; questions go where [SUPPORT.md](SUPPORT.md) says.

## Setup

You need [Bun](https://bun.sh) 1.3 or later. Everything else installs from the lockfile.

| Task            | Command                                                                 |
| --------------- | ----------------------------------------------------------------------- |
| Install deps    | `bun install`                                                           |
| Run from source | `bun src/cli.ts <area> <command> ...`                                   |
| Typecheck       | `bun run typecheck`                                                     |
| Tests           | `bun test` (one file: `bun test test/domains/intune.test.ts`)           |
| Format          | `bun run format` (CI runs `bun run format:check`)                       |
| Command docs    | `bun run docs` regenerates `docs/COMMANDS.md` and `docs/PERMISSIONS.md` |
| Build           | `bun run build` cross-compiles every target into `dist/`                |
| Install locally | `bun run install:local` (this machine, into `~/.local/bin/tenantctl`)   |
| Live check      | `bun run smoke:live` runs every read command against your tenant        |

Tests never touch a real tenant or the network: they run against a mock Graph with a throwaway config folder, and
approval prompts are stubbed. `smoke:live` needs `bun src/cli.ts auth login` first and never writes.

## Layout

```
src/cli.ts                 argv -> global flags -> command lookup -> execute -> render
src/registry.ts            every command, area summaries, lookup
src/core/types.ts          defineCommand, ArgSpec, Context, ListResult, Risk
src/core/execute.ts        write gating, human approval, confirm tokens (token mode), audit, permission hints
src/core/approval.ts       confirm modes, env-can-only-strengthen, OS dialog and terminal prompt
src/core/graph.ts          GraphClient: retries, paging, errors, $batch; seg() and odata() encoders
src/core/exo.ts            Exchange Online PowerShell bridge (fixed script, parameters as data)
src/core/resolve.ts        resolveUser / resolveGroup / pickOne (names -> ids, AMBIGUOUS errors)
src/core/auth.ts           MSAL per resource (Graph, Exchange, Defender)
src/core/secretStore.ts    encrypted token cache; key in Keychain / libsecret / DPAPI / file
src/domains/<area>/        the commands for one area
src/domains/troubleshoot/  composite diagnostics: verdict, ranked findings, next commands
test/helpers/graphMock.ts  mockContext(), runCommand(), fakeExo(), page(), graphError()
skills/tenantctl/SKILL.md  how an AI agent should use the CLI
scripts/                   docs generator, build, install scripts, release manifests, live smoke test
```

## Writing a command

```ts
export const devicesSync = defineCommand({
  domain: 'intune',
  name: 'devices sync', // `tenantctl intune devices sync`
  summary: 'Ask a device to check in with Intune now',
  description: 'When to use it and how to read the result.',
  examples: ['tenantctl intune devices sync DESKTOP-123'],
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

1. **Accept what a person would say.** Users by UPN, email, display name or id (`resolveUser`); groups by name,
   mail or id (`resolveGroup`); devices by name, serial or id; sites by URL; teams by name. Use `pickOne` so
   several matches become an `AMBIGUOUS` error listing the candidates, never a silent first pick.
2. **Compact output by default.** Return the fields an admin needs, with friendly names and derived facts
   (`daysSinceLastSync`, `forwardsExternally`). Return the raw Graph object when `ctx.full` is true.
3. **Lists return `ListResult`** (`{ items, truncated }`) via `ctx.graph.list(path, { limit: listLimit(args) })`,
   spread `...LIST_ARGS` into `args`, and send `$select`.
4. **Risk is honest.** `read` changes nothing. `write` is low-impact or easily reversed. `destructive` is
   disruptive, security-sensitive or irreversible and must implement `preview`, which resolves the target, shows
   its current state and carries a plain-language `action` string for the approval dialog. The executor handles
   gating, approval and audit; commands never check `allowWrites` themselves.
5. **Permissions are least privilege**, taken from the "Permissions" table of the endpoint's Microsoft Graph doc
   page for both delegated and application, plus `roles` when the caller also needs an admin role. They generate
   `tenantctl permissions` and the sign-in scopes, so they must be accurate.
6. **Verify every endpoint against current Microsoft docs.** Prefer v1.0; use beta only when v1.0 lacks the
   feature, and say so in the description.
7. **Encode user input**: `seg()` for path segments, `odata()` for filter literals. Never interpolate raw input.
8. **Composite views use `settle()`** so one forbidden section degrades to `{ error }` instead of failing.
   Troubleshoot findings carry a `next` that is a runnable command, because agents run it verbatim.
9. **Arg names** are camelCase (`--kebab-case` on the CLI) and must not clash with global flags (`output`,
   `fields`, `full`, `profile`, `confirm`, `dryRun`, `debug`, `help`, `version`). Mark secrets `secret: true`.
10. **Other APIs.** Defender for Endpoint goes through `ctx.defender`; Exchange features with no Graph API through
    `ctx.exo`, batching related cmdlets into one session and keeping cmdlet names constant.
11. **Many items, one call.** Use `ctx.graph.batch()` instead of one request per item; bulk commands read CSV files
    with `readCsvRows()` and report a result per row.

Errors are `CliError(code, message, { hint })`, with a hint that says what to do next. Exit codes: 0 ok, 1 error,
2 usage, 3 confirmation required (token mode), 4 sign-in required, 5 permission denied, 6 not found, 7 writes
disabled, 8 ambiguous, 9 throttled, 10 not approved by a person.

## Tests

Tests use `test/helpers/graphMock.ts` and cover behaviour a user would notice: the right Graph call and filter,
target resolution, output shaping, and that destructive commands preview and wait for approval. `fakeExo()` stands
in for Exchange PowerShell and `runCommand(..., { approve })` for the person. A bug fix starts with a test that
reproduces it. Test data uses `contoso.com`-style names, never real tenant data.

## Pull requests

1. For anything bigger than a bug fix, open an issue first so the approach can be agreed on.
2. Branch from `main`; one change per pull request. Pull requests are squash-merged.
3. Use [Conventional Commits](https://www.conventionalcommits.org/) for the title: `feat: ...`, `fix: ...`,
   `docs: ...`, `test: ...`, `build: ...`, `chore: ...`, optionally scoped (`fix(intune): ...`).
4. Add a line under `## [Unreleased]` in [CHANGELOG.md](CHANGELOG.md) for anything a user would notice, and run
   `bun run docs` when commands, arguments or permissions change.
5. CI must pass on macOS, Linux and Windows: format, typecheck, tests and up-to-date command docs.

## Releases

Maintainers move the `## [Unreleased]` entries under a new version heading, then push a `vX.Y.Z` tag. The release
workflow tests, builds every target, publishes the GitHub release, and updates the Homebrew tap.
