# Contributing

Thanks for helping with tenantctl. Small, tested pull requests are the easiest to review. Everyone taking part
follows the [Code of Conduct](CODE_OF_CONDUCT.md). Report security problems privately as
[SECURITY.md](SECURITY.md) describes, never in a public issue; questions go where [SUPPORT.md](SUPPORT.md) says.

## Setup

You need [Bun](https://bun.sh) 1.3 or later. Everything else installs from the lockfile:

```sh
bun install
bun run typecheck && bun test   # what CI runs, on macOS, Linux and Windows
bun src/cli.ts --help           # run from source
```

Tests never touch a real tenant or the network: they run against a mock Graph (`test/helpers/graphMock.ts`) with
a throwaway config folder, and approval prompts are stubbed. To try read commands against your own tenant, sign in
with `bun src/cli.ts auth login`; `bun run smoke:live` runs every read command and never writes.

## Checks

| Step   | Command                                                       |
| ------ | ------------------------------------------------------------- |
| format | `bun run format:check` (fix with `bun run format`)            |
| types  | `bun run typecheck`                                           |
| tests  | `bun test` (one file: `bun test test/domains/intune.test.ts`) |
| docs   | `bun run docs` when commands, arguments or permissions change |
| build  | `bun run build` cross-compiles every target into `dist/`      |

## Writing commands

[AGENTS.md](AGENTS.md) is the guide: how a command is defined, how targets are resolved from names, risk levels and
previews, least-privilege permissions from the Microsoft Graph docs, error codes, and what tests must cover. It is
written for AI coding agents and people alike.

## Guarantees reviewers hold the line on

- A `destructive` command implements `preview` and never makes a change without the executor's approval step;
  commands never check `allowWrites` themselves.
- Environment variables can only make approval stricter, never weaker.
- User input is encoded with `seg()` / `odata()` before it reaches a Graph path or filter; Exchange cmdlet names
  stay constant and parameters travel as data.
- Secrets are marked `secret: true`, redacted from the audit log and `rerun` hints, and never printed.
- Permissions match the endpoint's Microsoft Graph documentation, delegated and application.
- No tenant data in the repository: test fixtures use `contoso.com` style names.

## Pull requests

1. For anything bigger than a bug fix, open an issue first so the approach can be agreed on.
2. Branch from `main`; one change per pull request.
3. Use [Conventional Commits](https://www.conventionalcommits.org/): `feat: ...`, `fix: ...`, `docs: ...`,
   `test: ...`, `build: ...`, `chore: ...`, optionally scoped (`fix(intune): ...`).
4. Add a line under `## [Unreleased]` in [CHANGELOG.md](CHANGELOG.md) for anything a user would notice, and update
   the README or `docs/domains/` when commands or behaviour change.
5. The checks above pass and CI is green on macOS, Linux and Windows.

## Releases

Maintainers move the `## [Unreleased]` entries under a new version heading, then push a `vX.Y.Z` tag. The release
workflow tests, builds every target, publishes the GitHub release, and updates the Homebrew tap and Scoop bucket.
