import type { Profile } from './config.ts';
import type { ExoRunner } from './exo.ts';
import type { GraphClient } from './graph.ts';

export type ArgType = 'string' | 'number' | 'boolean' | 'string[]';

export interface ArgSpec {
  /** camelCase key passed to `run`; the CLI flag is its kebab-case form (`dryRun` -> `--dry-run`). */
  readonly name: string;
  readonly type: ArgType;
  readonly description: string;
  readonly required?: boolean;
  /** Accept the value as a bare positional (in declaration order) as well as `--flag`. */
  readonly positional?: boolean;
  readonly default?: string | number | boolean | readonly string[];
  readonly choices?: readonly string[];
  /** Single-letter short flag. */
  readonly short?: string;
  /** Value is a secret (password etc.): redacted from audit logs and confirm `rerun` hints. */
  readonly secret?: boolean;
  /** `string[]` only: each occurrence is one whole value (values may contain commas); repeat the flag instead. */
  readonly noSplit?: boolean;
}

type ArgValue<S extends ArgSpec> = S['type'] extends 'string'
  ? S['choices'] extends readonly (infer C extends string)[]
    ? C
    : string
  : S['type'] extends 'number'
    ? number
    : S['type'] extends 'boolean'
      ? boolean
      : string[];

/**
 * Parsed arguments for a command. Required args, args with a default, and booleans are always present;
 * everything else may be undefined.
 */
export type ArgsOf<T extends readonly ArgSpec[]> = {
  [S in T[number] as S['name']]: S extends { required: true } | { default: unknown } | { type: 'boolean' }
    ? ArgValue<S>
    : ArgValue<S> | undefined;
};

/**
 * - read: no tenant change.
 * - write: reversible/low-impact change (sync a device, add a group member). Needs writes enabled.
 * - destructive: disruptive or irreversible (wipe, delete, disable, password reset). Needs writes enabled AND
 *   a person's approval of the preview (OS dialog or terminal; in `token` mode a confirm token from a preview run).
 */
export type Risk = 'read' | 'write' | 'destructive';

/**
 * `delegated` and `application` are Microsoft Graph permissions ONLY: they become the scopes `tenantctl auth login`
 * requests, and Entra refuses the whole sign-in if one belongs to another API (AADSTS650053). Other APIs have
 * their own fields.
 */
export interface Permissions {
  /** Least-privileged delegated Graph scopes (work/school). */
  readonly delegated?: readonly string[];
  /** Least-privileged Graph application permissions. */
  readonly application?: readonly string[];
  /** Directory/Intune/Defender role the signed-in admin needs in delegated mode, if any. */
  readonly roles?: readonly string[];
  /**
   * Uses the Exchange Online PowerShell bridge: delegated `Exchange.Manage` (via `auth login --service exchange`)
   * or application `Exchange.ManageAsApp` on Office 365 Exchange Online, plus the Exchange RBAC role in `roles`.
   */
  readonly exchange?: boolean;
  /** Defender for Endpoint (WindowsDefenderATP) permissions, same names delegated and app-only. */
  readonly defender?: readonly string[];
}

export interface Context {
  readonly graph: GraphClient;
  /** Defender for Endpoint API (`https://api.security.microsoft.com/api`), same client semantics as Graph. */
  readonly defender: GraphClient;
  /** Exchange Online PowerShell bridge for admin features Graph lacks. */
  readonly exo: ExoRunner;
  readonly profileName: string;
  readonly profile: Profile;
  /** `--full`: return raw Graph objects instead of the compact shape. */
  readonly full: boolean;
  /** Progress/diagnostic output to stderr; never to stdout. */
  log(message: string): void;
  /**
   * Ask the human (terminal or OS dialog, per the profile's confirm mode) and throw DENIED unless they approve.
   * The executor calls this for destructive commands; commands call it only for local security settings.
   */
  requireApproval(summary: string): Promise<void>;
}

export interface CommandSpec<T extends readonly ArgSpec[] = readonly ArgSpec[]> {
  /** Top-level group, e.g. `intune`. */
  readonly domain: string;
  /** Command path under the domain, space separated, e.g. `devices wipe`. */
  readonly name: string;
  /** One line, imperative, shown in listings. */
  readonly summary: string;
  /** Extra guidance for agents: when to use it, what the output means, caveats. */
  readonly description?: string;
  readonly examples?: readonly string[];
  readonly args: T;
  readonly risk: Risk;
  readonly permissions: Permissions;
  run(ctx: Context, args: ArgsOf<T>): Promise<unknown>;
  /**
   * Describe exactly what a write would do, resolving targets (ids, names, current state). Required for
   * destructive commands; used by `--dry-run` and the confirmation step.
   */
  preview?(ctx: Context, args: ArgsOf<T>): Promise<unknown>;
}

/** Type-erased command, as stored in the registry. */
export interface AnyCommand {
  readonly domain: string;
  readonly name: string;
  readonly summary: string;
  readonly description?: string;
  readonly examples?: readonly string[];
  readonly args: readonly ArgSpec[];
  readonly risk: Risk;
  readonly permissions: Permissions;
  run(ctx: Context, args: Record<string, unknown>): Promise<unknown>;
  preview?(ctx: Context, args: Record<string, unknown>): Promise<unknown>;
}

export function defineCommand<const T extends readonly ArgSpec[]>(spec: CommandSpec<T>): AnyCommand {
  return spec as unknown as AnyCommand;
}

/** Return shape for list commands: the CLI moves `truncated` into the envelope's `meta`. */
export interface ListResult<T = unknown> {
  readonly items: T[];
  /** More results exist beyond the requested limit. */
  readonly truncated: boolean;
}

export function isListResult(value: unknown): value is ListResult {
  return (
    typeof value === 'object' &&
    value !== null &&
    Array.isArray((value as ListResult).items) &&
    typeof (value as ListResult).truncated === 'boolean'
  );
}
