/**
 * Read-only smoke run against a real tenant. Discovers sample targets (the signed-in user, a group, a device, a
 * team, a site...), then runs every `risk: 'read'` command with them, plus explicit probes of the filters Microsoft
 * does not document. Each result is ok / unavailable (licence, permission, setup) / broken (a bug here) /
 * skipped (no sample data in this tenant).
 *
 *   bun run smoke:live [--domain intune] [--only "entra users get"] [--json] [--concurrency 4]
 *
 * It never runs a write: anything not `risk: 'read'`, and the local auth/config areas, are refused outright.
 */
import { kebab } from '../src/core/args.ts';
import { resolveProfile } from '../src/core/config.ts';
import { CliError, toCliError } from '../src/core/errors.ts';
import { createContext, execute } from '../src/core/execute.ts';
import type { AnyCommand, Context } from '../src/core/types.ts';
import { COMMANDS, commandId } from '../src/registry.ts';

type Status = 'ok' | 'unavailable' | 'broken' | 'skipped';
interface Row {
  command: string;
  argv: string;
  status: Status;
  ms: number;
  note: string;
}

const LOCAL_AREAS: Record<string, true> = { auth: true, config: true, tools: true, permissions: true };
const flags = process.argv.slice(2);
const flag = (name: string) => {
  const index = flags.indexOf(`--${name}`);
  return index >= 0 ? flags[index + 1] : undefined;
};
const onlyDomain = flag('domain');
const onlyCommand = flag('only');
const concurrency = Number(flag('concurrency') ?? 4);

const { name: profileName, profile } = resolveProfile(flag('profile'));
const ctx = createContext({ profileName, profile, log: () => {} }, { full: false, debug: false });

/** First item of a collection, or undefined when the tenant has none / the caller cannot read it. */
async function first<T>(
  context: Context,
  path: string,
  query: Record<string, string> = {},
  pageSize: number | false = 100,
) {
  try {
    // Some collections (team channels) reject $top; pageSize false leaves it off.
    const { items } = await context.graph.list<T>(path, { query, limit: 1, pageSize });
    return items[0];
  } catch {
    return undefined;
  }
}

interface Ref {
  id: string;
}

/** Sample targets found in the tenant; each is undefined when none exists or it cannot be read. */
interface Samples {
  me?: { id: string; userPrincipalName: string };
  group?: Ref;
  managedDevice?: { id: string; deviceName: string };
  entraDevice?: Ref;
  sp?: { appId: string };
  ca?: Ref;
  compliance?: Ref;
  config?: Ref;
  mobileApp?: Ref;
  incident?: Ref;
  alert?: Ref;
  team?: Ref;
  channel?: Ref;
  site?: { webUrl: string };
  issue?: Ref;
  message?: Ref;
  call?: Ref;
  file?: { webUrl: string };
  accessReview?: Ref;
  ediscoveryCase?: Ref;
  auditQuery?: Ref;
}

/** First file (not folder) in a user's OneDrive root; the children endpoint can't filter on the file facet. */
async function firstFile(context: Context, userId: string): Promise<{ webUrl: string } | undefined> {
  try {
    const { items } = await context.graph.list<{ webUrl: string; file?: unknown }>(
      `/users/${userId}/drive/root/children`,
      {
        query: { $select: 'webUrl,file' },
        limit: 50,
      },
    );
    return items.find((item) => item.file);
  } catch {
    return undefined;
  }
}

async function discover(): Promise<Samples> {
  const me =
    profile.authMode === 'app'
      ? await first<{ id: string; userPrincipalName: string }>(ctx, '/users', {
          $select: 'id,userPrincipalName',
          $filter: "userType eq 'Member'",
        })
      : await ctx.graph
          .get<{ id: string; userPrincipalName: string }>('/me', { query: { $select: 'id,userPrincipalName' } })
          .catch(() => undefined);
  const team = await first<{ id: string }>(ctx, '/groups', {
    $filter: "resourceProvisioningOptions/Any(x:x eq 'Team')",
    $select: 'id',
  });
  const [
    group,
    managedDevice,
    entraDevice,
    sp,
    ca,
    compliance,
    config,
    mobileApp,
    incident,
    alert,
    site,
    issue,
    message,
    call,
    channel,
    file,
    accessReview,
    ediscoveryCase,
    auditQuery,
  ] = await Promise.all([
    first<{ id: string }>(ctx, '/groups', { $select: 'id' }),
    first<{ id: string; deviceName: string }>(ctx, '/deviceManagement/managedDevices', { $select: 'id,deviceName' }),
    first<{ id: string }>(ctx, '/devices', { $select: 'id' }),
    first<{ appId: string }>(ctx, '/servicePrincipals', {
      $select: 'appId',
      $filter: "servicePrincipalType eq 'Application'",
    }),
    first<{ id: string }>(ctx, '/identity/conditionalAccess/policies', { $select: 'id' }),
    first<{ id: string }>(ctx, '/deviceManagement/deviceCompliancePolicies', { $select: 'id' }),
    first<{ id: string }>(ctx, '/deviceManagement/deviceConfigurations', { $select: 'id' }),
    first<{ id: string }>(ctx, '/deviceAppManagement/mobileApps', { $select: 'id' }),
    first<{ id: string }>(ctx, '/security/incidents'),
    first<{ id: string }>(ctx, '/security/alerts_v2'),
    ctx.graph.get<{ webUrl: string }>('/sites/root', { query: { $select: 'webUrl' } }).catch(() => undefined),
    first<{ id: string }>(ctx, '/admin/serviceAnnouncement/issues', { $select: 'id' }),
    first<{ id: string }>(ctx, '/admin/serviceAnnouncement/messages', { $select: 'id' }),
    first<{ id: string }>(ctx, '/communications/callRecords'),
    team ? first<{ id: string }>(ctx, `/teams/${team.id}/channels`, { $select: 'id' }, false) : undefined,
    me ? firstFile(ctx, me.id) : undefined,
    first<{ id: string }>(ctx, '/identityGovernance/accessReviews/definitions', { $select: 'id' }),
    first<{ id: string }>(ctx, '/security/cases/ediscoveryCases', { $select: 'id' }),
    first<{ id: string }>(ctx, '/security/auditLog/queries', {}, false),
  ]);
  return {
    me,
    group,
    managedDevice,
    entraDevice,
    sp,
    ca,
    compliance,
    config,
    mobileApp,
    incident,
    alert,
    team,
    channel,
    site,
    issue,
    message,
    call,
    file,
    accessReview,
    ediscoveryCase,
    auditQuery,
  };
}

/** Optional arguments a command still needs to do anything useful (one of several, or delegated-only rules). */
const EXTRA_ARGS: Record<string, string[]> = {
  'defender hunt': ['print ok = 1', '--timespan', 'P1D'],
  'sharepoint sites list': ['--search', 'a'],
};

/** The value to pass for a required argument, or undefined when the tenant has nothing suitable. */
function sample(command: AnyCommand, arg: string, s: Samples): string | undefined {
  const id = commandId(command);
  const byCommand: Record<string, string | undefined> = {
    'entra ca get:policy': s.ca?.id,
    'intune compliance status:policy': s.compliance?.id,
    'intune config status:policy': s.config?.id,
    'intune apps status:app': s.mobileApp?.id,
    'intune apps assignments:app': s.mobileApp?.id,
    'entra devices get:device': s.entraDevice?.id,
    'defender incidents get:id': s.incident?.id,
    'defender alerts get:id': s.alert?.id,
    'health issue:id': s.issue?.id,
    'health message:id': s.message?.id,
    'teams calls get:id': s.call?.id,
    'entra access-reviews decisions:reviewId': s.accessReview?.id,
    'purview ediscovery custodians:case': s.ediscoveryCase?.id,
    'purview audit results:queryId': s.auditQuery?.id,
    'exchange rooms free:start': new Date(Date.now() + 86_400_000).toISOString(),
    'exchange rooms free:end': new Date(Date.now() + 90_000_000).toISOString(),
    'sharepoint files search:query': 'report',
    'graph get:path': profile.authMode === 'app' ? '/organization' : '/me',
  };
  const key = `${id}:${arg}`;
  if (key in byCommand) return byCommand[key];
  const upn = s.me?.userPrincipalName;
  const byName: Record<string, string | undefined> = {
    user: upn,
    users: upn,
    recipient: upn,
    mailbox: upn,
    owner: upn,
    group: s.group?.id,
    device: s.managedDevice?.deviceName,
    machine: s.managedDevice?.deviceName,
    app: s.sp?.appId,
    team: s.team?.id,
    channel: s.channel?.id,
    site: s.site?.webUrl,
    ref: s.file?.webUrl,
  };
  return byName[arg];
}

/** Explicit probes of filters Microsoft does not document, which only a real tenant can prove. */
function probes(s: Samples): [string, string[]][] {
  const upn = s.me?.userPrincipalName;
  return [
    ['entra signins', ['--failures-only', '--days', '1', '--top', '5']],
    ['entra users list', ['--unlicensed', '--top', '5']],
    ['entra users list', ['--inactive-days', '90', '--top', '5']],
    ['defender incidents list', ['--unassigned', '--top', '5']],
    ['health messages', ['--action-required', '--top', '5']],
    ['health issues', ['--active', '--top', '5']],
    ['intune devices list', ['--os', 'windows', '--top', '5']],
    ['intune devices list', ['--compliance', 'noncompliant', '--stale-days', '30', '--top', '5']],
    ...(upn
      ? ([
          ['entra users signins', [upn, '--failures-only', '--days', '1', '--top', '5']],
          ['intune devices list', ['--user', upn, '--top', '5']],
        ] as [string, string[]][])
      : []),
  ];
}

function classify(error: CliError): Status {
  if (['FORBIDDEN', 'AUTH_REQUIRED', 'CONFIG', 'NOT_FOUND', 'THROTTLED'].includes(error.code)) return 'unavailable';
  if (/licen[cs]e|subscription|not (?:enabled|provisioned)|premium|P1|P2|tenant does not have/i.test(error.message))
    return 'unavailable';
  return 'broken';
}

async function runOne(command: AnyCommand, argv: string[]): Promise<Row> {
  const id = commandId(command);
  if (command.risk !== 'read' || LOCAL_AREAS[command.domain]) throw new Error(`refusing non-read command ${id}`);
  const started = Date.now();
  try {
    await execute(
      command,
      id,
      argv,
      { full: false, debug: false, dryRun: false },
      { profileName, profile, log: () => {} },
    );
    return { command: id, argv: argv.join(' '), status: 'ok', ms: Date.now() - started, note: '' };
  } catch (raw) {
    const error = toCliError(raw);
    return {
      command: id,
      argv: argv.join(' '),
      status: classify(error),
      ms: Date.now() - started,
      note: `${error.code}: ${error.message}`.slice(0, 200),
    };
  }
}

async function main() {
  process.stderr.write(`Discovering sample targets in profile "${profileName}"...\n`);
  const samples = await discover();
  const jobs: (() => Promise<Row>)[] = [];
  const selected = COMMANDS.filter(
    (c) =>
      c.risk === 'read' &&
      !LOCAL_AREAS[c.domain] &&
      (!onlyDomain || c.domain === onlyDomain) &&
      (!onlyCommand || commandId(c) === onlyCommand),
  );
  for (const command of selected) {
    const argv: string[] = [];
    let missing: string | undefined;
    for (const arg of command.args.filter((a) => a.required)) {
      const value = sample(command, arg.name, samples);
      if (!value) {
        missing = arg.name;
        break;
      }
      if (arg.positional) argv.push(value);
      else argv.push(`--${kebab(arg.name)}`, value);
    }
    if (command.args.some((a) => a.name === 'top')) argv.push('--top', '3');
    if (command.args.some((a) => a.name === 'waitSeconds')) argv.push('--wait-seconds', '30');
    const id = commandId(command);
    argv.push(...(EXTRA_ARGS[id] ?? []));
    jobs.push(async () =>
      missing
        ? { command: id, argv: '', status: 'skipped', ms: 0, note: `no sample for <${missing}> in this tenant` }
        : runOne(command, argv),
    );
  }
  for (const [id, argv] of probes(samples)) {
    const command = COMMANDS.find((c) => commandId(c) === id);
    if (command && (!onlyDomain || command.domain === onlyDomain) && !onlyCommand)
      jobs.push(() => runOne(command, argv));
  }

  const rows: Row[] = new Array(jobs.length);
  let next = 0;
  await Promise.all(
    Array.from({ length: Math.max(1, concurrency) }, async () => {
      while (next < jobs.length) {
        const index = next++;
        rows[index] = await (jobs[index] as () => Promise<Row>)();
        const row = rows[index] as Row;
        process.stderr.write(`${row.status.padEnd(11)} ${row.command} ${row.argv}\n`);
      }
    }),
  );

  const counts = rows.reduce<Record<string, number>>(
    (acc, row) => ({ ...acc, [row.status]: (acc[row.status] ?? 0) + 1 }),
    {},
  );
  if (flags.includes('--json')) {
    process.stdout.write(`${JSON.stringify({ profile: profileName, counts, rows }, null, 2)}\n`);
  } else {
    const order: Record<Status, number> = { broken: 0, unavailable: 1, skipped: 2, ok: 3 };
    for (const row of [...rows].sort((a, b) => order[a.status] - order[b.status])) {
      if (row.status === 'ok') continue;
      process.stdout.write(
        `${row.status.toUpperCase().padEnd(11)} ${`${row.command} ${row.argv}`.trim()}\n            ${row.note}\n`,
      );
    }
    process.stdout.write(`\n${JSON.stringify(counts)}  (${rows.length} runs)\n`);
  }
  process.exitCode = counts.broken ? 1 : 0;
}

await main();
