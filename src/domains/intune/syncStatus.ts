import { seg } from '../../core/graph.ts';
import type { LiveRegion } from '../../core/live.ts';
import type { Context } from '../../core/types.ts';
import { compact } from '../../core/util.ts';
import { DEVICES_PATH } from './common.ts';

/**
 * One stage of a device sync, as the Intune admin center's "Device sync status" pane shows it. Read with the beta
 * `managedDevice/getSyncStatus()` function, which is in the Graph beta metadata but not yet in the API reference.
 */
export interface SyncComponent {
  name: string;
  /** none (not started), inProgress, success, failure, partialSuccess. */
  status: string;
  reportedDateTime?: string;
  /** e.g. "35/37" for policies, "12 offered" for applications. */
  moreInfo?: string | null;
}

/** finished = every stage reported a result; waitingForDevice = notified but the device has not connected yet. */
export type SyncState = 'notStarted' | 'waitingForDevice' | 'inProgress' | 'succeeded' | 'partialSuccess' | 'failed';

const LABELS: Record<string, string> = {
  notifyingDevice: 'Notifying device',
  deviceConnecting: 'Device connecting',
  policies: 'Policies',
  applications: 'Applications',
  scripts: 'Scripts',
  compliance: 'Compliance',
};
/** Stages listed before the device reports them, like the admin center's "Not started" rows. */
const EXPECTED = ['notifyingDevice', 'deviceConnecting', 'policies', 'applications', 'scripts'];
const ORDER = Object.keys(LABELS);
const FINISHED = new Set(['success', 'failure', 'partialSuccess']);
const SESSION = new Set(['notifyingDevice', 'deviceConnecting']);
/**
 * Stages that can keep running for many minutes after the rest finish (compliance re-evaluation was still in
 * progress 4+ minutes after the others in testing). They are reported but don't hold up the result; the admin
 * center's pane doesn't list compliance either.
 */
const TRAILING = new Set(['compliance']);
/** Allowed clock difference between this machine and Intune when telling this sync's stages from the last one's. */
const SKEW_MS = 15_000;

export const DONE_STATES: ReadonlySet<SyncState> = new Set(['succeeded', 'partialSuccess', 'failed']);

const label = (name: string) => LABELS[name] ?? name.replace(/([A-Z])/g, ' $1').replace(/^./, (c) => c.toUpperCase());
const rank = (name: string) => (ORDER.includes(name) ? ORDER.indexOf(name) : ORDER.length);

export async function fetchSyncStatus(ctx: Context, deviceId: string): Promise<SyncComponent[]> {
  const response = await ctx.graph.get<{ components?: SyncComponent[] }>(
    `${DEVICES_PATH}/${seg(deviceId)}/getSyncStatus()`,
    { version: 'beta' },
  );
  return [...(response.components ?? [])].sort((a, b) => rank(a.name) - rank(b.name));
}

/** The stages reported since `startedAt` (ms); Intune keeps returning the previous sync's stages until then. */
export function stagesSince(components: SyncComponent[], startedAt: number): SyncComponent[] {
  return components.filter((c) => !c.reportedDateTime || Date.parse(c.reportedDateTime) >= startedAt - SKEW_MS);
}

export function syncState(components: SyncComponent[]): SyncState {
  if (!components.length) return 'notStarted';
  if (components.some((c) => c.name === 'notifyingDevice' && c.status === 'failure')) return 'failed';
  const workloads = components.filter((c) => !SESSION.has(c.name));
  if (!workloads.length) {
    return components.some((c) => c.name === 'deviceConnecting') ? 'inProgress' : 'waitingForDevice';
  }
  if (components.some((c) => !TRAILING.has(c.name) && !FINISHED.has(c.status))) return 'inProgress';
  if (components.some((c) => c.status === 'failure')) return 'failed';
  if (components.some((c) => c.status === 'partialSuccess')) return 'partialSuccess';
  return 'succeeded';
}

/** A stage as tenantctl reports it. */
export interface SyncStageView {
  stage: string;
  status: string;
  reportedDateTime?: string;
  detail?: string;
}

export function shapeStages(components: SyncComponent[]): SyncStageView[] {
  return components.map(
    (c) =>
      compact({
        stage: label(c.name),
        status: c.status,
        reportedDateTime: c.reportedDateTime,
        detail: c.moreInfo ?? undefined,
      }) as SyncStageView,
  );
}

const SPINNER = ['◐', '◓', '◑', '◒'];
const STATUS_TEXT: Record<string, string> = {
  success: 'Completed',
  failure: 'Failed',
  partialSuccess: 'Partial success',
  inProgress: 'In progress',
  none: 'Not started',
};
const STATE_TEXT: Record<SyncState, string> = {
  notStarted: 'Waiting for Intune to notify the device',
  waitingForDevice: 'Notified; waiting for the device to connect (it must be on and online)',
  inProgress: 'Syncing',
  succeeded: 'Sync completed',
  partialSuccess: 'Sync finished with partial success',
  failed: 'Sync failed',
};

function clock(ms: number): string {
  const seconds = Math.floor(ms / 1000);
  return `${Math.floor(seconds / 60)}:${String(seconds % 60).padStart(2, '0')}`;
}

function localTime(iso?: string): string {
  if (!iso) return '';
  const date = new Date(iso);
  return Number.isNaN(date.getTime()) ? '' : date.toLocaleTimeString([], { hour12: false });
}

/** The lines of the live panel: header with progress bar and elapsed time, one row per stage, then a status line. */
export function syncPanel(
  deviceName: string,
  components: SyncComponent[],
  view: { elapsedMs: number; frame: number; final: boolean; color: boolean },
): string[] {
  const paint = (text: string, code: number) => (view.color ? `\x1b[${code}m${text}\x1b[0m` : text);
  const byName = new Map(components.map((c) => [c.name, c]));
  const names = [...new Set([...EXPECTED, ...components.map((c) => c.name)])].sort((a, b) => rank(a) - rank(b));
  const finished = names.filter((n) => FINISHED.has(byName.get(n)?.status ?? '')).length;
  const filled = Math.round((finished / names.length) * 20);
  const bar = paint('█'.repeat(filled), 34) + paint('░'.repeat(20 - filled), 2);
  const rows = names.map((name) => {
    const stage = byName.get(name);
    const status = stage?.status ?? 'none';
    const icon =
      status === 'success'
        ? paint('✓', 32)
        : status === 'failure'
          ? paint('✗', 31)
          : status === 'partialSuccess'
            ? paint('!', 33)
            : status === 'inProgress'
              ? paint(SPINNER[view.frame % SPINNER.length] as string, 36)
              : paint('○', 2);
    const text = STATUS_TEXT[status] ?? status;
    return `  ${icon} ${label(name).padEnd(18)} ${text.padEnd(16)} ${localTime(stage?.reportedDateTime).padEnd(9)} ${stage?.moreInfo ?? ''}`.trimEnd();
  });
  const state = syncState(components);
  const footer = view.final
    ? `  ${DONE_STATES.has(state) ? STATE_TEXT[state] : `Stopped watching: ${STATE_TEXT[state].toLowerCase()}`}`
    : paint('  Ctrl-C stops watching; the sync itself carries on.', 2);
  return [`Sync status · ${deviceName}  ${bar} ${finished}/${names.length}  ${clock(view.elapsedMs)}`, ...rows, footer];
}

export interface WaitOptions {
  timeoutSeconds: number;
  deviceName: string;
  live?: LiveRegion;
  /** Overridable for tests. */
  sleep?: (ms: number) => Promise<void>;
  tickMs?: number;
}

export interface SyncOutcome {
  state: SyncState;
  stages: SyncStageView[];
  waitedSeconds: number;
  timedOut: boolean;
}

/**
 * Poll getSyncStatus every 3 s until every stage of the sync started at `startedAt` has a result, or `timeoutSeconds`
 * pass. With a live region, the panel redraws every second so the spinner and clock move between polls.
 */
export async function waitForSync(
  ctx: Context,
  deviceId: string,
  startedAt: number,
  options: WaitOptions,
): Promise<SyncOutcome> {
  const sleep = options.sleep ?? ((ms: number) => Bun.sleep(ms));
  const pollEvery = options.live ? 3 : 1;
  const tickMs = options.tickMs ?? (options.live ? 1000 : 3000);
  const lastTick = Math.ceil((options.timeoutSeconds * 1000) / tickMs);
  let components: SyncComponent[] = [];
  for (let tick = 0; ; tick++) {
    if (tick % pollEvery === 0) components = stagesSince(await fetchSyncStatus(ctx, deviceId), startedAt);
    const state = syncState(components);
    const done = DONE_STATES.has(state);
    const timedOut = !done && tick >= lastTick;
    options.live?.render(
      syncPanel(options.deviceName, components, {
        elapsedMs: tick * tickMs,
        frame: tick,
        final: done || timedOut,
        color: options.live.color,
      }),
    );
    if (done || timedOut) {
      return { state, stages: shapeStages(components), waitedSeconds: Math.round((tick * tickMs) / 1000), timedOut };
    }
    await sleep(tickMs);
  }
}
