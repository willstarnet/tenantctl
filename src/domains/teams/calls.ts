import { CliError } from '../../core/errors.ts';
import { odata, seg } from '../../core/graph.ts';
import { listLimit, LIST_ARGS } from '../../core/listing.ts';
import { resolveUser } from '../../core/resolve.ts';
import { type AnyCommand, defineCommand } from '../../core/types.ts';
import { daysAgo, isGuid, settle } from '../../core/util.ts';

/**
 * Call records (communications/callRecords) are application-permission only: delegated sign-in cannot read
 * them. The deprecated `organizer`/`participants` properties stopped returning data on 2026-06-30, so
 * identities come from `organizer_v2`, `participants_v2` and endpoint `associatedIdentity`.
 */

const CALL_PERMISSIONS = { application: ['CallRecords.Read.All'] };

interface Identity {
  id?: string | null;
  displayName?: string | null;
}

interface CommunicationsIdentitySet {
  user?: Identity | null;
  guest?: Identity | null;
  phone?: Identity | null;
  application?: Identity | null;
}

interface ParticipantBase {
  id?: string;
  identity?: CommunicationsIdentitySet | null;
}

export interface CallRecord {
  id: string;
  type?: string;
  modalities?: string[];
  startDateTime?: string;
  endDateTime?: string;
  joinWebUrl?: string | null;
  organizer_v2?: ParticipantBase | null;
  sessions?: Session[];
  'sessions@odata.nextLink'?: string;
}

interface Endpoint {
  '@odata.type'?: string;
  associatedIdentity?: Identity | null;
  name?: string | null;
  userAgent?: { headerValue?: string | null; applicationVersion?: string | null } | null;
  feedback?: { rating?: string | null; text?: string | null } | null;
}

interface FailureInfo {
  stage?: string;
  reason?: string;
}

interface Session {
  id: string;
  caller?: Endpoint;
  callee?: Endpoint;
  failureInfo?: FailureInfo | null;
  segments?: Segment[];
}

interface NetworkInfo {
  connectionType?: string;
  wifiSignalStrength?: number | null;
  wifiBand?: string | null;
  receivedQualityEventRatio?: number | null;
  sentQualityEventRatio?: number | null;
  delayEventRatio?: number | null;
  bandwidthLowEventRatio?: number | null;
}

interface MediaStream {
  streamDirection?: 'callerToCallee' | 'calleeToCaller';
  averagePacketLossRate?: number | null;
  maxPacketLossRate?: number | null;
  averageJitter?: string | null;
  maxJitter?: string | null;
  averageRoundTripTime?: string | null;
  maxRoundTripTime?: string | null;
  averageAudioDegradation?: number | null;
  averageRatioOfConcealedSamples?: number | null;
}

interface Media {
  label?: string;
  callerNetwork?: NetworkInfo | null;
  calleeNetwork?: NetworkInfo | null;
  streams?: MediaStream[];
}

interface Segment {
  id: string;
  caller?: Endpoint;
  callee?: Endpoint;
  failureInfo?: FailureInfo | null;
  media?: Media[];
}

/** ISO 8601 duration (`PT0.045S`, `PT1M2.5S`) to milliseconds. */
export function durationMs(value: string | null | undefined): number | null {
  const match = value?.match(/^PT(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/);
  if (!match) return null;
  const [, h = '0', m = '0', s = '0'] = match;
  return Math.round((Number(h) * 3600 + Number(m) * 60 + Number(s)) * 1000);
}

function secondsBetween(start?: string, end?: string): number | null {
  if (!start || !end) return null;
  const ms = Date.parse(end) - Date.parse(start);
  return Number.isFinite(ms) ? Math.round(ms / 1000) : null;
}

function identityName(identity: CommunicationsIdentitySet | null | undefined): string | null {
  const who = identity?.user ?? identity?.guest ?? identity?.phone ?? identity?.application;
  return who?.displayName ?? who?.id ?? null;
}

/**
 * Thresholds Call Quality Dashboard uses to classify an audio stream as poor (any one is enough):
 * degradation > 1.0, round trip > 500 ms, packet loss > 10%, jitter > 30 ms, concealed samples > 7%.
 */
const POOR = { degradation: 1.0, roundTripMs: 500, packetLoss: 0.1, jitterMs: 30, concealed: 0.07 };

export function poorStreamReasons(stream: MediaStream): string[] {
  const reasons: string[] = [];
  const jitter = durationMs(stream.averageJitter);
  const roundTrip = durationMs(stream.averageRoundTripTime);
  if ((stream.averagePacketLossRate ?? 0) > POOR.packetLoss) reasons.push('packetLoss');
  if (jitter !== null && jitter > POOR.jitterMs) reasons.push('jitter');
  if (roundTrip !== null && roundTrip > POOR.roundTripMs) reasons.push('roundTrip');
  if ((stream.averageAudioDegradation ?? 0) > POOR.degradation) reasons.push('audioDegradation');
  if ((stream.averageRatioOfConcealedSamples ?? 0) > POOR.concealed) reasons.push('concealedSamples');
  return reasons;
}

interface ParticipantQuality {
  participant: string;
  id: string | null;
  client: string | null;
  connectionTypes: string[];
  minWifiSignalStrength: number | null;
  streams: number;
  poorStreams: number;
  poorReasons: string[];
  maxPacketLossRate: number | null;
  maxJitterMs: number | null;
  maxRoundTripMs: number | null;
  networkEventRatios: {
    receivedQuality: number | null;
    sentQuality: number | null;
    delay: number | null;
    bandwidthLow: number | null;
  };
  feedbackRating: string | null;
  verdict: 'poor' | 'ok' | 'no media data';
}

const maxOf = (current: number | null, next: number | null | undefined) =>
  typeof next === 'number' ? Math.max(current ?? next, next) : current;

type QualityAccumulator = ParticipantQuality & { reasonSet: Set<string>; connectionSet: Set<string> };

/**
 * Fold every segment's streams into one row per person. Each segment joins two endpoints (a user and the
 * Teams media service in meetings, two users in a peer-to-peer call); stream metrics count against every user
 * endpoint on the segment, and network facts against the side that reported them. Service endpoints are not
 * people and are skipped.
 */
export function summarizeCallQuality(sessions: Session[]): { participants: ParticipantQuality[]; failures: unknown[] } {
  const rows = new Map<string, QualityAccumulator>();
  const failures: unknown[] = [];
  const rowFor = (endpoint: Endpoint | undefined): QualityAccumulator | undefined => {
    if (!endpoint || endpoint['@odata.type']?.endsWith('serviceEndpoint')) return undefined;
    const identity = endpoint.associatedIdentity;
    const key = identity?.id ?? identity?.displayName ?? endpoint.name;
    if (!key) return undefined;
    let row = rows.get(key);
    if (!row) {
      row = {
        participant: identity?.displayName ?? endpoint.name ?? key,
        id: identity?.id ?? null,
        client: endpoint.userAgent?.headerValue ?? null,
        connectionTypes: [],
        minWifiSignalStrength: null,
        streams: 0,
        poorStreams: 0,
        poorReasons: [],
        maxPacketLossRate: null,
        maxJitterMs: null,
        maxRoundTripMs: null,
        networkEventRatios: { receivedQuality: null, sentQuality: null, delay: null, bandwidthLow: null },
        feedbackRating: null,
        verdict: 'no media data',
        reasonSet: new Set(),
        connectionSet: new Set(),
      };
      rows.set(key, row);
    }
    row.feedbackRating ??= endpoint.feedback?.rating ?? null;
    return row;
  };
  const addNetwork = (row: QualityAccumulator | undefined, network: NetworkInfo | null | undefined) => {
    if (!row || !network) return;
    if (network.connectionType && network.connectionType !== 'unknown') row.connectionSet.add(network.connectionType);
    const signal = network.wifiSignalStrength;
    if (typeof signal === 'number') row.minWifiSignalStrength = Math.min(row.minWifiSignalStrength ?? signal, signal);
    const ratios = row.networkEventRatios;
    ratios.receivedQuality = maxOf(ratios.receivedQuality, network.receivedQualityEventRatio);
    ratios.sentQuality = maxOf(ratios.sentQuality, network.sentQualityEventRatio);
    ratios.delay = maxOf(ratios.delay, network.delayEventRatio);
    ratios.bandwidthLow = maxOf(ratios.bandwidthLow, network.bandwidthLowEventRatio);
  };

  for (const session of sessions) {
    if (session.failureInfo?.reason) failures.push({ session: session.id, ...session.failureInfo });
    for (const segment of session.segments ?? []) {
      if (segment.failureInfo?.reason) failures.push({ segment: segment.id, ...segment.failureInfo });
      const caller = rowFor(segment.caller);
      const callee = rowFor(segment.callee);
      const people = [caller, callee].filter((row) => row !== undefined);
      for (const media of segment.media ?? []) {
        addNetwork(caller, media.callerNetwork);
        addNetwork(callee, media.calleeNetwork);
        for (const stream of media.streams ?? []) {
          const reasons = poorStreamReasons(stream);
          for (const row of people) {
            row.streams++;
            if (reasons.length) row.poorStreams++;
            for (const reason of reasons) row.reasonSet.add(reason);
            row.maxPacketLossRate = maxOf(
              row.maxPacketLossRate,
              stream.maxPacketLossRate ?? stream.averagePacketLossRate,
            );
            row.maxJitterMs = maxOf(row.maxJitterMs, durationMs(stream.maxJitter ?? stream.averageJitter));
            row.maxRoundTripMs = maxOf(
              row.maxRoundTripMs,
              durationMs(stream.maxRoundTripTime ?? stream.averageRoundTripTime),
            );
          }
        }
      }
    }
  }

  const participants = [...rows.values()].map(({ reasonSet, connectionSet, ...row }) => ({
    ...row,
    connectionTypes: [...connectionSet],
    poorReasons: [...reasonSet],
    verdict: row.streams === 0 ? ('no media data' as const) : row.poorStreams ? ('poor' as const) : ('ok' as const),
  }));
  return { participants, failures };
}

function shapeCallRecord(record: CallRecord) {
  return {
    id: record.id,
    type: record.type ?? null,
    modalities: record.modalities ?? [],
    startDateTime: record.startDateTime ?? null,
    endDateTime: record.endDateTime ?? null,
    durationSeconds: secondsBetween(record.startDateTime, record.endDateTime),
    organizer: identityName(record.organizer_v2?.identity),
  };
}

interface PstnCall {
  id: string;
  startDateTime?: string;
  endDateTime?: string;
  duration?: number;
  userPrincipalName?: string | null;
  userDisplayName?: string | null;
  callType?: string;
  callerNumber?: string;
  calleeNumber?: string;
  destinationName?: string;
  destinationContext?: string;
  charge?: number;
  currency?: string;
  operator?: string;
}

const REPORT_ROLES = ['Reports Reader', 'Teams Administrator', 'Teams Communications Administrator'];

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'teams',
    name: 'calls list',
    summary: 'List recent Teams calls and meetings (call records)',
    description:
      'Call records appear about 15 minutes to a few hours after a call ends and are kept 30 days. --user limits ' +
      'to calls that user took part in. Needs app-only auth (application permission CallRecords.Read.All); ' +
      'delegated sign-in cannot read call records. Next: `tenantctl teams calls get <id>` for per-participant quality.',
    examples: ['tenantctl teams calls list --user jane@contoso.com --days 2'],
    args: [
      { name: 'days', type: 'number', default: 1, description: 'Calls that started in the last N days (max 30).' },
      { name: 'user', type: 'string', description: 'Only calls this user joined (UPN, email, name or id).' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: { application: ['CallRecords.Read.All', 'User.Read.All'] },
    async run(ctx, args) {
      const filters = [`startDateTime ge ${daysAgo(args.days)}`];
      if (args.user) {
        const user = await resolveUser(ctx, args.user);
        filters.push(`participants_v2/any(p:p/id eq ${odata(user.id)})`);
      }
      // The list endpoint documents only $select and $filter, so no $top.
      const result = await ctx.graph.list<CallRecord>('/communications/callRecords', {
        query: { $filter: filters.join(' and ') },
        limit: listLimit(args),
        pageSize: false,
      });
      return { items: ctx.full ? result.items : result.items.map(shapeCallRecord), truncated: result.truncated };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'calls get',
    summary: 'Diagnose one call: per-participant audio/video quality, network and failures',
    description:
      'For "my Teams call was choppy/dropped". `quality` has one row per person: verdict poor|ok, the worst ' +
      'packet loss (0-1), jitter and round trip (ms), poor stream count with reasons, connection type (wifi/' +
      'wired) and weakest Wi-Fi signal. A stream is poor by Call Quality Dashboard rules: packet loss > 10%, ' +
      'jitter > 30 ms, round trip > 500 ms, degradation > 1.0 or concealed samples > 7%. networkEventRatios ' +
      'are the fraction of the call the client saw bad network; high `receivedQuality` or `delay` with ' +
      "connection wifi usually points at the user's Wi-Fi. `failures` lists setup/midcall failures. " +
      'Needs app-only auth (CallRecords.Read.All).',
    examples: ['tenantctl teams calls get 2f1a6c87-...'],
    args: [
      {
        name: 'id',
        type: 'string',
        required: true,
        positional: true,
        description: 'Call record id from `teams calls list`.',
      },
    ],
    risk: 'read',
    permissions: CALL_PERMISSIONS,
    async run(ctx, args) {
      const path = `/communications/callRecords/${seg(args.id)}`;
      // $expand takes either sessions or participants_v2, not both, so participants are a second call.
      const parts = await settle({
        record: ctx.graph.get<CallRecord>(path, { query: { $expand: 'sessions($expand=segments)' } }),
        participants: ctx.graph.list<ParticipantBase>(`${path}/participants_v2`, {
          limit: Number.POSITIVE_INFINITY,
          pageSize: false,
        }),
      });
      if ('error' in parts.record) {
        throw new CliError('NOT_FOUND', `Call record ${args.id}: ${parts.record.error.message}`, {
          hint: 'Find ids with `tenantctl teams calls list`; records are kept 30 days and need app-only auth.',
        });
      }
      const record = parts.record;
      const sessions = [...(record.sessions ?? [])];
      if (record['sessions@odata.nextLink']) {
        const more = await ctx.graph.list<Session>(record['sessions@odata.nextLink'], {
          limit: Number.POSITIVE_INFINITY,
          pageSize: false,
        });
        sessions.push(...more.items);
      }
      if (ctx.full) return { ...record, sessions, participants_v2: parts.participants };
      const { participants: quality, failures } = summarizeCallQuality(sessions);
      const poor = quality.filter((q) => q.verdict === 'poor').map((q) => q.participant);
      return {
        ...shapeCallRecord(record),
        joinWebUrl: record.joinWebUrl ?? null,
        participants:
          'error' in parts.participants
            ? parts.participants
            : parts.participants.items.map((p) => identityName(p.identity)),
        summary: poor.length
          ? `Poor quality for ${poor.join(', ')}`
          : quality.some((q) => q.verdict === 'ok')
            ? 'No poor streams detected'
            : 'No media quality data in this record',
        quality,
        failures,
      };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'calls pstn',
    summary: 'List phone (PSTN) calls made or received through Teams Phone',
    description:
      'Covers Microsoft Calling Plans and Operator Connect calls (not Direct Routing). Shows who called which ' +
      'number, duration in seconds and cost. Max range 90 days. Needs app-only auth (CallRecords.Read.All).',
    examples: ['tenantctl teams calls pstn --days 7'],
    args: [
      { name: 'days', type: 'number', default: 7, description: 'Calls that started in the last N days (max 90).' },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: CALL_PERMISSIONS,
    async run(ctx, args) {
      if (args.days <= 0 || args.days > 90) {
        throw new CliError('USAGE', '--days must be between 1 and 90', {
          hint: 'Graph limits PSTN call queries to 90 days.',
        });
      }
      const from = daysAgo(args.days);
      const to = new Date().toISOString();
      const result = await ctx.graph.list<PstnCall>(
        `/communications/callRecords/getPstnCalls(fromDateTime=${from},toDateTime=${to})`,
        { limit: listLimit(args), pageSize: false },
      );
      if (ctx.full) return result;
      return {
        items: result.items.map((c) => ({
          id: c.id,
          startDateTime: c.startDateTime ?? null,
          durationSeconds: c.duration ?? null,
          user: c.userPrincipalName ?? c.userDisplayName ?? null,
          callType: c.callType ?? null,
          callerNumber: c.callerNumber ?? null,
          calleeNumber: c.calleeNumber ?? null,
          destination: c.destinationName ?? null,
          destinationContext: c.destinationContext ?? null,
          charge: c.charge ?? null,
          currency: c.currency ?? null,
          operator: c.operator ?? null,
        })),
        truncated: result.truncated,
      };
    },
  }),

  defineCommand({
    domain: 'teams',
    name: 'activity',
    summary: 'Teams usage per user: last activity, chats, calls, meetings',
    description:
      'From the Microsoft 365 usage report, which lags about 2 days. Use to check whether someone actually uses ' +
      'Teams. If the tenant conceals user details in reports (Microsoft 365 admin center > Settings > Org ' +
      'settings > Reports), names and UPNs are replaced by hashes and --user cannot match. Delegated callers ' +
      'need a reports role such as Reports Reader.',
    examples: ['tenantctl teams activity --user jane@contoso.com', 'tenantctl teams activity --period D30 --top 20'],
    args: [
      { name: 'user', type: 'string', description: 'Only this user (UPN, email, name or id).' },
      {
        name: 'period',
        type: 'string',
        choices: ['D7', 'D30', 'D90', 'D180'],
        default: 'D7',
        description: 'Report window.',
      },
      ...LIST_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['Reports.Read.All'],
      application: ['Reports.Read.All'],
      roles: REPORT_ROLES,
    },
    async run(ctx, args) {
      const rows = await ctx.graph.report(`/reports/getTeamsUserActivityUserDetail(period=${odata(args.period)})`);
      let selected = rows;
      if (args.user) {
        const ref = args.user.trim();
        const keys =
          ref.includes('@') || isGuid(ref)
            ? [ref.toLowerCase()]
            : await resolveUser(ctx, ref).then((u) => [u.userPrincipalName.toLowerCase(), u.id.toLowerCase()]);
        selected = rows.filter(
          (r) =>
            keys.includes((r['User Principal Name'] ?? '').toLowerCase()) ||
            keys.includes((r['User Id'] ?? '').toLowerCase()),
        );
        if (!selected.length) {
          const concealed = rows.length > 0 && rows.every((r) => !(r['User Principal Name'] ?? '').includes('@'));
          throw new CliError('NOT_FOUND', `No Teams activity row for ${ref} in the ${args.period} report`, {
            hint: concealed
              ? 'User details are concealed in reports; an admin can turn that off in Microsoft 365 admin center > Settings > Org settings > Reports.'
              : 'The user may be unlicensed for Teams or new (reports lag ~2 days).',
          });
        }
      }
      const limit = listLimit(args);
      const items = selected.slice(0, limit);
      const number = (value: string | undefined) => (value === undefined || value === '' ? null : Number(value));
      return {
        items: ctx.full
          ? items
          : items.map((r) => ({
              user: r['User Principal Name'] ?? null,
              lastActivityDate: r['Last Activity Date'] || null,
              teamChatMessages: number(r['Team Chat Message Count']),
              privateChatMessages: number(r['Private Chat Message Count']),
              calls: number(r['Call Count']),
              meetings: number(r['Meeting Count']),
              meetingsOrganized: number(r['Meetings Organized Count']),
              isLicensed: r['Is Licensed'] ? r['Is Licensed'] === 'Yes' : null,
              reportRefreshDate: r['Report Refresh Date'] ?? null,
            })),
        truncated: selected.length > limit,
      };
    },
  }),
];
