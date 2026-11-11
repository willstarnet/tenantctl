import { CliError } from '../../core/errors.ts';
import { seg } from '../../core/graph.ts';
import { LIST_ARGS, listLimit } from '../../core/listing.ts';
import { type AnyCommand, type Context, defineCommand } from '../../core/types.ts';

interface RawRoom {
  id: string;
  displayName?: string | null;
  emailAddress?: string | null;
  building?: string | null;
  floorLabel?: string | null;
  floorNumber?: number | null;
  capacity?: number | null;
  bookingType?: string | null;
  isWheelChairAccessible?: boolean | null;
  videoDeviceName?: string | null;
  displayDeviceName?: string | null;
}

interface ScheduleInformation {
  scheduleId?: string;
  scheduleItems?: { status?: string; start?: { dateTime?: string }; end?: { dateTime?: string } }[];
  error?: { message?: string; responseCode?: string } | null;
}

const ROOM_SELECT =
  'id,displayName,emailAddress,building,floorLabel,floorNumber,capacity,bookingType,isWheelChairAccessible,videoDeviceName,displayDeviceName';

/** Schedules per getSchedule call (conservative; the API takes a collection). */
const SCHEDULES_PER_CALL = 20;

const FILTER_ARGS = [
  { name: 'building', type: 'string', description: 'Only rooms whose building contains this text.' },
  { name: 'capacity', type: 'number', description: 'Only rooms that seat at least N people.' },
] as const;

function shapeRoom(room: RawRoom) {
  return {
    name: room.displayName ?? null,
    email: room.emailAddress ?? null,
    building: room.building ?? null,
    floor: room.floorLabel ?? room.floorNumber ?? null,
    capacity: room.capacity ?? null,
    bookingType: room.bookingType ?? null,
    wheelchairAccessible: room.isWheelChairAccessible ?? null,
    video: room.videoDeviceName ?? null,
    display: room.displayDeviceName ?? null,
  };
}

/** Every room in the tenant matching --building / --capacity (filtered locally; room lists are small). */
async function findRooms(ctx: Context, args: { building?: string; capacity?: number }): Promise<RawRoom[]> {
  const { items } = await ctx.graph.list<RawRoom>('/places/microsoft.graph.room', {
    query: { $select: ROOM_SELECT },
    limit: Number.POSITIVE_INFINITY,
  });
  const building = args.building?.trim().toLowerCase();
  return items
    .filter((r) => !building || (r.building ?? '').toLowerCase().includes(building))
    .filter((r) => args.capacity === undefined || (r.capacity ?? 0) >= args.capacity)
    .sort((a, b) => (a.displayName ?? '').localeCompare(b.displayName ?? ''));
}

function parseInstant(value: string, flag: string): Date {
  const date = new Date(value);
  if (Number.isNaN(date.getTime()))
    throw new CliError('USAGE', `${flag} must be an ISO date-time, e.g. 2026-11-12T14:00:00Z`);
  return date;
}

/** Graph dateTimeTimeZone in UTC (no offset suffix). */
function utc(date: Date) {
  return { dateTime: date.toISOString().replace(/Z$/, ''), timeZone: 'UTC' };
}

/** A room is free when no schedule item other than `free` falls in the window. */
export function isScheduleFree(info: ScheduleInformation): boolean | null {
  if (info.error) return null;
  return !(info.scheduleItems ?? []).some((item) => (item.status ?? 'busy') !== 'free');
}

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'exchange',
    name: 'rooms list',
    summary: 'List meeting rooms with building, floor and capacity',
    description:
      'Rooms come from Places (room mailboxes with location metadata set in Exchange). A room missing here has no ' +
      'Places entry; its mailbox still exists (`tenantctl exchange mailbox details <room>`).',
    examples: ['tenantctl exchange rooms list --building "HQ" --capacity 8'],
    args: [...FILTER_ARGS, ...LIST_ARGS],
    risk: 'read',
    permissions: { delegated: ['Place.Read.All'], application: ['Place.Read.All'] },
    async run(ctx, args) {
      const rooms = await findRooms(ctx, args);
      const limit = listLimit(args);
      return {
        items: rooms.slice(0, limit).map((r) => (ctx.full ? r : shapeRoom(r))),
        truncated: rooms.length > limit,
      };
    },
  }),

  defineCommand({
    domain: 'exchange',
    name: 'rooms free',
    summary: 'Find meeting rooms free for a whole time window',
    description:
      'Reads free/busy (getSchedule) for every matching room and returns those with nothing but free time in ' +
      '[start, end). `busy` lists rooms with tentative/busy/oof items; `unknown` lists rooms whose availability ' +
      'could not be read. Times without an offset are UTC.',
    examples: ['tenantctl exchange rooms free --start 2026-11-12T14:00:00Z --end 2026-11-12T15:00:00Z --capacity 6'],
    args: [
      { name: 'start', type: 'string', required: true, description: 'Window start, ISO date-time.' },
      { name: 'end', type: 'string', required: true, description: 'Window end, ISO date-time.' },
      ...FILTER_ARGS,
    ],
    risk: 'read',
    permissions: {
      delegated: ['Place.Read.All', 'Calendars.ReadBasic'],
      application: ['Place.Read.All', 'Calendars.ReadBasic'],
    },
    async run(ctx, args) {
      const start = parseInstant(args.start, '--start');
      const end = parseInstant(args.end, '--end');
      if (end <= start) throw new CliError('USAGE', '--end must be after --start');
      const rooms = (await findRooms(ctx, args)).filter((r) => r.emailAddress);
      const chunks: RawRoom[][] = [];
      for (let i = 0; i < rooms.length; i += SCHEDULES_PER_CALL) chunks.push(rooms.slice(i, i + SCHEDULES_PER_CALL));
      const minutes = Math.min(1440, Math.max(5, Math.round((end.getTime() - start.getTime()) / 60_000)));
      const responses = await Promise.all(
        chunks.map((chunk) => {
          // App-only has no /me; any mailbox can ask, so the first room asks for its chunk.
          const path =
            ctx.profile.authMode === 'app'
              ? `/users/${seg(chunk[0]?.emailAddress as string)}/calendar/getSchedule`
              : '/me/calendar/getSchedule';
          return ctx.graph.post<{ value?: ScheduleInformation[] }>(path, {
            schedules: chunk.map((r) => r.emailAddress),
            startTime: utc(start),
            endTime: utc(end),
            availabilityViewInterval: minutes,
          });
        }),
      );
      const byEmail = new Map<string, ScheduleInformation>();
      for (const info of responses.flatMap((r) => r?.value ?? [])) {
        if (info.scheduleId) byEmail.set(info.scheduleId.toLowerCase(), info);
      }
      const free: RawRoom[] = [];
      const busy: string[] = [];
      const unknown: { room: string; error: string }[] = [];
      for (const room of rooms) {
        const info = byEmail.get((room.emailAddress as string).toLowerCase());
        const state = info ? isScheduleFree(info) : null;
        if (state === true) free.push(room);
        else if (state === false) busy.push(room.displayName ?? (room.emailAddress as string));
        else
          unknown.push({
            room: room.emailAddress as string,
            error: info?.error?.message ?? 'no availability returned',
          });
      }
      return {
        window: { start: start.toISOString(), end: end.toISOString() },
        checked: rooms.length,
        freeCount: free.length,
        free: free.map(shapeRoom),
        busy,
        unknown,
      };
    },
  }),
];
