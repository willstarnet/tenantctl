import { describe, expect, test } from 'bun:test';
import type { CliError } from '../../src/core/errors.ts';
import { commands, resolveTeam, summarizeCallQuality } from '../../src/domains/teams/index.ts';
import { commandById, mockContext, page, runCommand } from '../helpers/graphMock.ts';

const TEAM_ID = '11111111-1111-1111-1111-111111111111';
const USER_ID = '22222222-2222-2222-2222-222222222222';
const falcon = {
  id: TEAM_ID,
  displayName: 'Project Falcon',
  mail: 'falcon@contoso.com',
  resourceProvisioningOptions: ['Team'],
};
const jane = { id: USER_ID, displayName: 'Jane Doe', userPrincipalName: 'jane@contoso.com', mail: 'jane@contoso.com' };

const teamByName = { path: '/groups', reply: page([falcon]) };
const janeRoute = { path: '/users/jane@contoso.com', reply: { body: jane } };

describe('resolveTeam', () => {
  test('a display name filters team-enabled groups by exact name', async () => {
    const { ctx, calls } = mockContext([teamByName]);
    expect((await resolveTeam(ctx, 'Project Falcon')).id).toBe(TEAM_ID);
    expect(calls[0]?.query.$filter).toBe(
      "resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq 'Project Falcon'",
    );
  });

  test('falls back to a search, and several partial matches are AMBIGUOUS', async () => {
    const { ctx, calls } = mockContext([
      {
        path: '/groups',
        reply: (req) =>
          req.query.$search
            ? page([falcon, { ...falcon, id: 'other', displayName: 'Project Falcon Archive' }])
            : page([]),
      },
    ]);
    const error = (await resolveTeam(ctx, 'Falcon').catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('AMBIGUOUS');
    expect(error.details).toHaveLength(2);
    expect(calls[1]?.query.$search).toBe('"displayName:Falcon"');
    expect(calls[1]?.headers.ConsistencyLevel).toBe('eventual');
  });

  test('a Teams link resolves through its groupId, and a non-team group is rejected', async () => {
    const link = `https://teams.microsoft.com/l/team/19%3aabc%40thread.tacv2/conversations?groupId=${TEAM_ID}&tenantId=t`;
    const { ctx } = mockContext([{ path: `/groups/${TEAM_ID}`, reply: { body: falcon } }]);
    expect((await resolveTeam(ctx, link)).displayName).toBe('Project Falcon');
    const plain = mockContext([
      { path: `/groups/${TEAM_ID}`, reply: { body: { ...falcon, resourceProvisioningOptions: [] } } },
    ]);
    const error = (await resolveTeam(plain.ctx, TEAM_ID).catch((e: unknown) => e)) as CliError;
    expect(error.code).toBe('NOT_FOUND');
  });
});

describe('teams members', () => {
  test('members add posts an aadUserConversationMember bound to the user, as owner with --owner', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'teams members add'),
      ['Project Falcon', 'jane@contoso.com', '--owner'],
      [
        teamByName,
        janeRoute,
        {
          method: 'POST',
          path: `/teams/${TEAM_ID}/members`,
          reply: {
            status: 201,
            body: { id: 'm1', displayName: 'Jane Doe', roles: ['owner'], userId: USER_ID, email: 'jane@contoso.com' },
          },
        },
      ],
      { allowWrites: true },
    );
    const post = calls.find((c) => c.method === 'POST');
    expect(post?.body).toEqual({
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${USER_ID}')`,
    });
    expect(outcome).toEqual({
      kind: 'result',
      result: {
        team: 'Project Falcon',
        added: {
          displayName: 'Jane Doe',
          email: 'jane@contoso.com',
          role: 'owner',
          userId: USER_ID,
          membershipId: 'm1',
        },
      },
    });
  });

  test('members remove previews the membership, then deletes it once confirmed', async () => {
    const routes = [
      teamByName,
      janeRoute,
      { path: `/teams/${TEAM_ID}/members`, reply: page([{ id: 'membership-1', userId: USER_ID, roles: [] }]) },
      { method: 'DELETE', path: `/teams/${TEAM_ID}/members/membership-1`, reply: { status: 204 } },
    ];
    const remove = commandById(commands, 'teams members remove');
    const first = await runCommand(remove, ['Project Falcon', 'jane@contoso.com'], routes, { allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.method === 'DELETE')).toBe(false);
    const filterCall = first.calls.find((c) => c.path === `/teams/${TEAM_ID}/members`);
    expect(filterCall?.query.$filter).toBe(`(microsoft.graph.aadUserConversationMember/userId eq '${USER_ID}')`);
    const preview =
      first.outcome.kind === 'confirm' ? (first.outcome.preview as { currentRole: string; action: string }) : null;
    expect(preview?.currentRole).toBe('member');
    expect(preview?.action).toContain('Remove jane@contoso.com');

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(remove, ['Project Falcon', 'jane@contoso.com'], routes, {
      allowWrites: true,
      confirm: token,
    });
    expect(second.calls.filter((c) => c.method === 'DELETE').map((c) => c.path)).toEqual([
      `/teams/${TEAM_ID}/members/membership-1`,
    ]);
    expect(second.outcome).toEqual({
      kind: 'result',
      result: { team: 'Project Falcon', removed: 'jane@contoso.com', previousRole: 'member' },
    });
  });

  test('--role filters owners and guests from the roles array', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'teams members'),
      ['Project Falcon', '--role', 'guest'],
      [
        teamByName,
        {
          path: `/teams/${TEAM_ID}/members`,
          reply: page([
            { id: 'a', displayName: 'Owner', roles: ['owner'] },
            { id: 'b', displayName: 'Member', roles: [] },
            { id: 'c', displayName: 'Guest', roles: ['guest'] },
          ]),
        },
      ],
    );
    const result =
      outcome.kind === 'result' ? (outcome.result as { items: { displayName: string; role: string }[] }) : null;
    expect(result?.items.map((m) => [m.displayName, m.role])).toEqual([['Guest', 'guest']]);
  });
});

describe('teams create', () => {
  test('binds the standard template and adds the owner as a member', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'teams create'),
      ['--name', 'Falcon', '--owner', 'jane@contoso.com', '--visibility', 'public'],
      [
        janeRoute,
        {
          method: 'POST',
          path: '/teams',
          reply: {
            status: 202,
            headers: {
              Location: `/teams('${TEAM_ID}')/operations('op-1')`,
              'Content-Location': `/teams('${TEAM_ID}')`,
            },
          },
        },
      ],
      { allowWrites: true },
    );
    const post = calls.find((c) => c.method === 'POST');
    expect(post?.body).toEqual({
      'template@odata.bind': "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
      displayName: 'Falcon',
      visibility: 'public',
      members: [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: ['owner'],
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${USER_ID}')`,
        },
      ],
    });
    expect(outcome.kind === 'result' && outcome.result).toMatchObject({
      status: 'accepted',
      teamId: TEAM_ID,
      operationId: 'op-1',
      next: `tenantctl graph get "/teams('${TEAM_ID}')/operations('op-1')"`,
    });
  });

  test('archive is destructive, previews the archived state, then returns the operation id', async () => {
    const routes = [
      teamByName,
      { path: `/teams/${TEAM_ID}`, reply: { body: { id: TEAM_ID, isArchived: false, summary: { membersCount: 12 } } } },
      {
        method: 'POST',
        path: `/teams/${TEAM_ID}/archive`,
        reply: { status: 202, headers: { Location: `/teams('${TEAM_ID}')/operations('op-2')` } },
      },
    ];
    const archive = commandById(commands, 'teams archive');
    const first = await runCommand(archive, ['Project Falcon'], routes, { allowWrites: true });
    expect(first.outcome.kind).toBe('confirm');
    expect(first.calls.some((c) => c.method === 'POST')).toBe(false);
    const preview =
      first.outcome.kind === 'confirm'
        ? (first.outcome.preview as { team: { isArchived: boolean; members: number } })
        : null;
    expect(preview?.team).toMatchObject({ isArchived: false, members: 12 });

    const token = first.outcome.kind === 'confirm' ? first.outcome.token : '';
    const second = await runCommand(archive, ['Project Falcon'], routes, { allowWrites: true, confirm: token });
    expect(second.calls.filter((c) => c.method === 'POST').map((c) => c.path)).toEqual([`/teams/${TEAM_ID}/archive`]);
    expect(second.outcome.kind === 'result' && second.outcome.result).toMatchObject({
      teamId: TEAM_ID,
      operationId: 'op-2',
    });
  });
});

describe('teams presence', () => {
  test('resolves every user and asks getPresencesByUserId once', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'teams presence'),
      ['jane@contoso.com'],
      [
        janeRoute,
        {
          method: 'POST',
          path: '/communications/getPresencesByUserId',
          reply: {
            body: {
              value: [
                {
                  id: USER_ID,
                  availability: 'Busy',
                  activity: 'InACall',
                  statusMessage: { message: { content: 'Back at 3', contentType: 'text' } },
                  outOfOfficeSettings: { isOutOfOffice: false },
                },
              ],
            },
          },
        },
      ],
    );
    expect(calls.at(-1)?.body).toEqual({ ids: [USER_ID] });
    expect(outcome).toEqual({
      kind: 'result',
      result: [
        {
          user: 'jane@contoso.com',
          availability: 'Busy',
          activity: 'InACall',
          statusMessage: 'Back at 3',
          isOutOfOffice: false,
        },
      ],
    });
  });
});

describe('teams calls', () => {
  const user = (id: string, name: string) => ({
    '@odata.type': '#microsoft.graph.callRecords.participantEndpoint',
    associatedIdentity: { id, displayName: name },
  });
  const service = { '@odata.type': '#microsoft.graph.callRecords.serviceEndpoint' };
  const sessions = [
    {
      id: 's1',
      segments: [
        {
          id: 'seg1',
          caller: user('u-jane', 'Jane Doe'),
          callee: service,
          media: [
            {
              label: 'main-audio',
              callerNetwork: { connectionType: 'wifi', wifiSignalStrength: 35, receivedQualityEventRatio: 0.4 },
              streams: [
                {
                  streamDirection: 'callerToCallee',
                  averagePacketLossRate: 0.15,
                  maxPacketLossRate: 0.3,
                  averageJitter: 'PT0.045S',
                  maxJitter: 'PT0.12S',
                  averageRoundTripTime: 'PT0.08S',
                },
                {
                  streamDirection: 'calleeToCaller',
                  averagePacketLossRate: 0.01,
                  averageJitter: 'PT0.01S',
                  averageRoundTripTime: 'PT0.08S',
                },
              ],
            },
          ],
        },
      ],
    },
    {
      id: 's2',
      failureInfo: { stage: 'midcall', reason: 'unexpectedClientError' },
      segments: [
        {
          id: 'seg2',
          caller: user('u-bob', 'Bob Smith'),
          callee: service,
          media: [
            {
              callerNetwork: { connectionType: 'wired' },
              streams: [
                {
                  streamDirection: 'callerToCallee',
                  averagePacketLossRate: 0.001,
                  averageJitter: 'PT0.005S',
                  averageRoundTripTime: 'PT0.03S',
                },
              ],
            },
          ],
        },
      ],
    },
  ];

  test('summarizeCallQuality flags the participant whose streams cross CQD thresholds', () => {
    const { participants, failures } = summarizeCallQuality(sessions as never);
    const janeRow = participants.find((p) => p.participant === 'Jane Doe');
    expect(janeRow).toMatchObject({
      verdict: 'poor',
      streams: 2,
      poorStreams: 1,
      poorReasons: ['packetLoss', 'jitter'],
      maxPacketLossRate: 0.3,
      maxJitterMs: 120,
      maxRoundTripMs: 80,
      connectionTypes: ['wifi'],
      minWifiSignalStrength: 35,
      networkEventRatios: { receivedQuality: 0.4, sentQuality: null, delay: null, bandwidthLow: null },
    });
    expect(participants.find((p) => p.participant === 'Bob Smith')).toMatchObject({
      verdict: 'ok',
      poorStreams: 0,
      connectionTypes: ['wired'],
    });
    expect(participants).toHaveLength(2);
    expect(failures).toEqual([{ session: 's2', stage: 'midcall', reason: 'unexpectedClientError' }]);
  });

  test('calls get expands sessions and segments and names the poor participants', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'teams calls get'),
      ['call-1'],
      [
        {
          path: '/communications/callRecords/call-1',
          reply: {
            body: {
              id: 'call-1',
              type: 'groupCall',
              modalities: ['audio'],
              startDateTime: '2026-11-09T10:00:00Z',
              endDateTime: '2026-11-09T10:30:00Z',
              organizer_v2: { id: 'u-jane', identity: { user: { id: 'u-jane', displayName: 'Jane Doe' } } },
              sessions,
            },
          },
        },
        {
          path: '/communications/callRecords/call-1/participants_v2',
          reply: page([
            { id: 'u-jane', identity: { user: { displayName: 'Jane Doe' } } },
            { id: 'u-bob', identity: { user: { displayName: 'Bob Smith' } } },
          ]),
        },
      ],
    );
    expect(calls[0]?.query.$expand).toBe('sessions($expand=segments)');
    const result = outcome.kind === 'result' ? (outcome.result as Record<string, unknown>) : {};
    expect(result).toMatchObject({
      organizer: 'Jane Doe',
      durationSeconds: 1800,
      participants: ['Jane Doe', 'Bob Smith'],
      summary: 'Poor quality for Jane Doe',
    });
  });

  test('calls list filters on start time and participant id without $top', async () => {
    const { calls } = await runCommand(
      commandById(commands, 'teams calls list'),
      ['--user', 'jane@contoso.com', '--days', '2'],
      [janeRoute, { path: '/communications/callRecords', reply: page([]) }],
    );
    const list = calls.find((c) => c.path === '/communications/callRecords');
    expect(list?.query.$filter).toMatch(
      new RegExp(`^startDateTime ge \\S+Z and participants_v2/any\\(p:p/id eq '${USER_ID}'\\)$`),
    );
    expect(list?.query.$top).toBeUndefined();
  });
});

describe('teams activity', () => {
  test('picks the user row from the usage report and converts counts', async () => {
    const csv =
      'Report Refresh Date,User Id,User Principal Name,Last Activity Date,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Meetings Organized Count,Is Licensed\n' +
      '2026-11-08,u1,jane@contoso.com,2026-11-07,4,10,2,3,1,Yes\n' +
      '2026-11-08,u2,bob@contoso.com,,0,0,0,0,0,No\n';
    const { outcome, calls } = await runCommand(
      commandById(commands, 'teams activity'),
      ['--user', 'Jane@Contoso.com'],
      [{ path: "/reports/getTeamsUserActivityUserDetail(period='D7')", reply: { body: csv } }],
    );
    expect(calls[0]?.path).toBe("/reports/getTeamsUserActivityUserDetail(period='D7')");
    const result = outcome.kind === 'result' ? (outcome.result as { items: unknown[] }) : { items: [] };
    expect(result.items).toEqual([
      {
        user: 'jane@contoso.com',
        lastActivityDate: '2026-11-07',
        teamChatMessages: 4,
        privateChatMessages: 10,
        calls: 2,
        meetings: 3,
        meetingsOrganized: 1,
        isLicensed: true,
        reportRefreshDate: '2026-11-08',
      },
    ]);
  });
});
