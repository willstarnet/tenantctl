import { describe, expect, test } from 'bun:test';
import { commands, htmlToText } from '../../src/domains/health/index.ts';
import { commandById, page, runCommand } from '../helpers/graphMock.ts';

describe('health overview', () => {
  test('derives degraded services and a one-line summary, ignoring operational and restored ones', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'health overview'),
      [],
      [
        {
          path: '/admin/serviceAnnouncement/healthOverviews',
          reply: page([
            { id: 'Exchange', service: 'Exchange Online', status: 'ServiceOperational' },
            { id: 'microsoftteams', service: 'Microsoft Teams', status: 'serviceDegradation' },
            { id: 'OrgLiveID', service: 'Identity Service', status: 'serviceRestored' },
            { id: 'SharePoint', service: 'SharePoint Online', status: 'investigating' },
          ]),
        },
      ],
    );
    const result = outcome.kind === 'result' ? (outcome.result as Record<string, unknown>) : {};
    expect(result.degraded).toEqual([
      { service: 'Microsoft Teams', status: 'serviceDegradation' },
      { service: 'SharePoint Online', status: 'investigating' },
    ]);
    expect(result.recentlyRestored).toEqual([{ service: 'Identity Service', status: 'serviceRestored' }]);
    expect(result.summary).toBe(
      '2 of 4 services have an active issue: Microsoft Teams (serviceDegradation), SharePoint Online (investigating)',
    );
  });

  test('all healthy reads as such', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'health overview'),
      [],
      [
        {
          path: '/admin/serviceAnnouncement/healthOverviews',
          reply: page([{ id: 'Exchange', service: 'Exchange Online', status: 'serviceOperational' }]),
        },
      ],
    );
    expect(outcome.kind === 'result' && (outcome.result as { summary: string }).summary).toBe(
      'All 1 services are healthy',
    );
  });
});

describe('health issues', () => {
  test('combines service, active, classification and days into one filter', async () => {
    const { calls } = await runCommand(
      commandById(commands, 'health issues'),
      ['--service', "Exchange O'Line", '--active', '--classification', 'incident', '--days', '3'],
      [{ path: '/admin/serviceAnnouncement/issues', reply: page([]) }],
    );
    expect(calls[0]?.query.$filter).toMatch(
      /^service eq 'Exchange O''Line' and isResolved eq false and lastModifiedDateTime ge \S+Z and classification eq 'incident'$/,
    );
  });

  test('issue returns posts newest first as plain text', async () => {
    const { outcome } = await runCommand(
      commandById(commands, 'health issue'),
      ['EX123'],
      [
        {
          path: '/admin/serviceAnnouncement/issues/EX123',
          reply: {
            body: {
              id: 'EX123',
              title: 'Delayed mail',
              service: 'Exchange Online',
              status: 'serviceDegradation',
              classification: 'incident',
              isResolved: false,
              posts: [
                {
                  createdDateTime: '2026-11-09T08:00:00Z',
                  postType: 'regular',
                  description: { contentType: 'html', content: '<p>First</p>' },
                },
                {
                  createdDateTime: '2026-11-09T10:00:00Z',
                  postType: 'regular',
                  description: {
                    contentType: 'html',
                    content: '<p>Title: Delayed&nbsp;mail</p><p>Impact:<br>Some users</p>',
                  },
                },
              ],
            },
          },
        },
      ],
    );
    const result = outcome.kind === 'result' ? (outcome.result as { posts: { text: string }[] }) : { posts: [] };
    expect(result.posts.map((p) => p.text)).toEqual(['Title: Delayed mail\nImpact:\nSome users', 'First']);
  });
});

describe('health messages', () => {
  test('filters by service, major, category and action required', async () => {
    const { calls } = await runCommand(
      commandById(commands, 'health messages'),
      ['--service', 'Microsoft Teams', '--major', '--category', 'planForChange', '--action-required'],
      [{ path: '/admin/serviceAnnouncement/messages', reply: page([]) }],
    );
    expect(calls[0]?.query.$filter).toBe(
      "services/any(s:s eq 'Microsoft Teams') and isMajorChange eq true and category eq 'planForChange' and actionRequiredByDateTime ne null",
    );
  });
});

describe('htmlToText', () => {
  test('keeps structure as line breaks and bullets, drops tags, decodes entities', () => {
    const html =
      '<div><p>We&#39;re rolling out <b>new</b> features &amp; fixes:</p><ul><li>Item&nbsp;one</li><li>Item &lt;two&gt;</li></ul>' +
      '<style>.x{color:red}</style><p>More at <a href="https://x">this link</a>&#x2e;</p></div>';
    expect(htmlToText(html)).toBe(
      "We're rolling out new features & fixes:\n\n- Item one\n- Item <two>\nMore at this link.",
    );
    expect(htmlToText(null)).toBe('');
  });
});
