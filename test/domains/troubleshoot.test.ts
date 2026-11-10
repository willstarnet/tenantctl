import { describe, expect, test } from 'bun:test';
import { commands } from '../../src/domains/troubleshoot/index.ts';
import { analyzeSignins, type RawSignIn, signinFindings } from '../../src/domains/troubleshoot/signins.ts';
import type { ExoRunner } from '../../src/core/exo.ts';
import { commandById, fakeExo, graphError, type MockRoute, page, runCommand } from '../helpers/graphMock.ts';

const USER_ID = '11111111-1111-1111-1111-111111111111';
const UPN = 'jane@contoso.com';
const DEVICE_ID = '22222222-2222-2222-2222-222222222222';
const ENTRA_DEVICE_ID = '33333333-3333-3333-3333-333333333333';
const E3 = '05e9a617-0261-4cee-bb44-138d3ef5d965';
const hoursAgo = (h: number) => new Date(Date.now() - h * 3_600_000).toISOString();

interface Report {
  verdict: string;
  findings: { severity: string; area: string; issue: string; next?: string; detail?: string }[];
  unavailable: { section: string; code: string; message: string }[];
}

async function report(
  id: string,
  argv: string[],
  routes: MockRoute[],
  exo?: ExoRunner,
): Promise<Report & Record<string, unknown>> {
  const { outcome } = await runCommand(commandById(commands, id), argv, routes, { exo });
  if (outcome.kind !== 'result') throw new Error('expected a result');
  return outcome.result as Report & Record<string, unknown>;
}

function signin(hours: number, code: number, extra: Partial<RawSignIn> = {}): RawSignIn {
  return {
    id: `s-${hours}-${code}`,
    createdDateTime: hoursAgo(hours),
    userPrincipalName: UPN,
    appDisplayName: 'Office 365 Exchange Online',
    ipAddress: '203.0.113.5',
    status: { errorCode: code, failureReason: code ? 'failed' : null },
    ...extra,
  };
}

describe('sign-in analysis', () => {
  test('a failure followed by a success is history; one after the last success is still failing', () => {
    const analysis = analyzeSignins([signin(5, 50126), signin(3, 0), signin(1, 53003)], 7, false);
    expect(analysis.failures).toBe(2);
    const findings = signinFindings(analysis, UPN);
    expect(findings.find((f) => f.issue.includes('53003'))).toMatchObject({
      severity: 'high',
      next: `m365 troubleshoot signin ${UPN}`,
    });
    expect(findings.find((f) => f.issue.includes('50126'))).toMatchObject({ severity: 'low' });
    expect(findings.find((f) => f.issue.includes('50126'))?.issue).toContain('succeeded since');
  });

  test('benign interrupts produce no finding; device codes point at the device', () => {
    const analysis = analyzeSignins(
      [signin(2, 50140), signin(1, 53000, { deviceDetail: { displayName: 'LAPTOP-7' } })],
      7,
      false,
    );
    const findings = signinFindings(analysis, UPN);
    expect(findings).toHaveLength(1);
    expect(findings[0]?.next).toBe('m365 troubleshoot device "LAPTOP-7"');
  });

  test('wrong passwords from many IPs are flagged as possible password spray', () => {
    const attempts = Array.from({ length: 12 }, (_, i) => signin(i + 1, 50126, { ipAddress: `198.51.100.${i % 4}` }));
    const findings = signinFindings(analyzeSignins([...attempts, signin(20, 0)], 7, false), UPN);
    expect(findings[0]).toMatchObject({ severity: 'high', next: `m365 entra risk-detections --user ${UPN}` });
    expect(findings[0]?.detail).toContain('password spray');
  });

  test('failed Conditional Access policies are named with their required controls', () => {
    const blocked = signin(1, 53003, {
      appliedConditionalAccessPolicies: [
        {
          displayName: 'Require compliant device',
          result: 'failure',
          enforcedGrantControls: ['RequireCompliantDevice'],
        },
        { displayName: 'Require MFA', result: 'success', enforcedGrantControls: ['Mfa'] },
      ],
    });
    const group = analyzeSignins([blocked], 7, false).failureGroups[0];
    expect(group?.failedPolicies).toEqual(['Require compliant device (requires RequireCompliantDevice)']);
  });
});

describe('troubleshoot user', () => {
  const user = {
    id: USER_ID,
    displayName: 'Jane Doe',
    userPrincipalName: UPN,
    mail: UPN,
    accountEnabled: false,
    userType: 'Member',
    assignedLicenses: [{ skuId: E3, disabledPlans: [] }],
    licenseAssignmentStates: [{ skuId: E3, state: 'Error', error: 'CountViolation', assignedByGroup: null }],
  };
  const routes: MockRoute[] = [
    { path: `/users/${UPN}`, reply: { body: user } },
    { path: `/users/${USER_ID}`, reply: { body: user } },
    { path: `/users/${USER_ID}/manager`, reply: graphError(404, 'Request_ResourceNotFound', 'no manager') },
    {
      path: '/subscribedSkus',
      reply: page([{ skuId: E3, skuPartNumber: 'ENTERPRISEPACK', consumedUnits: 10, servicePlans: [] }]),
    },
    {
      path: `/users/${USER_ID}/authentication/methods`,
      reply: page([{ '@odata.type': '#microsoft.graph.passwordAuthenticationMethod', id: 'p' }]),
    },
    {
      path: `/reports/authenticationMethods/userRegistrationDetails/${USER_ID}`,
      reply: { body: { isMfaRegistered: false } },
    },
    {
      path: '/auditLogs/signIns',
      reply: page([
        signin(1, 53003, { appliedConditionalAccessPolicies: [{ displayName: 'Block unmanaged', result: 'failure' }] }),
      ]),
    },
    { path: `/identityProtection/riskyUsers/${USER_ID}`, reply: graphError(403, 'Forbidden', 'needs P2') },
    {
      path: '/deviceManagement/managedDevices',
      reply: page([
        {
          id: DEVICE_ID,
          deviceName: 'LAPTOP-7',
          userId: USER_ID,
          complianceState: 'noncompliant',
          lastSyncDateTime: hoursAgo(2),
        },
      ]),
    },
    { path: `/users/${USER_ID}/mailboxSettings`, reply: graphError(403, 'ErrorAccessDenied', 'Access is denied') },
    { path: '/domains', reply: page([{ id: 'contoso.com', isVerified: true }]) },
    {
      path: `/users/${USER_ID}/mailFolders/inbox/messageRules`,
      reply: page([
        {
          id: 'r1',
          displayName: '.',
          isEnabled: true,
          sequence: 1,
          actions: { forwardTo: [{ emailAddress: { address: 'drop@evil.example' } }] },
        },
      ]),
    },
  ];

  test('ranks account, license, sign-in, mailbox and device problems with next commands', async () => {
    const result = await report('troubleshoot user', [UPN], routes);
    const issues = result.findings.map((f) => `${f.severity}:${f.area}`);
    expect(issues.slice(0, 4).every((i) => i.startsWith('high:'))).toBe(true);
    expect(result.findings).toContainEqual(
      expect.objectContaining({
        severity: 'high',
        issue: 'Account is disabled',
        next: `m365 entra users enable ${UPN}`,
      }),
    );
    expect(result.findings).toContainEqual(
      expect.objectContaining({ severity: 'high', issue: 'License ENTERPRISEPACK failed to assign: CountViolation' }),
    );
    expect(result.findings).toContainEqual(
      expect.objectContaining({ area: 'mailbox', severity: 'high', next: `m365 exchange rules list ${UPN}` }),
    );
    expect(result.findings.find((f) => f.area === 'sign-in')?.detail).toContain('Blocked by: Block unmanaged');
    expect(result.findings).toContainEqual(expect.objectContaining({ area: 'mfa', severity: 'medium' }));
    expect(result.findings).toContainEqual(
      expect.objectContaining({ area: 'devices', next: 'm365 troubleshoot device "LAPTOP-7"' }),
    );
    expect(result.verdict).toStartWith(
      `${result.findings.filter((f) => f.severity === 'high' || f.severity === 'medium').length} problems for ${UPN}: Account is disabled`,
    );
  });

  test('sections the caller cannot read are reported, not fatal', async () => {
    const noExchange = fakeExo(() => {
      throw new Error('UnAuthorized');
    });
    const result = await report('troubleshoot user', [UPN], routes, noExchange);
    expect(result.unavailable.map((u) => `${u.section}:${u.code}`).sort()).toEqual([
      'autoReply:FORBIDDEN',
      'mailboxForwarding:GRAPH_ERROR',
      'risk:FORBIDDEN',
    ]);
    expect(result.verdict).toContain('could not check: risk, autoReply, mailboxForwarding');
  });

  test('mailbox-level forwarding outside the company is a high finding, checked through Exchange', async () => {
    const exo = fakeExo((step) =>
      step.cmdlet === 'Get-Mailbox'
        ? [
            {
              ForwardingSmtpAddress: 'smtp:drop@evil.example',
              ForwardingAddress: null,
              DeliverToMailboxAndForward: false,
            },
          ]
        : [],
    );
    const result = await report('troubleshoot user', [UPN], routes, exo);
    expect(exo.steps[0]).toMatchObject({ cmdlet: 'Get-Mailbox', params: { Identity: UPN } });
    expect(result.findings).toContainEqual(
      expect.objectContaining({
        severity: 'high',
        issue: 'Mailbox forwards all mail to drop@evil.example (outside the organization)',
        next: `m365 exchange forwarding ${UPN}`,
      }),
    );
  });
});

describe('troubleshoot device', () => {
  const device = {
    id: DEVICE_ID,
    deviceName: 'LAPTOP-7',
    userId: USER_ID,
    userPrincipalName: UPN,
    operatingSystem: 'Windows',
    complianceState: 'noncompliant',
    lastSyncDateTime: hoursAgo(24 * 12),
    isEncrypted: false,
    azureADDeviceId: ENTRA_DEVICE_ID,
    freeStorageSpaceInBytes: 5 * 1024 ** 3,
    totalStorageSpaceInBytes: 256 * 1024 ** 3,
    deviceActionResults: [{ actionName: 'wipe', actionState: 'pending', startDateTime: hoursAgo(1) }],
  };
  const routes: MockRoute[] = [
    { path: `/deviceManagement/managedDevices/${DEVICE_ID}`, reply: { body: device } },
    {
      path: `/deviceManagement/managedDevices/${DEVICE_ID}/deviceCompliancePolicyStates`,
      reply: page([
        {
          id: 'p1',
          displayName: 'Windows baseline',
          state: 'nonCompliant',
          settingStates: [
            {
              setting: 'Windows10CompliancePolicy.bitLockerEnabled',
              settingName: 'BitLocker',
              state: 'nonCompliant',
              currentValue: 'false',
            },
          ],
        },
      ]),
    },
    {
      path: `/deviceManagement/managedDevices/${DEVICE_ID}/deviceConfigurationStates`,
      reply: page([{ id: 'c1', displayName: 'Wi-Fi', state: 'error', settingStates: [] }]),
    },
    {
      path: '/devices',
      reply: page([{ id: 'obj', deviceId: ENTRA_DEVICE_ID, displayName: 'LAPTOP-7', accountEnabled: false }]),
    },
    {
      path: '/auditLogs/signIns',
      reply: page([
        signin(1, 53000, { deviceDetail: { deviceId: ENTRA_DEVICE_ID, displayName: 'LAPTOP-7' } }),
        signin(2, 50126, { deviceDetail: { deviceId: 'some-other-device' } }),
      ]),
    },
  ];

  test('explains non-compliance and every device-side blocker', async () => {
    const result = await report('troubleshoot device', [DEVICE_ID], routes);
    const byArea = Object.fromEntries(result.findings.map((f) => [f.area, f]));
    expect(byArea.compliance).toMatchObject({ severity: 'high', next: 'm365 intune devices compliance "LAPTOP-7"' });
    expect(byArea.compliance?.detail).toContain('"Windows baseline" fails BitLocker (current value false)');
    expect(byArea.sync).toMatchObject({ severity: 'medium', next: 'm365 intune devices sync "LAPTOP-7"' });
    expect(byArea.entra).toMatchObject({
      severity: 'high',
      issue: 'Device is disabled in Entra ID',
      next: 'm365 entra devices enable obj',
    });
    expect(byArea.configuration?.detail).toBe('Wi-Fi (error)');
    expect(byArea.actions).toMatchObject({ severity: 'high' });
    expect(byArea.storage).toMatchObject({ severity: 'low' });
    expect(result.findings.filter((f) => f.area === 'sign-in').map((f) => f.issue)).toEqual([
      '1 failed sign-in from this device: 53000 DeviceNotCompliant',
    ]);
  });

  test('a device without an Entra registration is called out and no Entra lookup is made', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'troubleshoot device'),
      [DEVICE_ID],
      [
        ...routes.filter((r) => r.path !== `/deviceManagement/managedDevices/${DEVICE_ID}`),
        {
          path: `/deviceManagement/managedDevices/${DEVICE_ID}`,
          reply: { body: { ...device, azureADDeviceId: '00000000-0000-0000-0000-000000000000' } },
        },
      ],
    );
    const result = (outcome as { result: Report }).result;
    expect(result.findings).toContainEqual(
      expect.objectContaining({ area: 'entra', issue: 'Not registered in Entra ID' }),
    );
    expect(calls.some((c) => c.path === '/devices')).toBe(false);
  });
});

describe('troubleshoot mail', () => {
  const TRACE = '/admin/exchange/tracing/messageTraces';
  const rules: MockRoute[] = [
    { path: `/users/${UPN}`, reply: { body: { id: USER_ID, userPrincipalName: UPN, displayName: 'Jane' } } },
    { path: '/domains', reply: page([{ id: 'contoso.com', isVerified: true }]) },
    {
      path: `/users/${USER_ID}/mailFolders/inbox/messageRules`,
      reply: page([
        {
          id: 'r1',
          displayName: 'Vendors',
          isEnabled: true,
          conditions: { senderContains: ['vendor'] },
          actions: { delete: true },
        },
      ]),
    },
  ];

  test('filters the trace by recipient and sender, and explains a quarantine from its last event', async () => {
    const { outcome, calls } = await runCommand(
      commandById(commands, 'troubleshoot mail'),
      [UPN, '--sender', 'billing@vendor.com'],
      [
        ...rules,
        {
          path: TRACE,
          reply: page([
            {
              id: 't1',
              senderAddress: 'billing@vendor.com',
              recipientAddress: UPN,
              subject: 'Invoice',
              status: 'quarantined',
              receivedDateTime: hoursAgo(3),
            },
          ]),
        },
        {
          path: `${TRACE}/t1/getDetailsByRecipient(recipientAddress='${UPN}')`,
          reply: page([
            { dateTime: hoursAgo(3), event: 'Receive', description: 'Received' },
            { dateTime: hoursAgo(3), event: 'Quarantine', description: 'Message quarantined as phish' },
          ]),
        },
      ],
    );
    const filter = calls.find((c) => c.path === TRACE)?.query.$filter ?? '';
    expect(filter).toContain(`recipientAddress eq '${UPN}'`);
    expect(filter).toContain("senderAddress eq 'billing@vendor.com'");
    const result = (outcome as { result: Report & { statusCounts: Record<string, number> } }).result;
    expect(result.statusCounts).toEqual({ quarantined: 1 });
    expect(result.findings[0]).toMatchObject({ severity: 'high', area: 'delivery' });
    expect(result.findings[0]?.detail).toContain('Message quarantined as phish');
  });

  test('delivered mail points at the inbox rule that deletes it', async () => {
    const result = await report(
      'troubleshoot mail',
      [UPN],
      [
        ...rules,
        {
          path: TRACE,
          reply: page([
            {
              id: 't2',
              senderAddress: 'a@vendor.com',
              recipientAddress: UPN,
              status: 'delivered',
              receivedDateTime: hoursAgo(1),
            },
          ]),
        },
      ],
    );
    expect(result.findings).toContainEqual(
      expect.objectContaining({
        area: 'inbox rules',
        issue: 'Inbox rule "Vendors" deletes matching mail',
        next: `m365 exchange rules list ${UPN}`,
      }),
    );
  });

  test('no trace at all means the message never reached Exchange Online', async () => {
    const result = await report('troubleshoot mail', [UPN], [...rules, { path: TRACE, reply: page([]) }]);
    expect(result.findings[0]?.issue).toContain('reached Exchange Online');
  });
});

describe('troubleshoot tenant', () => {
  test('flags outages, stalled directory sync, licence and secret problems; P2-only sections are unavailable', async () => {
    const result = await report(
      'troubleshoot tenant',
      [],
      [
        {
          path: '/admin/serviceAnnouncement/healthOverviews',
          reply: page([{ id: 'Exchange', service: 'Exchange Online', status: 'serviceDegradation' }]),
        },
        {
          path: '/organization',
          reply: page([
            { displayName: 'Contoso', onPremisesSyncEnabled: true, onPremisesLastSyncDateTime: hoursAgo(26) },
          ]),
        },
        {
          path: '/subscribedSkus',
          reply: page([
            {
              skuId: E3,
              skuPartNumber: 'ENTERPRISEPACK',
              appliesTo: 'User',
              consumedUnits: 100,
              prepaidUnits: { enabled: 100, warning: 0, suspended: 0 },
            },
          ]),
        },
        {
          path: '/applications',
          reply: page([
            {
              id: 'a',
              appId: 'x',
              displayName: 'Payroll sync',
              passwordCredentials: [{ keyId: 'k', endDateTime: hoursAgo(48) }],
              keyCredentials: [],
            },
          ]),
        },
        {
          path: '/deviceManagement/deviceCompliancePolicyDeviceStateSummary',
          reply: { body: { compliantDeviceCount: 90, nonCompliantDeviceCount: 10 } },
        },
        { path: '/identityProtection/riskyUsers', reply: graphError(403, 'Forbidden', 'Entra ID P2 required') },
        { path: '/security/incidents', reply: page([]) },
        { path: '/security/secureScores', reply: page([{ currentScore: 40, maxScore: 100 }]) },
      ],
    );
    const issues = result.findings.map((f) => f.issue);
    expect(issues).toContain('Microsoft reports an issue with Exchange Online (serviceDegradation)');
    expect(issues).toContain('Entra Connect has not synced for 26 hours');
    expect(issues).toContain('ENTERPRISEPACK: 0 of 100 licenses left');
    expect(issues).toContain('1 app secret/certificate has expired');
    expect(issues).toContain('10 of 100 devices are non-compliant or in error');
    expect(issues).toContain('Secure Score is 40%');
    expect(result.unavailable).toEqual([{ section: 'riskyUsers', code: 'FORBIDDEN', message: expect.any(String) }]);
  });
});
