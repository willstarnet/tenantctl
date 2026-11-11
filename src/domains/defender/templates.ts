import { CliError } from '../../core/errors.ts';

/**
 * Curated advanced hunting queries for the questions admins ask most. Placeholders are `{param}` and are
 * replaced with KQL string literals, so values never break out of the query. Tables and columns follow the
 * Defender XDR advanced hunting schema reference (learn.microsoft.com/defender-xdr/advanced-hunting-schema-tables).
 * Time range comes from the request Timespan, not from the query.
 */
export interface HuntTemplate {
  readonly description: string;
  readonly params: Record<string, string>;
  /** Product data the table needs, so the agent can explain an empty or failing result. */
  readonly requires: string;
  readonly kql: string;
}

const DEVICE_MATCH = '(DeviceName =~ {device} or DeviceName startswith strcat({device}, ".") or DeviceId =~ {device})';

export const HUNT_TEMPLATES: Record<string, HuntTemplate> = {
  'signins-for-user': {
    description: 'Entra ID sign-ins (interactive and non-interactive) for one user, newest first.',
    params: { user: 'User principal name.' },
    requires: 'Defender for Identity or Defender XDR with Entra ID P2 (EntraIdSignInEvents).',
    kql: `EntraIdSignInEvents
| where AccountUpn =~ {user}
| project Timestamp, Application, ResourceDisplayName, ErrorCode, IPAddress, Country, City, DeviceName, OSPlatform,
  ClientAppUsed, AuthenticationRequirement, ConditionalAccessStatus, RiskLevelDuringSignIn, IsManaged, IsCompliant
| order by Timestamp desc`,
  },
  'emails-from-sender': {
    description: 'Email delivered from a sender address: recipients, subject, verdict and where it landed.',
    params: { sender: 'Sender email address (header From or envelope MailFrom).' },
    requires: 'Defender for Office 365 Plan 2 or Defender XDR (EmailEvents).',
    kql: `EmailEvents
| where SenderFromAddress =~ {sender} or SenderMailFromAddress =~ {sender}
| project Timestamp, NetworkMessageId, InternetMessageId, SenderFromAddress, RecipientEmailAddress, Subject,
  EmailDirection, DeliveryAction, LatestDeliveryLocation, ThreatTypes, DetectionMethods
| order by Timestamp desc`,
  },
  'email-by-subject': {
    description: 'Email whose subject contains a phrase (case-insensitive): who sent it, who got it, verdict.',
    params: { subject: 'Text the subject contains.' },
    requires: 'Defender for Office 365 Plan 2 or Defender XDR (EmailEvents).',
    kql: `EmailEvents
| where Subject contains {subject}
| project Timestamp, NetworkMessageId, InternetMessageId, SenderFromAddress, RecipientEmailAddress, Subject,
  EmailDirection, DeliveryAction, LatestDeliveryLocation, ThreatTypes
| order by Timestamp desc`,
  },
  'url-clicks-for-user': {
    description: 'Safe Links clicks by a user (email, Teams, Office apps), including blocked and clicked-through.',
    params: { user: 'User principal name.' },
    requires: 'Defender for Office 365 with Safe Links (UrlClickEvents).',
    kql: `UrlClickEvents
| where AccountUpn =~ {user}
| project Timestamp, Url, ActionType, Workload, IsClickedThrough, ThreatTypes, IPAddress, NetworkMessageId
| order by Timestamp desc`,
  },
  'device-processes': {
    description: 'Processes started on a device with command lines and parent process.',
    params: { device: 'Device name (short or FQDN) or Defender device id.' },
    requires: 'Defender for Endpoint Plan 2 or Defender XDR (DeviceProcessEvents).',
    kql: `DeviceProcessEvents
| where ${DEVICE_MATCH}
| project Timestamp, DeviceName, AccountName, FileName, FolderPath, ProcessCommandLine, SHA256,
  InitiatingProcessFileName, InitiatingProcessCommandLine
| order by Timestamp desc`,
  },
  'device-network': {
    description: 'Network connections made from a device: remote IP/port/URL and the process that connected.',
    params: { device: 'Device name (short or FQDN) or Defender device id.' },
    requires: 'Defender for Endpoint Plan 2 or Defender XDR (DeviceNetworkEvents).',
    kql: `DeviceNetworkEvents
| where ${DEVICE_MATCH}
| project Timestamp, DeviceName, ActionType, RemoteIP, RemotePort, RemoteUrl, Protocol, LocalIP,
  InitiatingProcessFileName, InitiatingProcessCommandLine
| order by Timestamp desc`,
  },
  'device-logons': {
    description: 'Sign-ins to a device (local, RDP, network), failures with reason, and local admin logons.',
    params: { device: 'Device name (short or FQDN) or Defender device id.' },
    requires: 'Defender for Endpoint Plan 2 or Defender XDR (DeviceLogonEvents).',
    kql: `DeviceLogonEvents
| where ${DEVICE_MATCH}
| project Timestamp, DeviceName, ActionType, LogonType, AccountDomain, AccountName, Protocol, FailureReason,
  IsLocalAdmin, RemoteDeviceName, RemoteIP
| order by Timestamp desc`,
  },
  'alerts-for-user': {
    description: 'Alerts from every Defender product that list the user as evidence.',
    params: { user: 'User principal name or Entra object id.' },
    requires: 'Defender XDR (AlertEvidence, AlertInfo).',
    kql: `AlertEvidence
| where AccountUpn =~ {user} or AccountObjectId =~ {user}
| distinct AlertId
| join kind=inner (AlertInfo) on AlertId
| summarize arg_max(Timestamp, *) by AlertId
| project Timestamp, AlertId, Title, Severity, Category, ServiceSource, DetectionSource, AttackTechniques
| order by Timestamp desc`,
  },
  'alerts-for-device': {
    description: 'Alerts from every Defender product that list the device as evidence.',
    params: { device: 'Device name (short or FQDN) or Defender device id.' },
    requires: 'Defender XDR (AlertEvidence, AlertInfo).',
    kql: `AlertEvidence
| where ${DEVICE_MATCH}
| distinct AlertId
| join kind=inner (AlertInfo) on AlertId
| summarize arg_max(Timestamp, *) by AlertId
| project Timestamp, AlertId, Title, Severity, Category, ServiceSource, DetectionSource, AttackTechniques
| order by Timestamp desc`,
  },
  'file-hash-prevalence': {
    description: 'How widespread a file is: device count, first/last seen, names, paths and devices.',
    params: { sha256: 'SHA-256 hash of the file.' },
    requires: 'Defender for Endpoint Plan 2 or Defender XDR (DeviceFileEvents).',
    kql: `DeviceFileEvents
| where SHA256 =~ {sha256}
| summarize DeviceCount = dcount(DeviceId), FirstSeen = min(Timestamp), LastSeen = max(Timestamp),
  FileNames = make_set(FileName, 10), FolderPaths = make_set(FolderPath, 10), Devices = make_set(DeviceName, 50)
  by SHA256`,
  },
  'inbox-rule-changes': {
    description: 'Inbox rules created or changed for a mailbox (a common sign of a compromised account).',
    params: { user: 'User principal name of the mailbox owner or the actor.' },
    requires: 'Defender for Cloud Apps Office 365 connector or Defender XDR (CloudAppEvents).',
    kql: `CloudAppEvents
| where ActionType in~ ("New-InboxRule", "Set-InboxRule", "Enable-InboxRule", "UpdateInboxRules")
| where tostring(RawEventData.UserId) =~ {user} or AccountDisplayName =~ {user}
| project Timestamp, ActionType, Actor = tostring(RawEventData.UserId), AccountDisplayName, IPAddress, CountryCode,
  Parameters = RawEventData.Parameters
| order by Timestamp desc`,
  },
};

/** KQL single-quoted string literal; backslash escapes keep any input inside the literal. */
export function kqlString(value: string): string {
  const escaped = value.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/\r/g, '\\r').replace(/\n/g, '\\n');
  return `'${escaped}'`;
}

/** Turn repeated `--param key=value` flags into a map; the value keeps everything after the first `=`. */
export function parseParams(pairs: readonly string[]): Record<string, string> {
  const out: Record<string, string> = {};
  for (const pair of pairs) {
    const index = pair.indexOf('=');
    if (index <= 0) throw new CliError('USAGE', `--param "${pair}" must look like name=value`);
    out[pair.slice(0, index).trim()] = pair.slice(index + 1).trim();
  }
  return out;
}

export function renderTemplate(name: string, params: Record<string, string>): string {
  const template = HUNT_TEMPLATES[name];
  if (!template) {
    throw new CliError('USAGE', `Unknown hunting template "${name}"`, {
      hint: `Run \`tenantctl defender hunt templates\`. Available: ${Object.keys(HUNT_TEMPLATES).join(', ')}.`,
    });
  }
  const missing = Object.keys(template.params).filter((key) => !params[key]);
  if (missing.length) {
    throw new CliError('USAGE', `Template "${name}" needs ${missing.map((key) => `--param ${key}=...`).join(' ')}`, {
      hint: Object.entries(template.params)
        .map(([key, description]) => `${key}: ${description}`)
        .join(' '),
    });
  }
  const unknown = Object.keys(params).filter((key) => !(key in template.params));
  if (unknown.length) {
    throw new CliError('USAGE', `Template "${name}" has no parameter ${unknown.join(', ')}`, {
      hint: `Parameters: ${Object.keys(template.params).join(', ')}.`,
    });
  }
  return template.kql.replace(/\{(\w+)\}/g, (match, key: string) =>
    key in template.params ? kqlString(params[key] as string) : match,
  );
}
