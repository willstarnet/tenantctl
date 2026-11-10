---
name: m365
description: Administer and troubleshoot a Microsoft 365 tenant (Entra ID users/groups/sign-ins/MFA/Conditional Access, Intune devices and compliance, Defender incidents/alerts/hunting, Exchange mailboxes/rules/message trace, SharePoint/OneDrive sharing, Teams, service health) with the `m365` CLI. Use whenever the user asks about M365 users, devices, email delivery, sign-in problems, security alerts, file sharing or tenant health.
---

# m365: Microsoft 365 administration

`m365` is a CLI over Microsoft Graph. Run it through the shell and read its JSON output.

## Ground rules

1. **Always pass `--json`.** Every result is `{ "ok": true, "command", "data", "meta"? }` or
   `{ "ok": false, "error": { "code", "message", "hint", "details"? } }`. Follow the error's `hint`; it tells
   you what to do next.
2. **Use the names the user gives you.** Commands accept a UPN, email, display name, device name, serial
   number, site URL or team name. Don't look up ids first.
3. **Exit code 8 / `AMBIGUOUS`:** `error.details` lists the candidates. Ask the user which one they mean, then
   pass its `id`. Never pick one yourself.
4. **Start with `troubleshoot` for any problem report.** It checks every relevant area at once and returns a
   `verdict`, ranked `findings` (each with a `next` command), and `unavailable` sections. Tell the user the
   verdict and the high and medium findings in plain language, then offer the `next` steps.
5. **Don't guess commands.** Use `m365 <area> --help`, `m365 <area> <command> --help`, or `m365 tools --json`
   (every command with a JSON Schema for its arguments).
6. **Lists stop at `--top` (default 50).** When `meta.truncated` is true, say so, and use `--all` or narrower
   filters if the user needs everything.
7. **`--full` returns raw Graph objects.** Use it only when the compact view lacks a field you need. `--fields
   a,b.c` trims the output.

## Changing things: the safety protocol

The tenant is read-only until writes are enabled. Never enable writes yourself, and never change `confirmMode`.

- **Exit 7 / `WRITES_DISABLED`:** tell the user the change needs writes enabled and ask them to run
  `export M365_ALLOW_WRITES=1` (this session) or `m365 config set allowWrites true` (this profile). Meanwhile
  `--dry-run` shows what would happen.
- **Exit 3 / `confirmationRequired` (destructive commands):** nothing has changed yet. The response has a
  `preview` (the resolved target, its current state, and an `action` sentence) and `confirm.rerun`. You MUST:
  1. Show the user the preview in plain language: what will happen, to which user/device/item, and anything
     irreversible (a wipe erases the device; a delete is soft for 30 days for users only).
  2. Get explicit approval in the chat for this specific action. Approval for one target never covers another.
  3. Run `confirm.rerun` exactly, and tell the user to watch for an approval dialog (or a terminal prompt)
     from `m365`. They approve the change there; the command waits up to 5 minutes. The token expires after
     10 minutes and works only with the same arguments.
- **Exit 10 / `DENIED`:** the person did not approve, or `m365` could not ask them, for example because there
  was no desktop and no terminal. Nothing changed. Don't retry or look for a workaround; ask the user what they
  want.
- Commands marked `write` (sync a device, add a group member, set an auto-reply) run as soon as writes are on.
  Still say what you're about to do before you run them.
- Every write attempt is logged locally: `m365 config audit-log --json`.

## Sign-in and permission errors

| Error                                             | Do this                                                                                                                                                                                                                                                                                                                                                                                                |
| ------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `AUTH_REQUIRED` (exit 4)                          | Ask the user to run the login command the `hint` names and complete the device-code prompt, or run it yourself and relay the stderr message word for word: open https://login.microsoft.com/device and enter the code. No notification is sent. `m365 auth login` covers Graph; `--service exchange` covers Exchange admin commands; `--service defender` covers Defender for Endpoint device actions. |
| `FORBIDDEN` (exit 5)                              | The `hint` names the missing permission or role. Tell the user. Don't retry.                                                                                                                                                                                                                                                                                                                           |
| `AUTH_FAILED` with an AADSTS code                 | `m365 entra signin-error <code> --json` explains it.                                                                                                                                                                                                                                                                                                                                                   |
| `THROTTLED` (exit 9)                              | Wait a minute, then retry with a smaller `--top`.                                                                                                                                                                                                                                                                                                                                                      |
| An `unavailable` section in a troubleshoot result | Mention what couldn't be checked and why (licence or permission). Use the rest of the report.                                                                                                                                                                                                                                                                                                          |

`m365 auth status --json` shows the signed-in identity, tenant, admin roles, granted permissions and whether
writes are on. Check it when something unexpectedly fails.

## Playbooks

**"X can't sign in" / "X is locked out"**

1. `m365 troubleshoot user <x> --json`. The common causes:
   - disabled account: `entra users enable`
   - still-failing sign-in codes
   - Conditional Access block: `troubleshoot signin <x>` names the policy
   - non-compliant device: `troubleshoot device`
   - no MFA method: `entra users tap` issues a Temporary Access Pass
   - high risk: `entra risk-detections`
2. Password reset: `m365 entra users reset-password <x>` (destructive). It returns a temporary password once;
   give it to the user through a secure channel and don't repeat it later.

**"X got a new phone" / MFA reset**
`m365 entra users auth-methods <x>`, then `m365 entra users auth-method remove <x> <methodId>` for the old
device, and `m365 entra users tap <x>` so they can register the new one.

**"My laptop/phone can't get to email" / "device not compliant"**

1. `m365 troubleshoot device <name or serial>`. `findings[0].detail` names the failing setting.
2. Fixes: the user turns on the setting (e.g. BitLocker or an OS update), then `m365 intune devices sync
   <device>` (write).
3. Stale check-in: the device must be powered on and online to sync.

**Lost or stolen device**

1. `m365 intune devices get <device>` to confirm owner, serial and last check-in with the user.
2. Then offer, from least to most drastic:
   - `intune devices lock` (mobile)
   - `defender machines isolate <device> --comment "<why>"`: cuts it off the network but keeps Defender
     connected. Needs Defender for Endpoint and its own sign-in; see Sign-in.
   - `entra users revoke-sessions <owner>`
   - `entra devices disable <device>`
   - `intune devices retire`: removes company data
   - `intune devices wipe`: factory reset
3. Each is destructive: preview, chat approval, then the person approves the dialog. Retrieve the BitLocker key
   or LAPS password only when the user asks (`--reveal` is audit-logged by Microsoft).

**"I didn't get an email from Y"**

1. `m365 troubleshoot mail <recipient> --sender <y>`.
2. If it was delivered, look for inbox rules, or check Junk with `m365 exchange messages search <recipient>
   --from <y>`.
3. If it was quarantined: `m365 exchange quarantine list --recipient <recipient>`, then
   `m365 exchange quarantine release <identity> --user <recipient>`.
4. For mail-flow rules: `m365 exchange transport-rules`.

**Suspected compromised account**

1. `m365 troubleshoot user <x>`. Look for:
   - external forwarding, in inbox rules **and** on the mailbox itself
   - risk
   - odd sign-in locations
2. Then gather evidence:
   - `m365 entra signins --user <x> --days 7`
   - `m365 defender alerts for-user <x>`
   - `m365 purview audit mailbox-access <x> --days 7` shows what was read and by whom
   - `m365 entra apps consents <app>` for any unfamiliar app the user consented to
3. With approval, contain:
   - `entra users disable`
   - `entra users revoke-sessions`
   - `exchange forwarding clear`
   - `exchange rules delete` for malicious rules
   - `entra apps revoke-consent` (plus `entra apps disable`) for a malicious app
   - `entra users reset-password`
   - `entra risky-users confirm-compromised`

**Phishing campaign**

1. `m365 defender hunt --template emails-from-sender --param sender=<addr>` (or `email-by-subject`) to find
   recipients.
2. Then `m365 defender hunt --template url-clicks-for-user` per recipient.
3. Purge with `m365 exchange messages delete <user> <messageId>` (destructive), each with approval.

**Onboarding**

1. `m365 entra users create --display-name ... --upn ... --usage-location US --manager ...`. For several people,
   use `m365 entra users bulk-create <file.csv>` (run with `--dry-run` first to check every row).
2. `m365 entra licenses` to check stock, then `m365 entra users license add <upn> --sku <sku>`.
3. `m365 entra groups add-member <group> <upn>` (or `entra groups add-members <group> --from-csv <file>`), and
   `m365 teams members add <team> <upn>`.
4. Guests and contractors: `m365 entra users invite <email> --name "<name>" [--group <group>]`.

**Offboarding**
Each step needs approval:

1. `entra users disable`, then `entra users revoke-sessions`.
2. `exchange auto-reply set --on`.
3. Hand over mailbox and files:
   - `exchange mailbox convert <x> --to shared`, then `exchange permissions add <x> <manager> --access full`
   - `sharepoint onedrive grant <x> <manager>`
4. `intune devices retire`.
5. `entra users license remove` (after the mailbox is shared; a shared mailbox under 50 GB needs no license).

**"Who deleted / opened this?"**

- Files: `m365 purview audit file-deleted --site <url> [--file <name>]`
- Mailbox access: `m365 purview audit mailbox-access <mailbox>`
- Admin changes: `m365 purview audit admin-changes`
- Anything else: `m365 purview audit search --user <x> --operation <op>`
- Searches can take minutes. If one returns `next`, run that command later.

**"Push this app / policy to a group" / "what does group X get?"**

1. `m365 intune assignments <group>` shows what already targets the group.
2. To deploy, use `m365 intune apps assign <app> --intent required --group <group>`, `m365 intune compliance assign`,
   or `m365 intune config assign`. These are destructive because they affect every member; the preview shows the
   member count.

**"Export it to a spreadsheet"**
Add `--all --csv` to any list command and redirect to a file, e.g. `m365 intune devices list --compliance
noncompliant --all --csv > noncompliant.csv`.

**"Who can see this file?" / oversharing**
`m365 sharepoint files permissions <file or sharing URL>`: check `anonymousLinks` and `externalUsers`. To remove
access, `m365 sharepoint files unshare <url> <permissionId>` or `--all-links`.

**"Is Microsoft down?" / several users affected / health check**
`m365 troubleshoot tenant` (or just `m365 health overview`). Active issues:
`m365 health issues --active --service "<service>"`.

**Teams call quality**
`m365 teams calls list --user <x> --days 1`, then `m365 teams calls get <id>`. Per-participant `poor` flags give
the reason (packet loss, jitter, Wi-Fi). Call records need an app-only profile.

## Areas

| Area           | For                                                                                                                                                                                                                                                          |
| -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `troubleshoot` | user, signin, device, mail, tenant                                                                                                                                                                                                                           |
| `entra`        | users (lifecycle, invite, bulk-create, MFA, licenses, sign-ins), groups, licenses, subscriptions, signins, signin-error, audit, roles, apps (consents, revoke-consent, disable), ca, risky-users, risk-detections, access-reviews, partners, devices, tenant |
| `intune`       | devices (list/get/compliance/config-states/apps/bitlocker-key/laps/actions and remote actions), compliance, config, apps, assignments (assign/unassign), autopilot, report noncompliant                                                                      |
| `defender`     | incidents, alerts, hunt (KQL + templates), secure-score, threat-submit, machines (Defender for Endpoint: list/get/isolate/release/scan/collect-package/restrict/actions)                                                                                     |
| `exchange`     | mailbox (details/convert), auto-reply, rules, forwarding (set/clear), permissions, shared-mailbox, dl, quarantine, hold, archive, transport-rules, rooms, trace, messages, calendar, usage                                                                   |
| `sharepoint`   | sites, files (search/permissions/unshare/versions), onedrive (get/usage/grant/revoke), settings                                                                                                                                                              |
| `purview`      | audit (search/results/queries/file-deleted/mailbox-access/admin-changes), ediscovery (list/create/custodians/hold/release)                                                                                                                                   |
| `teams`        | list/get/members/channels, user, presence, create/archive, calls, activity                                                                                                                                                                                   |
| `health`       | overview, issues, messages                                                                                                                                                                                                                                   |
| `graph`        | raw `get` (read) and `post/patch/put/delete` (always destructive). Last resort only.                                                                                                                                                                         |

## Limits to state honestly

Some things are still out of reach. When the user asks for one, say so and name the tool that can:

- **SharePoint/PnP PowerShell:** classic SharePoint groups and site collection admins.
- **Teams PowerShell or admin center:** Teams policies and voice.
- **Defender portal:** live response, Tenant Allow/Block List.

Exchange admin commands need PowerShell 7 with the ExchangeOnlineManagement module and the Exchange sign-in. If
they fail with `CONFIG`, relay the install hint. Each `docs/domains/<area>.md` in the repo lists what is not
covered.
