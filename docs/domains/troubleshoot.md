# Troubleshoot

One command per support question. Each one reads several areas at the same time and returns:

- `verdict`: one sentence to relay to the person asking.
- `findings`: what is wrong, ranked `high` > `medium` > `low` > `info`. Each finding has `area`, `issue`,
  `detail` and `next`, the exact command to investigate or fix it.
- `unavailable`: sections that could not be checked, with the error code. A `FORBIDDEN` here usually means a
  missing Graph permission or licence. The rest of the report is still valid.
- The raw sections the findings were built from, for follow-up questions.

All troubleshoot commands only read. Fixes are the `next` commands, which go through the normal write gate and
the confirm step.

| Question                                                        | Command                                                                          |
| --------------------------------------------------------------- | -------------------------------------------------------------------------------- |
| "Jane can't sign in / get email / use Teams"                    | `m365 troubleshoot user jane@contoso.com`                                        |
| "Why exactly is her sign-in failing?"                           | `m365 troubleshoot signin jane@contoso.com [--app "Microsoft Teams"] [--days 2]` |
| "My laptop says it isn't compliant" / "is this device healthy?" | `m365 troubleshoot device LAPTOP-0042` (name, serial or id)                      |
| "I never got the email from the vendor"                         | `m365 troubleshoot mail jane@contoso.com --sender billing@vendor.com`            |
| "Is something wrong with Microsoft 365?" / weekly check         | `m365 troubleshoot tenant`                                                       |

## What each command checks

**user**

- The account: disabled, synced from AD (change it there instead).
- Licenses: assignment errors such as `CountViolation`, or no license at all.
- MFA: whether any method is registered.
- Entra ID Protection risk.
- Sign-ins over `--days`: failures grouped by error code and explained.
- The user's Intune devices: non-compliant or stale.
- Mailbox: inbox rules that forward outside the company or look like account takeover, and auto-reply.

**signin**

- Failures grouped by error code. Each group lists the apps, IPs, locations and devices involved.
- The Conditional Access policies whose result was `failure`, with the controls they required.
- The latest correlation id, for a Microsoft support case.
- `lastSuccess` and `latest`, to tell "still failing" from "failed earlier, fine now".
- Many `50126` failures from several IPs are flagged as possible password spray.

**device**

- Intune compliance, summed up in one sentence that names the failing setting (e.g. BitLocker off).
- Last check-in.
- The Entra ID device object: missing, disabled (causes 135011), or Intune compliant but Entra not yet updated.
- Configuration profiles in error or conflict.
- Encryption, low storage, jailbreak.
- Pending remote actions such as a wipe.
- The primary user's failed sign-ins from this device.

**mail**

- Message trace to the recipient, optionally narrowed by sender and subject.
- Delivery events for up to 5 messages that were not delivered (failed, quarantined, junked or pending), so the
  last event explains the outcome.
- The recipient's inbox rules that delete, move or forward mail, which explain mail that was delivered but
  can't be found.

**tenant**

- Microsoft service health.
- Entra Connect sync more than 3 hours old.
- Licenses running out or in the renewal grace period.
- App secrets and certificates expired or expiring within 30 days.
- Intune compliance totals.
- High-risk users, active high-severity Defender incidents, and Secure Score.

## Prerequisites

Each section needs what its underlying command needs. `m365 troubleshoot <cmd> --help` lists the union. If a
section is missing, the report says so in `unavailable` instead of failing:

- **Sign-in logs:** Entra ID P1. **Risk:** Entra ID P2. **Incidents:** Microsoft Defender XDR.
- **Another user's mailbox** (inbox rules, auto-reply): app-only auth. A delegated admin sign-in only reaches
  their own mailbox.
- **Message trace:** the one-time message-trace service principal (see `m365 exchange trace --help`).
