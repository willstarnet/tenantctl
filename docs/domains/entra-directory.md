# Entra ID: directory

Groups, licenses, sign-in and audit logs, admin roles, enterprise apps, Conditional Access, Identity
Protection risk, Entra device objects and the tenant itself: `tenantctl entra ...`. Per-user lifecycle commands are in
[entra-users.md](entra-users.md).

Groups accept a display name, mail address or id. Group members accept users, devices or groups by name or id.
Prefix a member with `user:`, `group:` or `device:` when a name could mean more than one kind of object.

## Common tasks

| Admin intent                                           | Command                                                                                                                                                                                        |
| ------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Orient in a tenant: name, domains, is AD sync healthy  | `tenantctl entra tenant`                                                                                                                                                                       |
| Do we have a spare license?                            | `tenantctl entra licenses`                                                                                                                                                                     |
| When does each subscription expire or renew?           | `tenantctl entra subscriptions` (`renewalSoon` = next lifecycle step within 30 days; `--renewal-soon` to filter)                                                                               |
| Find a group / list security or Microsoft 365 groups   | `tenantctl entra groups list --search finance`, `--type security`                                                                                                                              |
| What does this group grant, who owns it                | `tenantctl entra groups get "Finance Team"`                                                                                                                                                    |
| Who is in a group (including nested groups)            | `tenantctl entra groups members "Finance Team" --transitive --all`                                                                                                                             |
| Add people or devices to a group                       | `tenantctl entra groups add-member "Finance Team" jane@contoso.com device:FIN-01`                                                                                                              |
| Add everyone in a CSV to a group                       | `tenantctl --dry-run entra groups add-members "Finance Team" --from-csv people.csv`, then without `--dry-run`                                                                                  |
| Remove someone from a group                            | `tenantctl entra groups remove-member "Finance Team" jane@contoso.com` (destructive: you approve the preview)                                                                                  |
| Create a group                                         | `tenantctl entra groups create --name "Finance Laptops" --type security --members device:FIN-01`                                                                                               |
| Why can't anyone in the Sales office sign in?          | `tenantctl entra signins --failures-only --days 1`, filter with `--app`, `--ip`, `--country` or `--error-code`                                                                                 |
| What does error 53003 / AADSTS50126 mean?              | `tenantctl entra signin-error 53003` (offline; no argument lists the catalog)                                                                                                                  |
| Who reset Jane's password / who changed this group?    | `tenantctl entra audit --target jane@contoso.com --days 14`                                                                                                                                    |
| Who are our Global Admins?                             | `tenantctl entra roles list`                                                                                                                                                                   |
| Is Jane an admin (active or PIM-eligible)?             | `tenantctl entra roles user jane@contoso.com`                                                                                                                                                  |
| What can this third-party app access?                  | `tenantctl entra apps get "Zoom"`                                                                                                                                                              |
| Which app secrets or certificates are about to expire? | `tenantctl entra apps expiring --days 60`                                                                                                                                                      |
| Which users consented to this app, with which scopes?  | `tenantctl entra apps consents "Shady Mail"` (delegated grants per user/admin, application permissions, with ids)                                                                              |
| Illicit consent grant: cut the app off                 | `tenantctl entra apps revoke-consent "Shady Mail" --all-delegated --all-application`, then `tenantctl entra apps disable "Shady Mail"`, then `tenantctl entra signins --app <appId> --days 30` |
| Re-enable a disabled app                               | `tenantctl entra apps enable "Shady Mail"`                                                                                                                                                     |
| Which access reviews run, and who hasn't been reviewed | `tenantctl entra access-reviews list --active`, `tenantctl entra access-reviews decisions "Quarterly guest review" --pending`                                                                  |
| MSP: which customer GDAP relationships expire soon?    | `tenantctl entra partners --active` (`expiringSoon`, `daysLeft`, roles by name)                                                                                                                |
| Why was a sign-in blocked or asked for MFA?            | `tenantctl entra ca list`, then `tenantctl entra ca get "<policy>"` (or `tenantctl troubleshoot signin <user>`)                                                                                |
| Who is risky right now?                                | `tenantctl entra risky-users --state atRisk`                                                                                                                                                   |
| Why is Jane risky?                                     | `tenantctl entra risk-detections --user jane@contoso.com --days 30`                                                                                                                            |
| False positive / confirmed compromise                  | `tenantctl entra risky-users dismiss jane@contoso.com`, `tenantctl entra risky-users confirm-compromised jane@contoso.com`                                                                     |
| Stale devices in Entra                                 | `tenantctl entra devices list --stale-days 90 --all`                                                                                                                                           |
| Block a lost laptop from signing in                    | `tenantctl entra devices disable DESKTOP-123` (and `tenantctl intune devices retire` or `wipe`)                                                                                                |

## Reading the output

- **Group `type`** is derived from the group's flags: `tenantctl` (Microsoft 365 group), `security`, `distribution`,
  or `mail-security` (mail-enabled security group). The last two are managed in Exchange.
- **Group `isTeam`** means the group has a team. Use `tenantctl teams get` to see it.
- **Sign-ins** carry an `explanation` for each error code (meaning, what to do, whether it is `benign`
  noise such as an MFA interrupt). They also show `conditionalAccess` (success, failure, notApplied) and the
  device's compliance state.
- **Conditional Access `state`**: `report-only` policies only log what they would have done; they never block.
  Scope fields read `All`, `None`, or `N specific`. `ca get` resolves those ids to names.
- **`apps get`** splits granted permissions into `applicationPermissions` and `delegatedPermissions`.
  Application permissions are app-only roles that work without a user, so they matter most in a consent review.
  Delegated permissions are scopes users or an admin consented to.
- **`apps consents`** lists `delegated` grants (consent `admin` = all users, `user` = one user's own consent, with
  the principal) and `application` permissions. The `grantId` / `assignmentId` values feed `apps revoke-consent`.
  Revoking does not end access tokens already issued (about an hour); disable the app to stop new ones.
- **`subscriptions`**: `status` runs Enabled -> Warning (expired, grace) -> Suspended -> Deleted;
  `daysToNextLifecycle` counts to the next step (or the renewal date for auto-renewing subscriptions).
- **`tenant`** sets `dirSyncStale` when on-premises sync is on and the last sync is more than 3 hours old. Entra
  Connect normally syncs every 30 minutes.

## Licensing and roles

| Feature                                                    | Requirement                                                                                             |
| ---------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| Sign-in logs (`signins`, sign-in sections of troubleshoot) | Entra ID P1 or P2                                                                                       |
| Conditional Access                                         | Entra ID P1                                                                                             |
| Risky users, risk detections, PIM eligible roles           | Entra ID P2                                                                                             |
| Group-based licensing (`groups get` licenses)              | Entra ID P1                                                                                             |
| Access reviews (`access-reviews list`, `decisions`)        | Entra ID P2 or Entra ID Governance; Security Reader, Identity Governance Administrator or Global Reader |
| GDAP relationships (`partners`)                            | Run in the partner (MSP) tenant; Graph has no customer-side view                                        |
| Revoke consents, disable apps                              | Cloud Application Administrator; Privileged Role Administrator to revoke Microsoft Graph app roles      |
| Delegated reads                                            | Global Reader covers nearly everything; Security Reader for risk; Reports Reader for logs               |
| Delegated writes                                           | Groups Administrator (groups), Cloud Device Administrator (devices), Security Administrator (risk)      |

`tenantctl permissions --domain entra --mode application` prints the exact Graph permissions.

## Not available via Graph

| Task                                                                 | Use instead                                                                                                             |
| -------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| Unlock a Smart Lockout                                               | The lockout clears on its own, or reset the password (`tenantctl entra users reset-password`)                           |
| Per-user (legacy) MFA state                                          | Beta-only in Graph; Entra admin center > Per-user MFA                                                                   |
| Activate a PIM role for yourself or approve requests                 | Entra admin center > PIM (Graph supports it, but this CLI doesn't wrap it)                                              |
| Create or edit Conditional Access policies                           | Entra admin center (Graph supports it; this CLI only reads policies, so it can't lock anyone out by mistake)            |
| Entra Connect server health and sync errors per object               | Entra Connect Health portal                                                                                             |
| Named locations, authentication method policies, cross-tenant access | Entra admin center (or `tenantctl graph get` for reads)                                                                 |
| Which partners hold delegated admin (GDAP) access to our tenant      | Microsoft 365 admin center > Settings > Partner relationships (`tenantctl entra partners` covers the partner side only) |

Anything else Graph exposes is reachable through `tenantctl graph get <path>`.
