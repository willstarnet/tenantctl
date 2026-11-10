# Service health

Microsoft 365 service health (incidents and advisories) and the message center (upcoming changes). Check
service health first when several users report the same problem.

## Common tasks

| Admin asks                                             | Command                                                      |
| ------------------------------------------------------ | ------------------------------------------------------------ |
| Is Microsoft 365 down? Is anything degraded right now? | `m365 health overview`                                       |
| What is wrong with Exchange / Teams right now?         | `m365 health issues --service "Exchange Online" --active`    |
| Incidents in the last week                             | `m365 health issues --classification incident --days 7`      |
| Details and timeline of an incident                    | `m365 health issue EX123456`                                 |
| What changes are coming to Teams?                      | `m365 health messages --service "Microsoft Teams" --days 30` |
| What do we need to act on?                             | `m365 health messages --action-required`                     |
| Major changes only                                     | `m365 health messages --major`                               |
| Read a message center post                             | `m365 health message MC123456`                               |

Reading `health overview`: `degraded` lists services with an ongoing issue (investigating, service degradation,
service interruption, extended recovery...). `recentlyRestored` lists services whose last issue is closed
(service restored, post-incident review published, false positive) but that are not back to "operational" yet.
Service names in `overview` (e.g. "Exchange Online", "Microsoft Teams") are the values `--service` expects.

`health issue` and `health message` return Microsoft's HTML posts as plain text; issue posts are newest first.

## Licensing and roles

- Available in every Microsoft 365 tenant; no extra license.
- Permissions: `ServiceHealth.Read.All` (overview, issues) and `ServiceMessage.Read.All` (messages), delegated or
  application.
- Delegated callers need an admin role that can read service health, such as **Service Support Administrator**
  or **Global Reader**; the message center also opens to **Message Center Reader**.

## Not available via Graph

| Task                                                     | Use instead                                                      |
| -------------------------------------------------------- | ---------------------------------------------------------------- |
| Report an issue to Microsoft or open a support case      | Microsoft 365 admin center > Support                             |
| Configure service health email notifications             | Microsoft 365 admin center > Health > Service health > Customize |
| Message center preferences (email digests, Planner sync) | Microsoft 365 admin center > Message center > Preferences        |
| Network connectivity / Microsoft 365 network health      | Microsoft 365 admin center > Health > Network connectivity       |
| Azure service health                                     | Azure portal or Azure Resource Health / Resource Graph APIs      |
