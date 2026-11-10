# Teams

Teams, channels, membership, presence, call quality and usage. Teams are Microsoft 365 groups with Teams
enabled, so a team id is its group id. Commands accept a team display name, group id, group mail address or a
Teams link (`...?groupId=<id>`).

## Common tasks

| Admin asks                                                      | Command                                                                                                 |
| --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| Find a team / list archived teams                               | `m365 teams list --search falcon` / `m365 teams list --archived`                                        |
| Who owns this team? Is it archived? Can guests create channels? | `m365 teams get "Project Falcon"`                                                                       |
| Who is in the team / list only owners or guests                 | `m365 teams members "Project Falcon" --role owner`                                                      |
| Add Jane to a team (as owner)                                   | `m365 teams members add "Project Falcon" jane@contoso.com [--owner]`                                    |
| Remove Jane from a team                                         | `m365 teams members remove "Project Falcon" jane@contoso.com` (preview, then `--confirm`)               |
| Jane can see the team but not a channel                         | `m365 teams channels "Project Falcon"` then `m365 teams channels members "Project Falcon" "Leadership"` |
| Which teams is Jane in?                                         | `m365 teams user jane@contoso.com`                                                                      |
| Is Bob available / why is his status stuck?                     | `m365 teams presence bob@contoso.com [jane@contoso.com ...]`                                            |
| Create a team                                                   | `m365 teams create --name "Project Falcon" --owner jane@contoso.com`                                    |
| Archive / restore a finished team                               | `m365 teams archive "Project Falcon"` / `m365 teams unarchive "Project Falcon"`                         |
| "My Teams call was choppy"                                      | `m365 teams calls list --user jane@contoso.com --days 1`, then `m365 teams calls get <id>`              |
| Phone calls placed through Teams Phone                          | `m365 teams calls pstn --days 7`                                                                        |
| Does Jane actually use Teams?                                   | `m365 teams activity --user jane@contoso.com --period D30`                                              |

Reading `teams calls get`: each row in `quality` is a person. `verdict: poor` means at least one stream crossed
a Call Quality Dashboard threshold (packet loss > 10%, jitter > 30 ms, round trip > 500 ms, audio degradation

> 1.0 or concealed samples > 7%); `poorReasons` says which. `connectionTypes: ["wifi"]` with a low
> `minWifiSignalStrength` or high `networkEventRatios.receivedQuality`/`delay` points at the user's Wi-Fi.
> `failures` lists sessions or segments that failed to set up or dropped.

## Licensing and roles

- Teams commands work with any Microsoft 365 plan that includes Teams. Delegated callers who are not team
  owners need the **Teams Administrator** role to read or change teams they are not in.
- `teams calls list|get|pstn` need **app-only** auth with application permission `CallRecords.Read.All`;
  Graph offers no delegated access to call records. Records appear 15 minutes to a few hours after a call and
  are kept 30 days. `calls pstn` covers Calling Plans and Operator Connect (not Direct Routing) and accepts at
  most 90 days.
- `teams presence` needs `Presence.Read.All`.
- `teams activity` reads the Microsoft 365 usage report (`Reports.Read.All`); delegated callers also need a
  reports role such as **Reports Reader**. If the tenant conceals user details in reports (Microsoft 365 admin
  center > Settings > Org settings > Reports), names are hashed and `--user` cannot match.
- Adding guests to a team is supported only with delegated sign-in, not app-only.
- `teams create` with app-only auth requires `--owner`. Graph accepts the request and provisions the team in the
  background (minutes). `create`, `archive` and `unarchive` return the `teamId` and `operationId`; their `next`
  field is the command that polls the operation (wait 30+ seconds between checks).

## Not available via Graph

| Task                                                                                                                   | Use instead                                                                                                                                                                                                                                                                                                                |
| ---------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Create or edit calling, meeting, messaging, app setup or other Teams policies; assign most policies to users or groups | MicrosoftTeams PowerShell (`New-/Set-CsTeamsMeetingPolicy`, `Grant-CsTeamsCallingPolicy`, ...) or Teams admin center. Graph v1.0 only reads a user's effective policies (`/admin/teams/userConfigurations`) and resolves policy ids (`/admin/teams/policy/getPolicyId`), not wrapped by this CLI (`m365 graph get` works). |
| Voice routing: online voice routing policies, PSTN usages, Direct Routing gateways, dial plans, emergency locations    | MicrosoftTeams PowerShell (`*-CsOnlineVoiceRoute*`, `*-CsOnlinePSTNGateway`, `*-CsTenantDialPlan`, `*-CsOnlineLisLocation`) or Teams admin center                                                                                                                                                                          |
| Auto attendants, call queues, resource accounts' voice configuration                                                   | MicrosoftTeams PowerShell (`*-CsAutoAttendant`, `*-CsCallQueue`) or Teams admin center                                                                                                                                                                                                                                     |
| Phone number assignment                                                                                                | Available in Graph v1.0 (`/admin/teams/telephoneNumberManagement/numberAssignments`, assign/unassign) but not wrapped by this CLI yet; use `m365 graph get` or MicrosoftTeams PowerShell (`Set-CsPhoneNumberAssignment`)                                                                                                   |
| Teams devices (Teams Rooms, phones, displays): inventory, health, restart, config                                      | Teams admin center or Teams Rooms Pro portal. Graph has only a **beta** `teamwork/devices` API, not wrapped here                                                                                                                                                                                                           |
| Direct Routing call logs                                                                                               | Graph has `getDirectRoutingCalls`; not wrapped here (`m365 graph get "/communications/callRecords/getDirectRoutingCalls(fromDateTime=...,toDateTime=...)"`)                                                                                                                                                                |
| Call Quality Dashboard reports, building/subnet (tenant data) upload                                                   | Call Quality Dashboard (cqd.teams.microsoft.com)                                                                                                                                                                                                                                                                           |
| Org-wide Teams settings (guest access, external access/federation)                                                     | MicrosoftTeams PowerShell (`Set-CsTeamsClientConfiguration`, `Set-CsTenantFederationConfiguration`) or Teams admin center                                                                                                                                                                                                  |
