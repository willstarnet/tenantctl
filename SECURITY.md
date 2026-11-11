# Security policy

## Supported versions

Only the latest release gets fixes. Update by re-running the install script, or with
`brew upgrade tenantctl` / `scoop update tenantctl`.

## Reporting a vulnerability

Report it privately through GitHub:
<https://github.com/willstarnet/tenantctl/security/advisories/new>. Do not open a public issue,
pull request or discussion about it.

Include what you found, the tenantctl version (`tenantctl --version`), your OS, how you signed in
(delegated or app-only), and steps or a proof of concept. Never include tenant data, tokens, client
secrets or certificates. You will get a reply within a week; the fix ships in a new release with a
published advisory.

## Scope

tenantctl runs with whatever Microsoft Graph, Exchange and Defender permissions its sign-in holds,
so its own safeguards matter. A way to break one of these is in scope:

- a destructive command making a change without the person approving it in the dialog or terminal
  prompt (or, in `token` mode, without a valid confirm token bound to that command and arguments);
- an agent or environment variable weakening `confirmMode` or enabling writes without the person's
  approval;
- reading or recovering the token cache or its key (Keychain, libsecret, DPAPI, key file) as another
  user, or tokens, secrets or passwords appearing in output, logs, the audit log or `rerun` hints;
- user input reaching a Graph path, OData filter, KQL query, PowerShell script or approval dialog
  without encoding (injection);
- the install scripts or release archives installing something that doesn't match `SHA256SUMS`.

Out of scope: vulnerabilities in Microsoft 365 or Microsoft Graph themselves (report those to the
[Microsoft Security Response Center](https://msrc.microsoft.com/report/)), a malicious process already
running as the same OS user (it can edit the config directly; see the README), and modified builds.
