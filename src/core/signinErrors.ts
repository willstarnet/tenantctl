/**
 * Common Entra sign-in error codes (the `status.errorCode` on sign-in logs; `AADSTS<code>` in messages), with
 * what they mean and what an admin can do. Source: https://learn.microsoft.com/entra/identity-platform/reference-error-codes
 * Microsoft's live lookup for any code: https://login.microsoftonline.com/error?code=<code>
 */
export interface SigninErrorInfo {
  code: number;
  name: string;
  meaning: string;
  action: string;
  /** User-caused/expected noise rather than a problem to fix. */
  benign?: boolean;
}

export const SIGNIN_ERRORS: readonly SigninErrorInfo[] = [
  { code: 0, name: 'Success', meaning: 'Sign-in succeeded.', action: 'None.', benign: true },
  {
    code: 16000,
    name: 'InteractionRequired',
    meaning: 'The account is not in the tenant it tried to sign in to (often a guest picking the wrong account).',
    action: 'Have the user pick the right account, or invite them as a guest.',
  },
  {
    code: 50011,
    name: 'InvalidReplyTo',
    meaning: 'The app sent a redirect URI that is not registered on the app.',
    action: 'Add the redirect URI to the app registration.',
  },
  {
    code: 50020,
    name: 'UserUnauthorized',
    meaning: 'The account from an external identity provider does not exist in this tenant.',
    action: 'Invite the user as a guest, or have them use their account in this tenant.',
  },
  {
    code: 50034,
    name: 'UserAccountNotFound',
    meaning: 'No such account in the directory: mistyped username, or wrong tenant.',
    action: 'Check the UPN the user typed against `m365 entra users get`.',
  },
  {
    code: 50053,
    name: 'IdsLocked',
    meaning: 'Account locked by Smart Lockout after repeated bad passwords, or sign-in from a malicious IP blocked.',
    action: 'Wait for the lockout to expire or reset the password; check risky sign-ins if the user did not try.',
  },
  {
    code: 50055,
    name: 'InvalidPasswordExpiredPassword',
    meaning: 'The password has expired.',
    action: 'User resets it (SSPR) or an admin resets it.',
  },
  {
    code: 50056,
    name: 'InvalidOrNullPassword',
    meaning: 'No password exists in the directory for this user.',
    action: 'Set a password, or check password hash sync / passwordless setup.',
  },
  {
    code: 50057,
    name: 'UserDisabled',
    meaning: 'The account is disabled (accountEnabled = false).',
    action: 'Re-enable the account if the user should have access (`m365 entra users enable`).',
  },
  {
    code: 50058,
    name: 'UserInformationNotProvided',
    meaning: 'Silent sign-in found no session. Expected before the user signs in interactively.',
    action: 'None unless it persists with a visible failure.',
    benign: true,
  },
  {
    code: 50059,
    name: 'MissingTenantRealm',
    meaning: 'No tenant could be identified from the request or credentials.',
    action: 'User signs in with a full work account (user@domain).',
  },
  {
    code: 50072,
    name: 'UserStrongAuthEnrollmentRequiredInterrupt',
    meaning: 'The user must register for MFA before continuing.',
    action: 'User completes MFA registration at https://aka.ms/mysecurityinfo.',
  },
  {
    code: 50074,
    name: 'UserStrongAuthClientAuthNRequiredInterrupt',
    meaning: 'MFA was required and the user did not pass the challenge.',
    action: "Check the user's MFA methods (`m365 entra users auth-methods`); usually user retries.",
  },
  {
    code: 50076,
    name: 'UserStrongAuthClientAuthNRequired',
    meaning: 'MFA required by Conditional Access, per-user MFA, or a new location. Normal mid-flow interrupt.',
    action: 'None if a success follows; otherwise check MFA methods.',
    benign: true,
  },
  {
    code: 50079,
    name: 'UserStrongAuthEnrollmentRequired',
    meaning: 'Policy requires MFA and the user has not registered a method.',
    action: 'User registers at https://aka.ms/mysecurityinfo, or issue a Temporary Access Pass.',
  },
  {
    code: 50089,
    name: 'FlowTokenExpired',
    meaning: 'The sign-in flow took too long and expired.',
    action: 'User retries.',
    benign: true,
  },
  {
    code: 50097,
    name: 'DeviceAuthenticationRequired',
    meaning: 'Policy requires device authentication and the device did not provide it.',
    action: 'Check the device is Entra joined/registered and healthy (`m365 troubleshoot device`).',
  },
  {
    code: 50105,
    name: 'EntitlementGrantsNotFound',
    meaning: 'The app requires assignment and the user is not assigned to it.',
    action: 'Assign the user (or a group they are in) to the enterprise app.',
  },
  {
    code: 50125,
    name: 'PasswordResetRegistrationRequiredInterrupt',
    meaning: 'Sign-in interrupted for password reset or SSPR registration.',
    action: 'User completes the prompt.',
    benign: true,
  },
  {
    code: 50126,
    name: 'InvalidUserNameOrPassword',
    meaning: 'Wrong username or password.',
    action: 'User retries or resets the password. Many from unknown IPs suggests password spray: check risk.',
  },
  {
    code: 50128,
    name: 'InvalidDomainName',
    meaning: 'No tenant found for the domain in the request.',
    action: 'User checks the domain part of their username.',
  },
  {
    code: 50131,
    name: 'ConditionalAccessFailed',
    meaning: 'Conditional Access failure: bad Windows device state, suspicious activity, or a policy decision.',
    action: "Open the sign-in's appliedConditionalAccessPolicies to see which policy failed.",
  },
  {
    code: 50132,
    name: 'SsoArtifactInvalidOrExpired',
    meaning: 'Session invalid after a password expiry or recent password change.',
    action: 'User signs in again with the new password.',
  },
  {
    code: 50133,
    name: 'SsoArtifactRevoked',
    meaning: 'Session invalid after a password expiry or recent password change.',
    action: 'User signs in again with the new password.',
  },
  {
    code: 50140,
    name: 'KmsiInterrupt',
    meaning: '"Stay signed in?" prompt. Expected part of the flow.',
    action: 'None.',
    benign: true,
  },
  {
    code: 50144,
    name: 'InvalidPasswordExpiredOnPremPassword',
    meaning: 'The on-premises Active Directory password has expired.',
    action: 'Reset the password on-premises or via SSPR with writeback.',
  },
  {
    code: 50158,
    name: 'ExternalSecurityChallenge',
    meaning: 'An external security challenge (third-party MFA, custom control) was not satisfied.',
    action: 'User completes the external challenge; check the external provider if it keeps failing.',
  },
  {
    code: 50173,
    name: 'FreshTokenNeeded',
    meaning: 'The grant was revoked, typically by a password change or reset.',
    action: 'User signs in again.',
  },
  {
    code: 50199,
    name: 'CmsiInterrupt',
    meaning: 'User was asked to confirm the app (mobile browser redirect). Expected.',
    action: 'None.',
    benign: true,
  },
  {
    code: 51004,
    name: 'UserAccountNotInDirectory',
    meaning: 'The account does not exist in the directory the app chose.',
    action: 'User signs in to the correct tenant/account.',
  },
  {
    code: 53000,
    name: 'DeviceNotCompliant',
    meaning: 'Conditional Access requires a compliant device and this device is not compliant (or not enrolled).',
    action: 'Find why the device is non-compliant: `m365 troubleshoot device <name>`.',
  },
  {
    code: 53001,
    name: 'DeviceNotDomainJoined',
    meaning: 'Conditional Access requires a (hybrid) domain-joined device.',
    action: 'User signs in from a joined device.',
  },
  {
    code: 53002,
    name: 'ApplicationUsedIsNotAnApprovedApp',
    meaning: 'Conditional Access requires an approved client app (e.g. Outlook mobile) and another app was used.',
    action: 'User switches to an approved app.',
  },
  {
    code: 53003,
    name: 'BlockedByConditionalAccess',
    meaning: 'A Conditional Access policy blocked the sign-in.',
    action: 'Check appliedConditionalAccessPolicies on the sign-in to find the blocking policy.',
  },
  {
    code: 53004,
    name: 'ProofUpBlockedDueToRisk',
    meaning: 'MFA registration blocked because the sign-in was risky.',
    action: 'Investigate risk (`m365 entra risky-users`), then have the user register from a trusted location.',
  },
  {
    code: 65001,
    name: 'DelegationDoesNotExist',
    meaning: 'Nobody has consented to the app for this user/permission.',
    action: 'Grant consent (admin consent if user consent is off).',
  },
  {
    code: 65004,
    name: 'UserDeclinedConsent',
    meaning: 'The user declined the consent prompt.',
    action: 'User retries and accepts, or an admin grants consent.',
  },
  {
    code: 70043,
    name: 'BadTokenDueToSignInFrequency',
    meaning: 'Refresh token expired by a Conditional Access sign-in frequency policy.',
    action: 'User signs in again. Expected under that policy.',
    benign: true,
  },
  {
    code: 90072,
    name: 'PassThroughUserMfaError',
    meaning: "External account cannot satisfy this tenant's MFA (guest not set up here).",
    action: 'Check the guest account and cross-tenant MFA trust settings.',
  },
  {
    code: 90094,
    name: 'AdminConsentRequired',
    meaning: 'The app needs admin consent.',
    action: 'An admin grants consent to the app.',
  },
  {
    code: 135011,
    name: 'DeviceDisabled',
    meaning: 'The device used to sign in is disabled in Entra ID.',
    action: 'Re-enable the device object, or the user uses another device.',
  },
  {
    code: 500121,
    name: 'StrongAuthRequestFailed',
    meaning: 'Authentication failed during the MFA request: denied, timed out, or the method failed.',
    action: 'Check MFA methods; repeated unexpected prompts the user denied may mean someone has the password.',
  },
  {
    code: 530032,
    name: 'BlockedByConditionalAccessOnSecurityPolicy',
    meaning: 'A tenant security policy (e.g. security defaults, block legacy auth) blocked the request.',
    action: 'Check tenant-level security policies.',
  },
  {
    code: 700016,
    name: 'UnauthorizedClient_DoesNotMatchRequest',
    meaning: 'The application was not found in the tenant.',
    action: 'Check the app id and tenant; the app may need admin consent/provisioning.',
  },
  {
    code: 700082,
    name: 'ExpiredOrRevokedGrantInactiveToken',
    meaning: 'Refresh token expired from inactivity.',
    action: 'User signs in again.',
    benign: true,
  },
  {
    code: 7000215,
    name: 'InvalidClientSecret',
    meaning: 'The app presented a wrong client secret.',
    action: 'Check the secret value (not the secret id) configured in the calling app.',
  },
  {
    code: 7000222,
    name: 'InvalidClientSecretExpiredKeysProvided',
    meaning: "The app's client secret has expired.",
    action: 'Create a new secret or certificate (`m365 entra apps expiring` finds others about to lapse).',
  },
  {
    code: 81010,
    name: 'DesktopSsoAuthTokenInvalid',
    meaning: 'Seamless SSO failed: Kerberos ticket expired or invalid.',
    action: 'Usually transient; check Seamless SSO health if widespread.',
  },
];

const BY_CODE = new Map(SIGNIN_ERRORS.map((entry) => [entry.code, entry]));

export function signinErrorInfo(code: number | string): SigninErrorInfo | undefined {
  return BY_CODE.get(Number(String(code).replace(/^AADSTS/i, '')));
}

export function signinErrorLookupUrl(code: number | string): string {
  return `https://login.microsoftonline.com/error?code=${String(code).replace(/^AADSTS/i, '')}`;
}
