import { type AnyCommand, defineCommand } from '../../core/types.ts';

interface SharePointSettings {
  sharingCapability?: string;
  sharingDomainRestrictionMode?: string;
  sharingAllowedDomainList?: string[];
  sharingBlockedDomainList?: string[];
  isResharingByExternalUsersEnabled?: boolean;
  isRequireAcceptingUserToMatchInvitedUserEnabled?: boolean;
  idleSessionSignOut?: { isEnabled?: boolean; warnAfterInSeconds?: number; signOutAfterInSeconds?: number };
  isLegacyAuthProtocolsEnabled?: boolean;
  isSiteCreationEnabled?: boolean;
  isUnmanagedSyncAppForTenantRestricted?: boolean;
  allowedDomainGuidsForSyncApp?: string[];
  excludedFileExtensionsForSyncApp?: string[];
  isMacSyncAppEnabled?: boolean;
  isLoopEnabled?: boolean;
  deletedUserPersonalSiteRetentionPeriodInDays?: number;
  personalSiteDefaultStorageLimitInMB?: number;
  isSitesStorageLimitAutomatic?: boolean;
  siteCreationDefaultStorageLimitInMB?: number;
}

const SHARING_CAPABILITIES = [
  'disabled',
  'existingExternalUserSharingOnly',
  'externalUserSharingOnly',
  'externalUserAndGuestSharing',
] as const;

const SHARING_MEANING: Record<string, string> = {
  disabled: 'internal only: no external sharing',
  existingExternalUserSharingOnly: 'existing guests already in the directory only',
  externalUserSharingOnly: 'new and existing guests who sign in or use a verification code; no anonymous links',
  externalUserAndGuestSharing: 'anyone, including anonymous "Anyone with the link" links',
};

function shapeSettings(s: SharePointSettings) {
  return {
    sharingCapability: s.sharingCapability ?? null,
    sharingMeans: SHARING_MEANING[s.sharingCapability ?? ''] ?? null,
    anonymousLinksAllowed: s.sharingCapability === 'externalUserAndGuestSharing',
    sharingDomainRestrictionMode: s.sharingDomainRestrictionMode ?? null,
    sharingAllowedDomains: s.sharingAllowedDomainList ?? [],
    sharingBlockedDomains: s.sharingBlockedDomainList ?? [],
    isResharingByExternalUsersEnabled: s.isResharingByExternalUsersEnabled ?? null,
    guestMustUseInvitedAccount: s.isRequireAcceptingUserToMatchInvitedUserEnabled ?? null,
    idleSessionSignOut: s.idleSessionSignOut ?? null,
    isLegacyAuthProtocolsEnabled: s.isLegacyAuthProtocolsEnabled ?? null,
    isSiteCreationEnabled: s.isSiteCreationEnabled ?? null,
    syncRestrictedToDomainJoinedPcs: s.isUnmanagedSyncAppForTenantRestricted ?? null,
    allowedDomainGuidsForSyncApp: s.allowedDomainGuidsForSyncApp ?? [],
    excludedFileExtensionsForSyncApp: s.excludedFileExtensionsForSyncApp ?? [],
    isMacSyncAppEnabled: s.isMacSyncAppEnabled ?? null,
    isLoopEnabled: s.isLoopEnabled ?? null,
    deletedUserOneDriveRetentionDays: s.deletedUserPersonalSiteRetentionPeriodInDays ?? null,
    oneDriveDefaultStorageGb:
      typeof s.personalSiteDefaultStorageLimitInMB === 'number' ? s.personalSiteDefaultStorageLimitInMB / 1024 : null,
    isSitesStorageLimitAutomatic: s.isSitesStorageLimitAutomatic ?? null,
    siteCreationDefaultStorageLimitInMB: s.siteCreationDefaultStorageLimitInMB ?? null,
  };
}

const SETTINGS_PATH = '/admin/sharepoint/settings';

export const commands: AnyCommand[] = [
  defineCommand({
    domain: 'sharepoint',
    name: 'settings',
    summary: 'Show tenant SharePoint/OneDrive settings: external sharing, domain lists, idle sign-out, sync',
    description:
      'Answers "can people share externally", "are anonymous links allowed", "which domains can we share ' +
      'with". sharingCapability is the tenant-wide ceiling; individual sites can only be more restrictive ' +
      '(per-site sharing settings are not in Graph). Delegated callers need SharePoint Administrator or Global ' +
      'Reader. Change sharing with `sharepoint settings set --sharing`.',
    examples: ['tenantctl sharepoint settings'],
    args: [],
    risk: 'read',
    permissions: {
      delegated: ['SharePointTenantSettings.Read.All'],
      application: ['SharePointTenantSettings.Read.All'],
      roles: ['SharePoint Administrator', 'Global Reader'],
    },
    async run(ctx) {
      const settings = await ctx.graph.get<SharePointSettings>(SETTINGS_PATH);
      return ctx.full ? settings : shapeSettings(settings);
    },
  }),
  defineCommand({
    domain: 'sharepoint',
    name: 'settings set',
    summary: 'Change the tenant-wide external sharing level for SharePoint and OneDrive',
    description:
      'Tenant-wide and immediate. Lowering it disables existing links/guest access above the new level across ' +
      'every site and OneDrive (e.g. to existingExternalUserSharingOnly stops all anonymous links). Levels: ' +
      'disabled < existingExternalUserSharingOnly < externalUserSharingOnly < externalUserAndGuestSharing ' +
      '(anonymous links). Needs SharePoint Administrator in delegated mode.',
    examples: ['tenantctl sharepoint settings set --sharing externalUserSharingOnly'],
    args: [
      {
        name: 'sharing',
        type: 'string',
        required: true,
        choices: SHARING_CAPABILITIES,
        description: 'New tenant sharing capability.',
      },
    ],
    risk: 'destructive',
    permissions: {
      delegated: ['SharePointTenantSettings.ReadWrite.All'],
      application: ['SharePointTenantSettings.ReadWrite.All'],
      roles: ['SharePoint Administrator'],
    },
    async preview(ctx, args) {
      const current = await ctx.graph.get<SharePointSettings>(SETTINGS_PATH);
      return {
        current: current.sharingCapability ?? null,
        currentMeans: SHARING_MEANING[current.sharingCapability ?? ''] ?? null,
        new: args.sharing,
        newMeans: SHARING_MEANING[args.sharing],
        action:
          current.sharingCapability === args.sharing
            ? `No change: tenant sharing is already ${args.sharing}.`
            : `Set tenant-wide SharePoint/OneDrive sharing from ${current.sharingCapability} to ${args.sharing} ` +
              `(${SHARING_MEANING[args.sharing]}) for every site and OneDrive.`,
      };
    },
    async run(ctx, args) {
      const updated = await ctx.graph.patch<SharePointSettings>(SETTINGS_PATH, { sharingCapability: args.sharing });
      return {
        sharingCapability: updated?.sharingCapability ?? args.sharing,
        sharingMeans: SHARING_MEANING[args.sharing],
      };
    },
  }),
];
