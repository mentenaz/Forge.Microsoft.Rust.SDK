use serde::{Deserialize, Serialize};

/// Indicates the role of the sharing link.
///
/// Ported from PnPjs `SharingRole` enum
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(try_from = "i64", into = "i64")]
#[repr(i64)]
pub enum SharingRole {
    None = 0,
    View = 1,
    Edit = 2,
    Owner = 3,
}

impl std::convert::TryFrom<i64> for SharingRole {
    type Error = String;
    fn try_from(v: i64) -> std::result::Result<Self, String> {
        match v {
            0 => Ok(Self::None),
            1 => Ok(Self::View),
            2 => Ok(Self::Edit),
            3 => Ok(Self::Owner),
            _ => Err(format!("unknown SharingRole value {v}")),
        }
    }
}

impl From<SharingRole> for i64 {
    fn from(v: SharingRole) -> i64 {
        v as i64
    }
}

/// Ported from PnPjs `SPSharedObjectType` enum
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(try_from = "i64", into = "i64")]
#[repr(i64)]
pub enum SharedObjectType {
    Unknown = 0,
    File = 1,
    Folder = 2,
    Item = 3,
    List = 4,
    Web = 5,
    Max = 6,
}

impl std::convert::TryFrom<i64> for SharedObjectType {
    type Error = String;
    fn try_from(v: i64) -> std::result::Result<Self, String> {
        match v {
            0 => Ok(Self::Unknown),
            1 => Ok(Self::File),
            2 => Ok(Self::Folder),
            3 => Ok(Self::Item),
            4 => Ok(Self::List),
            5 => Ok(Self::Web),
            6 => Ok(Self::Max),
            _ => Err(format!("unknown SharedObjectType value {v}")),
        }
    }
}

impl From<SharedObjectType> for i64 {
    fn from(v: SharedObjectType) -> i64 {
        v as i64
    }
}

/// Ported from PnPjs `SharingDomainRestrictionMode` enum
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(try_from = "i64", into = "i64")]
#[repr(i64)]
pub enum SharingDomainRestrictionMode {
    None = 0,
    AllowList = 1,
    BlockList = 2,
}

impl std::convert::TryFrom<i64> for SharingDomainRestrictionMode {
    type Error = String;
    fn try_from(v: i64) -> std::result::Result<Self, String> {
        match v {
            0 => Ok(Self::None),
            1 => Ok(Self::AllowList),
            2 => Ok(Self::BlockList),
            _ => Err(format!("unknown SharingDomainRestrictionMode value {v}")),
        }
    }
}

impl From<SharingDomainRestrictionMode> for i64 {
    fn from(v: SharingDomainRestrictionMode) -> i64 {
        v as i64
    }
}

/// Summarizes the result of a sharing operation.
///
/// Ported from PnPjs `SharingOperationStatusCode` enum
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(try_from = "i64", into = "i64")]
#[repr(i64)]
pub enum SharingOperationStatusCode {
    CompletedSuccessfully = 0,
    AccessRequestsQueued = 1,
    NoResolvedUsers = -1,
    AccessDenied = -2,
    CrossSiteRequestNotSupported = -3,
    UnknowError = -4,
    EmailBodyTooLong = -5,
    ListUniqueScopesExceeded = -6,
    CapabilityDisabled = -7,
    ObjectNotSupported = -8,
    NestedGroupsNotSupported = -9,
}

impl std::convert::TryFrom<i64> for SharingOperationStatusCode {
    type Error = String;
    fn try_from(v: i64) -> std::result::Result<Self, String> {
        match v {
            0 => Ok(Self::CompletedSuccessfully),
            1 => Ok(Self::AccessRequestsQueued),
            -1 => Ok(Self::NoResolvedUsers),
            -2 => Ok(Self::AccessDenied),
            -3 => Ok(Self::CrossSiteRequestNotSupported),
            -4 => Ok(Self::UnknowError),
            -5 => Ok(Self::EmailBodyTooLong),
            -6 => Ok(Self::ListUniqueScopesExceeded),
            -7 => Ok(Self::CapabilityDisabled),
            -8 => Ok(Self::ObjectNotSupported),
            -9 => Ok(Self::NestedGroupsNotSupported),
            _ => Err(format!("unknown SharingOperationStatusCode value {v}")),
        }
    }
}

impl From<SharingOperationStatusCode> for i64 {
    fn from(v: SharingOperationStatusCode) -> i64 {
        v as i64
    }
}

/// The kinds of sharing links that can be created.
///
/// Ported from PnPjs `SharingLinkKind` enum
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default, Serialize, Deserialize)]
#[serde(try_from = "i64", into = "i64")]
#[repr(i64)]
pub enum SharingLinkKind {
    /// Uninitialized link.
    #[default]
    Uninitialized = 0,
    /// Direct link to the object being shared.
    Direct = 1,
    /// Organization-shareable link with view permissions.
    OrganizationView = 2,
    /// Organization-shareable link with edit permissions.
    OrganizationEdit = 3,
    /// View only anonymous link.
    AnonymousView = 4,
    /// Read/write anonymous link.
    AnonymousEdit = 5,
    /// Flexible sharing link, properties can change without affecting the URL.
    Flexible = 6,
}

impl std::convert::TryFrom<i64> for SharingLinkKind {
    type Error = String;
    fn try_from(v: i64) -> std::result::Result<Self, String> {
        match v {
            0 => Ok(Self::Uninitialized),
            1 => Ok(Self::Direct),
            2 => Ok(Self::OrganizationView),
            3 => Ok(Self::OrganizationEdit),
            4 => Ok(Self::AnonymousView),
            5 => Ok(Self::AnonymousEdit),
            6 => Ok(Self::Flexible),
            _ => Err(format!("unknown SharingLinkKind value {v}")),
        }
    }
}

impl From<SharingLinkKind> for i64 {
    fn from(v: SharingLinkKind) -> i64 {
        v as i64
    }
}

/// SharePoint role types used when sharing to a group.
///
/// Ported from PnPjs `RoleType` enum
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(try_from = "i64", into = "i64")]
#[repr(i64)]
pub enum RoleType {
    None = 0,
    Guest = 1,
    Reader = 2,
    Contributor = 3,
    WebDesigner = 4,
    Administrator = 5,
}

impl std::convert::TryFrom<i64> for RoleType {
    type Error = String;
    fn try_from(v: i64) -> std::result::Result<Self, String> {
        match v {
            0 => Ok(Self::None),
            1 => Ok(Self::Guest),
            2 => Ok(Self::Reader),
            3 => Ok(Self::Contributor),
            4 => Ok(Self::WebDesigner),
            5 => Ok(Self::Administrator),
            _ => Err(format!("unknown RoleType value {v}")),
        }
    }
}

impl From<RoleType> for i64 {
    fn from(v: RoleType) -> i64 {
        v as i64
    }
}

/// Basic principal (user/group) information.
///
/// Ported from PnPjs `IPrincipalInfo` (`packages/sp/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct PrincipalInfo {
    #[serde(rename = "Department", default)]
    pub department: Option<String>,
    #[serde(rename = "DisplayName", default)]
    pub display_name: Option<String>,
    #[serde(rename = "Email", default)]
    pub email: Option<String>,
    #[serde(rename = "JobTitle", default)]
    pub job_title: Option<String>,
    #[serde(rename = "LoginName", default)]
    pub login_name: Option<String>,
    #[serde(rename = "Mobile", default)]
    pub mobile: Option<String>,
    #[serde(rename = "PrincipalId", default)]
    pub principal_id: Option<i64>,
    #[serde(rename = "SIPAddress", default)]
    pub sip_address: Option<String>,
}

/// Email data used when sharing with an email notification.
///
/// Ported from PnPjs `ISharingEmailData`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Serialize)]
pub struct SharingEmailData {
    #[serde(rename = "subject", skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,
    #[serde(rename = "body")]
    pub body: String,
}

/// Options accepted by the `shareObject` operation.
///
/// Ported from PnPjs `IShareObjectOptions`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Serialize)]
pub struct ShareObjectOptions {
    #[serde(rename = "url", skip_serializing_if = "Option::is_none")]
    pub url: Option<String>,
    #[serde(rename = "loginNames", skip_serializing_if = "Option::is_none")]
    pub login_names: Option<String>,
    #[serde(rename = "role", skip_serializing_if = "Option::is_none")]
    pub role: Option<SharingRole>,
    #[serde(rename = "emailData", skip_serializing_if = "Option::is_none")]
    pub email_data: Option<SharingEmailData>,
    #[serde(rename = "group", skip_serializing_if = "Option::is_none")]
    pub group: Option<RoleType>,
    #[serde(rename = "propagateAcl", skip_serializing_if = "Option::is_none")]
    pub propagate_acl: Option<bool>,
    #[serde(
        rename = "includeAnonymousLinkInEmail",
        skip_serializing_if = "Option::is_none"
    )]
    pub include_anonymous_link_in_email: Option<bool>,
    #[serde(rename = "useSimplifiedRoles", skip_serializing_if = "Option::is_none")]
    pub use_simplified_roles: Option<bool>,
}

/// Settings for a sharing link to be created or updated.
///
/// Ported from PnPjs `IShareLinkSettings`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Serialize)]
pub struct ShareLinkSettings {
    #[serde(rename = "shareId", skip_serializing_if = "Option::is_none")]
    pub share_id: Option<String>,
    #[serde(rename = "linkKind")]
    pub link_kind: SharingLinkKind,
    #[serde(rename = "expiration", skip_serializing_if = "Option::is_none")]
    pub expiration: Option<String>,
    #[serde(rename = "role", skip_serializing_if = "Option::is_none")]
    pub role: Option<SharingRole>,
    #[serde(
        rename = "allowAnonymousAccess",
        skip_serializing_if = "Option::is_none"
    )]
    pub allow_anonymous_access: Option<bool>,
}

/// Body of the `shareLink` request.
///
/// Ported from PnPjs `IShareLinkRequest`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Serialize)]
pub struct ShareLinkRequest {
    #[serde(rename = "peoplePickerInput", skip_serializing_if = "Option::is_none")]
    pub people_picker_input: Option<String>,
    #[serde(rename = "createLink")]
    pub create_link: bool,
    #[serde(rename = "emailData", skip_serializing_if = "Option::is_none")]
    pub email_data: Option<SharingEmailData>,
    #[serde(rename = "settings")]
    pub settings: Option<ShareLinkSettings>,
}

/// Response of a `shareLink` call.
///
/// Ported from PnPjs `IShareLinkResponse`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct ShareLinkResponse {
    #[serde(rename = "sharingLinkInfo", default)]
    pub sharing_link_info: Option<SharingLinkInfo>,
}

/// Information about a sharing link.
///
/// Ported from PnPjs `ISharingLinkInfo`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct SharingLinkInfo {
    #[serde(rename = "AllowsAnonymousAccess", default)]
    pub allows_anonymous_access: Option<bool>,
    #[serde(rename = "Created", default)]
    pub created: Option<String>,
    #[serde(rename = "CreatedBy", default)]
    pub created_by: Option<PrincipalInfo>,
    #[serde(rename = "Expiration", default)]
    pub expiration: Option<String>,
    #[serde(rename = "IsActive", default)]
    pub is_active: Option<bool>,
    #[serde(rename = "IsEditLink", default)]
    pub is_edit_link: Option<bool>,
    #[serde(rename = "IsFormsLink", default)]
    pub is_forms_link: Option<bool>,
    #[serde(rename = "IsUnhealthy", default)]
    pub is_unhealthy: Option<bool>,
    #[serde(rename = "LastModified", default)]
    pub last_modified: Option<String>,
    #[serde(rename = "LastModifiedBy", default)]
    pub last_modified_by: Option<PrincipalInfo>,
    #[serde(rename = "LinkKind", default)]
    pub link_kind: Option<SharingLinkKind>,
    #[serde(rename = "ShareId", default)]
    pub share_id: Option<String>,
    #[serde(rename = "Url", default)]
    pub url: Option<String>,
}

/// Summary of a sharing operation.
///
/// Ported from PnPjs `ISharingResult`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct SharingResult {
    #[serde(rename = "PermissionsPageRelativeUrl", default)]
    pub permissions_page_relative_url: Option<String>,
    #[serde(rename = "UsersWithAccessRequests", default)]
    pub users_with_access_requests: Option<Vec<serde_json::Value>>,
    #[serde(rename = "StatusCode", default)]
    pub status_code: Option<SharingOperationStatusCode>,
    #[serde(rename = "ErrorMessage", default)]
    pub error_message: Option<String>,
    #[serde(rename = "UniquelyPermissionedUsers", default)]
    pub uniquely_permissioned_users: Option<Vec<UserSharingResult>>,
    #[serde(rename = "GroupsSharedWith", default)]
    pub groups_shared_with: Option<Vec<serde_json::Value>>,
    #[serde(rename = "GroupUsersAddedTo", default)]
    pub group_users_added_to: Option<serde_json::Value>,
    #[serde(rename = "UsersAddedToGroup", default)]
    pub users_added_to_group: Option<Vec<UserSharingResult>>,
    #[serde(rename = "InvitedUsers", default)]
    pub invited_users: Option<Vec<InvitationCreationResult>>,
    #[serde(rename = "Name", default)]
    pub name: Option<String>,
    #[serde(rename = "Url", default)]
    pub url: Option<String>,
    #[serde(rename = "IconUrl", default)]
    pub icon_url: Option<String>,
}

/// Result of an external user invitation.
///
/// Ported from PnPjs `IInvitationCreationResult`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct InvitationCreationResult {
    #[serde(rename = "Succeeded", default)]
    pub succeeded: Option<bool>,
    #[serde(rename = "Email", default)]
    pub email: Option<String>,
    #[serde(rename = "InvitationLink", default)]
    pub invitation_link: Option<String>,
}

/// Result of sharing with an individual user.
///
/// Ported from PnPjs `IUserSharingResult`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct UserSharingResult {
    #[serde(rename = "IsUserKnown", default)]
    pub is_user_known: Option<bool>,
    #[serde(rename = "Status", default)]
    pub status: Option<bool>,
    #[serde(rename = "Message", default)]
    pub message: Option<String>,
    #[serde(rename = "User", default)]
    pub user: Option<String>,
    #[serde(rename = "DisplayName", default)]
    pub display_name: Option<String>,
    #[serde(rename = "Email", default)]
    pub email: Option<String>,
    #[serde(rename = "CurrentRole", default)]
    pub current_role: Option<SharingRole>,
    #[serde(rename = "AllowedRoles", default)]
    pub allowed_roles: Option<Vec<SharingRole>>,
    #[serde(rename = "InvitationLink", default)]
    pub invitation_link: Option<String>,
}

/// A recipient entity for permission checks.
///
/// Ported from PnPjs `ISharingRecipient`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SharingRecipient {
    #[serde(rename = "email", skip_serializing_if = "Option::is_none")]
    pub email: Option<String>,
    #[serde(rename = "alias", skip_serializing_if = "Option::is_none")]
    pub alias: Option<String>,
}

/// The permission an entity has on a shared object.
///
/// Ported from PnPjs `ISharingEntityPermission`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct SharingEntityPermission {
    #[serde(rename = "inputEntity", default)]
    pub input_entity: Option<String>,
    #[serde(rename = "resolvedEntity", default)]
    pub resolved_entity: Option<String>,
    #[serde(rename = "hasAccess", default)]
    pub has_access: Option<bool>,
    #[serde(rename = "role", default)]
    pub role: Option<SharingRole>,
}

/// Parameters controlling the `getSharingInformation` call.
///
/// Ported from PnPjs `ISharingInformationRequest`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Serialize)]
pub struct SharingInformationRequest {
    #[serde(rename = "maxPrincipalsToReturn")]
    pub max_principals_to_return: i64,
    #[serde(
        rename = "clientSupportedFeatures",
        skip_serializing_if = "Option::is_none"
    )]
    pub client_supported_features: Option<String>,
    #[serde(rename = "populateInheritedLinks")]
    pub populate_inherited_links: bool,
}

/// Sharing-related data from `getSharingInformation`.
///
/// Ported from PnPjs `ISharingInformation`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct SharingInformation {
    #[serde(rename = "canAddExternalPrincipal", default)]
    pub can_add_external_principal: Option<bool>,
    #[serde(rename = "canAddInternalPrincipal", default)]
    pub can_add_internal_principal: Option<bool>,
    #[serde(rename = "canSendEmail", default)]
    pub can_send_email: Option<bool>,
    #[serde(rename = "canUseSimplifiedRoles", default)]
    pub can_use_simplified_roles: Option<bool>,
    #[serde(rename = "hasUniquePermissions", default)]
    pub has_unique_permissions: Option<bool>,
    #[serde(rename = "currentRole", default)]
    pub current_role: Option<SharingRole>,
    #[serde(rename = "requiresAccessApproval", default)]
    pub requires_access_approval: Option<bool>,
    #[serde(rename = "hasPendingAccessRequests", default)]
    pub has_pending_access_requests: Option<bool>,
    #[serde(rename = "pendingAccessRequestsLink", default)]
    pub pending_access_requests_link: Option<String>,
    #[serde(rename = "sharedObjectType", default)]
    pub shared_object_type: Option<SharedObjectType>,
    #[serde(rename = "directUrl", default)]
    pub direct_url: Option<String>,
    #[serde(rename = "webUrl", default)]
    pub web_url: Option<String>,
    #[serde(rename = "defaultLinkKind", default)]
    pub default_link_kind: Option<SharingLinkKind>,
    #[serde(rename = "domainRestrictionMode", default)]
    pub domain_restriction_mode: Option<SharingDomainRestrictionMode>,
    #[serde(rename = "RestrictedDomains", default)]
    pub restricted_domains: Option<String>,
    #[serde(rename = "anonymousLinkExpirationRestrictionDays", default)]
    pub anonymous_link_expiration_restriction_days: Option<f64>,
    #[serde(rename = "permissionsInformation", default)]
    pub permissions_information: Option<serde_json::Value>,
    #[serde(rename = "pickerSettings", default)]
    pub picker_settings: Option<serde_json::Value>,
}

/// The full sharing settings of a securable object.
///
/// Ported from PnPjs `IObjectSharingSettings`
/// (`packages/sp/sharing/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct ObjectSharingSettings {
    #[serde(rename = "WebUrl", default)]
    pub web_url: Option<String>,
    #[serde(rename = "ListId", default)]
    pub list_id: Option<String>,
    #[serde(rename = "ItemId", default)]
    pub item_id: Option<String>,
    #[serde(rename = "ItemName", default)]
    pub item_name: Option<String>,
    #[serde(rename = "ItemUrl", default)]
    pub item_url: Option<String>,
    #[serde(rename = "ObjectSharingInformation", default)]
    pub object_sharing_information: Option<serde_json::Value>,
    #[serde(rename = "AccessRequestMode", default)]
    pub access_request_mode: Option<bool>,
    #[serde(rename = "PermissionsOnlyMode", default)]
    pub permissions_only_mode: Option<bool>,
    #[serde(rename = "InheritingWebLink", default)]
    pub inheriting_web_link: Option<String>,
    #[serde(rename = "ShareByEmailEnabled", default)]
    pub share_by_email_enabled: Option<bool>,
    #[serde(rename = "IsGuestUser", default)]
    pub is_guest_user: Option<bool>,
    #[serde(rename = "HasEditRole", default)]
    pub has_edit_role: Option<bool>,
    #[serde(rename = "HasReadRole", default)]
    pub has_read_role: Option<bool>,
    #[serde(rename = "IsPictureLibrary", default)]
    pub is_picture_library: Option<bool>,
    #[serde(rename = "CanShareFolder", default)]
    pub can_share_folder: Option<bool>,
    #[serde(rename = "CanSendEmail", default)]
    pub can_send_email: Option<bool>,
    #[serde(rename = "DefaultShareLinkType", default)]
    pub default_share_link_type: Option<SharingLinkKind>,
    #[serde(rename = "SupportsAclPropagation", default)]
    pub supports_acl_propagation: Option<bool>,
    #[serde(rename = "CanCurrentUserShareInternally", default)]
    pub can_current_user_share_internally: Option<bool>,
    #[serde(rename = "CanCurrentUserShareExternally", default)]
    pub can_current_user_share_externally: Option<bool>,
    #[serde(rename = "CanCurrentUserRetrieveReadonlyLink", default)]
    pub can_current_user_retrieve_readonly_link: Option<bool>,
    #[serde(rename = "CanCurrentUserManageReadonlyLink", default)]
    pub can_current_user_manage_readonly_link: Option<bool>,
    #[serde(rename = "CanCurrentUserRetrieveReadWriteLink", default)]
    pub can_current_user_retrieve_read_write_link: Option<bool>,
    #[serde(rename = "CanCurrentUserManageReadWriteLink", default)]
    pub can_current_user_manage_read_write_link: Option<bool>,
    #[serde(rename = "CanCurrentUserRetrieveOrganizationReadonlyLink", default)]
    pub can_current_user_retrieve_organization_readonly_link: Option<bool>,
    #[serde(rename = "CanCurrentUserManageOrganizationReadonlyLink", default)]
    pub can_current_user_manage_organization_readonly_link: Option<bool>,
    #[serde(rename = "CanCurrentUserRetrieveOrganizationReadWriteLink", default)]
    pub can_current_user_retrieve_organization_read_write_link: Option<bool>,
    #[serde(rename = "CanCurrentUserManageOrganizationReadWriteLink", default)]
    pub can_current_user_manage_organization_read_write_link: Option<bool>,
    #[serde(rename = "CanSendLink", default)]
    pub can_send_link: Option<bool>,
    #[serde(rename = "ShowExternalSharingWarning", default)]
    pub show_external_sharing_warning: Option<bool>,
    #[serde(rename = "SharingPermissions", default)]
    pub sharing_permissions: Option<Vec<serde_json::Value>>,
    #[serde(rename = "SimplifiedRoles", default)]
    pub simplified_roles: Option<std::collections::HashMap<String, String>>,
    #[serde(rename = "GroupsList", default)]
    pub groups_list: Option<std::collections::HashMap<String, String>>,
    #[serde(rename = "Roles", default)]
    pub roles: Option<std::collections::HashMap<String, String>>,
    #[serde(rename = "SharePointSettings", default)]
    pub sharepoint_settings: Option<serde_json::Value>,
    #[serde(rename = "IsUserSiteAdmin", default)]
    pub is_user_site_admin: Option<bool>,
    #[serde(rename = "RequiredAnonymousLinkExpirationInDays", default)]
    pub required_anonymous_link_expiration_in_days: Option<i64>,
}

/// The empty GUID used as the default `shareId`.
pub const EMPTY_GUID: &str = "00000000-0000-0000-0000-000000000000";
