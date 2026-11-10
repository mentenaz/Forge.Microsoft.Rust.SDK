use serde::{Deserialize, Serialize};

/// Parameters for the Client People Picker query.
///
/// Ported from PnPjs `IClientPeoplePickerQueryParameters`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Serialize)]
pub struct PeoplePickerQueryParameters {
    #[serde(rename = "QueryString")]
    pub query_string: String,
    #[serde(
        rename = "AllowEmailAddresses",
        skip_serializing_if = "Option::is_none"
    )]
    pub allow_email_addresses: Option<bool>,
    #[serde(
        rename = "AllowMultipleEntities",
        skip_serializing_if = "Option::is_none"
    )]
    pub allow_multiple_entities: Option<bool>,
    #[serde(
        rename = "AllowOnlyEmailAddresses",
        skip_serializing_if = "Option::is_none"
    )]
    pub allow_only_email_addresses: Option<bool>,
    #[serde(rename = "AllUrlZones", skip_serializing_if = "Option::is_none")]
    pub all_url_zones: Option<bool>,
    #[serde(
        rename = "EnabledClaimProviders",
        skip_serializing_if = "Option::is_none"
    )]
    pub enabled_claim_providers: Option<String>,
    #[serde(rename = "ForceClaims", skip_serializing_if = "Option::is_none")]
    pub force_claims: Option<bool>,
    #[serde(rename = "MaximumEntitySuggestions")]
    pub maximum_entity_suggestions: i64,
    #[serde(rename = "PrincipalSource", skip_serializing_if = "Option::is_none")]
    pub principal_source: Option<i64>,
    #[serde(rename = "PrincipalType", skip_serializing_if = "Option::is_none")]
    pub principal_type: Option<i64>,
    #[serde(rename = "SharePointGroupID", skip_serializing_if = "Option::is_none")]
    pub sharepoint_group_id: Option<i64>,
    #[serde(rename = "WebApplicationID", skip_serializing_if = "Option::is_none")]
    pub web_application_id: Option<String>,
}

/// A resolved or matched people picker entity.
///
/// Ported from PnPjs `IPeoplePickerEntity`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct PeoplePickerEntity {
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "DisplayText")]
    pub display_text: String,
    #[serde(rename = "EntityData")]
    pub entity_data: PeoplePickerEntityData,
    #[serde(rename = "EntityType")]
    pub entity_type: String,
    #[serde(rename = "IsResolved")]
    pub is_resolved: bool,
    #[serde(rename = "Key")]
    pub key: String,
    #[serde(rename = "MultipleMatches")]
    pub multiple_matches: Vec<PeoplePickerEntityData>,
    #[serde(rename = "ProviderDisplayName")]
    pub provider_display_name: String,
    #[serde(rename = "ProviderName")]
    pub provider_name: String,
}

/// Data associated with a people picker entity.
///
/// Ported from PnPjs `IPeoplePickerEntityData`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct PeoplePickerEntityData {
    #[serde(rename = "AccountName", default)]
    pub account_name: Option<String>,
    #[serde(rename = "Department", default)]
    pub department: Option<String>,
    #[serde(rename = "Email", default)]
    pub email: Option<String>,
    #[serde(rename = "MobilePhone", default)]
    pub mobile_phone: Option<String>,
    #[serde(rename = "ObjectId", default)]
    pub object_id: Option<String>,
    #[serde(rename = "PrincipalType", default)]
    pub principal_type: Option<String>,
    #[serde(rename = "SPGroupID", default)]
    pub sp_group_id: Option<String>,
    #[serde(rename = "SPUserID", default)]
    pub sp_user_id: Option<String>,
    #[serde(rename = "Title", default)]
    pub title: Option<String>,
}

/// A hashtag from the trending tags or followed tags.
///
/// Ported from PnPjs `IHashTag` (`packages/sp/profiles/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct HashTag {
    #[serde(rename = "Name", default)]
    pub name: Option<String>,
    #[serde(rename = "UseCount", default)]
    pub use_count: Option<i64>,
}

/// Collection of trending or followed hashtags.
///
/// Ported from PnPjs `IHashTagCollection`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct HashTagCollection {
    #[serde(rename = "Items")]
    pub items: Vec<HashTag>,
}

/// User profile information returned by ProfileLoader.
///
/// Ported from PnPjs `IUserProfile` (`packages/sp/profiles/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct UserProfile {
    #[serde(rename = "AccountName", default)]
    pub account_name: Option<String>,
    #[serde(rename = "DisplayName", default)]
    pub display_name: Option<String>,
    #[serde(rename = "PersonalSite", default)]
    pub personal_site: Option<String>,
    #[serde(rename = "PersonalSiteCapabilities", default)]
    pub personal_site_capabilities: Option<i64>,
    #[serde(rename = "PersonalSiteFirstCreationError", default)]
    pub personal_site_first_creation_error: Option<String>,
    #[serde(rename = "PersonalSiteInstantiationState", default)]
    pub personal_site_instantiation_state: Option<i64>,
    #[serde(rename = "PublicUrl", default)]
    pub public_url: Option<String>,
}

/// Followed content URLs.
///
/// Ported from PnPjs `IFollowedContent`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct FollowedContent {
    #[serde(rename = "FollowedDocumentsUrl")]
    pub followed_documents_url: String,
    #[serde(rename = "FollowedSitesUrl")]
    pub followed_sites_url: String,
}
