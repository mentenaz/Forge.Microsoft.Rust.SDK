use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::{Deserialize, Serialize};
use serde_json::Value;

/// Result from creating or retrieving a site script.
///
/// Ported from PnPjs `ISiteScriptInfo` (`packages/sp/site-scripts/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteScriptInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "Content")]
    pub content: String,
    #[serde(rename = "Version")]
    pub version: String,
}

/// Data for updating a site script; all fields but `id` are optional in the
/// REST call.
///
/// Ported from PnPjs `ISiteScriptUpdateInfo`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SiteScriptUpdateInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title", skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(rename = "Description", skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(rename = "Content", skip_serializing_if = "Option::is_none")]
    pub content: Option<String>,
    #[serde(rename = "Version", skip_serializing_if = "Option::is_none")]
    pub version: Option<String>,
}

/// Configuration for what to include when extracting a site script from a web.
///
/// Ported from PnPjs `ISiteScriptSerializationInfo`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SiteScriptSerializationInfo {
    #[serde(rename = "IncludeBranding", skip_serializing_if = "Option::is_none")]
    pub include_branding: Option<bool>,
    #[serde(rename = "IncludedLists", skip_serializing_if = "Option::is_none")]
    pub included_lists: Option<Vec<String>>,
    #[serde(
        rename = "IncludeLinksToExportedItems",
        skip_serializing_if = "Option::is_none"
    )]
    pub include_links_to_exported_items: Option<bool>,
    #[serde(
        rename = "IncludeRegionalSettings",
        skip_serializing_if = "Option::is_none"
    )]
    pub include_regional_settings: Option<bool>,
    #[serde(
        rename = "IncludeSiteExternalSharingCapability",
        skip_serializing_if = "Option::is_none"
    )]
    pub include_site_external_sharing_capability: Option<bool>,
    #[serde(rename = "IncludeTheme", skip_serializing_if = "Option::is_none")]
    pub include_theme: Option<bool>,
}

/// Ported from PnPjs `ISiteScriptSerializationResult`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteScriptSerializationResult {
    #[serde(rename = "JSON")]
    pub json: String,
    #[serde(rename = "Warnings")]
    pub warnings: Vec<String>,
}

/// Ported from PnPjs `SiteScriptActionOutcome`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
/// Serializes/deserializes as its underlying numeric code, matching the wire
/// shape (no `serde_repr` dependency in this workspace, so implemented by hand).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum SiteScriptActionOutcome {
    Success = 0,
    Failure = 1,
    NoOp = 2,
    SucceededWithException = 3,
}

impl Serialize for SiteScriptActionOutcome {
    fn serialize<S: serde::Serializer>(
        &self,
        serializer: S,
    ) -> std::result::Result<S::Ok, S::Error> {
        serializer.serialize_u8(*self as u8)
    }
}

impl<'de> Deserialize<'de> for SiteScriptActionOutcome {
    fn deserialize<D: serde::Deserializer<'de>>(
        deserializer: D,
    ) -> std::result::Result<Self, D::Error> {
        match u8::deserialize(deserializer)? {
            0 => Ok(Self::Success),
            1 => Ok(Self::Failure),
            2 => Ok(Self::NoOp),
            3 => Ok(Self::SucceededWithException),
            other => Err(serde::de::Error::custom(format!(
                "unknown SiteScriptActionOutcome value: {other}"
            ))),
        }
    }
}

/// Ported from PnPjs `ISiteScriptActionResult`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteScriptActionResult {
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "OutcomeText")]
    pub outcome_text: String,
    #[serde(rename = "Outcome")]
    pub outcome: SiteScriptActionOutcome,
    #[serde(rename = "Target")]
    pub target: String,
}

/// Doubles embedded `'` and percent-encodes for use inside a URL-embedded
/// query-string parameter value (`?@title='...'`).
/// Ported from PnPjs `encodePath` (`packages/sp/utils/encode-path-str.ts`).
fn encode_query_value(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn utility_url(site_url: &str, method_name: &str) -> String {
    format!(
        "{}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.{method_name}",
        site_url.trim_end_matches('/')
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-site-scripts linked in?")
        })
}

#[pnp_operation(id = "sp.site_scripts.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_scripts_op() {}

#[pnp_operation(id = "sp.site_scripts.create", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_site_script_op() {}

#[pnp_operation(id = "sp.site_scripts.get_metadata", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_script_metadata_op() {}

#[pnp_operation(id = "sp.site_scripts.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_site_script_op() {}

#[pnp_operation(id = "sp.site_scripts.update", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn update_site_script_op() {}

#[pnp_operation(id = "sp.site_scripts.get_from_list", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_script_from_list_op() {}

#[pnp_operation(id = "sp.site_scripts.get_from_web", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_script_from_web_op() {}

#[pnp_operation(id = "sp.site_scripts.execute_action", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn execute_site_script_action_op() {}

/// Gets a list of information on all existing site scripts. Must be called
/// against the tenant admin site.
///
/// Ported from PnPjs `SiteScripts.getSiteScripts`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_scripts(client: &Client<'_>, site_url: &str) -> Result<Vec<SiteScriptInfo>> {
    let url = utility_url(site_url, "GetSiteScripts");
    let bytes = client
        .run_ladder(
            entry("sp.site_scripts.get_all"),
            "POST",
            &url,
            &[],
            Some(b"{}"),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Creates a new site script. `content` is the JSON value describing the
/// script actions.
///
/// Ported from PnPjs `SiteScripts.createSiteScript`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn create_site_script(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    description: &str,
    content: &Value,
) -> Result<SiteScriptInfo> {
    let url = format!(
        "{}?@title='{}'&@desc='{}'",
        utility_url(site_url, "CreateSiteScript(Title=@title,Description=@desc)"),
        encode_query_value(title),
        encode_query_value(description)
    );
    let body = serde_json::to_vec(content)?;
    let bytes = client
        .run_ladder(
            entry("sp.site_scripts.create"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets information about a specific site script, including its JSON content.
///
/// Ported from PnPjs `SiteScripts.getSiteScriptMetadata`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_script_metadata(
    client: &Client<'_>,
    site_url: &str,
    id: &str,
) -> Result<SiteScriptInfo> {
    let url = utility_url(site_url, "GetSiteScriptMetadata");
    let body = serde_json::to_vec(&serde_json::json!({ "id": id }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_scripts.get_metadata"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes a site script.
///
/// Ported from PnPjs `SiteScripts.deleteSiteScript`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn delete_site_script(client: &Client<'_>, site_url: &str, id: &str) -> Result<()> {
    let url = utility_url(site_url, "DeleteSiteScript");
    let body = serde_json::to_vec(&serde_json::json!({ "id": id }))?;
    client
        .run_ladder(
            entry("sp.site_scripts.delete"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(())
}

/// Updates a site script with new values. If `content` is provided it
/// overwrites `update_info.content` as a serialized JSON string, matching
/// PnPjs's behavior of stringifying the passed-in content object.
///
/// Ported from PnPjs `SiteScripts.updateSiteScript`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn update_site_script(
    client: &Client<'_>,
    site_url: &str,
    mut update_info: SiteScriptUpdateInfo,
    content: Option<&Value>,
) -> Result<SiteScriptInfo> {
    if let Some(content) = content {
        update_info.content = Some(content.to_string());
    }
    let url = utility_url(site_url, "UpdateSiteScript");
    let body = serde_json::to_vec(&serde_json::json!({ "updateInfo": update_info }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_scripts.update"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the site script syntax (JSON) for a specific list, given its
/// absolute URL.
///
/// Ported from PnPjs `SiteScripts.getSiteScriptFromList`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_script_from_list(
    client: &Client<'_>,
    site_url: &str,
    list_url: &str,
) -> Result<String> {
    let url = utility_url(site_url, "GetSiteScriptFromList");
    let body = serde_json::to_vec(&serde_json::json!({ "listUrl": list_url }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_scripts.get_from_list"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the site script syntax (JSON) for a specific web, given its absolute
/// URL and an extraction-configuration object.
///
/// Ported from PnPjs `SiteScripts.getSiteScriptFromWeb`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_script_from_web(
    client: &Client<'_>,
    site_url: &str,
    web_url: &str,
    info: &SiteScriptSerializationInfo,
) -> Result<SiteScriptSerializationResult> {
    let url = utility_url(site_url, "getSiteScriptFromWeb");
    let body = serde_json::to_vec(&serde_json::json!({ "webUrl": web_url, "info": info }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_scripts.get_from_web"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Executes the indicated site script action definition (a single-action
/// JSON fragment) on the current web.
///
/// Ported from PnPjs `SiteScripts.executeSiteScriptAction`
/// (`packages/sp/site-scripts/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn execute_site_script_action(
    client: &Client<'_>,
    site_url: &str,
    action_definition: &str,
) -> Result<SiteScriptActionResult> {
    let url = utility_url(site_url, "executeSiteScriptAction");
    let body = serde_json::to_vec(&serde_json::json!({ "actionDefinition": action_definition }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_scripts.execute_action"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
