use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::{Deserialize, Serialize};
use serde_json::json;

/// Result from creating or retrieving a site design.
///
/// Ported from PnPjs `ISiteDesignInfo` (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteDesignInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "WebTemplate")]
    pub web_template: String,
    #[serde(rename = "SiteScriptIds")]
    pub site_script_ids: Vec<String>,
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "PreviewImageUrl")]
    pub preview_image_url: String,
    #[serde(rename = "PreviewImageAltText")]
    pub preview_image_alt_text: String,
    #[serde(rename = "IsDefault")]
    pub is_default: bool,
    #[serde(rename = "Version")]
    pub version: String,
}

/// Data for creating a site design. All fields beyond `title`/`web_template`
/// are optional in the REST call.
///
/// Ported from PnPjs `ISiteDesignCreationInfo`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SiteDesignCreationInfo {
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "WebTemplate")]
    pub web_template: String,
    #[serde(rename = "SiteScriptIds", skip_serializing_if = "Option::is_none")]
    pub site_script_ids: Option<Vec<String>>,
    #[serde(rename = "Description", skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(rename = "PreviewImageUrl", skip_serializing_if = "Option::is_none")]
    pub preview_image_url: Option<String>,
    #[serde(
        rename = "PreviewImageAltText",
        skip_serializing_if = "Option::is_none"
    )]
    pub preview_image_alt_text: Option<String>,
    #[serde(rename = "IsDefault", skip_serializing_if = "Option::is_none")]
    pub is_default: Option<bool>,
    #[serde(rename = "ThumbnailUrl", skip_serializing_if = "Option::is_none")]
    pub thumbnail_url: Option<String>,
    #[serde(rename = "DesignPackageId", skip_serializing_if = "Option::is_none")]
    pub design_package_id: Option<String>,
    #[serde(rename = "TemplateFeatures", skip_serializing_if = "Option::is_none")]
    pub template_features: Option<Vec<String>>,
    #[serde(
        rename = "SupportedWebTemplates",
        skip_serializing_if = "Option::is_none"
    )]
    pub supported_web_templates: Option<Vec<String>>,
    #[serde(
        rename = "RequiresGroupConnected",
        skip_serializing_if = "Option::is_none"
    )]
    pub requires_group_connected: Option<bool>,
    #[serde(
        rename = "RequiresTeamsConnected",
        skip_serializing_if = "Option::is_none"
    )]
    pub requires_teams_connected: Option<bool>,
    #[serde(
        rename = "RequiresYammerConnected",
        skip_serializing_if = "Option::is_none"
    )]
    pub requires_yammer_connected: Option<bool>,
    #[serde(
        rename = "RequiresClassConnected",
        skip_serializing_if = "Option::is_none"
    )]
    pub requires_class_connected: Option<bool>,
    #[serde(
        rename = "RequiresSyntexLicense",
        skip_serializing_if = "Option::is_none"
    )]
    pub requires_syntex_license: Option<bool>,
    #[serde(rename = "IsTenantAdminOnly", skip_serializing_if = "Option::is_none")]
    pub is_tenant_admin_only: Option<bool>,
    /// `TemplateDesignType`: 0 = Site, 1 = List (kept as a raw discriminant —
    /// see PnPjs `TemplateDesignType` enum).
    #[serde(rename = "DesignType", skip_serializing_if = "Option::is_none")]
    pub design_type: Option<i64>,
    /// `ListDesignColor` discriminant (0-11) — see PnPjs `ListDesignColor` enum.
    #[serde(rename = "ListColor", skip_serializing_if = "Option::is_none")]
    pub list_color: Option<i64>,
    /// `ListDesignIcon` discriminant (0-11) — see PnPjs `ListDesignIcon` enum.
    #[serde(rename = "ListIcon", skip_serializing_if = "Option::is_none")]
    pub list_icon: Option<i64>,
    #[serde(rename = "TargetPlatforms", skip_serializing_if = "Option::is_none")]
    pub target_platforms: Option<Vec<String>>,
}

/// Data for updating a site design. All fields beyond `id` are optional; the
/// REST call resets any field to its default when omitted (per PnPjs docs on
/// `IsDefault`), so callers must resend fields they want preserved.
///
/// Ported from PnPjs `ISiteDesignUpdateInfo`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteDesignUpdateInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(flatten)]
    pub fields: SiteDesignCreationInfo,
}

/// Result from retrieving the rights for a site design.
///
/// Ported from PnPjs `ISiteDesignPrincipals`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteDesignPrincipals {
    #[serde(rename = "DisplayName")]
    pub display_name: String,
    #[serde(rename = "PrincipalName")]
    pub principal_name: String,
    #[serde(rename = "Rights")]
    pub rights: i64,
}

/// Ported from PnPjs `ISiteDesignTask` (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteDesignTask {
    #[serde(rename = "ID")]
    pub id: String,
    #[serde(rename = "LogonName")]
    pub logon_name: String,
    #[serde(rename = "SiteDesignID")]
    pub site_design_id: String,
    #[serde(rename = "SiteID")]
    pub site_id: String,
    #[serde(rename = "WebID")]
    pub web_id: String,
}

/// Ported from PnPjs `ISiteScriptActionStatus`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteScriptActionStatus {
    #[serde(rename = "ActionIndex")]
    pub action_index: i64,
    #[serde(rename = "ActionKey")]
    pub action_key: String,
    #[serde(rename = "ActionTitle")]
    pub action_title: String,
    #[serde(rename = "LastModified")]
    pub last_modified: i64,
    #[serde(rename = "OrdinalIndex")]
    pub ordinal_index: String,
    #[serde(rename = "OutcomeCode")]
    pub outcome_code: i64,
    #[serde(rename = "OutcomeText")]
    pub outcome_text: String,
    #[serde(rename = "SiteScriptID")]
    pub site_script_id: String,
    #[serde(rename = "SiteScriptIndex")]
    pub site_script_index: i64,
    #[serde(rename = "SiteScriptTitle")]
    pub site_script_title: String,
}

/// Ported from PnPjs `ISiteDesignRun` (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SiteDesignRun {
    #[serde(rename = "ID")]
    pub id: String,
    #[serde(rename = "SiteDesignID")]
    pub site_design_id: String,
    #[serde(rename = "SiteDesignTitle")]
    pub site_design_title: String,
    #[serde(rename = "SiteDesignVersion")]
    pub site_design_version: i64,
    #[serde(rename = "SiteID")]
    pub site_id: String,
    #[serde(rename = "StartTime")]
    pub start_time: i64,
    #[serde(rename = "WebID")]
    pub web_id: String,
}

fn site_designs_url(site_url: &str, method_name: &str) -> String {
    format!(
        "{}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.{method_name}",
        site_url.trim_end_matches('/')
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-site-designs linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.site_designs.create", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_site_design_op() {}

#[pnp_operation(id = "sp.site_designs.apply", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn apply_site_design_op() {}

#[pnp_operation(id = "sp.site_designs.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_designs_op() {}

#[pnp_operation(id = "sp.site_designs.get_metadata", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_design_metadata_op() {}

#[pnp_operation(id = "sp.site_designs.update", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn update_site_design_op() {}

#[pnp_operation(id = "sp.site_designs.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_site_design_op() {}

#[pnp_operation(id = "sp.site_designs.get_rights", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_design_rights_op() {}

#[pnp_operation(id = "sp.site_designs.grant_rights", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn grant_site_design_rights_op() {}

#[pnp_operation(id = "sp.site_designs.revoke_rights", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn revoke_site_design_rights_op() {}

#[pnp_operation(id = "sp.site_designs.add_task", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_site_design_task_op() {}

#[pnp_operation(id = "sp.site_designs.add_task_current_web", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_site_design_task_to_current_web_op() {}

#[pnp_operation(id = "sp.site_designs.get_task", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_design_task_op() {}

#[pnp_operation(id = "sp.site_designs.get_run", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_design_run_op() {}

#[pnp_operation(id = "sp.site_designs.get_run_status", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_design_run_status_op() {}

/// Creates a new site design available to users when they create a new site
/// from the SharePoint home page.
///
/// Ported from PnPjs `_SiteDesigns.createSiteDesign`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn create_site_design(
    client: &Client<'_>,
    site_url: &str,
    creation_info: &SiteDesignCreationInfo,
) -> Result<SiteDesignInfo> {
    let url = site_designs_url(site_url, "CreateSiteDesign");
    let body = serde_json::to_vec(&json!({ "info": creation_info }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.create"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Applies a site design to an existing site collection.
///
/// Ported from PnPjs `_SiteDesigns.applySiteDesign`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn apply_site_design(
    client: &Client<'_>,
    site_url: &str,
    site_design_id: &str,
    web_url: &str,
) -> Result<()> {
    let url = site_designs_url(site_url, "ApplySiteDesign");
    let body = serde_json::to_vec(&json!({ "siteDesignId": site_design_id, "webUrl": web_url }))?;
    client
        .run_ladder(
            entry("sp.site_designs.apply"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(())
}

/// Gets the list of available site designs.
///
/// Ported from PnPjs `_SiteDesigns.getSiteDesigns`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_designs(client: &Client<'_>, site_url: &str) -> Result<Vec<SiteDesignInfo>> {
    let url = site_designs_url(site_url, "GetSiteDesigns");
    let body = serde_json::to_vec(&json!({}))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.get_all"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets information about a specific site design.
///
/// Ported from PnPjs `_SiteDesigns.getSiteDesignMetadata`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_design_metadata(
    client: &Client<'_>,
    site_url: &str,
    id: &str,
) -> Result<SiteDesignInfo> {
    let url = site_designs_url(site_url, "GetSiteDesignMetadata");
    let body = serde_json::to_vec(&json!({ "id": id }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.get_metadata"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Updates a site design with new values. All fields beyond the site design
/// Id are optional in the REST call, but resetting `IsDefault` requires
/// resending it (SharePoint resets it to `false` when omitted).
///
/// Ported from PnPjs `_SiteDesigns.updateSiteDesign`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn update_site_design(
    client: &Client<'_>,
    site_url: &str,
    update_info: &SiteDesignUpdateInfo,
) -> Result<SiteDesignInfo> {
    let url = site_designs_url(site_url, "UpdateSiteDesign");
    let body = serde_json::to_vec(&json!({ "updateInfo": update_info }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.update"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes a site design.
///
/// Ported from PnPjs `_SiteDesigns.deleteSiteDesign`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn delete_site_design(client: &Client<'_>, site_url: &str, id: &str) -> Result<()> {
    let url = site_designs_url(site_url, "DeleteSiteDesign");
    let body = serde_json::to_vec(&json!({ "id": id }))?;
    client
        .run_ladder(
            entry("sp.site_designs.delete"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(())
}

/// Gets a list of principals that have access to a site design.
///
/// Ported from PnPjs `_SiteDesigns.getSiteDesignRights`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_design_rights(
    client: &Client<'_>,
    site_url: &str,
    id: &str,
) -> Result<Vec<SiteDesignPrincipals>> {
    let url = site_designs_url(site_url, "GetSiteDesignRights");
    let body = serde_json::to_vec(&json!({ "id": id }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.get_rights"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Grants access to a site design for one or more principals. `granted_rights`
/// is always `1` (the View right) per PnPjs's default and current SharePoint
/// support.
///
/// Ported from PnPjs `_SiteDesigns.grantSiteDesignRights`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn grant_site_design_rights(
    client: &Client<'_>,
    site_url: &str,
    id: &str,
    principal_names: &[String],
    granted_rights: u32,
) -> Result<()> {
    let url = site_designs_url(site_url, "GrantSiteDesignRights");
    let body = serde_json::to_vec(&json!({
        "grantedRights": granted_rights.to_string(),
        "id": id,
        "principalNames": principal_names,
    }))?;
    client
        .run_ladder(
            entry("sp.site_designs.grant_rights"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(())
}

/// Revokes access from a site design for one or more principals. If all
/// principals have rights revoked, the site design becomes viewable to
/// everyone.
///
/// Ported from PnPjs `_SiteDesigns.revokeSiteDesignRights`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn revoke_site_design_rights(
    client: &Client<'_>,
    site_url: &str,
    id: &str,
    principal_names: &[String],
) -> Result<()> {
    let url = site_designs_url(site_url, "RevokeSiteDesignRights");
    let body = serde_json::to_vec(&json!({ "id": id, "principalNames": principal_names }))?;
    client
        .run_ladder(
            entry("sp.site_designs.revoke_rights"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(())
}

/// Adds a site design task on the specified web to be invoked asynchronously.
///
/// Ported from PnPjs `_SiteDesigns.addSiteDesignTask`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn add_site_design_task(
    client: &Client<'_>,
    site_url: &str,
    web_url: &str,
    site_design_id: &str,
) -> Result<SiteDesignTask> {
    let url = site_designs_url(site_url, "AddSiteDesignTask");
    let body = serde_json::to_vec(&json!({ "webUrl": web_url, "siteDesignId": site_design_id }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.add_task"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Adds a site design task on the current web to be invoked asynchronously.
///
/// Ported from PnPjs `_SiteDesigns.addSiteDesignTaskToCurrentWeb`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn add_site_design_task_to_current_web(
    client: &Client<'_>,
    site_url: &str,
    site_design_id: &str,
) -> Result<SiteDesignTask> {
    let url = site_designs_url(site_url, "AddSiteDesignTaskToCurrentWeb");
    let body = serde_json::to_vec(&json!({ "siteDesignId": site_design_id }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.add_task_current_web"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Retrieves the site design task. Returns `None` once the task has finished
/// running and SharePoint no longer reports an `ID` for it (mirrors PnPjs's
/// `hOP(task, "ID") ? task : null` check).
///
/// Ported from PnPjs `_SiteDesigns.getSiteDesignTask`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_design_task(
    client: &Client<'_>,
    site_url: &str,
    id: &str,
) -> Result<Option<SiteDesignTask>> {
    let url = site_designs_url(site_url, "GetSiteDesignTask");
    let body = serde_json::to_vec(&json!({ "taskId": id }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.get_task"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    let value: serde_json::Value = serde_json::from_slice(&bytes)?;
    if value.get("ID").is_none() {
        return Ok(None);
    }
    Ok(Some(serde_json::from_value(value)?))
}

/// Retrieves a list of site design runs that have run on a specific web.
/// `site_design_id` is optional; when `None`, all runs are returned.
///
/// Ported from PnPjs `_SiteDesigns.getSiteDesignRun`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_design_run(
    client: &Client<'_>,
    site_url: &str,
    web_url: &str,
    site_design_id: Option<&str>,
) -> Result<Vec<SiteDesignRun>> {
    let url = site_designs_url(site_url, "GetSiteDesignRun");
    let body = serde_json::to_vec(&json!({ "webUrl": web_url, "siteDesignId": site_design_id }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.get_run"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Retrieves the status of a site design that has been run or is still
/// running.
///
/// Ported from PnPjs `_SiteDesigns.getSiteDesignRunStatus`
/// (`packages/sp/site-designs/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_site_design_run_status(
    client: &Client<'_>,
    site_url: &str,
    web_url: &str,
    run_id: &str,
) -> Result<Vec<SiteScriptActionStatus>> {
    let url = site_designs_url(site_url, "GetSiteDesignRunStatus");
    let body = serde_json::to_vec(&json!({ "webUrl": web_url, "runId": run_id }))?;
    let bytes = client
        .run_ladder(
            entry("sp.site_designs.get_run_status"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
