use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};

/// Status/result of a group-connected site creation or lookup.
///
/// Ported from PnPjs `IGroupSiteInfo` (`packages/sp/groupsitemanager/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GroupSiteInfo {
    #[serde(rename = "DocumentsUrl")]
    pub documents_url: String,
    #[serde(rename = "ErrorMessage")]
    pub error_message: Option<String>,
    #[serde(rename = "GroupId")]
    pub group_id: String,
    #[serde(rename = "SiteStatus")]
    pub site_status: i32,
    #[serde(rename = "SiteUrl")]
    pub site_url: String,
}

/// Ported from PnPjs `IEnsureTeamForGroupExResponse`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EnsureTeamForGroupExResponse {
    #[serde(rename = "teamsId")]
    pub teams_id: String,
    #[serde(rename = "teamsUrl")]
    pub teams_url: String,
}

/// Ported from PnPjs `IOrgLabelsContext`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OrgLabel {
    #[serde(rename = "DisplayName")]
    pub display_name: String,
    #[serde(rename = "LabelApplicableTo")]
    pub label_applicable_to: String,
    #[serde(rename = "ObjectId")]
    pub object_id: String,
}

/// Ported from PnPjs `IOrgLabelsContextList`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OrgLabelsContextList {
    #[serde(rename = "IsLastPage")]
    pub is_last_page: bool,
    #[serde(rename = "Labels")]
    pub labels: Vec<OrgLabel>,
}

/// Ported from PnPjs `IGroupSiteConversionInfo`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GroupSiteConversionInfo {
    #[serde(rename = "GroupType")]
    pub group_type: i32,
    #[serde(rename = "IsGroupifyDisabled")]
    pub is_groupify_disabled: bool,
    #[serde(rename = "IsRegionRestricted")]
    pub is_region_restricted: bool,
    #[serde(rename = "IsWrongPdl")]
    pub is_wrong_pdl: bool,
    #[serde(rename = "SuggestedMembers")]
    pub suggested_members: Vec<String>,
    #[serde(rename = "SuggestedOwners")]
    pub suggested_owners: Vec<String>,
    #[serde(rename = "UnsuggestablePrincipals")]
    pub unsuggestable_principals: Vec<String>,
}

/// Raw JSON: PnPjs's `IGroupCreationContext` mixes typed fields with
/// `{ Key, Value: any, ValueType }` tuples whose `Value` shape is
/// tenant-configuration-dependent and not documented, so (as with
/// `sp-appcatalog::AppInfo`) this is left as untyped JSON rather than
/// guessing field types.
pub type GroupCreationContext = Map<String, Value>;

fn group_site_manager_url(site_url: &str, method: &str) -> String {
    format!(
        "{}/_api/groupsitemanager/{method}",
        site_url.trim_end_matches('/')
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-groupsitemanager linked in?")
        })
}

#[pnp_operation(id = "sp.groupsitemanager.can_user_create_group", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn can_user_create_group_op() {}

#[pnp_operation(id = "sp.groupsitemanager.clear_current_user_teams_cache", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn clear_current_user_teams_cache_op() {}

#[pnp_operation(id = "sp.groupsitemanager.create", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_op() {}

#[pnp_operation(id = "sp.groupsitemanager.create_group", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_group_op() {}

#[pnp_operation(id = "sp.groupsitemanager.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_op() {}

#[pnp_operation(id = "sp.groupsitemanager.ensure_team_for_group", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn ensure_team_for_group_op() {}

#[pnp_operation(id = "sp.groupsitemanager.ensure_team_for_group_ex", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn ensure_team_for_group_ex_op() {}

#[pnp_operation(id = "sp.groupsitemanager.get_all_org_labels", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_all_org_labels_op() {}

#[pnp_operation(id = "sp.groupsitemanager.get_group_creation_context", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_group_creation_context_op() {}

#[pnp_operation(id = "sp.groupsitemanager.get_group_site_conversion_data", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_group_site_conversion_data_op() {}

#[pnp_operation(id = "sp.groupsitemanager.get_site_status", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_status_op() {}

#[pnp_operation(id = "sp.groupsitemanager.get_valid_site_url_from_alias", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_valid_site_url_from_alias_op() {}

#[pnp_operation(id = "sp.groupsitemanager.is_teamify_prompt_hidden", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn is_teamify_prompt_hidden_op() {}

#[pnp_operation(id = "sp.groupsitemanager.notebook", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn notebook_op() {}

/// Indicates whether the current caller can create Microsoft 365 groups.
///
/// Ported from PnPjs `GroupSiteManager.canUserCreateGroup`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn can_user_create_group(client: &Client<'_>, site_url: &str) -> Result<bool> {
    let url = group_site_manager_url(site_url, "CanUserCreateGroup");
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.can_user_create_group"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Clears the Teams cache for the current user/application. Bodiless POST.
///
/// Ported from PnPjs `GroupSiteManager.clearCurrentUserTeamsCache`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn clear_current_user_teams_cache(client: &Client<'_>, site_url: &str) -> Result<()> {
    let url = group_site_manager_url(site_url, "ClearCurrentUserTeamsCache");
    client
        .run_ladder(
            entry("sp.groupsitemanager.clear_current_user_teams_cache"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

/// Creates a SharePoint team site for the given Microsoft 365 group, or
/// returns its current creation status if already in progress.
///
/// Ported from PnPjs `GroupSiteManager.create`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn create(client: &Client<'_>, site_url: &str, group_id: &str) -> Result<GroupSiteInfo> {
    let url = group_site_manager_url(site_url, "Create");
    let body = serde_json::json!({ "groupId": group_id });
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.create"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Creates a Microsoft 365 group with a connected site. Does not work under
/// app-only (Azure AD application) auth per PnPjs's own notes.
///
/// Ported from PnPjs `GroupSiteManager.createGroup`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
#[allow(clippy::too_many_arguments)]
pub async fn create_group(
    client: &Client<'_>,
    site_url: &str,
    display_name: &str,
    alias: &str,
    is_public: bool,
    owner_principal_names: &[String],
    description: &str,
    creation_options: &[String],
) -> Result<GroupSiteInfo> {
    let url = group_site_manager_url(site_url, "CreateGroup");
    let body = serde_json::json!({
        "displayName": display_name,
        "alias": alias,
        "isPublic": is_public,
        "ownerPrincipalNames": owner_principal_names,
        "description": description,
        "creationOptions": creation_options,
    });
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.create_group"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes a group-connected site. Does not work under app-only auth per
/// PnPjs's own notes.
///
/// Ported from PnPjs `GroupSiteManager.delete`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn delete(client: &Client<'_>, site_url: &str, target_site_url: &str) -> Result<()> {
    let url = group_site_manager_url(site_url, "Delete");
    let body = serde_json::json!({ "siteUrl": target_site_url });
    client
        .run_ladder(
            entry("sp.groupsitemanager.delete"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(())
}

/// Creates a Team for the current (group-connected) site. Bodiless POST.
/// Response is assumed to be a bare JSON string (the team URL) per PnPjs's
/// `Promise<string>` return type — unverified against a live tenant.
///
/// Ported from PnPjs `GroupSiteManager.ensureTeamForGroup`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn ensure_team_for_group(client: &Client<'_>, site_url: &str) -> Result<String> {
    let url = group_site_manager_url(site_url, "EnsureTeamForGroup");
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.ensure_team_for_group"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Creates a Team for the current (group-connected) site, returning its id
/// and URL. Bodiless POST.
///
/// Ported from PnPjs `GroupSiteManager.ensureTeamForGroupEx`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn ensure_team_for_group_ex(
    client: &Client<'_>,
    site_url: &str,
) -> Result<EnsureTeamForGroupExResponse> {
    let url = group_site_manager_url(site_url, "EnsureTeamForGroupEx");
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.ensure_team_for_group_ex"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets tenant sensitivity/classification labels, one page at a time.
///
/// Ported from PnPjs `GroupSiteManager.getAllOrgLabels`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_all_org_labels(
    client: &Client<'_>,
    site_url: &str,
    page_number: i32,
) -> Result<OrgLabelsContextList> {
    let url = group_site_manager_url(site_url, "GetAllOrgLabels");
    let body = serde_json::json!({ "pageNumber": page_number });
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.get_all_org_labels"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets tenant-wide Microsoft 365 group creation configuration.
///
/// Ported from PnPjs `GroupSiteManager.getGroupCreationContext`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_group_creation_context(
    client: &Client<'_>,
    site_url: &str,
) -> Result<GroupCreationContext> {
    let url = group_site_manager_url(site_url, "GetGroupCreationContext");
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.get_group_creation_context"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets groupification configuration/eligibility for the current site.
///
/// Ported from PnPjs `GroupSiteManager.getGroupSiteConversionData`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_group_site_conversion_data(
    client: &Client<'_>,
    site_url: &str,
) -> Result<GroupSiteConversionInfo> {
    let url = group_site_manager_url(site_url, "GetGroupSiteConversionData");
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.get_group_site_conversion_data"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets group-connected site creation status for the given group.
///
/// Ported from PnPjs `GroupSiteManager.getSiteStatus`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_site_status(
    client: &Client<'_>,
    site_url: &str,
    group_id: &str,
) -> Result<GroupSiteInfo> {
    let url = group_site_manager_url(site_url, "GetSiteStatus");
    let body = serde_json::json!({ "groupId": group_id });
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.get_site_status"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Resolves a valid SharePoint site URL from a desired alias. Response is
/// assumed to be a bare JSON string per PnPjs's `Promise<string>` return
/// type — unverified against a live tenant.
///
/// Ported from PnPjs `GroupSiteManager.getValidSiteUrlFromAlias`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_valid_site_url_from_alias(
    client: &Client<'_>,
    site_url: &str,
    alias: &str,
    managed_path: Option<&str>,
    is_team_site: Option<bool>,
) -> Result<String> {
    let url = group_site_manager_url(site_url, "GetValidSiteUrlFromAlias");
    let body = serde_json::json!({
        "alias": alias,
        "managedPath": managed_path,
        "isTeamSite": is_team_site,
    });
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.get_valid_site_url_from_alias"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Indicates whether the "Teamify" prompt is hidden on a group-connected
/// site. `site_url_target` defaults to the current site context when `None`.
///
/// Ported from PnPjs `GroupSiteManager.isTeamifyPromptHidden`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn is_teamify_prompt_hidden(
    client: &Client<'_>,
    site_url: &str,
    site_url_target: Option<&str>,
) -> Result<bool> {
    let url = group_site_manager_url(site_url, "IsTeamifyPromptHidden");
    let body = serde_json::json!({ "siteUrl": site_url_target });
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.is_teamify_prompt_hidden"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the group-connected site's default OneNote notebook URL. Response
/// is assumed to be a bare JSON string per PnPjs's `Promise<string>` return
/// type — unverified against a live tenant.
///
/// Ported from PnPjs `GroupSiteManager.notebook`
/// (`packages/sp/groupsitemanager/types.ts`) @ pnpjs `8ee2375d`.
pub async fn notebook(client: &Client<'_>, site_url: &str, group_id: &str) -> Result<String> {
    let url = group_site_manager_url(site_url, "Notebook");
    let body = serde_json::json!({ "groupId": group_id });
    let bytes = client
        .run_ladder(
            entry("sp.groupsitemanager.notebook"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
