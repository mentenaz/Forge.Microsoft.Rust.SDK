use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use forge_m365_sp_site_users::SiteUserInfo;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;

/// Subset of SharePoint's `SP.Group` fields, hand-ported from PnPjs's
/// `ISiteGroupInfo` (`packages/sp/site-groups/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct SiteGroupInfo {
    #[serde(rename = "Id")]
    pub id: i64,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "LoginName")]
    pub login_name: String,
    #[serde(rename = "OwnerTitle")]
    pub owner_title: String,
    #[serde(rename = "IsHiddenInUI")]
    pub is_hidden_in_ui: bool,
    #[serde(rename = "PrincipalType")]
    pub principal_type: i64,
}

/// Doubles embedded `'` and percent-encodes for use inside a `getbyname('...')`
/// path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-site-groups linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.site_groups.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_all_op() {}

#[pnp_operation(id = "sp.site_groups.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_id_op() {}

#[pnp_operation(id = "sp.site_groups.get_by_name", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_name_op() {}

#[pnp_operation(id = "sp.site_groups.get_members", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_members_op() {}

/// Gets every site group.
///
/// Ported from PnPjs `Web.siteGroups` via `_SiteGroups`
/// (`packages/sp/site-groups/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_site_groups(client: &Client<'_>, site_url: &str) -> Result<Vec<SiteGroupInfo>> {
    let url = format!("{}/_api/web/sitegroups", site_url.trim_end_matches('/'));
    let bytes = client
        .run_ladder(entry("sp.site_groups.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a group by its SharePoint group id.
///
/// Ported from PnPjs `SiteGroups.getById` (`packages/sp/site-groups/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_group_by_id(
    client: &Client<'_>,
    site_url: &str,
    id: i64,
) -> Result<SiteGroupInfo> {
    let url = format!(
        "{}/_api/web/sitegroups({id})",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.site_groups.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a group by name.
///
/// Ported from PnPjs `SiteGroups.getByName` (`packages/sp/site-groups/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_group_by_name(
    client: &Client<'_>,
    site_url: &str,
    name: &str,
) -> Result<SiteGroupInfo> {
    let url = format!(
        "{}/_api/web/sitegroups/getbyname('{}')",
        site_url.trim_end_matches('/'),
        encode_segment(name)
    );
    let bytes = client
        .run_ladder(entry("sp.site_groups.get_by_name"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the members of a group. Reuses `forge_m365_sp_site_users::SiteUserInfo`
/// since SharePoint returns the same `SP.User` shape here as it does for
/// site users.
///
/// Ported from PnPjs `SiteGroup.users` (`packages/sp/site-groups/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_group_members(
    client: &Client<'_>,
    site_url: &str,
    group_id: i64,
) -> Result<Vec<SiteUserInfo>> {
    let url = format!(
        "{}/_api/web/sitegroups({group_id})/users",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.site_groups.get_members"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
