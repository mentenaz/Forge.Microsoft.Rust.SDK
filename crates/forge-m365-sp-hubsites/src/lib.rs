use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::Deserialize;

/// Subset of SharePoint's `SP.HubSite` fields, hand-ported from PnPjs's
/// `IHubSiteInfo` (`packages/sp/hubsites/types.ts`). Note the field is
/// `ID`, not `Id` — that's SharePoint's actual wire casing for this entity,
/// not a typo.
#[derive(Debug, Clone, Deserialize)]
pub struct HubSiteInfo {
    #[serde(rename = "ID")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "SiteId")]
    pub site_id: String,
    #[serde(rename = "SiteUrl")]
    pub site_url: String,
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "LogoUrl")]
    pub logo_url: String,
    #[serde(rename = "RequiresJoinApproval")]
    pub requires_join_approval: bool,
    #[serde(rename = "HideNameInNavigation")]
    pub hide_name_in_navigation: bool,
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-hubsites linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.hubsites.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_hub_sites_op() {}

#[pnp_operation(id = "sp.hubsites.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_hub_site_by_id_op() {}

/// Gets every hub site visible from this context.
///
/// Ported from PnPjs `HubSites` collection reads
/// (`packages/sp/hubsites/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_hub_sites(client: &Client<'_>, site_url: &str) -> Result<Vec<HubSiteInfo>> {
    let url = format!("{}/_api/hubsites", site_url.trim_end_matches('/'));
    let bytes = client
        .run_ladder(entry("sp.hubsites.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a hub site by its GUID id.
///
/// Ported from PnPjs `HubSites.getById` (`packages/sp/hubsites/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_hub_site_by_id(
    client: &Client<'_>,
    site_url: &str,
    hub_site_id: &str,
) -> Result<HubSiteInfo> {
    let url = format!(
        "{}/_api/hubsites/GetById?hubSiteId='{hub_site_id}'",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.hubsites.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
