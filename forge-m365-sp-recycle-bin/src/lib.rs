use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::Deserialize;

/// Subset of SharePoint's recycle-bin item fields, hand-ported from PnPjs's
/// `IRecycleBinItemObject` (`packages/sp/recycle-bin/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct RecycleBinItemInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "LeafName")]
    pub leaf_name: String,
    #[serde(rename = "DirName")]
    pub dir_name: String,
    #[serde(rename = "DeletedDate")]
    pub deleted_date: String,
    #[serde(rename = "DeletedByName")]
    pub deleted_by_name: String,
    #[serde(rename = "Size")]
    pub size: i64,
    #[serde(rename = "ItemType")]
    pub item_type: i64,
    #[serde(rename = "ItemState")]
    pub item_state: i64,
}

fn recycle_bin_item_url(site_url: &str, item_id: &str) -> String {
    format!(
        "{}/_api/web/RecycleBin('{item_id}')",
        site_url.trim_end_matches('/')
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-recycle-bin linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.recycle_bin.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_recycle_bin_items_op() {}

#[pnp_operation(id = "sp.recycle_bin.restore", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn restore_recycle_bin_item_op() {}

#[pnp_operation(
    id = "sp.recycle_bin.delete_object",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn delete_recycle_bin_item_op() {}

/// Gets the items currently in the site's (first-stage) recycle bin.
///
/// Ported from PnPjs `Web.recycleBin` via `_RecycleBin`
/// (`packages/sp/recycle-bin/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_recycle_bin_items(
    client: &Client<'_>,
    site_url: &str,
) -> Result<Vec<RecycleBinItemInfo>> {
    let url = format!("{}/_api/web/RecycleBin", site_url.trim_end_matches('/'));
    let bytes = client
        .run_ladder(entry("sp.recycle_bin.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Restores a recycle-bin item to its original location. Bodiless POST.
///
/// Ported from PnPjs `RecycleBinItem.restore`
/// (`packages/sp/recycle-bin/types.ts`) @ pnpjs `8ee2375d`.
pub async fn restore_recycle_bin_item(
    client: &Client<'_>,
    site_url: &str,
    item_id: &str,
) -> Result<()> {
    let url = format!("{}/Restore", recycle_bin_item_url(site_url, item_id));
    client
        .run_ladder(entry("sp.recycle_bin.restore"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Permanently deletes a single recycle-bin item (bypasses the second-stage
/// recycle bin — this is not reversible). Bodiless POST.
///
/// Ported from PnPjs `RecycleBinItem.delete`
/// (`packages/sp/recycle-bin/types.ts`) @ pnpjs `8ee2375d`.
pub async fn delete_recycle_bin_item(
    client: &Client<'_>,
    site_url: &str,
    item_id: &str,
) -> Result<()> {
    let url = format!("{}/DeleteObject", recycle_bin_item_url(site_url, item_id));
    client
        .run_ladder(
            entry("sp.recycle_bin.delete_object"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}
