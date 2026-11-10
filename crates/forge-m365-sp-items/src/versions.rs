use forge_m365_core::{registered_operations, Client, OperationEntry, Result};
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};

use crate::types::VersionInfo;

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-items linked in?")
        })
}

fn encode_list_title(title: &str) -> String {
    let doubled = title.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn versions_url(site_url: &str, list_title: &str, item_id: i64) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')/items({item_id})/versions",
        site_url.trim_end_matches('/'),
        encode_list_title(list_title)
    )
}

/// Gets the collection of versions associated with this item.
///
/// Ported from PnPjs `Item.versions` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_versions(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Vec<VersionInfo>> {
    let url = versions_url(site_url, list_title, item_id);
    let bytes = client
        .run_ladder(entry("sp.items.get_versions"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a specific version of a list item by its version ID.
///
/// Ported from PnPjs `ItemVersions.getById` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_version_by_id(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    version_id: i64,
) -> Result<VersionInfo> {
    let url = format!(
        "{}({version_id})",
        versions_url(site_url, list_title, item_id)
    );
    let bytes = client
        .run_ladder(entry("sp.items.get_version_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes a specific version of a list item.
///
/// Ported from PnPjs `ItemVersion.delete` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn delete_version(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    version_id: i64,
) -> Result<()> {
    let url = format!(
        "{}({version_id})",
        versions_url(site_url, list_title, item_id)
    );
    client
        .run_ladder(
            entry("sp.items.delete_version"),
            "DELETE",
            &url,
            &[("IF-MATCH", "*")],
            None,
        )
        .await?;
    Ok(())
}
