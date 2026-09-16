use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::{Deserialize, Serialize};

/// Ported from PnPjs `IRelatedItem`
/// (`packages/sp/related-items/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RelatedItem {
    #[serde(rename = "ListId")]
    pub list_id: String,
    #[serde(rename = "ItemId")]
    pub item_id: i64,
    #[serde(rename = "Url")]
    pub url: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "WebId")]
    pub web_id: String,
    #[serde(rename = "IconUrl")]
    pub icon_url: String,
}

fn manager_url(site_url: &str) -> String {
    format!(
        "{}/_api/SP.RelatedItemManager",
        site_url.trim_end_matches('/')
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-related-items linked in?")
        })
}

#[pnp_operation(
    id = "sp.related_items.get_related_items",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_related_items_op() {}

#[pnp_operation(
    id = "sp.related_items.get_page_one_related_items",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_page_one_related_items_op() {}

#[pnp_operation(
    id = "sp.related_items.add_single_link",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn add_single_link_op() {}

#[pnp_operation(
    id = "sp.related_items.add_single_link_to_url",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn add_single_link_to_url_op() {}

#[pnp_operation(
    id = "sp.related_items.add_single_link_from_url",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn add_single_link_from_url_op() {}

#[pnp_operation(
    id = "sp.related_items.delete_single_link",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn delete_single_link_op() {}

/// Gets all the related items for the given item specification.
///
/// Ported from PnPjs `_RelatedItemManager.getRelatedItems`
/// (`packages/sp/related-items/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_related_items(
    client: &Client<'_>,
    site_url: &str,
    source_list_name: &str,
    source_item_id: i64,
) -> Result<Vec<RelatedItem>> {
    let url = format!("{}.GetRelatedItems", manager_url(site_url));
    let body = serde_json::json!({
        "SourceItemID": source_item_id,
        "SourceListName": source_list_name,
    });
    let bytes = client
        .run_ladder(
            entry("sp.related_items.get_related_items"),
            "POST",
            &url,
            &[],
            Some(serde_json::to_vec(&body)?.as_slice()),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the first page of related items for the given item specification.
///
/// Ported from PnPjs `_RelatedItemManager.getPageOneRelatedItems`
/// (`packages/sp/related-items/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_page_one_related_items(
    client: &Client<'_>,
    site_url: &str,
    source_list_name: &str,
    source_item_id: i64,
) -> Result<Vec<RelatedItem>> {
    let url = format!("{}.GetPageOneRelatedItems", manager_url(site_url));
    let body = serde_json::json!({
        "SourceItemID": source_item_id,
        "SourceListName": source_list_name,
    });
    let bytes = client
        .run_ladder(
            entry("sp.related_items.get_page_one_related_items"),
            "POST",
            &url,
            &[],
            Some(serde_json::to_vec(&body)?.as_slice()),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Adds a single related-item link using full source/target specifications.
///
/// Ported from PnPjs `_RelatedItemManager.addSingleLink`
/// (`packages/sp/related-items/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[allow(clippy::too_many_arguments)]
pub async fn add_single_link(
    client: &Client<'_>,
    site_url: &str,
    source_list_name: &str,
    source_item_id: i64,
    source_web_url: &str,
    target_list_name: &str,
    target_item_id: i64,
    target_web_url: &str,
    try_add_reverse_link: bool,
) -> Result<()> {
    let url = format!("{}.AddSingleLink", manager_url(site_url));
    let body = serde_json::json!({
        "SourceItemID": source_item_id,
        "SourceListName": source_list_name,
        "SourceWebUrl": source_web_url,
        "TargetItemID": target_item_id,
        "TargetListName": target_list_name,
        "TargetWebUrl": target_web_url,
        "TryAddReverseLink": try_add_reverse_link,
    });
    client
        .run_ladder(
            entry("sp.related_items.add_single_link"),
            "POST",
            &url,
            &[],
            Some(serde_json::to_vec(&body)?.as_slice()),
        )
        .await?;
    Ok(())
}

/// Adds a related-item link from an item specified by list name/item id to
/// an item specified by url.
///
/// Ported from PnPjs `_RelatedItemManager.addSingleLinkToUrl`
/// (`packages/sp/related-items/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn add_single_link_to_url(
    client: &Client<'_>,
    site_url: &str,
    source_list_name: &str,
    source_item_id: i64,
    target_item_url: &str,
    try_add_reverse_link: bool,
) -> Result<()> {
    let url = format!("{}.AddSingleLinkToUrl", manager_url(site_url));
    let body = serde_json::json!({
        "SourceItemID": source_item_id,
        "SourceListName": source_list_name,
        "TargetItemUrl": target_item_url,
        "TryAddReverseLink": try_add_reverse_link,
    });
    client
        .run_ladder(
            entry("sp.related_items.add_single_link_to_url"),
            "POST",
            &url,
            &[],
            Some(serde_json::to_vec(&body)?.as_slice()),
        )
        .await?;
    Ok(())
}

/// Adds a related-item link from an item specified by url to an item
/// specified by list name/item id.
///
/// Ported from PnPjs `_RelatedItemManager.addSingleLinkFromUrl`
/// (`packages/sp/related-items/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn add_single_link_from_url(
    client: &Client<'_>,
    site_url: &str,
    source_item_url: &str,
    target_list_name: &str,
    target_item_id: i64,
    try_add_reverse_link: bool,
) -> Result<()> {
    let url = format!("{}.AddSingleLinkFromUrl", manager_url(site_url));
    let body = serde_json::json!({
        "SourceItemUrl": source_item_url,
        "TargetItemID": target_item_id,
        "TargetListName": target_list_name,
        "TryAddReverseLink": try_add_reverse_link,
    });
    client
        .run_ladder(
            entry("sp.related_items.add_single_link_from_url"),
            "POST",
            &url,
            &[],
            Some(serde_json::to_vec(&body)?.as_slice()),
        )
        .await?;
    Ok(())
}

/// Deletes a single related-item link.
///
/// Ported from PnPjs `_RelatedItemManager.deleteSingleLink`
/// (`packages/sp/related-items/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[allow(clippy::too_many_arguments)]
pub async fn delete_single_link(
    client: &Client<'_>,
    site_url: &str,
    source_list_name: &str,
    source_item_id: i64,
    source_web_url: &str,
    target_list_name: &str,
    target_item_id: i64,
    target_web_url: &str,
    try_delete_reverse_link: bool,
) -> Result<()> {
    let url = format!("{}.DeleteSingleLink", manager_url(site_url));
    let body = serde_json::json!({
        "SourceItemID": source_item_id,
        "SourceListName": source_list_name,
        "SourceWebUrl": source_web_url,
        "TargetItemID": target_item_id,
        "TargetListName": target_list_name,
        "TargetWebUrl": target_web_url,
        "TryDeleteReverseLink": try_delete_reverse_link,
    });
    client
        .run_ladder(
            entry("sp.related_items.delete_single_link"),
            "POST",
            &url,
            &[],
            Some(serde_json::to_vec(&body)?.as_slice()),
        )
        .await?;
    Ok(())
}
