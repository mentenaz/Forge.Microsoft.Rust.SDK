use forge_m365_core::{registered_operations, Client, OperationEntry, Result};
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde_json::{Map, Value};

use crate::types::{
    ItemDeleteParams, ItemParentId, ItemParentInfos, ListItemFormUpdateValue, ParentListInfo,
    ParentWebInfo, ResourcePath,
};

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

fn item_url(site_url: &str, list_title: &str, item_id: i64) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')/items({item_id})",
        site_url.trim_end_matches('/'),
        encode_list_title(list_title)
    )
}

/// Gets a list item by its integer ID.
///
/// Ported from PnPjs `Items.getById` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_by_id(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Map<String, Value>> {
    let url = item_url(site_url, list_title, item_id);
    let bytes = client
        .run_ladder(entry("sp.items.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a BCS item by its string identifier.
///
/// Ported from PnPjs `Items.getItemByStringId` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_by_string_id(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    string_id: &str,
) -> Result<Map<String, Value>> {
    let encoded_id = string_id.replace('\'', "''");
    let url = format!(
        "{}/_api/web/lists/getbytitle('{}')/items/getItemByStringId('{}')",
        site_url.trim_end_matches('/'),
        encode_list_title(list_title),
        encoded_id,
    );
    let bytes = client
        .run_ladder(entry("sp.items.get_by_string_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Moves the list item to the Recycle Bin and returns the identifier of the
/// new Recycle Bin item.
///
/// Ported from PnPjs `Item.recycle` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn recycle(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<String> {
    let url = format!("{}/recycle", item_url(site_url, list_title, item_id));
    let bytes = client
        .run_ladder(entry("sp.items.recycle"), "POST", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes the item with the specified options (e.g. bypass shared lock).
///
/// Ported from PnPjs `Item.deleteWithParams` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn delete_with_params(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    bypass_shared_lock: bool,
) -> Result<()> {
    let url = format!(
        "{}/DeleteWithParameters",
        item_url(site_url, list_title, item_id)
    );
    let params = ItemDeleteParams { bypass_shared_lock };
    let body = serde_json::json!({ "parameters": params });
    let payload = serde_json::to_vec(&body)?;
    client
        .run_ladder(
            entry("sp.items.delete_with_params"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(())
}

/// Gets a string representation of the full URL to the WOPI frame.
///
/// `action`: 0 = view, 1 = edit, 2 = mobileView, 3 = interactivePreview.
///
/// Ported from PnPjs `Item.getWopiFrameUrl` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_wopi_frame_url(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    action: i64,
) -> Result<String> {
    let url = format!(
        "{}/getWOPIFrameUrl(@action)?@action={action}",
        item_url(site_url, list_title, item_id),
    );
    let bytes = client
        .run_ladder(
            entry("sp.items.get_wopi_frame_url"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Validates and sets the values of the specified collection of fields for
/// the list item.
///
/// Ported from PnPjs `Item.validateUpdateListItem`
/// (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn validate_update_list_item(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    form_values: &[ListItemFormUpdateValue],
    b_new_document_update: bool,
) -> Result<Vec<ListItemFormUpdateValue>> {
    let url = format!(
        "{}/validateupdatelistitem",
        item_url(site_url, list_title, item_id)
    );
    let body = serde_json::json!({
        "formValues": form_values,
        "bNewDocumentUpdate": b_new_document_update,
    });
    let payload = serde_json::to_vec(&body)?;
    let bytes = client
        .run_ladder(
            entry("sp.items.validate_update_list_item"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the parent information for this item's list and web.
///
/// The raw SharePoint REST response nests `ParentWeb` inside `ParentList`.
/// This function reshapes it into the flat `ItemParentInfos` structure
/// matching the PnPjs convention.
///
/// Ported from PnPjs `Item.getParentInfos` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_parent_infos(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<ItemParentInfos> {
    let url = format!(
        "{}?$select=Id,ParentList/Id,ParentList/Title,ParentList/RootFolder/UniqueId,ParentList/RootFolder/ServerRelativeUrl,ParentList/RootFolder/ServerRelativePath,ParentList/ParentWeb/Id,ParentList/ParentWeb/Url,ParentList/ParentWeb/ServerRelativeUrl,ParentList/ParentWeb/ServerRelativePath&$expand=ParentList,ParentList/RootFolder,ParentList/ParentWeb",
        item_url(site_url, list_title, item_id),
    );
    let bytes = client
        .run_ladder(entry("sp.items.get_parent_infos"), "GET", &url, &[], None)
        .await?;
    let raw: serde_json::Value = serde_json::from_slice(&bytes)?;

    let item_id_val = raw["Id"].clone();
    let list = &raw["ParentList"];

    Ok(ItemParentInfos {
        item: ItemParentId {
            id: item_id_val.as_str().unwrap_or_default().to_string(),
        },
        parent_list: ParentListInfo {
            id: list["Id"].as_str().unwrap_or_default().to_string(),
            title: list["Title"].as_str().unwrap_or_default().to_string(),
            root_folder_server_relative_path: ResourcePath {
                decoded_url: list["RootFolder"]["ServerRelativePath"]["DecodedUrl"]
                    .as_str()
                    .unwrap_or_default()
                    .to_string(),
            },
            root_folder_server_relative_url: list["RootFolder"]["ServerRelativeUrl"]
                .as_str()
                .unwrap_or_default()
                .to_string(),
            root_folder_unique_id: list["RootFolder"]["UniqueId"]
                .as_str()
                .unwrap_or_default()
                .to_string(),
        },
        parent_web: ParentWebInfo {
            id: list["ParentWeb"]["Id"]
                .as_str()
                .unwrap_or_default()
                .to_string(),
            server_relative_path: ResourcePath {
                decoded_url: list["ParentWeb"]["ServerRelativePath"]["DecodedUrl"]
                    .as_str()
                    .unwrap_or_default()
                    .to_string(),
            },
            server_relative_url: list["ParentWeb"]["ServerRelativeUrl"]
                .as_str()
                .unwrap_or_default()
                .to_string(),
            url: list["ParentWeb"]["Url"]
                .as_str()
                .unwrap_or_default()
                .to_string(),
        },
    })
}
