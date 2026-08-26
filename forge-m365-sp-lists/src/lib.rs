use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;
use serde_json::{Map, Value};

/// Subset of SharePoint's `SP.List` fields, hand-ported from PnPjs's `IListInfo`
/// (`packages/sp/lists/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct ListInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "ItemCount")]
    pub item_count: i64,
    #[serde(rename = "BaseTemplate")]
    pub base_template: i64,
    #[serde(rename = "Hidden")]
    pub hidden: bool,
    /// The `__metadata.type` value item bodies must carry for this list to be
    /// recognized by `add_item`, e.g. `SP.Data.MyListListItem`.
    #[serde(rename = "ListItemEntityTypeFullName")]
    pub list_item_entity_type_full_name: String,
}

/// Doubles embedded `'` and percent-encodes for use inside an OData
/// `getByTitle('...')` path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_list_title(title: &str) -> String {
    let doubled = title.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn list_items_url(site_url: &str, list_title: &str) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')/items",
        site_url.trim_end_matches('/'),
        encode_list_title(list_title)
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-lists linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.lists.get_by_title", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_title_op() {}

#[pnp_operation(id = "sp.lists.items.get", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_items_op() {}

#[pnp_operation(id = "sp.lists.items.add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_item_op() {}

#[pnp_operation(id = "sp.lists.items.update", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn update_item_op() {}

#[pnp_operation(id = "sp.lists.items.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_item_op() {}

/// Gets a list's properties by title.
///
/// Ported from PnPjs `Lists.getByTitle` + `List.get` (`packages/sp/lists/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_list_by_title(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
) -> Result<ListInfo> {
    let url = format!(
        "{}/_api/web/lists/getbytitle('{}')?$select=Id,Title,ItemCount,BaseTemplate,Hidden,ListItemEntityTypeFullName",
        site_url.trim_end_matches('/'),
        encode_list_title(list_title)
    );
    let bytes = client
        .run_ladder(entry("sp.lists.get_by_title"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the items in a list. Item schemas are per-list and only known at
/// runtime, so results are returned as raw JSON objects rather than a typed
/// struct — same approach PnPjs takes (`Items.add`/collection reads return `any`).
///
/// Ported from PnPjs `Items` collection read (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_items(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
) -> Result<Vec<Map<String, Value>>> {
    let url = list_items_url(site_url, list_title);
    let bytes = client
        .run_ladder(entry("sp.lists.items.get"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Adds an item to a list. `fields` is the map of column-internal-name to
/// value (e.g. `{"Title": "New item"}`); the required `__metadata.type` is
/// fetched and attached automatically.
///
/// Ported from PnPjs `Items.add` (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn add_item(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    fields: Map<String, Value>,
) -> Result<Map<String, Value>> {
    let list = get_list_by_title(client, site_url, list_title).await?;

    let mut body = fields;
    let mut metadata = Map::new();
    metadata.insert(
        "type".to_string(),
        Value::String(list.list_item_entity_type_full_name),
    );
    body.insert("__metadata".to_string(), Value::Object(metadata));

    let url = list_items_url(site_url, list_title);
    let payload = serde_json::to_vec(&Value::Object(body))?;
    let bytes = client
        .run_ladder(
            entry("sp.lists.items.add"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Updates an existing list item. `fields` carries only the columns to
/// change; unlisted columns are left untouched (SharePoint MERGE semantics).
///
/// Ported from PnPjs `Item.update` (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn update_item(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    fields: Map<String, Value>,
) -> Result<()> {
    let url = format!("{}({item_id})", list_items_url(site_url, list_title));
    let payload = serde_json::to_vec(&Value::Object(fields))?;
    client
        .run_ladder(
            entry("sp.lists.items.update"),
            "PATCH",
            &url,
            &[("IF-MATCH", "*")],
            Some(&payload),
        )
        .await?;
    Ok(())
}

/// Deletes a list item.
///
/// Ported from PnPjs `Item.delete` (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn delete_item(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<()> {
    let url = format!("{}({item_id})", list_items_url(site_url, list_title));
    client
        .run_ladder(
            entry("sp.lists.items.delete"),
            "DELETE",
            &url,
            &[("IF-MATCH", "*")],
            None,
        )
        .await?;
    Ok(())
}
