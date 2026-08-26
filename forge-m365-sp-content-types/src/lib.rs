use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;
use serde_json::{Map, Value};

/// Mirrors PnPjs's `Id: { StringValue: string }` shape rather than
/// flattening it, since that's the actual wire shape SharePoint returns.
#[derive(Debug, Clone, Deserialize)]
pub struct ContentTypeId {
    #[serde(rename = "StringValue")]
    pub string_value: String,
}

/// Subset of SharePoint's `SP.ContentType` fields, hand-ported from PnPjs's
/// `IContentTypeInfo` (`packages/sp/content-types/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct ContentTypeInfo {
    #[serde(rename = "Id")]
    pub id: ContentTypeId,
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "Group")]
    pub group: String,
    #[serde(rename = "Hidden")]
    pub hidden: bool,
    #[serde(rename = "ReadOnly")]
    pub read_only: bool,
    #[serde(rename = "Sealed")]
    pub sealed: bool,
    #[serde(rename = "Scope")]
    pub scope: String,
}

/// Doubles embedded `'` and percent-encodes for use inside a `getbytitle`/
/// content-type-id path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-content-types linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.content_types.get_web", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_web_content_types_op() {}

#[pnp_operation(id = "sp.content_types.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_content_type_by_id_op() {}

#[pnp_operation(id = "sp.content_types.get_list", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_list_content_types_op() {}

#[pnp_operation(id = "sp.content_types.add_to_list", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_content_type_to_list_op() {}

/// Gets the site's web-level content types.
///
/// Ported from PnPjs `Web.contentTypes` via `_ContentTypes`
/// (`packages/sp/content-types/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_web_content_types(
    client: &Client<'_>,
    site_url: &str,
) -> Result<Vec<ContentTypeInfo>> {
    let url = format!("{}/_api/web/contenttypes", site_url.trim_end_matches('/'));
    let bytes = client
        .run_ladder(entry("sp.content_types.get_web"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a single web-level content type by id (e.g. `0x0101`).
///
/// Ported from PnPjs `ContentTypes.getById` (`packages/sp/content-types/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_content_type_by_id(
    client: &Client<'_>,
    site_url: &str,
    content_type_id: &str,
) -> Result<ContentTypeInfo> {
    let url = format!(
        "{}/_api/web/contenttypes('{}')",
        site_url.trim_end_matches('/'),
        encode_segment(content_type_id)
    );
    let bytes = client
        .run_ladder(entry("sp.content_types.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the content types available on a specific list.
///
/// Ported from PnPjs `List.contentTypes` via `_ContentTypes`
/// (`packages/sp/content-types/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_list_content_types(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
) -> Result<Vec<ContentTypeInfo>> {
    let url = format!(
        "{}/_api/web/lists/getbytitle('{}')/contenttypes",
        site_url.trim_end_matches('/'),
        encode_segment(list_title)
    );
    let bytes = client
        .run_ladder(entry("sp.content_types.get_list"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Makes an existing site (web-level) content type available on a list.
///
/// Ported from PnPjs `ContentTypes.addAvailableContentType`
/// (`packages/sp/content-types/types.ts`) @ pnpjs `8ee2375d`.
pub async fn add_content_type_to_list(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    content_type_id: &str,
) -> Result<()> {
    let url = format!(
        "{}/_api/web/lists/getbytitle('{}')/contenttypes/addAvailableContentType",
        site_url.trim_end_matches('/'),
        encode_segment(list_title)
    );
    let mut body = Map::new();
    body.insert(
        "contentTypeId".to_string(),
        Value::String(content_type_id.to_string()),
    );
    let payload = serde_json::to_vec(&Value::Object(body))?;
    client
        .run_ladder(
            entry("sp.content_types.add_to_list"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(())
}
