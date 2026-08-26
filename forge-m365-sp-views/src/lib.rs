use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;
use serde_json::Value;

/// Subset of SharePoint's `SP.View` fields, hand-ported from PnPjs's
/// `IViewInfo` (`packages/sp/views/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct ViewInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "Hidden")]
    pub hidden: bool,
    #[serde(rename = "DefaultView")]
    pub default_view: bool,
    #[serde(rename = "PersonalView")]
    pub personal_view: bool,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
    #[serde(rename = "ViewQuery")]
    pub view_query: String,
    #[serde(rename = "RowLimit")]
    pub row_limit: i64,
}

/// The list of internal field names shown in a view, hand-ported from
/// PnPjs's `_ViewFields` return shape (`packages/sp/views/types.ts`).
#[derive(Debug, Clone, Default)]
pub struct ViewFields {
    pub items: Vec<String>,
    pub schema_xml: String,
}

/// Doubles embedded `'` and percent-encodes for use inside a `getbytitle('...')`
/// path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn views_url(site_url: &str, list_title: &str) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')/views",
        site_url.trim_end_matches('/'),
        encode_segment(list_title)
    )
}

/// Some collection-typed properties come back wrapped as `{"results": [...]}`,
/// others bare — same defensive check used in `sp-sites`/`sp-search`.
fn unwrap_array(value: &Value) -> Vec<Value> {
    value
        .get("results")
        .unwrap_or(value)
        .as_array()
        .cloned()
        .unwrap_or_default()
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-views linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.views.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_views_op() {}

#[pnp_operation(id = "sp.views.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_view_by_id_op() {}

#[pnp_operation(id = "sp.views.get_by_title", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_view_by_title_op() {}

#[pnp_operation(id = "sp.views.get_fields", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_view_fields_op() {}

/// Gets every view defined on a list.
///
/// Ported from PnPjs `List.views` via `_Views`
/// (`packages/sp/views/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_views(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
) -> Result<Vec<ViewInfo>> {
    let url = views_url(site_url, list_title);
    let bytes = client
        .run_ladder(entry("sp.views.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a view by its GUID id.
///
/// Ported from PnPjs `Views.getById` (`packages/sp/views/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_view_by_id(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    view_id: &str,
) -> Result<ViewInfo> {
    let url = format!("{}('{}')", views_url(site_url, list_title), view_id);
    let bytes = client
        .run_ladder(entry("sp.views.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a view by its (case-sensitive) title.
///
/// Ported from PnPjs `Views.getByTitle` (`packages/sp/views/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_view_by_title(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    view_title: &str,
) -> Result<ViewInfo> {
    let url = format!(
        "{}/getbytitle('{}')",
        views_url(site_url, list_title),
        encode_segment(view_title)
    );
    let bytes = client
        .run_ladder(entry("sp.views.get_by_title"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the internal field names shown in a view, and the view's schema XML.
///
/// Ported from PnPjs `View.fields` via `_ViewFields`
/// (`packages/sp/views/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_view_fields(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    view_id: &str,
) -> Result<ViewFields> {
    let url = format!(
        "{}('{}')/viewfields",
        views_url(site_url, list_title),
        view_id
    );
    let bytes = client
        .run_ladder(entry("sp.views.get_fields"), "GET", &url, &[], None)
        .await?;
    let root: Value = serde_json::from_slice(&bytes)?;
    let items = unwrap_array(&root["Items"])
        .iter()
        .filter_map(|v| v.as_str().map(str::to_string))
        .collect();
    Ok(ViewFields {
        items,
        schema_xml: root["SchemaXml"].as_str().unwrap_or_default().to_string(),
    })
}
