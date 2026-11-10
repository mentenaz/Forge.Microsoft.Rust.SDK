use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;

/// Subset of SharePoint's `SP.Form` fields, hand-ported from PnPjs's
/// `IFormInfo` (`packages/sp/forms/types.ts`). `form_type` is the numeric
/// `PageType` enum value (PnPjs `packages/sp/types.ts`); not re-modeled as
/// a Rust enum here since this crate doesn't otherwise need `PageType`.
#[derive(Debug, Clone, Deserialize)]
pub struct FormInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "FormType")]
    pub form_type: i64,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
}

/// Doubles embedded `'` and percent-encodes for use inside a `getbytitle('...')`
/// path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn forms_url(site_url: &str, list_title: &str) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')/forms",
        site_url.trim_end_matches('/'),
        encode_segment(list_title)
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-forms linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.forms.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_forms_op() {}

#[pnp_operation(id = "sp.forms.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_form_by_id_op() {}

/// Gets the forms (display/edit/new) associated with a list.
///
/// Ported from PnPjs `List.forms` via `_Forms` (`packages/sp/forms/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_forms(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
) -> Result<Vec<FormInfo>> {
    let url = forms_url(site_url, list_title);
    let bytes = client
        .run_ladder(entry("sp.forms.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a single form by its GUID id.
///
/// Ported from PnPjs `Forms.getById` (`packages/sp/forms/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_form_by_id(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    form_id: &str,
) -> Result<FormInfo> {
    let url = format!("{}('{form_id}')", forms_url(site_url, list_title));
    let bytes = client
        .run_ladder(entry("sp.forms.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
