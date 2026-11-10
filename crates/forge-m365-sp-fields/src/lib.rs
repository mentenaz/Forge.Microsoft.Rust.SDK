pub mod creation;
pub mod field;
pub mod types;

use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};

/// Doubles embedded `'` and percent-encodes for use inside a
/// `getbytitle('...')` or `getbyinternalnameor-title('...')` path segment.
/// Ported from PnPjs `encodePath` (`packages/sp/utils/encode-path-str.ts`).
pub fn encode_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn fields_url(site_url: &str) -> String {
    format!("{}/_api/web/fields", site_url.trim_end_matches('/'))
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-fields linked in?")
        })
}

#[pnp_operation(id = "sp.fields.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_all_fields_op() {}

#[pnp_operation(id = "sp.fields.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_field_by_id_op() {}

#[pnp_operation(id = "sp.fields.get_by_title", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_field_by_title_op() {}

#[pnp_operation(id = "sp.fields.get_by_internal_name_or_title", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_field_by_internal_name_or_title_op() {}

#[pnp_operation(id = "sp.fields.add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_field_op() {}

#[pnp_operation(id = "sp.fields.create_field_as_xml", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_field_as_xml_op() {}

#[pnp_operation(id = "sp.fields.add_lookup", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_lookup_op() {}

#[pnp_operation(id = "sp.fields.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_field_op() {}

#[pnp_operation(id = "sp.fields.update", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn update_field_op() {}

#[pnp_operation(id = "sp.fields.set_show_in_display_form", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn set_show_in_display_form_op() {}

#[pnp_operation(id = "sp.fields.set_show_in_edit_form", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn set_show_in_edit_form_op() {}

#[pnp_operation(id = "sp.fields.set_show_in_new_form", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn set_show_in_new_form_op() {}

/// Gets all web-level fields.
///
/// Ported from PnPjs `_Fields` collection read
/// (`packages/sp/fields/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_all_fields(client: &Client<'_>, site_url: &str) -> Result<Vec<types::FieldInfo>> {
    let url = fields_url(site_url);
    let bytes = client
        .run_ladder(entry("sp.fields.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

// Re-exports for convenience
pub use creation::*;
pub use field::*;
pub use types::*;
