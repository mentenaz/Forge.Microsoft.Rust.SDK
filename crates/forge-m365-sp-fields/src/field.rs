use forge_m365_core::{registered_operations, Client, OperationEntry, Result};
use serde_json::{Map, Value};

use crate::types::FieldInfo;

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-fields linked in?")
        })
}

/// Deletes a field by its GUID id.
///
/// Ported from PnPjs `_Field.delete` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn delete_field(client: &Client<'_>, site_url: &str, field_id: &str) -> Result<()> {
    let url = format!(
        "{}/_api/web/fields('{}')",
        site_url.trim_end_matches('/'),
        field_id
    );
    client
        .run_ladder(
            entry("sp.fields.delete"),
            "DELETE",
            &url,
            &[("IF-MATCH", "*")],
            None,
        )
        .await?;
    Ok(())
}

/// Updates a field's properties.
///
/// `properties` is the map of columns to change (MERGE semantics).
///
/// Ported from PnPjs `_Field.update` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn update_field(
    client: &Client<'_>,
    site_url: &str,
    field_id: &str,
    properties: Map<String, Value>,
) -> Result<()> {
    let url = format!(
        "{}/_api/web/fields('{}')",
        site_url.trim_end_matches('/'),
        field_id
    );
    let payload = serde_json::to_vec(&Value::Object(properties))?;
    client
        .run_ladder(
            entry("sp.fields.update"),
            "PATCH",
            &url,
            &[("IF-MATCH", "*")],
            Some(&payload),
        )
        .await?;
    Ok(())
}

/// Sets whether a field appears in the display form.
///
/// Ported from PnPjs `_Field.setShowInDisplayForm`
/// (`packages/sp/fields/types.ts`) @ pnpjs `8ee2375d`.
pub async fn set_show_in_display_form(
    client: &Client<'_>,
    site_url: &str,
    field_id: &str,
    show: bool,
) -> Result<()> {
    let url = format!(
        "{}/_api/web/fields('{}')/setshowindisplayform({})",
        site_url.trim_end_matches('/'),
        field_id,
        show
    );
    client
        .run_ladder(
            entry("sp.fields.set_show_in_display_form"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

/// Sets whether a field appears in the edit form.
///
/// Ported from PnPjs `_Field.setShowInEditForm`
/// (`packages/sp/fields/types.ts`) @ pnpjs `8ee2375d`.
pub async fn set_show_in_edit_form(
    client: &Client<'_>,
    site_url: &str,
    field_id: &str,
    show: bool,
) -> Result<()> {
    let url = format!(
        "{}/_api/web/fields('{}')/setshowineditform({})",
        site_url.trim_end_matches('/'),
        field_id,
        show
    );
    client
        .run_ladder(
            entry("sp.fields.set_show_in_edit_form"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

/// Sets whether a field appears in the new form.
///
/// Ported from PnPjs `_Field.setShowInNewForm`
/// (`packages/sp/fields/types.ts`) @ pnpjs `8ee2375d`.
pub async fn set_show_in_new_form(
    client: &Client<'_>,
    site_url: &str,
    field_id: &str,
    show: bool,
) -> Result<()> {
    let url = format!(
        "{}/_api/web/fields('{}')/setshowinnewform({})",
        site_url.trim_end_matches('/'),
        field_id,
        show
    );
    client
        .run_ladder(
            entry("sp.fields.set_show_in_new_form"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

/// Gets a single field's metadata by its GUID id.
///
/// Ported from PnPjs `_Fields.getById` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_field_by_id(
    client: &Client<'_>,
    site_url: &str,
    field_id: &str,
) -> Result<FieldInfo> {
    let url = format!(
        "{}/_api/web/fields('{}')",
        site_url.trim_end_matches('/'),
        field_id
    );
    let bytes = client
        .run_ladder(entry("sp.fields.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a single field by its case-sensitive title.
///
/// Ported from PnPjs `_Fields.getByTitle` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_field_by_title(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
) -> Result<FieldInfo> {
    let url = format!(
        "{}/_api/web/fields/getbytitle('{}')",
        site_url.trim_end_matches('/'),
        crate::encode_segment(title)
    );
    let bytes = client
        .run_ladder(entry("sp.fields.get_by_title"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a single field by internal name or title.
///
/// Ported from PnPjs `_Fields.getByInternalNameOrTitle`
/// (`packages/sp/fields/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_field_by_internal_name_or_title(
    client: &Client<'_>,
    site_url: &str,
    name: &str,
) -> Result<FieldInfo> {
    let url = format!(
        "{}/_api/web/fields/getbyinternalnameor-title('{}')",
        site_url.trim_end_matches('/'),
        crate::encode_segment(name)
    );
    let bytes = client
        .run_ladder(
            entry("sp.fields.get_by_internal_name_or_title"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
