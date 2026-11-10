use forge_m365_core::{registered_operations, Client, OperationEntry, Result};
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde_json::{Map, Value};

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

/// Gets the effective base permissions for the item.
///
/// Ported from PnPjs `Item.effectiveBasePermissions`
/// (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_effective_base_permissions(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Map<String, Value>> {
    let url = format!(
        "{}/EffectiveBasePermissions",
        item_url(site_url, list_title, item_id)
    );
    let bytes = client
        .run_ladder(
            entry("sp.items.get_effective_base_permissions"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the effective base permissions for the item in a UI context.
///
/// Ported from PnPjs `Item.effectiveBasePermissionsForUI`
/// (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_effective_base_permissions_for_ui(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Map<String, Value>> {
    let url = format!(
        "{}/EffectiveBasePermissionsForUI",
        item_url(site_url, list_title, item_id)
    );
    let bytes = client
        .run_ladder(
            entry("sp.items.get_effective_base_permissions_for_ui"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the field values for this list item in their HTML representation.
///
/// Ported from PnPjs `Item.fieldValuesAsHTML`
/// (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_field_values_as_html(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Map<String, Value>> {
    let url = format!(
        "{}/FieldValuesAsHTML",
        item_url(site_url, list_title, item_id)
    );
    let bytes = client
        .run_ladder(
            entry("sp.items.get_field_values_as_html"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the field values for this list item in their text representation.
///
/// Ported from PnPjs `Item.fieldValuesAsText`
/// (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_field_values_as_text(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Map<String, Value>> {
    let url = format!(
        "{}/FieldValuesAsText",
        item_url(site_url, list_title, item_id)
    );
    let bytes = client
        .run_ladder(
            entry("sp.items.get_field_values_as_text"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the field values for this list item for use in editing controls.
///
/// Ported from PnPjs `Item.fieldValuesForEdit`
/// (`packages/sp/items/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_field_values_for_edit(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Map<String, Value>> {
    let url = format!(
        "{}/FieldValuesForEdit",
        item_url(site_url, list_title, item_id)
    );
    let bytes = client
        .run_ladder(
            entry("sp.items.get_field_values_for_edit"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
