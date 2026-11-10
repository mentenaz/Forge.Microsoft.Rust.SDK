use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde_json::{json, Map, Value};

/// The `WebPartsPersonalizationScope` a web part manager operates under.
///
/// Ported from PnPjs `WebPartsPersonalizationScope`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WebPartsPersonalizationScope {
    User = 0,
    Shared = 1,
}

/// PnPjs's own `_WebPartDefinitions`/`_WebPartDefinition` classes are
/// untyped (`_SPCollection`/`_SPInstance` with no generic type param) — the
/// `SP.WebPartDefinition`/`SP.WebPart` entity shape isn't modeled anywhere in
/// the reference source, so rather than invent a struct with unverified
/// field names, this crate returns raw JSON, same as `sp-lists`'s
/// `get_items` and `sp-appcatalog`'s `AppInfo` do for the same reason.
pub type WebPartDefinition = Map<String, Value>;

/// This crate covers only the classic (pre-modern-page)
/// `getLimitedWebPartManager` surface on a `.aspx` file — the client-side
/// (modern) page canvas/web-part model lives in PnPjs `clientside-pages`,
/// which is deliberately out of scope here (too large/complex to port this
/// pass; see the workspace README/AGENTS.md for status). This function
/// builds the base URL for that classic manager, scoped to one file.
fn web_part_manager_url(
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
) -> String {
    let encoded = utf8_percent_encode(
        &server_relative_file_path.replace('\'', "''"),
        NON_ALPHANUMERIC,
    )
    .to_string();
    format!(
        "{}/_api/web/getFileByServerRelativePath(decodedUrl='{encoded}')/getLimitedWebPartManager(scope={})",
        site_url.trim_end_matches('/'),
        scope as u8
    )
}

fn web_part_by_id_url(
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    definition_id: &str,
) -> String {
    format!(
        "{}/webparts/getbyid('{definition_id}')",
        web_part_manager_url(site_url, server_relative_file_path, scope)
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-webparts linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.webparts.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_all_op() {}

#[pnp_operation(id = "sp.webparts.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_id_op() {}

#[pnp_operation(id = "sp.webparts.get_by_control_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_control_id_op() {}

#[pnp_operation(id = "sp.webparts.export", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn export_op() {}

#[pnp_operation(id = "sp.webparts.import", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn import_op() {}

#[pnp_operation(id = "sp.webparts.save_changes", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn save_changes_op() {}

#[pnp_operation(id = "sp.webparts.move_to", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn move_to_op() {}

#[pnp_operation(id = "sp.webparts.close", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn close_op() {}

#[pnp_operation(id = "sp.webparts.open", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn open_op() {}

#[pnp_operation(id = "sp.webparts.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_op() {}

/// Gets the set of web part definitions contained by a page's web part
/// manager.
///
/// Ported from PnPjs `_WebPartDefinitions` collection reads
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_web_parts(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
) -> Result<Vec<WebPartDefinition>> {
    let url = format!(
        "{}/webparts",
        web_part_manager_url(site_url, server_relative_file_path, scope)
    );
    let bytes = client
        .run_ladder(entry("sp.webparts.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a web part definition from the collection by its storage id.
///
/// Ported from PnPjs `_WebPartDefinitions.getById`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_web_part_by_id(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    definition_id: &str,
) -> Result<WebPartDefinition> {
    let url = web_part_by_id_url(site_url, server_relative_file_path, scope, definition_id);
    let bytes = client
        .run_ladder(entry("sp.webparts.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a web part definition from the collection by the `WebPart.ID`
/// control id (distinct from the storage id used by `get_web_part_by_id`).
///
/// Ported from PnPjs `_WebPartDefinitions.getByControlId`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_web_part_by_control_id(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    control_id: &str,
) -> Result<WebPartDefinition> {
    let url = format!(
        "{}/webparts/getByControlId('{control_id}')",
        web_part_manager_url(site_url, server_relative_file_path, scope)
    );
    let bytes = client
        .run_ladder(
            entry("sp.webparts.get_by_control_id"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Exports a web part definition as XML, given its storage id.
///
/// Ported from PnPjs `_LimitedWebPartManager.export`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn export_web_part(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    definition_id: &str,
) -> Result<String> {
    let url = format!(
        "{}/ExportWebPart",
        web_part_manager_url(site_url, server_relative_file_path, scope)
    );
    let body = json!({ "webPartId": definition_id });
    let bytes = client
        .run_ladder(
            entry("sp.webparts.export"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Imports a web part from its `.dwp`/`.webpart` XML definition onto the
/// page. Returns the raw import result, whose shape PnPjs leaves untyped
/// (`Promise<any>`).
///
/// Ported from PnPjs `_LimitedWebPartManager.import`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn import_web_part(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    xml: &str,
) -> Result<Value> {
    let url = format!(
        "{}/ImportWebPart",
        web_part_manager_url(site_url, server_relative_file_path, scope)
    );
    let body = json!({ "webPartXml": xml });
    let bytes = client
        .run_ladder(
            entry("sp.webparts.import"),
            "POST",
            &url,
            &[],
            Some(&serde_json::to_vec(&body)?),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Saves changes made to a web part definition's properties/methods back to
/// the page. Bodiless POST.
///
/// Ported from PnPjs `_WebPartDefinition.saveChanges`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn save_web_part_changes(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    definition_id: &str,
) -> Result<Value> {
    let url = format!(
        "{}/SaveWebPartChanges",
        web_part_by_id_url(site_url, server_relative_file_path, scope, definition_id)
    );
    let bytes = client
        .run_ladder(entry("sp.webparts.save_changes"), "POST", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Moves a web part to a different zone/position on the page. Bodiless POST.
///
/// Ported from PnPjs `_WebPartDefinition.moveTo`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn move_web_part_to(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    definition_id: &str,
    zone_id: &str,
    zone_index: i64,
) -> Result<()> {
    let url = format!(
        "{}/MoveWebPartTo(zoneID='{zone_id}', zoneIndex={zone_index})",
        web_part_by_id_url(site_url, server_relative_file_path, scope, definition_id)
    );
    client
        .run_ladder(entry("sp.webparts.move_to"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Closes the web part; a no-op if it's already closed. Bodiless POST.
///
/// Ported from PnPjs `_WebPartDefinition.close`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn close_web_part(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    definition_id: &str,
) -> Result<()> {
    let url = format!(
        "{}/CloseWebPart",
        web_part_by_id_url(site_url, server_relative_file_path, scope, definition_id)
    );
    client
        .run_ladder(entry("sp.webparts.close"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Opens the web part; a no-op if it's already open. Bodiless POST.
///
/// Ported from PnPjs `_WebPartDefinition.open`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn open_web_part(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    definition_id: &str,
) -> Result<()> {
    let url = format!(
        "{}/OpenWebPart",
        web_part_by_id_url(site_url, server_relative_file_path, scope, definition_id)
    );
    client
        .run_ladder(entry("sp.webparts.open"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Removes a web part from the page. All its settings are lost. Bodiless
/// POST.
///
/// Ported from PnPjs `_WebPartDefinition.delete`
/// (`packages/sp/webparts/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn delete_web_part(
    client: &Client<'_>,
    site_url: &str,
    server_relative_file_path: &str,
    scope: WebPartsPersonalizationScope,
    definition_id: &str,
) -> Result<()> {
    let url = format!(
        "{}/DeleteWebPart",
        web_part_by_id_url(site_url, server_relative_file_path, scope, definition_id)
    );
    client
        .run_ladder(entry("sp.webparts.delete"), "POST", &url, &[], None)
        .await?;
    Ok(())
}
