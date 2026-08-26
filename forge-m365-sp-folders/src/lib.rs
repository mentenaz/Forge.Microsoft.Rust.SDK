use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;

/// Subset of SharePoint's `SP.Folder` fields, hand-ported from PnPjs's
/// `IFolderInfo` (`packages/sp/folders/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct FolderInfo {
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
    #[serde(rename = "ItemCount")]
    pub item_count: i64,
    #[serde(rename = "Exists")]
    pub exists: bool,
    #[serde(rename = "TimeCreated")]
    pub time_created: String,
    #[serde(rename = "TimeLastModified")]
    pub time_last_modified: String,
    #[serde(rename = "UniqueId")]
    pub unique_id: String,
}

/// Doubles embedded `'` and percent-encodes for use inside an OData
/// `decodedUrl='...'`/`DecodedUrl='...'` path segment. Ported from PnPjs
/// `encodePath` (`packages/sp/utils/encode-path-str.ts`).
fn encode_path_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn folder_by_path_url(site_url: &str, server_relative_path: &str) -> String {
    format!(
        "{}/_api/web/getFolderByServerRelativePath(decodedUrl='{}')",
        site_url.trim_end_matches('/'),
        encode_path_segment(server_relative_path)
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-folders linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.folders.get_by_path", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_path_op() {}

#[pnp_operation(id = "sp.folders.get_subfolders", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_subfolders_op() {}

#[pnp_operation(id = "sp.folders.add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_op() {}

#[pnp_operation(id = "sp.folders.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_op() {}

/// Gets a folder's properties by its server-relative path.
///
/// Ported from PnPjs `folderFromServerRelativePath`
/// (`packages/sp/folders/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_folder_by_path(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> Result<FolderInfo> {
    let url = format!(
        "{}?$select=Name,ServerRelativeUrl,ItemCount,Exists,TimeCreated,TimeLastModified,UniqueId",
        folder_by_path_url(site_url, server_relative_path)
    );
    let bytes = client
        .run_ladder(entry("sp.folders.get_by_path"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the immediate subfolders of the folder at `server_relative_path`.
///
/// Ported from PnPjs `Folder.folders` (`packages/sp/folders/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_subfolders(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> Result<Vec<FolderInfo>> {
    let url = format!(
        "{}/folders?$select=Name,ServerRelativeUrl,ItemCount,Exists,TimeCreated,TimeLastModified,UniqueId",
        folder_by_path_url(site_url, server_relative_path)
    );
    let bytes = client
        .run_ladder(entry("sp.folders.get_subfolders"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Creates a folder at `server_relative_url` (and any missing intermediate
/// folders in its path). Bodiless POST — SharePoint takes the path and
/// overwrite flag as URL parameters, not a JSON body.
///
/// Ported from PnPjs `Folders.addUsingPath` (`packages/sp/folders/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_folder(
    client: &Client<'_>,
    site_url: &str,
    server_relative_url: &str,
    overwrite: bool,
) -> Result<FolderInfo> {
    let url = format!(
        "{}/_api/web/folders/addUsingPath(DecodedUrl='{}',overwrite={overwrite})",
        site_url.trim_end_matches('/'),
        encode_path_segment(server_relative_url)
    );
    let bytes = client
        .run_ladder(entry("sp.folders.add"), "POST", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes a folder by its server-relative path.
///
/// Ported from PnPjs `Folder.delete` (`packages/sp/folders/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn delete_folder(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> Result<()> {
    let url = folder_by_path_url(site_url, server_relative_path);
    client
        .run_ladder(
            entry("sp.folders.delete"),
            "DELETE",
            &url,
            &[("IF-MATCH", "*")],
            None,
        )
        .await?;
    Ok(())
}
