use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;

/// Subset of SharePoint's `SP.File` fields, hand-ported from PnPjs's `IFileInfo`
/// (`packages/sp/files/types.ts`). `length` is a string because SharePoint
/// REST always emits `Edm.Int64` fields as JSON strings, in both `verbose`
/// and `nometadata` response formats.
#[derive(Debug, Clone, Deserialize)]
pub struct FileInfo {
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
    #[serde(rename = "Length")]
    pub length: String,
    #[serde(rename = "TimeLastModified")]
    pub time_last_modified: String,
    #[serde(rename = "UniqueId")]
    pub unique_id: String,
}

/// Doubles embedded `'` and percent-encodes for use inside an OData
/// `decodedUrl='...'` path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_path_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn file_by_path_url(site_url: &str, server_relative_path: &str) -> String {
    format!(
        "{}/_api/web/getFileByServerRelativePath(decodedUrl='{}')",
        site_url.trim_end_matches('/'),
        encode_path_segment(server_relative_path)
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-files linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.files.get_by_path", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_path_op() {}

#[pnp_operation(id = "sp.files.download", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn download_op() {}

#[pnp_operation(id = "sp.files.upload", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn upload_op() {}

#[pnp_operation(id = "sp.files.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_op() {}

/// Gets a file's properties by its server-relative path
/// (e.g. `/sites/team/Shared Documents/report.docx`).
///
/// Ported from PnPjs `Web.getFileByServerRelativePath` (`packages/sp/files/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_file_by_path(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> Result<FileInfo> {
    let url = format!(
        "{}?$select=Name,ServerRelativeUrl,Length,TimeLastModified,UniqueId",
        file_by_path_url(site_url, server_relative_path)
    );
    let bytes = client
        .run_ladder(entry("sp.files.get_by_path"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Downloads a file's raw content.
///
/// Ported from PnPjs `ReadableFile.getBuffer` (`packages/sp/files/readable-file.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn download_file(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> Result<Vec<u8>> {
    let url = format!(
        "{}/$value",
        file_by_path_url(site_url, server_relative_path)
    );
    client
        .run_ladder(entry("sp.files.download"), "GET", &url, &[], None)
        .await
}

/// Uploads `content` as a file named `file_name` into the folder at
/// `folder_server_relative_path`. Single-request upload only — SharePoint's
/// chunked-upload session (`Files.addChunked`/`startUpload`/`continueUpload`)
/// for large files is not yet ported.
///
/// Ported from PnPjs `Files.addUsingPath` (`packages/sp/files/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn upload_file(
    client: &Client<'_>,
    site_url: &str,
    folder_server_relative_path: &str,
    file_name: &str,
    content: &[u8],
    overwrite: bool,
) -> Result<FileInfo> {
    let url = format!(
        "{}/_api/web/getFolderByServerRelativePath(decodedUrl='{}')/files/AddUsingPath(decodedurl='{}',Overwrite={overwrite})",
        site_url.trim_end_matches('/'),
        encode_path_segment(folder_server_relative_path),
        encode_path_segment(file_name)
    );
    let bytes = client
        .run_ladder(
            entry("sp.files.upload"),
            "POST",
            &url,
            &[("Content-Type", "application/octet-stream")],
            Some(content),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes a file by its server-relative path.
///
/// Ported from PnPjs `File.delete` (`packages/sp/files/types.ts`) @ pnpjs `8ee2375d`.
pub async fn delete_file(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> Result<()> {
    let url = file_by_path_url(site_url, server_relative_path);
    client
        .run_ladder(
            entry("sp.files.delete"),
            "DELETE",
            &url,
            &[("IF-MATCH", "*")],
            None,
        )
        .await?;
    Ok(())
}
