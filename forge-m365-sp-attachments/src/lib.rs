use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;

/// Subset of SharePoint's `SP.Attachment` fields, hand-ported from PnPjs's
/// `IAttachmentInfo` (`packages/sp/attachments/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct AttachmentInfo {
    #[serde(rename = "FileName")]
    pub file_name: String,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
}

/// Doubles embedded `'` and percent-encodes for use inside a file-name path
/// segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn attachments_url(site_url: &str, list_title: &str, item_id: i64) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')/items({item_id})/AttachmentFiles",
        site_url.trim_end_matches('/'),
        encode_segment(list_title)
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-attachments linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.attachments.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_attachments_op() {}

#[pnp_operation(id = "sp.attachments.download", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn download_attachment_op() {}

#[pnp_operation(id = "sp.attachments.add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_attachment_op() {}

#[pnp_operation(id = "sp.attachments.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_attachment_op() {}

/// Gets the attachments on a list item.
///
/// Ported from PnPjs `Item.attachmentFiles` via `_Attachments`
/// (`packages/sp/attachments/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_attachments(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Vec<AttachmentInfo>> {
    let url = attachments_url(site_url, list_title, item_id);
    let bytes = client
        .run_ladder(entry("sp.attachments.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Downloads an attachment's raw content.
///
/// Ported from PnPjs `ReadableFile.getBuffer` as applied to `_Attachment`
/// (`packages/sp/attachments/types.ts`) @ pnpjs `8ee2375d`.
pub async fn download_attachment(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    file_name: &str,
) -> Result<Vec<u8>> {
    let url = format!(
        "{}('{}')/$value",
        attachments_url(site_url, list_title, item_id),
        encode_segment(file_name)
    );
    client
        .run_ladder(entry("sp.attachments.download"), "GET", &url, &[], None)
        .await
}

/// Adds an attachment to a list item. Raw-bytes POST, same as `sp-files`
/// upload — SharePoint takes the file content directly as the request body,
/// not JSON.
///
/// Ported from PnPjs `Attachments.add` (`packages/sp/attachments/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_attachment(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    file_name: &str,
    content: &[u8],
) -> Result<AttachmentInfo> {
    let url = format!(
        "{}/add(FileName='{}')",
        attachments_url(site_url, list_title, item_id),
        encode_segment(file_name)
    );
    let bytes = client
        .run_ladder(
            entry("sp.attachments.add"),
            "POST",
            &url,
            &[("Content-Type", "application/octet-stream")],
            Some(content),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes an attachment.
///
/// Ported from PnPjs `Attachment.delete` (`packages/sp/attachments/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn delete_attachment(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    file_name: &str,
) -> Result<()> {
    let url = format!(
        "{}('{}')",
        attachments_url(site_url, list_title, item_id),
        encode_segment(file_name)
    );
    client
        .run_ladder(
            entry("sp.attachments.delete"),
            "DELETE",
            &url,
            &[("IF-MATCH", "*")],
            None,
        )
        .await?;
    Ok(())
}
