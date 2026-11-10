use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;
use serde_json::{Map, Value};

/// A comment's author, hand-ported from PnPjs's `ICommentAuthorData`
/// (`packages/sp/comments/types.ts`). Unlike every other crate so far, this
/// API's JSON fields are camelCase, not PascalCase — it's SharePoint's newer
/// "likes and comments" service, not the classic `_api/web`/`SP.List` REST
/// surface, even though it's addressed under the same `_api/web/lists/...`
/// item path.
#[derive(Debug, Clone, Deserialize)]
pub struct CommentAuthorInfo {
    pub email: String,
    pub id: i64,
    #[serde(rename = "isActive")]
    pub is_active: bool,
    #[serde(rename = "loginName")]
    pub login_name: String,
    pub name: String,
}

/// A list-item comment, hand-ported from PnPjs's `ICommentInfo`
/// (`packages/sp/comments/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct CommentInfo {
    pub id: String,
    pub text: String,
    pub author: CommentAuthorInfo,
    #[serde(rename = "createdDate")]
    pub created_date: String,
    #[serde(rename = "isLikedByUser")]
    pub is_liked_by_user: bool,
    #[serde(rename = "likeCount")]
    pub like_count: i64,
    #[serde(rename = "replyCount")]
    pub reply_count: i64,
    #[serde(rename = "isReply")]
    pub is_reply: bool,
    #[serde(rename = "parentId")]
    pub parent_id: String,
}

/// Doubles embedded `'` and percent-encodes for use inside a `getbytitle('...')`
/// path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn comments_url(site_url: &str, list_title: &str, item_id: i64) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')/items({item_id})/comments",
        site_url.trim_end_matches('/'),
        encode_segment(list_title)
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-comments linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.comments.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_comments_op() {}

#[pnp_operation(id = "sp.comments.add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_comment_op() {}

#[pnp_operation(id = "sp.comments.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_comment_op() {}

#[pnp_operation(id = "sp.comments.like", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn like_comment_op() {}

#[pnp_operation(id = "sp.comments.unlike", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn unlike_comment_op() {}

/// Gets the comments on a list item.
///
/// Ported from PnPjs `Item.comments` via `_Comments`
/// (`packages/sp/comments/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_comments(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<Vec<CommentInfo>> {
    let url = comments_url(site_url, list_title, item_id);
    let bytes = client
        .run_ladder(entry("sp.comments.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Adds a comment to a list item.
///
/// Ported from PnPjs `Comments.add` (`packages/sp/comments/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_comment(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    text: &str,
) -> Result<CommentInfo> {
    let url = comments_url(site_url, list_title, item_id);
    let mut body = Map::new();
    body.insert("text".to_string(), Value::String(text.to_string()));
    let payload = serde_json::to_vec(&Value::Object(body))?;
    let bytes = client
        .run_ladder(entry("sp.comments.add"), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deletes a comment.
///
/// Ported from PnPjs `Comment.delete` (`packages/sp/comments/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn delete_comment(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    comment_id: &str,
) -> Result<()> {
    let url = format!(
        "{}({comment_id})",
        comments_url(site_url, list_title, item_id)
    );
    client
        .run_ladder(entry("sp.comments.delete"), "DELETE", &url, &[], None)
        .await?;
    Ok(())
}

/// Likes a comment as the current (delegated) user. Bodiless POST.
///
/// Ported from PnPjs `Comment.like` (`packages/sp/comments/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn like_comment(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    comment_id: &str,
) -> Result<()> {
    let url = format!(
        "{}({comment_id})/Like",
        comments_url(site_url, list_title, item_id)
    );
    client
        .run_ladder(entry("sp.comments.like"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Unlikes a comment as the current (delegated) user. Bodiless POST.
///
/// Ported from PnPjs `Comment.unlike` (`packages/sp/comments/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn unlike_comment(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    comment_id: &str,
) -> Result<()> {
    let url = format!(
        "{}({comment_id})/Unlike",
        comments_url(site_url, list_title, item_id)
    );
    client
        .run_ladder(entry("sp.comments.unlike"), "POST", &url, &[], None)
        .await?;
    Ok(())
}
