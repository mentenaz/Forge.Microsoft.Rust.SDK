use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::{Deserialize, Serialize};
use serde_json::Value;

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-social linked in?")
        })
}

fn following_url(site_url: &str) -> String {
    format!("{}/_api/social.following", site_url.trim_end_matches('/'))
}

/// `SP.Social.SocialActorType` — a single actor kind, per PnPjs
/// `SocialActorType` (`packages/sp/social/types.ts`).
pub mod actor_type {
    pub const USER: i64 = 0;
    pub const DOCUMENT: i64 = 1;
    pub const SITE: i64 = 2;
    pub const TAG: i64 = 3;
}

/// `SP.Social.SocialActorTypes` — a bitwise set of [`actor_type`] values used
/// to filter `my/followed(...)`/`my/followedcount(...)`, per PnPjs
/// `SocialActorTypes` (`packages/sp/social/types.ts`).
pub mod actor_types {
    pub const NONE: u32 = 0;
    pub const USER: u32 = 1 << 0;
    pub const DOCUMENT: u32 = 1 << 1;
    pub const SITE: u32 = 1 << 2;
    pub const TAG: u32 = 1 << 3;
    /// The set excludes documents and sites that do not have feeds.
    pub const EXCLUDE_CONTENT_WITHOUT_FEEDS: u32 = 268_435_456;
    /// The set includes group sites.
    pub const INCLUDE_GROUPS_SITES: u32 = 536_870_912;
    /// The set includes only items created within the last 24 hours.
    pub const WITHIN_LAST_24_HOURS: u32 = 1_073_741_824;
}

/// `SP.Social.SocialFollowResult`, the outcome of [`follow`].
///
/// Ported from PnPjs `SocialFollowResult` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SocialFollowResult {
    Ok,
    AlreadyFollowing,
    LimitReached,
    InternalError,
    /// A code the server returned that isn't one of the four documented
    /// values above.
    Other(i64),
}

impl From<i64> for SocialFollowResult {
    fn from(code: i64) -> Self {
        match code {
            0 => Self::Ok,
            1 => Self::AlreadyFollowing,
            2 => Self::LimitReached,
            3 => Self::InternalError,
            other => Self::Other(other),
        }
    }
}

impl<'de> Deserialize<'de> for SocialFollowResult {
    fn deserialize<D>(deserializer: D) -> std::result::Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        Ok(Self::from(i64::deserialize(deserializer)?))
    }
}

/// Input to [`follow`], [`is_followed`], and [`stop_following`].
///
/// Ported from PnPjs `ISocialActorInfo` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize)]
pub struct SocialActorInfo {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub account_name: Option<String>,
    pub actor_type: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content_uri: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub tag_guid: Option<String>,
}

/// A followed/following/suggested actor as returned by the `my/*` endpoints.
///
/// Ported from PnPjs `ISocialActor` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct SocialActor {
    pub actor_type: i64,
    pub id: String,
    pub uri: String,
    pub name: String,
    pub is_followed: bool,
    pub status: i64,
    pub can_follow: bool,
    pub image_uri: String,
    pub account_name: String,
    pub email_address: String,
    pub title: String,
    pub status_text: String,
    pub personal_site_uri: String,
    pub followed_content_uri: String,
    pub content_uri: String,
    pub library_uri: String,
    pub tag_guid: String,
}

/// Response shape from `my` itself.
///
/// Ported from PnPjs `IMySocialData` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct MySocialData {
    pub social_actor: SocialActor,
    pub my_followed_documents_uri: String,
    pub my_followed_sites_uri: String,
}

fn actor_request_body(actor_info: &SocialActorInfo) -> serde_json::Result<Vec<u8>> {
    #[derive(Serialize)]
    struct Wrapper<'a> {
        actor: &'a SocialActorInfo,
    }
    serde_json::to_vec(&Wrapper { actor: actor_info })
}

/// Unwraps a verbose-OData `{"<Key>": value}` envelope PnPjs peels off with
/// `hOP`, falling back to a bare value for a non-verbose/plain response.
fn unwrap_field(bytes: &[u8], key: &str) -> serde_json::Result<Value> {
    let v: Value = serde_json::from_slice(bytes)?;
    Ok(v.get(key).cloned().unwrap_or(v))
}

#[pnp_operation(id = "sp.social.get_followed_sites_uri", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_followed_sites_uri_op() {}

#[pnp_operation(
    id = "sp.social.get_followed_documents_uri",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_followed_documents_uri_op() {}

#[pnp_operation(id = "sp.social.follow", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn follow_op() {}

#[pnp_operation(id = "sp.social.is_followed", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn is_followed_op() {}

#[pnp_operation(id = "sp.social.stop_following", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn stop_following_op() {}

#[pnp_operation(id = "sp.social.my", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn my_op() {}

#[pnp_operation(id = "sp.social.my_followed", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn my_followed_op() {}

#[pnp_operation(
    id = "sp.social.my_followed_count",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn my_followed_count_op() {}

#[pnp_operation(id = "sp.social.my_followers", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn my_followers_op() {}

#[pnp_operation(id = "sp.social.my_suggestions", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn my_suggestions_op() {}

/// Gets a URI for the current user's followed-sites content.
///
/// Ported from PnPjs `_Social.getFollowedSitesUri`
/// (`packages/sp/social/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_followed_sites_uri(client: &Client<'_>, site_url: &str) -> Result<String> {
    let url = format!("{}/FollowedSitesUri", following_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.social.get_followed_sites_uri"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let v = unwrap_field(&bytes, "FollowedSitesUri")?;
    Ok(v.as_str().unwrap_or_default().to_string())
}

/// Gets a URI for the current user's followed-documents content.
///
/// Ported from PnPjs `_Social.getFollowedDocumentsUri`
/// (`packages/sp/social/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_followed_documents_uri(client: &Client<'_>, site_url: &str) -> Result<String> {
    let url = format!("{}/FollowedDocumentsUri", following_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.social.get_followed_documents_uri"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let v = unwrap_field(&bytes, "FollowedDocumentsUri")?;
    Ok(v.as_str().unwrap_or_default().to_string())
}

/// Follows an actor (user, document, site, or tag) for the current user.
///
/// Ported from PnPjs `_Social.follow` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn follow(
    client: &Client<'_>,
    site_url: &str,
    actor_info: &SocialActorInfo,
) -> Result<SocialFollowResult> {
    let url = format!("{}/follow", following_url(site_url));
    let body = actor_request_body(actor_info)?;
    let bytes = client
        .run_ladder(entry("sp.social.follow"), "POST", &url, &[], Some(&body))
        .await?;
    let v = unwrap_field(&bytes, "follow")?;
    let code = v.as_i64().unwrap_or(0);
    Ok(SocialFollowResult::from(code))
}

/// Checks whether the current user is following an actor.
///
/// Ported from PnPjs `_Social.isFollowed` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn is_followed(
    client: &Client<'_>,
    site_url: &str,
    actor_info: &SocialActorInfo,
) -> Result<bool> {
    let url = format!("{}/isfollowed", following_url(site_url));
    let body = actor_request_body(actor_info)?;
    let bytes = client
        .run_ladder(
            entry("sp.social.is_followed"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    let v = unwrap_field(&bytes, "isfollowed")?;
    Ok(v.as_bool().unwrap_or(false))
}

/// Stops following an actor for the current user. Bodiless-response POST.
///
/// Ported from PnPjs `_Social.stopFollowing` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn stop_following(
    client: &Client<'_>,
    site_url: &str,
    actor_info: &SocialActorInfo,
) -> Result<()> {
    let url = format!("{}/stopfollowing", following_url(site_url));
    let body = actor_request_body(actor_info)?;
    client
        .run_ladder(
            entry("sp.social.stop_following"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(())
}

/// Gets the current user's social summary (actor + followed-sites/documents
/// URIs).
///
/// Ported from PnPjs `_MySocial` invoked as a function
/// (`packages/sp/social/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn my(client: &Client<'_>, site_url: &str) -> Result<MySocialData> {
    let url = format!("{}/my", following_url(site_url));
    let bytes = client
        .run_ladder(entry("sp.social.my"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets users, documents, sites, and tags the current user is following.
/// `types` is a bitwise-OR'd set of [`actor_types`] values.
///
/// Ported from PnPjs `_MySocial.followed` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn my_followed(
    client: &Client<'_>,
    site_url: &str,
    types: u32,
) -> Result<Vec<SocialActor>> {
    let url = format!("{}/my/followed(types={types})", following_url(site_url));
    let bytes = client
        .run_ladder(entry("sp.social.my_followed"), "GET", &url, &[], None)
        .await?;
    let v = unwrap_field(&bytes, "Followed")?;
    let v = v.get("results").cloned().unwrap_or(v);
    Ok(serde_json::from_value(v)?)
}

/// Gets the count of users, documents, sites, and tags the current user is
/// following. `types` is a bitwise-OR'd set of [`actor_types`] values.
///
/// Ported from PnPjs `_MySocial.followedCount`
/// (`packages/sp/social/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn my_followed_count(client: &Client<'_>, site_url: &str, types: u32) -> Result<i64> {
    let url = format!(
        "{}/my/followedcount(types={types})",
        following_url(site_url)
    );
    let bytes = client
        .run_ladder(entry("sp.social.my_followed_count"), "GET", &url, &[], None)
        .await?;
    let v = unwrap_field(&bytes, "FollowedCount")?;
    Ok(v.as_i64().unwrap_or_default())
}

/// Gets the users who are following the current user.
///
/// Ported from PnPjs `_MySocial.followers` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn my_followers(client: &Client<'_>, site_url: &str) -> Result<Vec<SocialActor>> {
    let url = format!("{}/my/followers", following_url(site_url));
    let bytes = client
        .run_ladder(entry("sp.social.my_followers"), "GET", &url, &[], None)
        .await?;
    let v = unwrap_field(&bytes, "Followers")?;
    let v = v.get("results").cloned().unwrap_or(v);
    Ok(serde_json::from_value(v)?)
}

/// Gets users who the current user might want to follow.
///
/// Ported from PnPjs `_MySocial.suggestions` (`packages/sp/social/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn my_suggestions(client: &Client<'_>, site_url: &str) -> Result<Vec<SocialActor>> {
    let url = format!("{}/my/suggestions", following_url(site_url));
    let bytes = client
        .run_ladder(entry("sp.social.my_suggestions"), "GET", &url, &[], None)
        .await?;
    let v = unwrap_field(&bytes, "Suggestions")?;
    let v = v.get("results").cloned().unwrap_or(v);
    Ok(serde_json::from_value(v)?)
}
