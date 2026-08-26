use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;

/// Subset of SharePoint's `SP.User` fields, hand-ported from PnPjs's
/// `ISiteUserInfo` (`packages/sp/site-users/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct SiteUserInfo {
    #[serde(rename = "Id")]
    pub id: i64,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "LoginName")]
    pub login_name: String,
    #[serde(rename = "Email")]
    pub email: String,
    #[serde(rename = "PrincipalType")]
    pub principal_type: i64,
    #[serde(rename = "IsSiteAdmin")]
    pub is_site_admin: bool,
    #[serde(rename = "IsHiddenInUI")]
    pub is_hidden_in_ui: bool,
}

/// Doubles embedded `'` and percent-encodes for use inside a `getByX('...')`
/// path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-site-users linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.site_users.get_current_user", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_current_user_op() {}

#[pnp_operation(id = "sp.site_users.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_all_op() {}

#[pnp_operation(id = "sp.site_users.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_id_op() {}

#[pnp_operation(id = "sp.site_users.get_by_email", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_email_op() {}

#[pnp_operation(id = "sp.site_users.get_by_login_name", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_login_name_op() {}

/// Gets the caller's own user record. Requires a delegated (signed-in user)
/// token — with an app-only token there is no "current user", so SharePoint
/// either 400s or returns the app identity depending on tenant configuration.
///
/// Ported from PnPjs `Web.currentUser` (`packages/sp/site-users/web.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_current_user(client: &Client<'_>, site_url: &str) -> Result<SiteUserInfo> {
    let url = format!("{}/_api/web/currentuser", site_url.trim_end_matches('/'));
    let bytes = client
        .run_ladder(
            entry("sp.site_users.get_current_user"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets every user SharePoint has a record for on this site (this includes
/// anyone who has ever been granted access, not just current members).
///
/// Ported from PnPjs `Web.siteUsers` (`packages/sp/site-users/web.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_site_users(client: &Client<'_>, site_url: &str) -> Result<Vec<SiteUserInfo>> {
    let url = format!("{}/_api/web/siteusers", site_url.trim_end_matches('/'));
    let bytes = client
        .run_ladder(entry("sp.site_users.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a user by their SharePoint user id.
///
/// Ported from PnPjs `SiteUsers.getById` (`packages/sp/site-users/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_user_by_id(client: &Client<'_>, site_url: &str, id: i64) -> Result<SiteUserInfo> {
    let url = format!(
        "{}/_api/web/siteusers/getbyid({id})",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.site_users.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a user by email address.
///
/// Ported from PnPjs `SiteUsers.getByEmail` (`packages/sp/site-users/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_user_by_email(
    client: &Client<'_>,
    site_url: &str,
    email: &str,
) -> Result<SiteUserInfo> {
    let url = format!(
        "{}/_api/web/siteusers/getbyemail('{}')",
        site_url.trim_end_matches('/'),
        encode_segment(email)
    );
    let bytes = client
        .run_ladder(entry("sp.site_users.get_by_email"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a user by claims login name (e.g.
/// `i:0#.f|membership|user@contoso.com`).
///
/// The `!@v::` prefix is not a typo or a bug: it's SharePoint REST's literal-
/// value indexer syntax, which PnPjs uses specifically here because claims
/// login names contain `#`/`|`/`@` that survive better unencoded through this
/// path than through normal percent-encoding. Ported byte-for-byte from
/// PnPjs `SiteUsers.getByLoginName` (`packages/sp/site-users/types.ts`)
/// @ pnpjs `8ee2375d` — the login name is intentionally NOT percent-encoded,
/// matching the reference implementation.
pub async fn get_user_by_login_name(
    client: &Client<'_>,
    site_url: &str,
    login_name: &str,
) -> Result<SiteUserInfo> {
    let url = format!(
        "{}/_api/web/siteusers('!@v::{login_name}')",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(
            entry("sp.site_users.get_by_login_name"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
