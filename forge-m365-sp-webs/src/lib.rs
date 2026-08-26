use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::Deserialize;

/// Subset of SharePoint's `SP.Web` fields, hand-ported from PnPjs's `IWebInfo`
/// (`packages/sp/webs/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct WebInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "Url")]
    pub url: String,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "WebTemplate")]
    pub web_template: String,
}

#[derive(Debug, Deserialize)]
struct ParentWebQuery {
    #[serde(rename = "Url")]
    url: String,
    #[serde(rename = "ParentWeb")]
    parent_web: Option<ParentWebRef>,
}

#[derive(Debug, Deserialize)]
struct ParentWebRef {
    #[serde(rename = "ServerRelativeUrl")]
    server_relative_url: String,
}

/// `scheme://host` from an absolute url, e.g. `https://contoso.sharepoint.com`
/// from `https://contoso.sharepoint.com/sites/team`.
fn origin(url: &str) -> &str {
    if let Some(scheme_end) = url.find("://") {
        let after_scheme = scheme_end + 3;
        if let Some(slash) = url[after_scheme..].find('/') {
            return &url[..after_scheme + slash];
        }
    }
    url.trim_end_matches('/')
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-webs linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.webs.get", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_web_op() {}

#[pnp_operation(id = "sp.webs.get_subwebs", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_subwebs_op() {}

#[pnp_operation(id = "sp.webs.get_parent_web_url", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_parent_web_url_op() {}

/// Gets the web's own properties.
///
/// Ported from PnPjs `Web.get` (`packages/sp/webs/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_web(client: &Client<'_>, site_url: &str) -> Result<WebInfo> {
    let url = format!(
        "{}/_api/web?$select=Id,Title,Url,ServerRelativeUrl,Description,WebTemplate",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.webs.get"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets this web's immediate subwebs.
///
/// Ported from PnPjs `Web.webs` (`packages/sp/webs/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_subwebs(client: &Client<'_>, site_url: &str) -> Result<Vec<WebInfo>> {
    let url = format!(
        "{}/_api/web/webs?$select=Id,Title,Url,ServerRelativeUrl,Description,WebTemplate",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.webs.get_subwebs"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the absolute URL of this web's parent web, or `None` if this is
/// already the root web of the site collection.
///
/// Ported from PnPjs `Web.getParentWeb` (`packages/sp/webs/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_parent_web_url(client: &Client<'_>, site_url: &str) -> Result<Option<String>> {
    let url = format!(
        "{}/_api/web?$select=Url,ParentWeb/ServerRelativeUrl&$expand=ParentWeb",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.webs.get_parent_web_url"), "GET", &url, &[], None)
        .await?;
    let parsed: ParentWebQuery = serde_json::from_slice(&bytes)?;
    Ok(parsed
        .parent_web
        .map(|p| format!("{}{}", origin(&parsed.url), p.server_relative_url)))
}
