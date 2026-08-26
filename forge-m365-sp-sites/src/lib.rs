use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;

/// Subset of SharePoint's `SP.Site` fields, hand-ported from PnPjs's `ISiteInfo`
/// (`packages/sp/sites/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct SiteInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Url")]
    pub url: String,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
    #[serde(rename = "Classification")]
    pub classification: String,
    #[serde(rename = "IsHubSite")]
    pub is_hub_site: bool,
    #[serde(rename = "HubSiteId")]
    pub hub_site_id: String,
}

#[derive(Debug, Clone, Deserialize)]
struct RootWebUrl {
    #[serde(rename = "Url")]
    url: String,
}

/// Hand-ported from PnPjs's `IDocumentLibraryInformation`
/// (`packages/sp/sites/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct DocumentLibraryInfo {
    #[serde(rename = "AbsoluteUrl")]
    pub absolute_url: String,
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "IsDefaultDocumentLibrary")]
    pub is_default_document_library: bool,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
    #[serde(rename = "Title")]
    pub title: String,
}

/// Some `_api/sp.web.*` OData functions wrap their payload in an object keyed by
/// the function's PascalCase name (e.g. `{"GetDocumentLibraries": [...]}`); others
/// return the bare value. PnPjs handles both by checking for the key at runtime,
/// which this mirrors.
fn unwrap_result<T: serde::de::DeserializeOwned>(bytes: &[u8], key: &str) -> Result<T> {
    let mut value: serde_json::Value = serde_json::from_slice(bytes)?;
    if let Some(inner) = value.get_mut(key).map(serde_json::Value::take) {
        value = inner;
    }
    Ok(serde_json::from_value(value)?)
}

/// Builds the `@v='<value>'` aliased-parameter query fragment SharePoint REST
/// uses for OData function arguments, e.g. `_api/sp.web.getdocumentlibraries(@v)?@v=...`.
fn v_param(value: &str) -> String {
    utf8_percent_encode(&format!("'{value}'"), NON_ALPHANUMERIC).to_string()
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-sites linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.sites.get", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_site_op() {}

#[pnp_operation(id = "sp.sites.get_root_web_url", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_root_web_url_op() {}

#[pnp_operation(id = "sp.sites.get_document_libraries", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_document_libraries_op() {}

#[pnp_operation(id = "sp.sites.get_web_url_from_page_url", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_web_url_from_page_url_op() {}

/// Gets site-collection properties for `site_url`.
///
/// Ported from PnPjs `Site.get` (`packages/sp/sites/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_site(client: &Client<'_>, site_url: &str) -> Result<SiteInfo> {
    let url = format!(
        "{}/_api/site?$select=Id,Url,ServerRelativeUrl,Classification,IsHubSite,HubSiteId",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.sites.get"), "GET", &url, None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the absolute URL of the site collection's root web.
///
/// Ported from PnPjs `Site.getRootWeb` (`packages/sp/sites/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_root_web_url(client: &Client<'_>, site_url: &str) -> Result<String> {
    let url = format!(
        "{}/_api/site/rootweb?$select=Url",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(entry("sp.sites.get_root_web_url"), "GET", &url, None)
        .await?;
    let parsed: RootWebUrl = serde_json::from_slice(&bytes)?;
    Ok(parsed.url)
}

/// Gets the document libraries on the web at `absolute_web_url`.
///
/// Ported from PnPjs `Site.getDocumentLibraries` (`packages/sp/sites/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_document_libraries(
    client: &Client<'_>,
    site_url: &str,
    absolute_web_url: &str,
) -> Result<Vec<DocumentLibraryInfo>> {
    let url = format!(
        "{}/_api/sp.web.getdocumentlibraries(@v)?@v={}",
        site_url.trim_end_matches('/'),
        v_param(absolute_web_url)
    );
    let bytes = client
        .run_ladder(entry("sp.sites.get_document_libraries"), "GET", &url, None)
        .await?;
    unwrap_result(&bytes, "GetDocumentLibraries")
}

/// Gets the site url that hosts `absolute_page_url`.
///
/// Ported from PnPjs `Site.getWebUrlFromPageUrl` (`packages/sp/sites/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn get_web_url_from_page_url(
    client: &Client<'_>,
    site_url: &str,
    absolute_page_url: &str,
) -> Result<String> {
    let url = format!(
        "{}/_api/sp.web.getweburlfrompageurl(@v)?@v={}",
        site_url.trim_end_matches('/'),
        v_param(absolute_page_url)
    );
    let bytes = client
        .run_ladder(
            entry("sp.sites.get_web_url_from_page_url"),
            "GET",
            &url,
            None,
        )
        .await?;
    unwrap_result(&bytes, "GetWebUrlFromPageUrl")
}
