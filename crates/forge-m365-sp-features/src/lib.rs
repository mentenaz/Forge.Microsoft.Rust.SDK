pub mod types;

use crate::types::{FeatureAddResult, FeatureInfo};
use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde_json::{json, Value};

// ---------------------------------------------------------------------------
// Operation registrations

#[pnp_operation(id = "sp.features.site_list", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn site_list_op() {}

#[pnp_operation(id = "sp.features.site_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn site_by_id_op() {}

#[pnp_operation(id = "sp.features.site_add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn site_add_op() {}

#[pnp_operation(id = "sp.features.site_remove", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn site_remove_op() {}

#[pnp_operation(id = "sp.features.web_list", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn web_list_op() {}

#[pnp_operation(id = "sp.features.web_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn web_by_id_op() {}

#[pnp_operation(id = "sp.features.web_add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn web_add_op() {}

#[pnp_operation(id = "sp.features.web_remove", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn web_remove_op() {}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-features linked in?")
        })
}

// ---------------------------------------------------------------------------
// URL plumbing

/// `{site}/_api/{site|web}/features` — the collection a `Feature` sits in,
/// wired onto both `_Site` (`packages/sp/features/site.ts`) and `_Web`
/// (`packages/sp/features/web.ts`).
fn features_url(site_url: &str, scope: &str) -> String {
    format!("{}/_api/{scope}", site_url.trim_end_matches('/'))
}

fn features_collection_url(site_url: &str, scope: &str) -> String {
    format!("{}/features", features_url(site_url, scope))
}

/// `.../features('{id}')` — PnPjs `Features.getById`
/// (`packages/sp/features/types.ts:38`).
fn feature_by_id_url(site_url: &str, scope: &str, feature_id: &str) -> String {
    format!(
        "{}/features('{}')",
        features_url(site_url, scope),
        feature_id.replace('\'', "''")
    )
}

/// Unwraps the OData verbose envelope (`{"d": ...}` and the
/// `{"results": [...]}` form for collections) so the payload can be
/// deserialized directly.
fn unwrap_odata(bytes: &[u8]) -> Result<Vec<u8>> {
    let mut root: Value = serde_json::from_slice(bytes)?;
    if let Some(d) = root.get_mut("d").map(Value::take) {
        root = d;
    }
    if let Some(inner) = root.get_mut("results").map(Value::take) {
        root = inner;
    }
    Ok(serde_json::to_vec(&root)?)
}

/// Activation action body shared by `add` on both scopes. Key/order per
/// PnPjs `_Features.add` (`packages/sp/features/types.ts:21`).
fn add_body(feature_id: &str, force: bool) -> Result<Vec<u8>> {
    Ok(serde_json::to_vec(&json!({
        "featdefScope": 0,
        "featureId": feature_id,
        "force": force,
    }))?)
}

/// Deactivation body shared by `remove` (`types.ts:50`).
fn remove_body(feature_id: &str, force: bool) -> Result<Vec<u8>> {
    Ok(serde_json::to_vec(&json!({
        "featureId": feature_id,
        "force": force,
    }))?)
}

// ---------------------------------------------------------------------------
// Site-scoped operations

/// Returns all the features activated in the site.
pub async fn get_site_features(client: &Client<'_>, site_url: &str) -> Result<Vec<FeatureInfo>> {
    let url = features_collection_url(site_url, "site");
    let bytes = client
        .run_ladder(entry("sp.features.site_list"), "GET", &url, &[], None)
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Returns a single site feature by its guid.
pub async fn get_site_feature_by_id(
    client: &Client<'_>,
    site_url: &str,
    feature_id: &str,
) -> Result<FeatureInfo> {
    let url = feature_by_id_url(site_url, "site", feature_id);
    let bytes = client
        .run_ladder(entry("sp.features.site_by_id"), "GET", &url, &[], None)
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Activates a feature in the site. Posts to `features/add` with
/// `featdefScope: 0`, mirroring PnPjs `_Features.add`
/// (`packages/sp/features/types.ts:19-31`).
pub async fn activate_site_feature(
    client: &Client<'_>,
    site_url: &str,
    feature_id: &str,
    force: bool,
) -> Result<FeatureAddResult> {
    let url = format!("{}/add", features_collection_url(site_url, "site"));
    let payload = add_body(feature_id, force)?;
    let bytes = client
        .run_ladder(
            entry("sp.features.site_add"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    let unwrapped = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&unwrapped)?)
}

/// Deactivates a feature in the site. Posts to `features/remove`, mirroring
/// PnPjs `_Features.remove` (`packages/sp/features/types.ts:48-54`).
pub async fn deactivate_site_feature(
    client: &Client<'_>,
    site_url: &str,
    feature_id: &str,
    force: bool,
) -> Result<()> {
    let url = format!("{}/remove", features_collection_url(site_url, "site"));
    let payload = remove_body(feature_id, force)?;
    client
        .run_ladder(
            entry("sp.features.site_remove"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Web-scoped operations

/// Returns all the features activated in the web.
pub async fn get_web_features(client: &Client<'_>, site_url: &str) -> Result<Vec<FeatureInfo>> {
    let url = features_collection_url(site_url, "web");
    let bytes = client
        .run_ladder(entry("sp.features.web_list"), "GET", &url, &[], None)
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Returns a single web feature by its guid.
pub async fn get_web_feature_by_id(
    client: &Client<'_>,
    site_url: &str,
    feature_id: &str,
) -> Result<FeatureInfo> {
    let url = feature_by_id_url(site_url, "web", feature_id);
    let bytes = client
        .run_ladder(entry("sp.features.web_by_id"), "GET", &url, &[], None)
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Activates a feature in the web. Posts to `features/add` with
/// `featdefScope: 0` (`packages/sp/features/types.ts:19-31`).
pub async fn activate_web_feature(
    client: &Client<'_>,
    site_url: &str,
    feature_id: &str,
    force: bool,
) -> Result<FeatureAddResult> {
    let url = format!("{}/add", features_collection_url(site_url, "web"));
    let payload = add_body(feature_id, force)?;
    let bytes = client
        .run_ladder(
            entry("sp.features.web_add"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    let unwrapped = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&unwrapped)?)
}

/// Deactivates a feature in the web. Posts to `features/remove`
/// (`packages/sp/features/types.ts:48-54`).
pub async fn deactivate_web_feature(
    client: &Client<'_>,
    site_url: &str,
    feature_id: &str,
    force: bool,
) -> Result<()> {
    let url = format!("{}/remove", features_collection_url(site_url, "web"));
    let payload = remove_body(feature_id, force)?;
    client
        .run_ladder(
            entry("sp.features.web_remove"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(())
}
