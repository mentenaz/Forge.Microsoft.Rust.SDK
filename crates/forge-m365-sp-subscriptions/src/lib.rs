pub mod types;

use crate::types::Subscription;
use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde_json::Value;

// ---------------------------------------------------------------------------
// Operation registrations

#[pnp_operation(id = "sp.subscriptions.list", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn list_subscriptions_op() {}

#[pnp_operation(id = "sp.subscriptions.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_subscription_by_id_op() {}

#[pnp_operation(id = "sp.subscriptions.add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_subscription_op() {}

#[pnp_operation(id = "sp.subscriptions.update", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn update_subscription_op() {}

#[pnp_operation(id = "sp.subscriptions.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_subscription_op() {}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-subscriptions linked in?")
        })
}

// ---------------------------------------------------------------------------
// URL plumbing

/// Doubles embedded `'` and percent-encodes for use inside an OData path
/// segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_path_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

/// `{site}/_api/web/lists/getbytitle('...')` — the list instance the
/// subscriptions nav hangs off (`packages/sp/subscriptions/list.ts`).
fn list_base_url(site_url: &str, list_title: &str) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')",
        site_url.trim_end_matches('/'),
        encode_path_segment(list_title)
    )
}

/// `{site}/_api/web/lists/getbytitle('...')/subscriptions` — the collection
/// addressed by PnPjs `list.subscriptions`.
fn list_url(site_url: &str, list_title: &str) -> String {
    format!("{}/subscriptions", list_base_url(site_url, list_title))
}

/// `.../subscriptions('{id}')` — PnPjs `Subscriptions.getById`
/// (`packages/sp/subscriptions/types.ts:22`).
fn by_id_url(site_url: &str, list_title: &str, subscription_id: &str) -> String {
    format!(
        "{}/subscriptions('{}')",
        list_base_url(site_url, list_title),
        subscription_id.replace('\'', "''")
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

// ---------------------------------------------------------------------------
// Public operations

/// Returns all the webhook subscriptions created for the list.
///
/// Ported from PnPjs `_Subscriptions()` fetch
/// (`packages/sp/subscriptions/types.ts:16` — `getById` when given an id).
pub async fn list_subscriptions(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
) -> Result<Vec<Subscription>> {
    let url = list_url(site_url, list_title);
    let bytes = client
        .run_ladder(entry("sp.subscriptions.list"), "GET", &url, &[], None)
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Returns the webhook subscription with the given id.
///
/// Ported from PnPjs `Subscriptions.getById(id)` + `()`
/// (`packages/sp/subscriptions/types.ts:20-23`).
pub async fn get_subscription_by_id(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    subscription_id: &str,
) -> Result<Subscription> {
    let url = by_id_url(site_url, list_title, subscription_id);
    let bytes = client
        .run_ladder(entry("sp.subscriptions.get_by_id"), "GET", &url, &[], None)
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Creates a new webhook subscription on the list.
///
/// `expiration_date` must be `YYYY-MM-ddTHH:mm:ss+00:00` (max 6 months).
/// `clientState` is only sent when non-empty, matching PnPjs truthy check.
/// `resource` is the subscriptions collection URL (`this.toUrl()`).
///
/// Ported from PnPjs `_Subscriptions.add`
/// (`packages/sp/subscriptions/types.ts:32-45`).
pub async fn add_subscription(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    notification_url: &str,
    expiration_date: &str,
    client_state: Option<&str>,
) -> Result<Subscription> {
    let url = list_url(site_url, list_title);
    let mut body = serde_json::Map::new();
    body.insert(
        "expirationDateTime".to_string(),
        Value::String(expiration_date.to_string()),
    );
    body.insert(
        "notificationUrl".to_string(),
        Value::String(notification_url.to_string()),
    );
    body.insert("resource".to_string(), Value::String(url.clone()));
    if let Some(state) = client_state.filter(|s| !s.is_empty()) {
        body.insert("clientState".to_string(), Value::String(state.to_string()));
    }
    let payload = serde_json::to_vec(&Value::Object(body))?;
    let bytes = client
        .run_ladder(
            entry("sp.subscriptions.add"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    let unwrapped = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&unwrapped)?)
}

/// Renews/updates a webhook subscription. Only the supplied fields are sent
/// (`PATCH`), and empty strings are omitted like PnPjs's truthy check.
///
/// Ported from PnPjs `_Subscription.update`
/// (`packages/sp/subscriptions/types.ts:59-76`).
#[allow(clippy::too_many_arguments)]
pub async fn update_subscription(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    subscription_id: &str,
    expiration_date: Option<&str>,
    notification_url: Option<&str>,
    client_state: Option<&str>,
) -> Result<Subscription> {
    let url = by_id_url(site_url, list_title, subscription_id);
    let mut body = serde_json::Map::new();
    if let Some(date) = expiration_date.filter(|s| !s.is_empty()) {
        body.insert(
            "expirationDateTime".to_string(),
            Value::String(date.to_string()),
        );
    }
    if let Some(nurl) = notification_url.filter(|s| !s.is_empty()) {
        body.insert(
            "notificationUrl".to_string(),
            Value::String(nurl.to_string()),
        );
    }
    if let Some(state) = client_state.filter(|s| !s.is_empty()) {
        body.insert("clientState".to_string(), Value::String(state.to_string()));
    }
    let payload = serde_json::to_vec(&Value::Object(body))?;
    let bytes = client
        .run_ladder(
            entry("sp.subscriptions.update"),
            "PATCH",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    let unwrapped = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&unwrapped)?)
}

/// Removes the webhook subscription.
///
/// Ported from PnPjs `_Subscription.delete`
/// (`packages/sp/subscriptions/types.ts:82-84`).
pub async fn delete_subscription(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    subscription_id: &str,
) -> Result<()> {
    let url = by_id_url(site_url, list_title, subscription_id);
    client
        .run_ladder(entry("sp.subscriptions.delete"), "DELETE", &url, &[], None)
        .await?;
    Ok(())
}
