use serde::Deserialize;

/// A single webhook subscription on a list, as returned by the
/// `/_api/web/lists/getbytitle(...)/subscriptions` service
/// (`packages/sp/subscriptions/types.ts`). PnPjs treats the payload as
/// untyped (`Promise<any>`); fields are tolerant (`#[serde(default)]`).
#[derive(Debug, Clone, Default, Deserialize)]
pub struct Subscription {
    #[serde(rename = "id", default)]
    pub id: String,
    /// The resource (list) this webhook is subscribed to.
    #[serde(rename = "resource", default)]
    pub resource: String,
    /// URL that receives the notifications.
    #[serde(rename = "notificationUrl", default)]
    pub notification_url: String,
    /// Expiry in `YYYY-MM-ddTHH:mm:ss+00:00` form (max 6 months).
    #[serde(rename = "expirationDateTime", default)]
    pub expiration_date_time: String,
    /// Optional client-specific string echoed in notifications.
    #[serde(rename = "clientState", default)]
    pub client_state: Option<String>,
    #[serde(rename = "lastModifiedTime", default)]
    pub last_modified_time: Option<String>,
}
