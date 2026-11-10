pub mod types;

use crate::types::{InstalledLanguageInfo, RegionalSettingsInfo, TimeZoneInfo};
use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;
use serde_json::{json, Value};

// ---------------------------------------------------------------------------
// Operation registrations

#[pnp_operation(id = "sp.regional_settings.get", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn regional_settings_get_op() {}

#[pnp_operation(id = "sp.regional_settings.time_zone", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn time_zone_op() {}

#[pnp_operation(id = "sp.regional_settings.time_zones", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn time_zones_op() {}

#[pnp_operation(
    id = "sp.regional_settings.time_zone_by_id",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn time_zone_by_id_op() {}

#[pnp_operation(
    id = "sp.regional_settings.installed_languages",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn installed_languages_op() {}

#[pnp_operation(
    id = "sp.regional_settings.utc_to_local_time",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn utc_to_local_time_op() {}

#[pnp_operation(
    id = "sp.regional_settings.local_time_to_utc",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn local_time_to_utc_op() {}

#[pnp_operation(
    id = "sp.regional_settings.web_title_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn web_title_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.web_description_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn web_description_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.list_title_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn list_title_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.list_description_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn list_description_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.field_title_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn field_title_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.field_description_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn field_description_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.content_type_name_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn content_type_name_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.content_type_description_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn content_type_description_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.usercustomaction_title_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn usercustomaction_title_resource_op() {}

#[pnp_operation(
    id = "sp.regional_settings.usercustomaction_description_resource",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn usercustomaction_description_resource_op() {}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-regional-settings linked in?")
        })
}

// ---------------------------------------------------------------------------
// URL plumbing

/// Doubles embedded `'` and percent-encodes for use inside an OData path
/// segment. Ported from PnPjs `encodePath`.
fn encode_path_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

/// `{site}/_api/web/regionalsettings` (`packages/sp/regional-settings
/// /types.ts:11`).
fn regional_url(site_url: &str) -> String {
    format!(
        "{}/_api/web/regionalsettings",
        site_url.trim_end_matches('/')
    )
}

// The `getValueForUICulture` binder posts to
// `{queryable}/{resource}/getValueForUICulture` with `{cultureName}`
// (`packages/sp/regional-settings/funcs.ts`); see the async helper
// `post_resource_for_culture` below.

// ---------------------------------------------------------------------------
// Parsing helpers

/// Unwraps the OData verbose envelope (`{"d": ...}` and the
/// `{"results": [...]}` form) so the payload can be deserialized directly.
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

/// Resolves a string-valued payload. PnPjs peels the named key when the
/// service boxes the value (`hOP` check in `utcToLocalTime`/
/// `localTimeToUTC`; the culture bounder returns the raw string); here the
/// envelope is unwrapped first and a single trailing `{key: "..."}` object is
/// reduced the same way.
fn string_payload(bytes: &[u8]) -> Result<String> {
    let unwrapped = unwrap_odata(bytes)?;
    let value: Value = serde_json::from_slice(&unwrapped)?;
    match value {
        Value::String(s) => Ok(s),
        Value::Object(map) => {
            let (_, v) = map.into_iter().next().ok_or_else(|| {
                forge_m365_core::Error::Surface(Surface::SpRest, "empty object result".into())
            })?;
            v.as_str().map(str::to_string).ok_or_else(|| {
                forge_m365_core::Error::Surface(Surface::SpRest, "non-string result value".into())
            })
        }
        _ => Err(forge_m365_core::Error::Surface(
            Surface::SpRest,
            "non-string result".into(),
        )),
    }
}

// ---------------------------------------------------------------------------
// Regional settings + time zones

/// Returns the regional settings of the web.
pub async fn get_regional_settings(
    client: &Client<'_>,
    site_url: &str,
) -> Result<RegionalSettingsInfo> {
    let url = regional_url(site_url);
    let bytes = client
        .run_ladder(entry("sp.regional_settings.get"), "GET", &url, &[], None)
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Returns the web's current time zone
/// (`_RegionalSettings.timeZone`, `packages/sp/regional-settings/types.ts:17`).
pub async fn get_time_zone(client: &Client<'_>, site_url: &str) -> Result<TimeZoneInfo> {
    let url = format!("{}/timezone", regional_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.regional_settings.time_zone"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Returns all time zones SharePoint knows about.
pub async fn get_time_zones(client: &Client<'_>, site_url: &str) -> Result<Vec<TimeZoneInfo>> {
    let url = format!("{}/timezones", regional_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.regional_settings.time_zones"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Returns the time zone with the given integer id. PnPjs `_TimeZones.getById`
/// posts to `timezones/GetById({id})` (`types.ts:92`).
pub async fn get_time_zone_by_id(
    client: &Client<'_>,
    site_url: &str,
    id: i64,
) -> Result<TimeZoneInfo> {
    let url = format!("{}/timezones/GetById({id})", regional_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.regional_settings.time_zone_by_id"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice(&payload)?)
}

/// Returns the languages installed in the server farm. PnPjs reads the
/// `Items` property of the installedlanguages response (`types.ts:31`).
pub async fn get_installed_languages(
    client: &Client<'_>,
    site_url: &str,
) -> Result<Vec<InstalledLanguageInfo>> {
    let url = format!("{}/installedlanguages", regional_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.regional_settings.installed_languages"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    #[derive(Deserialize)]
    struct Items {
        #[serde(rename = "Items")]
        items: Vec<InstalledLanguageInfo>,
    }
    let payload = unwrap_odata(&bytes)?;
    Ok(serde_json::from_slice::<Items>(&payload)?.items)
}

/// Converts a UTC ISO-8601 string to the web's local time
/// (`_TimeZone.utcToLocalTime`, `types.ts:47`).
pub async fn utc_to_local_time(
    client: &Client<'_>,
    site_url: &str,
    utc_iso: &str,
) -> Result<String> {
    let url = format!(
        "{}/timezone/utctolocaltime('{}')",
        regional_url(site_url),
        utc_iso.replace('\'', "''")
    );
    let bytes = client
        .run_ladder(
            entry("sp.regional_settings.utc_to_local_time"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    string_payload(&bytes)
}

/// Converts a local ISO-8601 string to UTC
/// (`_TimeZone.localTimeToUTC`, `types.ts:66`).
pub async fn local_time_to_utc(
    client: &Client<'_>,
    site_url: &str,
    local_iso: &str,
) -> Result<String> {
    let url = format!(
        "{}/timezone/localtimetoutc('{}')",
        regional_url(site_url),
        local_iso.replace('\'', "''")
    );
    let bytes = client
        .run_ladder(
            entry("sp.regional_settings.local_time_to_utc"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    string_payload(&bytes)
}

// ---------------------------------------------------------------------------
// UI-culture resource strings (`IUserResources`)

async fn post_resource_for_culture(
    client: &Client<'_>,
    op_id: &'static str,
    queryable_url: &str,
    resource: &str,
    culture_name: &str,
) -> Result<String> {
    let url = format!("{queryable_url}/{resource}/getValueForUICulture");
    let payload = serde_json::to_vec(&json!({ "cultureName": culture_name }))?;
    let bytes = client
        .run_ladder(entry(op_id), "POST", &url, &[], Some(&payload))
        .await?;
    string_payload(&bytes)
}

/// Web title in the given culture (`web.ts`, `funcs.ts`).
pub async fn web_title_resource(
    client: &Client<'_>,
    site_url: &str,
    culture_name: &str,
) -> Result<String> {
    let url = format!("{}/_api/web", site_url.trim_end_matches('/'));
    post_resource_for_culture(
        client,
        "sp.regional_settings.web_title_resource",
        &url,
        "titleResource",
        culture_name,
    )
    .await
}

/// Web description in the given culture.
pub async fn web_description_resource(
    client: &Client<'_>,
    site_url: &str,
    culture_name: &str,
) -> Result<String> {
    let url = format!("{}/_api/web", site_url.trim_end_matches('/'));
    post_resource_for_culture(
        client,
        "sp.regional_settings.web_description_resource",
        &url,
        "descriptionResource",
        culture_name,
    )
    .await
}

fn list_queryable(site_url: &str, list_title: &str) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')",
        site_url.trim_end_matches('/'),
        encode_path_segment(list_title)
    )
}

/// List title in the given culture (`list.ts`).
pub async fn list_title_resource(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    culture_name: &str,
) -> Result<String> {
    post_resource_for_culture(
        client,
        "sp.regional_settings.list_title_resource",
        &list_queryable(site_url, list_title),
        "titleResource",
        culture_name,
    )
    .await
}

/// List description in the given culture (`list.ts`).
pub async fn list_description_resource(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    culture_name: &str,
) -> Result<String> {
    post_resource_for_culture(
        client,
        "sp.regional_settings.list_description_resource",
        &list_queryable(site_url, list_title),
        "descriptionResource",
        culture_name,
    )
    .await
}

fn field_queryable(site_url: &str, list_title: &str, field_name: &str) -> String {
    format!(
        "{}/fields/getByInternalNameOrTitle('{}')",
        list_queryable(site_url, list_title),
        encode_path_segment(field_name)
    )
}

/// Field title in the given culture (`field.ts`).
pub async fn field_title_resource(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    field_name: &str,
    culture_name: &str,
) -> Result<String> {
    post_resource_for_culture(
        client,
        "sp.regional_settings.field_title_resource",
        &field_queryable(site_url, list_title, field_name),
        "titleResource",
        culture_name,
    )
    .await
}

/// Field description in the given culture (`field.ts`).
pub async fn field_description_resource(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    field_name: &str,
    culture_name: &str,
) -> Result<String> {
    post_resource_for_culture(
        client,
        "sp.regional_settings.field_description_resource",
        &field_queryable(site_url, list_title, field_name),
        "descriptionResource",
        culture_name,
    )
    .await
}

/// Content type name in the given culture. PnPjs binds `nameResource` as the
/// content type's title resource (`content-type.ts`).
pub async fn content_type_name_resource(
    client: &Client<'_>,
    site_url: &str,
    content_type_id: &str,
    culture_name: &str,
) -> Result<String> {
    let url = format!(
        "{}/_api/web/contenttypes/getbyid('{}')",
        site_url.trim_end_matches('/'),
        content_type_id.replace('\'', "''")
    );
    post_resource_for_culture(
        client,
        "sp.regional_settings.content_type_name_resource",
        &url,
        "nameResource",
        culture_name,
    )
    .await
}

/// Content type description in the given culture (`content-type.ts`).
pub async fn content_type_description_resource(
    client: &Client<'_>,
    site_url: &str,
    content_type_id: &str,
    culture_name: &str,
) -> Result<String> {
    let url = format!(
        "{}/_api/web/contenttypes/getbyid('{}')",
        site_url.trim_end_matches('/'),
        content_type_id.replace('\'', "''")
    );
    post_resource_for_culture(
        client,
        "sp.regional_settings.content_type_description_resource",
        &url,
        "descriptionResource",
        culture_name,
    )
    .await
}

/// User custom action title in the given culture
/// (`user-custom-actions.ts`; web-scoped action instances).
pub async fn usercustomaction_title_resource(
    client: &Client<'_>,
    site_url: &str,
    action_id: &str,
    culture_name: &str,
) -> Result<String> {
    let url = format!(
        "{}/_api/web/usercustomactions('{}')",
        site_url.trim_end_matches('/'),
        action_id.replace('\'', "''")
    );
    post_resource_for_culture(
        client,
        "sp.regional_settings.usercustomaction_title_resource",
        &url,
        "titleResource",
        culture_name,
    )
    .await
}

/// User custom action description in the given culture
/// (`user-custom-actions.ts`).
pub async fn usercustomaction_description_resource(
    client: &Client<'_>,
    site_url: &str,
    action_id: &str,
    culture_name: &str,
) -> Result<String> {
    let url = format!(
        "{}/_api/web/usercustomactions('{}')",
        site_url.trim_end_matches('/'),
        action_id.replace('\'', "''")
    );
    post_resource_for_culture(
        client,
        "sp.regional_settings.usercustomaction_description_resource",
        &url,
        "descriptionResource",
        culture_name,
    )
    .await
}
