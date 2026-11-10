use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};

/// Bitflag values for `SP.Utilities.Utility`'s `PrincipalType` enum. Kept as
/// plain `u32` constants (rather than a Rust `enum`) because the wire value
/// is a bitmask a caller may OR together (e.g. `User | SharePointGroup`),
/// which a C-style enum can't represent cleanly.
///
/// Ported from PnPjs `PrincipalType` (`packages/sp/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub struct PrincipalType;
impl PrincipalType {
    pub const NONE: u32 = 0;
    pub const USER: u32 = 1;
    pub const DISTRIBUTION_LIST: u32 = 2;
    pub const SECURITY_GROUP: u32 = 4;
    pub const SHARE_POINT_GROUP: u32 = 8;
    pub const ALL: u32 =
        Self::USER | Self::DISTRIBUTION_LIST | Self::SECURITY_GROUP | Self::SHARE_POINT_GROUP;
}

/// Bitflag values for `SP.Utilities.Utility`'s `PrincipalSource` enum. See
/// [`PrincipalType`] for why this is constants rather than an enum.
///
/// Ported from PnPjs `PrincipalSource` (`packages/sp/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub struct PrincipalSource;
impl PrincipalSource {
    pub const NONE: u32 = 0;
    pub const USER_INFO_LIST: u32 = 1;
    pub const WINDOWS: u32 = 2;
    pub const MEMBERSHIP_PROVIDER: u32 = 4;
    pub const ROLE_PROVIDER: u32 = 8;
    pub const ALL: u32 =
        Self::USER_INFO_LIST | Self::WINDOWS | Self::MEMBERSHIP_PROVIDER | Self::ROLE_PROVIDER;
}

/// Ported from PnPjs `IPrincipalInfo` (`packages/sp/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct PrincipalInfo {
    pub department: String,
    pub display_name: String,
    pub email: String,
    pub job_title: String,
    pub login_name: String,
    pub mobile: String,
    pub principal_id: i64,
    pub principal_type: u32,
    #[serde(rename = "SIPAddress")]
    pub sip_address: String,
}

/// Ported from PnPjs `IEmailProperties` (`packages/sp/sputilities/types.ts`)
/// @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`. `additional_headers` is
/// a plain key/value map here; PnPjs remaps it into a `{Key,Value,ValueType}`
/// triple array before sending (see `sendEmail`'s wire encoding below).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EmailProperties {
    #[serde(rename = "To")]
    pub to: Vec<String>,
    #[serde(rename = "CC", skip_serializing_if = "Option::is_none")]
    pub cc: Option<Vec<String>>,
    #[serde(rename = "BCC", skip_serializing_if = "Option::is_none")]
    pub bcc: Option<Vec<String>>,
    #[serde(rename = "Subject")]
    pub subject: String,
    #[serde(rename = "Body")]
    pub body: String,
    #[serde(skip)]
    pub additional_headers: Option<std::collections::BTreeMap<String, String>>,
    #[serde(rename = "From", skip_serializing_if = "Option::is_none")]
    pub from: Option<String>,
}

fn utility_url(site_url: &str, method: &str) -> String {
    format!(
        "{}/_api/SP.Utilities.Utility.{method}",
        site_url.trim_end_matches('/')
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-utilities linked in?")
        })
}

#[pnp_operation(id = "sp.utilities.send_email", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn send_email_op() {}

#[pnp_operation(id = "sp.utilities.get_current_user_email_addresses", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_current_user_email_addresses_op() {}

#[pnp_operation(id = "sp.utilities.resolve_principal", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn resolve_principal_op() {}

#[pnp_operation(id = "sp.utilities.search_principals", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn search_principals_op() {}

#[pnp_operation(id = "sp.utilities.create_email_body_for_invitation", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_email_body_for_invitation_op() {}

#[pnp_operation(id = "sp.utilities.expand_groups_to_principals", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn expand_groups_to_principals_op() {}

/// Sends an email through SharePoint. Deprecated upstream by Microsoft in
/// favor of Graph's `sendMail`, but still a valid REST call.
///
/// Ported from PnPjs `Utilities.sendEmail`
/// (`packages/sp/sputilities/types.ts`) @ pnpjs `8ee2375d`.
pub async fn send_email(
    client: &Client<'_>,
    site_url: &str,
    properties: &EmailProperties,
) -> Result<()> {
    let mut props = serde_json::to_value(properties)?;
    if let Some(headers) = &properties.additional_headers {
        let encoded: Vec<Value> = headers
            .iter()
            .map(|(k, v)| json!({ "Key": k, "Value": v, "ValueType": "Edm.String" }))
            .collect();
        props["AdditionalHeaders"] = Value::Array(encoded);
    }
    let body = serde_json::to_vec(&json!({ "properties": props }))?;
    client
        .run_ladder(
            entry("sp.utilities.send_email"),
            "POST",
            &utility_url(site_url, "SendEmail"),
            &[],
            Some(&body),
        )
        .await?;
    Ok(())
}

/// Returns the current user's email addresses known to SharePoint, as a
/// single semicolon-delimited string (matches the raw SharePoint response
/// shape; PnPjs returns it unmodified too).
///
/// Ported from PnPjs `Utilities.getCurrentUserEmailAddresses`
/// (`packages/sp/sputilities/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_current_user_email_addresses(
    client: &Client<'_>,
    site_url: &str,
) -> Result<String> {
    let bytes = client
        .run_ladder(
            entry("sp.utilities.get_current_user_email_addresses"),
            "POST",
            &utility_url(site_url, "GetCurrentUserEmailAddresses"),
            &[],
            Some(b"{}"),
        )
        .await?;
    let value: Value = serde_json::from_slice(&bytes)?;
    Ok(odata_string_value(&value))
}

/// Resolves a principal (user/group) matching the given search input.
///
/// Ported from PnPjs `Utilities.resolvePrincipal`
/// (`packages/sp/sputilities/types.ts`) @ pnpjs `8ee2375d`.
#[allow(clippy::too_many_arguments)]
pub async fn resolve_principal(
    client: &Client<'_>,
    site_url: &str,
    input: &str,
    scopes: u32,
    sources: u32,
    input_is_email_only: bool,
    add_to_user_info_list: bool,
    match_user_info_list: bool,
) -> Result<PrincipalInfo> {
    let body = serde_json::to_vec(&json!({
        "input": input,
        "scopes": scopes,
        "sources": sources,
        "inputIsEmailOnly": input_is_email_only,
        "addToUserInfoList": add_to_user_info_list,
        "matchUserInfoList": match_user_info_list,
    }))?;
    let bytes = client
        .run_ladder(
            entry("sp.utilities.resolve_principal"),
            "POST",
            &utility_url(site_url, "ResolvePrincipalInCurrentContext"),
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Searches for principals matching the given input.
///
/// Ported from PnPjs `Utilities.searchPrincipals`
/// (`packages/sp/sputilities/types.ts`) @ pnpjs `8ee2375d`.
pub async fn search_principals(
    client: &Client<'_>,
    site_url: &str,
    input: &str,
    scopes: u32,
    sources: u32,
    group_name: &str,
    max_count: i32,
) -> Result<Vec<PrincipalInfo>> {
    let body = serde_json::to_vec(&json!({
        "input": input,
        "scopes": scopes,
        "sources": sources,
        "groupName": group_name,
        "maxCount": max_count,
    }))?;
    let bytes = client
        .run_ladder(
            entry("sp.utilities.search_principals"),
            "POST",
            &utility_url(site_url, "SearchPrincipalsUsingContextWeb"),
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the external URL to a document or resource, for use in an invitation
/// email body.
///
/// Ported from PnPjs `Utilities.createEmailBodyForInvitation`
/// (`packages/sp/sputilities/types.ts`) @ pnpjs `8ee2375d`.
pub async fn create_email_body_for_invitation(
    client: &Client<'_>,
    site_url: &str,
    page_address: &str,
) -> Result<String> {
    let body = serde_json::to_vec(&json!({ "pageAddress": page_address }))?;
    let bytes = client
        .run_ladder(
            entry("sp.utilities.create_email_body_for_invitation"),
            "POST",
            &utility_url(site_url, "CreateEmailBodyForInvitation"),
            &[],
            Some(&body),
        )
        .await?;
    let value: Value = serde_json::from_slice(&bytes)?;
    Ok(odata_string_value(&value))
}

/// Expands one or more groups into their member principals.
///
/// Ported from PnPjs `Utilities.expandGroupsToPrincipals`
/// (`packages/sp/sputilities/types.ts`) @ pnpjs `8ee2375d`.
pub async fn expand_groups_to_principals(
    client: &Client<'_>,
    site_url: &str,
    inputs: &[String],
    max_count: i32,
) -> Result<Vec<PrincipalInfo>> {
    let body = serde_json::to_vec(&json!({ "inputs": inputs, "maxCount": max_count }))?;
    let bytes = client
        .run_ladder(
            entry("sp.utilities.expand_groups_to_principals"),
            "POST",
            &utility_url(site_url, "ExpandGroupsToPrincipals"),
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// SharePoint's `odata=verbose` string-returning actions wrap the value as
/// `{"value": "..."}`; classic (non-verbose) responses return the bare JSON
/// string. Handle both since `AuthedTransport` requests verbose by default
/// but a caller-supplied transport may not.
fn odata_string_value(value: &Value) -> String {
    match value {
        Value::String(s) => s.clone(),
        Value::Object(map) => map
            .get("value")
            .and_then(|v| v.as_str())
            .unwrap_or_default()
            .to_string(),
        _ => String::new(),
    }
}
