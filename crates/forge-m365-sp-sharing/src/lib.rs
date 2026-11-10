pub mod types;

use crate::types::{
    ObjectSharingSettings, RoleType, ShareLinkResponse, ShareObjectOptions, SharingEmailData,
    SharingEntityPermission, SharingInformation, SharingInformationRequest, SharingLinkKind,
    SharingRecipient, SharingResult, SharingRole, EMPTY_GUID,
};
use forge_m365_core::{registered_operations, Client, Error, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;
use serde_json::{json, Value};

// ---------------------------------------------------------------------------
// Operation registrations

#[pnp_operation(id = "sp.sharing.role_definition_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn role_definition_id_op() {}

#[pnp_operation(id = "sp.sharing.associated_group_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn associated_group_id_op() {}

#[pnp_operation(id = "sp.sharing.item_share_with", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn item_share_with_op() {}

#[pnp_operation(id = "sp.sharing.item_share_link", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn item_share_link_op() {}

#[pnp_operation(id = "sp.sharing.item_check_permissions", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn item_check_permissions_op() {}

#[pnp_operation(id = "sp.sharing.item_sharing_information", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn item_sharing_information_op() {}

#[pnp_operation(id = "sp.sharing.item_object_sharing_settings", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn item_object_sharing_settings_op() {}

#[pnp_operation(id = "sp.sharing.item_unshare", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn item_unshare_op() {}

#[pnp_operation(id = "sp.sharing.item_delete_link_by_kind", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn item_delete_link_by_kind_op() {}

#[pnp_operation(id = "sp.sharing.item_unshare_link", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn item_unshare_link_op() {}

#[pnp_operation(id = "sp.sharing.file_share_with", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn file_share_with_op() {}

#[pnp_operation(id = "sp.sharing.file_share_link", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn file_share_link_op() {}

#[pnp_operation(id = "sp.sharing.file_check_permissions", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn file_check_permissions_op() {}

#[pnp_operation(id = "sp.sharing.file_sharing_information", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn file_sharing_information_op() {}

#[pnp_operation(id = "sp.sharing.file_object_sharing_settings", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn file_object_sharing_settings_op() {}

#[pnp_operation(id = "sp.sharing.file_unshare", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn file_unshare_op() {}

#[pnp_operation(id = "sp.sharing.file_delete_link_by_kind", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn file_delete_link_by_kind_op() {}

#[pnp_operation(id = "sp.sharing.file_unshare_link", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn file_unshare_link_op() {}

#[pnp_operation(id = "sp.sharing.folder_share_with", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn folder_share_with_op() {}

#[pnp_operation(id = "sp.sharing.folder_share_link", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn folder_share_link_op() {}

#[pnp_operation(id = "sp.sharing.folder_check_permissions", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn folder_check_permissions_op() {}

#[pnp_operation(id = "sp.sharing.folder_sharing_information", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn folder_sharing_information_op() {}

#[pnp_operation(id = "sp.sharing.folder_object_sharing_settings", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn folder_object_sharing_settings_op() {}

#[pnp_operation(id = "sp.sharing.folder_unshare", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn folder_unshare_op() {}

#[pnp_operation(id = "sp.sharing.folder_delete_link_by_kind", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn folder_delete_link_by_kind_op() {}

#[pnp_operation(id = "sp.sharing.folder_unshare_link", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn folder_unshare_link_op() {}

#[pnp_operation(id = "sp.sharing.web_share_with", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn web_share_with_op() {}

#[pnp_operation(id = "sp.sharing.web_share_object", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn web_share_object_op() {}

#[pnp_operation(id = "sp.sharing.web_share_object_raw", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn web_share_object_raw_op() {}

#[pnp_operation(id = "sp.sharing.web_unshare_object", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn web_unshare_object_op() {}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-sharing linked in?")
        })
}

// ---------------------------------------------------------------------------
// URL plumbing

/// Doubles embedded `'` and percent-encodes for use inside an OData
/// `decodedUrl='...'` path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_path_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

/// Instance URL for a list item share target:
/// `{site}/_api/web/lists/getbytitle('...')/items({id})`.
fn item_url(site_url: &str, list_title: &str, item_id: i64) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')/items({item_id})",
        site_url.trim_end_matches('/'),
        encode_path_segment(list_title)
    )
}

/// File share target. PnPjs `_File` resolves its item via the
/// `listItemAllFields` navigation (`getItem()` returns that nav instance
/// directly — no fetch), so sharing ops land on
/// `{site}/_api/web/getFileByServerRelativePath(decodedUrl='...')/listItemAllFields`
/// (`packages/sp/sharing/file.ts`, `files/types.ts:490`).
fn file_url(site_url: &str, server_relative_path: &str) -> String {
    format!(
        "{}/_api/web/getFileByServerRelativePath(decodedUrl='{}')/listItemAllFields",
        site_url.trim_end_matches('/'),
        encode_path_segment(server_relative_path)
    )
}

/// Folder share target, mirror of the file case through
/// `getFolderByServerRelativePath` (`packages/sp/folders/types.ts:302`).
fn folder_url(site_url: &str, server_relative_path: &str) -> String {
    format!(
        "{}/_api/web/getFolderByServerRelativePath(decodedUrl='{}')/listItemAllFields",
        site_url.trim_end_matches('/'),
        encode_path_segment(server_relative_path)
    )
}

/// Static `SP.Web.ShareObject` endpoint (`packages/sp/sharing/funcs.ts:217`).
fn share_object_static_url(site_url: &str) -> String {
    format!(
        "{}/_api/SP.Web.ShareObject?$expand=UsersWithAccessRequests,GroupsSharedWith",
        site_url.trim_end_matches('/')
    )
}

// ---------------------------------------------------------------------------
// Role lookups

#[derive(Debug, Deserialize)]
struct IdOnly {
    #[serde(rename = "Id")]
    id: i64,
}

/// Unwraps an OData verbose envelope (`{"d": ...}` and the `{"results": [...]}`
/// form for collections) so the payload can be deserialized directly. Same
/// defensive handling used across the SP crates (`sp-security`, `sp-views`).
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

/// Looks up the first role definition whose `RoleTypeKind` matches the given
/// role. `top_one` renders `$top=1` like PnPjs `getRoleValue`
/// (`packages/sp/sharing/funcs.ts:250`); `shareWith`'s lookup omits it
/// (`funcs.ts:191`).
async fn lookup_role_definition_id(
    client: &Client<'_>,
    site_url: &str,
    role: SharingRole,
    top_one: bool,
) -> Result<i64> {
    let role_filter = if role == SharingRole::Edit {
        i64::from(RoleType::Contributor)
    } else {
        i64::from(RoleType::Reader)
    };
    let top = if top_one { "&$top=1" } else { "" };
    let url = format!(
        "{}/_api/web/roledefinitions?$select=Id{top}&$filter=RoleTypeKind eq {role_filter}",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(
            entry("sp.sharing.role_definition_id"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let payload = unwrap_odata(&bytes)?;
    let defs: Vec<IdOnly> = serde_json::from_slice(&payload)?;
    defs.into_iter()
        .next()
        .map(|d| d.id)
        .ok_or_else(|| {
            Error::Surface(
                Surface::SpRest,
                "Could not locate associated role definition for supplied role. Edit and View are supported".into(),
            )
        })
}

/// Looks up the id of the associated member/visitor group for a `RoleType`,
/// mirroring PnPjs `getRoleValue`'s group branch
/// (`packages/sp/sharing/funcs.ts:234-246`).
async fn lookup_associated_group_id(
    client: &Client<'_>,
    site_url: &str,
    group: RoleType,
) -> Result<i64> {
    let endpoint = match group {
        RoleType::Contributor => "associatedmembergroup",
        RoleType::Reader | RoleType::Guest => "associatedvisitorgroup",
        other => {
            return Err(Error::Surface(
                Surface::SpRest,
                format!(
                    "Could not determine role value for supplied value {other:?}. Contributor, Reader, and Guest are supported"
                ),
            ))
        }
    };
    let url = format!(
        "{}/_api/web/{}?$select=Id",
        site_url.trim_end_matches('/'),
        endpoint
    );
    let bytes = client
        .run_ladder(
            entry("sp.sharing.associated_group_id"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let payload = unwrap_odata(&bytes)?;
    let group: IdOnly = serde_json::from_slice(&payload)?;
    Ok(group.id)
}

// ---------------------------------------------------------------------------
// Body builders

/// Serializes a list of login names into the people picker JSON string used
/// by the share endpoints. Mirrors `jsS(loginNames.map(login => ({ Key: login })))`
/// (`packages/sp/sharing/funcs.ts:185-187`).
fn people_picker_input(login_names: &[String]) -> Result<String> {
    let entries: Vec<Value> = login_names
        .iter()
        .map(|k| json!({ "Key": k.as_str() }))
        .collect();
    Ok(serde_json::to_string(&entries)?)
}

/// POST body for an instance-level `shareObject` call, emitted in the exact
/// field order of PnPjs `shareWith` (`packages/sp/sharing/funcs.ts:196-212`).
fn share_with_body(
    login_names: &[String],
    role_value: &str,
    require_signin: bool,
    propagate_acl: bool,
    email_data: Option<&SharingEmailData>,
) -> Result<Vec<u8>> {
    let body = match email_data {
        Some(email) => json!({
            "includeAnonymousLinkInEmail": require_signin,
            "peoplePickerInput": people_picker_input(login_names)?,
            "propagateAcl": propagate_acl,
            "roleValue": role_value,
            "useSimplifiedRoles": true,
            "emailBody": email.body,
            "emailSubject": email.subject.as_deref().unwrap_or(""),
            "sendEmail": true,
        }),
        None => json!({
            "includeAnonymousLinkInEmail": require_signin,
            "peoplePickerInput": people_picker_input(login_names)?,
            "propagateAcl": propagate_acl,
            "roleValue": role_value,
            "useSimplifiedRoles": true,
        }),
    };
    Ok(serde_json::to_vec(&body)?)
}

/// POST body for the static `SP.Web.ShareObject` call, emitted in the exact
/// field order of PnPjs `shareObject` (`packages/sp/sharing/funcs.ts:56-69`).
/// Note `includeAnonymousLinkInEmail`, `propagateAcl` and
/// `useSimplifiedRoles` are accepted by the method signature but never sent.
fn share_object_body(
    login_names: &[String],
    role_value: &str,
    share_url: &str,
    email_data: Option<&SharingEmailData>,
) -> Result<Vec<u8>> {
    let body = match email_data {
        Some(email) => json!({
            "emailBody": email.body,
            "emailSubject": email.subject.as_deref().unwrap_or("Shared with you."),
            "sendEmail": true,
            "peoplePickerInput": people_picker_input(login_names)?,
            "roleValue": role_value,
            "url": share_url,
        }),
        None => json!({
            "peoplePickerInput": people_picker_input(login_names)?,
            "roleValue": role_value,
            "url": share_url,
        }),
    };
    Ok(serde_json::to_vec(&body)?)
}

// ---------------------------------------------------------------------------
// Shared funcs — internal implementations used by item / file / folder

/// Instance-level `shareWith` flow. Role definitions are looked up with
/// `sp.sharing.role_definition_id`, then `shareObject` is POSTed to
/// `{base}/shareObject` (`packages/sp/sharing/funcs.ts:172-215`).
#[allow(clippy::too_many_arguments)]
async fn share_with_internal(
    client: &Client<'_>,
    site_url: &str,
    base: &str,
    op_id: &'static str,
    login_names: &[String],
    role: SharingRole,
    require_signin: bool,
    email_data: Option<&SharingEmailData>,
) -> Result<SharingResult> {
    let role_id = lookup_role_definition_id(client, site_url, role, false).await?;
    let body = share_with_body(
        login_names,
        &format!("role: {role_id}"),
        require_signin,
        false,
        email_data,
    )?;
    let url = format!("{}/shareObject", base.trim_end_matches('/'));
    let bytes = client
        .run_ladder(entry(op_id), "POST", &url, &[], Some(&body))
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// `getShareLink` (`packages/sp/sharing/funcs.ts:80-97`).
async fn get_share_link_internal(
    client: &Client<'_>,
    base: &str,
    op_id: &'static str,
    kind: SharingLinkKind,
    expiration: Option<String>,
) -> Result<ShareLinkResponse> {
    let url = format!("{}/shareLink", base.trim_end_matches('/'));
    let body = json!({
        "request": {
            "createLink": true,
            "emailData": null,
            "settings": {
                "expiration": expiration,
                "linkKind": kind,
            },
        },
    });
    let payload = serde_json::to_vec(&body)?;
    let bytes = client
        .run_ladder(entry(op_id), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// `checkPermissions` (`packages/sp/sharing/funcs.ts:104-108`).
async fn check_permissions_internal(
    client: &Client<'_>,
    base: &str,
    op_id: &'static str,
    recipients: &[SharingRecipient],
) -> Result<Vec<SharingEntityPermission>> {
    let url = format!("{}/checkPermissions", base.trim_end_matches('/'));
    let body = json!({ "recipients": recipients });
    let payload = serde_json::to_vec(&body)?;
    let bytes = client
        .run_ladder(entry(op_id), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// `getSharingInformation`, with `$select`/`$expand` on the method URL
/// (`packages/sp/sharing/funcs.ts:117-121`).
async fn get_sharing_information_internal(
    client: &Client<'_>,
    base: &str,
    op_id: &'static str,
    request: Option<&SharingInformationRequest>,
    expands: &[String],
    selects: &[String],
) -> Result<SharingInformation> {
    let selects: String = if selects.is_empty() {
        "*".into()
    } else {
        selects.join(",")
    };
    let mut url = format!(
        "{}/getSharingInformation?$select={}",
        base.trim_end_matches('/'),
        selects
    );
    if !expands.is_empty() {
        url.push_str(&format!("&$expand={}", expands.join(",")));
    }
    let body = json!({ "request": request });
    let payload = serde_json::to_vec(&body)?;
    let bytes = client
        .run_ladder(entry(op_id), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// `getObjectSharingSettings` (`packages/sp/sharing/funcs.ts:128-132`).
async fn get_object_sharing_settings_internal(
    client: &Client<'_>,
    base: &str,
    op_id: &'static str,
    use_simplified_roles: bool,
) -> Result<ObjectSharingSettings> {
    let url = format!("{}/getObjectSharingSettings", base.trim_end_matches('/'));
    let body = json!({ "useSimplifiedRoles": use_simplified_roles });
    let payload = serde_json::to_vec(&body)?;
    let bytes = client
        .run_ladder(entry(op_id), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// `unshare` (`packages/sp/sharing/funcs.ts:137-140`). PnPjs sends no body.
async fn unshare_internal(
    client: &Client<'_>,
    base: &str,
    op_id: &'static str,
) -> Result<SharingResult> {
    let url = format!("{}/unshareObject", base.trim_end_matches('/'));
    let bytes = client
        .run_ladder(entry(op_id), "POST", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// `deleteLinkByKind` (`packages/sp/sharing/funcs.ts:147-150`).
async fn delete_link_by_kind_internal(
    client: &Client<'_>,
    base: &str,
    op_id: &'static str,
    kind: SharingLinkKind,
) -> Result<()> {
    let url = format!("{}/deleteLinkByKind", base.trim_end_matches('/'));
    let body = json!({ "linkKind": kind });
    let payload = serde_json::to_vec(&body)?;
    client
        .run_ladder(entry(op_id), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(())
}

/// `unshareLink` (`packages/sp/sharing/funcs.ts:158-161`). An empty
/// `share_id` resolves to the empty guid default like PnPjs `emptyGuid`.
async fn unshare_link_internal(
    client: &Client<'_>,
    base: &str,
    op_id: &'static str,
    kind: SharingLinkKind,
    share_id: &str,
) -> Result<()> {
    let url = format!("{}/unshareLink", base.trim_end_matches('/'));
    let body = json!({
        "linkKind": kind,
        "shareId": if share_id.is_empty() { EMPTY_GUID } else { share_id },
    });
    let payload = serde_json::to_vec(&body)?;
    client
        .run_ladder(entry(op_id), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Item sharing

/// Shares a list item with one or more users.
///
/// Ported from PnPjs `_Item.shareWith` + shared funcs
/// (`packages/sp/sharing/item.ts`, `funcs.ts`) @ pnpjs `8ee2375d`.
#[allow(clippy::too_many_arguments)]
pub async fn item_share_with(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    login_names: &[String],
    role: SharingRole,
    require_signin: bool,
    email_data: Option<&SharingEmailData>,
) -> Result<SharingResult> {
    let base = item_url(site_url, list_title, item_id);
    share_with_internal(
        client,
        site_url,
        &base,
        "sp.sharing.item_share_with",
        login_names,
        role,
        require_signin,
        email_data,
    )
    .await
}

/// Gets a sharing link for a list item.
///
/// Ported from PnPjs `getShareLink` (@ pnpjs `8ee2375d`).
pub async fn item_get_share_link(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    kind: SharingLinkKind,
    expiration: Option<String>,
) -> Result<ShareLinkResponse> {
    let base = item_url(site_url, list_title, item_id);
    get_share_link_internal(
        client,
        &base,
        "sp.sharing.item_share_link",
        kind,
        expiration,
    )
    .await
}

/// Checks what permissions a list of recipients has on a list item.
///
/// Ported from PnPjs `checkSharingPermissions` (@ pnpjs `8ee2375d`).
pub async fn item_check_permissions(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    recipients: &[SharingRecipient],
) -> Result<Vec<SharingEntityPermission>> {
    let base = item_url(site_url, list_title, item_id);
    check_permissions_internal(
        client,
        &base,
        "sp.sharing.item_check_permissions",
        recipients,
    )
    .await
}

/// Gets sharing information for a list item.
///
/// Ported from PnPjs `getSharingInformation` (@ pnpjs `8ee2375d`).
pub async fn item_get_sharing_information(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    request: Option<&SharingInformationRequest>,
    expands: &[String],
    selects: &[String],
) -> Result<SharingInformation> {
    let base = item_url(site_url, list_title, item_id);
    get_sharing_information_internal(
        client,
        &base,
        "sp.sharing.item_sharing_information",
        request,
        expands,
        selects,
    )
    .await
}

/// Gets the object sharing settings for a list item.
///
/// Ported from PnPjs `getObjectSharingSettings` (@ pnpjs `8ee2375d`).
pub async fn item_get_object_sharing_settings(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    use_simplified_roles: bool,
) -> Result<ObjectSharingSettings> {
    let base = item_url(site_url, list_title, item_id);
    get_object_sharing_settings_internal(
        client,
        &base,
        "sp.sharing.item_object_sharing_settings",
        use_simplified_roles,
    )
    .await
}

/// Unshares a list item.
///
/// Ported from PnPjs `unshare` (@ pnpjs `8ee2375d`).
pub async fn item_unshare(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
) -> Result<SharingResult> {
    let base = item_url(site_url, list_title, item_id);
    unshare_internal(client, &base, "sp.sharing.item_unshare").await
}

/// Deletes a sharing link of the given kind on a list item.
///
/// Ported from PnPjs `deleteSharingLinkByKind` (@ pnpjs `8ee2375d`).
pub async fn item_delete_sharing_link_by_kind(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    kind: SharingLinkKind,
) -> Result<()> {
    let base = item_url(site_url, list_title, item_id);
    delete_link_by_kind_internal(client, &base, "sp.sharing.item_delete_link_by_kind", kind).await
}

/// Removes a sharing link from a list item.
///
/// Ported from PnPjs `unshareLink` (@ pnpjs `8ee2375d`).
pub async fn item_unshare_link(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    item_id: i64,
    kind: SharingLinkKind,
    share_id: &str,
) -> Result<()> {
    let base = item_url(site_url, list_title, item_id);
    unshare_link_internal(
        client,
        &base,
        "sp.sharing.item_unshare_link",
        kind,
        share_id,
    )
    .await
}

// ---------------------------------------------------------------------------
// File sharing

/// Shares a file with one or more users. The file's associated list item is
/// addressed directly through the `listItemAllFields` navigation, matching
/// the endpoint PnPjs builds after `_File` resolves `getItem()`
/// (`packages/sp/sharing/file.ts`) @ pnpjs `8ee2375d`.
pub async fn file_share_with(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    login_names: &[String],
    role: SharingRole,
    require_signin: bool,
    email_data: Option<&SharingEmailData>,
) -> Result<SharingResult> {
    let base = file_url(site_url, server_relative_path);
    share_with_internal(
        client,
        site_url,
        &base,
        "sp.sharing.file_share_with",
        login_names,
        role,
        require_signin,
        email_data,
    )
    .await
}

/// Gets a sharing link for a file.
///
/// Ported from PnPjs `getShareLink` (@ pnpjs `8ee2375d`).
pub async fn file_get_share_link(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    kind: SharingLinkKind,
    expiration: Option<String>,
) -> Result<ShareLinkResponse> {
    let base = file_url(site_url, server_relative_path);
    get_share_link_internal(
        client,
        &base,
        "sp.sharing.file_share_link",
        kind,
        expiration,
    )
    .await
}

/// Checks what permissions a list of recipients has on a file.
///
/// Ported from PnPjs `checkSharingPermissions` (@ pnpjs `8ee2375d`).
pub async fn file_check_permissions(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    recipients: &[SharingRecipient],
) -> Result<Vec<SharingEntityPermission>> {
    let base = file_url(site_url, server_relative_path);
    check_permissions_internal(
        client,
        &base,
        "sp.sharing.file_check_permissions",
        recipients,
    )
    .await
}

/// Gets sharing information for a file.
///
/// Ported from PnPjs `getSharingInformation` (@ pnpjs `8ee2375d`).
pub async fn file_get_sharing_information(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    request: Option<&SharingInformationRequest>,
    expands: &[String],
    selects: &[String],
) -> Result<SharingInformation> {
    let base = file_url(site_url, server_relative_path);
    get_sharing_information_internal(
        client,
        &base,
        "sp.sharing.file_sharing_information",
        request,
        expands,
        selects,
    )
    .await
}

/// Gets the object sharing settings for a file.
///
/// Ported from PnPjs `getObjectSharingSettings` (@ pnpjs `8ee2375d`).
pub async fn file_get_object_sharing_settings(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    use_simplified_roles: bool,
) -> Result<ObjectSharingSettings> {
    let base = file_url(site_url, server_relative_path);
    get_object_sharing_settings_internal(
        client,
        &base,
        "sp.sharing.file_object_sharing_settings",
        use_simplified_roles,
    )
    .await
}

/// Unshares a file.
///
/// Ported from PnPjs `unshare` (@ pnpjs `8ee2375d`).
pub async fn file_unshare(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> Result<SharingResult> {
    let base = file_url(site_url, server_relative_path);
    unshare_internal(client, &base, "sp.sharing.file_unshare").await
}

/// Deletes a sharing link of the given kind on a file.
///
/// Ported from PnPjs `deleteSharingLinkByKind` (@ pnpjs `8ee2375d`).
pub async fn file_delete_sharing_link_by_kind(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    kind: SharingLinkKind,
) -> Result<()> {
    let base = file_url(site_url, server_relative_path);
    delete_link_by_kind_internal(client, &base, "sp.sharing.file_delete_link_by_kind", kind).await
}

/// Removes a sharing link from a file.
///
/// Ported from PnPjs `unshareLink` (@ pnpjs `8ee2375d`).
pub async fn file_unshare_link(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    kind: SharingLinkKind,
    share_id: &str,
) -> Result<()> {
    let base = file_url(site_url, server_relative_path);
    unshare_link_internal(
        client,
        &base,
        "sp.sharing.file_unshare_link",
        kind,
        share_id,
    )
    .await
}

// ---------------------------------------------------------------------------
// Folder sharing

/// Shares a folder with one or more users, addressed through the
/// `listItemAllFields` navigation of the folder's associated list item
/// (`packages/sp/sharing/folder.ts`) @ pnpjs `8ee2375d`.
pub async fn folder_share_with(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    login_names: &[String],
    role: SharingRole,
    require_signin: bool,
    email_data: Option<&SharingEmailData>,
) -> Result<SharingResult> {
    let base = folder_url(site_url, server_relative_path);
    share_with_internal(
        client,
        site_url,
        &base,
        "sp.sharing.folder_share_with",
        login_names,
        role,
        require_signin,
        email_data,
    )
    .await
}

/// Gets a sharing link for a folder.
///
/// Ported from PnPjs `getShareLink` (@ pnpjs `8ee2375d`).
pub async fn folder_get_share_link(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    kind: SharingLinkKind,
    expiration: Option<String>,
) -> Result<ShareLinkResponse> {
    let base = folder_url(site_url, server_relative_path);
    get_share_link_internal(
        client,
        &base,
        "sp.sharing.folder_share_link",
        kind,
        expiration,
    )
    .await
}

/// Checks what permissions a list of recipients has on a folder.
///
/// Ported from PnPjs `checkSharingPermissions` (@ pnpjs `8ee2375d`).
pub async fn folder_check_permissions(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    recipients: &[SharingRecipient],
) -> Result<Vec<SharingEntityPermission>> {
    let base = folder_url(site_url, server_relative_path);
    check_permissions_internal(
        client,
        &base,
        "sp.sharing.folder_check_permissions",
        recipients,
    )
    .await
}

/// Gets sharing information for a folder.
///
/// Ported from PnPjs `getSharingInformation` (@ pnpjs `8ee2375d`).
pub async fn folder_get_sharing_information(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    request: Option<&SharingInformationRequest>,
    expands: &[String],
    selects: &[String],
) -> Result<SharingInformation> {
    let base = folder_url(site_url, server_relative_path);
    get_sharing_information_internal(
        client,
        &base,
        "sp.sharing.folder_sharing_information",
        request,
        expands,
        selects,
    )
    .await
}

/// Gets the object sharing settings for a folder.
///
/// Ported from PnPjs `getObjectSharingSettings` (@ pnpjs `8ee2375d`).
pub async fn folder_get_object_sharing_settings(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    use_simplified_roles: bool,
) -> Result<ObjectSharingSettings> {
    let base = folder_url(site_url, server_relative_path);
    get_object_sharing_settings_internal(
        client,
        &base,
        "sp.sharing.folder_object_sharing_settings",
        use_simplified_roles,
    )
    .await
}

/// Unshares a folder.
///
/// Ported from PnPjs `unshare` (@ pnpjs `8ee2375d`).
pub async fn folder_unshare(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> Result<SharingResult> {
    let base = folder_url(site_url, server_relative_path);
    unshare_internal(client, &base, "sp.sharing.folder_unshare").await
}

/// Deletes a sharing link of the given kind on a folder.
///
/// Ported from PnPjs `deleteSharingLinkByKind` (@ pnpjs `8ee2375d`).
pub async fn folder_delete_sharing_link_by_kind(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    kind: SharingLinkKind,
) -> Result<()> {
    let base = folder_url(site_url, server_relative_path);
    delete_link_by_kind_internal(client, &base, "sp.sharing.folder_delete_link_by_kind", kind).await
}

/// Removes a sharing link from a folder.
///
/// Ported from PnPjs `unshareLink` (@ pnpjs `8ee2375d`).
pub async fn folder_unshare_link(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
    kind: SharingLinkKind,
    share_id: &str,
) -> Result<()> {
    let base = folder_url(site_url, server_relative_path);
    unshare_link_internal(
        client,
        &base,
        "sp.sharing.folder_unshare_link",
        kind,
        share_id,
    )
    .await
}

// ---------------------------------------------------------------------------
// Web sharing

/// Shares this web with the supplied users by sharing the web's
/// `_layouts/15/aclinv.aspx?forSharing=1&mbypass=1` page.
///
/// Ported from PnPjs `_Web.shareWith` (`packages/sp/sharing/web.ts:47-56`)
/// @ pnpjs `8ee2375d`.
pub async fn web_share_with(
    client: &Client<'_>,
    site_url: &str,
    login_names: &[String],
    role: SharingRole,
    email_data: Option<&SharingEmailData>,
) -> Result<SharingResult> {
    let share_url = format!(
        "{}/_layouts/15/aclinv.aspx?forSharing=1&mbypass=1",
        site_url.trim_end_matches('/')
    );
    web_share_object(
        client,
        site_url,
        &share_url,
        login_names,
        role,
        None,
        email_data,
    )
    .await
}

/// Shares an arbitrary URL within the web using the static
/// `SP.Web.ShareObject` method.
///
/// `group`, when supplied (`RoleType::Contributor`, `Reader` or `Guest`),
/// resolves the associated member/visitor group id for the role value.
/// `includeAnonymousLinkInEmail`, `propagateAcl` and `useSimplifiedRoles`
/// are accepted by PnPjs's method signature but are not transmitted.
///
/// Ported from PnPjs `_Web.shareObject` (`packages/sp/sharing/web.ts:70-91`,
/// `funcs.ts:30-72`) @ pnpjs `8ee2375d`.
pub async fn web_share_object(
    client: &Client<'_>,
    site_url: &str,
    share_url: &str,
    login_names: &[String],
    role: SharingRole,
    group: Option<RoleType>,
    email_data: Option<&SharingEmailData>,
) -> Result<SharingResult> {
    let role_value = if let Some(group_type) = group {
        let id = lookup_associated_group_id(client, site_url, group_type).await?;
        format!("group: {id}")
    } else {
        let id = lookup_role_definition_id(client, site_url, role, true).await?;
        format!("role: {id}")
    };
    let body = share_object_body(login_names, &role_value, share_url, email_data)?;
    let url = share_object_static_url(site_url);
    let bytes = client
        .run_ladder(
            entry("sp.sharing.web_share_object"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Sends the supplied options straight to the static `SP.Web.ShareObject`
/// method without any preprocessing (`bypass = true`).
///
/// Ported from PnPjs `_Web.shareObjectRaw` (`packages/sp/sharing/web.ts:98-100`,
/// `funcs.ts:32-36`) @ pnpjs `8ee2375d`.
pub async fn web_share_object_raw(
    client: &Client<'_>,
    site_url: &str,
    options: &ShareObjectOptions,
) -> Result<SharingResult> {
    let body = serde_json::to_vec(options)?;
    let url = share_object_static_url(site_url);
    let bytes = client
        .run_ladder(
            entry("sp.sharing.web_share_object_raw"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Unshares the object at the supplied url within the web.
///
/// Ported from PnPjs `_Web.unshareObject` (`packages/sp/sharing/web.ts:107-108`)
/// @ pnpjs `8ee2375d`.
pub async fn web_unshare_object(
    client: &Client<'_>,
    site_url: &str,
    url: &str,
) -> Result<SharingResult> {
    let endpoint = format!("{}/_api/web/unshareObject", site_url.trim_end_matches('/'));
    let body = json!({ "url": url });
    let payload = serde_json::to_vec(&body)?;
    let bytes = client
        .run_ladder(
            entry("sp.sharing.web_unshare_object"),
            "POST",
            &endpoint,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
