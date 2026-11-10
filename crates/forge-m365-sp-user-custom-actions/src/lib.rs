use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};

/// Which collection a user custom action belongs to. SharePoint exposes the
/// identical `usercustomactions` shape off both `_api/web` and `_api/site`;
/// the URL differs only by this one path segment.
///
/// Mirrors PnPjs's separate `sp.web.userCustomActions` /
/// `sp.site.userCustomActions` accessors (`packages/sp/user-custom-actions/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Scope {
    Web,
    Site,
}

impl Scope {
    fn segment(self) -> &'static str {
        match self {
            Scope::Web => "web",
            Scope::Site => "site",
        }
    }
}

/// `SP.UserCustomAction` fields, hand-ported from PnPjs's `IUserCustomActionInfo`
/// (`packages/sp/user-custom-actions/types.ts`).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UserCustomActionInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "Location")]
    pub location: String,
    #[serde(rename = "Group")]
    pub group: String,
    #[serde(rename = "Sequence")]
    pub sequence: i64,
    #[serde(rename = "RegistrationId")]
    pub registration_id: String,
    /// `UserCustomActionRegistrationType`: None=0, List=1, ContentType=2,
    /// ProgId=3, FileType=4.
    #[serde(rename = "RegistrationType")]
    pub registration_type: i64,
    /// `UserCustomActionScope`: Unknown=0, Site=2, Web=3, List=4.
    #[serde(rename = "Scope")]
    pub scope: i64,
    #[serde(rename = "ScriptBlock")]
    pub script_block: Option<String>,
    #[serde(rename = "ScriptSrc")]
    pub script_src: Option<String>,
    #[serde(rename = "Url")]
    pub url: Option<String>,
    #[serde(rename = "ImageUrl")]
    pub image_url: Option<String>,
    #[serde(rename = "CommandUIExtension")]
    pub command_ui_extension: Option<String>,
    #[serde(rename = "VersionOfUserCustomAction")]
    pub version_of_user_custom_action: String,
    #[serde(rename = "Rights")]
    pub rights: Option<BasePermissions>,
}

/// `SP.BasePermissions`: a 64-bit permission mask split into two 32-bit
/// halves, hand-ported from PnPjs's `IBasePermissions`
/// (`packages/sp/security/types.ts`). SharePoint transmits both halves as
/// strings, not numbers.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct BasePermissions {
    #[serde(rename = "High")]
    pub high: String,
    #[serde(rename = "Low")]
    pub low: String,
}

fn actions_url(site_url: &str, scope: Scope) -> String {
    format!(
        "{}/_api/{}/usercustomactions",
        site_url.trim_end_matches('/'),
        scope.segment()
    )
}

fn action_url(site_url: &str, scope: Scope, id: &str) -> String {
    format!("{}('{id}')", actions_url(site_url, scope))
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!(
                "operation '{id}' not registered; is forge-m365-sp-user-custom-actions linked in?"
            )
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.user_custom_actions.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_all_op() {}

#[pnp_operation(id = "sp.user_custom_actions.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_id_op() {}

#[pnp_operation(id = "sp.user_custom_actions.add", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_op() {}

#[pnp_operation(id = "sp.user_custom_actions.update", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn update_op() {}

#[pnp_operation(id = "sp.user_custom_actions.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_op() {}

#[pnp_operation(id = "sp.user_custom_actions.clear", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn clear_op() {}

/// Gets all user custom actions in the given scope's collection.
///
/// Ported from PnPjs `_UserCustomActions` collection read
/// (`packages/sp/user-custom-actions/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_all(
    client: &Client<'_>,
    site_url: &str,
    scope: Scope,
) -> Result<Vec<UserCustomActionInfo>> {
    let url = actions_url(site_url, scope);
    let bytes = client
        .run_ladder(
            entry("sp.user_custom_actions.get_all"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a single user custom action by its GUID id.
///
/// Ported from PnPjs `_UserCustomActions.getById`
/// (`packages/sp/user-custom-actions/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_by_id(
    client: &Client<'_>,
    site_url: &str,
    scope: Scope,
    id: &str,
) -> Result<UserCustomActionInfo> {
    let url = action_url(site_url, scope, id);
    let bytes = client
        .run_ladder(
            entry("sp.user_custom_actions.get_by_id"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Creates a user custom action from a property bag (e.g. `Name`, `Location`,
/// `Sequence`, `ScriptBlock`/`ScriptSrc`).
///
/// Ported from PnPjs `_UserCustomActions.add`
/// (`packages/sp/user-custom-actions/types.ts`) @ pnpjs `8ee2375d`.
pub async fn add(
    client: &Client<'_>,
    site_url: &str,
    scope: Scope,
    properties: Map<String, Value>,
) -> Result<UserCustomActionInfo> {
    let url = actions_url(site_url, scope);
    let payload = serde_json::to_vec(&Value::Object(properties))?;
    let bytes = client
        .run_ladder(
            entry("sp.user_custom_actions.add"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Updates a user custom action with the supplied properties (SharePoint
/// MERGE semantics: unlisted fields are left untouched).
///
/// Ported from PnPjs `_UserCustomAction.update`
/// (`packages/sp/user-custom-actions/types.ts`) @ pnpjs `8ee2375d`.
pub async fn update(
    client: &Client<'_>,
    site_url: &str,
    scope: Scope,
    id: &str,
    properties: Map<String, Value>,
) -> Result<()> {
    let url = action_url(site_url, scope, id);
    let payload = serde_json::to_vec(&Value::Object(properties))?;
    client
        .run_ladder(
            entry("sp.user_custom_actions.update"),
            "PATCH",
            &url,
            &[("IF-MATCH", "*")],
            Some(&payload),
        )
        .await?;
    Ok(())
}

/// Deletes a single user custom action.
///
/// Ported from PnPjs `_UserCustomAction.delete` (via PnPjs's generic
/// `deleteable` mixin, `packages/sp/user-custom-actions/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn delete(client: &Client<'_>, site_url: &str, scope: Scope, id: &str) -> Result<()> {
    let url = action_url(site_url, scope, id);
    client
        .run_ladder(
            entry("sp.user_custom_actions.delete"),
            "DELETE",
            &url,
            &[("IF-MATCH", "*")],
            None,
        )
        .await?;
    Ok(())
}

/// Deletes every user custom action in the given scope's collection.
/// Bodiless POST.
///
/// Ported from PnPjs `_UserCustomActions.clear`
/// (`packages/sp/user-custom-actions/types.ts`) @ pnpjs `8ee2375d`.
pub async fn clear(client: &Client<'_>, site_url: &str, scope: Scope) -> Result<()> {
    let url = format!("{}/clear", actions_url(site_url, scope));
    client
        .run_ladder(
            entry("sp.user_custom_actions.clear"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}
