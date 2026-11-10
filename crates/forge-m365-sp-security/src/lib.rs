use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;
use serde_json::Value;

/// A role assignment's principal, when `Member` is expanded. Hand-ported
/// from PnPjs's `ISiteUserInfo`/`ISiteGroupInfo` shapes as seen through
/// `RoleAssignment` (`packages/sp/security/types.ts`) — a role assignment's
/// member can be either a user or a group, so only the fields common to both
/// are modeled here.
#[derive(Debug, Clone, Deserialize)]
pub struct RoleAssignmentMember {
    #[serde(rename = "Id")]
    pub id: i64,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "LoginName")]
    pub login_name: String,
    #[serde(rename = "PrincipalType")]
    pub principal_type: i64,
}

/// Subset of SharePoint's `SP.RoleDefinition` fields, hand-ported from
/// PnPjs's `IRoleDefinitionInfo` (`packages/sp/security/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct RoleDefinitionInfo {
    #[serde(rename = "Id")]
    pub id: i64,
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "Description")]
    pub description: String,
    #[serde(rename = "Hidden")]
    pub hidden: bool,
    #[serde(rename = "Order")]
    pub order: i64,
}

/// A role assignment with `Member` and `RoleDefinitionBindings` expanded —
/// PnPjs's own `IRoleAssignmentInfo` is just `{ PrincipalId }`, since PnPjs's
/// chainable query builder expects callers to `.expand()`/`.select()`
/// themselves. This SDK has no query builder yet, so the expansion is baked
/// into `get_role_assignments` directly to make the result useful as-is.
#[derive(Debug, Clone, Deserialize)]
pub struct RoleAssignmentInfo {
    #[serde(rename = "PrincipalId")]
    pub principal_id: i64,
    #[serde(rename = "Member")]
    pub member: Option<RoleAssignmentMember>,
    #[serde(rename = "RoleDefinitionBindings")]
    pub role_definition_bindings: Vec<RoleDefinitionInfo>,
}

/// SharePoint's permission mask, hand-ported from PnPjs's `IBasePermissions`
/// (`packages/sp/security/types.ts`). `low`/`high` are strings because
/// SharePoint REST always emits `Edm.Int64` fields as JSON strings.
#[derive(Debug, Clone, Deserialize)]
pub struct BasePermissions {
    #[serde(rename = "Low")]
    pub low: String,
    #[serde(rename = "High")]
    pub high: String,
}

/// Some collection-typed properties come back wrapped as `{"results": [...]}`,
/// others bare — same defensive check used in `sp-sites`/`sp-search`/`sp-views`.
fn take_results_or_self(value: &mut Value) {
    if let Some(inner) = value.get_mut("results").map(Value::take) {
        *value = inner;
    }
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-security linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.security.get_role_assignments", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_role_assignments_op() {}

#[pnp_operation(id = "sp.security.get_role_definitions", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_role_definitions_op() {}

#[pnp_operation(
    id = "sp.security.break_role_inheritance",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn break_role_inheritance_op() {}

#[pnp_operation(
    id = "sp.security.get_user_effective_permissions",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_user_effective_permissions_op() {}

/// Gets the role assignments (who has what permission level) on a securable
/// object — a web, list, or item. `securable_url` is that object's own REST
/// endpoint, e.g. `{site}/_api/web` or
/// `{site}/_api/web/lists/getbytitle('Tasks')`.
///
/// Ported from PnPjs `RoleAssignments` collection reads
/// (`packages/sp/security/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_role_assignments(
    client: &Client<'_>,
    securable_url: &str,
) -> Result<Vec<RoleAssignmentInfo>> {
    let url = format!(
        "{}/roleassignments?$expand=Member,RoleDefinitionBindings",
        securable_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(
            entry("sp.security.get_role_assignments"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;

    let mut root: Value = serde_json::from_slice(&bytes)?;
    take_results_or_self(&mut root);
    let Value::Array(items) = &mut root else {
        return Ok(Vec::new());
    };
    for item in items.iter_mut() {
        if let Some(bindings) = item.get_mut("RoleDefinitionBindings") {
            take_results_or_self(bindings);
        }
    }
    Ok(serde_json::from_value(root)?)
}

/// Gets the site's available role definitions (permission levels, e.g. "Full
/// Control", "Edit", "Read").
///
/// Ported from PnPjs `Web.roleDefinitions` via `_RoleDefinitions`
/// (`packages/sp/security/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_role_definitions(
    client: &Client<'_>,
    site_url: &str,
) -> Result<Vec<RoleDefinitionInfo>> {
    let url = format!(
        "{}/_api/web/roledefinitions",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(
            entry("sp.security.get_role_definitions"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Stops a securable object (web/list/item) from inheriting permissions from
/// its parent, giving it its own role assignments. Bodiless POST — the
/// parameters are URL-embedded, not a JSON body.
///
/// Ported from PnPjs `ISecurableMethods.breakRoleInheritance`
/// (`packages/sp/security/types.ts`) @ pnpjs `8ee2375d`.
pub async fn break_role_inheritance(
    client: &Client<'_>,
    securable_url: &str,
    copy_role_assignments: bool,
    clear_subscopes: bool,
) -> Result<()> {
    let url = format!(
        "{}/breakroleinheritance(copyroleassignments={copy_role_assignments},clearsubscopes={clear_subscopes})",
        securable_url.trim_end_matches('/')
    );
    client
        .run_ladder(
            entry("sp.security.break_role_inheritance"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

/// Gets a user's effective permission mask on the site, by login name.
///
/// Ported from PnPjs `ISecurableMethods.getUserEffectivePermissions`
/// (`packages/sp/security/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_user_effective_permissions(
    client: &Client<'_>,
    site_url: &str,
    login_name: &str,
) -> Result<BasePermissions> {
    let doubled = format!("'{}'", login_name.replace('\'', "''"));
    let encoded = utf8_percent_encode(&doubled, NON_ALPHANUMERIC);
    let url = format!(
        "{}/_api/web/getUserEffectivePermissions(@user)?@user={encoded}",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(
            entry("sp.security.get_user_effective_permissions"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
