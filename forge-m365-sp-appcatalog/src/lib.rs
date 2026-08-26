use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde_json::{Map, Value};

/// PnPjs's own `_AppCatalog`/`_App` classes are untyped (`_SPCollection`/
/// `_SPInstance` with no generic type param) — SharePoint's `SP.AppCatalog`
/// available-apps entity isn't modeled anywhere in the reference source, so
/// rather than invent a struct with unverified field names, this crate
/// returns raw JSON, same as `sp-lists`'s `get_items` does for the same
/// reason.
pub type AppInfo = Map<String, Value>;

fn app_catalog_url(site_url: &str) -> String {
    format!(
        "{}/_api/web/tenantappcatalog/AvailableApps",
        site_url.trim_end_matches('/')
    )
}

fn app_url(site_url: &str, app_id: &str) -> String {
    format!("{}/getById('{app_id}')", app_catalog_url(site_url))
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-appcatalog linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.appcatalog.get_all", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_available_apps_op() {}

#[pnp_operation(id = "sp.appcatalog.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_app_by_id_op() {}

#[pnp_operation(id = "sp.appcatalog.deploy", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn deploy_app_op() {}

#[pnp_operation(id = "sp.appcatalog.retract", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn retract_app_op() {}

#[pnp_operation(id = "sp.appcatalog.install", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn install_app_op() {}

#[pnp_operation(id = "sp.appcatalog.uninstall", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn uninstall_app_op() {}

#[pnp_operation(id = "sp.appcatalog.upgrade", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn upgrade_app_op() {}

#[pnp_operation(id = "sp.appcatalog.remove", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn remove_app_op() {}

/// Gets the apps available in the (tenant or site collection) app catalog
/// for this context.
///
/// Ported from PnPjs `AppCatalog` collection reads
/// (`packages/sp/appcatalog/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_available_apps(client: &Client<'_>, site_url: &str) -> Result<Vec<AppInfo>> {
    let url = app_catalog_url(site_url);
    let bytes = client
        .run_ladder(entry("sp.appcatalog.get_all"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a single app's details by id.
///
/// Ported from PnPjs `AppCatalog.getAppById`
/// (`packages/sp/appcatalog/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_app_by_id(client: &Client<'_>, site_url: &str, app_id: &str) -> Result<AppInfo> {
    let url = app_url(site_url, app_id);
    let bytes = client
        .run_ladder(entry("sp.appcatalog.get_by_id"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Deploys an app from the catalog. Must be called against the tenant app
/// catalog web, not an arbitrary site. Bodiless POST.
///
/// Ported from PnPjs `App.deploy` (`packages/sp/appcatalog/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn deploy_app(
    client: &Client<'_>,
    site_url: &str,
    app_id: &str,
    skip_feature_deployment: bool,
) -> Result<()> {
    let url = format!(
        "{}/Deploy({skip_feature_deployment})",
        app_url(site_url, app_id)
    );
    client
        .run_ladder(entry("sp.appcatalog.deploy"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Retracts a deployed app. Must be called against the tenant app catalog
/// web. Bodiless POST.
///
/// Ported from PnPjs `App.retract` (`packages/sp/appcatalog/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn retract_app(client: &Client<'_>, site_url: &str, app_id: &str) -> Result<()> {
    let url = format!("{}/Retract", app_url(site_url, app_id));
    client
        .run_ladder(entry("sp.appcatalog.retract"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Installs an already-deployed app on the current web. Bodiless POST.
///
/// Ported from PnPjs `App.install` (`packages/sp/appcatalog/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn install_app(client: &Client<'_>, site_url: &str, app_id: &str) -> Result<()> {
    let url = format!("{}/Install", app_url(site_url, app_id));
    client
        .run_ladder(entry("sp.appcatalog.install"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Uninstalls an installed app from the current web. Note: unlike files,
/// SharePoint does not send uninstalled solution packages to the recycle
/// bin. Bodiless POST.
///
/// Ported from PnPjs `App.uninstall` (`packages/sp/appcatalog/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn uninstall_app(client: &Client<'_>, site_url: &str, app_id: &str) -> Result<()> {
    let url = format!("{}/Uninstall", app_url(site_url, app_id));
    client
        .run_ladder(entry("sp.appcatalog.uninstall"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Upgrades an installed app on the current web. Bodiless POST.
///
/// Ported from PnPjs `App.upgrade` (`packages/sp/appcatalog/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn upgrade_app(client: &Client<'_>, site_url: &str, app_id: &str) -> Result<()> {
    let url = format!("{}/Upgrade", app_url(site_url, app_id));
    client
        .run_ladder(entry("sp.appcatalog.upgrade"), "POST", &url, &[], None)
        .await?;
    Ok(())
}

/// Removes an app from the catalog entirely. Must be called against the
/// tenant app catalog web. Bodiless POST.
///
/// Ported from PnPjs `App.remove` (`packages/sp/appcatalog/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn remove_app(client: &Client<'_>, site_url: &str, app_id: &str) -> Result<()> {
    let url = format!("{}/Remove", app_url(site_url, app_id));
    client
        .run_ladder(entry("sp.appcatalog.remove"), "POST", &url, &[], None)
        .await?;
    Ok(())
}
