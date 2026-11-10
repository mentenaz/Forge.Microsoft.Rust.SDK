use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!(
                "operation '{id}' not registered; is forge-m365-sp-publishing-sitepageservice linked in?"
            )
        })
}

#[pnp_operation(
    id = "sp.publishing_sitepageservice.get_current_user_memberships",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_current_user_memberships_op() {}

/// Gets the current user's unified-group memberships known to the
/// publishing site-page service.
///
/// Ported from PnPjs `_SitePageService.getCurrentUserMemberships`
/// (`packages/sp/publishing-sitepageservice/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_current_user_memberships(
    client: &Client<'_>,
    site_url: &str,
) -> Result<Vec<String>> {
    let url = format!(
        "{}/_api/SP.Publishing.SitePageService.GetCurrentUserMemberships",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(
            entry("sp.publishing_sitepageservice.get_current_user_memberships"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
