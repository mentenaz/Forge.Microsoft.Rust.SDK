use forge_m365_auth::{AuthedTransport, TokenClient};
use forge_m365_core::Client;
use forge_m365_sp_navigation::NavigationArea;

fn env(name: &str) -> String {
    std::env::var(name).unwrap_or_else(|_| {
        eprintln!("missing env var {name}");
        std::process::exit(2);
    })
}

fn cert_pem() -> Option<String> {
    if let Ok(p) = std::env::var("M365_CERT_PEM") {
        if !p.is_empty() {
            return Some(p.replace("\\n", "\n").replace("\\r", "\r"));
        }
    }
    std::env::var("M365_CERT_PEM_PATH")
        .ok()
        .map(|path| std::fs::read_to_string(&path).unwrap_or_else(|e| panic!("read {path}: {e}")))
}

#[tokio::main]
async fn main() {
    let tenant_id = env("M365_TENANT_ID");
    let client_id = env("M365_CLIENT_ID");
    let site_url = std::env::var("M365_SITE_URL").unwrap_or_else(|_| {
        eprintln!("missing env var M365_SITE_URL (e.g. https://tenant.sharepoint.com/sites/team)");
        std::process::exit(2);
    });

    let token_client = if let Some(secret) = std::env::var("M365_CLIENT_SECRET")
        .ok()
        .filter(|s| !s.is_empty())
    {
        eprintln!("method: client secret (app-only)");
        TokenClient::with_secret(&tenant_id, &client_id, &secret)
    } else if let Some(pem) = cert_pem() {
        eprintln!("method: client certificate (app-only)");
        TokenClient::with_certificate(&tenant_id, &pem, &client_id).expect("certificate client")
    } else {
        eprintln!("this example needs app-only credentials:");
        eprintln!("  set M365_CLIENT_SECRET or M365_CERT_PEM(_PATH)");
        std::process::exit(2);
    };

    let transport = AuthedTransport::new(token_client);
    let client = Client::new(&transport);

    match forge_m365_sp_navigation::get_navigation_nodes(
        &client,
        &site_url,
        NavigationArea::QuickLaunch,
    )
    .await
    {
        Ok(nodes) => println!("{} quicklaunch node(s):\n{nodes:#?}", nodes.len()),
        Err(e) => {
            eprintln!("get_navigation_nodes(QuickLaunch) FAILED: {e}");
            std::process::exit(1);
        }
    }

    match forge_m365_sp_navigation::get_menu_state(&client, &site_url, None).await {
        Ok(state) => println!(
            "menu state '{}': {} top-level node(s)",
            state.starting_node_title,
            state.nodes.len()
        ),
        Err(e) => {
            eprintln!("get_menu_state FAILED: {e}");
            std::process::exit(1);
        }
    }
}
