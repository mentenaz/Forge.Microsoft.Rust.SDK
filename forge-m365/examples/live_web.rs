use forge_m365_auth::{AuthedTransport, TokenClient};
use forge_m365_core::{Client, Ladder, OperationEntry, Surface};

fn env(name: &str) -> String {
    std::env::var(name).unwrap_or_else(|_| {
        eprintln!("missing env var {name}");
        std::process::exit(2);
    })
}

const GET_WEB: OperationEntry = OperationEntry {
    id: "sp.web.get",
    operation_path: "live_web::get_web",
    ladder: Ladder {
        primary: Surface::SpRest,
        fallback: &[],
    },
};

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

    let url = format!(
        "{site}/_api/web?$select=Title,ServerRelativeUrl",
        site = site_url.trim_end_matches('/')
    );

    match client.run_ladder(&GET_WEB, "GET", &url, None).await {
        Ok(bytes) => {
            let text = String::from_utf8_lossy(&bytes);
            println!("web title response:");
            println!(
                "{}",
                serde_json::to_string_pretty(
                    &serde_json::from_str::<serde_json::Value>(&text)
                        .unwrap_or(serde_json::Value::String(text.to_string()))
                )
                .unwrap()
            );
        }
        Err(e) => {
            eprintln!("FAILED: {e}");
            eprintln!();
            eprintln!("hints:");
            eprintln!("  - 401/403: app registration needs SharePoint application permission (e.g. Sites.Read.All or Sites.FullControl.All) with admin consent");
            eprintln!("  - 404: check M365_SITE_URL points at an existing site");
            std::process::exit(1);
        }
    }
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
