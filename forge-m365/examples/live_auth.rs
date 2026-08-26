use forge_m365_auth::{scope_for, BrowserConfig, TokenClient};
use forge_m365_core::Surface;

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
    if let Ok(path) = std::env::var("M365_CERT_PEM_PATH") {
        return Some(std::fs::read_to_string(&path).unwrap_or_else(|e| panic!("read {path}: {e}")));
    }
    None
}

fn main() {
    let tenant_id = env("M365_TENANT_ID");
    let client_id = env("M365_CLIENT_ID");

    let tokens: Vec<(String, String)> = if let Some(secret) = std::env::var("M365_CLIENT_SECRET")
        .ok()
        .filter(|s| !s.is_empty())
    {
        eprintln!("method: client secret");
        run_app_only(
            TokenClient::with_secret(&tenant_id, &client_id, &secret),
            &tenant_id,
        )
    } else if let Some(pem) = cert_pem() {
        eprintln!("method: client certificate");
        let client = TokenClient::with_certificate(&tenant_id, &pem, &client_id)
            .expect("certificate client");
        run_app_only(client, &tenant_id)
    } else {
        eprintln!("method: interactive browser (delegated)");
        let mut out = Vec::new();
        for (label, surface) in [("Graph", Surface::Graph), ("SharePoint", Surface::SpRest)] {
            let cfg = BrowserConfig::new(&tenant_id, &client_id, scope_for(surface, &tenant_id));
            match cfg.acquire_interactive() {
                Ok(tok) => out.push((
                    label.to_string(),
                    format!("OK ({} chars)", tok.access_token.len()),
                )),
                Err(e) => {
                    eprintln!("{label}: FAILED: {e}");
                    std::process::exit(1);
                }
            }
        }
        out
    };

    for (label, status) in tokens {
        println!("{label}: {status}");
    }
    println!("done");
}

fn run_app_only(client: TokenClient, _tenant_id: &str) -> Vec<(String, String)> {
    let rt = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .unwrap();
    let mut out = Vec::new();
    for (label, surface) in [("Graph", Surface::Graph), ("SharePoint", Surface::SpRest)] {
        match rt.block_on(client.acquire(surface)) {
            Ok(tok) => out.push((
                label.to_string(),
                format!(
                    "OK ({} chars, expires in {}s)",
                    tok.access_token.len(),
                    tok.expires_in.as_secs()
                ),
            )),
            Err(e) => {
                eprintln!("{label}: FAILED: {e}");
                std::process::exit(1);
            }
        }
    }
    out
}
