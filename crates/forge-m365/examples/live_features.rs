use forge_m365_auth::{AuthedTransport, TokenClient};
use forge_m365_core::Client;
use forge_m365_sp_features as features;

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
    let site_url = env("M365_SITE_URL");

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

    let site_features = features::get_site_features(&client, &site_url)
        .await
        .unwrap_or_else(|e| {
            eprintln!("get_site_features FAILED: {e}");
            std::process::exit(1);
        });
    println!("{} site feature(s):", site_features.len());
    for f in &site_features {
        println!("  {}", f.definition_id);
    }

    let web_features = features::get_web_features(&client, &site_url)
        .await
        .unwrap_or_else(|e| {
            eprintln!("get_web_features FAILED: {e}");
            std::process::exit(1);
        });
    println!("{} web feature(s):", web_features.len());
    for f in &web_features {
        println!("  {}", f.definition_id);
    }

    // Activation/deactivation mutate real feature state, so they are opt-in.
    if std::env::var("M365_FEATURES_WRITE_TEST").ok().as_deref() != Some("1") {
        eprintln!(
            "\nset M365_FEATURES_WRITE_TEST=1 to activate + deactivate a feature on the web \
             (set M365_FEATURES_GUID=<guid>)"
        );
        return;
    }

    let feature_id = std::env::var("M365_FEATURES_GUID").unwrap_or_else(|_| {
        eprintln!("missing env var M365_FEATURES_GUID");
        std::process::exit(2);
    });

    let activated = features::activate_web_feature(&client, &site_url, &feature_id, false)
        .await
        .unwrap_or_else(|e| {
            eprintln!("activate_web_feature FAILED: {e}");
            std::process::exit(1);
        });
    println!("activated feature {}", activated.definition_id);

    features::deactivate_web_feature(&client, &site_url, &feature_id, false)
        .await
        .unwrap_or_else(|e| {
            eprintln!("deactivate_web_feature FAILED: {e}");
            std::process::exit(1);
        });
    println!("deactivated feature {feature_id}");
}
