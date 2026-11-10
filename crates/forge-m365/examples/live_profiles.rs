use forge_m365_auth::{AuthedTransport, TokenClient};
use forge_m365_core::Client;

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
    let login_name = std::env::var("M365_PROFILE_LOGIN").unwrap_or_else(|_| {
        eprintln!("missing env var M365_PROFILE_LOGIN (e.g. i:0#.f|membership|user@tenant.com)");
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

    match forge_m365_sp_profiles::get_properties_for(&client, &site_url, &login_name).await {
        Ok(props) => println!("properties for {login_name}:\n{props:#}"),
        Err(e) => {
            eprintln!("get_properties_for FAILED: {e}");
            std::process::exit(1);
        }
    }

    match forge_m365_sp_profiles::am_i_following(&client, &site_url, &login_name).await {
        Ok(following) => println!("currently following {login_name}: {following}"),
        Err(e) => {
            eprintln!("am_i_following FAILED: {e}");
            std::process::exit(1);
        }
    }

    match forge_m365_sp_profiles::get_trending_tags(&client, &site_url).await {
        Ok(tags) => println!("{} trending tag(s):\n{tags:#?}", tags.items.len()),
        Err(e) => {
            eprintln!("get_trending_tags FAILED: {e}");
            std::process::exit(1);
        }
    }

    // Write ops mutate the signed-in user's real profile, so those are opt-in.
    if std::env::var("M365_PROFILE_WRITE_TEST").ok().as_deref() != Some("1") {
        eprintln!(
            "\nset M365_PROFILE_WRITE_TEST=1 to also exercise the write ops \
             (set_single_value_profile_property / set_my_profile_pic)"
        );
        return;
    }

    forge_m365_sp_profiles::set_single_value_profile_property(
        &client,
        &site_url,
        &login_name,
        "AboutMe",
        "forge-m365 live example",
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("set_single_value_profile_property FAILED: {e}");
        std::process::exit(1);
    });
    println!("AboutMe updated");
}
