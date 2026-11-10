use forge_m365_auth::{AuthedTransport, TokenClient};
use forge_m365_core::Client;
use forge_m365_sp_regional_settings as regional;

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

    let settings = regional::get_regional_settings(&client, &site_url)
        .await
        .unwrap_or_else(|e| {
            eprintln!("get_regional_settings FAILED: {e}");
            std::process::exit(1);
        });
    println!(
        "regional settings: locale {} date-sep {} time-sep {}",
        settings.locale_id, settings.date_separator, settings.time_separator
    );

    let tz = regional::get_time_zone(&client, &site_url)
        .await
        .unwrap_or_else(|e| {
            eprintln!("get_time_zone FAILED: {e}");
            std::process::exit(1);
        });
    println!("time zone #{:?}", tz.information.as_ref());
    let tz = regional::get_time_zone_by_id(&client, &site_url, 13)
        .await
        .unwrap_or_else(|e| {
            eprintln!("get_time_zone_by_id FAILED: {e}");
            std::process::exit(1);
        });
    println!("time zone 13: {}", tz.description);

    let tzs = regional::get_time_zones(&client, &site_url)
        .await
        .unwrap_or_else(|e| {
            eprintln!("get_time_zones FAILED: {e}");
            std::process::exit(1);
        });
    println!("{} time zones", tzs.len());

    let langs = regional::get_installed_languages(&client, &site_url)
        .await
        .unwrap_or_else(|e| {
            eprintln!("get_installed_languages FAILED: {e}");
            std::process::exit(1);
        });
    println!("{} installed language(s)", langs.len());

    let title = regional::web_title_resource(&client, &site_url, "en-US")
        .await
        .unwrap_or_else(|e| {
            eprintln!("web_title_resource FAILED: {e}");
            std::process::exit(1);
        });
    println!("web title (en-US): {title}");

    println!(
        "UTC now string: {}",
        regional::utc_to_local_time(&client, &site_url, "2027-01-01T12:00:00.000Z")
            .await
            .unwrap_or_else(|e| {
                eprintln!("utc_to_local_time FAILED: {e}");
                std::process::exit(1);
            })
    );
    println!(
        "local now string: {}",
        regional::local_time_to_utc(&client, &site_url, "2027-01-01T12:00:00.000Z")
            .await
            .unwrap_or_else(|e| {
                eprintln!("local_time_to_utc FAILED: {e}");
                std::process::exit(1);
            })
    );
}
