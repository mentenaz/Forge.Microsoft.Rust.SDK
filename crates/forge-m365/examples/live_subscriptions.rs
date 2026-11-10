use forge_m365_auth::{AuthedTransport, TokenClient};
use forge_m365_core::Client;
use forge_m365_sp_subscriptions as subscriptions;

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
    let list_title = env("M365_SUBSCRIPTIONS_LIST");

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

    let subs = subscriptions::list_subscriptions(&client, &site_url, &list_title)
        .await
        .unwrap_or_else(|e| {
            eprintln!("list_subscriptions FAILED: {e}");
            std::process::exit(1);
        });
    println!("{} webhook subscription(s) on '{list_title}':", subs.len());
    for s in &subs {
        println!(
            "  {} -> {} (expires {})",
            s.id, s.notification_url, s.expiration_date_time
        );
    }

    // Add/renew/delete operations mutate real subscription state, so they are
    // opt-in.
    if std::env::var("M365_SUBSCRIPTIONS_WRITE_TEST")
        .ok()
        .as_deref()
        != Some("1")
    {
        eprintln!(
            "\nset M365_SUBSCRIPTIONS_WRITE_TEST=1 to add, renew and delete a webhook \
             (notificationUrl must be reachable for SharePoint's validation ping)"
        );
        return;
    }

    let notification_url = env("M365_SUBSCRIPTIONS_NOTIFY_URL");
    let expires = std::env::var("M365_SUBSCRIPTIONS_EXPIRES")
        .unwrap_or_else(|_| "2027-02-27T00:00:00+00:00".to_string());

    let created = subscriptions::add_subscription(
        &client,
        &site_url,
        &list_title,
        &notification_url,
        &expires,
        Some("forge-m365-example"),
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("add_subscription FAILED: {e}");
        std::process::exit(1);
    });
    println!(
        "added subscription {} for {}",
        created.id, created.notification_url
    );

    let renewed = subscriptions::update_subscription(
        &client,
        &site_url,
        &list_title,
        &created.id,
        Some("2027-08-27T00:00:00+00:00"),
        None,
        None,
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("update_subscription FAILED: {e}");
        std::process::exit(1);
    });
    println!("renewed expiry: {}", renewed.expiration_date_time);

    subscriptions::delete_subscription(&client, &site_url, &list_title, &created.id)
        .await
        .unwrap_or_else(|e| {
            eprintln!("delete_subscription FAILED: {e}");
            std::process::exit(1);
        });
    println!("deleted subscription {}", created.id);
}
