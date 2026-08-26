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
    let list_title = std::env::var("M365_LIST_TITLE").unwrap_or_else(|_| {
        eprintln!("missing env var M365_LIST_TITLE (an existing list on that site, e.g. Tasks)");
        std::process::exit(2);
    });
    let item_id: i64 = std::env::var("M365_ITEM_ID")
        .unwrap_or_else(|_| {
            eprintln!("missing env var M365_ITEM_ID (an existing item id in that list)");
            std::process::exit(2);
        })
        .parse()
        .unwrap_or_else(|_| {
            eprintln!("M365_ITEM_ID must be an integer");
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

    match forge_m365_sp_attachments::get_attachments(&client, &site_url, &list_title, item_id).await
    {
        Ok(attachments) => println!("{} attachment(s):\n{attachments:#?}", attachments.len()),
        Err(e) => {
            eprintln!("get_attachments FAILED: {e}");
            std::process::exit(1);
        }
    }

    // Write round trip is opt-in: it mutates real data, so this example only
    // ever touches an attachment it creates itself.
    if std::env::var("M365_ATTACHMENTS_WRITE_TEST").ok().as_deref() != Some("1") {
        eprintln!("\nset M365_ATTACHMENTS_WRITE_TEST=1 to also round-trip add/download/delete on a throwaway attachment");
        return;
    }

    let file_name = "forge-m365-live-attachments-test.txt";
    let content = b"forge-m365 live_attachments test upload";

    let added = forge_m365_sp_attachments::add_attachment(
        &client,
        &site_url,
        &list_title,
        item_id,
        file_name,
        content,
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("add_attachment FAILED: {e}");
        std::process::exit(1);
    });
    println!("added: {added:#?}");

    let downloaded = forge_m365_sp_attachments::download_attachment(
        &client,
        &site_url,
        &list_title,
        item_id,
        file_name,
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("download_attachment FAILED: {e}");
        std::process::exit(1);
    });
    assert_eq!(downloaded, content, "downloaded content mismatch");
    println!("downloaded content matches ({} bytes)", downloaded.len());

    forge_m365_sp_attachments::delete_attachment(
        &client,
        &site_url,
        &list_title,
        item_id,
        file_name,
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("delete_attachment FAILED: {e}");
        std::process::exit(1);
    });
    println!("deleted {file_name}");
}
