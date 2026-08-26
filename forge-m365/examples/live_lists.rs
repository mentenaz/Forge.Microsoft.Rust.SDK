use forge_m365_auth::{AuthedTransport, TokenClient};
use forge_m365_core::Client;
use serde_json::{Map, Value};

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

    let list = match forge_m365_sp_lists::get_list_by_title(&client, &site_url, &list_title).await {
        Ok(list) => {
            println!("list: {list:#?}");
            list
        }
        Err(e) => {
            eprintln!("get_list_by_title FAILED: {e}");
            std::process::exit(1);
        }
    };

    match forge_m365_sp_lists::get_items(&client, &site_url, &list_title).await {
        Ok(items) => println!("{} item(s):\n{items:#?}", items.len()),
        Err(e) => {
            eprintln!("get_items FAILED: {e}");
            std::process::exit(1);
        }
    }

    // Write ops are opt-in: they mutate real data, so this example only
    // touches an item it creates itself, then cleans it up.
    if std::env::var("M365_LIST_WRITE_TEST").ok().as_deref() != Some("1") {
        eprintln!(
            "\nset M365_LIST_WRITE_TEST=1 to also round-trip add/update/delete on a throwaway item"
        );
        return;
    }

    println!(
        "\n--- write round trip (list entity type: {}) ---",
        list.list_item_entity_type_full_name
    );

    let mut fields = Map::new();
    fields.insert(
        "Title".to_string(),
        Value::String("forge-m365 live_lists test item".to_string()),
    );
    let created = forge_m365_sp_lists::add_item(&client, &site_url, &list_title, fields)
        .await
        .unwrap_or_else(|e| {
            eprintln!("add_item FAILED: {e}");
            std::process::exit(1);
        });
    let item_id = created["Id"].as_i64().expect("created item has an Id");
    println!("added item {item_id}");

    let mut update_fields = Map::new();
    update_fields.insert(
        "Title".to_string(),
        Value::String("forge-m365 live_lists test item (updated)".to_string()),
    );
    forge_m365_sp_lists::update_item(&client, &site_url, &list_title, item_id, update_fields)
        .await
        .unwrap_or_else(|e| {
            eprintln!("update_item FAILED: {e}");
            std::process::exit(1);
        });
    println!("updated item {item_id}");

    forge_m365_sp_lists::delete_item(&client, &site_url, &list_title, item_id)
        .await
        .unwrap_or_else(|e| {
            eprintln!("delete_item FAILED: {e}");
            std::process::exit(1);
        });
    println!("deleted item {item_id}");
}
