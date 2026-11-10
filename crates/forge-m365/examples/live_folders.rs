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
    let folder_path = std::env::var("M365_FOLDER_PATH").unwrap_or_else(|_| {
        eprintln!("missing env var M365_FOLDER_PATH (e.g. /sites/team/Shared Documents)");
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

    match forge_m365_sp_folders::get_folder_by_path(&client, &site_url, &folder_path).await {
        Ok(folder) => println!("folder: {folder:#?}"),
        Err(e) => {
            eprintln!("get_folder_by_path FAILED: {e}");
            std::process::exit(1);
        }
    }

    match forge_m365_sp_folders::get_subfolders(&client, &site_url, &folder_path).await {
        Ok(folders) => println!("{} subfolder(s):\n{folders:#?}", folders.len()),
        Err(e) => {
            eprintln!("get_subfolders FAILED: {e}");
            std::process::exit(1);
        }
    }

    // Write round trip is opt-in: it mutates real data, so this example only
    // ever touches a folder it creates itself, then cleans it up.
    if std::env::var("M365_FOLDERS_WRITE_TEST").ok().as_deref() != Some("1") {
        eprintln!(
            "\nset M365_FOLDERS_WRITE_TEST=1 to also round-trip add/delete on a throwaway folder"
        );
        return;
    }

    let new_folder_path = format!("{folder_path}/forge-m365-live-folders-test");
    let created = forge_m365_sp_folders::add_folder(&client, &site_url, &new_folder_path, true)
        .await
        .unwrap_or_else(|e| {
            eprintln!("add_folder FAILED: {e}");
            std::process::exit(1);
        });
    println!("created: {created:#?}");

    forge_m365_sp_folders::delete_folder(&client, &site_url, &created.server_relative_url)
        .await
        .unwrap_or_else(|e| {
            eprintln!("delete_folder FAILED: {e}");
            std::process::exit(1);
        });
    println!("deleted {}", created.server_relative_url);
}
