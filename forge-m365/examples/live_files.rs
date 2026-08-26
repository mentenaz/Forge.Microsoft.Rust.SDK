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

    // Read-only mode: inspect an existing file, if one was given.
    if let Ok(file_path) = std::env::var("M365_FILE_PATH") {
        match forge_m365_sp_files::get_file_by_path(&client, &site_url, &file_path).await {
            Ok(info) => println!("file: {info:#?}"),
            Err(e) => {
                eprintln!("get_file_by_path FAILED: {e}");
                std::process::exit(1);
            }
        }
        match forge_m365_sp_files::download_file(&client, &site_url, &file_path).await {
            Ok(bytes) => println!("downloaded {} byte(s)", bytes.len()),
            Err(e) => {
                eprintln!("download_file FAILED: {e}");
                std::process::exit(1);
            }
        }
    }

    // Write round trip is opt-in: it mutates real data, so this example only
    // ever touches a file it creates itself, then cleans it up.
    let Ok(folder_path) = std::env::var("M365_FOLDER_PATH") else {
        eprintln!(
            "\nset M365_FOLDER_PATH (a document library server-relative path, e.g. /sites/team/Shared Documents)"
        );
        eprintln!("to round-trip upload/download/delete on a throwaway file");
        return;
    };

    println!("\n--- write round trip (folder: {folder_path}) ---");

    let file_name = "forge-m365-live-files-test.txt";
    let content = b"forge-m365 live_files test upload";

    let uploaded = forge_m365_sp_files::upload_file(
        &client,
        &site_url,
        &folder_path,
        file_name,
        content,
        true,
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("upload_file FAILED: {e}");
        std::process::exit(1);
    });
    println!("uploaded: {uploaded:#?}");

    let downloaded =
        forge_m365_sp_files::download_file(&client, &site_url, &uploaded.server_relative_url)
            .await
            .unwrap_or_else(|e| {
                eprintln!("download_file FAILED: {e}");
                std::process::exit(1);
            });
    assert_eq!(
        downloaded, content,
        "downloaded content did not match uploaded content"
    );
    println!("downloaded content matches ({} bytes)", downloaded.len());

    forge_m365_sp_files::delete_file(&client, &site_url, &uploaded.server_relative_url)
        .await
        .unwrap_or_else(|e| {
            eprintln!("delete_file FAILED: {e}");
            std::process::exit(1);
        });
    println!("deleted {}", uploaded.server_relative_url);
}
