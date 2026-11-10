use forge_m365_auth::{AuthedTransport, TokenClient};
use forge_m365_core::Client;
use forge_m365_sp_sharing::types::{SharingLinkKind, SharingRecipient, SharingRole};

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
    let list_title = std::env::var("M365_SHARING_LIST").unwrap_or_else(|_| {
        eprintln!("missing env var M365_SHARING_LIST (list title holding the target item)");
        std::process::exit(2);
    });
    let item_id: i64 = std::env::var("M365_SHARING_ITEM_ID")
        .unwrap_or_else(|_| {
            eprintln!("missing env var M365_SHARING_ITEM_ID");
            std::process::exit(2);
        })
        .parse()
        .expect("M365_SHARING_ITEM_ID must be an integer");
    let server_relative_path = std::env::var("M365_SHARING_FILE_PATH").unwrap_or_else(|_| {
        eprintln!(
            "missing env var M365_SHARING_FILE_PATH (server-relative, e.g. {site_url}/Shared Documents/report.docx)"
        );
        std::process::exit(2);
    });
    let login_name = std::env::var("M365_SHARING_LOGIN").unwrap_or_else(|_| {
        eprintln!("missing env var M365_SHARING_LOGIN (e.g. i:0#.f|membership|user@tenant.com)");
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

    match forge_m365_sp_sharing::item_get_object_sharing_settings(
        &client,
        &site_url,
        &list_title,
        item_id,
        true,
    )
    .await
    {
        Ok(settings) => println!(
            "item {item_id} sharing settings:\nweb url: {}\nhas edit role: {}\ncan send email: {}",
            settings.web_url.as_deref().unwrap_or("?"),
            settings.has_edit_role.unwrap_or_default(),
            settings.can_send_email.unwrap_or_default(),
        ),
        Err(e) => {
            eprintln!("item_get_object_sharing_settings FAILED: {e}");
            std::process::exit(1);
        }
    }

    match forge_m365_sp_sharing::item_get_sharing_information(
        &client,
        &site_url,
        &list_title,
        item_id,
        None,
        &[],
        &[],
    )
    .await
    {
        Ok(info) => println!(
            "item {item_id} sharing info:\nunique permissions: {}\nshared object type: {:?}",
            info.has_unique_permissions.unwrap_or_default(),
            info.shared_object_type,
        ),
        Err(e) => {
            eprintln!("item_get_sharing_information FAILED: {e}");
            std::process::exit(1);
        }
    }

    match forge_m365_sp_sharing::item_check_permissions(
        &client,
        &site_url,
        &list_title,
        item_id,
        &[SharingRecipient {
            email: Some(login_name.clone()),
            alias: None,
        }],
    )
    .await
    {
        Ok(perms) => println!(
            "permissions for {login_name}: {}/{}",
            perms.len(),
            perms
                .iter()
                .filter(|p| p.has_access.unwrap_or_default())
                .count()
        ),
        Err(e) => {
            eprintln!("item_check_permissions FAILED: {e}");
            std::process::exit(1);
        }
    }

    match forge_m365_sp_sharing::file_get_object_sharing_settings(
        &client,
        &site_url,
        &server_relative_path,
        true,
    )
    .await
    {
        Ok(settings) => println!(
            "file sharing settings:\nitem name: {}\nshare by email enabled: {}",
            settings.item_name.as_deref().unwrap_or("?"),
            settings.share_by_email_enabled.unwrap_or_default(),
        ),
        Err(e) => {
            eprintln!("file_get_object_sharing_settings FAILED: {e}");
            std::process::exit(1);
        }
    }

    // Share/unshare ops create or remove real sharing state, so those are opt-in.
    if std::env::var("M365_SHARING_WRITE_TEST").ok().as_deref() != Some("1") {
        eprintln!(
            "\nset M365_SHARING_WRITE_TEST=1 to also exercise the write ops \
             (get/share a link, share / unshare the item with {login_name})"
        );
        return;
    }

    let link = forge_m365_sp_sharing::item_get_share_link(
        &client,
        &site_url,
        &list_title,
        item_id,
        SharingLinkKind::OrganizationView,
        None,
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("item_get_share_link FAILED: {e}");
        std::process::exit(1);
    });
    println!(
        "organization view share link: {}",
        link.sharing_link_info
            .as_ref()
            .and_then(|i| i.url.as_deref())
            .unwrap_or("?")
    );

    let share_result = forge_m365_sp_sharing::item_share_with(
        &client,
        &site_url,
        &list_title,
        item_id,
        std::slice::from_ref(&login_name),
        SharingRole::View,
        false,
        None,
    )
    .await
    .unwrap_or_else(|e| {
        eprintln!("item_share_with FAILED: {e}");
        std::process::exit(1);
    });
    println!(
        "share status: {:?} ({} unique){}",
        share_result.status_code,
        share_result
            .uniquely_permissioned_users
            .as_ref()
            .map(Vec::len)
            .unwrap_or(0),
        share_result
            .error_message
            .as_deref()
            .map(|m| format!("; error: {m}"))
            .unwrap_or_default(),
    );

    forge_m365_sp_sharing::item_unshare(&client, &site_url, &list_title, item_id)
        .await
        .unwrap_or_else(|e| {
            eprintln!("item_unshare FAILED: {e}");
            std::process::exit(1);
        });
    println!("item {item_id} unshared");
}
