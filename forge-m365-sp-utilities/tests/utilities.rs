use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_utilities::{
    create_email_body_for_invitation, expand_groups_to_principals,
    get_current_user_email_addresses, resolve_principal, search_principals, send_email,
    EmailProperties, PrincipalSource, PrincipalType,
};
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
    last_method: Mutex<Option<String>>,
    last_body: Mutex<Option<String>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            ..Default::default()
        }
    }
}

impl Transport for RecordingTransport {
    fn execute(
        &self,
        _surface: Surface,
        method: &str,
        url: &str,
        _headers: &[(&str, &str)],
        body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        *self.last_url.lock().unwrap() = Some(url.to_string());
        *self.last_method.lock().unwrap() = Some(method.to_string());
        *self.last_body.lock().unwrap() = body.map(|b| String::from_utf8_lossy(b).to_string());
        let response = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(response.into_bytes()) })
    }
}

const PRINCIPAL_JSON: &str = r#"{
    "Department": "Engineering",
    "DisplayName": "Alice",
    "Email": "alice@contoso.com",
    "JobTitle": "Engineer",
    "LoginName": "i:0#.f|membership|alice@contoso.com",
    "Mobile": "",
    "PrincipalId": 12,
    "PrincipalType": 1,
    "SIPAddress": ""
}"#;

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

#[tokio::test]
async fn send_email_posts_to_correct_url() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    let props = EmailProperties {
        to: vec!["bob@contoso.com".into()],
        subject: "Hi".into(),
        body: "Hello".into(),
        ..Default::default()
    };
    send_email(&client, SITE, &props).await.unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("_api/SP.Utilities.Utility.SendEmail"));
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"Subject\":\"Hi\""));
}

#[tokio::test]
async fn send_email_remaps_additional_headers() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    let mut headers = std::collections::BTreeMap::new();
    headers.insert("X-Custom".to_string(), "value".to_string());
    let props = EmailProperties {
        to: vec!["bob@contoso.com".into()],
        subject: "Hi".into(),
        body: "Hello".into(),
        additional_headers: Some(headers),
        ..Default::default()
    };
    send_email(&client, SITE, &props).await.unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"Key\":\"X-Custom\""));
    assert!(body.contains("\"ValueType\":\"Edm.String\""));
}

#[tokio::test]
async fn get_current_user_email_addresses_handles_bare_string() {
    let transport = RecordingTransport::with_response("\"alice@contoso.com\"");
    let client = Client::new(&transport);

    let email = get_current_user_email_addresses(&client, SITE)
        .await
        .unwrap();
    assert_eq!(email, "alice@contoso.com");
}

#[tokio::test]
async fn get_current_user_email_addresses_handles_odata_wrapper() {
    let transport = RecordingTransport::with_response(r#"{"value":"alice@contoso.com"}"#);
    let client = Client::new(&transport);

    let email = get_current_user_email_addresses(&client, SITE)
        .await
        .unwrap();
    assert_eq!(email, "alice@contoso.com");
}

#[tokio::test]
async fn resolve_principal_parses_response() {
    let transport = RecordingTransport::with_response(PRINCIPAL_JSON);
    let client = Client::new(&transport);

    let principal = resolve_principal(
        &client,
        SITE,
        "alice@contoso.com",
        PrincipalType::USER,
        PrincipalSource::ALL,
        true,
        false,
        false,
    )
    .await
    .unwrap();

    assert_eq!(principal.display_name, "Alice");
    assert_eq!(principal.principal_id, 12);
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("ResolvePrincipalInCurrentContext"));
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"scopes\":1"));
}

#[tokio::test]
async fn search_principals_builds_request() {
    let transport = RecordingTransport::with_response(&format!("[{PRINCIPAL_JSON}]"));
    let client = Client::new(&transport);

    let results = search_principals(
        &client,
        SITE,
        "alice",
        PrincipalType::ALL,
        PrincipalSource::ALL,
        "",
        10,
    )
    .await
    .unwrap();

    assert_eq!(results.len(), 1);
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("SearchPrincipalsUsingContextWeb"));
}

#[tokio::test]
async fn create_email_body_for_invitation_returns_string() {
    let transport =
        RecordingTransport::with_response(r#"{"value":"https://contoso.sharepoint.com/x"}"#);
    let client = Client::new(&transport);

    let body = create_email_body_for_invitation(&client, SITE, "/sites/team/doc.docx")
        .await
        .unwrap();
    assert_eq!(body, "https://contoso.sharepoint.com/x");
}

#[tokio::test]
async fn expand_groups_to_principals_parses_array() {
    let transport = RecordingTransport::with_response(&format!("[{PRINCIPAL_JSON}]"));
    let client = Client::new(&transport);

    let results = expand_groups_to_principals(&client, SITE, &["Team Owners".to_string()], 30)
        .await
        .unwrap();
    assert_eq!(results.len(), 1);
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("ExpandGroupsToPrincipals"));
}
