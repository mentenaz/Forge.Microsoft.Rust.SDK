use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_site_users::{
    get_current_user, get_site_users, get_user_by_email, get_user_by_id, get_user_by_login_name,
};
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            last_url: Mutex::new(None),
        }
    }
}

impl Transport for RecordingTransport {
    fn execute(
        &self,
        _surface: Surface,
        _method: &str,
        url: &str,
        _headers: &[(&str, &str)],
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        *self.last_url.lock().unwrap() = Some(url.to_string());
        let body = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

const USER_JSON: &str = r#"{"Id":7,"Title":"Jane Doe","LoginName":"i:0#.f|membership|jane@contoso.com","Email":"jane@contoso.com","PrincipalType":1,"IsSiteAdmin":false,"IsHiddenInUI":false}"#;

#[tokio::test]
async fn parses_current_user() {
    let transport = RecordingTransport::with_response(USER_JSON);
    let client = Client::new(&transport);

    let user = get_current_user(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(user.title, "Jane Doe");
    assert_eq!(user.id, 7);
}

#[tokio::test]
async fn parses_site_users_array() {
    let transport = RecordingTransport::with_response(&format!("[{USER_JSON}]"));
    let client = Client::new(&transport);

    let users = get_site_users(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(users.len(), 1);
    assert_eq!(users[0].email, "jane@contoso.com");
}

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(USER_JSON);
    let client = Client::new(&transport);

    get_user_by_id(&client, "https://contoso.sharepoint.com/sites/team", 7)
        .await
        .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/_api/web/siteusers/getbyid(7)"));
}

#[tokio::test]
async fn get_by_email_encodes_special_characters() {
    let transport = RecordingTransport::with_response(USER_JSON);
    let client = Client::new(&transport);

    get_user_by_email(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "jane+test@contoso.com",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.contains("getbyemail('jane%2Btest%40contoso%2Ecom')"));
}

#[tokio::test]
async fn get_by_login_name_does_not_encode_the_login_name() {
    let transport = RecordingTransport::with_response(USER_JSON);
    let client = Client::new(&transport);

    get_user_by_login_name(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "i:0#.f|membership|jane@contoso.com",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("siteusers('!@v::i:0#.f|membership|jane@contoso.com')"));
}
