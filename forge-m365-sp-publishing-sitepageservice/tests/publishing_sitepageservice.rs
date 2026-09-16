use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_publishing_sitepageservice::get_current_user_memberships;
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
    last_method: Mutex<Option<String>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            last_url: Mutex::new(None),
            last_method: Mutex::new(None),
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
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        *self.last_url.lock().unwrap() = Some(url.to_string());
        *self.last_method.lock().unwrap() = Some(method.to_string());
        let body = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

#[tokio::test]
async fn builds_correct_url_and_parses_response() {
    let transport = RecordingTransport::with_response(r#"["Group A","Group B"]"#);
    let client = Client::new(&transport);

    let memberships =
        get_current_user_memberships(&client, "https://contoso.sharepoint.com/sites/team")
            .await
            .unwrap();

    assert_eq!(memberships, vec!["Group A", "Group B"]);
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "GET"
    );
    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        "https://contoso.sharepoint.com/sites/team/_api/SP.Publishing.SitePageService.GetCurrentUserMemberships"
    );
}

#[tokio::test]
async fn trims_trailing_slash_from_site_url() {
    let transport = RecordingTransport::with_response("[]");
    let client = Client::new(&transport);

    get_current_user_memberships(&client, "https://contoso.sharepoint.com/sites/team/")
        .await
        .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(!url.contains("team//"));
}
