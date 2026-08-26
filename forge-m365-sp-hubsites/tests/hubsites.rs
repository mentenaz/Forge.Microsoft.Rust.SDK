use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_hubsites::{get_hub_site_by_id, get_hub_sites};
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

const HUB_SITE_JSON: &str = r#"{"ID":"99999999-9999-9999-9999-999999999999","Title":"Marketing Hub","SiteId":"aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa","SiteUrl":"https://contoso.sharepoint.com/sites/marketing-hub","Description":"","LogoUrl":"","RequiresJoinApproval":false,"HideNameInNavigation":false}"#;

#[tokio::test]
async fn parses_hub_sites_array() {
    let transport = RecordingTransport::with_response(&format!("[{HUB_SITE_JSON}]"));
    let client = Client::new(&transport);

    let hubs = get_hub_sites(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(hubs.len(), 1);
    assert_eq!(hubs[0].title, "Marketing Hub");
    assert_eq!(hubs[0].id, "99999999-9999-9999-9999-999999999999");
}

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(HUB_SITE_JSON);
    let client = Client::new(&transport);

    get_hub_site_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "99999999-9999-9999-9999-999999999999",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(
        url.ends_with("/_api/hubsites/GetById?hubSiteId='99999999-9999-9999-9999-999999999999'")
    );
}
