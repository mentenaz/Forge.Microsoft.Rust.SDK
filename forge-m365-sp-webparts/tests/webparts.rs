use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_webparts::{
    close_web_part, delete_web_part, export_web_part, get_web_part_by_control_id,
    get_web_part_by_id, get_web_parts, import_web_part, move_web_part_to, open_web_part,
    save_web_part_changes, WebPartsPersonalizationScope,
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
        let resp = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(resp.into_bytes()) })
    }
}

const SITE: &str = "https://contoso.sharepoint.com/sites/team";
const PAGE: &str = "/sites/team/SitePages/home.aspx";
const WP_JSON: &str = r#"{"ID":"dddddddd-dddd-dddd-dddd-dddddddddddd","Title":"My Web Part"}"#;

#[tokio::test]
async fn get_web_parts_builds_correct_url() {
    let transport = RecordingTransport::with_response(&format!("[{WP_JSON}]"));
    let client = Client::new(&transport);

    let parts = get_web_parts(&client, SITE, PAGE, WebPartsPersonalizationScope::Shared)
        .await
        .unwrap();

    assert_eq!(parts.len(), 1);
    assert_eq!(parts[0]["Title"], "My Web Part");
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.contains("getLimitedWebPartManager(scope=1)"));
    assert!(url.ends_with("/webparts"));
}

#[tokio::test]
async fn get_web_part_by_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(WP_JSON);
    let client = Client::new(&transport);

    get_web_part_by_id(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::User,
        "dddddddd-dddd-dddd-dddd-dddddddddddd",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.contains("getLimitedWebPartManager(scope=0)"));
    assert!(url.ends_with("/webparts/getbyid('dddddddd-dddd-dddd-dddd-dddddddddddd')"));
}

#[tokio::test]
async fn get_web_part_by_control_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(WP_JSON);
    let client = Client::new(&transport);

    get_web_part_by_control_id(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::Shared,
        "g_1",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/webparts/getByControlId('g_1')"));
}

#[tokio::test]
async fn export_sends_correct_body_and_returns_string() {
    let transport = RecordingTransport::with_response(r#""<webPart>xml</webPart>""#);
    let client = Client::new(&transport);

    let xml = export_web_part(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::Shared,
        "dddddddd-dddd-dddd-dddd-dddddddddddd",
    )
    .await
    .unwrap();

    assert_eq!(xml, "<webPart>xml</webPart>");
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/ExportWebPart"));
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"webPartId\":\"dddddddd-dddd-dddd-dddd-dddddddddddd\""));
}

#[tokio::test]
async fn import_sends_correct_body() {
    let transport = RecordingTransport::with_response("{}");
    let client = Client::new(&transport);

    import_web_part(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::Shared,
        "<webPart/>",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/ImportWebPart"));
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"webPartXml\":\"<webPart/>\""));
}

#[tokio::test]
async fn move_to_builds_correct_url() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    move_web_part_to(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::Shared,
        "dddddddd-dddd-dddd-dddd-dddddddddddd",
        "zone-1",
        2,
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/MoveWebPartTo(zoneID='zone-1', zoneIndex=2)"));
}

#[tokio::test]
async fn lifecycle_action_urls_are_correct() {
    let transport = RecordingTransport::with_response("{}");
    let client = Client::new(&transport);
    let id = "dddddddd-dddd-dddd-dddd-dddddddddddd";

    save_web_part_changes(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::Shared,
        id,
    )
    .await
    .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/SaveWebPartChanges"));

    close_web_part(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::Shared,
        id,
    )
    .await
    .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/CloseWebPart"));

    open_web_part(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::Shared,
        id,
    )
    .await
    .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/OpenWebPart"));

    delete_web_part(
        &client,
        SITE,
        PAGE,
        WebPartsPersonalizationScope::Shared,
        id,
    )
    .await
    .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/DeleteWebPart"));
}
