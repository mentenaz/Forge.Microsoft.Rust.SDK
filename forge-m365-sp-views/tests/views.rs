use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_views::{get_view_by_id, get_view_by_title, get_view_fields, get_views};
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

const VIEW_JSON: &str = r#"{"Id":"88888888-8888-8888-8888-888888888888","Title":"All Items","Hidden":false,"DefaultView":true,"PersonalView":false,"ServerRelativeUrl":"/sites/team/Lists/Tasks/AllItems.aspx","ViewQuery":"","RowLimit":30}"#;

#[tokio::test]
async fn parses_views_array() {
    let transport = RecordingTransport::with_response(&format!("[{VIEW_JSON}]"));
    let client = Client::new(&transport);

    let views = get_views(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
    )
    .await
    .unwrap();

    assert_eq!(views.len(), 1);
    assert_eq!(views[0].title, "All Items");
    assert!(views[0].default_view);
}

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(VIEW_JSON);
    let client = Client::new(&transport);

    get_view_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        "88888888-8888-8888-8888-888888888888",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("views('88888888-8888-8888-8888-888888888888')"));
}

#[tokio::test]
async fn get_by_title_encodes_and_builds_correct_url() {
    let transport = RecordingTransport::with_response(VIEW_JSON);
    let client = Client::new(&transport);

    get_view_by_title(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        "All Items",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("views/getbytitle('All%20Items')"));
}

#[tokio::test]
async fn parses_view_fields_wrapped_response() {
    let json = r#"{"Items":{"results":["ID","Title","Modified"]},"SchemaXml":"<ViewFields>...</ViewFields>"}"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let fields = get_view_fields(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        "88888888-8888-8888-8888-888888888888",
    )
    .await
    .unwrap();

    assert_eq!(fields.items, vec!["ID", "Title", "Modified"]);
    assert!(fields.schema_xml.starts_with("<ViewFields>"));
}

#[tokio::test]
async fn parses_view_fields_bare_response() {
    let json = r#"{"Items":["ID","Title"],"SchemaXml":""}"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let fields = get_view_fields(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        "88888888-8888-8888-8888-888888888888",
    )
    .await
    .unwrap();

    assert_eq!(fields.items, vec!["ID", "Title"]);
}
