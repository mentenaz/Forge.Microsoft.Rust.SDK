use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_recycle_bin::{
    delete_recycle_bin_item, get_recycle_bin_items, restore_recycle_bin_item,
};
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

const RECYCLE_ITEM_JSON: &str = r#"{"Id":"bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb","Title":"report.docx","LeafName":"report.docx","DirName":"sites/team/Shared Documents","DeletedDate":"2026-08-26T00:00:00Z","DeletedByName":"Jane Doe","Size":12345,"ItemType":1,"ItemState":1}"#;

#[tokio::test]
async fn parses_recycle_bin_items() {
    let transport = RecordingTransport::with_response(&format!("[{RECYCLE_ITEM_JSON}]"));
    let client = Client::new(&transport);

    let items = get_recycle_bin_items(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(items.len(), 1);
    assert_eq!(items[0].title, "report.docx");
}

#[tokio::test]
async fn restore_sends_bodiless_post_to_restore_url() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    restore_recycle_bin_item(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
    )
    .await
    .unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("RecycleBin('bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb')/Restore"));
}

#[tokio::test]
async fn delete_object_sends_bodiless_post_to_delete_url() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    delete_recycle_bin_item(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
    )
    .await
    .unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("RecycleBin('bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb')/DeleteObject"));
}
