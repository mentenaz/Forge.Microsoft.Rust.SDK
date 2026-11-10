use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_folders::{add_folder, delete_folder, get_folder_by_path, get_subfolders};
use std::collections::VecDeque;
use std::sync::Mutex;

#[derive(Debug, Clone)]
struct RecordedCall {
    method: String,
    url: String,
    headers: Vec<(String, String)>,
}

#[derive(Default)]
struct RecordingTransport {
    responses: Mutex<VecDeque<Result<Vec<u8>>>>,
    calls: Mutex<Vec<RecordedCall>>,
}

impl RecordingTransport {
    fn queue(self, response: &str) -> Self {
        self.responses
            .lock()
            .unwrap()
            .push_back(Ok(response.as_bytes().to_vec()));
        self
    }
}

impl Transport for RecordingTransport {
    fn execute(
        &self,
        _surface: Surface,
        method: &str,
        url: &str,
        headers: &[(&str, &str)],
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        let call = RecordedCall {
            method: method.to_string(),
            url: url.to_string(),
            headers: headers
                .iter()
                .map(|(k, v)| (k.to_string(), v.to_string()))
                .collect(),
        };
        Box::pin(async move {
            self.calls.lock().unwrap().push(call);
            self.responses
                .lock()
                .unwrap()
                .pop_front()
                .unwrap_or_else(|| Err(Error::Unsupported("script-exhausted")))
        })
    }
}

const FOLDER_JSON: &str = r#"{"Name":"Reports","ServerRelativeUrl":"/sites/team/Shared Documents/Reports","ItemCount":3,"Exists":true,"TimeCreated":"2026-01-01T00:00:00Z","TimeLastModified":"2026-08-26T00:00:00Z","UniqueId":"77777777-7777-7777-7777-777777777777"}"#;

#[tokio::test]
async fn parses_folder_info() {
    let transport = RecordingTransport::default().queue(FOLDER_JSON);
    let client = Client::new(&transport);

    let folder = get_folder_by_path(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "/sites/team/Shared Documents/Reports",
    )
    .await
    .unwrap();

    assert_eq!(folder.name, "Reports");
    assert_eq!(folder.item_count, 3);
}

#[tokio::test]
async fn parses_subfolders_array() {
    let transport = RecordingTransport::default().queue(&format!("[{FOLDER_JSON}]"));
    let client = Client::new(&transport);

    let folders = get_subfolders(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "/sites/team/Shared Documents",
    )
    .await
    .unwrap();

    assert_eq!(folders.len(), 1);
    let calls = transport.calls.lock().unwrap();
    assert!(calls[0].url.ends_with("')/folders?$select=Name,ServerRelativeUrl,ItemCount,Exists,TimeCreated,TimeLastModified,UniqueId"));
}

#[tokio::test]
async fn add_sends_bodiless_post_with_overwrite_flag() {
    let transport = RecordingTransport::default().queue(FOLDER_JSON);
    let client = Client::new(&transport);

    add_folder(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "/sites/team/Shared Documents/Reports",
        true,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.contains("overwrite=true"));
}

#[tokio::test]
async fn delete_sends_delete_with_if_match() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    delete_folder(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "/sites/team/Shared Documents/Reports",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "DELETE");
    assert!(calls[0]
        .headers
        .contains(&("IF-MATCH".to_string(), "*".to_string())));
}
