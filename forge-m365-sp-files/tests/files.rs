use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_files::{delete_file, download_file, get_file_by_path, upload_file};
use std::collections::VecDeque;
use std::sync::Mutex;

#[derive(Debug, Clone)]
struct RecordedCall {
    method: String,
    url: String,
    headers: Vec<(String, String)>,
    body: Option<Vec<u8>>,
}

#[derive(Default)]
struct RecordingTransport {
    responses: Mutex<VecDeque<Result<Vec<u8>>>>,
    calls: Mutex<Vec<RecordedCall>>,
}

impl RecordingTransport {
    fn queue_bytes(self, response: Vec<u8>) -> Self {
        self.responses.lock().unwrap().push_back(Ok(response));
        self
    }

    fn queue(self, response: &str) -> Self {
        self.queue_bytes(response.as_bytes().to_vec())
    }
}

impl Transport for RecordingTransport {
    fn execute(
        &self,
        _surface: Surface,
        method: &str,
        url: &str,
        headers: &[(&str, &str)],
        body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        let call = RecordedCall {
            method: method.to_string(),
            url: url.to_string(),
            headers: headers
                .iter()
                .map(|(k, v)| (k.to_string(), v.to_string()))
                .collect(),
            body: body.map(<[u8]>::to_vec),
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

const FILE_INFO_JSON: &str = r#"{"Name":"report.docx","ServerRelativeUrl":"/sites/team/Shared Documents/report.docx","Length":"12345","TimeLastModified":"2026-08-26T00:00:00Z","UniqueId":"44444444-4444-4444-4444-444444444444"}"#;

#[tokio::test]
async fn parses_file_info() {
    let transport = RecordingTransport::default().queue(FILE_INFO_JSON);
    let client = Client::new(&transport);

    let file = get_file_by_path(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "/sites/team/Shared Documents/report.docx",
    )
    .await
    .unwrap();

    assert_eq!(file.name, "report.docx");
    assert_eq!(file.length, "12345");
}

#[tokio::test]
async fn download_returns_raw_bytes_not_json() {
    let raw = vec![0u8, 159, 146, 150, 1, 2, 3];
    let transport = RecordingTransport::default().queue_bytes(raw.clone());
    let client = Client::new(&transport);

    let bytes = download_file(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "/sites/team/Shared Documents/photo.png",
    )
    .await
    .unwrap();

    assert_eq!(bytes, raw);

    let calls = transport.calls.lock().unwrap();
    assert!(calls[0].url.ends_with("/$value"));
}

#[tokio::test]
async fn upload_sends_octet_stream_and_overwrite_flag() {
    let transport = RecordingTransport::default().queue(FILE_INFO_JSON);
    let client = Client::new(&transport);

    let content = b"hello world";
    upload_file(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "/sites/team/Shared Documents",
        "report.docx",
        content,
        true,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.contains("Overwrite=true"));
    assert!(calls[0].headers.contains(&(
        "Content-Type".to_string(),
        "application/octet-stream".to_string()
    )));
    assert_eq!(calls[0].body.as_deref(), Some(content.as_slice()));
}

#[tokio::test]
async fn delete_sends_delete_with_if_match() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    delete_file(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "/sites/team/Shared Documents/report.docx",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "DELETE");
    assert!(calls[0]
        .headers
        .contains(&("IF-MATCH".to_string(), "*".to_string())));
}
