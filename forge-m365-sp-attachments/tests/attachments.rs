use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_attachments::{
    add_attachment, delete_attachment, download_attachment, get_attachments,
};
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

const ATTACHMENT_JSON: &str = r#"{"FileName":"notes.txt","ServerRelativeUrl":"/sites/team/Lists/Tasks/Attachments/1/notes.txt"}"#;

// The file name is percent-encoded in the URL (NON_ALPHANUMERIC, same as
// every other crate's path-segment encoding), so '.' becomes %2E — this is
// the *encoded* segment the assertions below check for, not the decoded
// FileName value the server returns in JSON (which stays "notes.txt").
const ENCODED_NAME: &str = "notes%2Etxt";

#[tokio::test]
async fn parses_attachments_array() {
    let transport = RecordingTransport::default().queue(&format!("[{ATTACHMENT_JSON}]"));
    let client = Client::new(&transport);

    let attachments = get_attachments(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
    )
    .await
    .unwrap();

    assert_eq!(attachments.len(), 1);
    assert_eq!(attachments[0].file_name, "notes.txt");
}

#[tokio::test]
async fn download_returns_raw_bytes_and_builds_value_url() {
    let raw = vec![1u8, 2, 3, 255];
    let transport = RecordingTransport::default().queue_bytes(raw.clone());
    let client = Client::new(&transport);

    let bytes = download_attachment(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
        "notes.txt",
    )
    .await
    .unwrap();

    assert_eq!(bytes, raw);
    let calls = transport.calls.lock().unwrap();
    assert!(calls[0]
        .url
        .ends_with(&format!("AttachmentFiles('{ENCODED_NAME}')/$value")));
}

#[tokio::test]
async fn add_sends_octet_stream_and_correct_url() {
    let transport = RecordingTransport::default().queue(ATTACHMENT_JSON);
    let client = Client::new(&transport);

    let content = b"hello attachment";
    add_attachment(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
        "notes.txt",
        content,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0]
        .url
        .ends_with(&format!("AttachmentFiles/add(FileName='{ENCODED_NAME}')")));
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

    delete_attachment(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
        "notes.txt",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "DELETE");
    assert!(calls[0]
        .url
        .ends_with(&format!("AttachmentFiles('{ENCODED_NAME}')")));
    assert!(calls[0]
        .headers
        .contains(&("IF-MATCH".to_string(), "*".to_string())));
}
