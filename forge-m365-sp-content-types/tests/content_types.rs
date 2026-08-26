use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_content_types::{
    add_content_type_to_list, get_content_type_by_id, get_list_content_types, get_web_content_types,
};
use serde_json::Value;
use std::collections::VecDeque;
use std::sync::Mutex;

#[derive(Debug, Clone)]
struct RecordedCall {
    method: String,
    url: String,
    body: Option<Vec<u8>>,
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
        _headers: &[(&str, &str)],
        body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        let call = RecordedCall {
            method: method.to_string(),
            url: url.to_string(),
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

const CONTENT_TYPE_JSON: &str = r#"{"Id":{"StringValue":"0x0101"},"Name":"Document","Description":"","Group":"_Hidden","Hidden":false,"ReadOnly":false,"Sealed":false,"Scope":"/sites/team"}"#;

#[tokio::test]
async fn parses_web_content_types() {
    let transport = RecordingTransport::default().queue(&format!("[{CONTENT_TYPE_JSON}]"));
    let client = Client::new(&transport);

    let cts = get_web_content_types(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(cts.len(), 1);
    assert_eq!(cts[0].name, "Document");
    assert_eq!(cts[0].id.string_value, "0x0101");
}

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::default().queue(CONTENT_TYPE_JSON);
    let client = Client::new(&transport);

    get_content_type_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "0x0101",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert!(calls[0].url.ends_with("/_api/web/contenttypes('0x0101')"));
}

#[tokio::test]
async fn get_list_content_types_builds_correct_url() {
    let transport = RecordingTransport::default().queue(&format!("[{CONTENT_TYPE_JSON}]"));
    let client = Client::new(&transport);

    let cts = get_list_content_types(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
    )
    .await
    .unwrap();

    assert_eq!(cts.len(), 1);
    let calls = transport.calls.lock().unwrap();
    assert!(calls[0]
        .url
        .ends_with("lists/getbytitle('Tasks')/contenttypes"));
}

#[tokio::test]
async fn add_to_list_sends_content_type_id_body() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    add_content_type_to_list(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        "0x0101",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0]
        .url
        .ends_with("lists/getbytitle('Tasks')/contenttypes/addAvailableContentType"));

    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["contentTypeId"], "0x0101");
}
