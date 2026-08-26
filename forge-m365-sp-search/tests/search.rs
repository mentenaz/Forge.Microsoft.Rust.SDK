use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_search::search;
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

const WRAPPED_RESPONSE: &str = r#"{"ElapsedTime":42,"PrimaryQueryResult":{"RelevantResults":{"RowCount":1,"TotalRows":1,"Table":{"Rows":{"results":[{"Cells":{"results":[{"Key":"Title","Value":"Report.docx","ValueType":"Edm.String"},{"Key":"Path","Value":"https://contoso.sharepoint.com/sites/team/report.docx","ValueType":"Edm.String"}]}}]}}}}}"#;

const BARE_RESPONSE: &str = r#"{"ElapsedTime":42,"PrimaryQueryResult":{"RelevantResults":{"RowCount":1,"TotalRows":1,"Table":{"Rows":[{"Cells":[{"Key":"Title","Value":"Report.docx","ValueType":"Edm.String"}]}]}}}}"#;

#[tokio::test]
async fn sends_query_text_row_limit_and_wrapped_select_properties() {
    let transport = RecordingTransport::default().queue(WRAPPED_RESPONSE);
    let client = Client::new(&transport);

    search(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "contoso",
        Some(5),
        &["Title", "Path"],
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.ends_with("/_api/search/postquery"));

    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["request"]["Querytext"], "contoso");
    assert_eq!(sent["request"]["RowLimit"], 5);
    assert_eq!(
        sent["request"]["SelectProperties"]["results"],
        serde_json::json!(["Title", "Path"])
    );
}

#[tokio::test]
async fn omits_optional_fields_when_not_given() {
    let transport = RecordingTransport::default().queue(WRAPPED_RESPONSE);
    let client = Client::new(&transport);

    search(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "contoso",
        None,
        &[],
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert!(sent["request"].get("RowLimit").is_none());
    assert!(sent["request"].get("SelectProperties").is_none());
}

#[tokio::test]
async fn parses_results_wrapped_response() {
    let transport = RecordingTransport::default().queue(WRAPPED_RESPONSE);
    let client = Client::new(&transport);

    let results = search(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "contoso",
        None,
        &[],
    )
    .await
    .unwrap();

    assert_eq!(results.row_count, 1);
    assert_eq!(results.total_rows, 1);
    assert_eq!(results.results.len(), 1);
    assert_eq!(results.results[0]["Title"], "Report.docx");
    assert_eq!(
        results.results[0]["Path"],
        "https://contoso.sharepoint.com/sites/team/report.docx"
    );
}

#[tokio::test]
async fn parses_bare_array_response() {
    let transport = RecordingTransport::default().queue(BARE_RESPONSE);
    let client = Client::new(&transport);

    let results = search(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "contoso",
        None,
        &[],
    )
    .await
    .unwrap();

    assert_eq!(results.results.len(), 1);
    assert_eq!(results.results[0]["Title"], "Report.docx");
}
