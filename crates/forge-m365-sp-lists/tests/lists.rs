use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_lists::{add_item, delete_item, get_items, get_list_by_title, update_item};
use serde_json::{json, Map, Value};
use std::collections::VecDeque;
use std::sync::Mutex;

#[derive(Debug, Clone)]
struct RecordedCall {
    method: String,
    url: String,
    headers: Vec<(String, String)>,
    body: Option<Vec<u8>>,
}

/// Returns scripted responses in call order (FIFO) and records what each call
/// actually sent, so tests can assert on request construction (headers,
/// method, `__metadata` body shape), not just response parsing.
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

const LIST_INFO_JSON: &str = r#"{"Id":"33333333-3333-3333-3333-333333333333","Title":"Tasks","ItemCount":2,"BaseTemplate":100,"Hidden":false,"ListItemEntityTypeFullName":"SP.Data.TasksListItem"}"#;

#[tokio::test]
async fn parses_list_info() {
    let transport = RecordingTransport::default().queue(LIST_INFO_JSON);
    let client = Client::new(&transport);

    let list = get_list_by_title(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
    )
    .await
    .unwrap();

    assert_eq!(list.title, "Tasks");
    assert_eq!(list.item_count, 2);
    assert_eq!(
        list.list_item_entity_type_full_name,
        "SP.Data.TasksListItem"
    );
}

#[tokio::test]
async fn get_list_by_title_encodes_quotes_in_url() {
    let transport = RecordingTransport::default().queue(LIST_INFO_JSON);
    let client = Client::new(&transport);

    get_list_by_title(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "O'Brien's List",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert!(calls[0].url.contains("O%27%27Brien%27%27s"));
}

#[tokio::test]
async fn parses_items() {
    let json = r#"[{"Id":1,"Title":"First"},{"Id":2,"Title":"Second"}]"#;
    let transport = RecordingTransport::default().queue(json);
    let client = Client::new(&transport);

    let items = get_items(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
    )
    .await
    .unwrap();

    assert_eq!(items.len(), 2);
    assert_eq!(items[0]["Title"], json!("First"));
}

#[tokio::test]
async fn add_item_fetches_entity_type_and_attaches_metadata() {
    let created_json = r#"{"Id":3,"Title":"New task"}"#;
    let transport = RecordingTransport::default()
        .queue(LIST_INFO_JSON)
        .queue(created_json);
    let client = Client::new(&transport);

    let mut fields = Map::new();
    fields.insert("Title".to_string(), Value::String("New task".to_string()));

    let created = add_item(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        fields,
    )
    .await
    .unwrap();

    assert_eq!(created["Id"], json!(3));

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls.len(), 2);
    assert_eq!(calls[0].method, "GET");
    assert_eq!(calls[1].method, "POST");
    assert!(calls[1].url.ends_with("/items"));

    let sent: Value = serde_json::from_slice(calls[1].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["Title"], json!("New task"));
    assert_eq!(sent["__metadata"]["type"], json!("SP.Data.TasksListItem"));
}

#[tokio::test]
async fn update_item_sends_patch_with_if_match() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    let mut fields = Map::new();
    fields.insert("Title".to_string(), Value::String("Renamed".to_string()));

    update_item(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        42,
        fields,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "PATCH");
    assert!(calls[0].url.ends_with("/items(42)"));
    assert!(calls[0]
        .headers
        .contains(&("IF-MATCH".to_string(), "*".to_string())));

    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["Title"], json!("Renamed"));
    assert!(sent.get("__metadata").is_none());
}

#[tokio::test]
async fn delete_item_sends_delete_with_if_match_and_no_body() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    delete_item(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        42,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "DELETE");
    assert!(calls[0].url.ends_with("/items(42)"));
    assert!(calls[0]
        .headers
        .contains(&("IF-MATCH".to_string(), "*".to_string())));
    assert!(calls[0].body.is_none());
}
