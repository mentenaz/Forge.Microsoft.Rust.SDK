use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_items::{item, item_properties, versions};
use serde_json::{json, Value};
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

const ITEM_JSON: &str = r#"{"Id":1,"Title":"Test Item","Author":{"Title":"Test User","Id":1}}"#;

const VERSION_JSON: &str = r#"{"VersionLabel":"1.0","Id":1025,"Created":"2024-01-01T00:00:00Z","CreatedBy":{"Title":"Test User","Id":1},"LastModified":"2024-01-01T00:00:00Z","LastModifiedBy":{"Title":"Test User","Id":1},"CheckInComment":"","Size":1024}"#;

const PARENT_INFOS_JSON: &str = r#"{"Id":1,"ParentList":{"Id":"11111111-1111-1111-1111-111111111111","Title":"Test List","RootFolder":{"UniqueId":"22222222-2222-2222-2222-222222222222","ServerRelativeUrl":"/sites/team/Lists/TestList","ServerRelativePath":{"DecodedUrl":"/sites/team/Lists/TestList"}},"ParentWeb":{"Id":"33333333-3333-3333-3333-333333333333","Url":"https://contoso.sharepoint.com/sites/team","ServerRelativeUrl":"/sites/team","ServerRelativePath":{"DecodedUrl":"/sites/team"}}}}"#;

const PERMS_JSON: &str = r#"{"High":32767,"Low":65535}"#;

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::default().queue(ITEM_JSON);
    let client = Client::new(&transport);

    item::get_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "GET");
    assert!(calls[0].url.contains("/items(42)"));
}

#[tokio::test]
async fn get_by_string_id_builds_correct_url() {
    let transport = RecordingTransport::default().queue(ITEM_JSON);
    let client = Client::new(&transport);

    item::get_by_string_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        "abc-123",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "GET");
    assert!(calls[0].url.contains("getItemByStringId('abc-123')"));
}

#[tokio::test]
async fn recycle_sends_post_to_recycle_endpoint() {
    let transport = RecordingTransport::default().queue(r#""recycle-bin-id-123""#);
    let client = Client::new(&transport);

    let id = item::recycle(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
    )
    .await
    .unwrap();

    assert_eq!(id, "recycle-bin-id-123");
    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.contains("/items(42)/recycle"));
}

#[tokio::test]
async fn delete_with_params_sends_bypass_lock_body() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    item::delete_with_params(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
        true,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.contains("/DeleteWithParameters"));

    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["parameters"]["BypassSharedLock"], json!(true));
}

#[tokio::test]
async fn get_wopi_frame_url_sends_correct_action() {
    let transport = RecordingTransport::default().queue(r#""https://example.com/wopi""#);
    let client = Client::new(&transport);

    let url = item::get_wopi_frame_url(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
        1,
    )
    .await
    .unwrap();

    assert_eq!(url, "https://example.com/wopi");
    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.contains("getWOPIFrameUrl(@action)"));
    assert!(calls[0].url.contains("@action=1"));
}

#[tokio::test]
async fn validate_update_list_item_sends_form_values() {
    let transport = RecordingTransport::default().queue("[]");
    let client = Client::new(&transport);

    let form_values = vec![forge_m365_sp_items::types::ListItemFormUpdateValue {
        field_name: "Title".to_string(),
        field_value: "Updated".to_string(),
    }];

    item::validate_update_list_item(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
        &form_values,
        false,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.contains("/validateupdatelistitem"));

    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["formValues"][0]["FieldName"], json!("Title"));
    assert_eq!(sent["bNewDocumentUpdate"], json!(false));
}

#[tokio::test]
async fn get_parent_infos_builds_select_expand_url() {
    let transport = RecordingTransport::default().queue(PARENT_INFOS_JSON);
    let client = Client::new(&transport);

    let infos = item::get_parent_infos(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
    )
    .await
    .unwrap();

    assert_eq!(infos.parent_list.title, "Test List");
    assert_eq!(
        infos.parent_web.url,
        "https://contoso.sharepoint.com/sites/team"
    );

    let calls = transport.calls.lock().unwrap();
    assert!(calls[0].url.contains("$select="));
    assert!(calls[0].url.contains("$expand="));
}

#[tokio::test]
async fn get_effective_base_permissions_builds_correct_url() {
    let transport = RecordingTransport::default().queue(PERMS_JSON);
    let client = Client::new(&transport);

    item_properties::get_effective_base_permissions(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "GET");
    assert!(calls[0].url.contains("/EffectiveBasePermissions"));
}

#[tokio::test]
async fn get_field_values_as_text_builds_correct_url() {
    let transport = RecordingTransport::default().queue(r#"{"Title":"Test"}"#);
    let client = Client::new(&transport);

    item_properties::get_field_values_as_text(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "GET");
    assert!(calls[0].url.contains("/FieldValuesAsText"));
}

#[tokio::test]
async fn get_versions_builds_correct_url() {
    let transport = RecordingTransport::default().queue(&format!("[{VERSION_JSON}]"));
    let client = Client::new(&transport);

    let ver_list = versions::get_versions(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
    )
    .await
    .unwrap();

    assert_eq!(ver_list.len(), 1);
    assert_eq!(ver_list[0].version_label, "1.0");
    assert_eq!(ver_list[0].id, 1025);

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "GET");
    assert!(calls[0].url.contains("/versions"));
}

#[tokio::test]
async fn get_version_by_id_builds_correct_url() {
    let transport = RecordingTransport::default().queue(VERSION_JSON);
    let client = Client::new(&transport);

    let ver = versions::get_version_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
        1025,
    )
    .await
    .unwrap();

    assert_eq!(ver.version_label, "1.0");

    let calls = transport.calls.lock().unwrap();
    assert!(calls[0].url.contains("/versions(1025)"));
}

#[tokio::test]
async fn delete_version_sends_delete_with_if_match() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    versions::delete_version(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        42,
        1025,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "DELETE");
    assert!(calls[0].url.contains("/versions(1025)"));
    assert!(calls[0]
        .headers
        .contains(&("IF-MATCH".to_string(), "*".to_string())));
}

#[tokio::test]
async fn get_by_id_parses_item_json() {
    let transport = RecordingTransport::default().queue(ITEM_JSON);
    let client = Client::new(&transport);

    let item = item::get_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Test List",
        1,
    )
    .await
    .unwrap();

    assert_eq!(item["Id"], json!(1));
    assert_eq!(item["Title"], "Test Item");
}

#[tokio::test]
async fn list_title_with_quotes_is_encoded() {
    let transport = RecordingTransport::default().queue(ITEM_JSON);
    let client = Client::new(&transport);

    item::get_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "O'Brien's List",
        1,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert!(calls[0]
        .url
        .contains("getbytitle('O%27%27Brien%27%27s%20List')"));
}
