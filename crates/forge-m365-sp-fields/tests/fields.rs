use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_fields::{creation, field, get_all_fields};
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

const FIELD_JSON: &str = r#"{"Id":"11111111-1111-1111-1111-111111111111","Title":"Status","InternalName":"Status","FieldTypeKind":6,"TypeAsString":"Choice","TypeDisplayName":"Choice","TypeShortDescription":"Choice (menu to fill in)","Description":"","Group":"Custom Columns","Hidden":false,"Required":false,"ReadOnlyField":false,"Indexed":false,"EnforceUniqueValues":false,"SchemaXml":"","Scope":"/sites/team","StaticName":"Status","DefaultValue":null,"Choices":["Active","Inactive"],"EntityPropertyName":"Status","Filterable":true,"Sortable":true,"Sealed":false,"FromBaseType":false}"#;

#[tokio::test]
async fn parses_field_array() {
    let transport = RecordingTransport::default().queue(&format!("[{FIELD_JSON}]"));
    let client = Client::new(&transport);

    let fields = get_all_fields(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(fields.len(), 1);
    assert_eq!(fields[0].title, "Status");
    assert_eq!(fields[0].internal_name, "Status");
    assert!(fields[0]
        .choices
        .as_ref()
        .unwrap()
        .contains(&"Active".to_string()));
}

#[tokio::test]
async fn get_all_fields_builds_correct_url() {
    let transport = RecordingTransport::default().queue("[]");
    let client = Client::new(&transport);

    get_all_fields(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    let url = transport.calls.lock().unwrap()[0].url.clone();
    assert!(url.ends_with("/_api/web/fields"));
}

#[tokio::test]
async fn get_field_by_id_builds_correct_url() {
    let transport = RecordingTransport::default().queue(FIELD_JSON);
    let client = Client::new(&transport);

    field::get_field_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "11111111-1111-1111-1111-111111111111",
    )
    .await
    .unwrap();

    let url = transport.calls.lock().unwrap()[0].url.clone();
    assert!(url.contains("fields('11111111-1111-1111-1111-111111111111')"));
}

#[tokio::test]
async fn get_field_by_title_encodes_quotes() {
    let transport = RecordingTransport::default().queue(FIELD_JSON);
    let client = Client::new(&transport);

    field::get_field_by_title(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "O'Brien's Field",
    )
    .await
    .unwrap();

    let url = transport.calls.lock().unwrap()[0].url.clone();
    assert!(url.contains("getbytitle('O%27%27Brien%27%27s%20Field')"));
}

#[tokio::test]
async fn add_text_field_posts_with_defaults() {
    let created_json = r#"{"Id":"22222222-2222-2222-2222-222222222222","Title":"My Text"}"#;
    let transport = RecordingTransport::default().queue(created_json);
    let client = Client::new(&transport);

    #[derive(serde::Serialize)]
    struct Empty {}
    let common = Empty {};

    creation::add_text(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "My Text",
        &common,
        None,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.ends_with("/_api/web/fields"));

    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["Title"], json!("My Text"));
    assert_eq!(sent["FieldTypeKind"], json!(2));
    assert_eq!(sent["MaxLength"], json!(255));
    assert_eq!(sent["__metadata"]["type"], json!("SP.FieldText"));
}

#[tokio::test]
async fn add_text_field_custom_max_length() {
    let created_json = r#"{"Id":"22222222-2222-2222-2222-222222222222","Title":"My Text"}"#;
    let transport = RecordingTransport::default().queue(created_json);
    let client = Client::new(&transport);

    #[derive(serde::Serialize)]
    struct Empty {}
    let common = Empty {};

    creation::add_text(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "My Text",
        &common,
        Some(500),
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["MaxLength"], json!(500));
}

#[tokio::test]
async fn add_choice_field_wraps_choices() {
    let created_json = r#"{"Id":"33333333-3333-3333-3333-333333333333","Title":"Priority"}"#;
    let transport = RecordingTransport::default().queue(created_json);
    let client = Client::new(&transport);

    #[derive(serde::Serialize)]
    struct Empty {}
    let common = Empty {};

    let choices = vec!["High".to_string(), "Medium".to_string(), "Low".to_string()];

    creation::add_choice(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Priority",
        &common,
        &choices,
        None,
        None,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["FieldTypeKind"], json!(6));
    assert_eq!(sent["Choices"]["results"], json!(["High", "Medium", "Low"]));
}

#[tokio::test]
async fn add_boolean_field_uses_correct_type() {
    let created_json = r#"{"Id":"44444444-4444-4444-4444-444444444444","Title":"Is Active"}"#;
    let transport = RecordingTransport::default().queue(created_json);
    let client = Client::new(&transport);

    #[derive(serde::Serialize)]
    struct Empty {}
    let common = Empty {};

    creation::add_boolean(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Is Active",
        &common,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["FieldTypeKind"], json!(8));
    assert_eq!(sent["__metadata"]["type"], json!("SP.Field"));
}

#[tokio::test]
async fn add_date_time_field_sets_defaults() {
    let created_json = r#"{"Id":"55555555-5555-5555-5555-555555555555","Title":"Start Date"}"#;
    let transport = RecordingTransport::default().queue(created_json);
    let client = Client::new(&transport);

    #[derive(serde::Serialize)]
    struct Empty {}
    let common = Empty {};

    creation::add_date_time(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Start Date",
        &common,
        None,
        None,
        None,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["FieldTypeKind"], json!(4));
    assert_eq!(sent["DateTimeCalendarType"], json!(1));
    assert_eq!(sent["DisplayFormat"], json!(0));
    assert_eq!(sent["FriendlyDisplayFormat"], json!(0));
}

#[tokio::test]
async fn add_multi_choice_wraps_choices() {
    let created_json = r#"{"Id":"66666666-6666-6666-6666-666666666666","Title":"Tags"}"#;
    let transport = RecordingTransport::default().queue(created_json);
    let client = Client::new(&transport);

    #[derive(serde::Serialize)]
    struct Empty {}
    let common = Empty {};

    let choices = vec!["Rust".to_string(), "WASM".to_string()];

    creation::add_multi_choice(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tags",
        &common,
        &choices,
        Some(true),
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["FieldTypeKind"], json!(15));
    assert_eq!(sent["Choices"]["results"], json!(["Rust", "WASM"]));
    assert_eq!(sent["FillInChoice"], json!(true));
}

#[tokio::test]
async fn delete_field_sends_delete_with_if_match() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    field::delete_field(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "11111111-1111-1111-1111-111111111111",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "DELETE");
    assert!(calls[0]
        .url
        .contains("fields('11111111-1111-1111-1111-111111111111')"));
    assert!(calls[0]
        .headers
        .contains(&("IF-MATCH".to_string(), "*".to_string())));
    assert!(calls[0].body.is_none());
}

#[tokio::test]
async fn update_field_sends_patch_with_if_match() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    let mut props = Map::new();
    props.insert(
        "Description".to_string(),
        Value::String("Updated".to_string()),
    );

    field::update_field(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "11111111-1111-1111-1111-111111111111",
        props,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "PATCH");
    assert!(calls[0]
        .url
        .contains("fields('11111111-1111-1111-1111-111111111111')"));
    assert!(calls[0]
        .headers
        .contains(&("IF-MATCH".to_string(), "*".to_string())));

    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["Description"], json!("Updated"));
}

#[tokio::test]
async fn set_show_in_display_form_posts_correct_url() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    field::set_show_in_display_form(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "11111111-1111-1111-1111-111111111111",
        true,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    assert!(calls[0].url.contains("setshowindisplayform(true)"));
}

#[tokio::test]
async fn set_show_in_new_form_posts_correct_url() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    field::set_show_in_new_form(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "11111111-1111-1111-1111-111111111111",
        false,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert!(calls[0].url.contains("setshowinnewform(false)"));
}

#[tokio::test]
async fn field_type_odata_mapping() {
    use forge_m365_sp_fields::FieldType;

    assert_eq!(FieldType::Text.odata_type(), "SP.FieldText");
    assert_eq!(FieldType::Choice.odata_type(), "SP.FieldChoice");
    assert_eq!(FieldType::Lookup.odata_type(), "SP.FieldLookup");
    assert_eq!(FieldType::User.odata_type(), "SP.FieldUser");
    assert_eq!(FieldType::Note.odata_type(), "SP.FieldMultiLineText");
    assert_eq!(FieldType::Invalid.odata_type(), "SP.Field");
}
