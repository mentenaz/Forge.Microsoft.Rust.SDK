use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_user_custom_actions::{add, clear, delete, get_all, get_by_id, update, Scope};
use serde_json::{Map, Value};
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
    last_method: Mutex<Option<String>>,
    last_headers: Mutex<Vec<(String, String)>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            last_url: Mutex::new(None),
            last_method: Mutex::new(None),
            last_headers: Mutex::new(Vec::new()),
        }
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
        *self.last_url.lock().unwrap() = Some(url.to_string());
        *self.last_method.lock().unwrap() = Some(method.to_string());
        *self.last_headers.lock().unwrap() = headers
            .iter()
            .map(|(k, v)| (k.to_string(), v.to_string()))
            .collect();
        let body = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

const ACTION_JSON: &str = r#"{"Id":"aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa","Name":"MyAction","Title":"My Action","Description":"","Location":"ScriptLink","Group":"","Sequence":1000,"RegistrationId":"","RegistrationType":0,"Scope":3,"ScriptBlock":"alert('hi');","ScriptSrc":null,"Url":null,"ImageUrl":null,"CommandUIExtension":null,"VersionOfUserCustomAction":"1.0.0.0","Rights":{"High":"0","Low":"0"}}"#;
const ACTION_ID: &str = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";

#[tokio::test]
async fn get_all_uses_web_or_site_segment() {
    let transport = RecordingTransport::with_response(&format!("[{ACTION_JSON}]"));
    let client = Client::new(&transport);

    let actions = get_all(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        Scope::Web,
    )
    .await
    .unwrap();
    assert_eq!(actions.len(), 1);
    assert_eq!(actions[0].name, "MyAction");
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/_api/web/usercustomactions"));

    get_all(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        Scope::Site,
    )
    .await
    .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/_api/site/usercustomactions"));
}

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(ACTION_JSON);
    let client = Client::new(&transport);

    let action = get_by_id(
        &client,
        "https://contoso.sharepoint.com",
        Scope::Web,
        ACTION_ID,
    )
    .await
    .unwrap();

    assert_eq!(action.id, ACTION_ID);
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with(&format!("usercustomactions('{ACTION_ID}')")));
}

#[tokio::test]
async fn add_sends_post_with_properties() {
    let transport = RecordingTransport::with_response(ACTION_JSON);
    let client = Client::new(&transport);

    let mut props = Map::new();
    props.insert("Name".to_string(), Value::String("MyAction".to_string()));
    props.insert(
        "Location".to_string(),
        Value::String("ScriptLink".to_string()),
    );

    add(&client, "https://contoso.sharepoint.com", Scope::Web, props)
        .await
        .unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/usercustomactions"));
}

#[tokio::test]
async fn update_sends_patch_with_if_match() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    let mut props = Map::new();
    props.insert("Title".to_string(), Value::String("Renamed".to_string()));

    update(
        &client,
        "https://contoso.sharepoint.com",
        Scope::Web,
        ACTION_ID,
        props,
    )
    .await
    .unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "PATCH"
    );
    let headers = transport.last_headers.lock().unwrap();
    assert!(headers.iter().any(|(k, v)| k == "IF-MATCH" && v == "*"));
}

#[tokio::test]
async fn delete_sends_delete_with_if_match() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    delete(
        &client,
        "https://contoso.sharepoint.com",
        Scope::Site,
        ACTION_ID,
    )
    .await
    .unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "DELETE"
    );
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with(&format!("usercustomactions('{ACTION_ID}')")));
    let headers = transport.last_headers.lock().unwrap();
    assert!(headers.iter().any(|(k, v)| k == "IF-MATCH" && v == "*"));
}

#[tokio::test]
async fn clear_sends_bodiless_post() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    clear(&client, "https://contoso.sharepoint.com", Scope::Web)
        .await
        .unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/usercustomactions/clear"));
}
