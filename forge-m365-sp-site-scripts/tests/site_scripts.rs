use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_site_scripts::{
    create_site_script, delete_site_script, execute_site_script_action, get_site_script_from_list,
    get_site_script_from_web, get_site_script_metadata, get_site_scripts, update_site_script,
    SiteScriptSerializationInfo, SiteScriptUpdateInfo,
};
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
    last_method: Mutex<Option<String>>,
    last_body: Mutex<Option<String>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            last_url: Mutex::new(None),
            last_method: Mutex::new(None),
            last_body: Mutex::new(None),
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
        body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        *self.last_url.lock().unwrap() = Some(url.to_string());
        *self.last_method.lock().unwrap() = Some(method.to_string());
        *self.last_body.lock().unwrap() = body.map(|b| String::from_utf8_lossy(b).to_string());
        let response = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(response.into_bytes()) })
    }
}

const SITE_SCRIPT_JSON: &str = r#"{"Id":"aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa","Title":"Contoso Script","Description":"desc","Content":"{}","Version":"1"}"#;
const SCRIPT_ID: &str = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
const ADMIN_SITE: &str = "https://contoso-admin.sharepoint.com";

#[tokio::test]
async fn get_site_scripts_posts_to_utility_endpoint() {
    let transport = RecordingTransport::with_response(&format!("[{SITE_SCRIPT_JSON}]"));
    let client = Client::new(&transport);

    let scripts = get_site_scripts(&client, ADMIN_SITE).await.unwrap();

    assert_eq!(scripts.len(), 1);
    assert_eq!(scripts[0].title, "Contoso Script");
    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!(
            "{ADMIN_SITE}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts"
        )
    );
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
}

#[tokio::test]
async fn create_site_script_encodes_title_and_description_in_url() {
    let transport = RecordingTransport::with_response(SITE_SCRIPT_JSON);
    let client = Client::new(&transport);

    let content = serde_json::json!({ "$schema": "schema.json", "actions": [] });
    create_site_script(&client, ADMIN_SITE, "Contoso's Script", "a desc", &content)
        .await
        .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.contains("CreateSiteScript(Title=@title,Description=@desc)?"));
    // `'` doubled to `''` then percent-encoded.
    assert!(url.contains("@title='Contoso%27%27s%20Script'"));
    assert!(url.contains("@desc='a%20desc'"));

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let parsed: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(parsed["actions"], serde_json::json!([]));
}

#[tokio::test]
async fn get_site_script_metadata_sends_id_body() {
    let transport = RecordingTransport::with_response(SITE_SCRIPT_JSON);
    let client = Client::new(&transport);

    let script = get_site_script_metadata(&client, ADMIN_SITE, SCRIPT_ID)
        .await
        .unwrap();

    assert_eq!(script.id, SCRIPT_ID);
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert_eq!(body, format!(r#"{{"id":"{SCRIPT_ID}"}}"#));
}

#[tokio::test]
async fn delete_site_script_sends_id_body() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    delete_site_script(&client, ADMIN_SITE, SCRIPT_ID)
        .await
        .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with(".SiteScriptUtility.DeleteSiteScript"));
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert_eq!(body, format!(r#"{{"id":"{SCRIPT_ID}"}}"#));
}

#[tokio::test]
async fn update_site_script_stringifies_content_override() {
    let transport = RecordingTransport::with_response(SITE_SCRIPT_JSON);
    let client = Client::new(&transport);

    let update_info = SiteScriptUpdateInfo {
        id: SCRIPT_ID.to_string(),
        title: Some("New title".to_string()),
        ..Default::default()
    };
    let content = serde_json::json!({ "actions": [] });

    update_site_script(&client, ADMIN_SITE, update_info, Some(&content))
        .await
        .unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let parsed: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(parsed["updateInfo"]["Title"], "New title");
    assert_eq!(parsed["updateInfo"]["Content"], content.to_string());
}

#[tokio::test]
async fn get_site_script_from_list_sends_list_url() {
    let transport = RecordingTransport::with_response(r#""{}""#);
    let client = Client::new(&transport);

    let script = get_site_script_from_list(
        &client,
        ADMIN_SITE,
        "https://contoso.sharepoint.com/sites/s/Lists/L",
    )
    .await
    .unwrap();

    assert_eq!(script, "{}");
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"listUrl\""));
}

#[tokio::test]
async fn get_site_script_from_web_sends_web_url_and_info() {
    const RESULT_JSON: &str = r#"{"JSON":"{}","Warnings":[]}"#;
    let transport = RecordingTransport::with_response(RESULT_JSON);
    let client = Client::new(&transport);

    let info = SiteScriptSerializationInfo {
        include_branding: Some(true),
        ..Default::default()
    };
    let result = get_site_script_from_web(
        &client,
        ADMIN_SITE,
        "https://contoso.sharepoint.com/sites/s",
        &info,
    )
    .await
    .unwrap();

    assert_eq!(result.json, "{}");
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let parsed: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(parsed["info"]["IncludeBranding"], true);
}

#[tokio::test]
async fn execute_site_script_action_parses_outcome() {
    const RESULT_JSON: &str = r#"{"Title":"t","OutcomeText":"","Outcome":0,"Target":"web"}"#;
    let transport = RecordingTransport::with_response(RESULT_JSON);
    let client = Client::new(&transport);

    let result = execute_site_script_action(&client, ADMIN_SITE, "{}")
        .await
        .unwrap();

    assert_eq!(
        result.outcome,
        forge_m365_sp_site_scripts::SiteScriptActionOutcome::Success
    );
}
