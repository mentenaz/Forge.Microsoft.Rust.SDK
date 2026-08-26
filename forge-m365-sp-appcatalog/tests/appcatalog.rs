use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_appcatalog::{
    deploy_app, get_app_by_id, get_available_apps, install_app, remove_app, retract_app,
    uninstall_app, upgrade_app,
};
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
    last_method: Mutex<Option<String>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            last_url: Mutex::new(None),
            last_method: Mutex::new(None),
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
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        *self.last_url.lock().unwrap() = Some(url.to_string());
        *self.last_method.lock().unwrap() = Some(method.to_string());
        let body = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

const APP_JSON: &str = r#"{"Id":"dddddddd-dddd-dddd-dddd-dddddddddddd","Title":"Contoso App","Deployed":true,"InstalledVersion":"1.0.0.0"}"#;
const APP_ID: &str = "dddddddd-dddd-dddd-dddd-dddddddddddd";

#[tokio::test]
async fn parses_available_apps() {
    let transport = RecordingTransport::with_response(&format!("[{APP_JSON}]"));
    let client = Client::new(&transport);

    let apps = get_available_apps(&client, "https://contoso-admin.sharepoint.com")
        .await
        .unwrap();

    assert_eq!(apps.len(), 1);
    assert_eq!(apps[0]["Title"], "Contoso App");
}

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(APP_JSON);
    let client = Client::new(&transport);

    get_app_by_id(&client, "https://contoso-admin.sharepoint.com", APP_ID)
        .await
        .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with(&format!("AvailableApps/getById('{APP_ID}')")));
}

#[tokio::test]
async fn deploy_sends_bodiless_post_with_flag() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    deploy_app(
        &client,
        "https://contoso-admin.sharepoint.com",
        APP_ID,
        true,
    )
    .await
    .unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with(&format!("getById('{APP_ID}')/Deploy(true)")));
}

#[tokio::test]
async fn action_urls_are_correct() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);
    let site = "https://contoso-admin.sharepoint.com";

    retract_app(&client, site, APP_ID).await.unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/Retract"));

    install_app(&client, site, APP_ID).await.unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/Install"));

    uninstall_app(&client, site, APP_ID).await.unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/Uninstall"));

    upgrade_app(&client, site, APP_ID).await.unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/Upgrade"));

    remove_app(&client, site, APP_ID).await.unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/Remove"));
}
