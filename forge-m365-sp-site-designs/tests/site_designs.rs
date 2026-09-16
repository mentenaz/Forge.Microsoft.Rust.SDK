use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_site_designs::{
    add_site_design_task, add_site_design_task_to_current_web, apply_site_design,
    create_site_design, delete_site_design, get_site_design_metadata, get_site_design_rights,
    get_site_design_run, get_site_design_run_status, get_site_design_task, get_site_designs,
    grant_site_design_rights, revoke_site_design_rights, update_site_design,
    SiteDesignCreationInfo, SiteDesignUpdateInfo,
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

const SITE: &str = "https://contoso.sharepoint.com/sites/team";
const DESIGN_ID: &str = "11111111-1111-1111-1111-111111111111";
const DESIGN_JSON: &str = r#"{"Id":"11111111-1111-1111-1111-111111111111","Title":"Contoso Design","WebTemplate":"64","SiteScriptIds":["22222222-2222-2222-2222-222222222222"],"Description":"desc","PreviewImageUrl":"","PreviewImageAltText":"","IsDefault":false,"Version":"1"}"#;

fn base_url() -> String {
    format!("{SITE}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility")
}

#[tokio::test]
async fn create_site_design_posts_wrapped_info() {
    let transport = RecordingTransport::with_response(DESIGN_JSON);
    let client = Client::new(&transport);

    let info = SiteDesignCreationInfo {
        title: "Contoso Design".into(),
        web_template: "64".into(),
        ..Default::default()
    };
    let created = create_site_design(&client, SITE, &info).await.unwrap();

    assert_eq!(created.title, "Contoso Design");
    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!("{}.CreateSiteDesign", base_url())
    );
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let value: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(value["info"]["Title"], "Contoso Design");
    assert_eq!(value["info"]["WebTemplate"], "64");
}

#[tokio::test]
async fn apply_site_design_sends_id_and_web_url() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    apply_site_design(
        &client,
        SITE,
        DESIGN_ID,
        "https://contoso.sharepoint.com/sites/new",
    )
    .await
    .unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let value: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(value["siteDesignId"], DESIGN_ID);
    assert_eq!(value["webUrl"], "https://contoso.sharepoint.com/sites/new");
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with(".ApplySiteDesign"));
}

#[tokio::test]
async fn get_site_designs_parses_array() {
    let transport = RecordingTransport::with_response(&format!("[{DESIGN_JSON}]"));
    let client = Client::new(&transport);

    let designs = get_site_designs(&client, SITE).await.unwrap();

    assert_eq!(designs.len(), 1);
    assert_eq!(designs[0].id, DESIGN_ID);
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with(".GetSiteDesigns"));
}

#[tokio::test]
async fn get_site_design_metadata_sends_id() {
    let transport = RecordingTransport::with_response(DESIGN_JSON);
    let client = Client::new(&transport);

    get_site_design_metadata(&client, SITE, DESIGN_ID)
        .await
        .unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let value: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(value["id"], DESIGN_ID);
}

#[tokio::test]
async fn update_site_design_wraps_update_info() {
    let transport = RecordingTransport::with_response(DESIGN_JSON);
    let client = Client::new(&transport);

    let update = SiteDesignUpdateInfo {
        id: DESIGN_ID.to_string(),
        fields: SiteDesignCreationInfo {
            title: "Renamed".into(),
            is_default: Some(true),
            ..Default::default()
        },
    };
    update_site_design(&client, SITE, &update).await.unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let value: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(value["updateInfo"]["Id"], DESIGN_ID);
    assert_eq!(value["updateInfo"]["Title"], "Renamed");
    assert_eq!(value["updateInfo"]["IsDefault"], true);
}

#[tokio::test]
async fn delete_site_design_sends_id() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    delete_site_design(&client, SITE, DESIGN_ID).await.unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let value: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(value["id"], DESIGN_ID);
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with(".DeleteSiteDesign"));
}

#[tokio::test]
async fn get_site_design_rights_parses_principals() {
    let transport = RecordingTransport::with_response(
        r#"[{"DisplayName":"Alice","PrincipalName":"alice@contoso.com","Rights":1}]"#,
    );
    let client = Client::new(&transport);

    let rights = get_site_design_rights(&client, SITE, DESIGN_ID)
        .await
        .unwrap();

    assert_eq!(rights.len(), 1);
    assert_eq!(rights[0].principal_name, "alice@contoso.com");
}

#[tokio::test]
async fn grant_rights_stringifies_granted_rights() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    grant_site_design_rights(
        &client,
        SITE,
        DESIGN_ID,
        &["alice@contoso.com".to_string()],
        1,
    )
    .await
    .unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let value: serde_json::Value = serde_json::from_str(&body).unwrap();
    assert_eq!(value["grantedRights"], "1");
    assert_eq!(value["id"], DESIGN_ID);
    assert_eq!(value["principalNames"][0], "alice@contoso.com");
}

#[tokio::test]
async fn revoke_rights_sends_principal_list() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    revoke_site_design_rights(&client, SITE, DESIGN_ID, &["alice@contoso.com".to_string()])
        .await
        .unwrap();

    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with(".RevokeSiteDesignRights"));
}

#[tokio::test]
async fn add_task_urls_are_correct() {
    let transport = RecordingTransport::with_response(
        r#"{"ID":"33333333-3333-3333-3333-333333333333","LogonName":"i:0#.f|membership|alice@contoso.com","SiteDesignID":"11111111-1111-1111-1111-111111111111","SiteID":"44444444-4444-4444-4444-444444444444","WebID":"55555555-5555-5555-5555-555555555555"}"#,
    );
    let client = Client::new(&transport);

    let task = add_site_design_task(
        &client,
        SITE,
        "https://contoso.sharepoint.com/sites/new",
        DESIGN_ID,
    )
    .await
    .unwrap();
    assert_eq!(task.site_design_id, DESIGN_ID);
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with(".AddSiteDesignTask"));

    add_site_design_task_to_current_web(&client, SITE, DESIGN_ID)
        .await
        .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with(".AddSiteDesignTaskToCurrentWeb"));
}

#[tokio::test]
async fn get_task_returns_none_once_finished() {
    let transport = RecordingTransport::with_response("{}");
    let client = Client::new(&transport);

    let task = get_site_design_task(&client, SITE, "task-id")
        .await
        .unwrap();
    assert!(task.is_none());
}

#[tokio::test]
async fn get_task_returns_some_while_running() {
    let transport = RecordingTransport::with_response(
        r#"{"ID":"33333333-3333-3333-3333-333333333333","LogonName":"x","SiteDesignID":"11111111-1111-1111-1111-111111111111","SiteID":"44444444-4444-4444-4444-444444444444","WebID":"55555555-5555-5555-5555-555555555555"}"#,
    );
    let client = Client::new(&transport);

    let task = get_site_design_task(&client, SITE, "task-id")
        .await
        .unwrap();
    assert!(task.is_some());
}

#[tokio::test]
async fn get_run_and_run_status_urls_are_correct() {
    let transport = RecordingTransport::with_response("[]");
    let client = Client::new(&transport);
    let web_url = "https://contoso.sharepoint.com/sites/new";

    get_site_design_run(&client, SITE, web_url, Some(DESIGN_ID))
        .await
        .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with(".GetSiteDesignRun"));

    get_site_design_run_status(&client, SITE, web_url, "run-id")
        .await
        .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with(".GetSiteDesignRunStatus"));
}
