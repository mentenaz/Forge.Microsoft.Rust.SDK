use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_groupsitemanager::{
    can_user_create_group, clear_current_user_teams_cache, create, create_group, delete,
    ensure_team_for_group, ensure_team_for_group_ex, get_all_org_labels,
    get_group_creation_context, get_group_site_conversion_data, get_site_status,
    get_valid_site_url_from_alias, is_teamify_prompt_hidden, notebook,
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
        *self.last_body.lock().unwrap() = body.map(|b| String::from_utf8_lossy(b).into_owned());
        let response = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(response.into_bytes()) })
    }
}

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

#[tokio::test]
async fn can_user_create_group_parses_bool() {
    let transport = RecordingTransport::with_response("true");
    let client = Client::new(&transport);

    let result = can_user_create_group(&client, SITE).await.unwrap();

    assert!(result);
    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!("{SITE}/_api/groupsitemanager/CanUserCreateGroup")
    );
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "GET"
    );
}

#[tokio::test]
async fn clear_current_user_teams_cache_sends_bodiless_post() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    clear_current_user_teams_cache(&client, SITE).await.unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    assert!(transport.last_body.lock().unwrap().is_none());
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/ClearCurrentUserTeamsCache"));
}

#[tokio::test]
async fn create_sends_group_id_and_parses_site_info() {
    const RESPONSE: &str = r#"{"DocumentsUrl":"https://contoso.sharepoint.com/sites/team/Shared Documents","ErrorMessage":null,"GroupId":"11111111-1111-1111-1111-111111111111","SiteStatus":2,"SiteUrl":"https://contoso.sharepoint.com/sites/team"}"#;
    let transport = RecordingTransport::with_response(RESPONSE);
    let client = Client::new(&transport);

    let info = create(&client, SITE, "11111111-1111-1111-1111-111111111111")
        .await
        .unwrap();

    assert_eq!(info.site_status, 2);
    assert_eq!(info.group_id, "11111111-1111-1111-1111-111111111111");
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"groupId\":\"11111111-1111-1111-1111-111111111111\""));
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/Create"));
}

#[tokio::test]
async fn create_group_sends_all_fields() {
    const RESPONSE: &str = r#"{"DocumentsUrl":"","ErrorMessage":null,"GroupId":"22222222-2222-2222-2222-222222222222","SiteStatus":0,"SiteUrl":""}"#;
    let transport = RecordingTransport::with_response(RESPONSE);
    let client = Client::new(&transport);

    let owners = vec!["alice@contoso.com".to_string()];
    let options = vec!["HubSiteId:33333333-3333-3333-3333-333333333333".to_string()];

    create_group(
        &client,
        SITE,
        "Marketing",
        "marketing",
        true,
        &owners,
        "Marketing team site",
        &options,
    )
    .await
    .unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"displayName\":\"Marketing\""));
    assert!(body.contains("\"alias\":\"marketing\""));
    assert!(body.contains("\"isPublic\":true"));
    assert!(body.contains("alice@contoso.com"));
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/CreateGroup"));
}

#[tokio::test]
async fn delete_sends_target_site_url() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    delete(&client, SITE, "https://contoso.sharepoint.com/sites/old")
        .await
        .unwrap();

    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("https://contoso.sharepoint.com/sites/old"));
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/Delete"));
}

#[tokio::test]
async fn ensure_team_for_group_parses_string_response() {
    let transport =
        RecordingTransport::with_response(r#""https://teams.microsoft.com/l/team/abc""#);
    let client = Client::new(&transport);

    let url = ensure_team_for_group(&client, SITE).await.unwrap();

    assert_eq!(url, "https://teams.microsoft.com/l/team/abc");
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/EnsureTeamForGroup"));
}

#[tokio::test]
async fn ensure_team_for_group_ex_parses_response() {
    const RESPONSE: &str = r#"{"teamsId":"44444444-4444-4444-4444-444444444444","teamsUrl":"https://teams.microsoft.com/l/team/xyz"}"#;
    let transport = RecordingTransport::with_response(RESPONSE);
    let client = Client::new(&transport);

    let resp = ensure_team_for_group_ex(&client, SITE).await.unwrap();

    assert_eq!(resp.teams_id, "44444444-4444-4444-4444-444444444444");
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/EnsureTeamForGroupEx"));
}

#[tokio::test]
async fn get_all_org_labels_sends_page_number() {
    const RESPONSE: &str = r#"{"IsLastPage":true,"Labels":[{"DisplayName":"Confidential","LabelApplicableTo":"Site","ObjectId":"55555555-5555-5555-5555-555555555555"}]}"#;
    let transport = RecordingTransport::with_response(RESPONSE);
    let client = Client::new(&transport);

    let labels = get_all_org_labels(&client, SITE, 1).await.unwrap();

    assert!(labels.is_last_page);
    assert_eq!(labels.labels.len(), 1);
    assert_eq!(labels.labels[0].display_name, "Confidential");
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"pageNumber\":1"));
}

#[tokio::test]
async fn get_group_creation_context_is_raw_json() {
    const RESPONSE: &str = r#"{"SitePath":"/sites","ShowSelfServiceSiteCreation":true}"#;
    let transport = RecordingTransport::with_response(RESPONSE);
    let client = Client::new(&transport);

    let ctx = get_group_creation_context(&client, SITE).await.unwrap();

    assert_eq!(ctx["SitePath"], "/sites");
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "GET"
    );
}

#[tokio::test]
async fn get_group_site_conversion_data_parses_response() {
    const RESPONSE: &str = r#"{"GroupType":0,"IsGroupifyDisabled":false,"IsRegionRestricted":false,"IsWrongPdl":false,"SuggestedMembers":[],"SuggestedOwners":[],"UnsuggestablePrincipals":[]}"#;
    let transport = RecordingTransport::with_response(RESPONSE);
    let client = Client::new(&transport);

    let info = get_group_site_conversion_data(&client, SITE).await.unwrap();

    assert!(!info.is_groupify_disabled);
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "GET"
    );
}

#[tokio::test]
async fn get_site_status_sends_group_id() {
    const RESPONSE: &str = r#"{"DocumentsUrl":"","ErrorMessage":null,"GroupId":"66666666-6666-6666-6666-666666666666","SiteStatus":1,"SiteUrl":""}"#;
    let transport = RecordingTransport::with_response(RESPONSE);
    let client = Client::new(&transport);

    let info = get_site_status(&client, SITE, "66666666-6666-6666-6666-666666666666")
        .await
        .unwrap();

    assert_eq!(info.site_status, 1);
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/GetSiteStatus"));
}

#[tokio::test]
async fn get_valid_site_url_from_alias_sends_optional_fields() {
    let transport =
        RecordingTransport::with_response(r#""https://contoso.sharepoint.com/teams/finance""#);
    let client = Client::new(&transport);

    let url = get_valid_site_url_from_alias(&client, SITE, "finance", Some("/teams"), Some(true))
        .await
        .unwrap();

    assert_eq!(url, "https://contoso.sharepoint.com/teams/finance");
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"alias\":\"finance\""));
    assert!(body.contains("\"managedPath\":\"/teams\""));
    assert!(body.contains("\"isTeamSite\":true"));
}

#[tokio::test]
async fn is_teamify_prompt_hidden_parses_bool() {
    let transport = RecordingTransport::with_response("false");
    let client = Client::new(&transport);

    let hidden = is_teamify_prompt_hidden(&client, SITE, None).await.unwrap();

    assert!(!hidden);
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("\"siteUrl\":null"));
}

#[tokio::test]
async fn notebook_sends_group_id_and_parses_string() {
    let transport = RecordingTransport::with_response(
        r#""https://contoso.sharepoint.com/sites/team/notebook""#,
    );
    let client = Client::new(&transport);

    let url = notebook(&client, SITE, "77777777-7777-7777-7777-777777777777")
        .await
        .unwrap();

    assert_eq!(url, "https://contoso.sharepoint.com/sites/team/notebook");
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    assert!(body.contains("77777777-7777-7777-7777-777777777777"));
}
