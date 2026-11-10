use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_site_groups::{
    get_group_by_id, get_group_by_name, get_group_members, get_site_groups,
};
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            last_url: Mutex::new(None),
        }
    }
}

impl Transport for RecordingTransport {
    fn execute(
        &self,
        _surface: Surface,
        _method: &str,
        url: &str,
        _headers: &[(&str, &str)],
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        *self.last_url.lock().unwrap() = Some(url.to_string());
        let body = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

const GROUP_JSON: &str = r#"{"Id":5,"Title":"Team Owners","Description":"","LoginName":"Team Owners","OwnerTitle":"Team Owners","IsHiddenInUI":false,"PrincipalType":8}"#;

const USER_JSON: &str = r#"{"Id":7,"Title":"Jane Doe","LoginName":"i:0#.f|membership|jane@contoso.com","Email":"jane@contoso.com","PrincipalType":1,"IsSiteAdmin":false,"IsHiddenInUI":false}"#;

#[tokio::test]
async fn parses_site_groups_array() {
    let transport = RecordingTransport::with_response(&format!("[{GROUP_JSON}]"));
    let client = Client::new(&transport);

    let groups = get_site_groups(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(groups.len(), 1);
    assert_eq!(groups[0].title, "Team Owners");
}

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(GROUP_JSON);
    let client = Client::new(&transport);

    get_group_by_id(&client, "https://contoso.sharepoint.com/sites/team", 5)
        .await
        .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/_api/web/sitegroups(5)"));
}

#[tokio::test]
async fn get_by_name_encodes_and_builds_correct_url() {
    let transport = RecordingTransport::with_response(GROUP_JSON);
    let client = Client::new(&transport);

    get_group_by_name(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Team Owners",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("getbyname('Team%20Owners')"));
}

#[tokio::test]
async fn parses_group_members_as_site_user_info() {
    let transport = RecordingTransport::with_response(&format!("[{USER_JSON}]"));
    let client = Client::new(&transport);

    let members = get_group_members(&client, "https://contoso.sharepoint.com/sites/team", 5)
        .await
        .unwrap();

    assert_eq!(members.len(), 1);
    assert_eq!(members[0].email, "jane@contoso.com");

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/_api/web/sitegroups(5)/users"));
}
