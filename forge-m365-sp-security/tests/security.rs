use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_security::{
    break_role_inheritance, get_role_assignments, get_role_definitions,
    get_user_effective_permissions,
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

#[tokio::test]
async fn parses_role_assignments_with_wrapped_bindings() {
    let json = r#"{"results":[{"PrincipalId":5,"Member":{"Id":5,"Title":"Team Owners","LoginName":"Team Owners","PrincipalType":8},"RoleDefinitionBindings":{"results":[{"Id":1073741829,"Name":"Full Control","Description":"","Hidden":false,"Order":1}]}}]}"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let assignments = get_role_assignments(
        &client,
        "https://contoso.sharepoint.com/sites/team/_api/web",
    )
    .await
    .unwrap();

    assert_eq!(assignments.len(), 1);
    assert_eq!(assignments[0].principal_id, 5);
    assert_eq!(assignments[0].member.as_ref().unwrap().title, "Team Owners");
    assert_eq!(assignments[0].role_definition_bindings.len(), 1);
    assert_eq!(
        assignments[0].role_definition_bindings[0].name,
        "Full Control"
    );
}

#[tokio::test]
async fn parses_role_assignments_bare_arrays() {
    let json = r#"[{"PrincipalId":5,"Member":null,"RoleDefinitionBindings":[{"Id":1073741826,"Name":"Read","Description":"","Hidden":false,"Order":2}]}]"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let assignments = get_role_assignments(
        &client,
        "https://contoso.sharepoint.com/sites/team/_api/web",
    )
    .await
    .unwrap();

    assert_eq!(assignments.len(), 1);
    assert!(assignments[0].member.is_none());
    assert_eq!(assignments[0].role_definition_bindings[0].name, "Read");
}

#[tokio::test]
async fn parses_role_definitions() {
    let json =
        r#"[{"Id":1073741829,"Name":"Full Control","Description":"","Hidden":false,"Order":1}]"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let defs = get_role_definitions(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(defs.len(), 1);
    assert_eq!(defs[0].name, "Full Control");
}

#[tokio::test]
async fn break_role_inheritance_sends_bodiless_post_with_flags() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    break_role_inheritance(
        &client,
        "https://contoso.sharepoint.com/sites/team/_api/web",
        true,
        false,
    )
    .await
    .unwrap();

    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.contains("copyroleassignments=true"));
    assert!(url.contains("clearsubscopes=false"));
}

#[tokio::test]
async fn get_user_effective_permissions_parses_int64_strings() {
    let json = r#"{"Low":"1000000000","High":"0"}"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let perms = get_user_effective_permissions(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "i:0#.f|membership|jane@contoso.com",
    )
    .await
    .unwrap();

    assert_eq!(perms.low, "1000000000");
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.contains("getUserEffectivePermissions(@user)?@user="));
}
