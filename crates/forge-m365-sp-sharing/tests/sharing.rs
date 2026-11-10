use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_sharing as sharing;
use forge_m365_sp_sharing::types::{
    RoleType, ShareObjectOptions, SharingLinkKind, SharingOperationStatusCode, SharingRecipient,
    SharingRole,
};
use std::collections::VecDeque;
use std::sync::Mutex;

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

const EMPTY_GUID: &str = "00000000-0000-0000-0000-000000000000";

const ITEM_BASE: &str =
    "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/items(3)";

const FILE_BASE: &str = "https://contoso.sharepoint.com/sites/team/_api/web/getFileByServerRelativePath(decodedUrl='%2Fsites%2Fteam%2FShared%20Documents%2Freport%2Edocx')/listItemAllFields";

const FOLDER_BASE: &str = "https://contoso.sharepoint.com/sites/team/_api/web/getFolderByServerRelativePath(decodedUrl='%2Fsites%2Fteam%2FShared%20Documents')/listItemAllFields";

#[derive(Debug, Clone)]
struct RecordedCall {
    method: String,
    url: String,
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

    fn assert_last(&self, expected: (&str, impl AsRef<str>)) {
        let calls = self.calls.lock().unwrap();
        let call = calls.last().expect("no call was recorded");
        assert_eq!(
            (call.method.as_str(), call.url.as_str()),
            (expected.0, expected.1.as_ref()),
            "unexpected request"
        );
    }

    fn assert_last_with_body(&self, expected: (&str, impl AsRef<str>), body: &str) {
        self.assert_last(expected);
        let calls = self.calls.lock().unwrap();
        let call = calls.last().unwrap();
        let sent =
            String::from_utf8(call.body.clone().unwrap_or_default()).expect("body was not utf-8");
        assert_eq!(sent, body, "unexpected body");
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
        let call = RecordedCall {
            method: method.to_string(),
            url: url.to_string(),
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

#[tokio::test]
async fn item_share_with_looks_up_role_then_shares() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"results":[{"Id":3}]}}"#)
        .queue(r#"{"StatusCode":0,"ErrorMessage":null}"#);
    let client = Client::new(&transport);

    let result = sharing::item_share_with(
        &client,
        SITE,
        "Tasks",
        3,
        &["i:0#.f|membership|user@contoso.com".to_string()],
        SharingRole::View,
        false,
        None,
    )
    .await
    .unwrap();

    assert_eq!(
        result.status_code,
        Some(SharingOperationStatusCode::CompletedSuccessfully)
    );
    let calls = transport.calls.lock().unwrap();
    assert_eq!(
        (calls[0].method.as_str(), calls[0].url.as_str()),
        (
            "GET",
            "https://contoso.sharepoint.com/sites/team/_api/web/roledefinitions?$select=Id&$filter=RoleTypeKind eq 2"
        )
    );
    let sent = String::from_utf8(calls[1].body.clone().unwrap_or_default()).unwrap();
    let share_url = format!("{ITEM_BASE}/shareObject");
    assert_eq!(
        (calls[1].method.as_str(), calls[1].url.as_str()),
        ("POST", share_url.as_str())
    );
    assert_eq!(
        sent,
        r#"{"includeAnonymousLinkInEmail":false,"peoplePickerInput":"[{\"Key\":\"i:0#.f|membership|user@contoso.com\"}]","propagateAcl":false,"roleValue":"role: 3","useSimplifiedRoles":true}"#,
        "unexpected body"
    );
}

#[tokio::test]
async fn item_share_with_edit_role_and_email() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"results":[{"Id":6}]}}"#)
        .queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    let email = sharing::types::SharingEmailData {
        subject: None,
        body: "Please review".into(),
    };
    sharing::item_share_with(
        &client,
        SITE,
        "Tasks",
        3,
        &["user@contoso.com".to_string()],
        SharingRole::Edit,
        true,
        Some(&email),
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(
        (calls[0].method.as_str(), calls[0].url.as_str()),
        (
            "GET",
            "https://contoso.sharepoint.com/sites/team/_api/web/roledefinitions?$select=Id&$filter=RoleTypeKind eq 3"
        )
    );
    let sent = String::from_utf8(calls[1].body.clone().unwrap_or_default()).unwrap();
    assert_eq!(
        sent,
        r#"{"emailBody":"Please review","emailSubject":"","includeAnonymousLinkInEmail":true,"peoplePickerInput":"[{\"Key\":\"user@contoso.com\"}]","propagateAcl":false,"roleValue":"role: 6","sendEmail":true,"useSimplifiedRoles":true}"#,
        "unexpected body"
    );
}

#[tokio::test]
async fn item_get_share_link() {
    let transport = RecordingTransport::default().queue(
        r#"{"sharingLinkInfo":{"ShareId":"11111111-1111-1111-1111-111111111111","Url":"https://contoso.sharepoint.com/sites/team/Tasks/3"}}"#,
    );
    let client = Client::new(&transport);

    let result = sharing::item_get_share_link(
        &client,
        SITE,
        "Tasks",
        3,
        SharingLinkKind::OrganizationView,
        None,
    )
    .await
    .unwrap();

    assert_eq!(
        result.sharing_link_info.unwrap().share_id.as_deref(),
        Some("11111111-1111-1111-1111-111111111111")
    );
    transport.assert_last_with_body(
        ("POST", &format!("{ITEM_BASE}/shareLink")),
        r#"{"request":{"createLink":true,"emailData":null,"settings":{"expiration":null,"linkKind":2}}}"#,
    );
}

#[tokio::test]
async fn item_check_permissions() {
    let transport = RecordingTransport::default().queue(
        r#"[{"inputEntity":"user@contoso.com","resolvedEntity":"User","hasAccess":true,"role":2}]"#,
    );
    let client = Client::new(&transport);

    let perms = sharing::item_check_permissions(
        &client,
        SITE,
        "Tasks",
        3,
        &[SharingRecipient {
            email: Some("user@contoso.com".into()),
            alias: None,
        }],
    )
    .await
    .unwrap();

    assert_eq!(perms[0].has_access, Some(true));
    transport.assert_last_with_body(
        ("POST", &format!("{ITEM_BASE}/checkPermissions")),
        r#"{"recipients":[{"email":"user@contoso.com"}]}"#,
    );
}

#[tokio::test]
async fn item_get_sharing_information_default_projection() {
    let transport = RecordingTransport::default().queue(
        r#"{"canAddExternalPrincipal":true,"webUrl":"https://contoso.sharepoint.com/sites/team"}"#,
    );
    let client = Client::new(&transport);

    let info = sharing::item_get_sharing_information(&client, SITE, "Tasks", 3, None, &[], &[])
        .await
        .unwrap();

    assert_eq!(info.can_add_external_principal, Some(true));
    transport.assert_last_with_body(
        (
            "POST",
            &format!("{ITEM_BASE}/getSharingInformation?$select=*"),
        ),
        r#"{"request":null}"#,
    );
}

#[tokio::test]
async fn item_get_sharing_information_custom_projection() {
    let transport = RecordingTransport::default().queue(r#"{"directUrl":"/sites/team/Tasks/3"}"#);
    let client = Client::new(&transport);

    sharing::item_get_sharing_information(
        &client,
        SITE,
        "Tasks",
        3,
        None,
        &["PermissionsInformation".to_string()],
        &["WebUrl".to_string(), "CurrentUser".to_string()],
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        (
            "POST",
            &format!(
                "{ITEM_BASE}/getSharingInformation?$select=WebUrl,CurrentUser&$expand=PermissionsInformation"
            ),
        ),
        r#"{"request":null}"#,
    );
}

#[tokio::test]
async fn item_get_object_sharing_settings() {
    let transport = RecordingTransport::default()
        .queue(r#"{"WebUrl":"https://contoso.sharepoint.com/sites/team","HasEditRole":true,"DefaultShareLinkType":2}"#);
    let client = Client::new(&transport);

    let settings = sharing::item_get_object_sharing_settings(&client, SITE, "Tasks", 3, true)
        .await
        .unwrap();

    assert_eq!(settings.web_url.as_deref(), Some(SITE));
    assert_eq!(settings.has_edit_role, Some(true));
    transport.assert_last_with_body(
        ("POST", &format!("{ITEM_BASE}/getObjectSharingSettings")),
        r#"{"useSimplifiedRoles":true}"#,
    );
}

#[tokio::test]
async fn item_unshare() {
    let transport = RecordingTransport::default().queue(r#"{"StatusCode":0,"ErrorMessage":null}"#);
    let client = Client::new(&transport);

    let result = sharing::item_unshare(&client, SITE, "Tasks", 3)
        .await
        .unwrap();

    assert_eq!(
        result.status_code,
        Some(SharingOperationStatusCode::CompletedSuccessfully)
    );
    transport.assert_last(("POST", &format!("{ITEM_BASE}/unshareObject")));
}

#[tokio::test]
async fn item_delete_sharing_link_by_kind() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    sharing::item_delete_sharing_link_by_kind(
        &client,
        SITE,
        "Tasks",
        3,
        SharingLinkKind::AnonymousView,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{ITEM_BASE}/deleteLinkByKind")),
        r#"{"linkKind":4}"#,
    );
}

#[tokio::test]
async fn item_unshare_link_defaults_to_empty_guid() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    sharing::item_unshare_link(
        &client,
        SITE,
        "Tasks",
        3,
        SharingLinkKind::AnonymousView,
        "",
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{ITEM_BASE}/unshareLink")),
        &format!(r#"{{"linkKind":4,"shareId":"{EMPTY_GUID}"}}"#),
    );
}

#[tokio::test]
async fn item_unshare_link_explicit_share_id() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    sharing::item_unshare_link(
        &client,
        SITE,
        "Tasks",
        3,
        SharingLinkKind::AnonymousView,
        "9b9b9b9b-9b9b-9b9b-9b9b-9b9b9b9b9b9b",
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{ITEM_BASE}/unshareLink")),
        r#"{"linkKind":4,"shareId":"9b9b9b9b-9b9b-9b9b-9b9b-9b9b9b9b9b9b"}"#,
    );
}

#[tokio::test]
async fn item_share_with_missing_role_definition_errors() {
    let transport = RecordingTransport::default().queue(r#"{"d":{"results":[]}}"#);
    let client = Client::new(&transport);

    let err = sharing::item_share_with(
        &client,
        SITE,
        "Tasks",
        3,
        &["user@contoso.com".to_string()],
        SharingRole::View,
        false,
        None,
    )
    .await
    .unwrap_err();

    assert!(matches!(
        err,
        Error::Surface(_, msg) if msg.contains("Could not locate associated role definition")
    ));
}

#[tokio::test]
async fn file_share_with() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"results":[{"Id":3}]}}"#)
        .queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    sharing::file_share_with(
        &client,
        SITE,
        "/sites/team/Shared Documents/report.docx",
        &["user@contoso.com".to_string()],
        SharingRole::View,
        false,
        None,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FILE_BASE}/shareObject")),
        r#"{"includeAnonymousLinkInEmail":false,"peoplePickerInput":"[{\"Key\":\"user@contoso.com\"}]","propagateAcl":false,"roleValue":"role: 3","useSimplifiedRoles":true}"#,
    );
}

#[tokio::test]
async fn file_get_share_link() {
    let transport = RecordingTransport::default().queue(r#"{"sharingLinkInfo":{"IsActive":true}}"#);
    let client = Client::new(&transport);

    sharing::file_get_share_link(
        &client,
        SITE,
        "/sites/team/Shared Documents/report.docx",
        SharingLinkKind::OrganizationEdit,
        None,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FILE_BASE}/shareLink")),
        r#"{"request":{"createLink":true,"emailData":null,"settings":{"expiration":null,"linkKind":3}}}"#,
    );
}

#[tokio::test]
async fn file_check_permissions() {
    let transport = RecordingTransport::default()
        .queue(r#"[{"inputEntity":"a@contoso.com","hasAccess":false,"role":0}]"#);
    let client = Client::new(&transport);

    sharing::file_check_permissions(
        &client,
        SITE,
        "/sites/team/Shared Documents/report.docx",
        &[SharingRecipient {
            email: Some("a@contoso.com".into()),
            alias: None,
        }],
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FILE_BASE}/checkPermissions")),
        r#"{"recipients":[{"email":"a@contoso.com"}]}"#,
    );
}

#[tokio::test]
async fn file_get_sharing_information() {
    let transport = RecordingTransport::default().queue(
        r#"{"hasUniquePermissions":false,"webUrl":"https://contoso.sharepoint.com/sites/team"}"#,
    );
    let client = Client::new(&transport);

    sharing::file_get_sharing_information(
        &client,
        SITE,
        "/sites/team/Shared Documents/report.docx",
        None,
        &[],
        &[],
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        (
            "POST",
            &format!("{FILE_BASE}/getSharingInformation?$select=*"),
        ),
        r#"{"request":null}"#,
    );
}

#[tokio::test]
async fn file_get_object_sharing_settings() {
    let transport = RecordingTransport::default().queue(r#"{"ItemName":"report.docx"}"#);
    let client = Client::new(&transport);

    sharing::file_get_object_sharing_settings(
        &client,
        SITE,
        "/sites/team/Shared Documents/report.docx",
        true,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FILE_BASE}/getObjectSharingSettings")),
        r#"{"useSimplifiedRoles":true}"#,
    );
}

#[tokio::test]
async fn file_unshare() {
    let transport = RecordingTransport::default().queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    sharing::file_unshare(&client, SITE, "/sites/team/Shared Documents/report.docx")
        .await
        .unwrap();

    transport.assert_last(("POST", &format!("{FILE_BASE}/unshareObject")));
}

#[tokio::test]
async fn file_delete_sharing_link_by_kind() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    sharing::file_delete_sharing_link_by_kind(
        &client,
        SITE,
        "/sites/team/Shared Documents/report.docx",
        SharingLinkKind::Flexible,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FILE_BASE}/deleteLinkByKind")),
        r#"{"linkKind":6}"#,
    );
}

#[tokio::test]
async fn file_unshare_link() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    sharing::file_unshare_link(
        &client,
        SITE,
        "/sites/team/Shared Documents/report.docx",
        SharingLinkKind::AnonymousView,
        "",
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FILE_BASE}/unshareLink")),
        &format!(r#"{{"linkKind":4,"shareId":"{EMPTY_GUID}"}}"#),
    );
}

#[tokio::test]
async fn folder_share_with() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"results":[{"Id":3}]}}"#)
        .queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    sharing::folder_share_with(
        &client,
        SITE,
        "/sites/team/Shared Documents",
        &["user@contoso.com".to_string()],
        SharingRole::View,
        false,
        None,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FOLDER_BASE}/shareObject")),
        r#"{"includeAnonymousLinkInEmail":false,"peoplePickerInput":"[{\"Key\":\"user@contoso.com\"}]","propagateAcl":false,"roleValue":"role: 3","useSimplifiedRoles":true}"#,
    );
}

#[tokio::test]
async fn folder_get_share_link() {
    let transport = RecordingTransport::default().queue(r#"{"sharingLinkInfo":{"IsActive":true}}"#);
    let client = Client::new(&transport);

    sharing::folder_get_share_link(
        &client,
        SITE,
        "/sites/team/Shared Documents",
        SharingLinkKind::Direct,
        None,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FOLDER_BASE}/shareLink")),
        r#"{"request":{"createLink":true,"emailData":null,"settings":{"expiration":null,"linkKind":1}}}"#,
    );
}

#[tokio::test]
async fn folder_check_permissions() {
    let transport =
        RecordingTransport::default().queue(r#"[{"inputEntity":"user@contoso.com","role":2}]"#);
    let client = Client::new(&transport);

    sharing::folder_check_permissions(
        &client,
        SITE,
        "/sites/team/Shared Documents",
        &[SharingRecipient {
            email: None,
            alias: Some("user".into()),
        }],
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FOLDER_BASE}/checkPermissions")),
        r#"{"recipients":[{"alias":"user"}]}"#,
    );
}

#[tokio::test]
async fn folder_get_sharing_information() {
    let transport = RecordingTransport::default()
        .queue(r#"{"sharedObjectType":2,"directUrl":"/sites/team/Shared Documents"}"#);
    let client = Client::new(&transport);

    sharing::folder_get_sharing_information(
        &client,
        SITE,
        "/sites/team/Shared Documents",
        None,
        &[],
        &[],
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        (
            "POST",
            &format!("{FOLDER_BASE}/getSharingInformation?$select=*"),
        ),
        r#"{"request":null}"#,
    );
}

#[tokio::test]
async fn folder_get_object_sharing_settings() {
    let transport = RecordingTransport::default().queue(r#"{"CanShareFolder":true}"#);
    let client = Client::new(&transport);

    sharing::folder_get_object_sharing_settings(
        &client,
        SITE,
        "/sites/team/Shared Documents",
        false,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FOLDER_BASE}/getObjectSharingSettings")),
        r#"{"useSimplifiedRoles":false}"#,
    );
}

#[tokio::test]
async fn folder_unshare() {
    let transport = RecordingTransport::default().queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    sharing::folder_unshare(&client, SITE, "/sites/team/Shared Documents")
        .await
        .unwrap();

    transport.assert_last(("POST", &format!("{FOLDER_BASE}/unshareObject")));
}

#[tokio::test]
async fn folder_delete_sharing_link_by_kind() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    sharing::folder_delete_sharing_link_by_kind(
        &client,
        SITE,
        "/sites/team/Shared Documents",
        SharingLinkKind::AnonymousEdit,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FOLDER_BASE}/deleteLinkByKind")),
        r#"{"linkKind":5}"#,
    );
}

#[tokio::test]
async fn folder_unshare_link() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    sharing::folder_unshare_link(
        &client,
        SITE,
        "/sites/team/Shared Documents",
        SharingLinkKind::OrganizationView,
        "",
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{FOLDER_BASE}/unshareLink")),
        &format!(r#"{{"linkKind":2,"shareId":"{EMPTY_GUID}"}}"#),
    );
}

const ACLINV_URL: &str =
    "https://contoso.sharepoint.com/sites/team/_layouts/15/aclinv.aspx?forSharing=1&mbypass=1";

#[tokio::test]
async fn web_share_with_shares_aclinv_page() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"results":[{"Id":3}]}}"#)
        .queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    sharing::web_share_with(
        &client,
        SITE,
        &["user@contoso.com".to_string()],
        SharingRole::View,
        None,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(
        (calls[0].method.as_str(), calls[0].url.as_str()),
        (
            "GET",
            "https://contoso.sharepoint.com/sites/team/_api/web/roledefinitions?$select=Id&$top=1&$filter=RoleTypeKind eq 2"
        )
    );
    let sent = String::from_utf8(calls[1].body.clone().unwrap_or_default()).unwrap();
    assert_eq!(
        (calls[1].method.as_str(), calls[1].url.as_str()),
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/SP.Web.ShareObject?$expand=UsersWithAccessRequests,GroupsSharedWith"
        )
    );
    assert_eq!(
        sent,
        format!(
            r#"{{"peoplePickerInput":"[{{\"Key\":\"user@contoso.com\"}}]","roleValue":"role: 3","url":"{ACLINV_URL}"}}"#
        )
    );
}

#[tokio::test]
async fn web_share_object_with_email_subject_present() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"results":[{"Id":4}]}}"#)
        .queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    let email = sharing::types::SharingEmailData {
        subject: Some("Docs".into()),
        body: "See below".into(),
    };
    sharing::web_share_object(
        &client,
        SITE,
        &format!("{SITE}/Shared Documents/report.docx"),
        &["user@contoso.com".to_string()],
        SharingRole::Edit,
        None,
        Some(&email),
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/SP.Web.ShareObject?$expand=UsersWithAccessRequests,GroupsSharedWith",
        ),
        r#"{"emailBody":"See below","emailSubject":"Docs","peoplePickerInput":"[{\"Key\":\"user@contoso.com\"}]","roleValue":"role: 4","sendEmail":true,"url":"https://contoso.sharepoint.com/sites/team/Shared Documents/report.docx"}"#,
    );
}

#[tokio::test]
async fn web_share_object_email_subject_defaults_when_absent() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"results":[{"Id":4}]}}"#)
        .queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    let email = sharing::types::SharingEmailData {
        subject: None,
        body: "See below".into(),
    };
    sharing::web_share_object(
        &client,
        SITE,
        &format!("{SITE}/Shared Documents/report.docx"),
        &["user@contoso.com".to_string()],
        SharingRole::Edit,
        None,
        Some(&email),
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/SP.Web.ShareObject?$expand=UsersWithAccessRequests,GroupsSharedWith",
        ),
        r#"{"emailBody":"See below","emailSubject":"Shared with you.","peoplePickerInput":"[{\"Key\":\"user@contoso.com\"}]","roleValue":"role: 4","sendEmail":true,"url":"https://contoso.sharepoint.com/sites/team/Shared Documents/report.docx"}"#,
    );
}

#[tokio::test]
async fn web_share_object_with_group_uses_associated_member_group() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"Id":7}}"#)
        .queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    sharing::web_share_object(
        &client,
        SITE,
        &format!("{SITE}/Shared Documents/report.docx"),
        &["user@contoso.com".to_string()],
        SharingRole::Edit,
        Some(RoleType::Contributor),
        None,
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/SP.Web.ShareObject?$expand=UsersWithAccessRequests,GroupsSharedWith",
        ),
        r#"{"peoplePickerInput":"[{\"Key\":\"user@contoso.com\"}]","roleValue":"group: 7","url":"https://contoso.sharepoint.com/sites/team/Shared Documents/report.docx"}"#,
    );
}

#[tokio::test]
async fn web_share_object_raw_sends_options_untouched() {
    let transport = RecordingTransport::default().queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    let options = ShareObjectOptions {
        url: Some(format!("{SITE}/Shared Documents/report.docx")),
        login_names: Some(r#"[{"Key":"user@contoso.com"}]"#.into()),
        role: Some(SharingRole::View),
        email_data: None,
        group: None,
        propagate_acl: Some(true),
        include_anonymous_link_in_email: None,
        use_simplified_roles: Some(false),
    };
    sharing::web_share_object_raw(&client, SITE, &options)
        .await
        .unwrap();

    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/SP.Web.ShareObject?$expand=UsersWithAccessRequests,GroupsSharedWith",
        ),
        r#"{"url":"https://contoso.sharepoint.com/sites/team/Shared Documents/report.docx","loginNames":"[{\"Key\":\"user@contoso.com\"}]","role":1,"propagateAcl":true,"useSimplifiedRoles":false}"#,
    );
}

#[tokio::test]
async fn web_unshare_object() {
    let transport = RecordingTransport::default().queue(r#"{"StatusCode":0}"#);
    let client = Client::new(&transport);

    sharing::web_unshare_object(
        &client,
        SITE,
        &format!("{SITE}/Shared Documents/report.docx"),
    )
    .await
    .unwrap();

    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/web/unshareObject",
        ),
        r#"{"url":"https://contoso.sharepoint.com/sites/team/Shared Documents/report.docx"}"#,
    );
}

#[tokio::test]
async fn role_lookup_error_propagates_from_group() {
    let transport = RecordingTransport::default().queue(r#"{"d":{"Id":7}}"#);
    let client = Client::new(&transport);

    let err = sharing::web_share_object(
        &client,
        SITE,
        &format!("{SITE}/Shared Documents/report.docx"),
        &["user@contoso.com".to_string()],
        SharingRole::Edit,
        Some(RoleType::Administrator),
        None,
    )
    .await
    .unwrap_err();

    assert!(matches!(
        err,
        Error::Surface(_, msg) if msg.contains("Could not determine role value")
    ));
}
