use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_clientside_pages::{
    discard_page_checkout, promote_to_news, recycle_page, schedule_publish, set_author_by_id,
    set_author_by_login_name, set_banner_image_from_external_url, share_page_by_email,
    BannerImageProps, ClientsidePage,
};
use std::sync::Mutex;

#[derive(Debug, Clone)]
struct Call {
    method: String,
    url: String,
    body: Option<Vec<u8>>,
}

#[derive(Default)]
struct RecordingTransport {
    responses: Vec<(&'static str, &'static str)>,
    default_response: Mutex<Option<String>>,
    calls: Mutex<Vec<Call>>,
}

impl RecordingTransport {
    fn with_response(response: &'static str) -> Self {
        Self {
            responses: Vec::new(),
            default_response: Mutex::new(Some(response.to_string())),
            calls: Mutex::new(Vec::new()),
        }
    }

    fn with_url_matches(responses: Vec<(&'static str, &'static str)>) -> Self {
        Self {
            responses,
            default_response: Mutex::new(None),
            calls: Mutex::new(Vec::new()),
        }
    }

    fn last(&self) -> Call {
        self.calls.lock().unwrap().last().unwrap().clone()
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
        self.calls.lock().unwrap().push(Call {
            method: method.to_string(),
            url: url.to_string(),
            body: body.map(|b| b.to_vec()),
        });
        let response = self
            .responses
            .iter()
            .find(|(needle, _)| url.contains(needle))
            .map(|(_, resp)| resp.to_string())
            .or_else(|| self.default_response.lock().unwrap().clone())
            .unwrap_or_default();
        Box::pin(async move { Ok(response.into_bytes()) })
    }
}

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

const PAGE_JSON: &str = r#"{
    "AbsoluteUrl": "https://contoso.sharepoint.com/sites/team/SitePages/home.aspx",
    "AuthorByline": null,
    "BannerImageUrl": "",
    "BannerThumbnailUrl": "",
    "ContentTypeId": null,
    "Description": "",
    "DoesUserHaveEditPermission": true,
    "FileName": "home.aspx",
    "FirstPublished": "",
    "Id": 42,
    "IsPageCheckedOutToCurrentUser": true,
    "IsWebWelcomePage": false,
    "Modified": "",
    "PageLayoutType": "Article",
    "Path": { "DecodedUrl": "https://contoso.sharepoint.com/sites/team/SitePages/home.aspx" },
    "PromotedState": 0,
    "Title": "Home",
    "TopicHeader": null,
    "UniqueId": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
    "Url": "SitePages/home.aspx",
    "Version": "1.0",
    "VersionInfo": { "LastVersionCreated": "", "LastVersionCreatedBy": "" },
    "AlternativeUrlMap": "",
    "CanvasContent1": "[]",
    "LayoutWebpartsContent": "null"
}"#;

async fn dummy_page() -> ClientsidePage {
    // Built from the same fixture used for get_client_side_page in
    // page_ops.rs, via the public parse path so it stays in sync with the
    // real wire shape.
    let transport = SyncOnceTransport(PAGE_JSON);
    let client = Client::new(&transport);
    forge_m365_sp_clientside_pages::get_client_side_page(&client, SITE, 42)
        .await
        .unwrap()
}

struct SyncOnceTransport(&'static str);
impl Transport for SyncOnceTransport {
    fn execute(
        &self,
        _surface: Surface,
        _method: &str,
        _url: &str,
        _headers: &[(&str, &str)],
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        let body = self.0.to_string();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

#[tokio::test]
async fn discard_page_checkout_posts_and_reparses() {
    let transport = RecordingTransport::with_response(PAGE_JSON);
    let client = Client::new(&transport);

    let page = discard_page_checkout(&client, SITE, 42).await.unwrap();

    assert_eq!(page.metadata.id, 42);
    let call = transport.last();
    assert_eq!(call.method, "POST");
    assert!(call.url.ends_with("/_api/sitepages/pages(42)/discardPage"));
}

#[tokio::test]
async fn promote_to_news_parses_bool() {
    let transport = RecordingTransport::with_response("true");
    let client = Client::new(&transport);

    let promoted = promote_to_news(&client, SITE, 42).await.unwrap();

    assert!(promoted);
    let call = transport.last();
    assert!(call.url.ends_with("/promoteToNews"));
}

#[tokio::test]
async fn schedule_publish_sends_wrapped_date_and_parses_version() {
    let transport = RecordingTransport::with_response("\"2.0\"");
    let client = Client::new(&transport);

    let version = schedule_publish(&client, SITE, 42, "2024-06-01T00:00:00Z")
        .await
        .unwrap();

    assert_eq!(version, "2.0");
    let call = transport.last();
    assert!(call.url.ends_with("/SchedulePublish"));
    let body: serde_json::Value = serde_json::from_slice(&call.body.unwrap()).unwrap();
    assert_eq!(body["sitePage"]["PublishStartDate"], "2024-06-01T00:00:00Z");
}

#[tokio::test]
async fn share_page_by_email_sends_expected_body() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    share_page_by_email(
        &client,
        SITE,
        "https://contoso.sharepoint.com/sites/team/SitePages/home.aspx",
        &["a@contoso.com".to_string(), "b@contoso.com".to_string()],
        "check this out",
    )
    .await
    .unwrap();

    let call = transport.last();
    assert!(call
        .url
        .ends_with("/_api/SP.Publishing.RichSharing/SharePageByEmail"));
    let body: serde_json::Value = serde_json::from_slice(&call.body.unwrap()).unwrap();
    assert_eq!(body["message"], "check this out");
    assert_eq!(body["recipientEmails"][0], "a@contoso.com");
}

#[tokio::test]
async fn set_author_by_login_name_updates_page_state() {
    let transport = RecordingTransport::with_response(
        r#"[{"UserPrincipalName":"user@contoso.com","Title":"A User"}]"#,
    );
    let client = Client::new(&transport);
    let mut page = dummy_page().await;

    set_author_by_login_name(
        &client,
        SITE,
        &mut page,
        "i:0#.f|membership|user@contoso.com",
    )
    .await
    .unwrap();

    assert_eq!(
        page.metadata.author_byline,
        Some(vec!["user@contoso.com".to_string()])
    );
    assert_eq!(page.layout.properties.authors[0].name, "A User");
    let call = transport.last();
    assert!(call.url.contains("LoginName eq"));
}

#[tokio::test]
async fn set_author_by_id_resolves_login_name_then_profile() {
    let transport = RecordingTransport::with_url_matches(vec![
        (
            "Id eq 7",
            r#"[{"LoginName":"i:0#.f|membership|user@contoso.com"}]"#,
        ),
        (
            "LoginName eq",
            r#"[{"UserPrincipalName":"user@contoso.com","Title":"A User"}]"#,
        ),
    ]);
    let client = Client::new(&transport);
    let mut page = dummy_page().await;

    set_author_by_id(&client, SITE, &mut page, 7).await.unwrap();

    assert_eq!(
        page.metadata.author_byline,
        Some(vec!["user@contoso.com".to_string()])
    );
    assert_eq!(transport.calls.lock().unwrap().len(), 2);
}

#[tokio::test]
async fn set_author_by_login_name_errors_when_not_found() {
    let transport = RecordingTransport::with_response("[]");
    let client = Client::new(&transport);
    let mut page = dummy_page().await;

    let err = set_author_by_login_name(&client, SITE, &mut page, "nobody")
        .await
        .unwrap_err();

    assert!(err.to_string().contains("no site user"));
}

#[tokio::test]
async fn set_banner_image_from_external_url_uploads_and_sets_local_state() {
    let transport = RecordingTransport::with_response(
        r#"{"ServerRelativeUrl":"/sites/team/SiteAssets/SitePages/home/banner.png"}"#,
    );
    let client = Client::new(&transport);
    let mut page = dummy_page().await;

    set_banner_image_from_external_url(
        &client,
        SITE,
        &mut page,
        "https://example.com/images/banner.png",
        Some(BannerImageProps {
            alt_text: Some("A banner".to_string()),
            ..Default::default()
        }),
    )
    .await
    .unwrap();

    assert_eq!(
        page.metadata.banner_image_url,
        "/sites/team/SiteAssets/SitePages/home/banner.png"
    );
    assert_eq!(
        page.layout.properties.alt_text,
        Some("A banner".to_string())
    );
    let call = transport.last();
    assert!(call.url.contains("AddImageFromExternalUrl"));
    assert!(call.url.contains("pageName=%27home%27"));
}

#[tokio::test]
async fn set_banner_image_relativizes_absolute_url() {
    let mut page = dummy_page().await;

    page.set_banner_image(
        "https://contoso.sharepoint.com/sites/team/SiteAssets/banner.png",
        None,
    );

    assert_eq!(
        page.metadata.banner_image_url,
        "/sites/team/SiteAssets/banner.png"
    );
    assert_eq!(page.layout.properties.image_source_type, Some(2));
}

#[tokio::test]
async fn recycle_page_posts_to_recycle_endpoint() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    recycle_page(&client, SITE, 42).await.unwrap();

    let call = transport.last();
    assert_eq!(call.method, "POST");
    assert!(call
        .url
        .ends_with("/_api/web/lists/getbytitle('Site Pages')/items(42)/recycle"));
}

#[tokio::test]
async fn delete_page_sends_delete_with_if_match() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    forge_m365_sp_clientside_pages::delete_page(&client, SITE, 42)
        .await
        .unwrap();

    let call = transport.last();
    assert_eq!(call.method, "DELETE");
    assert!(call
        .url
        .ends_with("/_api/web/lists/getbytitle('Site Pages')/items(42)"));
}
