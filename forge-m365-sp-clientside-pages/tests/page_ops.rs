use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_clientside_pages::{
    get_client_side_page, get_client_side_web_parts, get_page_item_ref, load_client_side_page,
};
use std::sync::Mutex;

/// Records every request made and can be configured to return different
/// canned responses depending on the URL, so `load_client_side_page`'s two
/// sequential requests (item ref, then page data) can be exercised in one
/// test.
#[derive(Default)]
struct RecordingTransport {
    responses: Vec<(&'static str, &'static str)>,
    default_response: Mutex<Option<String>>,
    urls: Mutex<Vec<String>>,
}

impl RecordingTransport {
    fn with_response(response: &'static str) -> Self {
        Self {
            responses: Vec::new(),
            default_response: Mutex::new(Some(response.to_string())),
            urls: Mutex::new(Vec::new()),
        }
    }

    fn with_url_matches(responses: Vec<(&'static str, &'static str)>) -> Self {
        Self {
            responses,
            default_response: Mutex::new(None),
            urls: Mutex::new(Vec::new()),
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
        self.urls.lock().unwrap().push(url.to_string());
        let body = self
            .responses
            .iter()
            .find(|(needle, _)| url.contains(needle))
            .map(|(_, resp)| resp.to_string())
            .or_else(|| self.default_response.lock().unwrap().clone())
            .unwrap_or_default();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

const PAGE_JSON: &str = r#"{
    "AbsoluteUrl": "https://contoso.sharepoint.com/sites/team/SitePages/home.aspx",
    "AuthorByline": ["i:0#.f|membership|user@contoso.com"],
    "BannerImageUrl": "https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png",
    "BannerThumbnailUrl": "",
    "ContentTypeId": null,
    "Description": "A test page",
    "DoesUserHaveEditPermission": true,
    "FileName": "home.aspx",
    "FirstPublished": "2024-01-01T00:00:00Z",
    "Id": 42,
    "IsPageCheckedOutToCurrentUser": false,
    "IsWebWelcomePage": true,
    "Modified": "2024-02-01T00:00:00Z",
    "PageLayoutType": "Home",
    "Path": { "DecodedUrl": "https://contoso.sharepoint.com/sites/team/SitePages/home.aspx" },
    "PromotedState": 0,
    "Title": "Home",
    "TopicHeader": null,
    "UniqueId": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
    "Url": "SitePages/home.aspx",
    "Version": "3.0",
    "VersionInfo": { "LastVersionCreated": "2024-02-01T00:00:00Z", "LastVersionCreatedBy": "user@contoso.com" },
    "AlternativeUrlMap": "",
    "CanvasContent1": "[]",
    "LayoutWebpartsContent": "null"
}"#;

const ITEM_REF_JSON: &str = r#"{"Id": 42, "CommentsDisabled": true}"#;

#[tokio::test]
async fn get_client_side_page_parses_metadata_and_content() {
    let transport = RecordingTransport::with_response(PAGE_JSON);
    let client = Client::new(&transport);

    let page = get_client_side_page(&client, SITE, 42).await.unwrap();

    assert_eq!(page.metadata.id, 42);
    assert_eq!(page.metadata.title, "Home");
    assert_eq!(page.metadata.page_layout_type, "Home");
    assert_eq!(page.metadata.content_type_id, None);
    assert_eq!(
        page.metadata.author_byline,
        Some(vec!["i:0#.f|membership|user@contoso.com".to_string()])
    );
    assert!(page.content.sections.is_empty());
    assert_eq!(page.layout.properties.layout_type, "FullWidthImage");
    assert_eq!(page.comments_disabled, None);

    let url = transport.urls.lock().unwrap().last().unwrap().clone();
    assert_eq!(url, format!("{SITE}/_api/sitepages/pages(42)"));
}

#[tokio::test]
async fn get_page_item_ref_builds_percent_encoded_url() {
    let transport = RecordingTransport::with_response(ITEM_REF_JSON);
    let client = Client::new(&transport);

    let item_ref = get_page_item_ref(&client, SITE, "/sites/team/Site Pages/home.aspx")
        .await
        .unwrap();

    assert_eq!(item_ref.id, 42);
    assert!(item_ref.comments_disabled);

    let url = transport.urls.lock().unwrap().last().unwrap().clone();
    assert!(url.contains(
        "getFileByServerRelativePath(decodedUrl='%2Fsites%2Fteam%2FSite%20Pages%2Fhome%2Easpx')"
    ));
    assert!(url.ends_with("/ListItemAllFields?$select=Id,CommentsDisabled"));
}

#[tokio::test]
async fn load_client_side_page_composes_item_ref_then_page_fetch() {
    let transport = RecordingTransport::with_url_matches(vec![
        ("ListItemAllFields", ITEM_REF_JSON),
        ("sitepages/pages(42)", PAGE_JSON),
    ]);
    let client = Client::new(&transport);

    let page = load_client_side_page(&client, SITE, "/sites/team/SitePages/home.aspx")
        .await
        .unwrap();

    assert_eq!(page.metadata.id, 42);
    assert_eq!(page.comments_disabled, Some(true));

    let urls = transport.urls.lock().unwrap().clone();
    assert_eq!(urls.len(), 2, "item ref then page data, in that order");
    assert!(urls[0].contains("ListItemAllFields"));
    assert!(urls[1].contains("sitepages/pages(42)"));
}

#[tokio::test]
async fn get_client_side_web_parts_parses_manifest_array() {
    let json = r#"[{
        "ComponentType": 1,
        "Id": "cccccccc-cccc-cccc-cccc-cccccccccccc",
        "Manifest": "{\"preconfiguredEntries\":[]}",
        "ManifestType": 1,
        "Name": "TestWebPart",
        "Status": 0
    }]"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let parts = get_client_side_web_parts(&client, SITE).await.unwrap();

    assert_eq!(parts.len(), 1);
    assert_eq!(parts[0].name, "TestWebPart");
    let url = transport.urls.lock().unwrap().last().unwrap().clone();
    assert_eq!(url, format!("{SITE}/_api/web/GetClientSideWebParts"));
}

#[tokio::test]
async fn parses_full_canvas_content_when_non_empty() {
    let page_json = PAGE_JSON.replace(
        r#""CanvasContent1": "[]""#,
        r#""CanvasContent1": "[{\"controlType\":4,\"displayMode\":2,\"emphasis\":{},\"id\":\"bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb\",\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"controlIndex\":1,\"sectionFactor\":12,\"layoutIndex\":1},\"anchorComponentId\":\"bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb\",\"editorType\":\"CKEditor\",\"addedFromPersistedData\":true,\"innerHTML\":\"<p>hi</p>\"},null]""#,
    );
    let transport = RecordingTransport::with_response(Box::leak(page_json.into_boxed_str()));
    let client = Client::new(&transport);

    let page = get_client_side_page(&client, SITE, 42).await.unwrap();

    assert_eq!(page.content.sections.len(), 1);
    assert_eq!(page.content.sections[0].columns[0].controls.len(), 1);
}
