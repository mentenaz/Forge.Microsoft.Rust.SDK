use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_clientside_pages::{
    copy_client_side_page, create_client_side_page, save_client_side_page, save_page_as_template,
    ColumnControl, Page, TextData,
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
    fn with_url_matches(responses: Vec<(&'static str, &'static str)>) -> Self {
        Self {
            responses,
            default_response: Mutex::new(None),
            calls: Mutex::new(Vec::new()),
        }
    }

    fn calls_matching(&self, needle: &str) -> Vec<Call> {
        self.calls
            .lock()
            .unwrap()
            .iter()
            .filter(|c| c.url.contains(needle))
            .cloned()
            .collect()
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

fn page_json(id: i64, title: &str) -> String {
    page_json_checked_out(id, title, true)
}

fn page_json_checked_out(id: i64, title: &str, checked_out: bool) -> String {
    format!(
        r#"{{
        "AbsoluteUrl": "https://contoso.sharepoint.com/sites/team/SitePages/{title}.aspx",
        "AuthorByline": null,
        "BannerImageUrl": "",
        "BannerThumbnailUrl": "",
        "ContentTypeId": null,
        "Description": "",
        "DoesUserHaveEditPermission": true,
        "FileName": "{title}.aspx",
        "FirstPublished": "",
        "Id": {id},
        "IsPageCheckedOutToCurrentUser": {checked_out},
        "IsWebWelcomePage": false,
        "Modified": "",
        "PageLayoutType": "Article",
        "Path": {{ "DecodedUrl": "https://contoso.sharepoint.com/sites/team/SitePages/{title}.aspx" }},
        "PromotedState": 0,
        "Title": "{title}",
        "TopicHeader": null,
        "UniqueId": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        "Url": "SitePages/{title}.aspx",
        "Version": "1.0",
        "VersionInfo": {{ "LastVersionCreated": "", "LastVersionCreatedBy": "" }},
        "AlternativeUrlMap": "",
        "CanvasContent1": "[]",
        "LayoutWebpartsContent": "null"
    }}"#
    )
}

async fn dummy_page(id: i64, title: &str) -> forge_m365_sp_clientside_pages::ClientsidePage {
    let json = page_json(id, title);
    let transport = SyncOnceTransport(json);
    let client = Client::new(&transport);
    forge_m365_sp_clientside_pages::get_client_side_page(&client, SITE, id)
        .await
        .unwrap()
}

struct SyncOnceTransport(String);
impl Transport for SyncOnceTransport {
    fn execute(
        &self,
        _surface: Surface,
        _method: &str,
        _url: &str,
        _headers: &[(&str, &str)],
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        let body = self.0.clone();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

#[tokio::test]
async fn save_client_side_page_posts_body_then_reloads() {
    let saved_json = page_json(42, "home");
    let responses: Vec<(&'static str, &'static str)> = vec![
        ("savepage", ""),
        (
            "sitepages/pages(42)",
            Box::leak(saved_json.into_boxed_str()),
        ),
    ];
    let transport = RecordingTransport::with_url_matches(responses);
    let client = Client::new(&transport);
    let mut page = dummy_page(42, "home").await;
    page.content.push_dummy_section_for_test();

    let published = save_client_side_page(&client, SITE, &mut page, false)
        .await
        .unwrap();

    assert!(published);
    let save_call = transport.calls_matching("savepage");
    assert_eq!(save_call.len(), 1);
    assert_eq!(save_call[0].method, "POST");
    let body: serde_json::Value =
        serde_json::from_slice(save_call[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(body["Title"], "home");
    assert!(body["CanvasContent1"].as_str().unwrap().contains("hello"));

    // reloaded from the second (page-data) response
    assert_eq!(page.metadata.id, 42);
    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls.len(), 2, "save then reload");
}

#[tokio::test]
async fn save_client_side_page_publishes_when_requested() {
    // Post-publish, a real server reports the page as checked back in.
    let saved_json = page_json_checked_out(42, "home", false);
    let transport = RecordingTransport::with_url_matches(vec![
        ("savepage", ""),
        ("publish", "true"),
        (
            "sitepages/pages(42)",
            Box::leak(saved_json.into_boxed_str()),
        ),
    ]);
    let client = Client::new(&transport);
    let mut page = dummy_page(42, "home").await;

    let published = save_client_side_page(&client, SITE, &mut page, true)
        .await
        .unwrap();

    assert!(published);
    assert_eq!(transport.calls_matching("publish").len(), 1);
    assert!(!page.metadata.is_page_checked_out_to_current_user);
}

#[tokio::test]
async fn create_client_side_page_posts_then_saves_then_sets_final_title() {
    let created_json = page_json(99, "newpage");
    let saved_json = page_json(99, "newpage");
    // Order matters: RecordingTransport matches the first needle found as a
    // substring, and the bare creation URL ("…/_api/sitepages/pages") is
    // itself a substring of every subsequent action URL
    // ("…/_api/sitepages/pages(99)/savepage" etc.), so the more specific
    // needles must be listed first.
    let transport = RecordingTransport::with_url_matches(vec![
        ("savepage", ""),
        (
            "sitepages/pages(99)",
            Box::leak(saved_json.into_boxed_str()),
        ),
        (
            "_api/sitepages/pages",
            Box::leak(created_json.into_boxed_str()),
        ),
    ]);
    let client = Client::new(&transport);

    let page = create_client_side_page(
        &client,
        SITE,
        "newpage.aspx",
        "New Page Title",
        "Article",
        0,
    )
    .await
    .unwrap();

    // the odd PnPjs-parity behavior: final in-memory title is the caller's
    // requested title, even though only "newpage" (the filename) was ever
    // actually saved to the server
    assert_eq!(page.metadata.title, "New Page Title");

    let create_calls = transport.calls_matching("_api/sitepages/pages");
    let create_call = create_calls
        .iter()
        .find(|c| c.method == "POST" && !c.url.contains("savepage"))
        .unwrap();
    let body: serde_json::Value =
        serde_json::from_slice(create_call.body.as_ref().unwrap()).unwrap();
    assert_eq!(body["PageLayoutType"], "Article");

    let save_calls = transport.calls_matching("savepage");
    let save_body: serde_json::Value =
        serde_json::from_slice(save_calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(
        save_body["Title"], "newpage",
        "saved with the filename-derived title, not the final one"
    );
}

#[tokio::test]
async fn save_page_as_template_posts_then_saves() {
    let template_json = page_json(100, "template-of-home");
    // Realistic reload fixture: save_page_as_template sets the new page's
    // title to the source's title *before* saving (unlike
    // create_client_side_page's after-save quirk), so a real server's
    // post-save state reflects "home", not the server-generated
    // "template-of-home" placeholder from the initial SavePageAsTemplate response.
    let saved_json = page_json(100, "home");
    let transport = RecordingTransport::with_url_matches(vec![
        (
            "SavePageAsTemplate",
            Box::leak(template_json.into_boxed_str()),
        ),
        ("savepage", ""),
        ("publish", "true"),
        (
            "sitepages/pages(100)",
            Box::leak(saved_json.into_boxed_str()),
        ),
    ]);
    let client = Client::new(&transport);
    let source = dummy_page(42, "home").await;

    let template = save_page_as_template(&client, SITE, &source, true)
        .await
        .unwrap();

    assert_eq!(
        template.metadata.title, "home",
        "template inherits the source page's title"
    );
    assert_eq!(transport.calls_matching("SavePageAsTemplate").len(), 1);
}

#[tokio::test]
async fn copy_client_side_page_creates_then_copies_content_then_saves() {
    let created_json = page_json(200, "copy-of-home");
    let saved_json = page_json(200, "copy-of-home");
    // See the ordering note in create_client_side_page_posts_then_saves_then_sets_final_title.
    let transport = RecordingTransport::with_url_matches(vec![
        ("savepage", ""),
        ("publish", "true"),
        (
            "sitepages/pages(200)",
            Box::leak(saved_json.into_boxed_str()),
        ),
        (
            "_api/sitepages/pages",
            Box::leak(created_json.into_boxed_str()),
        ),
    ]);
    let client = Client::new(&transport);

    let mut source = dummy_page(42, "home").await;
    source.content.push_dummy_section_for_test();
    source.set_topic_header("Announcements");

    let copy = copy_client_side_page(
        &client,
        SITE,
        &source,
        "copy-of-home",
        "Copy of Home",
        true,
        0,
    )
    .await
    .unwrap();

    // `copy`'s in-memory content reflects whatever the final reload
    // returned (the mock fixture's, here) — the real assertion for
    // "did copy_page_content_to actually transfer the canvas" is on the
    // save request body it sent, below.
    let _ = &copy;
    // Two savepage calls happen: one inside create_client_side_page (empty
    // content, for the freshly-created shell page) and one inside
    // copy_page_content_to (the actual content copy) — the latter is last.
    let save_calls = transport.calls_matching("savepage");
    assert_eq!(save_calls.len(), 2);
    let save_body: serde_json::Value =
        serde_json::from_slice(save_calls[1].body.as_ref().unwrap()).unwrap();
    assert!(save_body["CanvasContent1"]
        .as_str()
        .unwrap()
        .contains("hello"));
    assert_eq!(save_body["TopicHeader"], "Announcements");
}

/// Test-only helper trait so the write-ops tests can cheaply build a page
/// with one non-empty section without duplicating canvas-construction
/// boilerplate already covered by `tests/canvas.rs`.
trait PushDummySection {
    fn push_dummy_section_for_test(&mut self);
}

impl PushDummySection for Page {
    fn push_dummy_section_for_test(&mut self) {
        let idx = self.add_section();
        self.sections[idx].add_control(ColumnControl::Text(TextData::new("hello")));
    }
}
