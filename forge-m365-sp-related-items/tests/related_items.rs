use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_related_items::{
    add_single_link, add_single_link_from_url, add_single_link_to_url, delete_single_link,
    get_page_one_related_items, get_related_items,
};
use serde_json::Value;
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
    last_method: Mutex<Option<String>>,
    last_body: Mutex<Option<Vec<u8>>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            ..Default::default()
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
        *self.last_body.lock().unwrap() = body.map(|b| b.to_vec());
        let resp = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(resp.into_bytes()) })
    }
}

const ITEM_JSON: &str = r#"{"ListId":"aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa","ItemId":1,"Url":"/sites/team/Lists/L/1_.000","Title":"Item 1","WebId":"bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb","IconUrl":"icgen.gif"}"#;
const SITE: &str = "https://contoso.sharepoint.com/sites/team";

#[tokio::test]
async fn get_related_items_posts_correct_body() {
    let transport = RecordingTransport::with_response(&format!("[{ITEM_JSON}]"));
    let client = Client::new(&transport);

    let items = get_related_items(&client, SITE, "Tasks", 1).await.unwrap();
    assert_eq!(items.len(), 1);
    assert_eq!(items[0].title, "Item 1");

    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!("{SITE}/_api/SP.RelatedItemManager.GetRelatedItems")
    );
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let body: Value =
        serde_json::from_slice(&transport.last_body.lock().unwrap().clone().unwrap()).unwrap();
    assert_eq!(body["SourceListName"], "Tasks");
    assert_eq!(body["SourceItemID"], 1);
}

#[tokio::test]
async fn get_page_one_related_items_uses_correct_url() {
    let transport = RecordingTransport::with_response("[]");
    let client = Client::new(&transport);

    get_page_one_related_items(&client, SITE, "Tasks", 1)
        .await
        .unwrap();

    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!("{SITE}/_api/SP.RelatedItemManager.GetPageOneRelatedItems")
    );
}

#[tokio::test]
async fn add_single_link_sends_full_spec() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    add_single_link(
        &client,
        SITE,
        "Tasks",
        1,
        "https://contoso.sharepoint.com/sites/team",
        "Docs",
        2,
        "https://contoso.sharepoint.com/sites/other",
        true,
    )
    .await
    .unwrap();

    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!("{SITE}/_api/SP.RelatedItemManager.AddSingleLink")
    );
    let body: Value =
        serde_json::from_slice(&transport.last_body.lock().unwrap().clone().unwrap()).unwrap();
    assert_eq!(body["TargetListName"], "Docs");
    assert_eq!(body["TryAddReverseLink"], true);
}

#[tokio::test]
async fn add_single_link_to_url_and_from_url_build_correct_urls() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    add_single_link_to_url(
        &client,
        SITE,
        "Tasks",
        1,
        "https://contoso.sharepoint.com/x",
        false,
    )
    .await
    .unwrap();
    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!("{SITE}/_api/SP.RelatedItemManager.AddSingleLinkToUrl")
    );

    add_single_link_from_url(
        &client,
        SITE,
        "https://contoso.sharepoint.com/x",
        "Docs",
        2,
        false,
    )
    .await
    .unwrap();
    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!("{SITE}/_api/SP.RelatedItemManager.AddSingleLinkFromUrl")
    );
}

#[tokio::test]
async fn delete_single_link_sends_full_spec() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);

    delete_single_link(
        &client,
        SITE,
        "Tasks",
        1,
        "https://contoso.sharepoint.com/sites/team",
        "Docs",
        2,
        "https://contoso.sharepoint.com/sites/other",
        false,
    )
    .await
    .unwrap();

    assert_eq!(
        transport.last_url.lock().unwrap().clone().unwrap(),
        format!("{SITE}/_api/SP.RelatedItemManager.DeleteSingleLink")
    );
    let body: Value =
        serde_json::from_slice(&transport.last_body.lock().unwrap().clone().unwrap()).unwrap();
    assert_eq!(body["TryDeleteReverseLink"], false);
}
