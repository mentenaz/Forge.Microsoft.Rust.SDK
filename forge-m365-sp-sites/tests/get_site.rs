use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_sites::{
    get_document_libraries, get_root_web_url, get_site, get_web_url_from_page_url,
};

struct FixedResponse(&'static str);

impl Transport for FixedResponse {
    fn execute(
        &self,
        _surface: Surface,
        _method: &str,
        _url: &str,
        _headers: &[(&str, &str)],
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        Box::pin(async move { Ok(self.0.as_bytes().to_vec()) })
    }
}

#[tokio::test]
async fn parses_site_info() {
    let json = r#"{"Id":"11111111-1111-1111-1111-111111111111","Url":"https://contoso.sharepoint.com/sites/team","ServerRelativeUrl":"/sites/team","Classification":"","IsHubSite":false,"HubSiteId":"00000000-0000-0000-0000-000000000000"}"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let site = get_site(&client, "https://contoso.sharepoint.com/sites/team/")
        .await
        .unwrap();

    assert_eq!(site.id, "11111111-1111-1111-1111-111111111111");
    assert_eq!(site.server_relative_url, "/sites/team");
    assert!(!site.is_hub_site);
}

#[tokio::test]
async fn parses_root_web_url() {
    let json = r#"{"Url":"https://contoso.sharepoint.com/sites/team"}"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let url = get_root_web_url(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(url, "https://contoso.sharepoint.com/sites/team");
}

#[tokio::test]
async fn parses_document_libraries_wrapped_response() {
    let json = r#"{"GetDocumentLibraries":[{"AbsoluteUrl":"https://contoso.sharepoint.com/sites/team/Shared Documents","Id":"22222222-2222-2222-2222-222222222222","IsDefaultDocumentLibrary":true,"ServerRelativeUrl":"/sites/team/Shared Documents","Title":"Documents"}]}"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let libs = get_document_libraries(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "https://contoso.sharepoint.com/sites/team",
    )
    .await
    .unwrap();

    assert_eq!(libs.len(), 1);
    assert_eq!(libs[0].title, "Documents");
    assert!(libs[0].is_default_document_library);
}

#[tokio::test]
async fn parses_document_libraries_bare_response() {
    let json = r#"[{"AbsoluteUrl":"https://contoso.sharepoint.com/sites/team/Shared Documents","Id":"22222222-2222-2222-2222-222222222222","IsDefaultDocumentLibrary":true,"ServerRelativeUrl":"/sites/team/Shared Documents","Title":"Documents"}]"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let libs = get_document_libraries(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "https://contoso.sharepoint.com/sites/team",
    )
    .await
    .unwrap();

    assert_eq!(libs.len(), 1);
}

#[tokio::test]
async fn parses_web_url_from_page_url_wrapped_response() {
    let json = r#"{"GetWebUrlFromPageUrl":"https://contoso.sharepoint.com/sites/team"}"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let url = get_web_url_from_page_url(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "https://contoso.sharepoint.com/sites/team/SitePages/Home.aspx",
    )
    .await
    .unwrap();

    assert_eq!(url, "https://contoso.sharepoint.com/sites/team");
}

#[tokio::test]
async fn parses_web_url_from_page_url_bare_response() {
    let json = r#""https://contoso.sharepoint.com/sites/team""#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let url = get_web_url_from_page_url(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "https://contoso.sharepoint.com/sites/team/SitePages/Home.aspx",
    )
    .await
    .unwrap();

    assert_eq!(url, "https://contoso.sharepoint.com/sites/team");
}
