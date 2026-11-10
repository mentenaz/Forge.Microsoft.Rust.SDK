use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_webs::{get_parent_web_url, get_subwebs, get_web};

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
async fn parses_web_info() {
    let json = r#"{"Id":"55555555-5555-5555-5555-555555555555","Title":"Team Site","Url":"https://contoso.sharepoint.com/sites/team","ServerRelativeUrl":"/sites/team","Description":"","WebTemplate":"STS"}"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let web = get_web(&client, "https://contoso.sharepoint.com/sites/team/")
        .await
        .unwrap();

    assert_eq!(web.title, "Team Site");
    assert_eq!(web.server_relative_url, "/sites/team");
}

#[tokio::test]
async fn parses_subwebs() {
    let json = r#"[{"Id":"66666666-6666-6666-6666-666666666666","Title":"Sub","Url":"https://contoso.sharepoint.com/sites/team/sub","ServerRelativeUrl":"/sites/team/sub","Description":"","WebTemplate":"STS"}]"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let webs = get_subwebs(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(webs.len(), 1);
    assert_eq!(webs[0].title, "Sub");
}

#[tokio::test]
async fn parses_parent_web_url_when_present() {
    let json = r#"{"Url":"https://contoso.sharepoint.com/sites/team/sub","ParentWeb":{"ServerRelativeUrl":"/sites/team"}}"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let parent = get_parent_web_url(&client, "https://contoso.sharepoint.com/sites/team/sub")
        .await
        .unwrap();

    assert_eq!(
        parent.as_deref(),
        Some("https://contoso.sharepoint.com/sites/team")
    );
}

#[tokio::test]
async fn root_web_has_no_parent() {
    let json = r#"{"Url":"https://contoso.sharepoint.com/sites/team","ParentWeb":null}"#;
    let transport = FixedResponse(json);
    let client = Client::new(&transport);

    let parent = get_parent_web_url(&client, "https://contoso.sharepoint.com/sites/team")
        .await
        .unwrap();

    assert_eq!(parent, None);
}
