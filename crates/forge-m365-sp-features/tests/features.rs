use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_features as features;
use std::collections::VecDeque;
use std::sync::Mutex;

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

const GUID: &str = "9a234d20-0000-0000-0000-000000000000";

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
async fn get_site_features_lists_collection() {
    let transport = RecordingTransport::default().queue(&format!(
        r#"{{"d":{{"results":[{{"DefinitionId":"{GUID}"}}]}}}}"#
    ));
    let client = Client::new(&transport);

    let infos = features::get_site_features(&client, SITE).await.unwrap();

    assert_eq!(infos.len(), 1);
    assert_eq!(infos[0].definition_id, GUID);
    transport.assert_last(("GET", &format!("{SITE}/_api/site/features")));
}

#[tokio::test]
async fn get_web_features_lists_collection() {
    let transport = RecordingTransport::default().queue(r#"{"d":{"results":[]}}"#);
    let client = Client::new(&transport);

    let infos = features::get_web_features(&client, SITE).await.unwrap();

    assert!(infos.is_empty());
    transport.assert_last(("GET", &format!("{SITE}/_api/web/features")));
}

#[tokio::test]
async fn get_site_feature_by_id_fetches_instance() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"DefinitionId":"9a234d20-0000-0000-0000-000000000000"}}"#);
    let client = Client::new(&transport);

    let info = features::get_site_feature_by_id(&client, SITE, GUID)
        .await
        .unwrap();

    assert_eq!(info.definition_id, GUID);
    transport.assert_last((
        "GET",
        "https://contoso.sharepoint.com/sites/team/_api/site/features('9a234d20-0000-0000-0000-000000000000')",
    ));
}

#[tokio::test]
async fn get_web_feature_by_id_fetches_instance() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"DefinitionId":"9a234d20-0000-0000-0000-000000000000"}}"#);
    let client = Client::new(&transport);

    let info = features::get_web_feature_by_id(&client, SITE, GUID)
        .await
        .unwrap();

    assert_eq!(info.definition_id, GUID);
    transport.assert_last((
        "GET",
        "https://contoso.sharepoint.com/sites/team/_api/web/features('9a234d20-0000-0000-0000-000000000000')",
    ));
}

#[tokio::test]
async fn activate_site_feature_posts_add() {
    let transport =
        RecordingTransport::default().queue(&format!(r#"{{"d":{{"DefinitionId":"{GUID}"}}}}"#));
    let client = Client::new(&transport);

    let result = features::activate_site_feature(&client, SITE, GUID, false)
        .await
        .unwrap();

    assert_eq!(result.definition_id, GUID);
    transport.assert_last_with_body(
        ("POST", &format!("{SITE}/_api/site/features/add")),
        &format!(r#"{{"featdefScope":0,"featureId":"{GUID}","force":false}}"#),
    );
}

#[tokio::test]
async fn activate_web_feature_posts_add_with_force() {
    let transport =
        RecordingTransport::default().queue(&format!(r#"{{"d":{{"DefinitionId":"{GUID}"}}}}"#));
    let client = Client::new(&transport);

    let result = features::activate_web_feature(&client, SITE, GUID, true)
        .await
        .unwrap();

    assert_eq!(result.definition_id, GUID);
    transport.assert_last_with_body(
        ("POST", &format!("{SITE}/_api/web/features/add")),
        &format!(r#"{{"featdefScope":0,"featureId":"{GUID}","force":true}}"#),
    );
}

#[tokio::test]
async fn deactivate_site_feature_posts_remove() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    features::deactivate_site_feature(&client, SITE, GUID, false)
        .await
        .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{SITE}/_api/site/features/remove")),
        &format!(r#"{{"featureId":"{GUID}","force":false}}"#),
    );
}

#[tokio::test]
async fn deactivate_web_feature_posts_remove_with_force() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    features::deactivate_web_feature(&client, SITE, GUID, true)
        .await
        .unwrap();

    transport.assert_last_with_body(
        ("POST", &format!("{SITE}/_api/web/features/remove")),
        &format!(r#"{{"featureId":"{GUID}","force":true}}"#),
    );
}
