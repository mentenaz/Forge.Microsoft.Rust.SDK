use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_subscriptions as subscriptions;
use std::collections::VecDeque;
use std::sync::Mutex;

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

const LIST_BASE: &str =
    "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions";

const ID: &str = "abc-123";

const NOTIFY: &str = "https://hook.example/notify";

const EXPIRES: &str = "2027-02-27T00:00:00+00:00";

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

fn sub_json(id: &str) -> String {
    format!(
        r#"{{"id":"{id}","resource":"{LIST_BASE}","notificationUrl":"{NOTIFY}","expirationDateTime":"{EXPIRES}","clientState":"cs","lastModifiedTime":"2027-02-20T09:00:00Z"}}"#
    )
}

#[tokio::test]
async fn list_gets_collection() {
    let transport = RecordingTransport::default()
        .queue(&format!(r#"{{"d":{{"results":[{}]}}}}"#, sub_json("one")));
    let client = Client::new(&transport);

    let subs = subscriptions::list_subscriptions(&client, SITE, "Tasks")
        .await
        .unwrap();

    assert_eq!(subs.len(), 1);
    assert_eq!(subs[0].id, "one");
    assert_eq!(subs[0].notification_url, NOTIFY);
    assert_eq!(subs[0].client_state.as_deref(), Some("cs"));
    transport.assert_last(("GET", LIST_BASE));
}

#[tokio::test]
async fn get_by_id_fetches_instance() {
    let transport = RecordingTransport::default().queue(r#"{"d":{}}"#);
    let client = Client::new(&transport);

    let sub = subscriptions::get_subscription_by_id(&client, SITE, "Tasks", ID)
        .await
        .unwrap();

    assert_eq!(sub.id, "");
    transport.assert_last((
        "GET",
        "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions('abc-123')",
    ));
}

#[tokio::test]
async fn get_by_id_parses_instance() {
    let transport = RecordingTransport::default().queue(&format!(r#"{{"d":{}}}"#, sub_json(ID)));
    let client = Client::new(&transport);

    let sub = subscriptions::get_subscription_by_id(&client, SITE, "Tasks", ID)
        .await
        .unwrap();

    assert_eq!(sub.id, ID);
    assert_eq!(sub.expiration_date_time, EXPIRES);
    assert!(sub.last_modified_time.is_some());
    transport.assert_last((
        "GET",
        "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions('abc-123')",
    ));
}

#[tokio::test]
async fn add_posts_with_client_state() {
    let transport = RecordingTransport::default().queue(&format!(r#"{{"d":{}}}"#, sub_json(ID)));
    let client = Client::new(&transport);

    let sub = subscriptions::add_subscription(&client, SITE, "Tasks", NOTIFY, EXPIRES, Some("cs"))
        .await
        .unwrap();

    assert_eq!(sub.id, ID);
    transport.assert_last_with_body(
        ("POST", LIST_BASE),
        &format!(
            r#"{{"clientState":"cs","expirationDateTime":"{EXPIRES}","notificationUrl":"{NOTIFY}","resource":"{LIST_BASE}"}}"#
        ),
    );
}

#[tokio::test]
async fn add_omits_empty_client_state() {
    let transport = RecordingTransport::default().queue(&format!(r#"{{"d":{}}}"#, sub_json(ID)));
    let client = Client::new(&transport);

    let sub = subscriptions::add_subscription(&client, SITE, "Tasks", NOTIFY, EXPIRES, None)
        .await
        .unwrap();

    assert_eq!(sub.id, ID);
    transport.assert_last_with_body(
        ("POST", LIST_BASE),
        &format!(
            r#"{{"expirationDateTime":"{EXPIRES}","notificationUrl":"{NOTIFY}","resource":"{LIST_BASE}"}}"#
        ),
    );
}

#[tokio::test]
async fn add_drops_empty_string_client_state() {
    let transport = RecordingTransport::default().queue(&format!(r#"{{"d":{}}}"#, sub_json(ID)));
    let client = Client::new(&transport);

    let sub = subscriptions::add_subscription(&client, SITE, "Tasks", NOTIFY, EXPIRES, Some(""))
        .await
        .unwrap();

    assert_eq!(sub.id, ID);
    transport.assert_last_with_body(
        ("POST", LIST_BASE),
        &format!(
            r#"{{"expirationDateTime":"{EXPIRES}","notificationUrl":"{NOTIFY}","resource":"{LIST_BASE}"}}"#
        ),
    );
}

#[tokio::test]
async fn update_patches_supplied_fields() {
    let transport = RecordingTransport::default().queue(&format!(r#"{{"d":{}}}"#, sub_json(ID)));
    let client = Client::new(&transport);

    let sub = subscriptions::update_subscription(
        &client,
        SITE,
        "Tasks",
        ID,
        Some(EXPIRES),
        Some(NOTIFY),
        Some("new-cs"),
    )
    .await
    .unwrap();

    assert_eq!(sub.id, ID);
    transport.assert_last_with_body(
        (
            "PATCH",
            "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions('abc-123')",
        ),
        &format!(
            r#"{{"clientState":"new-cs","expirationDateTime":"{EXPIRES}","notificationUrl":"{NOTIFY}"}}"#
        ),
    );
}

#[tokio::test]
async fn update_partial_only_expiration() {
    let transport = RecordingTransport::default().queue(&format!(r#"{{"d":{}}}"#, sub_json(ID)));
    let client = Client::new(&transport);

    let _ =
        subscriptions::update_subscription(&client, SITE, "Tasks", ID, Some(EXPIRES), None, None)
            .await
            .unwrap();

    transport.assert_last_with_body(
        (
            "PATCH",
            "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions('abc-123')",
        ),
        &format!(r#"{{"expirationDateTime":"{EXPIRES}"}}"#),
    );
}

#[tokio::test]
async fn update_drops_empty_strings() {
    let transport = RecordingTransport::default().queue(&format!(r#"{{"d":{}}}"#, sub_json(ID)));
    let client = Client::new(&transport);

    let _ =
        subscriptions::update_subscription(&client, SITE, "Tasks", ID, None, Some(""), Some(""))
            .await
            .unwrap();

    transport.assert_last_with_body(
        (
            "PATCH",
            "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions('abc-123')",
        ),
        r#"{}"#,
    );
}

#[tokio::test]
async fn update_empty_patches_empty_object() {
    let transport = RecordingTransport::default().queue(&format!(r#"{{"d":{}}}"#, sub_json(ID)));
    let client = Client::new(&transport);

    let _ = subscriptions::update_subscription(&client, SITE, "Tasks", ID, None, None, None)
        .await
        .unwrap();

    transport.assert_last_with_body(
        (
            "PATCH",
            "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions('abc-123')",
        ),
        r#"{}"#,
    );
}

#[tokio::test]
async fn delete_issues_delete() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    subscriptions::delete_subscription(&client, SITE, "Tasks", ID)
        .await
        .unwrap();

    transport.assert_last((
        "DELETE",
        "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions('abc-123')",
    ));
}

#[tokio::test]
async fn title_with_quote_encoded() {
    let transport = RecordingTransport::default().queue(r#"{"d":{"results":[]}}"#);
    let client = Client::new(&transport);

    let subs = subscriptions::list_subscriptions(&client, SITE, "O'Brien's List")
        .await
        .unwrap();

    assert!(subs.is_empty());
    transport.assert_last((
        "GET",
        "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('O%27%27Brien%27%27s%20List')/subscriptions",
    ));
}

#[tokio::test]
async fn by_id_doubles_quote_in_id() {
    let transport = RecordingTransport::default().queue(r#"{"d":{}}"#);
    let client = Client::new(&transport);

    let _ = subscriptions::get_subscription_by_id(&client, SITE, "Tasks", "a'b")
        .await
        .unwrap();

    transport.assert_last((
        "GET",
        "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/subscriptions('a''b')",
    ));
}
