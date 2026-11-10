use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_comments::{
    add_comment, delete_comment, get_comments, like_comment, unlike_comment,
};
use serde_json::Value;
use std::collections::VecDeque;
use std::sync::Mutex;

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

const COMMENT_JSON: &str = r#"{"id":"1","text":"Looks good","author":{"email":"jane@contoso.com","id":7,"isActive":true,"loginName":"i:0#.f|membership|jane@contoso.com","name":"Jane Doe"},"createdDate":"2026-08-26T00:00:00Z","isLikedByUser":false,"likeCount":0,"replyCount":0,"isReply":false,"parentId":""}"#;

#[tokio::test]
async fn parses_comments_array() {
    let transport = RecordingTransport::default().queue(&format!("[{COMMENT_JSON}]"));
    let client = Client::new(&transport);

    let comments = get_comments(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
    )
    .await
    .unwrap();

    assert_eq!(comments.len(), 1);
    assert_eq!(comments[0].text, "Looks good");
    assert_eq!(comments[0].author.name, "Jane Doe");
}

#[tokio::test]
async fn add_sends_text_body() {
    let transport = RecordingTransport::default().queue(COMMENT_JSON);
    let client = Client::new(&transport);

    add_comment(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
        "Looks good",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "POST");
    let sent: Value = serde_json::from_slice(calls[0].body.as_ref().unwrap()).unwrap();
    assert_eq!(sent["text"], "Looks good");
}

#[tokio::test]
async fn delete_sends_delete_to_comment_id_url() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);

    delete_comment(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
        "1",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls[0].method, "DELETE");
    assert!(calls[0].url.ends_with("/comments(1)"));
}

#[tokio::test]
async fn like_and_unlike_send_correct_action_urls() {
    let transport = RecordingTransport::default().queue("").queue("");
    let client = Client::new(&transport);

    like_comment(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
        "1",
    )
    .await
    .unwrap();
    unlike_comment(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        1,
        "1",
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert!(calls[0].url.ends_with("/comments(1)/Like"));
    assert!(calls[1].url.ends_with("/comments(1)/Unlike"));
}
