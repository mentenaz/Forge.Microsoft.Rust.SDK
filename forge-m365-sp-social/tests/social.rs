use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_social::{
    follow, get_followed_documents_uri, get_followed_sites_uri, is_followed, my, my_followed,
    my_followed_count, my_followers, my_suggestions, stop_following, SocialActorInfo,
    SocialFollowResult,
};
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
            last_url: Mutex::new(None),
            last_method: Mutex::new(None),
            last_body: Mutex::new(None),
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
        let response = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(response.into_bytes()) })
    }
}

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

#[tokio::test]
async fn get_followed_sites_uri_unwraps_verbose_envelope() {
    let transport = RecordingTransport::with_response(r#"{"FollowedSitesUri":"https://x/sites"}"#);
    let client = Client::new(&transport);

    let uri = get_followed_sites_uri(&client, SITE).await.unwrap();

    assert_eq!(uri, "https://x/sites");
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert_eq!(
        url,
        "https://contoso.sharepoint.com/sites/team/_api/social.following/FollowedSitesUri"
    );
}

#[tokio::test]
async fn get_followed_documents_uri_handles_bare_string() {
    let transport = RecordingTransport::with_response(r#""https://x/docs""#);
    let client = Client::new(&transport);

    let uri = get_followed_documents_uri(&client, SITE).await.unwrap();

    assert_eq!(uri, "https://x/docs");
}

#[tokio::test]
async fn follow_posts_wrapped_actor_body() {
    let transport = RecordingTransport::with_response(r#"{"follow":0}"#);
    let client = Client::new(&transport);
    let actor = SocialActorInfo {
        account_name: Some("i:0#.f|membership|user@contoso.com".to_string()),
        actor_type: 0,
        content_uri: None,
        id: None,
        tag_guid: None,
    };

    let result = follow(&client, SITE, &actor).await.unwrap();

    assert_eq!(result, SocialFollowResult::Ok);
    assert_eq!(
        transport.last_method.lock().unwrap().clone().unwrap(),
        "POST"
    );
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/_api/social.following/follow"));
    let body = transport.last_body.lock().unwrap().clone().unwrap();
    let body: serde_json::Value = serde_json::from_slice(&body).unwrap();
    assert_eq!(body["actor"]["accountName"], serde_json::Value::Null);
}

#[tokio::test]
async fn follow_maps_unknown_code_to_other() {
    let transport = RecordingTransport::with_response(r#"{"follow":42}"#);
    let client = Client::new(&transport);
    let actor = SocialActorInfo {
        account_name: None,
        actor_type: 0,
        content_uri: None,
        id: Some("x".to_string()),
        tag_guid: None,
    };

    let result = follow(&client, SITE, &actor).await.unwrap();

    assert_eq!(result, SocialFollowResult::Other(42));
}

#[tokio::test]
async fn is_followed_parses_bool() {
    let transport = RecordingTransport::with_response(r#"{"isfollowed":true}"#);
    let client = Client::new(&transport);
    let actor = SocialActorInfo {
        account_name: None,
        actor_type: 2,
        content_uri: None,
        id: Some("site-id".to_string()),
        tag_guid: None,
    };

    let followed = is_followed(&client, SITE, &actor).await.unwrap();

    assert!(followed);
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/isfollowed"));
}

#[tokio::test]
async fn stop_following_posts_bodiless_response() {
    let transport = RecordingTransport::with_response("");
    let client = Client::new(&transport);
    let actor = SocialActorInfo {
        account_name: None,
        actor_type: 1,
        content_uri: Some("https://x/doc.docx".to_string()),
        id: None,
        tag_guid: None,
    };

    stop_following(&client, SITE, &actor).await.unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/stopfollowing"));
}

#[tokio::test]
async fn my_parses_full_summary() {
    let json = r#"{
        "SocialActor": {
            "ActorType": 0, "Id": "u1", "Uri": "https://x/u1", "Name": "User One",
            "IsFollowed": false, "Status": 0, "CanFollow": true, "ImageUri": "",
            "AccountName": "user@contoso.com", "EmailAddress": "user@contoso.com",
            "Title": "User One", "StatusText": "", "PersonalSiteUri": "https://x/personal",
            "FollowedContentUri": "", "ContentUri": "", "LibraryUri": "", "TagGuid": ""
        },
        "MyFollowedDocumentsUri": "https://x/docs",
        "MyFollowedSitesUri": "https://x/sites"
    }"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let data = my(&client, SITE).await.unwrap();

    assert_eq!(data.social_actor.id, "u1");
    assert_eq!(data.my_followed_sites_uri, "https://x/sites");
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/my"));
}

#[tokio::test]
async fn my_followed_unwraps_odata_results() {
    let json = r#"{"Followed":{"results":[
        {"ActorType":2,"Id":"s1","Uri":"https://x/s1","Name":"Site One","IsFollowed":true,
         "Status":0,"CanFollow":true,"ImageUri":"","AccountName":"","EmailAddress":"",
         "Title":"","StatusText":"","PersonalSiteUri":"","FollowedContentUri":"",
         "ContentUri":"https://x/s1","LibraryUri":"","TagGuid":""}
    ]}}"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let followed = my_followed(&client, SITE, 4).await.unwrap();

    assert_eq!(followed.len(), 1);
    assert_eq!(followed[0].id, "s1");
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/my/followed(types=4)"));
}

#[tokio::test]
async fn my_followed_count_handles_bare_number() {
    let transport = RecordingTransport::with_response("3");
    let client = Client::new(&transport);

    let count = my_followed_count(&client, SITE, 1).await.unwrap();

    assert_eq!(count, 3);
    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/my/followedcount(types=1)"));
}

#[tokio::test]
async fn my_followers_and_suggestions_hit_expected_urls() {
    let transport = RecordingTransport::with_response(r#"{"Followers":{"results":[]}}"#);
    let client = Client::new(&transport);
    my_followers(&client, SITE).await.unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/my/followers"));

    let transport = RecordingTransport::with_response(r#"{"Suggestions":{"results":[]}}"#);
    let client = Client::new(&transport);
    my_suggestions(&client, SITE).await.unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("/my/suggestions"));
}
