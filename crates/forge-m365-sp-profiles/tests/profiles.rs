use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_profiles as profiles;
use std::collections::VecDeque;
use std::sync::Mutex;

#[derive(Debug, Clone)]
struct RecordedCall {
    method: String,
    url: String,
    headers: Vec<(String, String)>,
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

    fn assert_called(&self, expected: (&str, &str)) {
        let calls = self.calls.lock().unwrap();
        let call = calls.last().expect("no call was recorded");
        assert_eq!(
            (call.method.as_str(), call.url.as_str()),
            expected,
            "unexpected request"
        );
    }

    fn assert_called_with_body(&self, expected: (&str, &str), body: &str) {
        self.assert_called(expected);
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
        headers: &[(&str, &str)],
        body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        let call = RecordedCall {
            method: method.to_string(),
            url: url.to_string(),
            headers: headers
                .iter()
                .map(|(k, v)| (k.to_string(), v.to_string()))
                .collect(),
            body: body.map(|b| b.to_vec()),
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
async fn test_get_edit_profile_link() {
    let transport = RecordingTransport::default()
        .queue(r#""https://contoso.sharepoint.com/_layouts/15/EditProfile.aspx""#);
    let client = Client::new(&transport);
    let result =
        profiles::get_edit_profile_link(&client, "https://contoso.sharepoint.com/sites/test")
            .await
            .unwrap();
    assert_eq!(
        result,
        "https://contoso.sharepoint.com/_layouts/15/EditProfile.aspx"
    );
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/EditProfileLink",
    ));
}

#[tokio::test]
async fn test_get_is_my_people_list_public() {
    let transport = RecordingTransport::default().queue("true");
    let client = Client::new(&transport);
    let result = profiles::get_is_my_people_list_public(
        &client,
        "https://contoso.sharepoint.com/sites/test",
    )
    .await
    .unwrap();
    assert!(result);
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/IsMyPeopleListPublic",
    ));
}

#[tokio::test]
async fn test_am_i_followed_by() {
    let transport = RecordingTransport::default().queue("true");
    let client = Client::new(&transport);
    let result = profiles::am_i_followed_by(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
    )
    .await
    .unwrap();
    assert!(result);
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/amifollowedby(@v)?@v='i:0#.f|membership|user@contoso.com'",
    ));
}

#[tokio::test]
async fn test_am_i_following() {
    let transport = RecordingTransport::default().queue("false");
    let client = Client::new(&transport);
    let result = profiles::am_i_following(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
    )
    .await
    .unwrap();
    assert!(!result);
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/amifollowing(@v)?@v='i:0#.f|membership|user@contoso.com'",
    ));
}

#[tokio::test]
async fn test_get_followed_tags() {
    let transport = RecordingTransport::default().queue(r#"["sharepoint","rust"]"#);
    let client = Client::new(&transport);
    let result =
        profiles::get_followed_tags(&client, "https://contoso.sharepoint.com/sites/test", 20)
            .await
            .unwrap();
    assert_eq!(result, vec!["sharepoint".to_string(), "rust".to_string()]);
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/getfollowedtags(20)",
    ));
}

#[tokio::test]
async fn test_get_followers_for() {
    let transport = RecordingTransport::default()
        .queue(r#"[{"AccountName":"i:0#.f|membership|user1@contoso.com","DisplayName":"User 1"}]"#);
    let client = Client::new(&transport);
    let result = profiles::get_followers_for(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
    )
    .await
    .unwrap();
    assert_eq!(result.len(), 1);
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/getfollowersfor(@v)?@v='i:0#.f|membership|user@contoso.com'",
    ));
}

#[tokio::test]
async fn test_get_my_followers() {
    let transport = RecordingTransport::default().queue(r#"[{"DisplayName":"User 1"}]"#);
    let client = Client::new(&transport);
    let result = profiles::get_my_followers(&client, "https://contoso.sharepoint.com/sites/test")
        .await
        .unwrap();
    assert_eq!(result.len(), 1);
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/getmyfollowers",
    ));
}

#[tokio::test]
async fn test_get_my_properties() {
    let transport = RecordingTransport::default()
        .queue(r#"{"AccountName":"i:0#.f|membership|user@contoso.com","DisplayName":"Test User"}"#);
    let client = Client::new(&transport);
    let result = profiles::get_my_properties(&client, "https://contoso.sharepoint.com/sites/test")
        .await
        .unwrap();
    assert!(result.is_object());
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/getmyproperties",
    ));
}

#[tokio::test]
async fn test_get_people_followed_by() {
    let transport = RecordingTransport::default().queue(r#"[{"DisplayName":"User 1"}]"#);
    let client = Client::new(&transport);
    let result = profiles::get_people_followed_by(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
    )
    .await
    .unwrap();
    assert_eq!(result.len(), 1);
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/getpeoplefollowedby(@v)?@v='i:0#.f|membership|user@contoso.com'",
    ));
}

#[tokio::test]
async fn test_get_properties_for() {
    let transport = RecordingTransport::default()
        .queue(r#"{"AccountName":"i:0#.f|membership|user@contoso.com","DisplayName":"Test User"}"#);
    let client = Client::new(&transport);
    let result = profiles::get_properties_for(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
    )
    .await
    .unwrap();
    assert_eq!(result["DisplayName"], "Test User");
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/getpropertiesfor(@v)?@v='i:0#.f|membership|user@contoso.com'",
    ));
}

#[tokio::test]
async fn test_get_trending_tags() {
    let transport = RecordingTransport::default()
        .queue(r#"{"Items":[{"Name":"sharepoint","UseCount":42},{"Name":"rust","UseCount":17}]}"#);
    let client = Client::new(&transport);
    let result = profiles::get_trending_tags(&client, "https://contoso.sharepoint.com/sites/test")
        .await
        .unwrap();
    assert_eq!(result.items.len(), 2);
    assert_eq!(result.items[0].use_count, Some(42));
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/gettrendingtags",
    ));
}

#[tokio::test]
async fn test_get_user_profile_property_for() {
    let transport = RecordingTransport::default().queue(r#""user@contoso.com""#);
    let client = Client::new(&transport);
    let result = profiles::get_user_profile_property_for(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
        "WorkEmail",
    )
    .await
    .unwrap();
    assert_eq!(result, "user@contoso.com");
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/getuserprofilepropertyfor(accountname=@v, propertyname='WorkEmail')?@v='i:0#.f|membership|user@contoso.com'",
    ));
}

#[tokio::test]
async fn test_is_following() {
    let transport = RecordingTransport::default().queue("true");
    let client = Client::new(&transport);
    let result = profiles::is_following(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|follower@contoso.com",
        "i:0#.f|membership|followee@contoso.com",
    )
    .await
    .unwrap();
    assert!(result);
    transport.assert_called((
        "GET",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/isfollowing(possiblefolloweraccountname=@v, possiblefolloweeaccountname=@y)?@v='i:0#.f|membership|follower@contoso.com'&@y='i:0#.f|membership|followee@contoso.com'",
    ));
}

#[tokio::test]
async fn test_hide_suggestion() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);
    profiles::hide_suggestion(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
    )
    .await
    .unwrap();
    transport.assert_called((
        "POST",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/hidesuggestion(@v)?@v='i:0#.f|membership|user@contoso.com'",
    ));
}

#[tokio::test]
async fn test_set_single_value_profile_property() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);
    profiles::set_single_value_profile_property(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
        "AboutMe",
        "Hello world",
    )
    .await
    .unwrap();
    transport.assert_called_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/SetSingleValueProfileProperty",
        ),
        r#"{"accountName":"i:0#.f|membership|user@contoso.com","propertyName":"AboutMe","propertyValue":"Hello world"}"#,
    );
}

#[tokio::test]
async fn test_set_multi_valued_profile_property() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);
    profiles::set_multi_valued_profile_property(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        "i:0#.f|membership|user@contoso.com",
        "Skills",
        &["Rust".to_string(), "TypeScript".to_string()],
    )
    .await
    .unwrap();
    transport.assert_called_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/SetMultiValuedProfileProperty",
        ),
        r#"{"accountName":"i:0#.f|membership|user@contoso.com","propertyName":"Skills","propertyValues":["Rust","TypeScript"]}"#,
    );
}

#[tokio::test]
async fn test_set_my_profile_pic() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);
    profiles::set_my_profile_pic(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        &[0xff, 0xd8, 0xff, 0xdb],
    )
    .await
    .unwrap();
    transport.assert_called((
        "POST",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.peoplemanager/setmyprofilepicture",
    ));
    let calls = transport.calls.lock().unwrap();
    let call = calls.last().unwrap();
    let content_type = call
        .headers
        .iter()
        .find(|(k, _)| k == "Content-Type")
        .expect("missing Content-Type header");
    assert_eq!(content_type.1, "application/octet-stream");
    assert_eq!(call.body.as_deref(), Some(&[0xff, 0xd8, 0xff, 0xdb][..]));
}

#[tokio::test]
async fn test_get_user_profile() {
    let transport = RecordingTransport::default()
        .queue(r#"{"AccountName":"i:0#.f|membership|user@contoso.com","DisplayName":"Test User"}"#);
    let client = Client::new(&transport);
    let result = profiles::get_user_profile(&client, "https://contoso.sharepoint.com/sites/test")
        .await
        .unwrap();
    assert_eq!(
        result.account_name,
        Some("i:0#.f|membership|user@contoso.com".to_string())
    );
    transport.assert_called((
        "POST",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.profileloader/getprofileloader/getuserprofile",
    ));
}

#[tokio::test]
async fn test_get_owner_user_profile() {
    let transport = RecordingTransport::default().queue(
        r#"{"AccountName":"i:0#.f|membership|admin@contoso.com","DisplayName":"Admin User"}"#,
    );
    let client = Client::new(&transport);
    let result =
        profiles::get_owner_user_profile(&client, "https://contoso.sharepoint.com/sites/test")
            .await
            .unwrap();
    assert_eq!(
        result.account_name,
        Some("i:0#.f|membership|admin@contoso.com".to_string())
    );
    transport.assert_called((
        "POST",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.profileloader/getowneruserprofile",
    ));
}

#[tokio::test]
async fn test_create_personal_site() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);
    profiles::create_personal_site(&client, "https://contoso.sharepoint.com/sites/test", true)
        .await
        .unwrap();
    transport.assert_called((
        "POST",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.profileloader/getprofileloader/getuserprofile/createpersonalsiteenque(true)",
    ));
}

#[tokio::test]
async fn test_create_personal_site_enqueue_bulk() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);
    profiles::create_personal_site_enqueue_bulk(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        &[
            "user1@contoso.com".to_string(),
            "user2@contoso.com".to_string(),
        ],
    )
    .await
    .unwrap();
    transport.assert_called_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.profileloader/getprofileloader/createpersonalsiteenqueuebulk",
        ),
        r#"{"emailIDs":["user1@contoso.com","user2@contoso.com"]}"#,
    );
}

#[tokio::test]
async fn test_share_all_social_data() {
    let transport = RecordingTransport::default().queue("");
    let client = Client::new(&transport);
    profiles::share_all_social_data(&client, "https://contoso.sharepoint.com/sites/test", true)
        .await
        .unwrap();
    transport.assert_called((
        "POST",
        "https://contoso.sharepoint.com/sites/test/_api/sp.userprofiles.profileloader/getprofileloader/getuserprofile/shareallsocialdata(true)",
    ));
}

#[tokio::test]
async fn test_client_people_picker_resolve_user() {
    let transport = RecordingTransport::default().queue(
        r#"{"Description":"user@contoso.com","DisplayText":"Test User","EntityData":{"Email":"user@contoso.com"},"EntityType":"User","IsResolved":true,"Key":"i:0#.f|membership|user@contoso.com","MultipleMatches":[],"ProviderDisplayName":"Tenant","ProviderName":"tenant"}"#,
    );
    let client = Client::new(&transport);
    let params = profiles::types::PeoplePickerQueryParameters {
        query_string: "user@contoso.com".to_string(),
        maximum_entity_suggestions: 15,
        ..Default::default()
    };
    let result = profiles::client_people_picker_resolve_user(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        &params,
    )
    .await
    .unwrap();
    assert!(result.is_resolved);
    assert_eq!(result.key, "i:0#.f|membership|user@contoso.com");
    transport.assert_called((
        "POST",
        "https://contoso.sharepoint.com/sites/test/_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface/clientpeoplepickerresolveuser",
    ));
}

#[tokio::test]
async fn test_client_people_picker_search_user() {
    let transport = RecordingTransport::default().queue(
        r#"[{"Description":"","DisplayText":"User 1","EntityData":{},"EntityType":"User","IsResolved":false,"Key":"user1","MultipleMatches":[],"ProviderDisplayName":"","ProviderName":""},{"Description":"","DisplayText":"User 2","EntityData":{},"EntityType":"User","IsResolved":false,"Key":"user2","MultipleMatches":[],"ProviderDisplayName":"","ProviderName":""}]"#,
    );
    let client = Client::new(&transport);
    let params = profiles::types::PeoplePickerQueryParameters {
        query_string: "user".to_string(),
        maximum_entity_suggestions: 15,
        ..Default::default()
    };
    let result = profiles::client_people_picker_search_user(
        &client,
        "https://contoso.sharepoint.com/sites/test",
        &params,
    )
    .await
    .unwrap();
    assert_eq!(result.len(), 2);
    transport.assert_called((
        "POST",
        "https://contoso.sharepoint.com/sites/test/_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface/clientpeoplepickersearchuser",
    ));
}
