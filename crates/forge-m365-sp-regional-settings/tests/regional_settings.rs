use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_regional_settings as rs;
use std::collections::VecDeque;
use std::sync::Mutex;

const SITE: &str = "https://contoso.sharepoint.com/sites/team";

const REGION: &str = "https://contoso.sharepoint.com/sites/team/_api/web/regionalsettings";

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

fn region_payload() -> &'static str {
    r#"{"d":{"AdjustHijriDays":0,"AlternateCalendarType":0,"AM":"AM","CalendarType":1,"Collation":15,"CollationLCID":2070,"DateFormat":1,"DateSeparator":"/","DecimalSeparator":".","DigitGrouping":"3;0","FirstDayOfWeek":0,"FirstWeekOfYear":0,"IsEastAsia":false,"IsRightToLeft":false,"IsUIRightToLeft":false,"ListSeparator":",","LocaleId":1033,"NegativeSign":"-","NegNumberMode":1,"PM":"PM","PositiveSign":"","ShowWeeks":false,"ThousandSeparator":",","Time24":false,"TimeMarkerPosition":1,"TimeSeparator":":","WorkDayEndHour":1020,"WorkDays":62,"WorkDayStartHour":480}}"#
}

#[tokio::test]
async fn get_regional_settings_parses_full() {
    let transport = RecordingTransport::default().queue(region_payload());
    let client = Client::new(&transport);

    let info = rs::get_regional_settings(&client, SITE).await.unwrap();

    assert_eq!(info.locale_id, 1033);
    assert_eq!(info.time_separator, ":");
    assert!(!info.show_weeks);
    assert_eq!(info.am, "AM");
    transport.assert_last(("GET", REGION));
}

#[tokio::test]
async fn get_time_zone_fetches_instance() {
    let transport = RecordingTransport::default().queue(
        r#"{"d":{"Description":"(UTC-08:00) Pacific Time (US and Canada)","Id":13,"Information":{"Bias":480,"DaylightBias":-60,"StandardBias":0}}}"#,
    );
    let client = Client::new(&transport);

    let tz = rs::get_time_zone(&client, SITE).await.unwrap();

    assert_eq!(tz.id, 13);
    assert!(tz.description.contains("Pacific Time"));
    assert_eq!(tz.information.as_ref().unwrap().bias, 480);
    transport.assert_last(("GET", &format!("{REGION}/timezone")));
}

#[tokio::test]
async fn get_time_zones_lists_collection() {
    let transport = RecordingTransport::default().queue(
        r#"{"d":{"results":[{"Description":"A","Id":1,"Information":{"Bias":0,"DaylightBias":0,"StandardBias":0}},{"Description":"B","Id":2,"Information":{"Bias":60,"DaylightBias":0,"StandardBias":0}}]}}"#,
    );
    let client = Client::new(&transport);

    let tzs = rs::get_time_zones(&client, SITE).await.unwrap();

    assert_eq!(tzs.len(), 2);
    assert_eq!(tzs[1].id, 2);
    transport.assert_last(("GET", &format!("{REGION}/timezones")));
}

#[tokio::test]
async fn get_time_zone_by_id_posts_get_by_id() {
    let transport = RecordingTransport::default().queue(
        r#"{"d":{"Description":"(UTC-08:00) Pacific Time (US and Canada)","Id":13,"Information":{"Bias":480,"DaylightBias":-60,"StandardBias":0}}}"#,
    );
    let client = Client::new(&transport);

    let tz = rs::get_time_zone_by_id(&client, SITE, 13).await.unwrap();

    assert_eq!(tz.id, 13);
    transport.assert_last(("POST", &format!("{REGION}/timezones/GetById(13)")));
}

#[tokio::test]
async fn get_installed_languages_unwraps_items() {
    let transport = RecordingTransport::default().queue(
        r#"{"d":{"Items":[{"DisplayName":"English","LanguageTag":"en-US","Lcid":1033},{"DisplayName":"French","LanguageTag":"fr-FR","Lcid":1036}]}}"#,
    );
    let client = Client::new(&transport);

    let langs = rs::get_installed_languages(&client, SITE).await.unwrap();

    assert_eq!(langs.len(), 2);
    assert_eq!(langs[0].language_tag, "en-US");
    assert_eq!(langs[1].lcid, 1036);
    transport.assert_last(("GET", &format!("{REGION}/installedlanguages")));
}

#[tokio::test]
async fn utc_to_local_time_peels_boxed_key() {
    let transport = RecordingTransport::default()
        .queue(r#"{"d":{"UTCToLocalTime":"2027-01-01T04:00:00.000Z"}}"#);
    let client = Client::new(&transport);

    let local = rs::utc_to_local_time(&client, SITE, "2027-01-01T12:00:00.000Z")
        .await
        .unwrap();

    assert_eq!(local, "2027-01-01T04:00:00.000Z");
    transport.assert_last((
        "POST",
        &format!("{REGION}/timezone/utctolocaltime('2027-01-01T12:00:00.000Z')"),
    ));
}

#[tokio::test]
async fn local_time_to_utc_handles_bare_string() {
    let transport = RecordingTransport::default().queue(r#""2027-01-01T20:00:00.000Z""#);
    let client = Client::new(&transport);

    let utc = rs::local_time_to_utc(&client, SITE, "2027-01-01T12:00:00.000Z")
        .await
        .unwrap();

    assert_eq!(utc, "2027-01-01T20:00:00.000Z");
    transport.assert_last((
        "POST",
        &format!("{REGION}/timezone/localtimetoutc('2027-01-01T12:00:00.000Z')"),
    ));
}

#[tokio::test]
async fn web_title_resource_posts_culture() {
    let transport = RecordingTransport::default().queue(r#""Team""#);
    let client = Client::new(&transport);

    let title = rs::web_title_resource(&client, SITE, "en-US")
        .await
        .unwrap();

    assert_eq!(title, "Team");
    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/web/titleResource/getValueForUICulture",
        ),
        r#"{"cultureName":"en-US"}"#,
    );
}

#[tokio::test]
async fn web_description_resource_posts_culture() {
    let transport = RecordingTransport::default().queue(r#"{"d":"A team site"}"#);
    let client = Client::new(&transport);

    let desc = rs::web_description_resource(&client, SITE, "fr-FR")
        .await
        .unwrap();

    assert_eq!(desc, "A team site");
    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/web/descriptionResource/getValueForUICulture",
        ),
        r#"{"cultureName":"fr-FR"}"#,
    );
}

#[tokio::test]
async fn list_title_resource_uses_list_queryable() {
    let transport = RecordingTransport::default().queue(r#""Tasks""#);
    let client = Client::new(&transport);

    let title = rs::list_title_resource(&client, SITE, "Tasks", "de-DE")
        .await
        .unwrap();

    assert_eq!(title, "Tasks");
    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/titleResource/getValueForUICulture",
        ),
        r#"{"cultureName":"de-DE"}"#,
    );
}

#[tokio::test]
async fn field_title_resource_uses_get_by_internal_name() {
    let transport = RecordingTransport::default().queue(r#""Priority""#);
    let client = Client::new(&transport);

    let title = rs::field_title_resource(&client, SITE, "Tasks", "Priority", "en-US")
        .await
        .unwrap();

    assert_eq!(title, "Priority");
    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('Tasks')/fields/getByInternalNameOrTitle('Priority')/titleResource/getValueForUICulture",
        ),
        r#"{"cultureName":"en-US"}"#,
    );
}

#[tokio::test]
async fn content_type_uses_name_resource() {
    let transport = RecordingTransport::default().queue(r#""Task""#);
    let client = Client::new(&transport);

    let name = rs::content_type_name_resource(&client, SITE, "0x0108", "en-US")
        .await
        .unwrap();

    assert_eq!(name, "Task");
    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/web/contenttypes/getbyid('0x0108')/nameResource/getValueForUICulture",
        ),
        r#"{"cultureName":"en-US"}"#,
    );
}

#[tokio::test]
async fn content_type_description_resource_posts() {
    let transport = RecordingTransport::default().queue(r#""A task item""#);
    let client = Client::new(&transport);

    let desc = rs::content_type_description_resource(&client, SITE, "0x0108", "en-US")
        .await
        .unwrap();

    assert_eq!(desc, "A task item");
    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/web/contenttypes/getbyid('0x0108')/descriptionResource/getValueForUICulture",
        ),
        r#"{"cultureName":"en-US"}"#,
    );
}

#[tokio::test]
async fn usercustomaction_title_resource_posts() {
    let transport = RecordingTransport::default().queue(r#""Custom""#);
    let client = Client::new(&transport);

    let title = rs::usercustomaction_title_resource(&client, SITE, "uc-1", "en-US")
        .await
        .unwrap();

    assert_eq!(title, "Custom");
    transport.assert_last_with_body(
        (
            "POST",
            "https://contoso.sharepoint.com/sites/team/_api/web/usercustomactions('uc-1')/titleResource/getValueForUICulture",
        ),
        r#"{"cultureName":"en-US"}"#,
    );
}

#[tokio::test]
async fn list_title_encodes_quotes_in_title() {
    let transport = RecordingTransport::default().queue(r#""O'B's""#);
    let client = Client::new(&transport);

    let title = rs::list_title_resource(&client, SITE, "O'Brien's", "en-US")
        .await
        .unwrap();

    assert_eq!(title, "O'B's");
    transport.assert_last((
        "POST",
        "https://contoso.sharepoint.com/sites/team/_api/web/lists/getbytitle('O%27%27Brien%27%27s')/titleResource/getValueForUICulture",
    ));
}

#[tokio::test]
async fn web_title_resource_rejects_non_string() {
    let transport = RecordingTransport::default().queue(r#"{"Result":42}"#);
    let client = Client::new(&transport);

    let err = rs::web_title_resource(&client, SITE, "en-US")
        .await
        .unwrap_err();

    assert!(matches!(err, Error::Surface(Surface::SpRest, _)));
}
