use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_forms::{get_form_by_id, get_forms};
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            last_url: Mutex::new(None),
        }
    }
}

impl Transport for RecordingTransport {
    fn execute(
        &self,
        _surface: Surface,
        _method: &str,
        url: &str,
        _headers: &[(&str, &str)],
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        *self.last_url.lock().unwrap() = Some(url.to_string());
        let body = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

const FORM_JSON: &str = r#"{"Id":"cccccccc-cccc-cccc-cccc-cccccccccccc","FormType":100,"ServerRelativeUrl":"/sites/team/Lists/Tasks/DispForm.aspx"}"#;

#[tokio::test]
async fn parses_forms_array() {
    let transport = RecordingTransport::with_response(&format!("[{FORM_JSON}]"));
    let client = Client::new(&transport);

    let forms = get_forms(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
    )
    .await
    .unwrap();

    assert_eq!(forms.len(), 1);
    assert_eq!(forms[0].form_type, 100);
}

#[tokio::test]
async fn get_by_id_builds_correct_url() {
    let transport = RecordingTransport::with_response(FORM_JSON);
    let client = Client::new(&transport);

    get_form_by_id(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        "Tasks",
        "cccccccc-cccc-cccc-cccc-cccccccccccc",
    )
    .await
    .unwrap();

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("forms('cccccccc-cccc-cccc-cccc-cccccccccccc')"));
}
