use forge_m365_core::{Client, Error, Result, Surface, Transport};
use forge_m365_sp_column_defaults::{
    build_defaults_xml, clear_folder_default_column_values, get_default_column_values,
    get_folder_default_column_values, set_default_column_values, set_folder_default_column_values,
    ColumnDefaultValue, FieldDefaultInput, FieldDefaultProps, TaxonomyTermRef,
};
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

    fn queue_err(self, err: Error) -> Self {
        self.responses.lock().unwrap().push_back(Err(err));
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

const SITE: &str = "https://contoso.sharepoint.com/sites/team";
const DEFAULTS_XML: &str = "<MetadataDefaults><a href=\"/sites/team/Shared%20Documents\"><DefaultValue FieldName=\"Status\">Active</DefaultValue><DefaultValue FieldName=\"Region\">5;#North|guid-1</DefaultValue></a><a href=\"/sites/team/Shared%20Documents/Sub%20Folder\"><DefaultValue FieldName=\"Status\">Inactive</DefaultValue></a></MetadataDefaults>";

fn root_folder_json(path: &str) -> String {
    format!(r#"{{"ServerRelativePath":{{"DecodedUrl":"{path}"}}}}"#)
}

fn parent_web_json(url: &str) -> String {
    format!(r#"{{"ParentWeb":{{"Url":"{url}"}}}}"#)
}

#[tokio::test]
async fn parses_defaults_from_file_across_folders() {
    let transport = RecordingTransport::default()
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue(DEFAULTS_XML);
    let client = Client::new(&transport);

    let defaults = get_default_column_values(&client, SITE, "Documents")
        .await
        .unwrap();

    assert_eq!(defaults.len(), 3);
    assert_eq!(defaults[0].name, "Status");
    assert_eq!(defaults[0].path, "/sites/team/Shared Documents");
    assert_eq!(defaults[0].value, "Active");
    assert_eq!(defaults[1].name, "Region");
    assert_eq!(defaults[1].value, "5;#North|guid-1");
    assert_eq!(defaults[2].path, "/sites/team/Shared Documents/Sub Folder");
    assert_eq!(defaults[2].value, "Inactive");
}

#[tokio::test]
async fn missing_defaults_file_returns_empty() {
    let transport = RecordingTransport::default()
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue_err(Error::Surface(
            Surface::SpRest,
            "HTTP 404: not found".into(),
        ));
    let client = Client::new(&transport);

    let defaults = get_default_column_values(&client, SITE, "Documents")
        .await
        .unwrap();
    assert!(defaults.is_empty());
}

#[tokio::test]
async fn non_404_error_propagates() {
    let transport = RecordingTransport::default()
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue_err(Error::Surface(Surface::SpRest, "HTTP 500: boom".into()));
    let client = Client::new(&transport);

    let err = get_default_column_values(&client, SITE, "Documents")
        .await
        .unwrap_err();
    assert!(matches!(err, Error::Surface(_, msg) if msg.contains("500")));
}

#[tokio::test]
async fn get_file_url_is_built_correctly() {
    let transport = RecordingTransport::default()
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue("<MetadataDefaults></MetadataDefaults>");
    let client = Client::new(&transport);

    get_default_column_values(&client, SITE, "Documents")
        .await
        .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls.len(), 3);
    assert!(calls[2].url.contains("getFileByServerRelativePath"));
    assert!(calls[2].url.ends_with("/$value"));
    assert!(calls[2].url.contains("Forms"));
    assert!(calls[2].url.contains("LocationBasedDefaults"));
}

#[test]
fn builds_xml_for_single_and_multi_and_taxonomy_values() {
    let defaults = vec![
        FieldDefaultInput {
            name: "Status".into(),
            path: "/sites/team/Shared Documents".into(),
            value: ColumnDefaultValue::Single("Active".into()),
        },
        FieldDefaultInput {
            name: "Tags".into(),
            path: "/sites/team/Shared Documents".into(),
            value: ColumnDefaultValue::Multi(vec!["A".into(), "B".into()]),
        },
        FieldDefaultInput {
            name: "Region".into(),
            path: "/sites/team/Shared Documents".into(),
            value: ColumnDefaultValue::Taxonomy(TaxonomyTermRef {
                wss_id: "5".into(),
                term_name: "North".into(),
                term_id: "guid-1".into(),
            }),
        },
    ];
    let field_types = vec![
        ("Status".to_string(), "Text".to_string()),
        ("Tags".to_string(), "MultiChoice".to_string()),
        ("Region".to_string(), "Taxonomy".to_string()),
    ];

    let xml = build_defaults_xml(&defaults, &field_types).unwrap();

    assert!(xml.starts_with("<MetadataDefaults>"));
    assert!(xml.contains("<a href=\"/sites/team/Shared%20Documents\">"));
    assert!(xml.contains("<DefaultValue FieldName=\"Status\">Active</DefaultValue>"));
    assert!(xml.contains("<DefaultValue FieldName=\"Tags\">A;B</DefaultValue>"));
    assert!(xml.contains("<DefaultValue FieldName=\"Region\">5;#North|guid-1</DefaultValue>"));
}

#[test]
fn multi_value_rejected_for_single_only_field_type() {
    let defaults = vec![FieldDefaultInput {
        name: "Status".into(),
        path: "/sites/team/Shared Documents".into(),
        value: ColumnDefaultValue::Multi(vec!["A".into(), "B".into()]),
    }];
    let field_types = vec![("Status".to_string(), "Text".to_string())];

    let err = build_defaults_xml(&defaults, &field_types).unwrap_err();
    assert!(matches!(err, Error::Unsupported(_)));
}

#[test]
fn unknown_field_errors() {
    let defaults = vec![FieldDefaultInput {
        name: "Ghost".into(),
        path: "/sites/team/Shared Documents".into(),
        value: ColumnDefaultValue::Single("x".into()),
    }];
    let err = build_defaults_xml(&defaults, &[]).unwrap_err();
    assert!(matches!(err, Error::Unsupported(_)));
}

#[tokio::test]
async fn set_default_column_values_posts_defaults_file() {
    let transport = RecordingTransport::default()
        .queue(r#"[{"InternalName":"Status","TypeAsString":"Text"}]"#)
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue("");
    let client = Client::new(&transport);

    set_default_column_values(
        &client,
        SITE,
        "Documents",
        &[FieldDefaultInput {
            name: "Status".into(),
            path: "/sites/team/Shared Documents".into(),
            value: ColumnDefaultValue::Single("Active".into()),
        }],
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    assert_eq!(calls.len(), 4);
    assert!(calls[0].url.contains("/fields?"));
    assert!(calls[3].method == "POST");
    assert!(calls[3]
        .url
        .contains("files/add(overwrite=true,url='client_LocationBasedDefaults.html')"));
    let body = String::from_utf8(calls[3].body.clone().unwrap()).unwrap();
    assert!(body.contains("<DefaultValue FieldName=\"Status\">Active</DefaultValue>"));
}

#[tokio::test]
async fn get_folder_defaults_filters_by_path() {
    let transport = RecordingTransport::default()
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue(DEFAULTS_XML);
    let client = Client::new(&transport);

    let defaults = get_folder_default_column_values(
        &client,
        SITE,
        "Documents",
        "/sites/team/Shared Documents/Sub Folder",
    )
    .await
    .unwrap();

    assert_eq!(defaults.len(), 1);
    assert_eq!(defaults[0].name, "Status");
    assert_eq!(defaults[0].value, "Inactive");
}

#[tokio::test]
async fn set_folder_defaults_merges_with_existing() {
    // get_default_column_values (3 calls) inside set_folder_default_column_values,
    // then get_field_types + root_folder + parent_web + POST inside set_default_column_values.
    let transport = RecordingTransport::default()
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue(DEFAULTS_XML)
        .queue(r#"[{"InternalName":"Status","TypeAsString":"Text"},{"InternalName":"Region","TypeAsString":"Taxonomy"}]"#)
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue("");
    let client = Client::new(&transport);

    set_folder_default_column_values(
        &client,
        SITE,
        "Documents",
        "/sites/team/Shared Documents",
        &[FieldDefaultProps {
            name: "Status".into(),
            value: ColumnDefaultValue::Single("Draft".into()),
        }],
        true,
    )
    .await
    .unwrap();

    let calls = transport.calls.lock().unwrap();
    let post = calls.last().unwrap();
    let body = String::from_utf8(post.body.clone().unwrap()).unwrap();
    // Status was updated in place, Region (taxonomy, carried over raw) and
    // the untouched sub-folder default survive the merge.
    assert!(body.contains("<DefaultValue FieldName=\"Status\">Draft</DefaultValue>"));
    assert!(body.contains("<DefaultValue FieldName=\"Region\">5;#North|guid-1</DefaultValue>"));
    assert!(body.contains("Sub%20Folder"));
}

#[tokio::test]
async fn clear_folder_defaults_drops_that_folder_only() {
    let transport = RecordingTransport::default()
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue(DEFAULTS_XML)
        .queue(r#"[{"InternalName":"Status","TypeAsString":"Text"},{"InternalName":"Region","TypeAsString":"Taxonomy"}]"#)
        .queue(&root_folder_json("/sites/team/Shared Documents"))
        .queue(&parent_web_json(SITE))
        .queue("");
    let client = Client::new(&transport);

    clear_folder_default_column_values(&client, SITE, "Documents", "/sites/team/Shared Documents")
        .await
        .unwrap();

    let calls = transport.calls.lock().unwrap();
    let body = String::from_utf8(calls.last().unwrap().body.clone().unwrap()).unwrap();
    // The root folder's own defaults (Status/Region) are dropped, but the
    // sub-folder keeps its own unrelated "Status" default untouched.
    assert!(!body.contains("Active"));
    assert!(!body.contains("Region"));
    assert!(body.contains("Sub%20Folder"));
    assert!(body.contains("Inactive"));
}
