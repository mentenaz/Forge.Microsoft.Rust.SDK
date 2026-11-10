use forge_m365_core::{Client, Result, Surface, Transport};
use forge_m365_sp_navigation::{
    add_navigation_node, get_menu_state, get_navigation_nodes, NavigationArea,
};
use serde_json::Value;
use std::sync::Mutex;

#[derive(Default)]
struct RecordingTransport {
    response: Mutex<Option<String>>,
    last_url: Mutex<Option<String>>,
    last_body: Mutex<Option<Vec<u8>>>,
}

impl RecordingTransport {
    fn with_response(response: &str) -> Self {
        Self {
            response: Mutex::new(Some(response.to_string())),
            last_url: Mutex::new(None),
            last_body: Mutex::new(None),
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
        body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        *self.last_url.lock().unwrap() = Some(url.to_string());
        *self.last_body.lock().unwrap() = body.map(<[u8]>::to_vec);
        let body = self.response.lock().unwrap().clone().unwrap_or_default();
        Box::pin(async move { Ok(body.into_bytes()) })
    }
}

const NODE_JSON: &str = r#"{"Id":1,"Title":"Home","Url":"/sites/team","IsVisible":true,"IsExternal":false,"IsDocLib":false}"#;

#[tokio::test]
async fn get_nodes_builds_quicklaunch_and_topnav_urls() {
    let transport = RecordingTransport::with_response(&format!("[{NODE_JSON}]"));
    let client = Client::new(&transport);

    let nodes = get_navigation_nodes(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        NavigationArea::QuickLaunch,
    )
    .await
    .unwrap();
    assert_eq!(nodes.len(), 1);
    assert_eq!(nodes[0].title, "Home");
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("navigation/quicklaunch"));

    get_navigation_nodes(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        NavigationArea::TopNavigationBar,
    )
    .await
    .unwrap();
    assert!(transport
        .last_url
        .lock()
        .unwrap()
        .as_ref()
        .unwrap()
        .ends_with("navigation/topnavigationbar"));
}

#[tokio::test]
async fn add_node_sends_correct_body() {
    let transport = RecordingTransport::with_response(NODE_JSON);
    let client = Client::new(&transport);

    add_navigation_node(
        &client,
        "https://contoso.sharepoint.com/sites/team",
        NavigationArea::QuickLaunch,
        "Home",
        "/sites/team",
        true,
    )
    .await
    .unwrap();

    let sent: Value =
        serde_json::from_slice(transport.last_body.lock().unwrap().as_ref().unwrap()).unwrap();
    assert_eq!(sent["Title"], "Home");
    assert_eq!(sent["Url"], "/sites/team");
    assert_eq!(sent["IsVisible"], true);
}

#[tokio::test]
async fn get_menu_state_parses_wrapped_nested_nodes() {
    let json = r#"{"StartingNodeTitle":"Home","Nodes":{"results":[{"Key":"root","Title":"Home","SimpleUrl":"/sites/team","IsHidden":false,"Nodes":{"results":[{"Key":"child","Title":"Sub","SimpleUrl":"/sites/team/sub","IsHidden":false,"Nodes":{"results":[]}}]}}]}}"#;
    let transport = RecordingTransport::with_response(json);
    let client = Client::new(&transport);

    let state = get_menu_state(&client, "https://contoso.sharepoint.com/sites/team", None)
        .await
        .unwrap();

    assert_eq!(state.nodes.len(), 1);
    assert_eq!(state.nodes[0].title, "Home");
    assert_eq!(state.nodes[0].nodes.len(), 1);
    assert_eq!(state.nodes[0].nodes[0].title, "Sub");

    let url = transport.last_url.lock().unwrap().clone().unwrap();
    assert!(url.ends_with("/_api/navigation/MenuState"));
}
