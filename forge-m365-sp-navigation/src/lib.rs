use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde::Deserialize;
use serde_json::{Map, Value};

/// Which navigation collection an operation targets, hand-ported from
/// PnPjs's `Navigation.quicklaunch`/`Navigation.topNavigationBar`
/// (`packages/sp/navigation/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NavigationArea {
    QuickLaunch,
    TopNavigationBar,
}

impl NavigationArea {
    fn path_segment(self) -> &'static str {
        match self {
            NavigationArea::QuickLaunch => "quicklaunch",
            NavigationArea::TopNavigationBar => "topnavigationbar",
        }
    }
}

/// Subset of SharePoint's navigation-node fields, hand-ported from PnPjs's
/// `INavNodeInfo` (`packages/sp/navigation/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct NavNodeInfo {
    #[serde(rename = "Id")]
    pub id: i64,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "Url")]
    pub url: String,
    #[serde(rename = "IsVisible")]
    pub is_visible: bool,
    #[serde(rename = "IsExternal")]
    pub is_external: bool,
    #[serde(rename = "IsDocLib")]
    pub is_doc_lib: bool,
}

/// A node in a `getMenuState` response tree, hand-ported from PnPjs's
/// `IMenuNode` (`packages/sp/navigation/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct MenuNode {
    #[serde(rename = "Key")]
    pub key: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "SimpleUrl")]
    pub simple_url: String,
    #[serde(rename = "IsHidden")]
    pub is_hidden: bool,
    #[serde(rename = "Nodes")]
    pub nodes: Vec<MenuNode>,
}

/// A `getMenuState` response, hand-ported from PnPjs's `IMenuNodeCollection`
/// (`packages/sp/navigation/types.ts`).
#[derive(Debug, Clone, Deserialize)]
pub struct MenuNodeCollection {
    #[serde(rename = "StartingNodeTitle")]
    pub starting_node_title: String,
    #[serde(rename = "Nodes")]
    pub nodes: Vec<MenuNode>,
}

/// Some collection-typed properties come back wrapped as `{"results": [...]}`,
/// others bare — same defensive check used in `sp-sites`/`sp-search`. Applied
/// recursively here since `Nodes` nests at every level of the menu tree.
fn normalize_nodes(value: &mut Value) {
    if let Some(inner) = value.get_mut("results").map(Value::take) {
        *value = inner;
    }
    if let Value::Array(items) = value {
        for item in items.iter_mut() {
            if let Some(nodes) = item.get_mut("Nodes") {
                normalize_nodes(nodes);
            }
        }
    }
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-navigation linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.navigation.get_nodes", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_navigation_nodes_op() {}

#[pnp_operation(id = "sp.navigation.add_node", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn add_navigation_node_op() {}

#[pnp_operation(id = "sp.navigation.get_menu_state", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_menu_state_op() {}

/// Gets the nodes in the quicklaunch or top navigation bar.
///
/// Ported from PnPjs `Navigation.quicklaunch`/`topNavigationBar`
/// (`packages/sp/navigation/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_navigation_nodes(
    client: &Client<'_>,
    site_url: &str,
    area: NavigationArea,
) -> Result<Vec<NavNodeInfo>> {
    let url = format!(
        "{}/_api/web/navigation/{}",
        site_url.trim_end_matches('/'),
        area.path_segment()
    );
    let bytes = client
        .run_ladder(entry("sp.navigation.get_nodes"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Adds a node to the quicklaunch or top navigation bar.
///
/// Ported from PnPjs `NavigationNodes.add` (`packages/sp/navigation/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_navigation_node(
    client: &Client<'_>,
    site_url: &str,
    area: NavigationArea,
    title: &str,
    node_url: &str,
    visible: bool,
) -> Result<NavNodeInfo> {
    let url = format!(
        "{}/_api/web/navigation/{}",
        site_url.trim_end_matches('/'),
        area.path_segment()
    );
    let mut body = Map::new();
    body.insert("Title".to_string(), Value::String(title.to_string()));
    body.insert("Url".to_string(), Value::String(node_url.to_string()));
    body.insert("IsVisible".to_string(), Value::Bool(visible));
    let payload = serde_json::to_vec(&Value::Object(body))?;
    let bytes = client
        .run_ladder(
            entry("sp.navigation.add_node"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets a dump of the site's navigation menu state, as used by modern menu
/// controls. `map_provider_name` selects a non-default `SiteMapProvider`;
/// pass `None` for the site's default provider.
///
/// Ported from PnPjs `NavigationService.getMenuState`
/// (`packages/sp/navigation/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_menu_state(
    client: &Client<'_>,
    site_url: &str,
    map_provider_name: Option<&str>,
) -> Result<MenuNodeCollection> {
    let url = format!(
        "{}/_api/navigation/MenuState",
        site_url.trim_end_matches('/')
    );
    let mut body = Map::new();
    body.insert("depth".to_string(), Value::from(10));
    body.insert("menuNodeKey".to_string(), Value::Null);
    body.insert("customProperties".to_string(), Value::Null);
    body.insert(
        "mapProviderName".to_string(),
        map_provider_name.map_or(Value::Null, |s| Value::String(s.to_string())),
    );
    let payload = serde_json::to_vec(&Value::Object(body))?;
    let bytes = client
        .run_ladder(
            entry("sp.navigation.get_menu_state"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    let mut root: Value = serde_json::from_slice(&bytes)?;
    if let Some(nodes) = root.get_mut("Nodes") {
        normalize_nodes(nodes);
    }
    Ok(serde_json::from_value(root)?)
}
