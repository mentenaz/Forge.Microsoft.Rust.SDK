use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde_json::{Map, Value};
use std::collections::HashMap;

/// A single search result row, flattened from SharePoint's `Key`/`Value` cell
/// array into a plain map. Ported from PnPjs `SearchResults.formatSearchResults`
/// (`packages/sp/search/query.ts`).
pub type SearchResultRow = HashMap<String, String>;

#[derive(Debug, Clone, Default)]
pub struct SearchResults {
    pub elapsed_time: u64,
    pub row_count: u64,
    pub total_rows: u64,
    pub results: Vec<SearchResultRow>,
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-search linked in?")
        })
}

// This body is intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP call
// lives in the public fn below, which looks the entry back up by id.
#[pnp_operation(id = "sp.search.query", primary = Surface::Search, fallback = [])]
#[allow(dead_code)]
async fn search_op() {}

/// Runs a search query. `select_properties` is optional (pass `&[]` for the
/// server default set); SharePoint's search request DTO wraps collection
/// properties as `{"results": [...]}` regardless of the response odata
/// format, which is why this needs its own body construction rather than a
/// plain `serde_json` struct.
///
/// Ported from PnPjs `Search.run` (`packages/sp/search/query.ts`)
/// @ pnpjs `8ee2375d`. Satisfies the `_api/search/query` requirement
/// `SPEC.md` §7 names for communication-site enumeration — SharePoint also
/// exposes a lighter GET-based `_api/search/query` endpoint for the same
/// feature, but `postquery` is what PnPjs uses, so that's what's ported here.
pub async fn search(
    client: &Client<'_>,
    site_url: &str,
    query_text: &str,
    row_limit: Option<u32>,
    select_properties: &[&str],
) -> Result<SearchResults> {
    let mut request = Map::new();
    request.insert(
        "Querytext".to_string(),
        Value::String(query_text.to_string()),
    );
    if let Some(n) = row_limit {
        request.insert("RowLimit".to_string(), Value::from(n));
    }
    if !select_properties.is_empty() {
        let mut wrapped = Map::new();
        wrapped.insert(
            "results".to_string(),
            Value::Array(
                select_properties
                    .iter()
                    .map(|s| Value::String((*s).to_string()))
                    .collect(),
            ),
        );
        request.insert("SelectProperties".to_string(), Value::Object(wrapped));
    }
    let mut payload_body = Map::new();
    payload_body.insert("request".to_string(), Value::Object(request));

    let url = format!("{}/_api/search/postquery", site_url.trim_end_matches('/'));
    let payload = serde_json::to_vec(&Value::Object(payload_body))?;
    let bytes = client
        .run_ladder(entry("sp.search.query"), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(parse_search_response(&bytes)?)
}

/// Some SharePoint versions wrap collection-typed properties as
/// `{"results": [...]}`; others return the bare array. Mirrors the defensive
/// check PnPjs's `formatSearchResults` does.
fn unwrap_array(value: &Value) -> Vec<Value> {
    value
        .get("results")
        .unwrap_or(value)
        .as_array()
        .cloned()
        .unwrap_or_default()
}

fn parse_search_response(bytes: &[u8]) -> serde_json::Result<SearchResults> {
    let root: Value = serde_json::from_slice(bytes)?;
    // Some responses wrap the whole payload as {"postquery": {...}}.
    let root = root.get("postquery").cloned().unwrap_or(root);

    let relevant = &root["PrimaryQueryResult"]["RelevantResults"];
    let rows = unwrap_array(&relevant["Table"]["Rows"]);

    let results = rows
        .iter()
        .map(|row| {
            unwrap_array(&row["Cells"])
                .iter()
                .map(|cell| {
                    let key = cell["Key"].as_str().unwrap_or_default().to_string();
                    let value = cell["Value"]
                        .as_str()
                        .map(str::to_string)
                        .unwrap_or_else(|| cell["Value"].to_string());
                    (key, value)
                })
                .collect::<SearchResultRow>()
        })
        .collect();

    Ok(SearchResults {
        elapsed_time: root["ElapsedTime"].as_u64().unwrap_or(0),
        row_count: relevant["RowCount"].as_u64().unwrap_or(0),
        total_rows: relevant["TotalRows"].as_u64().unwrap_or(0),
        results,
    })
}
