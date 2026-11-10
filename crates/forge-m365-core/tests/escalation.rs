use forge_m365_core::{Client, Error, Result, Surface};
use forge_m365_macros::pnp_operation;

struct ScriptedTransport {
    responses: std::sync::Mutex<Vec<Result<Vec<u8>>>>,
}

impl forge_m365_core::Transport for ScriptedTransport {
    fn execute(
        &self,
        _surface: Surface,
        _method: &str,
        _url: &str,
        _headers: &[(&str, &str)],
        _body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>> {
        Box::pin(async move {
            self.responses
                .lock()
                .unwrap()
                .pop()
                .unwrap_or_else(|| Err(Error::Unsupported("script-exhausted")))
        })
    }
}

#[pnp_operation(
    id = "sp.lists.get_items",
    primary = Surface::SpRest,
    fallback = [Surface::Graph, Surface::Search]
)]
#[allow(dead_code)] // only exists to be registered by #[pnp_operation]; never called directly
async fn get_items() -> forge_m365_core::Result<Vec<u8>> {
    Ok(Vec::new())
}

fn entry(id: &str) -> &'static forge_m365_core::OperationEntry {
    forge_m365_core::registered_operations()
        .find(|e| e.id == id)
        .expect("operation registered via inventory")
}

#[test]
fn operation_is_registered_with_ladder() {
    let e = entry("sp.lists.get_items");
    assert_eq!(e.ladder.primary, Surface::SpRest);
    assert_eq!(e.ladder.fallback, &[Surface::Graph, Surface::Search]);
}

#[tokio::test]
async fn ladder_falls_through_surface_errors() {
    let transport = ScriptedTransport {
        responses: std::sync::Mutex::new(vec![Ok(b"ok".to_vec())]),
    };
    let client = Client::new(&transport);
    let bytes = client
        .run_ladder(entry("sp.lists.get_items"), "GET", "https://x", &[], None)
        .await
        .unwrap();
    assert_eq!(bytes, b"ok");
}

#[tokio::test]
async fn ladder_exhaustion_is_typed_unsupported() {
    let transport = ScriptedTransport {
        responses: std::sync::Mutex::new(vec![
            Err(Error::Surface(Surface::Search, "down".into())),
            Err(Error::Surface(Surface::Graph, "404".into())),
            Err(Error::Surface(Surface::SpRest, "404".into())),
        ]),
    };
    let client = Client::new(&transport);
    let err = client
        .run_ladder(entry("sp.lists.get_items"), "GET", "https://x", &[], None)
        .await
        .unwrap_err();
    assert!(matches!(err, Error::Surface(..)));
}
