#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Surface {
    SpRest,
    Graph,
    Search,
}

#[derive(Debug)]
pub struct Ladder {
    pub primary: Surface,
    pub fallback: &'static [Surface],
}

#[derive(Debug)]
pub struct OperationEntry {
    pub id: &'static str,
    pub operation_path: &'static str,
    pub ladder: Ladder,
}

inventory::collect!(OperationEntry);

pub use inventory;

pub fn registered_operations() -> impl Iterator<Item = &'static OperationEntry> {
    inventory::iter::<OperationEntry>.into_iter()
}

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("operation '{0}' exhausted its escalation ladder")]
    Unsupported(&'static str),
    #[error("http error: {0}")]
    Http(#[from] reqwest::Error),
    #[error("surface {0:?} rejected the request: {1}")]
    Surface(Surface, String),
}

pub type Result<T> = std::result::Result<T, Error>;

pub trait Transport: Send + Sync {
    fn execute(
        &self,
        surface: Surface,
        method: &str,
        url: &str,
        body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + '_>>;
}

pub struct Client<'a> {
    transport: &'a dyn Transport,
}

impl<'a> Client<'a> {
    pub fn new(transport: &'a dyn Transport) -> Self {
        Self { transport }
    }

    pub async fn run_ladder(
        &self,
        entry: &OperationEntry,
        method: &str,
        url: &str,
        body: Option<&[u8]>,
    ) -> Result<Vec<u8>> {
        let mut last_err = None;

        for surface in std::iter::once(&entry.ladder.primary).chain(entry.ladder.fallback.iter()) {
            match self.transport.execute(*surface, method, url, body).await {
                Ok(bytes) => return Ok(bytes),
                Err(Error::Surface(s, msg)) => last_err = Some(Error::Surface(s, msg)),
                Err(e) => return Err(e),
            }
        }

        Err(match last_err {
            Some(e) => e,
            None => Error::Unsupported(entry.id),
        })
    }
}
