use crate::token_client::{TokenClient, TokenResponse};
use forge_m365_core::{Error, Result, Surface, Transport};
use std::collections::HashMap;
use std::sync::Mutex;

pub struct AuthedTransport {
    tokens: TokenClient,
    http: reqwest::Client,
    cache: Mutex<HashMap<Surface, String>>,
}

impl AuthedTransport {
    pub fn new(tokens: TokenClient) -> Self {
        Self {
            tokens,
            http: reqwest::Client::new(),
            cache: Mutex::new(HashMap::new()),
        }
    }

    async fn bearer(&self, surface: Surface) -> Result<String> {
        if let Some(tok) = self.cache.lock().unwrap().get(&surface) {
            return Ok(tok.clone());
        }
        let tok: TokenResponse = self
            .tokens
            .acquire(surface)
            .await
            .map_err(|e| Error::Surface(surface, e.to_string()))?;
        self.cache
            .lock()
            .unwrap()
            .insert(surface, tok.access_token.clone());
        Ok(tok.access_token)
    }

    pub async fn invalidate(&self, surface: Surface) {
        self.cache.lock().unwrap().remove(&surface);
    }

    pub async fn raw(
        &self,
        surface: Surface,
        method: &str,
        url: &str,
        headers: &[(&str, &str)],
        body: Option<&[u8]>,
    ) -> Result<reqwest::Response> {
        let token = self.bearer(surface).await?;
        let method = reqwest::Method::from_bytes(method.as_bytes())
            .map_err(|e| Error::Surface(surface, e.to_string()))?;
        let mut req = self
            .http
            .request(method, url)
            .bearer_auth(token)
            .header("Accept", "application/json;odata=nometadata");
        // SharePoint REST needs the verbose content type to recognize a body's
        // __metadata.type on write requests; callers (e.g. file upload, which
        // sends raw bytes) can override by passing their own Content-Type,
        // since reqwest's `.header()` appends rather than replaces.
        let caller_sets_content_type = headers
            .iter()
            .any(|(name, _)| name.eq_ignore_ascii_case("content-type"));
        if body.is_some() && !caller_sets_content_type {
            req = req.header("Content-Type", "application/json;odata=verbose");
        }
        for (name, value) in headers {
            req = req.header(*name, *value);
        }
        if let Some(b) = body {
            req = req.body(b.to_vec());
        }
        let resp = req.send().await.map_err(Error::Http)?;
        if resp.status() == 401 {
            self.invalidate(surface).await;
        }
        Ok(resp)
    }
}

impl Transport for AuthedTransport {
    fn execute<'a>(
        &'a self,
        surface: Surface,
        method: &str,
        url: &str,
        headers: &[(&str, &str)],
        body: Option<&[u8]>,
    ) -> std::pin::Pin<Box<dyn std::future::Future<Output = Result<Vec<u8>>> + Send + 'a>> {
        let method = method.to_string();
        let url = url.to_string();
        let headers: Vec<(String, String)> = headers
            .iter()
            .map(|(k, v)| (k.to_string(), v.to_string()))
            .collect();
        let body = body.map(<[u8]>::to_vec);
        Box::pin(async move {
            let headers: Vec<(&str, &str)> = headers
                .iter()
                .map(|(k, v)| (k.as_str(), v.as_str()))
                .collect();
            let resp = self
                .raw(surface, &method, &url, &headers, body.as_deref())
                .await?;
            let status = resp.status();
            let bytes = resp.bytes().await.map_err(Error::Http)?.to_vec();
            if status.is_success() {
                Ok(bytes)
            } else {
                Err(Error::Surface(
                    surface,
                    format!("HTTP {status}: {}", String::from_utf8_lossy(&bytes)),
                ))
            }
        })
    }
}
