use crate::token_client::{AuthError, TokenResponse};
use base64::engine::general_purpose::URL_SAFE_NO_PAD;
use base64::Engine;
use rand::RngCore;
use rsa::sha2::{Digest, Sha256};
use std::io::{BufRead, BufReader, Write};
use std::net::TcpListener;

#[derive(Clone)]
pub struct BrowserConfig {
    pub tenant_id: String,
    pub client_id: String,
    pub scopes: String,
}

fn b64url(bytes: &[u8]) -> String {
    URL_SAFE_NO_PAD.encode(bytes)
}

fn random_b64url() -> String {
    let mut bytes = [0u8; 32];
    rand::thread_rng().fill_bytes(&mut bytes);
    b64url(&bytes)
}

fn pct(s: &str) -> String {
    percent_encoding::utf8_percent_encode(s, percent_encoding::NON_ALPHANUMERIC).to_string()
}

impl BrowserConfig {
    pub fn new(
        tenant_id: impl Into<String>,
        client_id: impl Into<String>,
        scopes: impl Into<String>,
    ) -> Self {
        Self {
            tenant_id: tenant_id.into(),
            client_id: client_id.into(),
            scopes: scopes.into(),
        }
    }

    pub fn acquire_interactive(&self) -> Result<TokenResponse, AuthError> {
        let listener =
            TcpListener::bind("localhost:0").map_err(|e| AuthError::Browser(e.to_string()))?;
        let port = listener
            .local_addr()
            .map_err(|e| AuthError::Browser(e.to_string()))?
            .port();

        let redirect_uri = format!("http://localhost:{port}/");
        let state = random_b64url();
        let verifier = random_b64url();
        let challenge = b64url(&Sha256::digest(verifier.as_bytes()));

        let auth_url = format!(
            "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize\
?client_id={client}&response_type=code&redirect_uri={redirect}&response_mode=query\
&scope={scope}&state={state}&code_challenge={challenge}&code_challenge_method=S256&prompt=select_account",
            tenant = self.tenant_id,
            client = pct(&self.client_id),
            redirect = pct(&redirect_uri),
            scope = pct(&self.scopes),
            state = state,
            challenge = challenge,
        );

        open::that_detached(&auth_url)
            .map_err(|e| AuthError::Browser(format!("cannot open browser: {e}")))?;

        eprintln!("waiting for sign-in in your browser (listening on {redirect_uri})...");

        let (stream, _) = listener
            .accept()
            .map_err(|e| AuthError::Browser(e.to_string()))?;
        let mut reader = BufReader::new(
            stream
                .try_clone()
                .map_err(|e| AuthError::Browser(e.to_string()))?,
        );

        let mut request_line = String::new();
        reader
            .read_line(&mut request_line)
            .map_err(|e| AuthError::Browser(e.to_string()))?;

        let mut line = String::new();
        loop {
            line.clear();
            match reader.read_line(&mut line) {
                Ok(0) | Err(_) => break,
                Ok(_) if line == "\r\n" || line == "\n" => break,
                Ok(_) => {}
            }
        }

        let response_page = "HTTP/1.1 200 OK\r\nContent-Type: text/html\r\nConnection: close\r\n\r\n\
<html><body><h2>Sign-in complete.</h2>You can close this window and return to the app.</body></html>";
        let _ = (&stream).write_all(response_page.as_bytes());
        drop(stream);

        let path = request_line.split_whitespace().nth(1).unwrap_or_default();
        let query = path.split('?').nth(1).unwrap_or_default();
        let params = parse_query(query);
        let get = |k: &str| {
            params
                .iter()
                .find(|(key, _)| key == k)
                .map(|(_, v)| v.clone())
        };

        if get("state").as_deref() != Some(state.as_str()) {
            return Err(AuthError::Browser("state mismatch".into()));
        }
        if let Some(err) = get("error") {
            return Err(AuthError::Browser(format!(
                "authorization failed: {} ({})",
                err,
                get("error_description").unwrap_or_default()
            )));
        }
        let code = get("code")
            .ok_or_else(|| AuthError::Browser("no authorization code in redirect".to_string()))?;

        self.redeem_code(&code, &verifier, &redirect_uri)
    }

    fn redeem_code(
        &self,
        code: &str,
        verifier: &str,
        redirect_uri: &str,
    ) -> Result<TokenResponse, AuthError> {
        let endpoint = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.tenant_id
        );
        let form = [
            ("grant_type", "authorization_code"),
            ("client_id", self.client_id.as_str()),
            ("code", code),
            ("redirect_uri", redirect_uri),
            ("code_verifier", verifier),
        ];
        let resp = reqwest::blocking::Client::new()
            .post(&endpoint)
            .form(&form)
            .send()
            .map_err(AuthError::Http)?;
        let status = resp.status().as_u16();
        let body = resp.text().map_err(AuthError::Http)?;
        if status != 200 {
            return Err(AuthError::Endpoint { status, body });
        }
        let json: serde_json::Value =
            serde_json::from_str(&body).map_err(|_| AuthError::Endpoint {
                status,
                body: format!("unparseable: {body}"),
            })?;
        Ok(TokenResponse {
            access_token: json["access_token"]
                .as_str()
                .unwrap_or_default()
                .to_string(),
            expires_in: std::time::Duration::from_secs(json["expires_in"].as_u64().unwrap_or(3600)),
        })
    }
}

fn parse_query(q: &str) -> Vec<(String, String)> {
    q.split('&')
        .filter_map(|pair| {
            let (k, v) = pair.split_once('=')?;
            let dec = |s: &str| {
                percent_encoding::percent_decode_str(&s.replace('+', " "))
                    .decode_utf8()
                    .map(|c| c.to_string())
                    .unwrap_or_default()
            };
            Some((dec(k), dec(v)))
        })
        .collect()
}
