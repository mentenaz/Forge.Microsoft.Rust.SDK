use crate::assertion::{AssertionError, ClientAssertion};
use forge_m365_core::Surface;
use thiserror::Error;

#[derive(Debug, Clone)]
pub struct TokenResponse {
    pub access_token: String,
    pub expires_in: std::time::Duration,
}

#[derive(Error, Debug)]
pub enum AuthError {
    #[error("client assertion error: {0}")]
    Assertion(#[from] AssertionError),
    #[error("http error: {0}")]
    Http(#[from] reqwest::Error),
    #[error("token endpoint returned {status}: {body}")]
    Endpoint { status: u16, body: String },
    #[error("browser flow error: {0}")]
    Browser(String),
    #[error("device code flow error: {0}")]
    DeviceCode(String),
}

impl From<AuthError> for forge_m365_core::Error {
    fn from(e: AuthError) -> Self {
        match e {
            AuthError::Http(h) => forge_m365_core::Error::Http(h),
            other => forge_m365_core::Error::Surface(Surface::SpRest, other.to_string()),
        }
    }
}

pub const ASSERTION_TYPE: &str = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer";

pub fn scope_for(surface: Surface, tenant_id: &str) -> String {
    match surface {
        Surface::Graph => "https://graph.microsoft.com/.default".to_string(),
        Surface::SpRest | Surface::Search => {
            format!("https://{tenant_id}.sharepoint.com/.default")
        }
    }
}

pub enum ClientAuth {
    Secret(String),
    Certificate(std::sync::Arc<ClientAssertion>),
}

#[derive(Clone)]
pub struct TokenClient {
    http: reqwest::Client,
    tenant_id: String,
    client_id: String,
    auth: std::sync::Arc<ClientAuth>,
}

impl TokenClient {
    pub fn with_secret(
        tenant_id: impl Into<String>,
        client_id: impl Into<String>,
        client_secret: impl Into<String>,
    ) -> Self {
        Self {
            http: reqwest::Client::new(),
            tenant_id: tenant_id.into(),
            client_id: client_id.into(),
            auth: std::sync::Arc::new(ClientAuth::Secret(client_secret.into())),
        }
    }

    pub fn with_certificate(
        tenant_id: impl Into<String>,
        certificate_pem: &str,
        client_id: impl Into<String>,
    ) -> std::result::Result<Self, AuthError> {
        let client_id = client_id.into();
        Ok(Self {
            http: reqwest::Client::new(),
            tenant_id: tenant_id.into(),
            client_id: client_id.clone(),
            auth: std::sync::Arc::new(ClientAuth::Certificate(std::sync::Arc::new(
                ClientAssertion::from_pem(certificate_pem, client_id)?,
            ))),
        })
    }

    pub fn signed_jwt(&self, tenant_id: &str) -> std::result::Result<String, AuthError> {
        match &*self.auth {
            ClientAuth::Certificate(a) => a.signed_jwt(tenant_id).map_err(AuthError::from),
            ClientAuth::Secret(_) => {
                Err(AuthError::Browser("no certificate configured".to_string()))
            }
        }
    }

    pub async fn acquire(&self, surface: Surface) -> std::result::Result<TokenResponse, AuthError> {
        let endpoint = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.tenant_id
        );

        let mut form: Vec<(&str, String)> = vec![("scope", scope_for(surface, &self.tenant_id))];

        match &*self.auth {
            ClientAuth::Secret(secret) => {
                form.push(("grant_type", "client_credentials".to_string()));
                form.push(("client_id", self.client_id.clone()));
                form.push(("client_secret", secret.clone()));
            }
            ClientAuth::Certificate(assertion) => {
                form.push(("grant_type", "client_credentials".to_string()));
                form.push(("client_id", self.client_id.clone()));
                form.push(("client_assertion", assertion.signed_jwt(&self.tenant_id)?));
                form.push(("client_assertion_type", ASSERTION_TYPE.to_string()));
            }
        }

        let resp = self.http.post(&endpoint).form(&form).send().await?;
        let status = resp.status().as_u16();
        let body = resp.text().await?;

        if status != 200 {
            return Err(AuthError::Endpoint { status, body });
        }

        let json: serde_json::Value =
            serde_json::from_str(&body).map_err(|_| AuthError::Endpoint {
                status,
                body: format!("unparseable token response: {body}"),
            })?;

        Ok(TokenResponse {
            access_token: json["access_token"]
                .as_str()
                .ok_or_else(|| AuthError::Endpoint {
                    status: 200,
                    body: "missing access_token".to_string(),
                })?
                .to_string(),
            expires_in: std::time::Duration::from_secs(json["expires_in"].as_u64().unwrap_or(3600)),
        })
    }
}
