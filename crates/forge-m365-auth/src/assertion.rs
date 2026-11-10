use base64::engine::general_purpose::URL_SAFE_NO_PAD;
use base64::Engine;
use rand::RngCore;
use rsa::pkcs1v15::SigningKey;
use rsa::pkcs8::DecodePrivateKey;
use rsa::sha2::{Digest, Sha256};
use rsa::signature::{SignatureEncoding, Signer};
use rsa::RsaPrivateKey;
use serde_json::json;

#[derive(Debug, thiserror::Error)]
pub enum AssertionError {
    #[error("invalid private key PEM: {0}")]
    Key(#[from] rsa::pkcs8::Error),
    #[error("failed to sign client assertion: {0}")]
    Sign(String),
}

pub struct ClientAssertion {
    key: RsaPrivateKey,
    key_id: String,
    client_id: String,
}

fn b64url(bytes: &[u8]) -> String {
    URL_SAFE_NO_PAD.encode(bytes)
}

fn jti() -> String {
    let mut bytes = [0u8; 16];
    rand::thread_rng().fill_bytes(&mut bytes);
    b64url(&bytes)
}

impl ClientAssertion {
    pub fn from_pem(pem: &str, client_id: impl Into<String>) -> Result<Self, AssertionError> {
        let key = RsaPrivateKey::from_pkcs8_pem(pem)?;
        let digest = Sha256::digest(pem.as_bytes());
        Ok(Self {
            key,
            key_id: b64url(&digest),
            client_id: client_id.into(),
        })
    }

    pub fn signed_jwt(&self, tenant_id: &str) -> Result<String, AssertionError> {
        let header = b64url(
            json!({"alg": "RS256", "typ": "JWT", "x5t": self.key_id})
                .to_string()
                .as_bytes(),
        );
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock before 1970")
            .as_secs();
        let claims = json!({
            "aud": format!("https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"),
            "iss": self.client_id,
            "sub": self.client_id,
            "jti": jti(),
            "nbf": now,
            "exp": now + 600,
        });
        let payload = b64url(claims.to_string().as_bytes());
        let signing_input = format!("{header}.{payload}");

        let signature = SigningKey::<Sha256>::new(self.key.clone()).sign(signing_input.as_bytes());

        Ok(format!("{signing_input}.{}", b64url(&signature.to_bytes())))
    }
}
