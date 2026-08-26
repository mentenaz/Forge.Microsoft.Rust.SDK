use crate::token_client::{AuthError, TokenResponse};
use std::time::{Duration, Instant};

/// Info to show the user so they can complete sign-in on another device or
/// tab. This is the "open this URL" / "poll this code" event `SPEC.md` §5
/// calls for — the host renders it however fits (Tauri window, GPUI toast,
/// CLI println), rather than this crate printing anything itself.
#[derive(Debug, Clone)]
pub struct DeviceCodePrompt {
    pub verification_uri: String,
    pub user_code: String,
    pub message: String,
    pub expires_in: Duration,
}

/// Deliberately synchronous (like `BrowserConfig::acquire_interactive`), not
/// `async fn`: polling needs a sleep between attempts, and sleeping via a
/// specific async runtime would tie this library crate to that runtime
/// (`SPEC.md` §6 forbids owning a runtime). Callers on an async host run
/// this on a blocking thread (e.g. `tokio::task::spawn_blocking`).
#[derive(Clone)]
pub struct DeviceCodeConfig {
    pub tenant_id: String,
    pub client_id: String,
    pub scopes: String,
}

impl DeviceCodeConfig {
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

    /// Starts the device code flow, invokes `on_prompt` once with the
    /// verification URL/code to show the user, then polls until the user
    /// completes sign-in, declines, or the code expires.
    pub fn acquire(
        &self,
        on_prompt: impl FnOnce(DeviceCodePrompt),
    ) -> Result<TokenResponse, AuthError> {
        let http = reqwest::blocking::Client::new();

        let devicecode_endpoint = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/devicecode",
            self.tenant_id
        );
        let resp = http
            .post(&devicecode_endpoint)
            .form(&[
                ("client_id", self.client_id.as_str()),
                ("scope", self.scopes.as_str()),
            ])
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
                body: format!("unparseable device code response: {body}"),
            })?;

        let device_code = json["device_code"]
            .as_str()
            .ok_or_else(|| AuthError::Endpoint {
                status: 200,
                body: "missing device_code".to_string(),
            })?
            .to_string();
        let expires_in = Duration::from_secs(json["expires_in"].as_u64().unwrap_or(900));
        let mut interval = Duration::from_secs(json["interval"].as_u64().unwrap_or(5));

        on_prompt(DeviceCodePrompt {
            verification_uri: json["verification_uri"]
                .as_str()
                .unwrap_or("https://microsoft.com/devicelogin")
                .to_string(),
            user_code: json["user_code"].as_str().unwrap_or_default().to_string(),
            message: json["message"].as_str().unwrap_or_default().to_string(),
            expires_in,
        });

        let token_endpoint = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.tenant_id
        );
        let deadline = Instant::now() + expires_in;

        loop {
            std::thread::sleep(interval);
            if Instant::now() >= deadline {
                return Err(AuthError::DeviceCode(
                    "code expired before sign-in completed".to_string(),
                ));
            }

            let resp = http
                .post(&token_endpoint)
                .form(&[
                    ("grant_type", "urn:ietf:params:oauth:grant-type:device_code"),
                    ("client_id", self.client_id.as_str()),
                    ("device_code", device_code.as_str()),
                ])
                .send()
                .map_err(AuthError::Http)?;
            let status = resp.status().as_u16();
            let body = resp.text().map_err(AuthError::Http)?;

            if status == 200 {
                let json: serde_json::Value =
                    serde_json::from_str(&body).map_err(|_| AuthError::Endpoint {
                        status,
                        body: format!("unparseable token response: {body}"),
                    })?;
                return Ok(TokenResponse {
                    access_token: json["access_token"]
                        .as_str()
                        .ok_or_else(|| AuthError::Endpoint {
                            status: 200,
                            body: "missing access_token".to_string(),
                        })?
                        .to_string(),
                    expires_in: Duration::from_secs(json["expires_in"].as_u64().unwrap_or(3600)),
                });
            }

            let error_json: serde_json::Value =
                serde_json::from_str(&body).unwrap_or(serde_json::Value::Null);
            match classify_poll_error(status, error_json["error"].as_str()) {
                PollAction::KeepPolling => {}
                PollAction::SlowDown => interval += Duration::from_secs(5),
                PollAction::Fail => {
                    let desc = error_json["error_description"].as_str().unwrap_or(&body);
                    return Err(AuthError::DeviceCode(desc.to_string()));
                }
            }
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
enum PollAction {
    KeepPolling,
    SlowDown,
    Fail,
}

/// Interprets a non-200 token-endpoint response during device code polling,
/// per the OAuth2 Device Authorization Grant (RFC 8628 §3.5): `authorization_pending`
/// means keep polling, `slow_down` means back off, anything else
/// (`authorization_declined`, `expired_token`, `bad_verification_code`, ...) fails.
fn classify_poll_error(status: u16, error: Option<&str>) -> PollAction {
    if status != 400 {
        return PollAction::Fail;
    }
    match error {
        Some("authorization_pending") => PollAction::KeepPolling,
        Some("slow_down") => PollAction::SlowDown,
        _ => PollAction::Fail,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn authorization_pending_keeps_polling() {
        assert_eq!(
            classify_poll_error(400, Some("authorization_pending")),
            PollAction::KeepPolling
        );
    }

    #[test]
    fn slow_down_backs_off() {
        assert_eq!(
            classify_poll_error(400, Some("slow_down")),
            PollAction::SlowDown
        );
    }

    #[test]
    fn declined_fails() {
        assert_eq!(
            classify_poll_error(400, Some("authorization_declined")),
            PollAction::Fail
        );
    }

    #[test]
    fn expired_fails() {
        assert_eq!(
            classify_poll_error(400, Some("expired_token")),
            PollAction::Fail
        );
    }

    #[test]
    fn non_400_status_fails() {
        assert_eq!(classify_poll_error(500, None), PollAction::Fail);
    }
}
