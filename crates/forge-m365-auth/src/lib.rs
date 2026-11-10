mod assertion;
mod browser;
mod device_code;
mod token_client;
mod transport;

pub use assertion::{AssertionError, ClientAssertion};
pub use browser::BrowserConfig;
pub use device_code::{DeviceCodeConfig, DeviceCodePrompt};
pub use token_client::{
    scope_for, AuthError, ClientAuth, TokenClient, TokenResponse, ASSERTION_TYPE,
};
pub use transport::AuthedTransport;

use std::time::{Duration, Instant};

#[derive(Debug, Clone)]
pub struct CachedToken {
    pub access_token: String,
    acquired_at: Instant,
    expires_in: Duration,
}

impl CachedToken {
    pub fn is_expired(&self) -> bool {
        self.acquired_at.elapsed() >= self.expires_in
    }
}
