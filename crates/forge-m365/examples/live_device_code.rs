use forge_m365_auth::{scope_for, DeviceCodeConfig};
use forge_m365_core::Surface;

fn env(name: &str) -> String {
    std::env::var(name).unwrap_or_else(|_| {
        eprintln!("missing env var {name}");
        std::process::exit(2);
    })
}

fn main() {
    let tenant_id = env("M365_TENANT_ID");
    let client_id = env("M365_CLIENT_ID");

    // Device code is a delegated flow: needs a public-client app registration
    // (same requirement as the interactive browser flow) with a signed-in
    // user available to complete the prompt.
    let config = DeviceCodeConfig::new(
        &tenant_id,
        &client_id,
        scope_for(Surface::Graph, &tenant_id),
    );

    let result = config.acquire(|prompt| {
        println!("{}", prompt.message);
        println!(
            "(or open {} and enter code {})",
            prompt.verification_uri, prompt.user_code
        );
        println!(
            "waiting up to {}s for sign-in...",
            prompt.expires_in.as_secs()
        );
    });

    match result {
        Ok(token) => {
            println!(
                "acquired token, expires in {}s (access_token length: {})",
                token.expires_in.as_secs(),
                token.access_token.len()
            );
        }
        Err(e) => {
            eprintln!("device code flow FAILED: {e}");
            std::process::exit(1);
        }
    }
}
