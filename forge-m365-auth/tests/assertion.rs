use base64::engine::general_purpose::URL_SAFE_NO_PAD;
use base64::Engine;
use forge_m365_auth::{AuthError, TokenClient};
use rsa::pkcs8::EncodePrivateKey;
use rsa::RsaPrivateKey;

fn test_key_pem() -> String {
    let mut rng = rand::thread_rng();
    let key = RsaPrivateKey::new(&mut rng, 2048).expect("keygen");
    key.to_pkcs8_pem(rsa::pkcs8::LineEnding::LF)
        .expect("pem")
        .to_string()
}

#[test]
fn assertion_is_three_base64url_segments() {
    let pem = test_key_pem();
    let client = TokenClient::with_certificate("tenant", &pem, "client-id").expect("client");
    let jwt = client.signed_jwt("tenant").expect("jwt");
    let parts: Vec<&str> = jwt.split('.').collect();
    assert_eq!(parts.len(), 3);
    for p in &parts {
        assert!(URL_SAFE_NO_PAD.decode(p).is_ok(), "segment not b64url: {p}");
    }

    let header: serde_json::Value =
        serde_json::from_slice(&URL_SAFE_NO_PAD.decode(parts[0]).unwrap()).unwrap();
    assert_eq!(header["alg"], "RS256");

    let claims: serde_json::Value =
        serde_json::from_slice(&URL_SAFE_NO_PAD.decode(parts[1]).unwrap()).unwrap();
    assert_eq!(
        claims["aud"],
        "https://login.microsoftonline.com/tenant/oauth2/v2.0/token"
    );
    assert_eq!(claims["iss"], "client-id");
}

#[test]
fn bad_pem_is_rejected() {
    assert!(matches!(
        TokenClient::with_certificate("t", "not a pem", "c"),
        Err(AuthError::Assertion(_))
    ));
}
