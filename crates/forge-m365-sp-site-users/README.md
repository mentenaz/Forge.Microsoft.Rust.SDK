# forge-m365-sp-site-users

SharePoint site-user read operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this surface
is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\site-users` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `SiteUsers.getById/getByEmail/getByLoginName`, `Web.currentUser`, `Web.siteUsers` |

Only read operations are ported. `add`/`removeById`/`removeByLoginName`/
`update` (write ops), and the `groups` accessor (needs `sp-site-groups`,
not yet ported) from the PnPjs `_SiteUsers`/`_SiteUser` classes are not.

## The `!@v::` login-name quirk

`get_user_by_login_name` builds `siteusers('!@v::{login_name}')` with the
login name interpolated **unencoded** — this is intentional, not a bug.
`!@v::` is SharePoint REST's literal-value indexer syntax; PnPjs uses it
specifically here because claims login names
(`i:0#.f|membership|user@domain.com`) contain `#`/`|`/`@` that survive this
path better than normal percent-encoding through `getByLoginName(...)`
(which Microsoft has deprecated in some tenant configurations). Ported
byte-for-byte from PnPjs rather than "fixed," since the exact behavior here
is a documented tenant-dependent SharePoint quirk, not an encoding bug to
correct.

## Verification status

Wiring — URL construction for every accessor (including confirming
`get_user_by_email` percent-encodes while `get_user_by_login_name`
deliberately does not) and response parsing — is covered by
`tests/site_users.rs` against a recording transport. Compiled ≠ verified
against a live tenant — see `AGENTS.md` for the live-verification record
before relying on this against production data. `get_current_user`
additionally needs a delegated (signed-in user) token to be meaningful; it
has no clear behavior under app-only auth.
