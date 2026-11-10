# forge-m365-sp-user-custom-actions

SharePoint user custom action operations: list/get/add/update/delete/clear
the script-link and menu-extension custom actions registered on a web or
site collection. Hand-ported from PnPjs, per `SPEC.md` §4 (SharePoint REST
has no reliable CSDL coverage, so this surface is ported rather than
generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\user-custom-actions` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `UserCustomActions.getById/add/clear`, `UserCustomAction.update/delete` |

PnPjs exposes this collection twice — `sp.web.userCustomActions` and
`sp.site.userCustomActions` — with identical shapes differing only by the
`web`/`site` URL segment. This crate models that as a single `Scope { Web,
Site }` parameter on every function rather than duplicating the module.

`UserCustomActionInfo` is a typed struct (PnPjs's `IUserCustomActionInfo` is
well-typed, unlike list items). `Rights` (`SP.BasePermissions`) is modeled
locally as `BasePermissions { high: String, low: String }` rather than
depending on `forge-m365-sp-security`'s permission types, to keep this crate
self-contained; `RegistrationType` and `Scope` are left as raw `i64` (their
PnPjs enum values are documented in doc comments) rather than strict Rust
enums, since an unexpected value from the server should deserialize instead
of erroring.

Not yet ported: nothing from this module — full read/write/clear surface is
covered.

## Verification status

Wiring — URL construction (including the `Scope`-driven `web`/`site`
segment), `IF-MATCH` header on update/delete, and response parsing — is
covered by `tests/user_custom_actions.rs` against a recording transport.
Compiled ≠ verified against a live tenant — see `AGENTS.md` for the
live-verification record before relying on this against production data;
`add`/`update`/`delete`/`clear` all mutate real customization state once run
live.
