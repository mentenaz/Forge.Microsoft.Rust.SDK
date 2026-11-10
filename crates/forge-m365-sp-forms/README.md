# forge-m365-sp-forms

SharePoint list-form read operations (display/edit/new form URLs). Hand-
ported from PnPjs, per `SPEC.md` §4 (SharePoint REST has no reliable CSDL
coverage, so this surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\forms` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Forms.getById`, collection reads |

This is a read-only, 2-operation surface in PnPjs itself — there's no
write API to defer here, unlike most other crates in this SDK.

## Verification status

Wiring — URL construction and response parsing — is covered by
`tests/forms.rs` against a recording transport. Compiled ≠ verified against
a live tenant — see `AGENTS.md` for the live-verification record before
relying on this against production data.
