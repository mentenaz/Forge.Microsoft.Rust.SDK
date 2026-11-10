# forge-m365-sp-utilities

SharePoint `SP.Utilities.Utility` static-method operations: send email,
resolve/search principals, expand groups to their member principals, and
build an invitation email body. Hand-ported from PnPjs, per `SPEC.md` §4
(SharePoint REST has no reliable CSDL coverage, so this surface is ported
rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\sputilities\types.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Utilities.sendEmail/getCurrentUserEmailAddresses/resolvePrincipal/searchPrincipals/createEmailBodyForInvitation/expandGroupsToPrincipals` |
| PnPjs | `E:\sources\pnpjs\packages\sp\types.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `IPrincipalInfo`, `PrincipalType`, `PrincipalSource` |

`PrincipalType`/`PrincipalSource` are ported as plain `u32` associated
constants rather than a Rust `enum`, since PnPjs's own types are bitmasks a
caller may OR together (e.g. `PrincipalType::USER | PrincipalType::SHARE_POINT_GROUP`),
which a C-style enum can't represent.

`EmailProperties::additional_headers` is a plain `BTreeMap<String, String>`
on the Rust side; `send_email` remaps it into SharePoint's expected
`{Key, Value, ValueType: "Edm.String"}` triple array on the wire, matching
PnPjs's `sendEmail` remap (`packages/sp/sputilities/types.ts`).

`get_current_user_email_addresses` and `create_email_body_for_invitation`
return a bare `String`: SharePoint's OData-verbose responses for these
string-returning actions wrap the value as `{"value": "..."}`, while a
classic/non-verbose response returns the bare JSON string — both are
handled.

Not ported: `IWikiPageCreationInfo`'s wiki-page creation flow isn't exposed
by a method on `_Utilities` in this PnPjs version (only the type is
exported, unused elsewhere in the package), so there's no reference call
site to port faithfully.

## Verification status

Wiring — URL construction, request body shape (including the additional-
headers remap), and response parsing (both odata-verbose-wrapped and bare
string shapes) — is covered by `tests/utilities.rs` against a recording
transport. Compiled ≠ verified against a live tenant — see `AGENTS.md` for
the live-verification record before relying on this against production
data. `send_email` sends real mail once run live; `resolve_principal`/
`search_principals`/`expand_groups_to_principals` can add users to the
hidden user info list depending on flags passed.
