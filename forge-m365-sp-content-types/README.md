# forge-m365-sp-content-types

SharePoint content-type operations, at both web and list scope. Hand-ported
from PnPjs, per `SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage,
so this surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\content-types` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `ContentTypes.getById`, `ContentTypes.addAvailableContentType`, collection reads |

`ContentTypeId` mirrors PnPjs's `Id: { StringValue: string }` shape rather
than flattening it to a plain string, since that's the actual wire shape
SharePoint returns.

Not yet ported: creating new content types (`ContentTypes.add`), updating/
deleting a content type, field links, workflow associations, and the
`parent` accessor.

## Verification status

Wiring — request construction (web vs. list-scoped URLs, `addAvailableContentType`
body shape) and response parsing (including the nested `Id.StringValue`
field) — is covered by `tests/content_types.rs` against a recording
transport. Compiled ≠ verified against a live tenant — see `AGENTS.md` for
the live-verification record before relying on this against production
data, and note that `add_content_type_to_list` mutates real data once run
live.
