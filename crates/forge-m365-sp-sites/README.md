# forge-m365-sp-sites

SharePoint site-collection (`_api/site`) operations. Hand-ported from PnPjs,
per `SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this
surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\sites` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Site.get`, `Site.getRootWeb` |

Only read (GET) operations are ported so far: `get_site`, `get_root_web_url`,
`get_document_libraries`, `get_web_url_from_page_url`. Site creation/deletion,
`openWebById`, `exists`, and logo management from the PnPjs `_Site` class are
not yet ported — those are POST-based and need the request-body /
`Content-Type` handling that `forge-m365-sp-lists` CRUD will add to
`AuthedTransport`, so they're deferred rather than half-done here.

## Verification status

Wiring (operation registration + response parsing, including both the
`{"GetX": ...}`-wrapped and bare-value response shapes SharePoint uses for
`_api/sp.web.*` OData functions) is covered by `tests/get_site.rs` against a
scripted transport. Compiled ≠ verified against a live tenant — see
`AGENTS.md` for the live-verification record before relying on this against
production data.
