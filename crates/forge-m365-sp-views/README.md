# forge-m365-sp-views

SharePoint list-view read operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this
surface is ported rather than generated). Complements `sp-lists`.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\views` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Views.getById/getByTitle`, `ViewFields` reads, collection reads |

`ViewInfo` covers a practical subset of SharePoint's ~35-field `SP.View`
entity (id, title, hidden/default/personal flags, url, query, row limit) —
the same subset-over-completeness approach used for `SiteInfo`/`WebInfo`/
`ListInfo` in the other crates.

Not yet ported: `add`, `update`, `delete`, `renderAsHtml`, `setViewXml`,
and the `ViewFields` write ops (`add`/`move`/`remove`/`removeAll`).

## Verification status

Wiring — URL construction for every accessor and response parsing
(including both the `{"results": [...]}`-wrapped and bare shapes SharePoint
uses for `ViewFields.Items`) — is covered by `tests/views.rs` against a
recording transport. Compiled ≠ verified against a live tenant — see
`AGENTS.md` for the live-verification record before relying on this against
production data.
