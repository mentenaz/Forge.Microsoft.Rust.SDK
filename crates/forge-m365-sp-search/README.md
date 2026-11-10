# forge-m365-sp-search

SharePoint search (`_api/search/postquery`). Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this surface
is ported rather than generated).

Built to close the gap `SPEC.md` §7 names explicitly: communication sites
(no M365 group anchor) aren't fully visible via Graph, and enumerating them
requires `_api/search/query`. SharePoint also exposes a lighter GET-based
`_api/search/query` endpoint for the same feature; `postquery` (POST) is
what PnPjs uses, so that's what's ported here.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\search\query.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Search.run`, `SearchResults` (response flattening) |

Not yet ported: `SearchQueryBuilder`, `getPage` (pagination), suggestions
(`packages/sp/search/suggest.ts`), refiners, sort lists, and most of
`ISearchQuery`'s optional fields beyond `Querytext`/`RowLimit`/
`SelectProperties`.

## A wire-format quirk worth knowing

SharePoint's search request/response DTOs wrap collection-typed properties
as `{"results": [...]}` rather than a bare JSON array — independent of the
`odata=nometadata`/`odata=verbose` Accept/Content-Type negotiation that
governs every other `_api/*` endpoint in this SDK. This crate handles both
the wrapped and bare shapes defensively on the response side (mirroring
PnPjs's own defensive check) and always sends the wrapped shape on the
request side for `SelectProperties`.

## Verification status

Wiring — request construction (`Querytext`/`RowLimit`/`SelectProperties`
body shape) and response parsing (both wrapped and bare row/cell arrays) —
is covered by `tests/search.rs` against a recording transport. Compiled ≠
verified against a live tenant — see `AGENTS.md` for the live-verification
record before relying on this against production data.
