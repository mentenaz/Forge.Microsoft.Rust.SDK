# forge-m365-sp-lists

SharePoint list-container and list-item CRUD. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this surface
is ported rather than generated). Completes the "CRUD list items" leg of the
M1 vertical slice (`SPEC.md` §8).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\lists` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Lists.getByTitle`, `List.get` |
| PnPjs | `E:\sources\pnpjs\packages\sp\items` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Items` (add, collection read), `Item.update`, `Item.delete` |
| PnPjs | `E:\sources\pnpjs\packages\sp\utils\encode-path-str.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `encodePath` (list-title escaping) |

List/library management (create, delete, `ensure*`), field-level operations,
CAML queries, change tokens, and versions from the PnPjs `_Lists`/`_List`/
`_Item` classes are not yet ported.

## Why this crate required a core change

Item add/update/delete need request headers the read-only `sp-sites` crate
never used: `Content-Type: application/json;odata=verbose` (so SharePoint
recognizes a body's `__metadata.type`) and `IF-MATCH` on update/delete. The
`Transport` trait and `Client::run_ladder` gained a `headers` parameter to
carry these; `AuthedTransport` now sets the verbose content type by default
whenever a body is present.

## Verification status

Wiring — request construction (method, URL, headers, `__metadata` body
shape) and response parsing — is covered by `tests/lists.rs` against a
transport that records what was sent, not just what it returns. Compiled ≠
verified against a live tenant — see `AGENTS.md` for the live-verification
record before relying on this against production data, and note that write
operations (`add_item`/`update_item`/`delete_item`) mutate real data once
run live.
