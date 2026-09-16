# forge-m365-sp-related-items

SharePoint `SP.RelatedItemManager` operations: read related items for a list
item and add/delete related-item links (optionally with a reverse link on
the target). Hand-ported from PnPjs, per `SPEC.md` §4 (SharePoint REST has
no reliable CSDL coverage, so this surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\related-items` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_RelatedItemManager.getRelatedItems/getPageOneRelatedItems/addSingleLink/addSingleLinkToUrl/addSingleLinkFromUrl/deleteSingleLink` |

This is the entire PnPjs `related-items` package's public surface — ported
completely. PnPjs exposes this manager as `web.relatedItems`; this crate
exposes the same operations as free functions taking a `site_url` instead,
matching this workspace's convention (see `sp-appcatalog`, `sp-features`,
etc.) of not modeling PnPjs's fluent object graph.

## Verification status

Wiring — URL construction, POST body shape, and response parsing — is
covered by `tests/related_items.rs` against a recording transport. Compiled
≠ verified against a live tenant — see `AGENTS.md` for the
live-verification record before relying on this against production data,
and note that `add_single_link*`/`delete_single_link` mutate real data once
run live.
