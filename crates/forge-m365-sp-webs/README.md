# forge-m365-sp-webs

SharePoint web (`_api/web`) operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this surface
is ported rather than generated).

This crate replaces the ad hoc `_api/web` call that used to live inline in
`forge-m365/examples/live_web.rs` (now `live_webs.rs`, rewritten to call
this crate instead) — the same pattern `sp-sites` established for `_api/site`.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\webs` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Web.get`, `Web.webs`, `Web.getParentWeb` |

Only read operations are ported so far (`get_web`, `get_subwebs`,
`get_parent_web_url`). Web creation/deletion/update, theming, storage
entities, and web templates from the PnPjs `_Web`/`_Webs` classes are not
yet ported.

## Verification status

Wiring (operation registration + response parsing, including the
`ParentWeb: null` case for root webs) is covered by `tests/webs.rs` against
a scripted transport. Compiled ≠ verified against a live tenant — see
`AGENTS.md` for the live-verification record before relying on this against
production data.
