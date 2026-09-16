# forge-m365-sp-publishing-sitepageservice

SharePoint `SP.Publishing.SitePageService` read: the current user's unified
M365-group memberships as known to the publishing site-page service.
Hand-ported from PnPjs, per `SPEC.md` §4 (SharePoint REST has no reliable
CSDL coverage, so this surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\publishing-sitepageservice` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_SitePageService.getCurrentUserMemberships` |

This is the entire PnPjs `publishing-sitepageservice` package — it exposes
exactly one method, so this crate ports it exactly and completely.

## Verification status

Wiring — URL construction and response parsing — is covered by
`tests/publishing_sitepageservice.rs` against a recording transport.
Compiled ≠ verified against a live tenant — see `AGENTS.md` for the
live-verification record before relying on this against production data.
