# forge-m365-sp-hubsites

SharePoint hub-site read operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this
surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\hubsites` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `HubSites` collection reads, `HubSites.getById` |

`HubSiteInfo.id` maps to the JSON field `ID` (not `Id`) — that's
SharePoint's actual wire casing for this entity, not a transcription typo.

Not yet ported: `HubSite.getSite` (resolving the associated `ISite`
instance — straightforward to add once needed, just calls back into
`sp-sites`) and the tenant-admin hub-site management operations
(`packages/sp-admin`, out of scope for this crate).

## Verification status

Wiring — URL construction (including the `GetById?hubSiteId='...'` query
form) and response parsing — is covered by `tests/hubsites.rs` against a
recording transport. Compiled ≠ verified against a live tenant — see
`AGENTS.md` for the live-verification record before relying on this against
production data.
