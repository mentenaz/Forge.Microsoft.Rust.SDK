# forge-m365-sp-site-groups

SharePoint site-group read operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this
surface is ported rather than generated). Closes the `.groups` accessor gap
`sp-site-users` deferred.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\site-groups` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `SiteGroups.getById/getByName`, `SiteGroup.users`, collection reads |

`get_group_members` reuses `forge_m365_sp_site_users::SiteUserInfo` rather
than redefining an equivalent struct, since SharePoint returns the same
`SP.User` shape for a group's members as it does for `_api/web/siteusers`.

Not yet ported: `add`, `removeById`/`removeByLoginName`, `update`, and
`setUserAsOwner` (all write ops).

## Verification status

Wiring — URL construction for every accessor and response parsing
(including reusing the `sp-site-users` type for group members) — is
covered by `tests/site_groups.rs` against a recording transport. Compiled ≠
verified against a live tenant — see `AGENTS.md` for the live-verification
record before relying on this against production data.
