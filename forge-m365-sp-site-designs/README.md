# forge-m365-sp-site-designs

SharePoint site-design operations: create/apply/update/delete site designs,
manage their view-rights principals, kick off and poll asynchronous
site-design runs. Hand-ported from PnPjs, per `SPEC.md` §4 (SharePoint REST
has no reliable CSDL coverage, so this surface is ported rather than
generated).

Unlike most `sp-*` crates, this isn't the classic `_api/web` collection
style — every operation is a bodiless-URL POST to the static
`SiteScriptUtility` service:
`{site}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.<MethodName>`,
with all parameters carried in the JSON body (wrapper shape varies per
method — some wrap a single object under a named key like `info` or
`updateInfo`, others send bare named fields).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\site-designs` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_SiteDesigns` — create/apply/get/update/delete, rights grant/revoke, task add/get, run get/get-status |

All fourteen methods on PnPjs's `_SiteDesigns` class are ported:
`createSiteDesign`, `applySiteDesign`, `getSiteDesigns`,
`getSiteDesignMetadata`, `updateSiteDesign`, `deleteSiteDesign`,
`getSiteDesignRights`, `grantSiteDesignRights`, `revokeSiteDesignRights`,
`addSiteDesignTask`, `addSiteDesignTaskToCurrentWeb`, `getSiteDesignTask`,
`getSiteDesignRun`, `getSiteDesignRunStatus`.

`get_site_design_task` mirrors PnPjs's own quirk: SharePoint keeps returning
a (near-empty) object after the task finishes, so PnPjs checks for the
presence of the `ID` field and returns `null` once it's gone; this crate
does the same, returning `Option<SiteDesignTask>`.

`TemplateDesignType`, `ListDesignColor`, and `ListDesignIcon` are kept as raw
`i64` discriminants on `SiteDesignCreationInfo` rather than typed Rust enums
(no `serde_repr`-style dependency is in the workspace yet, and adding one for
three lightly-used enums didn't seem worth it) — see PnPjs's enum definitions
in `packages/sp/site-designs/types.ts` for the discriminant meanings if you
need to set them.

## Scope note

Site designs are a tenant-admin-level feature: creating/updating/deleting
designs and granting/revoking rights require SharePoint admin or the
tenant's Site Script feature enabled on the calling site, and
`applySiteDesign`/task operations act on a *target* site collection distinct
from the one hosting the utility endpoint — same "utility is called from one
place, acts on another" shape as `sp-appcatalog`'s tenant-app-catalog
requirement.

## Verification status

Wiring — URL construction for every method and JSON body shape (including
the varying wrapper keys) — is covered by `tests/site_designs.rs` against a
recording transport that inspects the request body, not just the mocked
response. Compiled ≠ verified against a live tenant — see `AGENTS.md` for
the live-verification record before relying on this against production
data. `createSiteDesign`/`updateSiteDesign`/`deleteSiteDesign`/
`applySiteDesign`/rights grant-revoke all mutate tenant-wide state once run
live.
