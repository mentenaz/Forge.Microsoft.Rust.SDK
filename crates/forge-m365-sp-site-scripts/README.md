# forge-m365-sp-site-scripts

SharePoint tenant site-script operations: list, create, inspect, update,
delete site scripts, extract a site script from an existing list/web, and
execute a single site-script action. Hand-ported from PnPjs, per `SPEC.md`
§4 (SharePoint REST has no reliable CSDL coverage, so this surface is ported
rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\site-scripts\types.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `SiteScripts.getSiteScripts/createSiteScript/getSiteScriptMetadata/deleteSiteScript/updateSiteScript/getSiteScriptFromList/getSiteScriptFromWeb/executeSiteScriptAction` |
| PnPjs | `E:\sources\pnpjs\packages\sp\utils\encode-path-str.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `encodePath` (title/description URL-embedded query encoding) |

Unlike the classic `_api/web/...` collection surface most other `sp-*`
crates model, every site-script operation is a POST to the static utility
endpoint `_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.<MethodName>`.
This is the same service family as `forge-m365-sp-site-designs` (a
companion crate, kept separate per this repo's fine-grained-workspace rule)
and, per Microsoft's documentation, most of these calls must be made against
the **tenant admin site** (`https://<tenant>-admin.sharepoint.com`), not an
arbitrary site collection.

`createSiteScript`'s title/description are not sent in the JSON body — PnPjs
embeds them as literal, percent-encoded query-string parameters on the action
URL itself (`CreateSiteScript(Title=@title,Description=@desc)?@title='...'&@desc='...'`);
this crate reproduces that exactly via a local `encode_query_value` helper
that doubles embedded `'` then percent-encodes, matching PnPjs's `encodePath`.

`get_site_script_from_list`/`get_site_script_from_web` and
`execute_site_script_action` are typed per the PnPjs interfaces, but their
exact wire response shape (verbose-OData wrapping, if any) is unverified —
see below.

## Verification status

Wiring — URL construction (including the query-string title/description
encoding), JSON body shape, and response parsing — is covered by
`tests/site_scripts.rs` against a recording transport. Compiled ≠ verified
against a live tenant — see `AGENTS.md` for the live-verification record
before relying on this against production data. `create_site_script`,
`update_site_script`, and `delete_site_script` mutate tenant-wide site-script
state once run live.
