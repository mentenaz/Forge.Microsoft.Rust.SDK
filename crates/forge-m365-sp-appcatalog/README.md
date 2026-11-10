# forge-m365-sp-appcatalog

SharePoint tenant app-catalog operations: list/inspect available apps and
drive their deploy/retract/install/uninstall/upgrade/remove lifecycle.
Hand-ported from PnPjs, per `SPEC.md` §4 (SharePoint REST has no reliable
CSDL coverage, so this surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\appcatalog` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `AppCatalog.getAppById`, collection reads, `App.deploy/retract/install/uninstall/upgrade/remove` |

`AppInfo` is `serde_json::Map<String, Value>`, not a typed struct — PnPjs's
own `_AppCatalog`/`_App` classes are untyped (`_SPCollection`/`_SPInstance`
with no generic parameter), so there's no reference field list to port
faithfully. Same approach `sp-lists::get_items` takes for the same reason.

Several deploy/retract/install/etc. operations only succeed when called
against the **tenant app catalog web specifically**, not an arbitrary site —
that's a SharePoint requirement, not a limitation of this crate.

Not yet ported: `AppCatalog.add` (upload an `.sppkg`/`.app` package — needs
the `binaryStringRequestBody` header convention, which behaves differently
enough from the plain raw-bytes upload `sp-files`/`sp-attachments` use that
it deserves its own deliberate pass rather than a guess) and
`syncSolutionToTeams` (a multi-step operation: resolve the app-catalog list,
filter to the matching item, then sync — also deferred pending a dedicated
pass).

## Verification status

Wiring — URL construction for every accessor/action and response parsing —
is covered by `tests/appcatalog.rs` against a recording transport. Compiled
≠ verified against a live tenant — see `AGENTS.md` for the
live-verification record before relying on this against production data,
and note that every write op here (deploy through remove) mutates tenant-
wide state once run live.
