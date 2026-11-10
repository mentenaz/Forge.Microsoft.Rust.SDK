# forge-m365-sp-webparts

Classic (pre-modern-page) SharePoint web-part manager operations: list/get a
page's web part definitions, export/import web part XML, and
move/open/close/delete a web part instance. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this surface
is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\webparts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_LimitedWebPartManager` (`export`, `import`, `webparts`/`scope` accessors), `_WebPartDefinitions` (`getById`, `getByControlId`), `_WebPartDefinition` (`saveChanges`, `moveTo`, `close`, `open`, `delete`), `File.getLimitedWebPartManager` (`packages/sp/webparts/file.ts`) for the base URL shape |

`WebPartDefinition` is `serde_json::Map<String, Value>`, not a typed struct —
PnPjs's own `_WebPartDefinitions`/`_WebPartDefinition` classes are untyped
(`_SPCollection`/`_SPInstance` with no generic parameter), so there's no
reference field list to port faithfully. Same approach `sp-lists::get_items`
and `sp-appcatalog::AppInfo` take for the same reason.

**Deliberately out of scope**: the client-side (modern) page canvas/web-part
model — PnPjs `packages/sp/clientside-pages` (~1800 lines: page sections,
columns, control data serialization) — is a separate, much larger surface
that has not been ported this pass. This crate covers only the classic
`getLimitedWebPartManager` surface on a `.aspx` file, which is what
`Exception is thrown if the file is not an ASPX page` in the PnPjs source
implies: it does not work against modern client-side pages at all.

## Verification status

Wiring — URL construction for every accessor/action and response parsing —
is covered by `tests/webparts.rs` against a recording transport. Compiled ≠
verified against a live tenant — see `AGENTS.md` for the live-verification
record before relying on this against production data, and note that every
write op here (import through delete) mutates the target page's web part
layout once run live.
