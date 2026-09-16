# forge-m365-sp-column-defaults

SharePoint document-library column (field) default values: read/write the
per-folder metadata defaults SharePoint applies to new items in a document
library. Hand-ported from PnPjs, per `SPEC.md` §4 (SharePoint REST has no
reliable CSDL coverage, so this surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\column-defaults\list.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_List.getDefaultColumnValues`, `_List.setDefaultColumnValues` |
| PnPjs | `E:\sources\pnpjs\packages\sp\column-defaults\folder.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_Folder.getDefaultColumnValues`, `_Folder.setDefaultColumnValues`, `_Folder.clearDefaultColumnValues` |
| PnPjs | `E:\sources\pnpjs\packages\sp\column-defaults\types.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `IFieldDefault`/`IFieldDefaultProps`/`AllowedDefaultColumnValues` shapes |

This surface is unusual among the ported crates: it isn't a single REST verb.
SharePoint stores column defaults as a small XML document
(`Forms/client_LocationBasedDefaults.html`) inside the library's root folder,
which the client reads, patches, and re-uploads via the plain file-read/
file-add endpoints `sp-files`/`sp-folders` also use. This crate hand-rolls a
minimal parser/serializer for that fixed `<MetadataDefaults><a
href="..."><DefaultValue FieldName="...">...</DefaultValue></a></MetadataDefaults>`
schema (regex-free, matching the flat non-nested structure PnPjs's own
regex-based parse assumes) rather than pulling in an XML crate dependency.

### Deliberate adaptations (not gaps)

- PnPjs's `_Folder.getDefaultColumnValues`/`setDefaultColumnValues` discover
  their owning list at runtime by reading the folder's `vti_x005f_listname`
  property and walking up through a `Web` object. Every function in this
  crate instead takes `list_title` as an explicit parameter, matching how
  every other crate in this workspace is called — so that discovery step
  doesn't apply here; `get_folder_default_column_values`/
  `set_folder_default_column_values`/`clear_folder_default_column_values`
  otherwise implement the same filter/merge/clear semantics.
- `ColumnDefaultValue::Raw` has no PnPjs equivalent. It exists because
  `get_default_column_values` always returns plain, already-formatted string
  values (as PnPjs's own getter does), but re-running those through the
  `TypeAsString` formatting switch during a merge would require reverse-
  parsing a taxonomy value back into `{wssId, termName, termId}` — which
  PnPjs itself does not attempt either (its JS is untyped enough that this
  path silently produces `undefined` in practice). `Raw` carries an existing
  value through a merge unchanged instead of guessing at that reverse parse.

### Not yet ported

`_List.setDefaultColumnValues` also ensures a `LocationBasedMetadataDefaultsReceiver
ItemAdded` event receiver exists on the list (via `_api/web/lists(...)/eventreceivers`
+ `Lists.add`-shaped body) so SharePoint actually honors the written defaults
for newly added items. No `sp-event-receivers`-shaped crate exists yet to
host that generic capability, so it's deferred — call sites should register
that receiver themselves until then, or wait for that crate.

## Verification status

Wiring — URL construction across the multi-request read (rootfolder +
parent-web + file `$value`) and write (`fields` lookup + file `add`) paths,
XML parse/serialize round-tripping, and per-`TypeAsString` value formatting —
is covered by `tests/column_defaults.rs` against a queued recording
transport. Compiled ≠ verified against a live tenant — see `AGENTS.md` for
the live-verification record before relying on this against production data;
`set_default_column_values` and its folder-level wrappers overwrite the
defaults file for the whole library, so a write here can silently drop
another folder's defaults if the caller doesn't pass a complete `defaults`
list (or use `merge = true` at the folder level).
