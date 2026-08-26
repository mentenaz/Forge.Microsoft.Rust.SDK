# forge-m365-sp-folders

SharePoint folder operations. Hand-ported from PnPjs, per `SPEC.md` §4
(SharePoint REST has no reliable CSDL coverage, so this surface is ported
rather than generated). Complements `sp-files`, which already depends on
`getFolderByServerRelativePath` for uploads.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\folders` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `folderFromServerRelativePath`, `Folder.folders`, `Folders.addUsingPath`, `Folder.delete` |

Not yet ported: move/copy, `addSubFolderUsingPath`, storage metrics,
recycle, `getItem` (associated list item), and `deleteWithParams`.

## Verification status

Wiring — request construction (method, URL, `overwrite` flag, `IF-MATCH`
on delete) and response parsing — is covered by `tests/folders.rs` against
a recording transport. Compiled ≠ verified against a live tenant — see
`AGENTS.md` for the live-verification record before relying on this against
production data, and note that `add_folder`/`delete_folder` mutate real
data once run live.
