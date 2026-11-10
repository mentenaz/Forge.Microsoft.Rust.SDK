# forge-m365-sp-items

SharePoint list-item instance operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this
surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\items` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_Items` collection, `_Item` instance, `_ItemVersions`, `_ItemVersion` |

## Module structure

| Module | Contents |
|---|---|
| `item` | Instance operations: `get_by_id`, `get_by_string_id`, `recycle`, `delete_with_params`, `get_wopi_frame_url`, `validate_update_list_item`, `get_parent_infos` |
| `item_properties` | Sub-resource reads: `get_effective_base_permissions`, `get_effective_base_permissions_for_ui`, `get_field_values_as_html`, `get_field_values_as_text`, `get_field_values_for_edit` |
| `versions` | Version history: `get_versions`, `get_version_by_id`, `delete_version` |
| `types` | `ItemDeleteParams`, `ItemUpdateResult`, `ItemImageUploadResult`, `ListItemFormUpdateValue`, `ItemParentInfos`, `VersionInfo`, `ValueInfo`, `ResourcePath` |

## Relationship to `forge-m365-sp-lists`

`sp-lists` covers **list-level** CRUD (`get_list_by_title`, `get_items`,
`add_item`, `update_item`, `delete_item`) and list metadata.
`sp-items` covers **item-level** instance operations — fetching a single
item by ID, recycling, WOPI URLs, field value views, version history,
validate-update, and parent-info resolution. Both crates use
`Map<String, Value>` for dynamic item schemas (list columns are only
known at runtime).

## Verification status

Wiring — URL construction, response parsing, request body shapes,
and header assertions (`IF-MATCH`) — is covered by `tests/items.rs`
against a recording transport. Compiled ≠ verified against a live tenant —
see `AGENTS.md` for the live-verification record before relying on this
against production data.
