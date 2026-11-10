# forge-m365-sp-fields

SharePoint field operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this
surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\fields` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_Fields` collection, `_Field` instance, all `add*` typed creation methods, `IFieldInfo`, enums |

## Module structure

Separation of concerns across three internal modules:

| Module | Contents |
|---|---|
| `types` | `FieldInfo`, `FieldType` enum, creation enums (`CalendarType`, `DateTimeFieldFormat`, `UrlFieldFormat`, etc.), creation prop structs (`AddTextProps`, `AddChoiceProps`, etc.) |
| `field` | Field instance operations: `delete_field`, `update_field`, `set_show_in_*_form`, `get_field_by_id`, `get_field_by_title`, `get_field_by_internal_name_or_title` |
| `creation` | Typed field creation: `add_field`, `create_field_as_xml`, `add_text`, `add_calculated`, `add_date_time`, `add_number`, `add_currency`, `add_multiline_text`, `add_url`, `add_user`, `add_lookup`, `add_choice`, `add_multi_choice`, `add_boolean`, `add_location`, `add_image_field`, `add_dependent_lookup_field` |
| `lib` | `get_all_fields` (collection read), `encode_segment`, re-exports |

`FieldInfo` covers a practical subset of SharePoint's ~32-field `SP.Field`
entity — the same subset-over-completeness approach used for `SiteInfo`/
`WebInfo`/`ListInfo` in the other crates.

## Verification status

Wiring — URL construction for every accessor, response parsing,
request body shapes (`__metadata.type`, choice `{"results": [...]}` wrapping),
and header assertions (`IF-MATCH`) — is covered by `tests/fields.rs`
against a recording transport. Compiled ≠ verified against a live tenant —
see `AGENTS.md` for the live-verification record before relying on this
against production data.
