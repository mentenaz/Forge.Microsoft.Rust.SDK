# forge-m365-sp-regional-settings

SharePoint web regional settings, time zones and UI-culture resource strings
for the forge-m365 SDK.

Hand-ported from PnP.js `packages/sp/regional-settings`
@pnpjs `8ee2375d`. The reference is the pure-REST endpoint map; parity must be
auditable against the listed source files.

## Surface

SharePoint REST — the `regionalsettings` nav on the web plus the
cross-cutting `IUserResources` binder (`getValueForUICulture`).

### Regional settings / time zones (all on `{site}/_api/web/regionalsettings`)

| Operation | HTTP | Endpoint | PnPjs source |
|---|---|---|---|
| `get_regional_settings` | `GET` | `/regionalsettings` | `_RegionalSettings` fetch |
| `get_time_zone` | `GET` | `/timezone` | `types.ts:17` (`timeZone`) |
| `get_time_zones` | `GET` | `/timezones` | `types.ts:24` (`timeZones`) |
| `get_time_zone_by_id` | `POST` | `/timezones/GetById({id})` | `types.ts:92` |
| `get_installed_languages` | `GET` | `/installedlanguages` | `types.ts:31` |
| `utc_to_local_time` | `POST` | `/timezone/utctolocaltime('{iso}')` | `types.ts:47` |
| `local_time_to_utc` | `POST` | `/timezone/localtimetoutc('{iso}')` | `types.ts:66` |

### UI-culture resource strings (`IUserResources`)

All POST `{queryable}/{title|description|name}Resource/getValueForUICulture`
with `{ "cultureName": "..." }` (`funcs.ts`):

| Operation | Queryable base |
|---|---|
| `web_title_resource` / `web_description_resource` | `{site}/_api/web` |
| `list_title_resource` / `list_description_resource` | `{site}/_api/web/lists/getbytitle('…')` |
| `field_title_resource` / `field_description_resource` | `…/fields/getByInternalNameOrTitle('…')` |
| `content_type_name_resource` / `content_type_description_resource` | `{site}/_api/web/contenttypes/getbyid('…')` |
| `usercustomaction_title_resource` / `usercustomaction_description_resource` | `{site}/_api/web/usercustomactions('…')` |

## Notes

- Content type "title" is bound to `nameResource` (`content-type.ts:10`), all
  other types use `titleResource`.
- String-valued responses are lenient: an OData boxed result
  (`{"d": {"UTCToLocalTime": "…"}}`) and a bare JSON string both resolve,
  mirroring the PnPjs `hOP` fallback in `utcToLocalTime`/`localTimeToUTC`.
- `get_installed_languages` reads the `Items` property of the response.
- `utcToLocalTime`/`localTimeToUTC` accept ISO-8601 strings; PnPjs's
  `Date` values are not modeled (callers serialize with `toISOString`).
- `regionalsettings`, `timezone(s)`, `installedlanguages` and `GetById(...)`
  are lowercase per the PnPjs paths; `GetByInternalNameOrTitle` and
  `getValueForUICulture` keep PowerShell/REST casing.

## Live-verification status

All 17 operations are wired and unit-tested against recording transports;
none are live-verified against SharePoint (no tenant). These POSTs return
values and don't mutate state. See the workspace `AGENTS.md` for the
live-testing rule.

## Reference

- `E:\sources\pnpjs\packages\sp\regional-settings\{funcs,list,field,content-type,user-custom-actions,web,types,index}.ts`