# forge-m365-sp-features

SharePoint site/web feature activation operations for the forge-m365 SDK.

Hand-ported from PnP.js `packages/sp/features`
@pnpjs `8ee2375d`. The reference is the pure-REST endpoint map; parity must be
auditable against the listed source files.

## Surface

SharePoint REST — both scopes the `features` nav is wired onto in PnPjs
(`site.ts`, `web.ts`):

| Operation | HTTP | Endpoint | PnPjs source |
|---|---|---|---|
| `get_site_features` | `GET` | `{site}/_api/site/features` | `_Site.features` fetch |
| `get_site_feature_by_id` | `GET` | `{site}/_api/site/features('{id}')` | `types.ts:38` (`getById`) |
| `activate_site_feature` | `POST` | `{site}/_api/site/features/add` | `types.ts:19` (`add`) |
| `deactivate_site_feature` | `POST` | `{site}/_api/site/features/remove` | `types.ts:48` (`remove`) |
| `get_web_features` | `GET` | `{site}/_api/web/features` | `_Web.features` fetch |
| `get_web_feature_by_id` | `GET` | `{site}/_api/web/features('{id}')` | `types.ts:38` |
| `activate_web_feature` | `POST` | `{site}/_api/web/features/add` | `types.ts:19` |
| `deactivate_web_feature` | `POST` | `{site}/_api/web/features/remove` | `types.ts:48` |

## Notes

- PnPjs `add` posts the body `{ featdefScope: 0, featureId, force }` to the
  `features/add` action and returns `IFeatureAddResult { data, feature }`.
  The `feature` half is a nav handle this library doesn't model, so
  `activate_*` returns just the `data` half (`FeatureAddResult`/`FeatureInfo`,
  both a `DefinitionId`).
- `force` defaults to `false` (PnPjs default parameter).
- `featureId` must be a feature GUID; embedded `'` is doubled inside
  `features('{id}')` (PnPjs concats raw).
- `remove` returns no payload from the service; `deactivate_*` returns `()`.
- All responses cross the OData v3 verbose envelope (including the `add`
  action), unwrapped before deserialization.

## Write-op status

`activate_*`/`deactivate_*` are wired and unit-tested against recording
transports but **not** live-verified against SharePoint (no tenant). See the
workspace `AGENTS.md` for the live-testing rule.

## Reference

- `E:\sources\pnpjs\packages\sp\features\{site,web,types,index}.ts`