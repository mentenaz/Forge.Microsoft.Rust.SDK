# forge-m365-sp-subscriptions

SharePoint list webhook (subscription) operations for the forge-m365 SDK.

Hand-ported from PnP.js `packages/sp/subscriptions`
@pnpjs `8ee2375d`. The reference is the pure-REST endpoint map; parity must be
auditable against the listed source files.

## Surface

SharePoint REST — list-scoped, all on
`{site}/_api/web/lists/getbytitle('...')/subscriptions`:

| Operation | HTTP | PnPjs source | Notes |
|---|---|---|---|
| `list_subscriptions` | `GET` collection | `pages/sp/subscriptions/types.ts` (`Subscriptions()`) | Returns all webhook subscriptions for the list |
| `get_subscription_by_id` | `GET` instance | `types.ts:20` (`getById`) | Single webhook subscription |
| `add_subscription` | `POST` | `types.ts:32` (`add`) | Requires `notificationUrl` + `expirationDateTime` (`YYYY-MM-ddTHH:mm:ss+00:00`, max 6 months); optional `clientState`; `resource` is the subscriptions collection URL (`this.toUrl()`) |
| `update_subscription` | `PATCH` | `types.ts:59` (`update`) | Renews the subscription; only supplied fields sent, empty strings omitted (PnPjs truthy check) |
| `delete_subscription` | `DELETE` | `types.ts:82` (`delete`) | Removes the webhook subscription |

## Notes

- PnPjs treats responses as untyped (`Promise<any>`); the crate models the
  payload as a tolerant `Subscription` struct (`#[serde(default)]` over
  `id`, `resource`, `notificationUrl`, `expirationDateTime`, `clientState`,
  `lastModifiedTime`).
- `resource` (POST) mirrors PnPjs `this.toUrl()`: the full collection URL
  including the site root.
- `subscriptionId` is passed through inside `subscriptions('{id}')` with
  embedded `'` doubled (PnPjs concats it raw; the SDK uses the same OData
  escaping it applies to list titles elsewhere).
- All responses cross the OData v3 verbose envelope (`{"d": {...}}` /
  `{"d": {"results": [...]}}`), unwrapped before deserialization.

## Write-op status

All five operations are wired and unit-tested against recording transports;
`add`/`update`/`delete` are **not** live-verified against SharePoint
(no tenant). See the workspace `AGENTS.md` for the live-testing rule.

## Reference

- `E:\sources\pnpjs\packages\sp\subscriptions\{list,types,index}.ts`