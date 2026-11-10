# forge-m365-sp-navigation

SharePoint navigation operations: quicklaunch, top navigation bar, and the
modern menu-state service. Hand-ported from PnPjs, per `SPEC.md` §4
(SharePoint REST has no reliable CSDL coverage, so this surface is ported
rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\navigation` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Navigation.quicklaunch`/`topNavigationBar`, `NavigationNodes.add`, `NavigationService.getMenuState` |

`get_menu_state` hits `_api/navigation/MenuState` — a different base path
than every other operation here (`_api/web/navigation/...`), matching
PnPjs's `_NavigationService` constructor, which deliberately rebases to the
site collection's `_api/navigation` rather than the current web's
`_api/web/navigation`.

Not yet ported: `delete`, `moveAfter`, node `update`, `children` (nested
node collections), and `getMenuNodeKey`.

## Verification status

Wiring — URL construction (including the two different base paths) and
response parsing (including the `{"results": [...]}` wrapping applied
recursively to `Nodes` at every depth of the menu-state tree) — is covered
by `tests/navigation.rs` against a recording transport. Compiled ≠ verified
against a live tenant — see `AGENTS.md` for the live-verification record
before relying on this against production data, and note that
`add_navigation_node` mutates real data once run live.
