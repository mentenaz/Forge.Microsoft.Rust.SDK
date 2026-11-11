# AGENTS.md — Forge.Microsoft.Rust.SDK

Context for AI agents continuing work in this repository. The full
specification lives in [`SPEC.md`](SPEC.md) — read it first. This file is the
working-state summary and the rules for contributing.

## Quick reference: build / lint / test

```bash
# All-in-one (what CI runs — must all pass before committing)
cargo fmt --all -- --check
cargo clippy --workspace --all-targets -- -D warnings
cargo build --workspace --all-targets
cargo test --workspace

# Single crate
cargo test -p forge-m365-sp-lists
cargo clippy -p forge-m365-sp-lists --all-targets -- -D warnings

# Format (auto-fix)
cargo fmt --all
```

No `rustfmt.toml` or `clippy.toml` — default Rust formatter/linter settings.

## What this is

A full-scope Rust implementation of the Microsoft 365 ecosystem (Graph +
SharePoint REST + admin surfaces), using PnP PowerShell / PnP Framework /
PnP.js as the functional spec. Private until stable; crates publish to
crates.io as `forge-m365-*` when ready. Consumed as a library by Tauri, GPUI
and terminal apps.

## Current state: M0 done, M1 code-complete (live verification pending), M2 done

Workspace has 40 crates: `forge-m365` (facade), `-core` (Error, Surface,
Ladder, `OperationEntry`, `Transport` trait w/ `headers` param,
`Client::run_ladder`), `-auth`, `-macros`, `-sp-sites`, `-sp-webs`,
`-sp-lists`, `-sp-files`, `-sp-search`, `-sp-site-users`, `-sp-folders`,
`-sp-content-types`, `-sp-site-groups`, `-sp-views`, `-sp-security`,
`-sp-sharing`, `-sp-subscriptions`, `-sp-features`,
`-sp-regional-settings`, `-sp-attachments`, `-sp-hubsites`,
`-sp-recycle-bin`, `-sp-comments`, `-sp-navigation`, `-sp-forms`,
`-sp-appcatalog`, `-sp-fields`, `-sp-items`, `-sp-profiles`,
`-sp-related-items`, `-sp-publishing-sitepageservice`, `-sp-webparts`,
`-sp-column-defaults`, `-sp-groupsitemanager`, `-sp-site-scripts`,
`-sp-utilities`, `-sp-site-designs`, `-sp-user-custom-actions`,
`-sp-social`, `-sp-clientside-pages`.
Escalation-ladder tests pass (`cargo test --workspace`, 322 tests, fmt/
clippy `-D warnings`/build/test all clean workspace-wide).

M2 domain progress: `sp-sites`, `sp-webs`, `sp-lists`, `sp-files`,
`sp-search`, `sp-site-users`, `sp-folders`, `sp-content-types`,
`sp-site-groups`, `sp-views`, `sp-security`, `sp-attachments`,
`sp-hubsites`, `sp-recycle-bin`, `sp-comments`, `sp-navigation`,
`sp-forms`, `sp-fields`, `sp-items`, `sp-profiles`, `sp-sharing` ported
(read-only + list/item/file/folder/attachment/comment
CRUD + role assignments/definitions + recycle-bin restore/delete + nav
node add/menu-state + field CRUD/creation types + item versions/permissions/recycle/WOPI/validate + user profiles (PeopleManager/ProfileLoader/ClientPeoplePicker) + list webhook subscriptions (add/renew/delete)); write ops beyond that — site/list/web create-delete,
chunked upload, search pagination, user/group add-remove, content-type
create-delete, view mutations, role assignment add-remove, attachment
setContent/recycle, recycle-bin bulk ops, comment replies, profile
set/personal-site provisioning notes — are deferred
and noted in each crate's README; `sp-forms` has no write API in PnPjs to
defer). `sp-site-groups` reuses `SiteUserInfo` from `sp-site-users` (first
cross-crate type reuse). `sp-comments` is the first crate covering
SharePoint's newer camelCase "likes and comments" service rather than the
classic PascalCase `_api/web` surface. `sp-appcatalog` is also ported.
`sp-sharing` covers the file/folder/item/web sharing surface (`shareWith`,
static/instance `shareObject`, getSharingInformation, object sharing
settings, unshare, sharing-link create/delete); its share/unshare writes are
live-test only, and `includeAnonymousLinkInEmail`/`propagateAcl`/
`useSimplifiedRoles` on the web static `shareObject` are accepted but not
transmitted, matching PnPjs. `sp-subscriptions` covers list webhook
subscriptions (list/renew/delete per `packages/sp/subscriptions`); its
add/renew/delete writes are live-test only. `sp-features` covers site/web
feature activation (`features` nav + `features/add`/`features/remove`
actions per `packages/sp/features`); its activate/deactivate writes are
live-test only. `sp-regional-settings` covers web regional settings, time
zones (`timezone`/`timezones`/`GetById`/`utctolocaltime`/`localtimetoutc`),
installed languages, and the cross-cutting `IUserResources`
`{title|description|name}Resource/getValueForUICulture` binder on web/list/
field/content-type/user-custom-action (per `packages/sp/regional-settings`);
all its POSTs are read-style operations (none mutate state).

Latest M2 batch (this session) ported ten more PnPjs `sp/*` domains, all
fmt/clippy/build/test-clean against mock transports: `sp-related-items`
(`SP.RelatedItemManager` — get/add/delete related-item links, complete
surface); `sp-publishing-sitepageservice` (`SP.Publishing.SitePageService`
— its one method, `getCurrentUserMemberships`, complete surface);
`sp-webparts` (classic `LimitedWebPartManager` — list/get/export/import web
parts, move/open/close/delete a web part instance; the *modern*
client-side-page canvas model is a separate PnPjs package, deliberately not
covered here); `sp-column-defaults` (per-folder default column values —
reads/patches the `Forms/client_LocationBasedDefaults.html` XML file rather
than a plain JSON REST verb); `sp-groupsitemanager` (`_api/groupsitemanager`
— M365-group site creation/status/teamify, tenant-root scoped, not
site-scoped); `sp-site-scripts` and `sp-site-designs` (both hang off the
same static `_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.
SiteScriptUtility.<MethodName>` POST-only action-verb endpoint family — kept
as two crates per the fine-grained-workspace rule even though they share a
URL convention); `sp-utilities` (`SP.Utilities.Utility` static methods —
sendEmail, resolvePrincipal, searchPrincipals, expandGroupsToPrincipals,
createEmailBodyForInvitation); `sp-user-custom-actions` (web- and
site-scoped custom action CRUD via one `Scope` parameter); `sp-social`
(`_api/social.following` — legacy, superseded by Viva Engage/Graph in
modern tenants, but ported completely since it's still documented and
live). Bugs the tests caught and fixed during integration: a `null`-vs-
`String` field (`GroupSiteInfo::error_message`) in `sp-groupsitemanager`,
a missing `#[serde(rename_all = "PascalCase")]` on `sp-utilities`'s
`PrincipalInfo`, and wrong slash-vs-dot method-separator assumptions in
`sp-site-designs`'s tests (implementation was already correct).

`clientside-pages` (~1800 lines in PnPjs — the modern client-side page
canvas/section/column/control serialization model) is being ported as its
own multi-phase pass rather than folded into a batch, given its size and
that it's a recursive JSON data model rather than a flat set of REST verbs.
**Phase 1 is done**: `forge-m365-sp-clientside-pages` (40th crate) — the
canvas data model (`Position`, `Emphasis`, `ColumnData`, `PageSettingsSlice`,
`TextData`, `WebPartData`, `Column`, `Section`, `Page`) and the exact JSON
codec/tree-merge algorithm for `CanvasContent1`/`LayoutWebpartsContent`
(`mergePartToTree`/`mergeColumnToTree`/`getOrCreateSection`/`reindex`,
including the emphasis-stamping and `<`→`&lt;` escaping `getControls`/
`specialSaveHandling` perform). Deliberately has **no** dependency on
`forge-m365-core`/`forge-m365-macros` — it does no network I/O, so its 9
tests assert exact JSON shape offline rather than only wiring against a
mock transport, unlike every other crate in this workspace. Two bugs the
tests caught: `inner_html` needs `#[serde(rename = "innerHTML")]` (serde's
camelCase heuristic doesn't know `HTML` is an acronym, so it produced
`innerHtml`), and parsing our own `to_canvas_content_json()` output back
via `from_canvas_content_json()` failed on the trailing `null` page-settings
placeholder (PnPjs always appends `_pageSettings ?? null`; the Rust
`Control` deserializer now filters `Value::Null` entries before decoding).
**Phase 2 is also done**: read-only network operations on top of the
Phase 1 model — `get_client_side_page` (GET `_api/sitepages/pages({id})`),
`get_page_item_ref` (resolves a page's list-item id + `CommentsDisabled`
from its server-relative path via `getFileByServerRelativePath(...)/
ListItemAllFields`), `load_client_side_page` (composes the two, mirroring
PnPjs's `Web.loadClientsidePage`/`ClientsidePageFromFile`), and
`get_client_side_web_parts` (GET `_api/web/GetClientSideWebParts`, the
manifest catalog `ClientsideWebpart.fromComponentDef` resolves a component
id against). This is the crate's first use of `forge-m365-core`/
`forge-m365-macros` — added as dependencies in this phase, not Phase 1. 5
more tests (14 total for the crate), still wiring-only for this phase since
it does real network calls. `PageMetadata`/`RawPageData` model PnPjs's
`IPageData` (the two content-string fields are parsed out into
`ClientsidePage::content`/`layout` instead of kept raw).
**Phases 3 and 4 are now also done — `clientside-pages` M2 work is
complete.** Phase 3 (lifecycle ops that don't touch the canvas tree):
`discard_page_checkout`, `promote_to_news` (simplified — PnPjs's
never-published-yet auto-publish-first workaround isn't reproduced),
`schedule_publish`, `share_page_by_email`, `set_author_by_id`/
`set_author_by_login_name` (resolve via `_api/web/siteusers`, then mutate
the in-memory page — persisted by a later `save`), `set_banner_image`
(pure in-memory setter) / `set_banner_image_from_external_url` (uploads via
`_api/sitepages/AddImageFromExternalUrl`), `recycle_page`/`delete_page`
(against `_api/web/lists/getbytitle('Site Pages')/items({id})` — a
documented simplification of PnPjs's `EnsureClientRenderedSitePagesLibrary`
list-resolution indirection, since "Site Pages" is the standard library
title). Phase 4 (the canvas write path): `save_client_side_page` (POSTs
`savepage` then optionally `publish`, then reloads from the server —
mirrors PnPjs's `save()` exactly, banner-image-dirty re-linking sub-flow
excepted), `create_client_side_page` (faithfully reproduces PnPjs's odd
title-set-after-save quirk from `CreateClientsidePage`), `save_page_as_template`,
`copy_client_side_page`/`copy_page_content_to` (banner cross-site GUID
re-linking sub-flow excepted — both are documented in the crate README, not
silent gaps). A correctness fix from this pass: `ClientsidePage::layout` is
no longer `Option<LayoutPartsContent>` — PnPjs's `_layoutPart` is *never*
actually absent (the constructor always assigns
`LayoutPartsContent::default_title_area()` before parsing whatever the
server sent), so a `null` `LayoutWebpartsContent` now fills in that default
rather than surfacing as `None`; this matters for `save()` correctness
(`getLayoutWebpartsContent` must never emit bare `null` once a page exists
in memory). 16 more tests (30 total for the crate: 9 Phase 1 + 5 Phase 2 +
11 Phase 3 + 5 Phase 4), still wiring/offline-only — no live-tenant
verification, same standing caveat as the rest of this workspace. See the
crate's README for the exact scope/simplifications of each phase.

Auth status: client-credentials flow **live-verified** against a real tenant
for both audiences (Graph + SharePoint), via client secret. Supports secret
AND certificate (RS256 JWT assertion, hand-rolled), interactive browser flow
(auth-code + PKCE on loopback, `BrowserConfig::acquire_interactive()` — needs
app-registration "Mobile and desktop" platform + public-client flows enabled;
untested because tenant lacks portal permissions), and device code flow
(`DeviceCodeConfig::acquire()`, RFC 8628 — untested live). All auth flows are
deliberately synchronous fns (not `async fn`): they block on I/O (TCP accept,
polling sleep) rather than depend on a specific async runtime, per the
library's no-owned-runtime constraint; async hosts run them via
`spawn_blocking`. Example: `cargo run --example live_auth -p forge-m365`
picks method from env vars.
NOTE: Kaspersky AV on this machine false-positives freshly built binaries
(VHO:Trojan-Ransom.Convagent.gen); `target\*` must stay in its exclusions.

CI (`ci.yml`): runs on ubuntu-latest + windows-latest in parallel. Steps:
fmt → clippy (with `-D warnings`) → build → test. All four must pass.

M1 vertical slice (auth + sp-sites + sp-lists + sp-files) is code-complete:
built, fmt/clippy/test-clean, wiring-tested against mock transports
(`tests/*.rs` in each crate). **Not yet live-verified** — `live_sites.rs`,
`live_webs.rs`, `live_lists.rs`, `live_files.rs`, `live_search.rs`,
`live_site_users.rs`, `live_folders.rs`, `live_content_types.rs`,
`live_site_groups.rs`, `live_views.rs`, `live_security.rs`,
`live_attachments.rs`, `live_hubsites.rs`, `live_recycle_bin.rs`,
`live_comments.rs`, `live_navigation.rs`, `live_forms.rs`,
`live_fields.rs`, `live_items.rs`, `live_profiles.rs`, `live_sharing.rs`,
`live_subscriptions.rs`, `live_features.rs`, `live_regional_settings.rs`,
`live_device_code.rs` exist
under `forge-m365/examples/` but none have been
run against a real tenant (owner has confirmed they lack a full admin
M365/SPO environment to test against — this may remain permanently
unverified; keep building but never claim live success that didn't happen).
Do not claim any of this surface works against SharePoint until one of
those examples has actually succeeded.

sp-lists/sp-files needed a core change: `Transport::execute` /
`Client::run_ladder` gained a `headers: &[(&str, &str)]` param.
`AuthedTransport` defaults `Content-Type: application/json;odata=verbose`
whenever a body is present, unless the caller already supplied a
Content-Type (file upload overrides it to `application/octet-stream`).

Next: `clientside-pages` is functionally complete (all 4 phases) —
**M2 as a whole is now done**. What remains open: no `live_*.rs` examples
exist yet for the ten M2-batch crates from a previous session
(`related_items`, `publishing_sitepageservice`, `webparts`,
`column_defaults`, `groupsitemanager`, `site_scripts`, `utilities`,
`site_designs`, `user_custom_actions`, `social`) or for `clientside-pages`
— only mock-transport/offline wiring tests exist for these. Adding those
examples, then starting M3 (Graph domains via generator) are the open
threads. Owner has directed continuing without per-crate check-ins since
live testing isn't currently possible either way.

## Locked decisions (do not relitigate without cause)

- Repo/crate naming: `forge-m365-*`, facade `forge-m365`. No dots in crate
  names; avoid `microsoft-`/`pnp-` prefixes on crates.io.
- Fine-grained workspace: one small crate per endpoint domain.
- API generation is **hybrid**: Graph generated from CSDL/OpenAPI;
  SharePoint REST hand-ported from PnPjs (`E:\sources\pnpjs\packages\sp`).
- Every operation declares an escalation ladder at compile time via a
  proc-macro (`#[pnp_operation(...)]`), patterned on Forge.Rust.SDK's
  `#[forge_action]` + inventory (`E:\Forge.Rust.SDK`).
- Auth emits UI events, never blocks on stdin; dual token audiences managed
  transparently.
- Library-only: async fns, no owned Tokio runtime.
- CSOM is out of scope as a protocol; its operations map to REST or become
  typed `Unsupported`.

## Graph / M3 plan (decided 2026-11-11 — see SPEC.md §8 M3 for the phase list)

- **Core hardening comes first** (G0), before any generated code.
- **Generator** is a standalone tool (`tools/graph-gen`, outside the workspace members) that reads a
  pinned `microsoftgraph/msgraph-metadata` OpenAPI file (v1.0 only at first) and writes
  **checked-in** Rust source. No `build.rs`, no build-time network.
- **Granularity**: ~15 workload crates (`forge-m365-graph-<workload>`) + one shared
  `forge-m365-graph-types` crate. Not one crate per resource, not 1:1 with PnPjs `graph/*`.
- **First slice**: the SharePoint-overlapping surface (sites, drives/drive items, lists, list items),
  so the escalation ladder gets real Graph fallbacks for `sp-sites`/`sp-lists`/`sp-files`/`sp-items`.
- Every generated crate README records the metadata repo + commit it was generated from.
- Third-party crates were reviewed and rejected — see the section below; never depend on `msgraph-rs` (GPL).

## Reference sources on disk (verified paths)

| Path | What |
|---|---|
| `E:\powershell` | PnP PowerShell C# repo (800+ cmdlets) — functional spec |
| `E:\pnpframework` | PnP Framework @ `e7631b62` |
| `E:\pnpcore` | PnP Core SDK @ `d90d3216f` |
| `E:\sources\pnpjs` | PnP.js — pure REST endpoint map (best hand-port reference) |
| `E:\sources\PowerShell` | PowerShell v7.4.0 source |
| `E:\sources\MSAL-dotnet` | MSAL.NET 4.85.2 |
| `E:\sources\ApplicationInsights-dotnet` | App Insights 2.21.0 |
| `E:\sources\TextCopy` | TextCopy 6.2.1 |
| `E:\sources\dotnet-runtime` | .NET BCL v9.0.5 |
| `E:\sources\msgraph-sdk-dotnet` | Graph SDK snapshot |
| `E:\sources\CSOM-binaries` | closed-source SPO CSOM nupkg (reference only) |
| `E:\Forge.Rust.SDK` | sibling project — macro/inventory pattern to imitate |

## Third-party Rust crates reviewed (2026-11-11)

Surveyed on crates.io; none is a dependency candidate. Do not re-survey unless a new crate appears.

| Crate | Verdict |
|---|---|
| `msgraph-rs` 0.1.2 | ~244 lines (token + get/list user). Published once in 2024, abandoned. **GPL-3.0-only** — incompatible with our MIT; never depend on it or copy from it. |
| `sharepoint` 0.3.0 | 2017, hyper 0.11, username/password cookie auth that modern tenants block. LGPL-3.0. Nothing useful. |
| `sharepoint-cli` 0.1.1 (MIT, active) | CLI, not an SDK (~8.6k lines; its REST layer is read-only GETs for sites/drives/files/search). Read-only **reference**, not a dependency. Source: https://github.com/rvben/sharepoint-cli |

What `sharepoint-cli` is useful for (re-implement in our own style; if code is ever lifted verbatim,
record its MIT license + source in the crate README):

- **Core hardening reference** — our core has none of these yet: 429/5xx retry with `Retry-After`
  (seconds and HTTP-date), `@odata.nextLink` / `odata.nextLink` paging, status → typed error mapping
  (incl. verbose 500 bodies), token cache + session store, device-code/browser flow edge cases
  (its device-code is 547 lines vs. our 213 — compare for `slow_down`/interval handling).
- **SharePoint REST quirks to turn into tests** in our crates: file `Length` may arrive as a string;
  search `LastModifiedTime` zero-fraction timestamps; document-library listing filter
  `BaseTemplate eq 101 or 700 and Hidden eq false` with `$expand=RootFolder`; download via
  `GetFileByServerRelativePath(decodedurl='…')/$value` (check `sp-files` uses the `decodedurl` form).
- **Design contrast** — its `Backend` enum switches Graph vs. REST by config with no automatic
  fallback; our escalation ladder is the more ambitious design. Its behaviour is evidence for the
  quirks above only — it does **not** live-verify our crates.
- **Open question noted**: nothing in our crates fetches a request digest (`contextinfo`). Bearer-token
  calls shouldn't need one, but write ops are unverified — keep it flagged.

## Domain knowledge that must survive sessions

- **Communication sites** (no M365 group) are not fully queryable via Graph;
  they require SharePoint REST (`_api/web`, `_api/search/query`). Both API
  clients are permanent first-class citizens of this SDK.
- `_api` is OData **v3-flavored**, Graph is **v4** — core parses both.
- SharePoint and Graph use different token audiences.
- Some tenant-admin endpoints are undocumented; model from PnPjs `sp-admin`
  and observed behavior, and mark them as such in code comments.

## Working rules for agents

1. Read `SPEC.md` before proposing changes; keep it current when decisions
   change. This repo's docs are the source of truth once code exists.
2. When porting a module, record the reference source + commit (e.g.
   "pnpjs@<sha> packages/sp/files") in that crate's README. Parity must be
   auditable.
3. Never describe generated or ported API code as tested against a live
   tenant unless it actually was. Compiled ≠ verified.
4. Publishing (crates.io, repo visibility change) is the owner's action, not
   the agent's.
5. Do not create issues/PRs unprompted.
