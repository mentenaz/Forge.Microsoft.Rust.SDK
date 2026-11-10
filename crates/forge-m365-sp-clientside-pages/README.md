# forge-m365-sp-clientside-pages

SharePoint modern/client-side page **canvas** data model and JSON codec.
Hand-ported from PnPjs, per `SPEC.md` §4 (SharePoint REST has no reliable
CSDL coverage, so this surface is ported rather than generated).

This was ported in 4 phases — see `AGENTS.md` for the overall M2 status.
**All 4 phases are done.** Scope of this crate today:

**Phase 1 (offline model + codec):**
- The canvas data model (`Position`, `Emphasis`, `ColumnData`,
  `PageSettingsSlice`, `TextData`, `WebPartData`, `Column`, `Section`,
  `Page`) matching PnPjs's `IPosition`/`IClientControlEmphasis`/
  `IClientsidePageColumnData`/`IClientsidePageSettingsSlice`/
  `IClientsideTextData`/`IClientsideWebPartData`.
- The exact JSON (de)serialization for the two list-item fields SharePoint
  stores a page's canvas in: `CanvasContent1` (a flat array of controls) and
  `LayoutWebpartsContent` (the header/banner web part, `null` or a
  single-element array).
- The tree-merge algorithm that turns that flat array into a
  sections → columns → controls tree and back
  (`mergePartToTree`/`mergeColumnToTree`/`getOrCreateSection`/`reindex`),
  including the emphasis-stamping and `<` → `&lt;` escaping
  `getControls`/`specialSaveHandling` perform on save.
- Enough tree-building API (`Page::add_section`, `Section::add_column`,
  `Section::add_control`, `WebPartData::new`, `TextData::new`) to construct
  a new page's canvas from scratch, not just parse an existing one.

**Phase 2 (read-only network operations, on top of Phase 1's model):**
- `get_client_side_page` — GET `_api/sitepages/pages({id})`, parsed into a
  [`ClientsidePage`] (metadata + the already-parsed `content`/`layout`).
- `get_page_item_ref` — resolves a page file's list-item id and
  `CommentsDisabled` flag from its server-relative path.
- `load_client_side_page` — composes the two above, mirroring PnPjs's
  `Web.loadClientsidePage`/`ClientsidePageFromFile`.
- `get_client_side_web_parts` — GET `_api/web/GetClientSideWebParts`, the
  manifest catalog `ClientsideWebpart.fromComponentDef` resolves a
  component id against (reading those manifests to build a `WebPartData`
  is left to the caller — see "Deliberate simplifications" below).

**Phase 3 (lifecycle operations that don't touch the canvas tree):**
- `discard_page_checkout`, `promote_to_news`, `schedule_publish`,
  `share_page_by_email`.
- `set_author_by_id`/`set_author_by_login_name` — resolve a site user via
  `_api/web/siteusers`, then mutate the in-memory `ClientsidePage` (a
  subsequent `save_client_side_page` persists it, matching PnPjs's own
  setter semantics).
- `ClientsidePage::set_banner_image` (pure in-memory setter) and
  `set_banner_image_from_external_url` (uploads via
  `_api/sitepages/AddImageFromExternalUrl`, then calls the setter).
- `ClientsidePage::set_topic_header` (pure in-memory setter mirroring the
  `topicHeader` property's PnPjs setter, used by `copy_page_content_to`).
- `recycle_page`/`delete_page`.

**Phase 4 (the canvas write path):**
- `save_client_side_page` — POSTs `savepage` (optionally `publish`
  afterward), then reloads the page from the server, mirroring PnPjs's
  `save()`.
- `create_client_side_page` — POSTs a new page shell, then saves it.
- `save_page_as_template`.
- `copy_client_side_page`/`copy_page_content_to` — copies canvas content
  and select layout properties from one page onto another, then saves.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\clientside-pages\types.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_ClientsidePage` canvas/tree logic (`getCanvasContent1`/`getLayoutWebpartsContent`/`setControls`/`getControls`/`mergePartToTree`/`mergeColumnToTree`/`getOrCreateSection`/`specialSaveHandling`), `CanvasSection`, `CanvasColumn`, `ColumnControl`/`ClientsideText`/`ClientsideWebpart`, and the `I*Data`/`IPosition`/`IClientControlEmphasis`/`ILayoutPartsContent` interfaces |
| PnPjs | `E:\sources\pnpjs\packages\sp\clientside-pages\funcs.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `getNextOrder`, `reindex` |

## Why `forge-m365-core`/`forge-m365-macros` arrived in Phase 2, not Phase 1

Phase 1's data model and codec do no network I/O and register no
`#[pnp_operation]` — pure data model and codec, unlike every other `sp-*`
crate in this workspace at the time it landed. That was deliberate:
`clientside-pages` is unusually large (~1800 lines in PnPjs, a recursive
canvas/section/column/control JSON model rather than a flat set of REST
verbs) and mixes several genuinely separable concerns. Splitting the
model/codec out as its own buildable, independently-testable phase — one
that could assert exact JSON shape offline, without a mock transport — was
judged clearer than attempting the whole surface at once. Phase 2 adds
`forge-m365-core`/`forge-m365-macros` now that there are real network calls
to make; Phase 1's types and codec fns remain fully usable without a
`Client` for anyone who only needs to build/parse canvas JSON offline.

## Deliberate simplifications and gaps

Everything below is a documented choice, not an oversight — check here
before assuming a PnPjs behavior was ported faithfully.

- **`recycle_page`/`delete_page`** hit `_api/web/lists/getbytitle('Site
  Pages')/items({id})` directly. PnPjs's `getItem()` instead resolves the
  library via `_api/lists/EnsureClientRenderedSitePagesLibrary` (defensive
  against a renamed/relocated library). "Site Pages" is the standard,
  virtually-always-correct title, so this is a pragmatic simplification —
  it will misbehave against a tenant whose Site Pages library was renamed.
- **`promote_to_news`** doesn't reproduce PnPjs's `promoteNewsImpl`
  workaround, which auto-publishes a page first if it has never been
  published (a server-bug mitigation). Publish the page yourself first if
  it's new.
- **`save_client_side_page`** doesn't reproduce the `_bannerImageDirty`
  branch of PnPjs's `save()`, which — only when a banner image was just set
  from a raw uploaded file rather than an external URL — issues a small
  batch of GETs to compute a `getpreview.ashx`-based banner URL carrying
  the image's site/web/file GUIDs. Banners set via
  `set_banner_image_from_external_url` are unaffected.
- **`copy_page_content_to`** doesn't reproduce PnPjs's cross-site banner-
  image re-linking (parsing site/web/file GUIDs out of the source's
  `getpreview.ashx` banner URL and resolving the equivalent file in the
  destination site). The destination page keeps whatever banner
  `create_client_side_page` gave it.
- **Web part manifest resolution** (the rest of
  `ClientsideWebpart.fromComponentDef`): `get_client_side_web_parts`
  fetches the raw manifests; parsing the `IClientSidePageComponentManifest`
  JSON string inside a manifest's `Manifest` field and extracting
  `preconfiguredEntries[0]` to build a `WebPartData` is left to the caller.
- **`create_client_side_page`** faithfully reproduces an odd PnPjs quirk:
  the page is saved once with `title` set to the filename, and only
  afterward is the in-memory `title` set to the caller's requested value —
  without a second save. This is PnPjs's actual behavior
  (`CreateClientsidePage`), not a bug introduced here.
- **Not ported at all**: `Web.addRepostPage`/`addFullPageApp`, the
  collection-level "list all pages" surface, and the `RepostPage`/
  `PromotedState` action endpoints beyond `promote_to_news`.

## Verification status

**Phase 1** (`tests/canvas.rs`): parsing a realistic `CanvasContent1`
payload into the expected tree shape, round-tripping it back to JSON
(including the emphasis-stamping and `<`/`&lt;` escaping), reindexing, the
page-settings `null` quirk (PnPjs always appends a page-settings slot to
the array, `null` when none was ever loaded — this is reproduced
faithfully, not treated as a bug), building a page from scratch,
`LayoutWebpartsContent` null/present round-tripping, and the GUID format
`new_guid()` produces. Because this part of the crate does no network I/O,
these tests assert exact JSON shape rather than only request wiring.

**Phase 2** (`tests/page_ops.rs`): URL construction (including percent-
encoding a server-relative path with spaces) and response parsing for all
four functions against a recording transport, plus that
`load_client_side_page` issues its two requests in the right order and
threads `CommentsDisabled` through correctly.

**Phase 3** (`tests/lifecycle.rs`): URL/method construction and response
parsing for every lifecycle op, `set_author_by_id`'s two-step resolution
(id → login name → profile) issuing its requests in order, the "no such
user" error path, the banner-image relativization heuristic, and that
`set_banner_image_from_external_url` threads the upload response into the
in-memory setter correctly.

**Phase 4** (`tests/write_ops.rs`): `save_client_side_page`'s save-then-
optionally-publish-then-reload request sequence and body shape,
`create_client_side_page`'s title-set-after-save quirk (asserting the save
body actually carries the filename-derived title, not the final one),
`save_page_as_template` inheriting the source's title, and
`copy_client_side_page`/`copy_page_content_to` sending the copied canvas
content and topic header in its (second, content-carrying) save request.

Like every other `sp-*` crate's network layer, Phases 2-4 are wiring tests
against a mock transport, not live-tenant verification — see `AGENTS.md`
for the live-verification record before relying on this against production
data. Every write operation here (`save_client_side_page` and everything
built on it, `set_author_*`, `set_banner_image_from_external_url`,
`recycle_page`/`delete_page`) mutates real page state once run live.
