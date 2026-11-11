# Roadmap — Forge.Microsoft.Rust.SDK

Forward-looking plan, written 2026-11-11. It complements, and does not replace:

- [`SPEC.md`](SPEC.md) — specification and milestone table (§8).
- [`AGENTS.md`](AGENTS.md) — current working state and locked decisions (source of truth for what exists).

Effort figures are rough guesses, in working sessions (one sitting of a few hours, with an AI agent
writing and the owner reviewing). They exclude live verification, which depends on tenant access.
Never describe any item below as working against a live tenant until a live example has actually
succeeded — compiled and mock-tested is not verified.

## Where we are

- M0 done. M1 code-complete, live verification pending. M2 (SharePoint REST ports) done: 40 crates.
- Only client-credentials auth is live-verified. Everything else is compiled + mock-transport tested.
- Gaps found when checking docs against code:
  - The escalation ladder is never exercised across surfaces: all 273 registered operations have
    `fallback = []`, and `Surface::Graph` is unused by any domain crate.
  - `run_ladder` returns the last `Error::Surface` when all surfaces fail; `Error::Unsupported` only
    appears for an empty ladder.
  - `forge-m365-core` (~90 lines) has no retry/backoff, paging helper, batching, caching, or typed
    HTTP errors (every failure is `Error::Surface(String)`).
  - `forge-m365-auth` has no managed identity, no WAM broker, no token cache.
  - No Graph crates, no `sp-admin`, no facade `graph` module.
  - 11 crates have no `live_*.rs` example (`related-items`, `publishing-sitepageservice`, `webparts`,
    `column-defaults`, `groupsitemanager`, `site-scripts`, `site-designs`, `utilities`,
    `user-custom-actions`, `social`, `clientside-pages`).
  - Many write operations are deferred (see each crate's README).

## M3 — Graph (decided 2026-11-11)

Decisions: core hardening first; standalone generator (`tools/graph-gen`, outside workspace members)
writing checked-in source; ~15 workload crates (`forge-m365-graph-<workload>`) plus a shared
`forge-m365-graph-types`; first slice is the SharePoint-overlapping surface; v1.0 only at first.

### Spike (1 session)
Parse the pinned `msgraph-metadata` v1.0 OpenAPI file and generate the `/sites` slice into one crate.
Goal: learn how an off-the-shelf OpenAPI parsing crate copes with the file, and how hard the type
rules are, before committing to G1/G3 estimates.

Size to plan for (measured): v1.0 `openapi.yaml` is 44.8 MB, ~997k lines, 1,435 paths, 18,004
operations, 5,341 schemas. Generation must be allow-list driven; the spec itself is not committed.

### G0 — Core hardening (2–3 sessions)
- Typed errors: status code + parsed Graph / SharePoint error body, so callers can tell 404 from 403.
- Retry/backoff on 429 and 5xx, honoring `Retry-After` (seconds and HTTP-date).
- Paging helper: `@odata.nextLink` (Graph) and `odata.nextLink` (SharePoint).
- Decide which statuses are "surface-specific" (fall through the ladder) vs. abort.
- `run_ladder` returns `Unsupported` on exhaustion, as the docs promise.
- Turn the SharePoint REST quirks in AGENTS.md into tests (string `Length`, zero-fraction search
  timestamps, `decodedurl` download paths, verbose 500 bodies).
- Done when: the new tests pass under the full gate (fmt, clippy `-D warnings`, build, test).

### G1 — Generator and types (4–6 sessions)
- `tools/graph-gen`: reads a pinned OpenAPI file plus per-workload allow-list TOML
  (`include_paths`, `include_methods`, optional renames, optional ladder fallback wiring).
- Pipeline: parse and resolve `$ref`s → filter to allow-list → collect reachable schemas →
  map to Rust types → emit with `quote` → format → stamp crate README with the metadata commit.
- Type rules: optional fields, unknown-field tolerance, enums with unknown-variant fallback,
  `@odata.type` polymorphism (tagged enum or raw-JSON fallback), open types and extension properties.
- Generated files carry a "do not edit" header and live in a `generated/` module; hand-written
  wrappers sit in separate files.
- Done when: the types crate compiles; snapshot tests cover the type rules; golden tests round-trip
  sample Graph payloads.

### G2 — First slice and a real ladder (3–4 sessions)
- Generate sites, drives/drive items, lists, and list-items crates.
- Add Graph fallbacks to `sp-sites`, `sp-lists`, `sp-files`, `sp-items` where an equivalent exists.
- Ladder tests that fail the primary surface and assert the fallback request shape.
- `live_graph_*.rs` examples (unverifiable without a tenant; say so in each).
- Done when: the ladder has non-empty fallbacks, covered by tests.

### G3 — Remaining workloads (10–20 sessions)
Users, groups, teams, mail, calendar, planner, and others, prioritized by need, using PnPjs
`packages/graph/*` as the functional spec. Mostly a TOML file per workload plus quirk fixes; the most
elastic phase.

### G4 — Batching, facade, token cache (2–3 sessions)
- `$batch` in core; facade `graph` module in `forge-m365`; file-backed token cache.

**Totals**: through G2 ≈ 10–14 sessions; full M3 ≈ 22–37 sessions.

**Risks**: polymorphism/type quirks (could double G1); parser performance on 45 MB (could add 1–2
sessions); G3 scope; reviewing large generated diffs.

## M4 — Auth and admin surfaces

- Managed identity; Windows WAM broker (platform-gated feature).
- `sp-admin` (tenant, office-tenant, site-properties) hand-ported from `pnpjs/packages/sp-admin`;
  mark undocumented endpoints as such.
- Compare our device-code flow against `sharepoint-cli` (see AGENTS.md notes) for `slow_down` /
  interval edge cases.

## Cross-cutting backlog (any time)

- Deferred SharePoint write operations: site/list/web create-delete, chunked upload, search
  pagination, user/group add-remove, content-type create-delete, view mutations, role assignment
  add-remove, attachment `setContent`/recycle, recycle-bin bulk ops, comment replies, profile
  set / personal-site provisioning, and the `clientside-pages` simplifications noted in its README.
- The 11 missing `live_*.rs` examples listed above.
- Confirm request-digest behaviour for write operations (bearer-token calls should not need one;
  currently unverified and nothing fetches `contextinfo`).
- Doc housekeeping: `SPEC.md` header still says "planning / pre-implementation" and its workspace
  layout predates the move to `crates/`; `AGENTS.md` does not mention the move.

## M5 — Release readiness

- Parity audit against the PnP PowerShell cmdlet list.
- Documentation pass.
- crates.io publish — the owner's action only; do not publish or change repo visibility from an agent.

## Live verification (blocked on tenant access)

The owner has no full admin M365/SPO environment. Until one is available, interactive-browser,
device-code, and every domain crate remain unverified. Keep building; keep the caveat in every README.
