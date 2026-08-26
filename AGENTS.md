# AGENTS.md — Forge.Microsoft.Rust.SDK

Context for AI agents continuing work in this repository. The full
specification lives in [`SPEC.md`](SPEC.md) — read it first. This file is the
working-state summary and the rules for contributing.

## What this is

A full-scope Rust implementation of the Microsoft 365 ecosystem (Graph +
SharePoint REST + admin surfaces), using PnP PowerShell / PnP Framework /
PnP.js as the functional spec. Private until stable; crates publish to
crates.io as `forge-m365-*` when ready. Consumed as a library by Tauri, GPUI
and terminal apps.

## Current state: M0 done, M1 code-complete (live verification pending), M2 in progress

Workspace has 16 crates: `forge-m365` (facade), `-core` (Error, Surface,
Ladder, `OperationEntry`, `Transport` trait w/ `headers` param,
`Client::run_ladder`), `-auth`, `-macros`, `-sp-sites`, `-sp-webs`,
`-sp-lists`, `-sp-files`, `-sp-search`, `-sp-site-users`, `-sp-folders`,
`-sp-content-types`, `-sp-site-groups`, `-sp-views`, `-sp-security`.
Escalation-ladder tests pass (`cargo test --workspace`, 60 tests as of
`sp-security`).

M2 domain progress: `sp-sites`, `sp-webs`, `sp-lists`, `sp-files`,
`sp-search`, `sp-site-users`, `sp-folders`, `sp-content-types`,
`sp-site-groups`, `sp-views`, `sp-security` ported (read-only + list/item +
file/folder CRUD + role assignments/definitions; write ops beyond that —
site/list/web create-delete, chunked upload, search pagination, user/group
add-remove, content-type create-delete, view mutations, role assignment
add-remove — are deferred and noted in each crate's README).
`sp-site-groups` reuses `SiteUserInfo` from `sp-site-users` (first
cross-crate type reuse). ~27 PnPjs `sp/*` packages remain untouched
(fields, sharing, forms, hubsites, recycle-bin, attachments, navigation,
comments, appcatalog, ...).

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

M1 vertical slice (auth + sp-sites + sp-lists + sp-files) is code-complete:
built, fmt/clippy/test-clean, wiring-tested against mock transports
(`tests/*.rs` in each crate). **Not yet live-verified** — `live_sites.rs`,
`live_webs.rs`, `live_lists.rs`, `live_files.rs`, `live_search.rs`,
`live_site_users.rs`, `live_folders.rs`, `live_content_types.rs`,
`live_site_groups.rs`, `live_views.rs`, `live_security.rs`,
`live_device_code.rs` exist under `forge-m365/examples/` but none have been
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

Next: continue M2 (remaining ~30 PnPjs `sp/*` domains — see `SPEC.md` §8,
full package list surveyed under `E:\sources\pnpjs\packages\sp\`). Owner
has directed continuing without per-crate check-ins since live testing
isn't currently possible either way.

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
