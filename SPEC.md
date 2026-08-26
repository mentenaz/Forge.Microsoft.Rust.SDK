# Forge.Microsoft.Rust.SDK — Specification

**Status**: planning / pre-implementation · **Visibility**: private until stable
**Repo**: https://github.com/mentenaz/Forge.Microsoft.Rust.SDK

A full-scope Rust implementation of the Microsoft 365 ecosystem — Microsoft
Graph, SharePoint REST and the tenant-admin surfaces — replacing the C#
(PnP Framework / PnP Core SDK / PnP PowerShell) and TypeScript (PnP.js)
stacks. Designed to be consumed as a library from Tauri apps, GPUI apps and
terminal applications.

Agent context lives in [`AGENTS.md`](AGENTS.md) alongside this document.

---

## 1. Goals and non-goals

### Goals

1. **Full parity ambition** with PnP PowerShell + PnPjs functionality,
   built incrementally by milestone.
2. **Library-first**: no owned runtime, no CLI coupling. Hosts drive execution.
3. **Resilient data access**: every operation declares a fallback ladder so a
   failure on one API surface degrades gracefully instead of erroring out.
4. **Auditable parity**: each ported operation records which reference source
   and commit it was ported from.

### Non-goals

- Not a CLI tool itself (a terminal app is a *consumer*).
- No CSOM wire-protocol reimplementation (closed-source; see §7).
- Not affiliated with the PnP community org or Microsoft; independent
  implementation against public APIs.

---

## 2. Naming

| Artifact | Name |
|---|---|
| Repository | `Forge.Microsoft.Rust.SDK` |
| Crate namespace | `forge-m365-*` |
| Facade crate | `forge-m365` |
| Example domain crates | `forge-m365-core`, `forge-m365-auth`, `forge-m365-sp-files`, `forge-m365-sp-lists`, `forge-m365-sp-sites`, `forge-m365-graph-users` |

Rules: dots are invalid in crate names (`pnp.core` ✗); hyphens in crate names
become underscores in imports (`use forge_m365::sp::lists`). The `microsoft-`
and `pnp-` prefixes are avoided on crates.io for trademark reasons.

---

## 3. Workspace architecture

Fine-grained workspace following the pattern proven in
`E:\Forge.Rust.SDK` (proc-macro + inventory registration, thin dispatch,
isolated logic crates). One small crate per endpoint domain; a facade crate
re-exports the full surface.

```
forge-m365/                  facade crate — pub use of everything below
forge-m365-core/             HTTP pipeline, OData v3/v4 handling, retry,
                             batching, caching, escalation-ladder engine
forge-m365-auth/             token acquisition, all flows (§5), token cache,
                             dual-audience management (Graph vs SharePoint)
forge-m365-macros/           #[pnp_operation(...)] proc-macro: registers an
                             operation with its API-surface ladder at compile
                             time (modeled on Forge.Rust.SDK #[forge_action])
forge-m365-graph/            Graph domain root (CSDL/OpenAPI-generated types)
forge-m365-graph-users/      generated per-domain crates…
forge-m365-sp/               SharePoint domain root (hand-ported endpoints)
forge-m365-sp-sites/         hand-ported per-domain crates…
forge-m365-sp-lists/
forge-m365-sp-files/
forge-m365-sp-search/
…                            ~40+ sp domains mirroring PnPjs packages/sp/*,
                             plus graph domains mirroring PnPjs packages/graph/*
```

Sizing note: target 100–300 lines of *logic* per domain crate is accepted;
generated type crates may exceed it mechanically.

### The escalation ladder (core concept)

Each operation self-describes its API surfaces at compile time:

```rust
#[pnp_operation(
    id = "sp.lists.get_items",
    primary = SpRest("GET {site}/_api/web/lists/getbytitle('{list}')/items"),
    fallback = [Graph("GET /sites/{site}/lists/{list}/items"), Legacy(Search)],
    unsupported = "taxonomy-classic"   // typed Unsupported error if all fail
)]
pub async fn get_items(...) -> Result<ItemPage, Error> { ... }
```

Runtime behavior on failure: try `primary`; on surface-specific failure
(404 endpoint-missing, feature-not-enabled, permission shape mismatch) fall
through the declared list; if exhausted, return a typed
`Error::Unsupported(reason)` — never a raw HTTP error where a workaround
exists.

---

## 4. API generation strategy (hybrid)

| Surface | Strategy | Rationale |
|---|---|---|
| Microsoft Graph | **Generated** from CSDL/OpenAPI metadata | Complete, current metadata published by Microsoft; free typed structs |
| SharePoint `_api` REST | **Hand-ported** from PnPjs | CSDL incomplete/inconsistent; PnPjs encodes years of quirk knowledge |
| Tenant admin | Hand-ported from `pnpjs/packages/sp-admin` + PnP Core admin layer | Partially undocumented; needs deliberate modeling |

Both sides sit behind uniform traits so callers never see which surface served
a call.

---

## 5. Authentication

All flows in scope, sequenced:

1. **v0.0.1** — client credentials (app-only) with certificate
2. device code
3. interactive browser (local loopback redirect)
4. managed identity; Windows WAM broker (platform-gated)

Design constraints:

- **Dual audience**: Graph tokens and SharePoint tokens are different
  audiences; core manages both transparently per request.
- **UI-event driven**: auth emits events ("open this URL", "poll this code")
  via callback/channel rather than blocking on stdin, so Tauri can render a
  window and GPUI/CLI can each present their own UX.
- Token cache: file-backed, cross-platform path handling; MSAL.NET token
  cache semantics as reference.

---

## 6. Runtime model

- Pure `async fn` library surface; **no** `#[tokio::main]`, no runtime
  ownership inside the SDK.
- Compatible hosts: Tokio (Tauri), GPUI executor, smol/async-std, block-on CLIs.
- HTTP: `reqwest` (+ middleware traits for behaviors: retry, logging, caching).
- Feature flags keep platform-specific code (WAM broker, keychain) optional.

---

## 7. Known constraints

1. **CSOM is unportable.** `Microsoft.SharePointOnline.CSOM` is closed-source
   binary-only (nupkg archived at `E:\sources\CSOM-binaries`, v
   `16.1.27515.12000`; no symbols exist). CSOM-only operations are either
   reimplemented against documented REST equivalents or surfaced as typed
   `Unsupported`. Reference for "what is CSOM-only": PnP Core's model layer
   marks which entities bypass REST.
2. **Communication sites / non-group-connected sites are not fully visible to
   Graph.** They lack a Microsoft 365 group anchor; enumeration requires
   SharePoint REST (`_api/web`, `_api/search/query`) and admin endpoints.
   Both API clients are therefore permanent first-class citizens.
3. **OData dialects differ**: `_api` is OData v3-flavored, Graph is v4.
   Core must parse/model both.
4. Some tenant-admin endpoints are undocumented; they are modeled from
   observed traffic and community knowledge (PnPjs `sp-admin`).

---

## 8. Milestone plan

| Milestone | Scope |
|---|---|
| M0 | Workspace scaffold, `core` pipeline skeleton, auth client-credentials, CI, name reservation on crates.io |
| M1 (v0.0.1) | Auth + `sp-sites`, `sp-lists`, `sp-files`: enumerate sites, CRUD list items, upload/download files end-to-end against a live dev tenant. Exercises auth → dual audience → OData parse → escalation ladder |
| M2 | Device-code auth; remaining `sp-*` domains ported from PnPjs |
| M3 | Graph domains via generator; facade complete |
| M4 | Interactive browser auth, managed identity; admin surfaces |
| M5 | Parity audit vs PnP PowerShell cmdlet list; docs; crates.io publish |

---

## 9. Reference sources (local paths)

| Source | Path | Version/pin | Use |
|---|---|---|---|
| This repo | `E:\Forge.Microsoft.Rust.SDK` | private | — |
| Planning workspace (this doc's origin) | `E:\Microsoft.Rust.Crates` → renamed | local git | historical notes |
| PnP PowerShell (C#) | `E:\powershell` | branch state at planning time | functional spec: 800+ cmdlets |
| PnP Framework | `E:\pnpframework` | commit `e7631b62` (per dependencies.json) | provisioning behavior reference |
| PnP Core SDK | `E:\pnpcore` | commit `d90d3216f` (per dependencies.json) | REST layer = primary C# reference; entity→surface mapping |
| PnP.js | `E:\sources\pnpjs` | depth-1 main | **pure REST endpoint map**; `packages/sp/*` = hand-port reference, `packages/core`+`queryable` = pipeline architecture |
| PowerShell 7.4 | `E:\sources\PowerShell` | tag v7.4.0 | System.Management.Automation semantics |
| MSAL.NET | `E:\sources\MSAL-dotnet` | tag 4.85.2 | token flows/cache semantics; includes Extensions.Msal |
| Application Insights | `E:\sources\ApplicationInsights-dotnet` | tag 2.21.0 | telemetry schema reference |
| TextCopy | `E:\sources\TextCopy` | tag 6.2.1 | clipboard utility reference |
| .NET BCL | `E:\sources\dotnet-runtime` | tag v9.0.5 | Bcl.Cryptography / Bcl.AsyncInterfaces internals |
| MS Graph SDK (.NET) | `E:\sources\msgraph-sdk-dotnet` | main snapshot | Graph surface reference |
| SPO CSOM binaries | `E:\sources\CSOM-binaries` | nupkg 16.1.27515.12000 | closed-source; decompile-for-reference only |

Parity rule: every ported module records its source repo + commit in its crate
README (dependencies.json-style pinning) so parity claims stay auditable.
