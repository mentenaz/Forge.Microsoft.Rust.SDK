# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Required reading before any change

- [`SPEC.md`](SPEC.md) — full specification: architecture, workspace layout, API generation
  strategy, auth flows, escalation ladder, milestones, reference sources with pinned local
  paths/commits.
- [`AGENTS.md`](AGENTS.md) — working-state summary: which crates exist, what's ported vs.
  deferred, live-verification status, locked decisions, domain knowledge that must survive
  sessions. **This is the current source of truth for project state** — it is updated every
  session; this file is not.

Read both in full before proposing architectural changes. Don't re-derive decisions already
locked in AGENTS.md's "Locked decisions" section, and don't duplicate SPEC.md/AGENTS.md content
into new docs.

## Build / lint / test

```bash
# All-in-one — this is what CI (ci.yml, ubuntu-latest + windows-latest) runs; all four must pass
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

No `rustfmt.toml` or `clippy.toml` — default Rust formatter/linter settings apply.

Live-tenant examples live under `forge-m365/examples/live_*.rs` (one per domain crate, plus
`live_auth.rs` and `live_device_code.rs`). Run with `cargo run --example live_sites -p forge-m365`
etc. **None have been run against a real tenant except `live_auth.rs`'s client-credentials path**
— see AGENTS.md before claiming any ported surface "works." Compiled ≠ verified.

Kaspersky AV on the dev machine false-positives freshly built binaries
(`VHO:Trojan-Ransom.Convagent.gen`); `target\*` must stay excluded.

## What this is

A full-scope Rust implementation of the Microsoft 365 ecosystem (Graph + SharePoint REST + admin
surfaces), using PnP PowerShell / PnP Framework / PnP.js as the functional spec — not a port of
their code, but a from-scratch Rust implementation built to match their behavior. Library-only,
consumed by Tauri/GPUI/terminal hosts. Private until stable; crates publish to crates.io as
`forge-m365-*` when ready.

## Architecture

**Fine-grained workspace**: one small crate per SharePoint/Graph endpoint domain
(~100–300 lines of logic each; mechanically-generated crates may exceed that), modeled on the
sibling project `E:\Forge.Rust.SDK` (proc-macro + inventory registration, thin dispatch, isolated
logic crates). `forge-m365` is the facade crate that re-exports everything.

**Escalation ladder** (`forge-m365-core`): every operation is registered at compile time as an
`OperationEntry { id, operation_path, ladder: Ladder { primary: Surface, fallback: &[Surface] } }`
via `inventory::collect!`. `Client::run_ladder` walks `primary` then each `fallback` surface in
order, returning the first success; a surface-specific failure (`Error::Surface`) falls through to
the next entry, any other error aborts immediately, and exhausting the ladder yields
`Error::Unsupported(id)` — the SDK should never surface a raw HTTP error where a documented
fallback exists. `forge-m365-macros`' `#[pnp_operation(...)]` is the intended authoring surface for
this registration (see `SPEC.md` §3), patterned on Forge.Rust.SDK's `#[forge_action]`.

**`Transport` trait** (`forge-m365-core`): `execute(surface, method, url, headers, body) -> Result<Vec<u8>>`,
async via a boxed future (no async-trait dependency). `forge-m365-auth`'s `AuthedTransport`
implements it, injecting bearer tokens for the right audience and defaulting
`Content-Type: application/json;odata=verbose` whenever a body is present (unless the caller
already set a Content-Type — file upload overrides to `application/octet-stream`). Hosts/tests
supply their own `Transport` impl; test crates use a recording transport to assert on the actual
request shape (method, URL, headers, body), not just the mocked response.

**Hybrid API generation**: Microsoft Graph will be generated from CSDL/OpenAPI metadata (not yet
started — see AGENTS.md M2/M3 status). SharePoint `_api` REST is hand-ported from PnPjs
(`E:\sources\pnpjs\packages\sp\<domain>`) because CSDL coverage there is incomplete/inconsistent.
Every ported crate records its PnPjs source path + commit SHA in its own README so parity claims
stay auditable — check AGENTS.md and the crate's README for what's ported vs. deferred before
assuming a write operation exists.

**Two OData dialects, one core**: `_api` (SharePoint REST) is OData v3-flavored; Graph is v4.
`forge-m365-core` has to parse/model both — don't assume Graph-shaped response handling works
for SharePoint REST or vice versa.

**Auth** (`forge-m365-auth`): UI-event driven — flows emit "open this URL" / "poll this code"
rather than blocking on stdin, so each host can render its own UX. All flows are deliberately
*synchronous* functions (not `async fn`): they block on I/O (TCP accept for the loopback
redirect, device-code polling sleep) rather than depending on a specific async runtime, per the
library's no-owned-runtime rule; async hosts call them via `spawn_blocking`. Graph and SharePoint
are separate token audiences managed transparently — a caller never manually juggles two tokens.
Client-credentials (secret and certificate/RS256 JWT assertion) is live-verified against a real
tenant; interactive-browser and device-code flows are implemented but untested live (see AGENTS.md
for why).

**CSOM is out of scope as a wire protocol** — `Microsoft.SharePointOnline.CSOM` is closed-source
(binary nupkg reference only at `E:\sources\CSOM-binaries`). CSOM-only operations map to a
documented REST equivalent or become a typed `Unsupported` error; never reimplemented from the
binary.

**Communication sites** (no M365 group anchor) aren't fully visible via Graph and require
SharePoint REST (`_api/web`, `_api/search/query`) plus admin endpoints. Both API clients (Graph
and SharePoint REST) are permanent first-class citizens of this SDK, not a legacy path to
deprecate in favor of Graph.

**Library-only runtime model**: plain `async fn` surface (outside the deliberately-sync auth
flows above), no owned Tokio runtime, no `#[tokio::main]` inside the SDK. Hosts
(Tokio/GPUI-executor/smol/block-on CLIs) drive execution.

## Working rules specific to this repo (from AGENTS.md — do not relitigate)

- Crate namespace `forge-m365-*`, facade crate `forge-m365`. No dots in crate names; avoid
  `microsoft-`/`pnp-` prefixes (crates.io trademark reasons).
- When porting a module from a reference source, record that source + commit (e.g.
  "pnpjs@\<sha\> packages/sp/files") in the crate's README so parity stays auditable. Never claim
  generated/ported code was tested against a live tenant unless it actually was — compiled ≠
  tested, and tested-with-a-mock-transport ≠ verified live.
- Publishing to crates.io or changing repo visibility is the owner's action only. Do not create
  issues/PRs unprompted.
- Reference source checkouts are pinned, read-only local paths under `E:\` (PnP PowerShell, PnP
  Framework, PnP Core SDK, PnP.js, MSAL.NET, msgraph-sdk-dotnet, etc.) — see `SPEC.md` §9 and
  AGENTS.md's table for exact paths and commit pins. Consult those directly rather than guessing
  paths or versions.
