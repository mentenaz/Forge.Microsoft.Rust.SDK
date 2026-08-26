# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Current state: PLANNING / PRE-IMPLEMENTATION

No Rust code exists yet — no `Cargo.toml`, no workspace, no source files. There
are no build/lint/test commands to run because there is nothing to build yet.
The next concrete step is **M0**: workspace scaffold + `core` pipeline
skeleton + client-credentials auth (see `SPEC.md` §8). Do not assume a
workspace layout beyond what §3 of `SPEC.md` describes — it hasn't been
created.

## Required reading before any change

- [`SPEC.md`](SPEC.md) — full specification: architecture, workspace layout,
  API generation strategy, auth flows, escalation ladder, milestones,
  reference sources with pinned local paths/commits.
- [`AGENTS.md`](AGENTS.md) — working-state summary: locked decisions, domain
  knowledge that must survive sessions, working rules for agents.

These two files are the source of truth for this repo. Read them in full
before proposing architectural changes — do not re-derive decisions that are
already locked in `AGENTS.md` ("Locked decisions" section), and do not
duplicate their content into new docs.

## Working rules specific to this repo (from AGENTS.md — do not relitigate)

- Crate namespace `forge-m365-*`, facade crate `forge-m365`. No dots in crate
  names; avoid `microsoft-`/`pnp-` prefixes (crates.io trademark reasons).
- Fine-grained workspace: one small crate per endpoint domain (~100–300 lines
  of logic each; generated crates may exceed this mechanically), modeled on
  the sibling project `E:\Forge.Rust.SDK` (proc-macro + inventory
  registration, thin dispatch, isolated logic crates).
- API generation is **hybrid**: Microsoft Graph is generated from CSDL/OpenAPI
  metadata; SharePoint `_api` REST is hand-ported from PnPjs
  (`E:\sources\pnpjs\packages\sp`) because CSDL coverage there is
  incomplete/inconsistent.
- Every operation declares an escalation ladder at compile time via
  `#[pnp_operation(...)]` (primary surface → fallback surfaces → typed
  `Unsupported`) — see `SPEC.md` §3 for the macro shape. Never let an
  operation fail with a raw HTTP error where a documented fallback exists.
- Auth is UI-event driven (emits "open this URL" / "poll this code" via
  callback/channel) and never blocks on stdin, so each host (Tauri/GPUI/CLI)
  can render its own UX. Graph and SharePoint use separate token audiences
  managed transparently by `forge-m365-core`/`forge-m365-auth`.
- Library-only: plain `async fn` surface, no owned Tokio runtime, no
  `#[tokio::main]` inside the SDK. Hosts (Tokio/GPUI executor/smol/block-on
  CLIs) drive execution.
- CSOM (`Microsoft.SharePointOnline.CSOM`) is closed-source and out of scope
  as a wire protocol. CSOM-only operations map to a documented REST
  equivalent or become a typed `Unsupported` error — never reimplemented.
- `_api` is OData v3-flavored, Graph is v4 — `forge-m365-core` must parse and
  model both dialects.
- Communication sites (no M365 group anchor) are not fully visible via Graph;
  SharePoint REST (`_api/web`, `_api/search/query`) and admin endpoints are
  required for them. Both API clients are permanent first-class citizens —
  don't treat SharePoint REST as a legacy path to be deprecated.
- When porting a module from a reference source, record that source + commit
  (e.g. "pnpjs@\<sha\> packages/sp/files") in the crate's README so parity
  stays auditable. Never claim generated/ported code was tested against a
  live tenant unless it actually was — compiled ≠ verified.
- Publishing to crates.io or changing repo visibility is the owner's action
  only. Do not create issues/PRs unprompted.

## Reference source checkouts (read-only, local to this machine)

`SPEC.md` §9 and `AGENTS.md` list pinned local paths used as functional-spec
and hand-port references (PnP PowerShell, PnP Framework, PnP Core SDK, PnP.js,
MSAL.NET, msgraph-sdk-dotnet, etc., all under `E:\` — e.g. `E:\sources\pnpjs`,
`E:\pnpcore`, `E:\powershell`). Consult those tables directly rather than
guessing paths or versions; they carry exact commit/tag pins that matter for
parity claims.
