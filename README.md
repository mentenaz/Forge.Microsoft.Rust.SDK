# Forge.Microsoft.Rust.SDK

Full-scope Rust implementation of the Microsoft 365 ecosystem — Microsoft
Graph, SharePoint REST and tenant-admin surfaces — replacing the C#
(PnP Framework / PnP Core SDK / PnP PowerShell) and TypeScript (PnP.js)
stacks. Consumed as a library by Tauri, GPUI and terminal applications.

**Status**: planning / pre-implementation. Private until stable; crates
publish to crates.io as `forge-m365-*` when ready.

## Documentation

- [`SPEC.md`](SPEC.md) — full specification: architecture, API generation
  strategy, auth flows, escalation ladders, milestones, reference sources
- [`AGENTS.md`](AGENTS.md) — agent context: locked decisions, current state,
  working rules

## At a glance

| | |
|---|---|
| Crate namespace | `forge-m365-*`, facade crate `forge-m365` |
| Graph surface | generated from CSDL/OpenAPI metadata |
| SharePoint surface | hand-ported from PnPjs (pure REST endpoint map) |
| Resilience | per-operation fallback ladder: Graph → SP REST → legacy → typed `Unsupported` |
| Auth | client credentials → device code → interactive browser → managed identity; UI-event driven, dual token audiences |
| Runtime | async library only; host drives execution (Tauri / GPUI / CLI) |
| Known gap | SPO CSOM is closed-source — CSOM-only ops map to REST or become typed unsupported errors |

Functional spec sources: PnP PowerShell (~800 cmdlets), PnP Core SDK,
PnP.js — pinned local checkouts listed in `SPEC.md` §9.
