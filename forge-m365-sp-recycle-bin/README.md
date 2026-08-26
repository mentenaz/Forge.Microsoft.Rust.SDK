# forge-m365-sp-recycle-bin

SharePoint recycle-bin operations. Hand-ported from PnPjs, per `SPEC.md` §4
(SharePoint REST has no reliable CSDL coverage, so this surface is ported
rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\recycle-bin` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `RecycleBin` collection reads, `RecycleBinItem.restore`/`delete` |

Only single-item operations are ported. The bulk operations
(`deleteAll`, `deleteAllSecondStageItems`, `moveAllToSecondStage`,
`restoreAll`) are deliberately not — their blast radius (every item in the
recycle bin, tenant-wide in effect) is high enough that they're left for a
deliberate follow-up rather than bundled in here by default. `moveToSecondStage`
(single item) is also deferred.

## Verification status

Wiring — URL construction (`Restore`/`DeleteObject` action URLs) and
response parsing — is covered by `tests/recycle_bin.rs` against a recording
transport. Compiled ≠ verified against a live tenant — see `AGENTS.md` for
the live-verification record before relying on this against production
data, and note that `delete_recycle_bin_item` is **not reversible** — it
bypasses the second-stage recycle bin entirely.
