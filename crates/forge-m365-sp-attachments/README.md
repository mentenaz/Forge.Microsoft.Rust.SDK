# forge-m365-sp-attachments

SharePoint list-item attachment operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this
surface is ported rather than generated). Complements `sp-lists`.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\attachments` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Attachments.add`, collection reads, `Attachment.delete`, `ReadableFile.getBuffer` |

Not yet ported: `setContent` (replace an attachment's bytes) and `recycle`
(soft-delete to the recycle bin instead of a hard delete).

## Verification status

Wiring — request construction (raw-bytes `add`, `$value` download URL,
`IF-MATCH` on delete) and response parsing — is covered by
`tests/attachments.rs` against a recording transport. Compiled ≠ verified
against a live tenant — see `AGENTS.md` for the live-verification record
before relying on this against production data, and note that
`add_attachment`/`delete_attachment` mutate real data once run live.
