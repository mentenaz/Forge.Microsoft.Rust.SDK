# forge-m365-sp-files

SharePoint file upload/download. Hand-ported from PnPjs, per `SPEC.md` §4
(SharePoint REST has no reliable CSDL coverage, so this surface is ported
rather than generated). Completes the M1 vertical slice (`SPEC.md` §8):
auth + `sp-sites` + `sp-lists` + `sp-files`, live-verified end to end.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\files` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Files.addUsingPath`, `File.delete`, `Web.getFileByServerRelativePath` |
| PnPjs | `E:\sources\pnpjs\packages\sp\files\readable-file.ts` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `ReadableFile.getBuffer` (download) |

Not yet ported: chunked upload for large files (`Files.addChunked` /
`startUpload` / `continueUpload` / `finishUpload`), check-in/check-out,
versions, copy/move, and folder management (`packages/sp/folders`) beyond
the single `getFolderByServerRelativePath` call `upload_file` needs.

## Why upload needed another core fix

`AddUsingPath` takes the file's raw bytes as the request body, not JSON —
unlike every write operation `sp-lists` added. The default
`Content-Type: application/json;odata=verbose` that `AuthedTransport` sets
for any request with a body would have been wrong here, and since
`reqwest`'s `.header()` appends rather than replaces, the caller's own
`Content-Type` would have collided with it instead of overriding it.
`AuthedTransport::raw` now only applies the default when the caller hasn't
already supplied a `Content-Type` in `headers`.

## Verification status

Wiring — request construction (method, URL, headers, raw body vs JSON body)
and response parsing — is covered by `tests/files.rs` against a recording
transport. Compiled ≠ verified against a live tenant — see `AGENTS.md` for
the live-verification record before relying on this against production
data, and note that `upload_file`/`delete_file` mutate real data once run
live.
