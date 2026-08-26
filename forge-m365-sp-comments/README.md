# forge-m365-sp-comments

SharePoint list-item comment operations. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this
surface is ported rather than generated). Complements `sp-lists`.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\comments` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `Comments.add`, collection reads, `Comment.delete`/`like`/`unlike` |

## A different wire format from every other crate

Unlike every other crate in this SDK, this API's JSON fields are
**camelCase** (`text`, `createdDate`, `likeCount`), not PascalCase
(`Title`, `Created`). It's SharePoint's newer "likes and comments" service,
addressed under the same `_api/web/lists/.../items(id)/comments` item path
as the classic REST surface but with its own distinct response schema —
not a typo or an inconsistency introduced during porting.

Not yet ported: replies (`Comments.replies`/`Replies.add`), `Comments.clear`
(delete all), and likes/mentions detail beyond the summary counts already
on `CommentInfo`.

## Verification status

Wiring — request construction (`text` body shape, `Like`/`Unlike`/delete
action URLs) and response parsing (including the nested camelCase
`author` object) — is covered by `tests/comments.rs` against a recording
transport. Compiled ≠ verified against a live tenant — see `AGENTS.md` for
the live-verification record before relying on this against production
data, and note that `like_comment`/`unlike_comment` need a delegated
(signed-in user) token to be meaningful — app-only auth has no "current
user" to like as.
