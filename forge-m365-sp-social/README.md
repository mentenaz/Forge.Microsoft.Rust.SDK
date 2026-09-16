# forge-m365-sp-social

SharePoint social/following operations: follow/unfollow/check-follow an
actor (user, document, site, or tag), and read the current user's followed
sites/documents/followers/suggestions. Hand-ported from PnPjs, per
`SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so this surface
is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\social` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `_Social` (follow/isFollowed/stopFollowing/getFollowedSitesUri/getFollowedDocumentsUri), `_MySocial` (my/followed/followedCount/followers/suggestions) |

This is a legacy/deprecated SharePoint surface (`_api/social.following`),
superseded in modern tenants by Viva Engage/Graph, but it's still the
documented API and is ported completely and faithfully — nothing in this
PnPjs package is skipped.

`SocialActorType`/`SocialActorTypes`/`SocialStatusCode` are ported as plain
integer constants (`actor_type`/`actor_types` modules) rather than Rust
enums since `SocialActorTypes` is a bitwise set (not a discriminant) and
`SocialStatusCode` only ever appears embedded in `SocialActor::status`,
never as a return type worth strongly typing on its own.
`SocialFollowResult` — the one enum actually returned as a value from
`follow` — is a proper Rust enum with an `Other(i64)` fallback for any
undocumented code, so an unrecognized server response deserializes instead
of failing.

## Verification status

Wiring — URL construction for every accessor/action, the verbose-OData
`{"<Key>": value}`/bare-value envelope PnPjs's `hOP()` peels off, and
`{"results": [...]}` collection unwrapping — is covered by `tests/social.rs`
against a recording transport. Compiled ≠ verified against a live tenant —
see `AGENTS.md` for the live-verification record before relying on this
against production data, and note that `follow`/`stop_following` mutate a
real user's following list once run live.
