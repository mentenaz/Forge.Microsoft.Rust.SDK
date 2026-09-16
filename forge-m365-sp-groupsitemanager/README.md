# forge-m365-sp-groupsitemanager

Microsoft 365 group-connected site creation, status lookup, and related
tenant configuration reads (`_api/groupsitemanager/*`). Hand-ported from
PnPjs, per `SPEC.md` §4 (SharePoint REST has no reliable CSDL coverage, so
this surface is ported rather than generated).

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\groupsitemanager` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `GroupSiteManager.canUserCreateGroup/clearCurrentUserTeamsCache/create/createGroup/delete/ensureTeamForGroup/ensureTeamForGroupEx/getAllOrgLabels/getGroupCreationContext/getGroupSiteConversionData/getSiteStatus/getValidSiteUrlFromAlias/isTeamifyPromptHidden/notebook` |

`GroupCreationContext` is `serde_json::Map<String, Value>`, not a typed
struct — PnPjs's own `IGroupCreationContext` interface mixes clean fields
with `{ Key, Value: any, ValueType }` tuples whose `Value` shape is
tenant-configuration-dependent and undocumented, so (as with
`sp-appcatalog::AppInfo`) there's no reference field list to port faithfully.

`ensure_team_for_group`, `get_valid_site_url_from_alias`, and `notebook`
assume their response body is a bare JSON string, matching PnPjs's
`Promise<string>` return type for those methods — this is unverified against
a live tenant (see "Verification status" below).

Not yet ported: the Teams-channel-enumeration methods
(`getCurrentUserJoinedTeams`, `getCurrentUserSharedChannelMemberGroups`,
`getCurrentUserTeamConnectedMemberGroups`, `getTeamChannelFilesUrl`,
`getTeamChannels`, `getTeamChannelsDirect`, `getTeamChannelsEx`,
`getTeamChannelsWithSiteUrl`, `getUserSharedChannelMemberGroups`,
`getUserTeamConnectedMemberGroups`, `pinToTeam`) and the generic `call<T>`
escape hatch — PnPjs documents most of these as requiring
`@pnp/sp` "behaviors" (delegated/interactive-session headers) and as not
working under Azure AD application (client-credentials) context at all,
which this SDK's current auth model doesn't yet distinguish from ordinary
app-only calls. Deferred pending a deliberate pass on delegated-auth
behaviors rather than shipping calls known not to work under this SDK's
primary (app-only) auth flow.

## Verification status

Wiring — URL construction for every accessor/action and request/response
JSON shape — is covered by `tests/groupsitemanager.rs` against a recording
transport. Compiled ≠ verified against a live tenant — see `AGENTS.md` for
the live-verification record before relying on this against production
data. `create`, `create_group`, and `delete` create or destroy real
group-connected sites once run live.
