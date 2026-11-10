# forge-m365-sp-security

SharePoint role assignment/definition (permissions) operations. Hand-ported
from PnPjs, per `SPEC.md` §4 (SharePoint REST has no reliable CSDL
coverage, so this surface is ported rather than generated). Complements
`sp-site-users`/`sp-site-groups`.

## Port source

| Source | Path | Commit | Scope |
|---|---|---|---|
| PnPjs | `E:\sources\pnpjs\packages\sp\security` | `8ee2375dc3ce720926774cc87168e1251c9ed584` | `RoleAssignments` reads, `RoleDefinitions.getById`, `ISecurableMethods.breakRoleInheritance`/`getUserEffectivePermissions` |

`get_role_assignments` deliberately expands `Member`/`RoleDefinitionBindings`
even though PnPjs's own `IRoleAssignmentInfo` is just `{ PrincipalId }` —
PnPjs expects callers to chain `.expand()`/`.select()` themselves, but this
SDK has no query builder yet, so the useful shape is baked in directly
rather than shipping a result nobody could use without a second round trip.

`get_role_assignments`/`break_role_inheritance` take a `securable_url`
(the securable object's own REST endpoint — a web, list, or item) rather
than a `site_url`, since role assignments apply generically to any of those
in SharePoint's security model (PnPjs's `ISecurableMethods` mixin).

Not yet ported: adding/removing role assignments, creating/updating/deleting
role definitions, `resetRoleInheritance`, `userHasPermissions`/
`currentUserHasPermissions`, and `firstUniqueAncestorSecurableObject`.

## Verification status

Wiring — request construction (`breakroleinheritance` URL-embedded flags,
`getUserEffectivePermissions`'s `@user` aliased parameter) and response
parsing (including both wrapped and bare shapes for the outer array and the
nested `RoleDefinitionBindings` collection) — is covered by
`tests/security.rs` against a recording transport. Compiled ≠ verified
against a live tenant — see `AGENTS.md` for the live-verification record
before relying on this against production data, and note that
`break_role_inheritance` mutates real permissions once run live.
