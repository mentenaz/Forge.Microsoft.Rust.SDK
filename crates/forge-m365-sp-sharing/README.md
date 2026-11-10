# forge-m365-sp-sharing

SharePoint web/file/folder/list-item sharing operations for the forge-m365 SDK,
hand-ported from PnP.js [`packages/sp/sharing`](https://github.com/pnp/pnpjs/tree/8ee2375d/packages/sp/sharing).

## Port source

- **Ref:** PnP.js commit `8ee2375d3ce720926774cc87168e1251c9ed584`
- **Path:** `packages/sp/sharing/`

## Covered operations

### Item sharing (`lists/getbytitle('...')/items({id})`)

| Operation | HTTP | REST path |
|-----------|------|-----------|
| `item_share_with(login_names, role, require_signin, email_data)` | GET+POST | `/roledefinitions`, `/shareObject` |
| `item_get_share_link(kind, expiration)` | POST | `/shareLink` |
| `item_check_permissions(recipients)` | POST | `/checkPermissions` |
| `item_get_sharing_information(request, expands, selects)` | POST | `/getSharingInformation` |
| `item_get_object_sharing_settings(use_simplified_roles)` | POST | `/getObjectSharingSettings` |
| `item_unshare()` | POST | `/unshareObject` |
| `item_delete_sharing_link_by_kind(kind)` | POST | `/deleteLinkByKind` |
| `item_unshare_link(kind, share_id)` | POST | `/unshareLink` |

### File sharing (`getFileByServerRelativePath(decodedUrl='...')/listItemAllFields`)

| Operation | HTTP | REST path |
|-----------|------|-----------|
| `file_share_with(...)` | GET+POST | `/roledefinitions`, `/shareObject` |
| `file_get_share_link(...)` | POST | `/shareLink` |
| `file_check_permissions(...)` | POST | `/checkPermissions` |
| `file_get_sharing_information(...)` | POST | `/getSharingInformation` |
| `file_get_object_sharing_settings(...)` | POST | `/getObjectSharingSettings` |
| `file_unshare()` | POST | `/unshareObject` |
| `file_delete_sharing_link_by_kind(...)` | POST | `/deleteLinkByKind` |
| `file_unshare_link(...)` | POST | `/unshareLink` |

### Folder sharing (`getFolderByServerRelativePath(decodedUrl='...')/listItemAllFields`)

| Operation | HTTP | REST path |
|-----------|------|-----------|
| `folder_share_with(...)` | GET+POST | `/roledefinitions`, `/shareObject` |
| `folder_get_share_link(...)` | POST | `/shareLink` |
| `folder_check_permissions(...)` | POST | `/checkPermissions` |
| `folder_get_sharing_information(...)` | POST | `/getSharingInformation` |
| `folder_get_object_sharing_settings(...)` | POST | `/getObjectSharingSettings` |
| `folder_unshare()` | POST | `/unshareObject` |
| `folder_delete_sharing_link_by_kind(...)` | POST | `/deleteLinkByKind` |
| `folder_unshare_link(...)` | POST | `/unshareLink` |

File/folder sharing targets the object's `listItemAllFields` navigation
instead of fetching the list item first — matching the endpoint PnPjs builds
when `_File`/`_Folder` resolve their item via `getItem()`.

### Web sharing (`_api/web`)

| Operation | HTTP | REST path |
|-----------|------|-----------|
| `web_share_with(login_names, role, email_data)` | GET+POST | `/roledefinitions`, `SP.Web.ShareObject` |
| `web_share_object(url, login_names, role, group, email_data)` | GET+POST | `/roledefinitions`/`associatedmembergroup`/`associatedvisitorgroup`, `SP.Web.ShareObject` |
| `web_share_object_raw(options)` | POST | `SP.Web.ShareObject` |
| `web_unshare_object(url)` | POST | `/unshareObject` |

### Internal lookups

| Operation | HTTP | REST path |
|-----------|------|-----------|
| `sp.sharing.role_definition_id` | GET | `/roledefinitions?$select=Id&$filter=RoleTypeKind eq ...` |
| `sp.sharing.associated_group_id` | GET | `/associatedmembergroup` / `/associatedvisitorgroup` |

`web_share_with` shares the web's `_layouts/15/aclinv.aspx?forSharing=1&mbypass=1`
page. `web_share_object` defaults `includeAnonymousLinkInEmail`,
`propagateAcl` and `useSimplifiedRoles` to false/false/true in PnPjs's
signature but never transmits them — the Rust signature therefore omits them.
An empty `share_id` on `*_unshare_link` resolves to the empty GUID
(`00000000-0000-0000-0000-000000000000`) like PnPjs's `emptyGuid` default.