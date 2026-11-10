use serde::{Deserialize, Serialize};

/// Parameters for deleting an item with options.
///
/// Ported from PnPjs `IItemDeleteParams` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Serialize)]
pub struct ItemDeleteParams {
    /// If true, delete or recycle a file when the LockType
    /// value is SPLockType.Shared or SPLockType.None.
    /// When false, delete or recycle the file when
    /// the LockType value SPLockType.None.
    #[serde(rename = "BypassSharedLock")]
    pub bypass_shared_lock: bool,
}

/// Result returned after updating a list item.
///
/// Ported from PnPjs `IItemUpdateResultData` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct ItemUpdateResult {
    /// The ETag value returned in the response header after update.
    #[serde(rename = "etag")]
    pub etag: Option<String>,
}

/// Result returned after uploading an image to an item field.
///
/// Ported from PnPjs `IItemImageUploadResult` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct ItemImageUploadResult {
    #[serde(rename = "Name")]
    pub name: String,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
    #[serde(rename = "UniqueId")]
    pub unique_id: String,
}

/// A value to validate and set on a list item field.
///
/// Ported from PnPjs `IListItemFormUpdateValue` (`packages/sp/lists/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ListItemFormUpdateValue {
    #[serde(rename = "FieldName")]
    pub field_name: String,
    #[serde(rename = "FieldValue")]
    pub field_value: String,
}

/// Parent list and web information for an item.
///
/// Ported from PnPjs `IItemParentInfos` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct ItemParentInfos {
    #[serde(rename = "Item")]
    pub item: ItemParentId,
    #[serde(rename = "ParentList")]
    pub parent_list: ParentListInfo,
    #[serde(rename = "ParentWeb")]
    pub parent_web: ParentWebInfo,
}

#[derive(Debug, Clone, Deserialize)]
pub struct ItemParentId {
    #[serde(rename = "Id")]
    pub id: String,
}

#[derive(Debug, Clone, Deserialize)]
pub struct ParentListInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "RootFolderServerRelativePath")]
    pub root_folder_server_relative_path: ResourcePath,
    #[serde(rename = "RootFolderServerRelativeUrl")]
    pub root_folder_server_relative_url: String,
    #[serde(rename = "RootFolderUniqueId")]
    pub root_folder_unique_id: String,
}

#[derive(Debug, Clone, Deserialize)]
pub struct ParentWebInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "ServerRelativePath")]
    pub server_relative_path: ResourcePath,
    #[serde(rename = "ServerRelativeUrl")]
    pub server_relative_url: String,
    #[serde(rename = "Url")]
    pub url: String,
}

/// A SharePoint resource path object.
///
/// Ported from PnPjs `IResourcePath` (`packages/sp/utils/to-resource-path.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct ResourcePath {
    #[serde(rename = "DecodedUrl")]
    pub decoded_url: String,
}

/// Information about a single version of a list item.
///
/// Ported from PnPjs `IItemVersionInfo` (`packages/sp/items/types.ts`)
/// @ pnpjs `8ee2375d`.
#[derive(Debug, Clone, Deserialize)]
pub struct VersionInfo {
    #[serde(rename = "VersionLabel")]
    pub version_label: String,
    #[serde(rename = "Id")]
    pub id: i64,
    #[serde(rename = "Created")]
    pub created: String,
    #[serde(rename = "CreatedBy")]
    pub created_by: ValueInfo,
    #[serde(rename = "LastModified")]
    pub last_modified: String,
    #[serde(rename = "LastModifiedBy")]
    pub last_modified_by: ValueInfo,
    #[serde(rename = "CheckInComment")]
    pub check_in_comment: String,
    #[serde(rename = "Size")]
    pub size: i64,
}

/// A user or principal value reference in SharePoint REST responses.
#[derive(Debug, Clone, Deserialize)]
pub struct ValueInfo {
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "Id")]
    pub id: Option<i64>,
}
