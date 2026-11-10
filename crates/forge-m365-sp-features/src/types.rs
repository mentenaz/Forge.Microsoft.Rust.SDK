use serde::Deserialize;

/// A feature as returned by the site/web `features` collection service
/// (`packages/sp/features/types.ts`). PnPjs models it with a lone
/// `DefinitionId`; the payload tolerates the extra fields SharePoint returns.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct FeatureInfo {
    #[serde(rename = "DefinitionId", default)]
    pub definition_id: String,
}

/// Boxed value returned by the add (activation) action — in PnPjs terms the
/// `data` half of `IFeatureAddResult` (the `feature` half is a nav handle,
/// which this library does not model).
#[derive(Debug, Clone, Default, Deserialize)]
pub struct FeatureAddResult {
    #[serde(rename = "DefinitionId", default)]
    pub definition_id: String,
}
