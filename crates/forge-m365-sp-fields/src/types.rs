use serde::{Deserialize, Serialize};

/// Specifies the type of a SharePoint field.
///
/// Ported from PnPjs `FieldTypes` enum (`packages/sp/fields/types.ts`).
/// Deserialized from the integer `FieldTypeKind` value SharePoint returns.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(try_from = "i64", into = "i64")]
#[repr(i64)]
pub enum FieldType {
    Invalid = 0,
    Integer = 1,
    Text = 2,
    Note = 3,
    DateTime = 4,
    Counter = 5,
    Choice = 6,
    Lookup = 7,
    Boolean = 8,
    Number = 9,
    Currency = 10,
    Url = 11,
    Computed = 12,
    Threading = 13,
    Guid = 14,
    MultiChoice = 15,
    GridChoice = 16,
    Calculated = 17,
    File = 18,
    Attachments = 19,
    User = 20,
    Recurrence = 21,
    CrossProjectLink = 22,
    ModStat = 23,
    Error = 24,
    ContentTypeId = 25,
    PageSeparator = 26,
    ThreadIndex = 27,
    WorkflowStatus = 28,
    AllDayEvent = 29,
    WorkflowEventType = 30,
    Location = 33,
    Image = 34,
}

impl std::convert::TryFrom<i64> for FieldType {
    type Error = String;
    fn try_from(v: i64) -> std::result::Result<Self, String> {
        match v {
            0 => Ok(Self::Invalid),
            1 => Ok(Self::Integer),
            2 => Ok(Self::Text),
            3 => Ok(Self::Note),
            4 => Ok(Self::DateTime),
            5 => Ok(Self::Counter),
            6 => Ok(Self::Choice),
            7 => Ok(Self::Lookup),
            8 => Ok(Self::Boolean),
            9 => Ok(Self::Number),
            10 => Ok(Self::Currency),
            11 => Ok(Self::Url),
            12 => Ok(Self::Computed),
            13 => Ok(Self::Threading),
            14 => Ok(Self::Guid),
            15 => Ok(Self::MultiChoice),
            16 => Ok(Self::GridChoice),
            17 => Ok(Self::Calculated),
            18 => Ok(Self::File),
            19 => Ok(Self::Attachments),
            20 => Ok(Self::User),
            21 => Ok(Self::Recurrence),
            22 => Ok(Self::CrossProjectLink),
            23 => Ok(Self::ModStat),
            24 => Ok(Self::Error),
            25 => Ok(Self::ContentTypeId),
            26 => Ok(Self::PageSeparator),
            27 => Ok(Self::ThreadIndex),
            28 => Ok(Self::WorkflowStatus),
            29 => Ok(Self::AllDayEvent),
            30 => Ok(Self::WorkflowEventType),
            33 => Ok(Self::Location),
            34 => Ok(Self::Image),
            _ => Err(format!("unknown FieldType value: {v}")),
        }
    }
}

impl From<FieldType> for i64 {
    fn from(ft: FieldType) -> i64 {
        ft as i64
    }
}

impl FieldType {
    /// Maps a `FieldType` to the OData `__metadata.type` string SharePoint
    /// expects when creating a field.
    ///
    /// Ported from PnPjs `mapFieldTypeEnumToString`
    /// (`packages/sp/fields/types.ts`).
    pub fn odata_type(self) -> &'static str {
        match self {
            FieldType::Calculated => "SP.FieldCalculated",
            FieldType::Choice => "SP.FieldChoice",
            FieldType::Computed => "SP.FieldComputed",
            FieldType::Currency => "SP.FieldCurrency",
            FieldType::DateTime => "SP.FieldDateTime",
            FieldType::GridChoice => "SP.FieldRatingScale",
            FieldType::Guid => "SP.FieldGuid",
            FieldType::Image => "SP.FieldMultiLineText",
            FieldType::Integer => "SP.FieldNumber",
            FieldType::Location => "SP.FieldLocation",
            FieldType::Lookup => "SP.FieldLookup",
            FieldType::ModStat => "SP.FieldChoice",
            FieldType::MultiChoice => "SP.FieldMultiChoice",
            FieldType::Note => "SP.FieldMultiLineText",
            FieldType::Number => "SP.FieldNumber",
            FieldType::Text => "SP.FieldText",
            FieldType::Url => "SP.FieldUrl",
            FieldType::User => "SP.FieldUser",
            FieldType::WorkflowStatus => "SP.FieldChoice",
            FieldType::WorkflowEventType => "SP.FieldNumber",
            _ => "SP.Field",
        }
    }
}

/// Calendar type for date/time fields.
///
/// Ported from PnPjs `CalendarType` enum (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[repr(i64)]
pub enum CalendarType {
    Gregorian = 1,
    Japan = 3,
    Taiwan = 4,
    Korea = 5,
    Hijri = 6,
    Thai = 7,
    Hebrew = 8,
    GregorianMEFrench = 9,
    GregorianArabic = 10,
    GregorianXLITEnglish = 11,
    GregorianXLITFrench = 12,
    KoreaJapanLunar = 14,
    ChineseLunar = 15,
    SakaEra = 16,
    UmAlQura = 23,
}

/// Date/time display format.
///
/// Ported from PnPjs `DateTimeFieldFormatType` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[repr(i64)]
pub enum DateTimeFieldFormat {
    DateOnly = 0,
    DateTime = 1,
}

/// Friendly display format for date/time fields.
///
/// Ported from PnPjs `DateTimeFieldFriendlyFormatType`
/// (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[repr(i64)]
pub enum DateTimeFieldFriendlyFormat {
    Unspecified = 0,
    Disabled = 1,
    Relative = 2,
}

/// URL field display format.
///
/// Ported from PnPjs `UrlFieldFormatType` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[repr(i64)]
pub enum UrlFieldFormat {
    Hyperlink = 0,
    Image = 1,
}

/// User field selection mode.
///
/// Ported from PnPjs `FieldUserSelectionMode` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[repr(i64)]
pub enum FieldUserSelectionMode {
    PeopleOnly = 0,
    PeopleAndGroups = 1,
}

/// Choice field display format.
///
/// Ported from PnPjs `ChoiceFieldFormatType` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[repr(i64)]
pub enum ChoiceFieldFormat {
    Dropdown = 0,
    RadioButtons = 1,
}

/// Options controlling field creation behavior.
///
/// Ported from PnPjs `AddFieldOptions` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AddFieldOptions(pub u32);

impl AddFieldOptions {
    pub const DEFAULT: Self = Self(0);
    pub const ADD_TO_DEFAULT_CONTENT_TYPE: Self = Self(1);
    pub const ADD_TO_NO_CONTENT_TYPE: Self = Self(2);
    pub const ADD_TO_ALL_CONTENT_TYPES: Self = Self(4);
    pub const ADD_FIELD_INTERNAL_NAME_HINT: Self = Self(8);
    pub const ADD_FIELD_TO_DEFAULT_VIEW: Self = Self(16);
    pub const ADD_FIELD_CHECK_DISPLAY_NAME: Self = Self(32);
}

/// Common properties for field creation requests.
///
/// Ported from PnPjs `IFieldCreationProperties` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct FieldCreationProperties {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub group: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hidden: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub required: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub indexed: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub enforce_unique_values: Option<bool>,
    #[serde(rename = "DefaultFormula", skip_serializing_if = "Option::is_none")]
    pub default_formula: Option<String>,
    #[serde(rename = "ValidationFormula", skip_serializing_if = "Option::is_none")]
    pub validation_formula: Option<String>,
    #[serde(rename = "ValidationMessage", skip_serializing_if = "Option::is_none")]
    pub validation_message: Option<String>,
}

/// Properties specific to text field creation.
///
/// Ported from PnPjs `AddTextProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddTextProps {
    #[serde(rename = "MaxLength", skip_serializing_if = "Option::is_none")]
    pub max_length: Option<i64>,
}

/// Properties specific to calculated field creation.
///
/// Ported from PnPjs `AddCalculatedProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddCalculatedProps {
    #[serde(rename = "Formula", skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    #[serde(rename = "DateFormat", skip_serializing_if = "Option::is_none")]
    pub date_format: Option<DateTimeFieldFormat>,
    #[serde(rename = "OutputType", skip_serializing_if = "Option::is_none")]
    pub output_type: Option<FieldType>,
}

/// Properties specific to date/time field creation.
///
/// Ported from PnPjs `AddDateTimeProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddDateTimeProps {
    #[serde(
        rename = "DateTimeCalendarType",
        skip_serializing_if = "Option::is_none"
    )]
    pub calendar_type: Option<CalendarType>,
    #[serde(rename = "DisplayFormat", skip_serializing_if = "Option::is_none")]
    pub display_format: Option<DateTimeFieldFormat>,
    #[serde(
        rename = "FriendlyDisplayFormat",
        skip_serializing_if = "Option::is_none"
    )]
    pub friendly_display_format: Option<DateTimeFieldFriendlyFormat>,
}

/// Properties specific to number field creation.
///
/// Ported from PnPjs `AddNumberProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddNumberProps {
    #[serde(rename = "MinimumValue", skip_serializing_if = "Option::is_none")]
    pub minimum_value: Option<f64>,
    #[serde(rename = "MaximumValue", skip_serializing_if = "Option::is_none")]
    pub maximum_value: Option<f64>,
}

/// Properties specific to currency field creation.
///
/// Ported from PnPjs `AddCurrencyProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddCurrencyProps {
    #[serde(rename = "MinimumValue", skip_serializing_if = "Option::is_none")]
    pub minimum_value: Option<f64>,
    #[serde(rename = "MaximumValue", skip_serializing_if = "Option::is_none")]
    pub maximum_value: Option<f64>,
    #[serde(rename = "CurrencyLocaleId", skip_serializing_if = "Option::is_none")]
    pub currency_locale_id: Option<i64>,
}

/// Properties specific to multiline text field creation.
///
/// Ported from PnPjs `AddMultilineTextProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddMultilineTextProps {
    #[serde(rename = "AllowHyperlink", skip_serializing_if = "Option::is_none")]
    pub allow_hyperlink: Option<bool>,
    #[serde(rename = "AppendOnly", skip_serializing_if = "Option::is_none")]
    pub append_only: Option<bool>,
    #[serde(rename = "NumberOfLines", skip_serializing_if = "Option::is_none")]
    pub number_of_lines: Option<i64>,
    #[serde(rename = "RestrictedMode", skip_serializing_if = "Option::is_none")]
    pub restricted_mode: Option<bool>,
    #[serde(rename = "RichText", skip_serializing_if = "Option::is_none")]
    pub rich_text: Option<bool>,
}

/// Properties specific to URL field creation.
///
/// Ported from PnPjs `AddUrlProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddUrlProps {
    #[serde(rename = "DisplayFormat", skip_serializing_if = "Option::is_none")]
    pub display_format: Option<UrlFieldFormat>,
}

/// Properties specific to user field creation.
///
/// Ported from PnPjs `AddUserProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddUserProps {
    #[serde(rename = "SelectionMode", skip_serializing_if = "Option::is_none")]
    pub selection_mode: Option<FieldUserSelectionMode>,
}

/// Properties specific to choice field creation.
///
/// Ported from PnPjs `AddChoiceProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddChoiceProps {
    pub choices: Vec<String>,
    #[serde(rename = "EditFormat", skip_serializing_if = "Option::is_none")]
    pub edit_format: Option<ChoiceFieldFormat>,
    #[serde(rename = "FillInChoice", skip_serializing_if = "Option::is_none")]
    pub fill_in_choice: Option<bool>,
}

/// Properties specific to multi-choice field creation.
///
/// Ported from PnPjs `AddMultiChoiceProps` (`packages/sp/fields/types.ts`).
#[derive(Debug, Clone, Default, Serialize)]
pub struct AddMultiChoiceProps {
    pub choices: Vec<String>,
    #[serde(rename = "FillInChoice", skip_serializing_if = "Option::is_none")]
    pub fill_in_choice: Option<bool>,
}

/// Information about a SharePoint field.
///
/// Ported from PnPjs `IFieldInfo` (`packages/sp/fields/types.ts`).
/// Uses a practical subset of the ~35-field SharePoint entity.
#[derive(Debug, Clone, Deserialize)]
pub struct FieldInfo {
    #[serde(rename = "Id")]
    pub id: String,
    #[serde(rename = "Title")]
    pub title: String,
    #[serde(rename = "InternalName")]
    pub internal_name: String,
    #[serde(rename = "FieldTypeKind")]
    pub field_type_kind: FieldType,
    #[serde(rename = "TypeAsString", default)]
    pub type_as_string: String,
    #[serde(rename = "TypeDisplayName", default)]
    pub type_display_name: String,
    #[serde(rename = "TypeShortDescription", default)]
    pub type_short_description: String,
    #[serde(rename = "Description", default)]
    pub description: String,
    #[serde(rename = "Group", default)]
    pub group: String,
    #[serde(rename = "Hidden", default)]
    pub hidden: bool,
    #[serde(rename = "Required", default)]
    pub required: bool,
    #[serde(rename = "ReadOnlyField", default)]
    pub read_only_field: bool,
    #[serde(rename = "Indexed", default)]
    pub indexed: bool,
    #[serde(rename = "EnforceUniqueValues", default)]
    pub enforce_unique_values: bool,
    #[serde(rename = "SchemaXml", default)]
    pub schema_xml: String,
    #[serde(rename = "Scope", default)]
    pub scope: String,
    #[serde(rename = "StaticName", default)]
    pub static_name: String,
    #[serde(rename = "DefaultValue", default)]
    pub default_value: Option<String>,
    #[serde(rename = "Choices", default)]
    pub choices: Option<Vec<String>>,
    #[serde(rename = "EntityPropertyName", default)]
    pub entity_property_name: String,
    #[serde(rename = "Filterable", default)]
    pub filterable: bool,
    #[serde(rename = "Sortable", default)]
    pub sortable: bool,
    #[serde(rename = "Sealed", default)]
    pub sealed: bool,
    #[serde(rename = "FromBaseType", default)]
    pub from_base_type: bool,
}
