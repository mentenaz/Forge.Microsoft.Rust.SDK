use serde::Deserialize;

/// `IRegionalSettingsInfo` — the web `regionalsettings` service payload
/// (`packages/sp/regional-settings/types.ts:102`). All fields are PascalCase
/// per the OData v3 response.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct RegionalSettingsInfo {
    #[serde(rename = "AdjustHijriDays", default)]
    pub adjust_hijri_days: i64,
    #[serde(rename = "AlternateCalendarType", default)]
    pub alternate_calendar_type: i64,
    #[serde(rename = "AM", default)]
    pub am: String,
    #[serde(rename = "CalendarType", default)]
    pub calendar_type: i64,
    #[serde(rename = "Collation", default)]
    pub collation: i64,
    #[serde(rename = "CollationLCID", default)]
    pub collation_lcid: i64,
    #[serde(rename = "DateFormat", default)]
    pub date_format: i64,
    #[serde(rename = "DateSeparator", default)]
    pub date_separator: String,
    #[serde(rename = "DecimalSeparator", default)]
    pub decimal_separator: String,
    #[serde(rename = "DigitGrouping", default)]
    pub digit_grouping: String,
    #[serde(rename = "FirstDayOfWeek", default)]
    pub first_day_of_week: i64,
    #[serde(rename = "FirstWeekOfYear", default)]
    pub first_week_of_year: i64,
    #[serde(rename = "IsEastAsia", default)]
    pub is_east_asia: bool,
    #[serde(rename = "IsRightToLeft", default)]
    pub is_right_to_left: bool,
    #[serde(rename = "IsUIRightToLeft", default)]
    pub is_ui_right_to_left: bool,
    #[serde(rename = "ListSeparator", default)]
    pub list_separator: String,
    #[serde(rename = "LocaleId", default)]
    pub locale_id: i64,
    #[serde(rename = "NegativeSign", default)]
    pub negative_sign: String,
    #[serde(rename = "NegNumberMode", default)]
    pub neg_number_mode: i64,
    #[serde(rename = "PM", default)]
    pub pm: String,
    #[serde(rename = "PositiveSign", default)]
    pub positive_sign: String,
    #[serde(rename = "ShowWeeks", default)]
    pub show_weeks: bool,
    #[serde(rename = "ThousandSeparator", default)]
    pub thousand_separator: String,
    #[serde(rename = "Time24", default)]
    pub time_24: bool,
    #[serde(rename = "TimeMarkerPosition", default)]
    pub time_marker_position: i64,
    #[serde(rename = "TimeSeparator", default)]
    pub time_separator: String,
    #[serde(rename = "WorkDayEndHour", default)]
    pub work_day_end_hour: i64,
    #[serde(rename = "WorkDays", default)]
    pub work_days: i64,
    #[serde(rename = "WorkDayStartHour", default)]
    pub work_day_start_hour: i64,
}

/// `IInstalledLanguageInfo` (`types.ts:134`).
#[derive(Debug, Clone, Default, Deserialize)]
pub struct InstalledLanguageInfo {
    #[serde(rename = "DisplayName", default)]
    pub display_name: String,
    #[serde(rename = "LanguageTag", default)]
    pub language_tag: String,
    #[serde(rename = "Lcid", default)]
    pub lcid: i64,
}

/// `ITimeZoneInfo` (`types.ts:140`).
#[derive(Debug, Clone, Default, Deserialize)]
pub struct TimeZoneInfo {
    #[serde(rename = "Description", default)]
    pub description: String,
    #[serde(rename = "Id", default)]
    pub id: i64,
    #[serde(rename = "Information", default)]
    pub information: Option<TimeZoneInformation>,
}

#[derive(Debug, Clone, Default, Deserialize)]
pub struct TimeZoneInformation {
    #[serde(rename = "Bias", default)]
    pub bias: i64,
    #[serde(rename = "DaylightBias", default)]
    pub daylight_bias: i64,
    #[serde(rename = "StandardBias", default)]
    pub standard_bias: i64,
}
