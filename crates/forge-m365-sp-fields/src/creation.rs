use forge_m365_core::{registered_operations, Client, OperationEntry, Result};
use serde::Serialize;
use serde_json::{Map, Value};

use crate::types::FieldType;

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-fields linked in?")
        })
}

fn fields_url(site_url: &str) -> String {
    format!("{}/_api/web/fields", site_url.trim_end_matches('/'))
}

/// Builds the JSON body for the `add` endpoint: merges the metadata type
/// with the caller-supplied creation properties map.
fn build_add_body(
    title: &str,
    field_type: FieldType,
    props: &Map<String, Value>,
) -> Result<Vec<u8>> {
    let mut body = props.clone();
    body.insert("Title".to_string(), Value::String(title.to_string()));
    body.insert(
        "FieldTypeKind".to_string(),
        Value::Number((field_type as i64).into()),
    );
    let mut metadata = Map::new();
    metadata.insert(
        "type".to_string(),
        Value::String(field_type.odata_type().to_string()),
    );
    body.insert("__metadata".to_string(), Value::Object(metadata));
    Ok(serde_json::to_vec(&Value::Object(body))?)
}

/// Generic field creation — posts to `_api/web/fields` with the given
/// `FieldTypeKind` and optional creation properties.
///
/// Ported from PnPjs `_Fields.add` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_field(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    field_type: FieldType,
    properties: &Map<String, Value>,
) -> Result<Value> {
    let url = fields_url(site_url);
    let payload = build_add_body(title, field_type, properties)?;
    let bytes = client
        .run_ladder(entry("sp.fields.add"), "POST", &url, &[], Some(&payload))
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Creates a field from raw XML schema.
///
/// Ported from PnPjs `_Fields.createFieldAsXml`
/// (`packages/sp/fields/types.ts`) @ pnpjs `8ee2375d`.
pub async fn create_field_as_xml(
    client: &Client<'_>,
    site_url: &str,
    schema_xml: &str,
) -> Result<Value> {
    let url = format!("{}/createfieldasxml", fields_url(site_url));
    let mut body = Map::new();
    let mut params = Map::new();
    params.insert(
        "SchemaXml".to_string(),
        Value::String(schema_xml.to_string()),
    );
    body.insert("parameters".to_string(), Value::Object(params));
    let payload = serde_json::to_vec(&Value::Object(body))?;
    let bytes = client
        .run_ladder(
            entry("sp.fields.create_field_as_xml"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Adds a text field (FieldTypeKind = 2). Default MaxLength: 255.
///
/// Ported from PnPjs `_Fields.addText` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_text(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    max_length: Option<i64>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    props.insert(
        "MaxLength".to_string(),
        Value::Number(max_length.unwrap_or(255).into()),
    );
    add_field(client, site_url, title, FieldType::Text, &props).await
}

/// Adds a calculated field (FieldTypeKind = 17). Default OutputType: Text.
///
/// Ported from PnPjs `_Fields.addCalculated` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_calculated(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    formula: Option<&str>,
    output_type: Option<FieldType>,
    date_format: Option<i64>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    if let Some(f) = formula {
        props.insert("Formula".to_string(), Value::String(f.to_string()));
    }
    props.insert(
        "OutputType".to_string(),
        Value::Number((output_type.unwrap_or(FieldType::Text) as i64).into()),
    );
    if let Some(df) = date_format {
        props.insert("DateFormat".to_string(), Value::Number(df.into()));
    }
    add_field(client, site_url, title, FieldType::Calculated, &props).await
}

/// Adds a date/time field (FieldTypeKind = 4). Defaults: Gregorian, DateOnly, Unspecified.
///
/// Ported from PnPjs `_Fields.addDateTime` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_date_time(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    calendar_type: Option<i64>,
    display_format: Option<i64>,
    friendly_format: Option<i64>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    props.insert(
        "DateTimeCalendarType".to_string(),
        Value::Number(calendar_type.unwrap_or(1).into()),
    );
    props.insert(
        "DisplayFormat".to_string(),
        Value::Number(display_format.unwrap_or(0).into()),
    );
    props.insert(
        "FriendlyDisplayFormat".to_string(),
        Value::Number(friendly_format.unwrap_or(0).into()),
    );
    add_field(client, site_url, title, FieldType::DateTime, &props).await
}

/// Adds a number field (FieldTypeKind = 9).
///
/// Ported from PnPjs `_Fields.addNumber` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_number(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    min_value: Option<f64>,
    max_value: Option<f64>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    if let Some(min) = min_value {
        props.insert("MinimumValue".to_string(), Value::from(min));
    }
    if let Some(max) = max_value {
        props.insert("MaximumValue".to_string(), Value::from(max));
    }
    add_field(client, site_url, title, FieldType::Number, &props).await
}

/// Adds a currency field (FieldTypeKind = 10). Default CurrencyLocaleId: 1033.
///
/// Ported from PnPjs `_Fields.addCurrency` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_currency(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    min_value: Option<f64>,
    max_value: Option<f64>,
    currency_locale_id: Option<i64>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    if let Some(min) = min_value {
        props.insert("MinimumValue".to_string(), Value::from(min));
    }
    if let Some(max) = max_value {
        props.insert("MaximumValue".to_string(), Value::from(max));
    }
    props.insert(
        "CurrencyLocaleId".to_string(),
        Value::Number(currency_locale_id.unwrap_or(1033).into()),
    );
    add_field(client, site_url, title, FieldType::Currency, &props).await
}

/// Adds a multiline text field (FieldTypeKind = 3). Defaults: RichText, AllowHyperlink, 6 lines.
///
/// Ported from PnPjs `_Fields.addMultilineText` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
#[allow(clippy::too_many_arguments)]
pub async fn add_multiline_text(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    allow_hyperlink: Option<bool>,
    append_only: Option<bool>,
    number_of_lines: Option<i64>,
    restricted_mode: Option<bool>,
    rich_text: Option<bool>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    props.insert(
        "AllowHyperlink".to_string(),
        Value::Bool(allow_hyperlink.unwrap_or(true)),
    );
    props.insert(
        "AppendOnly".to_string(),
        Value::Bool(append_only.unwrap_or(false)),
    );
    props.insert(
        "NumberOfLines".to_string(),
        Value::Number(number_of_lines.unwrap_or(6).into()),
    );
    props.insert(
        "RestrictedMode".to_string(),
        Value::Bool(restricted_mode.unwrap_or(false)),
    );
    props.insert(
        "RichText".to_string(),
        Value::Bool(rich_text.unwrap_or(true)),
    );
    add_field(client, site_url, title, FieldType::Note, &props).await
}

/// Adds a URL field (FieldTypeKind = 11). Default: Hyperlink format.
///
/// Ported from PnPjs `_Fields.addUrl` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_url(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    display_format: Option<i64>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    props.insert(
        "DisplayFormat".to_string(),
        Value::Number(display_format.unwrap_or(0).into()),
    );
    add_field(client, site_url, title, FieldType::Url, &props).await
}

/// Adds a user field (FieldTypeKind = 20). Default: PeopleAndGroups.
///
/// Ported from PnPjs `_Fields.addUser` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_user(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    selection_mode: Option<i64>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    props.insert(
        "SelectionMode".to_string(),
        Value::Number(selection_mode.unwrap_or(1).into()),
    );
    add_field(client, site_url, title, FieldType::User, &props).await
}

/// Adds a lookup field (FieldTypeKind = 7).
///
/// Ported from PnPjs `_Fields.addLookup` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_lookup(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    lookup_list_id: &str,
    lookup_web_id: &str,
    lookup_field_name: &str,
) -> Result<Value> {
    let url = format!("{}/AddField", fields_url(site_url));
    let mut params = Map::new();
    params.insert("Title".to_string(), Value::String(title.to_string()));
    params.insert(
        "FieldTypeKind".to_string(),
        Value::Number((FieldType::Lookup as i64).into()),
    );
    params.insert(
        "LookupListId".to_string(),
        Value::String(lookup_list_id.to_string()),
    );
    params.insert(
        "LookupWebId".to_string(),
        Value::String(lookup_web_id.to_string()),
    );
    params.insert(
        "LookupFieldName".to_string(),
        Value::String(lookup_field_name.to_string()),
    );
    let mut body = Map::new();
    body.insert("parameters".to_string(), Value::Object(params));
    let payload = serde_json::to_vec(&Value::Object(body))?;
    let bytes = client
        .run_ladder(
            entry("sp.fields.add_lookup"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Adds a choice field (FieldTypeKind = 6). Choices are wrapped in `{"results": [...]}`.
///
/// Ported from PnPjs `_Fields.addChoice` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_choice(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    choices: &[String],
    edit_format: Option<i64>,
    fill_in_choice: Option<bool>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    let results: Vec<Value> = choices.iter().map(|c| Value::String(c.clone())).collect();
    let mut wrapped = Map::new();
    wrapped.insert("results".to_string(), Value::Array(results));
    props.insert("Choices".to_string(), Value::Object(wrapped));
    if let Some(ef) = edit_format {
        props.insert("EditFormat".to_string(), Value::Number(ef.into()));
    }
    if let Some(fi) = fill_in_choice {
        props.insert("FillInChoice".to_string(), Value::Bool(fi));
    }
    add_field(client, site_url, title, FieldType::Choice, &props).await
}

/// Adds a multi-choice field (FieldTypeKind = 15). Choices are wrapped in `{"results": [...]}`.
///
/// Ported from PnPjs `_Fields.addMultiChoice` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_multi_choice(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
    choices: &[String],
    fill_in_choice: Option<bool>,
) -> Result<Value> {
    let mut props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    let results: Vec<Value> = choices.iter().map(|c| Value::String(c.clone())).collect();
    let mut wrapped = Map::new();
    wrapped.insert("results".to_string(), Value::Array(results));
    props.insert("Choices".to_string(), Value::Object(wrapped));
    if let Some(fi) = fill_in_choice {
        props.insert("FillInChoice".to_string(), Value::Bool(fi));
    }
    add_field(client, site_url, title, FieldType::MultiChoice, &props).await
}

/// Adds a boolean field (FieldTypeKind = 8).
///
/// Ported from PnPjs `_Fields.addBoolean` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_boolean(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
) -> Result<Value> {
    let props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    add_field(client, site_url, title, FieldType::Boolean, &props).await
}

/// Adds a location field (FieldTypeKind = 33).
///
/// Ported from PnPjs `_Fields.addLocation` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_location(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
) -> Result<Value> {
    let props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    add_field(client, site_url, title, FieldType::Location, &props).await
}

/// Adds an image field (FieldTypeKind = 34).
///
/// Ported from PnPjs `_Fields.addImageField` (`packages/sp/fields/types.ts`)
/// @ pnpjs `8ee2375d`.
pub async fn add_image_field(
    client: &Client<'_>,
    site_url: &str,
    title: &str,
    common: &impl Serialize,
) -> Result<Value> {
    let props: Map<String, Value> = serde_json::from_slice(&serde_json::to_vec(common)?)?;
    add_field(client, site_url, title, FieldType::Image, &props).await
}

/// Creates a secondary (dependent) lookup field.
///
/// Ported from PnPjs `_Fields.addDependentLookupField`
/// (`packages/sp/fields/types.ts`) @ pnpjs `8ee2375d`.
pub async fn add_dependent_lookup_field(
    client: &Client<'_>,
    site_url: &str,
    display_name: &str,
    primary_lookup_field_id: &str,
    show_field: &str,
) -> Result<Value> {
    let url = format!(
        "{}/adddependentlookupfield(displayName='{}', primarylookupfieldid='{}', showfield='{}')",
        fields_url(site_url),
        crate::encode_segment(display_name),
        primary_lookup_field_id,
        show_field
    );
    let bytes = client
        .run_ladder(
            entry("sp.fields.add_dependent_lookup"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
