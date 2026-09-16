use forge_m365_core::{registered_operations, Client, Error, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::Deserialize;

const DEFAULTS_FILE_NAME: &str = "client_LocationBasedDefaults.html";

/// A parsed column default as read back from the defaults file: a field's
/// internal name, the (unencoded) server-relative folder path it applies to,
/// and its raw stored value. Ported from PnPjs `IFieldDefault`
/// (`packages/sp/column-defaults/types.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldDefault {
    pub name: String,
    pub path: String,
    pub value: String,
}

/// A single taxonomy term reference, as PnPjs's `AllowedDefaultColumnValues`
/// taxonomy shape (`{ wssId, termName, termId }`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TaxonomyTermRef {
    pub wss_id: String,
    pub term_name: String,
    pub term_id: String,
}

/// The value to set for one field default, in whatever shape its field type
/// requires. Ported from the `AllowedDefaultColumnValues` union and the
/// `TypeAsString` switch in PnPjs `_List.setDefaultColumnValues`
/// (`packages/sp/column-defaults/list.ts`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ColumnDefaultValue {
    /// `Boolean` / `Currency` / `Text` / `DateTime` / `Number` / `Choice` /
    /// `User` / a bare `MultiChoice` or `UserMulti` value.
    Single(String),
    /// `MultiChoice` (joined with `;`) or `UserMulti` (joined with `;#`).
    Multi(Vec<String>),
    /// `Taxonomy` / `TaxonomyFieldType`.
    Taxonomy(TaxonomyTermRef),
    /// `TaxonomyMulti` / `TaxonomyFieldTypeMulti`.
    TaxonomyMulti(Vec<TaxonomyTermRef>),
    /// An already-formatted value, written back unchanged regardless of
    /// field type. Used internally to round-trip existing defaults (as
    /// returned by [`get_default_column_values`], which always yields plain
    /// strings) through a merge without needing to reverse-parse a taxonomy
    /// value back into structured form.
    Raw(String),
}

/// A field default to write, paired with the server-relative folder path it
/// applies to. Ported from PnPjs `IFieldDefault` used as input to
/// `_List.setDefaultColumnValues`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldDefaultInput {
    pub name: String,
    pub path: String,
    pub value: ColumnDefaultValue,
}

/// A field default to write against a single, already-known folder. Ported
/// from PnPjs `IFieldDefaultProps`, used as input to
/// `_Folder.setDefaultColumnValues` (`packages/sp/column-defaults/folder.ts`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldDefaultProps {
    pub name: String,
    pub value: ColumnDefaultValue,
}

#[derive(Debug, Deserialize)]
struct ResourcePath {
    #[serde(rename = "DecodedUrl")]
    decoded_url: String,
}

#[derive(Debug, Deserialize)]
struct RootFolderResponse {
    #[serde(rename = "ServerRelativePath")]
    server_relative_path: ResourcePath,
}

#[derive(Debug, Deserialize)]
struct ParentWebUrl {
    #[serde(rename = "Url")]
    url: String,
}

#[derive(Debug, Deserialize)]
struct ParentWebResponse {
    #[serde(rename = "ParentWeb")]
    parent_web: ParentWebUrl,
}

#[derive(Debug, Deserialize)]
struct FieldTypeInfo {
    #[serde(rename = "InternalName")]
    internal_name: String,
    #[serde(rename = "TypeAsString")]
    type_as_string: String,
}

/// Doubles embedded `'` and percent-encodes for use inside an OData
/// `decodedUrl='...'` path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`).
fn encode_path_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn list_url(site_url: &str, list_title: &str) -> String {
    format!(
        "{}/_api/web/lists/getbytitle('{}')",
        site_url.trim_end_matches('/'),
        encode_path_segment(list_title)
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-column-defaults linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.column_defaults.get", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_op() {}

#[pnp_operation(id = "sp.column_defaults.set", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn set_op() {}

fn is_not_found(err: &Error) -> bool {
    matches!(err, Error::Surface(_, msg) if msg.starts_with("HTTP 404"))
}

fn format_default_value(field_type: &str, value: &ColumnDefaultValue) -> Result<String> {
    use ColumnDefaultValue::*;

    let single_only = || -> Result<&str> {
        match value {
            Single(v) => Ok(v.as_str()),
            _ => Err(Error::Unsupported(
                "column-defaults: this field type does not support multiple values",
            )),
        }
    };

    if let Raw(v) = value {
        return Ok(v.clone());
    }

    Ok(match field_type {
        "Boolean" | "Currency" | "Text" | "DateTime" | "Number" | "Choice" | "User" => {
            single_only()?.to_string()
        }
        "MultiChoice" => match value {
            Single(v) => v.clone(),
            Multi(vs) => vs.join(";"),
            _ => {
                return Err(Error::Unsupported(
                    "column-defaults: MultiChoice needs a Single or Multi value",
                ))
            }
        },
        "UserMulti" => match value {
            Single(v) => v.clone(),
            Multi(vs) => vs.join(";#"),
            _ => {
                return Err(Error::Unsupported(
                    "column-defaults: UserMulti needs a Single or Multi value",
                ))
            }
        },
        "Taxonomy" | "TaxonomyFieldType" => match value {
            Taxonomy(t) => format!("{};#{}|{}", t.wss_id, t.term_name, t.term_id),
            _ => {
                return Err(Error::Unsupported(
                    "column-defaults: Taxonomy field needs a single Taxonomy value",
                ))
            }
        },
        "TaxonomyMulti" | "TaxonomyFieldTypeMulti" => {
            let terms: &[TaxonomyTermRef] = match value {
                Taxonomy(t) => std::slice::from_ref(t),
                TaxonomyMulti(ts) => ts.as_slice(),
                _ => return Err(Error::Unsupported(
                    "column-defaults: TaxonomyMulti field needs a Taxonomy or TaxonomyMulti value",
                )),
            };
            terms
                .iter()
                .map(|t| format!("{};#{}|{}", t.wss_id, t.term_name, t.term_id))
                .collect::<Vec<_>>()
                .join(";#")
        }
        // PnPjs's switch has no default arm and falls through to an empty
        // string for any TypeAsString it doesn't recognize.
        _ => String::new(),
    })
}

/// Extracts non-nested `<tag ...>...</tag>` blocks. The defaults file's XML
/// is always flat (no nested `<a>` or `<DefaultValue>` elements), so a plain
/// scan is faithful to PnPjs's regex-based parse without pulling in an XML
/// crate dependency.
fn extract_tags<'a>(xml: &'a str, tag: &str) -> Vec<&'a str> {
    let open_prefix = format!("<{tag}");
    let close = format!("</{tag}>");
    let mut out = Vec::new();
    let mut rest = xml;
    while let Some(start) = rest.find(&open_prefix) {
        let after_open = &rest[start..];
        if let Some(end) = after_open.find(&close) {
            let end = end + close.len();
            out.push(&after_open[..end]);
            rest = &after_open[end..];
        } else {
            break;
        }
    }
    out
}

fn attr_value<'a>(tag: &'a str, attr: &str) -> Option<&'a str> {
    let needle = format!("{attr}=\"");
    let start = tag.find(&needle)? + needle.len();
    let end = tag[start..].find('"')? + start;
    Some(&tag[start..end])
}

fn tag_inner_text(tag: &str, tag_name: &str) -> String {
    let open_end = tag.find('>').map(|i| i + 1).unwrap_or(0);
    let close = format!("</{tag_name}>");
    let close_start = tag.rfind(&close).unwrap_or(tag.len());
    tag[open_end..close_start].to_string()
}

fn decode_percent(s: &str) -> String {
    percent_encoding::percent_decode_str(s)
        .decode_utf8_lossy()
        .into_owned()
}

/// Parses the `<MetadataDefaults><a href="...">/<DefaultValue FieldName="...">
/// ...</DefaultValue></a></MetadataDefaults>` document PnPjs writes/reads at
/// `Forms/client_LocationBasedDefaults.html`. Ported from PnPjs's
/// `_List.getDefaultColumnValues` regex parse (`packages/sp/column-defaults/list.ts`).
fn parse_defaults_xml(xml: &str) -> Vec<FieldDefault> {
    let mut out = Vec::new();
    for a_tag in extract_tags(xml, "a") {
        let Some(href) = attr_value(a_tag, "href") else {
            continue;
        };
        let path = decode_percent(href);
        for dv_tag in extract_tags(a_tag, "DefaultValue") {
            let Some(field_name) = attr_value(dv_tag, "FieldName") else {
                continue;
            };
            let value = tag_inner_text(dv_tag, "DefaultValue");
            out.push(FieldDefault {
                name: field_name.to_string(),
                path: path.clone(),
                value,
            });
        }
    }
    out
}

/// Serializes defaults back into the `<MetadataDefaults>` XML document,
/// grouping by folder path the same way PnPjs's `_List.setDefaultColumnValues`
/// does. `field_types` must contain every field referenced by `defaults`
/// (internal name -> `TypeAsString`), e.g. from the list's `fields` endpoint.
///
/// Ported from PnPjs `_List.setDefaultColumnValues`
/// (`packages/sp/column-defaults/list.ts`).
pub fn build_defaults_xml(
    defaults: &[FieldDefaultInput],
    field_types: &[(String, String)],
) -> Result<String> {
    let mut by_path: Vec<(String, Vec<&FieldDefaultInput>)> = Vec::new();
    for d in defaults {
        if let Some(group) = by_path.iter_mut().find(|(p, _)| p == &d.path) {
            group.1.push(d);
        } else {
            by_path.push((d.path.clone(), vec![d]));
        }
    }

    let mut path_defaults = Vec::with_capacity(by_path.len());
    for (path, fields) in &by_path {
        let mut tags = String::new();
        for f in fields {
            let field_type = field_types
                .iter()
                .find(|(name, _)| name == &f.name)
                .map(|(_, ty)| ty.as_str())
                .ok_or(Error::Unsupported(
                    "column-defaults: field does not exist in the list",
                ))?;
            let value = format_default_value(field_type, &f.value)?;
            tags.push_str(&format!(
                "<DefaultValue FieldName=\"{}\">{}</DefaultValue>",
                f.name, value
            ));
        }
        let href = path.replace(' ', "%20");
        path_defaults.push(format!("<a href=\"{href}\">{tags}</a>"));
    }

    Ok(format!(
        "<MetadataDefaults>{}</MetadataDefaults>",
        path_defaults.join("")
    ))
}

/// Fetches `InternalName`/`TypeAsString` for every non-hidden field on the
/// list, needed to format default values correctly (see
/// [`build_defaults_xml`]).
///
/// Ported from the `fieldDefs` lookup in PnPjs `_List.setDefaultColumnValues`
/// (`packages/sp/column-defaults/list.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
async fn get_field_types(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
) -> Result<Vec<(String, String)>> {
    let url = format!(
        "{}/fields?$select=InternalName,TypeAsString&$filter=Hidden ne true",
        list_url(site_url, list_title)
    );
    let bytes = client
        .run_ladder(entry("sp.column_defaults.set"), "GET", &url, &[], None)
        .await?;
    let fields: Vec<FieldTypeInfo> = serde_json::from_slice(&bytes)?;
    Ok(fields
        .into_iter()
        .map(|f| (f.internal_name, f.type_as_string))
        .collect())
}

async fn root_folder_path(client: &Client<'_>, site_url: &str, list_title: &str) -> Result<String> {
    let url = format!(
        "{}/rootfolder?$select=ServerRelativePath",
        list_url(site_url, list_title)
    );
    let bytes = client
        .run_ladder(entry("sp.column_defaults.get"), "GET", &url, &[], None)
        .await?;
    let resp: RootFolderResponse = serde_json::from_slice(&bytes)?;
    Ok(resp.server_relative_path.decoded_url)
}

async fn parent_web_url(client: &Client<'_>, site_url: &str, list_title: &str) -> Result<String> {
    let url = format!(
        "{}?$select=ParentWeb/Url&$expand=ParentWeb",
        list_url(site_url, list_title)
    );
    let bytes = client
        .run_ladder(entry("sp.column_defaults.get"), "GET", &url, &[], None)
        .await?;
    let resp: ParentWebResponse = serde_json::from_slice(&bytes)?;
    Ok(resp.parent_web.url)
}

fn defaults_file_path(root_folder_decoded_url: &str) -> String {
    format!(
        "/{}/Forms/{DEFAULTS_FILE_NAME}",
        root_folder_decoded_url.trim_start_matches('/')
    )
}

/// Reads every column default stored on a document library, across all
/// folders. Returns an empty list if the defaults file has never been
/// created (a 404 from SharePoint), matching PnPjs's behavior.
///
/// Ported from PnPjs `_List.getDefaultColumnValues`
/// (`packages/sp/column-defaults/list.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_default_column_values(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
) -> Result<Vec<FieldDefault>> {
    let root_folder = root_folder_path(client, site_url, list_title).await?;
    let web_url = parent_web_url(client, site_url, list_title).await?;
    let path = defaults_file_path(&root_folder);
    let url = format!(
        "{}/_api/web/getFileByServerRelativePath(decodedUrl='{}')/$value",
        web_url.trim_end_matches('/'),
        encode_path_segment(&path)
    );

    let bytes = match client
        .run_ladder(entry("sp.column_defaults.get"), "GET", &url, &[], None)
        .await
    {
        Ok(b) => b,
        Err(e) if is_not_found(&e) => return Ok(Vec::new()),
        Err(e) => return Err(e),
    };
    let xml = String::from_utf8_lossy(&bytes);
    Ok(parse_defaults_xml(&xml))
}

/// Replaces *all* column defaults on a document library with `defaults`
/// (across every folder — this is not a per-folder merge; see
/// [`set_folder_default_column_values`] for that). Looks up field types
/// itself to format values correctly, then writes the defaults file via
/// `Files.add(overwrite=true)`.
///
/// Note: PnPjs also ensures a `LocationBasedMetadataDefaultsReceiver
/// ItemAdded` event receiver exists on the list after writing the file (so
/// SharePoint actually applies the defaults to newly added items). That
/// receiver-registration step needs `_api/web/lists(...)/eventreceivers` and
/// `Lists.add`-shaped bodies this crate doesn't otherwise touch; it is
/// deferred here — call it out to the caller in your own migration/setup code
/// until a `sp-event-receivers`-shaped crate exists to host it properly.
///
/// Ported from PnPjs `_List.setDefaultColumnValues`
/// (`packages/sp/column-defaults/list.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn set_default_column_values(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    defaults: &[FieldDefaultInput],
) -> Result<()> {
    let field_types = get_field_types(client, site_url, list_title).await?;
    let xml = build_defaults_xml(defaults, &field_types)?;

    let root_folder = root_folder_path(client, site_url, list_title).await?;
    let web_url = parent_web_url(client, site_url, list_title).await?;
    let forms_path = format!("/{}/Forms", root_folder.trim_start_matches('/'));
    let url = format!(
        "{}/_api/web/getFolderByServerRelativePath(decodedUrl='{}')/files/add(overwrite=true,url='{DEFAULTS_FILE_NAME}')",
        web_url.trim_end_matches('/'),
        encode_path_segment(&forms_path)
    );

    client
        .run_ladder(
            entry("sp.column_defaults.set"),
            "POST",
            &url,
            &[("Content-Type", "text/plain")],
            Some(xml.as_bytes()),
        )
        .await?;
    Ok(())
}

/// Gets the column defaults that apply to one specific folder (identified by
/// its server-relative path), by fetching the whole list's defaults and
/// filtering to that path.
///
/// PnPjs's `_Folder.getDefaultColumnValues` instead discovers its owning list
/// by reading the folder's `vti_x005f_listname` property and walking up to a
/// `Web`/`List` object; this crate's functions take `list_title` explicitly
/// instead (as every other crate here does), so that discovery step doesn't
/// apply — this is a deliberate adaptation to this SDK's calling convention,
/// not a gap in the port.
///
/// Ported from PnPjs `_Folder.getDefaultColumnValues`
/// (`packages/sp/column-defaults/folder.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_folder_default_column_values(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    folder_server_relative_path: &str,
) -> Result<Vec<FieldDefault>> {
    let all = get_default_column_values(client, site_url, list_title).await?;
    Ok(all
        .into_iter()
        .filter(|v| v.path.eq_ignore_ascii_case(folder_server_relative_path))
        .collect())
}

/// Sets the column defaults for one folder. When `merge` is `true` (PnPjs's
/// default), existing defaults for other folders and other fields on this
/// folder are preserved; when `false`, every existing default for this
/// folder is dropped first.
///
/// See [`get_folder_default_column_values`] for why this crate takes
/// `list_title` explicitly instead of resolving it from a folder object.
///
/// Ported from PnPjs `_Folder.setDefaultColumnValues`
/// (`packages/sp/column-defaults/folder.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn set_folder_default_column_values(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    folder_server_relative_path: &str,
    field_defaults: &[FieldDefaultProps],
    merge: bool,
) -> Result<()> {
    let existing = get_default_column_values(client, site_url, list_title).await?;
    let mut filtered: Vec<FieldDefaultInput> = existing
        .into_iter()
        .filter(|f| merge || !f.path.eq_ignore_ascii_case(folder_server_relative_path))
        .map(|f| FieldDefaultInput {
            name: f.name,
            path: f.path,
            value: ColumnDefaultValue::Raw(f.value),
        })
        .collect();

    for d in field_defaults {
        if let Some(existing) = filtered
            .iter_mut()
            .find(|f| f.name == d.name && f.path.eq_ignore_ascii_case(folder_server_relative_path))
        {
            existing.value = d.value.clone();
        } else {
            filtered.push(FieldDefaultInput {
                name: d.name.clone(),
                path: folder_server_relative_path.to_string(),
                value: d.value.clone(),
            });
        }
    }

    set_default_column_values(client, site_url, list_title, &filtered).await
}

/// Clears every column default for one folder.
///
/// Ported from PnPjs `_Folder.clearDefaultColumnValues`
/// (`packages/sp/column-defaults/folder.ts`) @ pnpjs `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn clear_folder_default_column_values(
    client: &Client<'_>,
    site_url: &str,
    list_title: &str,
    folder_server_relative_path: &str,
) -> Result<()> {
    set_folder_default_column_values(
        client,
        site_url,
        list_title,
        folder_server_relative_path,
        &[],
        false,
    )
    .await
}
