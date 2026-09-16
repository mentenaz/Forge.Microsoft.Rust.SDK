//! SharePoint modern/client-side page **canvas** data model, JSON codec, and
//! read operations.
//!
//! Phase 1 (offline): the data model and the exact JSON (de)serialization
//! SharePoint uses for the `CanvasContent1` / `LayoutWebpartsContent`
//! list-item fields, plus the tree-merge algorithm PnPjs's `_ClientsidePage`
//! uses to turn a flat control array into a sections → columns → controls
//! tree (and back). This part does no network I/O, which is why its tests
//! assert exact byte-for-byte JSON shape rather than only wiring, unlike
//! most of this workspace.
//!
//! Phase 2 (this module's [`get_client_side_page`], [`load_client_side_page`],
//! [`get_page_item_ref`], [`get_client_side_web_parts`]): read-only network
//! operations built on top of the Phase 1 model, using `forge-m365-core`'s
//! `Client`/escalation-ladder machinery like every other `sp-*` crate.
//!
//! Deliberately **not** in scope here: `save()`/`saveAsTemplate()` (writing
//! a built/edited canvas back to SharePoint), page lifecycle (create,
//! delete, recycle, checkout, promote/schedule-publish, copy, share,
//! banner image, author), and resolving a web part's title/description/
//! properties from a manifest ([`WebPartData::new`] still requires the
//! caller to supply those, though [`get_client_side_web_parts`] now gets you
//! the manifests to read them from). Those are later phases — see this
//! crate's README.

use forge_m365_core::{
    registered_operations, Client, Error as CoreError, OperationEntry, Result as CoreResult,
    Surface,
};
use forge_m365_macros::pnp_operation;
use percent_encoding::{utf8_percent_encode, NON_ALPHANUMERIC};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::HashMap;

/// A position within the canvas, as PnPjs's `IPosition`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Position {
    pub zone_index: i64,
    pub section_index: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub control_index: Option<i64>,
    /// One of 0, 2, 4, 6, 8, 12 (twelfths of the section width). `None` when
    /// omitted, as PnPjs's type allows.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub section_factor: Option<i64>,
    pub layout_index: i64,
}

/// `IClientControlEmphasis` — serializes as `{}` when unset and
/// `{"zoneEmphasis": n}` when set, matching PnPjs exactly (this is not a
/// plain `Option<u8>` field on the wire, so it needs a custom impl).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Emphasis {
    pub zone_emphasis: Option<u8>,
}

impl Emphasis {
    pub fn none() -> Self {
        Self {
            zone_emphasis: None,
        }
    }

    /// Mirrors `_ClientsidePage.getEmphasisObj`: values outside `1..=3` are
    /// normalized to "no emphasis".
    pub fn from_section_value(value: u8) -> Self {
        if !(1..=3).contains(&value) {
            Self::none()
        } else {
            Self {
                zone_emphasis: Some(value),
            }
        }
    }
}

impl Serialize for Emphasis {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeMap;
        match self.zone_emphasis {
            Some(v) => {
                let mut map = serializer.serialize_map(Some(1))?;
                map.serialize_entry("zoneEmphasis", &v)?;
                map.end()
            }
            None => serializer.serialize_map(Some(0))?.end(),
        }
    }
}

impl<'de> Deserialize<'de> for Emphasis {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        #[derive(Deserialize)]
        struct Raw {
            #[serde(default, rename = "zoneEmphasis")]
            zone_emphasis: Option<u8>,
        }
        Ok(Self {
            zone_emphasis: Raw::deserialize(deserializer)?.zone_emphasis,
        })
    }
}

/// `IClientsidePageColumnData` — an empty canvas column (`controlType: 0`,
/// no `pageSettingsSlice`).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ColumnData {
    pub control_type: i64,
    pub display_mode: i64,
    pub emphasis: Emphasis,
    pub position: Position,
}

impl ColumnData {
    /// Mirrors `CanvasColumn.Default`.
    pub fn default_for(zone_index: i64, layout_index: i64, section_factor: i64) -> Self {
        Self {
            control_type: 0,
            display_mode: 2,
            emphasis: Emphasis::none(),
            position: Position {
                zone_index,
                section_index: 1,
                control_index: None,
                section_factor: Some(section_factor),
                layout_index,
            },
        }
    }
}

/// `IClientsidePageSettingsSlice` — a page-level metadata control PnPjs
/// stores as an extra entry in the same flat control array (`controlType:
/// 0`, but with a `pageSettingsSlice` key that distinguishes it from a
/// column).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageSettingsSlice {
    pub control_type: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_attributes: Option<Vec<String>>,
    pub page_settings_slice: PageSettingsSliceInner,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PageSettingsSliceInner {
    pub is_default_description: bool,
    pub is_default_thumbnail: bool,
}

/// `IClientsideTextData` (`controlType: 4`).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextData {
    pub control_type: i64,
    pub display_mode: i64,
    pub emphasis: Emphasis,
    pub id: String,
    pub position: Position,
    pub anchor_component_id: String,
    pub editor_type: String,
    pub added_from_persisted_data: bool,
    #[serde(rename = "innerHTML")]
    pub inner_html: String,
}

impl TextData {
    /// Mirrors the `ClientsideText` constructor: generates a fresh id and
    /// mirrors it into `anchorComponentId`, per PnPjs.
    pub fn new(text: impl Into<String>) -> Self {
        let id = new_guid();
        Self {
            control_type: 4,
            display_mode: 2,
            emphasis: Emphasis::none(),
            anchor_component_id: id.clone(),
            id,
            position: Position {
                zone_index: 1,
                section_index: 1,
                control_index: None,
                section_factor: Some(12),
                layout_index: 1,
            },
            editor_type: "CKEditor".to_string(),
            added_from_persisted_data: false,
            inner_html: text.into(),
        }
    }
}

/// `serverProcessedContent` on a web part, per PnPjs.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ServerProcessedContent {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_strings: Option<HashMap<String, String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub searchable_plain_texts: Option<HashMap<String, String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_sources: Option<HashMap<String, String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub links: Option<HashMap<String, String>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WebPartInnerData {
    pub id: String,
    pub instance_id: String,
    pub title: String,
    pub description: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub server_processed_content: Option<ServerProcessedContent>,
    pub data_version: String,
    pub properties: Value,
}

/// `IClientsideWebPartData` (`controlType: 3`).
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct WebPartData {
    pub control_type: i64,
    pub display_mode: i64,
    pub emphasis: Emphasis,
    pub id: String,
    pub position: Position,
    pub web_part_id: String,
    pub reserved_height: i64,
    pub reserved_width: i64,
    pub added_from_persisted_data: bool,
    pub web_part_data: WebPartInnerData,
}

impl WebPartData {
    /// Builds a web part control directly from a known component id and
    /// preconfigured title/description/properties. PnPjs's
    /// `ClientsideWebpart.fromComponentDef` derives these same fields from a
    /// `GetClientSideWebParts` manifest response instead — that network
    /// lookup is out of scope for this crate (see README); callers who
    /// already know the component id (e.g. from a prior manifest fetch, or
    /// a well-known built-in web part id) can use this directly.
    pub fn new(
        web_part_id: impl Into<String>,
        title: impl Into<String>,
        description: impl Into<String>,
        properties: Value,
    ) -> Self {
        let id = new_guid();
        Self {
            control_type: 3,
            display_mode: 2,
            emphasis: Emphasis::none(),
            id,
            position: Position {
                zone_index: 1,
                section_index: 1,
                control_index: None,
                section_factor: Some(12),
                layout_index: 1,
            },
            web_part_id: web_part_id.into(),
            reserved_height: 500,
            reserved_width: 500,
            added_from_persisted_data: false,
            web_part_data: WebPartInnerData {
                id: String::new(),
                instance_id: String::new(),
                title: title.into(),
                description: description.into(),
                server_processed_content: None,
                data_version: "1.0".to_string(),
                properties,
            },
        }
    }
}

/// A control that lives inside a [`Column`]: either a text block or a web
/// part. Mirrors PnPjs's `ColumnControl<T>` subclasses (`ClientsideText`,
/// `ClientsideWebpart`).
#[derive(Debug, Clone)]
pub enum ColumnControl {
    Text(TextData),
    WebPart(Box<WebPartData>),
}

impl ColumnControl {
    fn position(&self) -> &Position {
        match self {
            Self::Text(t) => &t.position,
            Self::WebPart(w) => &w.position,
        }
    }

    fn position_mut(&mut self) -> &mut Position {
        match self {
            Self::Text(t) => &mut t.position,
            Self::WebPart(w) => &mut w.position,
        }
    }

    fn emphasis_mut(&mut self) -> &mut Emphasis {
        match self {
            Self::Text(t) => &mut t.emphasis,
            Self::WebPart(w) => &mut w.emphasis,
        }
    }
}

/// One entry of the flat `CanvasContent1` JSON array before it's merged
/// into the sections/columns tree. Mirrors the four cases
/// `_ClientsidePage.setControls`'s `switch (controlType)` handles.
#[derive(Debug, Clone)]
pub enum Control {
    /// An empty column (`controlType: 0`, no `pageSettingsSlice`).
    Column(ColumnData),
    /// The page-settings entry (`controlType: 0`, has `pageSettingsSlice`).
    PageSettings(PageSettingsSlice),
    WebPart(Box<WebPartData>),
    Text(TextData),
}

impl Serialize for Control {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        match self {
            Self::Column(c) => c.serialize(serializer),
            Self::PageSettings(p) => p.serialize(serializer),
            Self::WebPart(w) => w.serialize(serializer),
            Self::Text(t) => t.serialize(serializer),
        }
    }
}

impl<'de> Deserialize<'de> for Control {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        let v = Value::deserialize(deserializer)?;
        // PnPjs: `hOP(controls[i], "controlType") ? controls[i].controlType : 0`
        let control_type = v.get("controlType").and_then(Value::as_i64).unwrap_or(0);
        match control_type {
            0 if v.get("pageSettingsSlice").is_some() => serde_json::from_value(v)
                .map(Control::PageSettings)
                .map_err(serde::de::Error::custom),
            0 => serde_json::from_value(v)
                .map(Control::Column)
                .map_err(serde::de::Error::custom),
            3 => serde_json::from_value(v)
                .map(|w| Control::WebPart(Box::new(w)))
                .map_err(serde::de::Error::custom),
            4 => serde_json::from_value(v)
                .map(Control::Text)
                .map_err(serde::de::Error::custom),
            other => Err(serde::de::Error::custom(format!(
                "unrecognized canvas controlType {other} (PnPjs itself drops these when building \
                 the section/column tree; forge-m365-sp-clientside-pages surfaces it instead of \
                 silently discarding data)"
            ))),
        }
    }
}

/// A column within a [`Section`]. Mirrors `CanvasColumn`.
#[derive(Debug, Clone)]
pub struct Column {
    pub data: ColumnData,
    pub controls: Vec<ColumnControl>,
}

impl Column {
    pub fn factor(&self) -> i64 {
        self.data.position.section_factor.unwrap_or(12)
    }

    /// Appends a control to this column, updating its position fields to
    /// match the column (mirrors `ColumnControl.onColumnChange`). Order is
    /// finalized by [`Page::reindex`].
    pub fn add_control(&mut self, mut control: ColumnControl) {
        {
            let factor = self.factor();
            let zone_index = self.data.position.zone_index;
            let section_index = self.data.position.section_index;
            let layout_index = self.data.position.layout_index;
            let pos = control.position_mut();
            pos.section_factor = Some(factor);
            pos.zone_index = zone_index;
            pos.section_index = section_index;
            pos.layout_index = layout_index;
        }
        self.controls.push(control);
    }
}

/// A section on the page. Mirrors `CanvasSection`. `layout_index` is `1` for
/// a normal section, `2` for the page's (single, special-cased) vertical
/// section.
#[derive(Debug, Clone)]
pub struct Section {
    pub order: i64,
    pub layout_index: i64,
    pub emphasis: u8,
    pub columns: Vec<Column>,
}

impl Section {
    /// Adds a column with the given factor (one of 0, 2, 4, 6, 8, 12),
    /// mirroring `CanvasSection.addColumn`.
    pub fn add_column(&mut self, factor: i64) -> usize {
        let order = next_order(self.columns.iter().map(|c| c.data.position.section_index));
        let mut data = ColumnData::default_for(self.order, self.layout_index, factor);
        data.position.section_index = order;
        self.columns.push(Column {
            data,
            controls: Vec::new(),
        });
        self.columns.len() - 1
    }

    /// Index of the default (first) column, creating a full-width one if
    /// the section has none yet, mirroring `CanvasSection.defaultColumn`.
    pub fn default_column_index(&mut self) -> usize {
        if self.columns.is_empty() {
            self.add_column(12);
        }
        0
    }

    /// Adds a control to the section's default column, mirroring
    /// `CanvasSection.addControl`.
    pub fn add_control(&mut self, control: ColumnControl) {
        let idx = self.default_column_index();
        self.columns[idx].add_control(control);
    }
}

fn next_order(orders: impl Iterator<Item = i64>) -> i64 {
    orders.max().map(|m| m + 1).unwrap_or(1)
}

/// A parsed/constructed client-side page canvas: the sections/columns/
/// controls tree plus the page-settings slice, if present. This is the
/// in-memory counterpart to the `CanvasContent1` JSON field. Mirrors the
/// tree-related parts of `_ClientsidePage` (everything except the network
/// operations — see the crate-level docs).
#[derive(Debug, Clone, Default)]
pub struct Page {
    pub sections: Vec<Section>,
    pub page_settings: Option<PageSettingsSlice>,
}

impl Page {
    /// Adds a normal (`layoutIndex: 1`) section, mirroring
    /// `_ClientsidePage.addSection`.
    pub fn add_section(&mut self) -> usize {
        let order = next_order(self.sections.iter().map(|s| s.order));
        self.sections.push(Section {
            order,
            layout_index: 1,
            emphasis: 0,
            columns: Vec::new(),
        });
        self.sections.len() - 1
    }

    /// Adds the page's vertical section (`layoutIndex: 2`), mirroring
    /// `_ClientsidePage.addVerticalSection`. SharePoint pages have at most
    /// one; this crate doesn't enforce that, matching PnPjs (which doesn't
    /// either).
    pub fn add_vertical_section(&mut self) -> usize {
        let order = next_order(self.sections.iter().map(|s| s.order));
        self.sections.push(Section {
            order,
            layout_index: 2,
            emphasis: 0,
            columns: Vec::new(),
        });
        self.sections.len() - 1
    }

    pub fn has_vertical_section(&self) -> bool {
        self.sections.iter().any(|s| s.layout_index == 2)
    }

    fn get_or_create_section(&mut self, order: i64, layout_index: i64, emphasis: u8) -> usize {
        if let Some(idx) = self
            .sections
            .iter()
            .position(|s| s.order == order && s.layout_index == layout_index)
        {
            return idx;
        }
        let idx = if layout_index == 2 {
            self.add_vertical_section()
        } else {
            self.add_section()
        };
        self.sections[idx].order = order;
        self.sections[idx].emphasis = emphasis;
        idx
    }

    /// Mirrors `_ClientsidePage.mergePartToTree`.
    fn merge_part_to_tree(&mut self, control: ColumnControl) {
        let pos = *control.position();
        let zone_emphasis = match &control {
            ColumnControl::Text(t) => t.emphasis.zone_emphasis,
            ColumnControl::WebPart(w) => w.emphasis.zone_emphasis,
        }
        .unwrap_or(0);

        let section_idx =
            self.get_or_create_section(pos.zone_index, pos.layout_index, zone_emphasis);
        let section = &mut self.sections[section_idx];
        let column_idx = section
            .columns
            .iter()
            .position(|c| c.data.position.section_index == pos.section_index)
            .unwrap_or_else(|| section.add_column(pos.section_factor.unwrap_or(12)));
        section.columns[column_idx].controls.push(control);
    }

    /// Mirrors `_ClientsidePage.mergeColumnToTree`.
    fn merge_column_to_tree(&mut self, column: Column) {
        let order = column.data.position.zone_index;
        let layout_index = column.data.position.layout_index;
        let emphasis = column.data.emphasis.zone_emphasis.unwrap_or(0);
        let section_idx = self.get_or_create_section(order, layout_index, emphasis);
        self.sections[section_idx].columns.push(column);
    }

    /// Normalizes every section/column/control order to be 1-based and
    /// contiguous, mirroring `reindex()` in `funcs.ts`.
    pub fn reindex(&mut self) {
        for (i, section) in self.sections.iter_mut().enumerate() {
            section.order = (i + 1) as i64;
            for (j, column) in section.columns.iter_mut().enumerate() {
                column.data.position.section_index = (j + 1) as i64;
                for (k, control) in column.controls.iter_mut().enumerate() {
                    control.position_mut().control_index = Some((k + 1) as i64);
                }
            }
        }
    }

    /// Parses a `CanvasContent1` JSON string into a [`Page`], running the
    /// same tree-merge algorithm as PnPjs's `_ClientsidePage.setControls`.
    ///
    /// Ported from PnPjs `_ClientsidePage.setControls`
    /// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
    /// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
    pub fn from_canvas_content_json(json: &str) -> serde_json::Result<Self> {
        let raw: Vec<Value> = serde_json::from_str(json)?;
        let mut page = Page::default();
        // A `null` entry is the trailing page-settings slot when none was
        // ever set (see `to_canvas_content_json`'s doc comment) — not a
        // control, so it's skipped rather than fed to `Control`'s
        // deserializer.
        let controls = raw
            .into_iter()
            .filter(|v| !v.is_null())
            .map(serde_json::from_value::<Control>)
            .collect::<serde_json::Result<Vec<_>>>()?;
        for control in controls {
            match control {
                Control::PageSettings(s) => page.page_settings = Some(s),
                Control::Column(data) => page.merge_column_to_tree(Column {
                    data,
                    controls: Vec::new(),
                }),
                Control::WebPart(w) => page.merge_part_to_tree(ColumnControl::WebPart(w)),
                Control::Text(t) => page.merge_part_to_tree(ColumnControl::Text(t)),
            }
        }
        page.reindex();
        Ok(page)
    }

    /// Serializes this page back to the flat `CanvasContent1` JSON shape,
    /// running the same emphasis-stamping and escaping PnPjs's
    /// `_ClientsidePage.getControls`/`specialSaveHandling` perform.
    ///
    /// Note: PnPjs always appends the page-settings slot to this array even
    /// when none was ever loaded (`canvasData.push(this._pageSettings)`
    /// where the field is `undefined`, which `JSON.stringify` renders as
    /// `null`). This is reproduced here for byte-for-byte parity, so a page
    /// with `page_settings: None` still serializes a trailing `null`.
    ///
    /// Ported from PnPjs `_ClientsidePage.getControls`/`getCanvasContent1`
    /// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
    /// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
    pub fn to_canvas_content_json(&mut self) -> serde_json::Result<String> {
        self.reindex();
        let mut out: Vec<Value> = Vec::new();

        for section in &self.sections {
            for column in &section.columns {
                if column.controls.is_empty() {
                    let entry = ColumnData {
                        control_type: 0,
                        display_mode: column.data.display_mode,
                        emphasis: Emphasis::from_section_value(section.emphasis),
                        position: column.data.position,
                    };
                    out.push(serde_json::to_value(entry)?);
                } else {
                    for control in &column.controls {
                        let mut control = control.clone();
                        *control.emphasis_mut() = Emphasis::from_section_value(section.emphasis);
                        special_save_handling(&mut control);
                        let value = match control {
                            ColumnControl::Text(t) => serde_json::to_value(t)?,
                            ColumnControl::WebPart(w) => serde_json::to_value(w)?,
                        };
                        out.push(value);
                    }
                }
            }
        }

        out.push(match &self.page_settings {
            Some(s) => serde_json::to_value(s)?,
            None => Value::Null,
        });

        serde_json::to_string(&out)
    }
}

/// Mirrors `_ClientsidePage.specialSaveHandling` (PnPjs issue #1690 / #2561):
/// `<` must be escaped as `&lt;` in a web part's
/// `serverProcessedContent.searchablePlainTexts` values before saving.
fn special_save_handling(control: &mut ColumnControl) {
    if let ColumnControl::WebPart(w) = control {
        if let Some(spc) = w.web_part_data.server_processed_content.as_mut() {
            if let Some(texts) = spc.searchable_plain_texts.as_mut() {
                for v in texts.values_mut() {
                    *v = v.replace('<', "&lt;");
                }
            }
        }
    }
}

/// The single-element `LayoutWebpartsContent` payload (the page header/
/// banner web part). PnPjs's `ILayoutPartsContent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutPartsContent {
    pub id: String,
    pub instance_id: String,
    pub title: String,
    pub description: String,
    pub server_processed_content: LayoutServerProcessedContent,
    pub data_version: String,
    pub properties: LayoutProperties,
    pub reserved_height: i64,
}

impl LayoutPartsContent {
    /// The default "title area" layout part PnPjs's `_ClientsidePage`
    /// constructor always assigns before `fromJSON` runs — a page's
    /// `_layoutPart` is never actually null/absent in PnPjs, only the wire
    /// value of `LayoutWebpartsContent` can be. Used by
    /// [`ClientsidePage::from_raw`] to fill in a missing layout part, and by
    /// [`create_client_side_page`] for a brand-new page.
    ///
    /// Mirrors PnPjs `_ClientsidePage.getDefaultLayoutPart`
    /// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
    /// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
    pub fn default_title_area() -> Self {
        Self {
            id: "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788".to_string(),
            instance_id: "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788".to_string(),
            title: "Title area".to_string(),
            description: "Title Region Description".to_string(),
            server_processed_content: LayoutServerProcessedContent {
                html_strings: Some(HashMap::new()),
                searchable_plain_texts: Some(HashMap::new()),
                image_sources: Some(HashMap::new()),
                links: Some(HashMap::new()),
                custom_metadata: None,
            },
            data_version: "1.4".to_string(),
            properties: LayoutProperties {
                title: String::new(),
                image_source_type: None,
                layout_type: "FullWidthImage".to_string(),
                text_alignment: "Left".to_string(),
                show_topic_header: false,
                show_publish_date: false,
                topic_header: String::new(),
                enable_gradient_effect: true,
                author_byline: Vec::new(),
                authors: Vec::new(),
                web_id: None,
                site_id: None,
                list_id: None,
                unique_id: None,
                translate_x: None,
                translate_y: None,
                alt_text: None,
                has_title_been_committed: None,
            },
            reserved_height: 280,
        }
    }
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutServerProcessedContent {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_strings: Option<HashMap<String, String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub searchable_plain_texts: Option<HashMap<String, String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_sources: Option<HashMap<String, String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub links: Option<HashMap<String, String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub custom_metadata: Option<Value>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutProperties {
    pub title: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub image_source_type: Option<i64>,
    pub layout_type: String,
    pub text_alignment: String,
    pub show_topic_header: bool,
    pub show_publish_date: bool,
    pub topic_header: String,
    pub enable_gradient_effect: bool,
    pub author_byline: Vec<String>,
    #[serde(default)]
    pub authors: Vec<LayoutAuthor>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub web_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub site_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub list_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub unique_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub translate_x: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub translate_y: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub alt_text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub has_title_been_committed: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct LayoutAuthor {
    pub id: String,
    pub upn: String,
    pub name: String,
    pub role: String,
}

/// Parses a `LayoutWebpartsContent` JSON string (either the literal `null`
/// or a single-element array) into the layout part it wraps, if any.
///
/// Ported from PnPjs `_ClientsidePage.load`'s `LayoutWebpartsContent`
/// handling (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub fn parse_layout_webparts_content(json: &str) -> serde_json::Result<Option<LayoutPartsContent>> {
    let value: Value = serde_json::from_str(json)?;
    match value {
        Value::Null => Ok(None),
        Value::Array(mut items) if !items.is_empty() => {
            serde_json::from_value(items.remove(0)).map(Some)
        }
        _ => Ok(None),
    }
}

/// Serializes a layout part back to the `LayoutWebpartsContent` wire shape:
/// `null` when absent, a single-element array otherwise.
///
/// Ported from PnPjs `_ClientsidePage.getLayoutWebpartsContent`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub fn to_layout_webparts_content_json(
    part: Option<&LayoutPartsContent>,
) -> serde_json::Result<String> {
    match part {
        Some(p) => serde_json::to_string(&[p]),
        None => Ok("null".to_string()),
    }
}

/// A v4 GUID formatted like PnPjs's `getGUID()` (lowercase, hyphenated,
/// version/variant bits set correctly) — used for control/instance ids that
/// don't come from the server.
pub fn new_guid() -> String {
    let mut bytes = [0u8; 16];
    for b in bytes.iter_mut() {
        *b = rand::random();
    }
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    format!(
        "{:02x}{:02x}{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
        bytes[0], bytes[1], bytes[2], bytes[3],
        bytes[4], bytes[5],
        bytes[6], bytes[7],
        bytes[8], bytes[9],
        bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15],
    )
}

// ---------------------------------------------------------------------
// Phase 2: read-only network operations
// ---------------------------------------------------------------------

/// `Path.DecodedUrl` on `IPageData`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct PagePath {
    pub decoded_url: String,
}

/// `VersionInfo` on `IPageData`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct VersionInfo {
    pub last_version_created: String,
    pub last_version_created_by: String,
}

/// A client-side page's server-tracked metadata — everything in PnPjs's
/// `IPageData` except `CanvasContent1`/`LayoutWebpartsContent`, which are
/// parsed into [`ClientsidePage::content`]/[`ClientsidePage::layout`]
/// instead of kept as raw strings. The `odata.*` fields PnPjs's interface
/// carries for its own fluent-query plumbing are omitted — nothing in this
/// crate needs them.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct PageMetadata {
    pub absolute_url: String,
    #[serde(default)]
    pub author_byline: Option<Vec<String>>,
    pub banner_image_url: String,
    pub banner_thumbnail_url: String,
    #[serde(default)]
    pub content_type_id: Option<String>,
    pub description: String,
    pub does_user_have_edit_permission: bool,
    pub file_name: String,
    pub first_published: String,
    pub id: i64,
    pub is_page_checked_out_to_current_user: bool,
    pub is_web_welcome_page: bool,
    pub modified: String,
    pub page_layout_type: String,
    pub path: PagePath,
    pub promoted_state: i64,
    pub title: String,
    #[serde(default)]
    pub topic_header: Option<String>,
    pub unique_id: String,
    pub url: String,
    pub version: String,
    pub version_info: VersionInfo,
    #[serde(default)]
    pub alternative_url_map: Option<String>,
}

/// The raw wire shape of `_api/sitepages/pages({id})`: [`PageMetadata`]'s
/// fields plus the two canvas content strings, flattened into one object as
/// SharePoint actually returns it.
#[derive(Debug, Clone, Deserialize)]
struct RawPageData {
    #[serde(flatten)]
    metadata: PageMetadata,
    #[serde(rename = "CanvasContent1")]
    canvas_content1: String,
    #[serde(rename = "LayoutWebpartsContent")]
    layout_webparts_content: String,
}

/// A fully loaded client-side page: metadata plus the already-parsed canvas
/// tree and header/banner layout part. The in-memory counterpart to a
/// loaded PnPjs `_ClientsidePage` instance (minus the network-write half —
/// see the crate-level docs for what's still deferred).
#[derive(Debug, Clone)]
pub struct ClientsidePage {
    pub metadata: PageMetadata,
    pub content: Page,
    /// Never actually absent once a page exists in memory — PnPjs's
    /// `_ClientsidePage` constructor always assigns
    /// [`LayoutPartsContent::default_title_area`] before parsing whatever
    /// `LayoutWebpartsContent` the server sent, so a wire value of `null`
    /// (a page whose header/banner was never customized) becomes the
    /// default here rather than `None`. Use [`parse_layout_webparts_content`]
    /// directly if you need to distinguish "never customized" from "has the
    /// exact default values" at the wire level.
    pub layout: LayoutPartsContent,
    /// Whether comments are disabled on this page. Only known when the page
    /// was loaded via [`load_client_side_page`] (which resolves it from the
    /// underlying list item, per PnPjs's `_ClientsidePage.load`); `None`
    /// when loaded via [`get_client_side_page`] alone, since
    /// `_api/sitepages/pages({id})` doesn't carry it.
    pub comments_disabled: Option<bool>,
}

impl ClientsidePage {
    fn from_raw(raw: RawPageData) -> serde_json::Result<Self> {
        let content = Page::from_canvas_content_json(&raw.canvas_content1)?;
        let layout = parse_layout_webparts_content(&raw.layout_webparts_content)?
            .unwrap_or_else(LayoutPartsContent::default_title_area);
        Ok(Self {
            metadata: raw.metadata,
            content,
            layout,
            comments_disabled: None,
        })
    }
}

/// The list-item identity behind a page file, as returned by
/// `getFileByServerRelativePath(...)/ListItemAllFields?$select=Id,CommentsDisabled`
/// — mirrors what PnPjs's `_ClientsidePage.load` reads off the underlying
/// file's item before fetching the page itself.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct PageItemRef {
    pub id: i64,
    pub comments_disabled: bool,
}

/// A client-side web part's manifest entry, as returned by
/// `_api/web/GetClientSideWebParts`. PnPjs's `IClientsidePageComponent`.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "PascalCase")]
pub struct ClientsidePageComponent {
    pub component_type: i64,
    pub id: String,
    pub manifest: String,
    pub manifest_type: i64,
    pub name: String,
    pub status: i64,
}

/// Doubles embedded `'` and percent-encodes for use inside an OData
/// `decodedUrl='...'` path segment. Ported from PnPjs `encodePath`
/// (`packages/sp/utils/encode-path-str.ts`) — duplicated locally rather
/// than depending on `forge-m365-sp-files` for four lines of logic, per
/// this workspace's usual practice for this helper.
fn encode_path_segment(value: &str) -> String {
    let doubled = value.replace('\'', "''");
    utf8_percent_encode(&doubled, NON_ALPHANUMERIC).to_string()
}

fn pages_url(site_url: &str, id: i64) -> String {
    format!(
        "{}/_api/sitepages/pages({id})",
        site_url.trim_end_matches('/')
    )
}

fn file_item_ref_url(site_url: &str, server_relative_path: &str) -> String {
    format!(
        "{}/_api/web/getFileByServerRelativePath(decodedUrl='{}')/ListItemAllFields?$select=Id,CommentsDisabled",
        site_url.trim_end_matches('/'),
        encode_path_segment(server_relative_path)
    )
}

fn client_side_web_parts_url(site_url: &str) -> String {
    format!(
        "{}/_api/web/GetClientSideWebParts",
        site_url.trim_end_matches('/')
    )
}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-clientside-pages linked in?")
        })
}

// These bodies are intentionally empty: #[pnp_operation] only uses the attribute
// to register the ladder in the inventory at compile time. The real HTTP calls
// live in the public fns below, which look the entry back up by id.
#[pnp_operation(id = "sp.clientside_pages.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_id_op() {}

#[pnp_operation(
    id = "sp.clientside_pages.get_page_item_ref",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_page_item_ref_op() {}

#[pnp_operation(
    id = "sp.clientside_pages.get_client_side_web_parts",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_client_side_web_parts_op() {}

/// Gets a client-side page by its list-item id.
///
/// Doesn't know whether comments are disabled on the page (see
/// [`ClientsidePage::comments_disabled`]) — use [`load_client_side_page`]
/// if you need that and only have the page's path.
///
/// Ported from PnPjs `_ClientsidePage.load`'s `_api/sitepages/pages({id})`
/// GET (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_client_side_page(
    client: &Client<'_>,
    site_url: &str,
    id: i64,
) -> CoreResult<ClientsidePage> {
    let url = pages_url(site_url, id);
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.get_by_id"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let raw: RawPageData = serde_json::from_slice(&bytes)?;
    Ok(ClientsidePage::from_raw(raw)?)
}

/// Resolves a page file's list-item id and comments-disabled flag from its
/// server-relative path (e.g. `/sites/team/SitePages/home.aspx`).
///
/// Ported from PnPjs `_ClientsidePage.load`'s
/// `getItem<{Id, CommentsDisabled}>("Id", "CommentsDisabled")` call
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_page_item_ref(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> CoreResult<PageItemRef> {
    let url = file_item_ref_url(site_url, server_relative_path);
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.get_page_item_ref"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Loads a client-side page from its server-relative path: resolves the
/// underlying list item first (for its id and comments-disabled flag), then
/// fetches the page itself. This is a composite of [`get_page_item_ref`]
/// and [`get_client_side_page`] — it doesn't register its own ladder entry.
///
/// Ported from PnPjs `Web.loadClientsidePage`/`ClientsidePageFromFile`
/// (`packages/sp/clientside-pages/web.ts`, `types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn load_client_side_page(
    client: &Client<'_>,
    site_url: &str,
    server_relative_path: &str,
) -> CoreResult<ClientsidePage> {
    let item_ref = get_page_item_ref(client, site_url, server_relative_path).await?;
    let mut page = get_client_side_page(client, site_url, item_ref.id).await?;
    page.comments_disabled = Some(item_ref.comments_disabled);
    Ok(page)
}

/// Gets the client-side web part manifests available in this web — the
/// catalog `ClientsideWebpart.fromComponentDef`/`Web.addFullPageApp` search
/// by component id to build a [`WebPartData`] with a known title/
/// description/properties (that construction step itself is a later
/// phase — see the crate README).
///
/// Ported from PnPjs `_Web.getClientsideWebParts`
/// (`packages/sp/clientside-pages/web.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn get_client_side_web_parts(
    client: &Client<'_>,
    site_url: &str,
) -> CoreResult<Vec<ClientsidePageComponent>> {
    let url = client_side_web_parts_url(site_url);
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.get_client_side_web_parts"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

// ---------------------------------------------------------------------
// Phase 3: lifecycle operations that don't touch the canvas tree
// ---------------------------------------------------------------------

/// Properties that control how a banner image displays. PnPjs's
/// `IBannerImageProps`.
#[derive(Debug, Clone, Default)]
pub struct BannerImageProps {
    pub alt_text: Option<String>,
    pub image_source_type: Option<i64>,
    pub translate_x: Option<f64>,
    pub translate_y: Option<f64>,
}

#[derive(Debug, Deserialize)]
struct SiteUserLoginNameRow {
    #[serde(rename = "LoginName")]
    login_name: String,
}

#[derive(Debug, Deserialize)]
struct SiteUserProfileLite {
    #[serde(rename = "UserPrincipalName")]
    user_principal_name: String,
    #[serde(rename = "Title")]
    title: String,
}

impl ClientsidePage {
    /// Sets the page's banner image, purely in memory — call
    /// [`save_client_side_page`] to persist it. A simplified port: PnPjs's
    /// `isUrlAbsolute` + host-stripping regex requires the host to look
    /// like a domain (`some.tld/...`); this strips the scheme and
    /// authority of any `http(s)://` URL unconditionally, which is
    /// equivalent for every real SharePoint absolute URL and simpler to
    /// reason about here.
    ///
    /// Ported from PnPjs `_ClientsidePage.setBannerImage`
    /// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
    /// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
    pub fn set_banner_image(&mut self, url: impl Into<String>, props: Option<BannerImageProps>) {
        let mut url = url.into();
        if let Some(rest) = url
            .strip_prefix("https://")
            .or_else(|| url.strip_prefix("http://"))
        {
            if let Some(slash) = rest.find('/') {
                url = format!("/{}", &rest[slash + 1..]);
            }
        }

        self.metadata.banner_image_url = url.clone();
        let mut image_sources = HashMap::new();
        image_sources.insert("imageSource".to_string(), url.clone());
        self.layout.server_processed_content.image_sources = Some(image_sources);
        self.metadata.banner_thumbnail_url = url;
        self.layout.properties.image_source_type = Some(2);

        if let Some(props) = props {
            if props.translate_x.is_some() {
                self.layout.properties.translate_x = props.translate_x;
            }
            if props.translate_y.is_some() {
                self.layout.properties.translate_y = props.translate_y;
            }
            if props.image_source_type.is_some() {
                self.layout.properties.image_source_type = props.image_source_type;
            }
            if props.alt_text.is_some() {
                self.layout.properties.alt_text = props.alt_text;
            }
        }
    }

    /// Sets the page's topic header, purely in memory. Mirrors PnPjs's
    /// `topicHeader` setter: updates both the page metadata and the layout
    /// part together, and clears `showTopicHeader` when the header is
    /// emptied.
    ///
    /// Ported from PnPjs `_ClientsidePage`'s `topicHeader` setter
    /// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
    /// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
    pub fn set_topic_header(&mut self, value: impl Into<String>) {
        let value = value.into();
        self.layout.properties.topic_header = value.clone();
        if value.is_empty() {
            self.layout.properties.show_topic_header = false;
            self.metadata.topic_header = None;
        } else {
            self.metadata.topic_header = Some(value);
        }
    }
}

fn action_url(site_url: &str, id: i64, action: &str) -> String {
    format!("{}/{action}", pages_url(site_url, id))
}

fn site_pages_item_url(site_url: &str, id: i64) -> String {
    // A documented simplification of PnPjs's `getItem()`, which resolves
    // the Site Pages library via `EnsureClientRenderedSitePagesLibrary`
    // rather than assuming its title — see this crate's README.
    format!(
        "{}/_api/web/lists/getbytitle('Site Pages')/items({id})",
        site_url.trim_end_matches('/')
    )
}

fn add_image_from_external_url_url(
    site_url: &str,
    image_file_name: &str,
    page_name: &str,
    external_url: &str,
) -> String {
    let q = |v: &str| utf8_percent_encode(&format!("'{v}'"), NON_ALPHANUMERIC).to_string();
    format!(
        "{}/_api/sitepages/AddImageFromExternalUrl?imageFileName={}&pageName={}&externalUrl={}&$select=ServerRelativeUrl",
        site_url.trim_end_matches('/'),
        q(image_file_name),
        q(page_name),
        q(external_url)
    )
}

#[pnp_operation(id = "sp.clientside_pages.discard_checkout", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn discard_checkout_op() {}

#[pnp_operation(id = "sp.clientside_pages.promote_to_news", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn promote_to_news_op() {}

#[pnp_operation(id = "sp.clientside_pages.schedule_publish", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn schedule_publish_op() {}

#[pnp_operation(id = "sp.clientside_pages.share_by_email", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn share_by_email_op() {}

#[pnp_operation(
    id = "sp.clientside_pages.get_site_user_by_id",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_site_user_by_id_op() {}

#[pnp_operation(
    id = "sp.clientside_pages.get_site_user_by_login_name",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn get_site_user_by_login_name_op() {}

#[pnp_operation(id = "sp.clientside_pages.recycle", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn recycle_op() {}

#[pnp_operation(id = "sp.clientside_pages.delete", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_op() {}

#[pnp_operation(
    id = "sp.clientside_pages.set_banner_image_from_external_url",
    primary = Surface::SpRest,
    fallback = []
)]
#[allow(dead_code)]
async fn set_banner_image_from_external_url_op() {}

/// Discards this page's checkout, reloading it from the server response.
///
/// Ported from PnPjs `_ClientsidePage.discardPageCheckout`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn discard_page_checkout(
    client: &Client<'_>,
    site_url: &str,
    id: i64,
) -> CoreResult<ClientsidePage> {
    let url = action_url(site_url, id, "discardPage");
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.discard_checkout"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    let raw: RawPageData = serde_json::from_slice(&bytes)?;
    Ok(ClientsidePage::from_raw(raw)?)
}

/// Promotes this page as a news item. A simplified port: PnPjs's
/// `promoteNewsImpl` first auto-publishes the page if it has never been
/// published (working around a server bug), which needs
/// [`save_client_side_page`]/full page state; that auto-publish step is
/// not reproduced here — publish the page yourself first if it's new.
///
/// Ported from PnPjs `_ClientsidePage.promoteToNews`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn promote_to_news(client: &Client<'_>, site_url: &str, id: i64) -> CoreResult<bool> {
    let url = action_url(site_url, id, "promoteToNews");
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.promote_to_news"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Schedules a page for publishing, returning the version that was
/// scheduled. `publish_date` is an ISO-8601 datetime string.
///
/// Ported from PnPjs `_ClientsidePage.schedulePublish`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn schedule_publish(
    client: &Client<'_>,
    site_url: &str,
    id: i64,
    publish_date: &str,
) -> CoreResult<String> {
    let url = action_url(site_url, id, "SchedulePublish");
    let body = serde_json::to_vec(
        &serde_json::json!({ "sitePage": { "PublishStartDate": publish_date } }),
    )?;
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.schedule_publish"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Shares a page's preview content by email.
///
/// Ported from PnPjs `_ClientsidePage.share`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn share_page_by_email(
    client: &Client<'_>,
    site_url: &str,
    page_url: &str,
    emails: &[String],
    message: &str,
) -> CoreResult<()> {
    let url = format!(
        "{}/_api/SP.Publishing.RichSharing/SharePageByEmail",
        site_url.trim_end_matches('/')
    );
    let body = serde_json::to_vec(&serde_json::json!({
        "recipientEmails": emails,
        "message": message,
        "url": page_url,
    }))?;
    client
        .run_ladder(
            entry("sp.clientside_pages.share_by_email"),
            "POST",
            &url,
            &[],
            Some(&body),
        )
        .await?;
    Ok(())
}

/// Sets the page's author from a site-user integer id, purely in memory —
/// call [`save_client_side_page`] to persist it.
///
/// Ported from PnPjs `_ClientsidePage.setAuthorById`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn set_author_by_id(
    client: &Client<'_>,
    site_url: &str,
    page: &mut ClientsidePage,
    author_id: i64,
) -> CoreResult<()> {
    let url = format!(
        "{}/_api/web/siteusers?$filter=Id eq {author_id}&$select=LoginName",
        site_url.trim_end_matches('/')
    );
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.get_site_user_by_id"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let users: Vec<SiteUserLoginNameRow> = serde_json::from_slice(&bytes)?;
    let user = users.into_iter().next().ok_or_else(|| {
        CoreError::Surface(Surface::SpRest, format!("no site user with id {author_id}"))
    })?;
    set_author_by_login_name(client, site_url, page, &user.login_name).await
}

/// Sets the page's author from a site-user login name, purely in memory —
/// call [`save_client_side_page`] to persist it.
///
/// Ported from PnPjs `_ClientsidePage.setAuthorByLoginName`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn set_author_by_login_name(
    client: &Client<'_>,
    site_url: &str,
    page: &mut ClientsidePage,
    author_login_name: &str,
) -> CoreResult<()> {
    let url = format!(
        "{}/_api/web/siteusers?$filter=LoginName eq '{}'&$select=UserPrincipalName,Title",
        site_url.trim_end_matches('/'),
        encode_path_segment(author_login_name)
    );
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.get_site_user_by_login_name"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    let users: Vec<SiteUserProfileLite> = serde_json::from_slice(&bytes)?;
    let user = users.into_iter().next().ok_or_else(|| {
        CoreError::Surface(
            Surface::SpRest,
            format!("no site user with login name '{author_login_name}'"),
        )
    })?;

    page.metadata.author_byline = Some(vec![user.user_principal_name.clone()]);
    page.layout.properties.author_byline = vec![user.user_principal_name.clone()];
    page.layout.properties.authors = vec![LayoutAuthor {
        id: author_login_name.to_string(),
        upn: user.user_principal_name,
        name: user.title,
        role: String::new(),
    }];
    Ok(())
}

/// Sets the page's banner image from an externally-hosted image, uploading
/// it into the page's library first.
///
/// Ported from PnPjs `_ClientsidePage.setBannerImageFromExternalUrl`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn set_banner_image_from_external_url(
    client: &Client<'_>,
    site_url: &str,
    page: &mut ClientsidePage,
    external_url: &str,
    props: Option<BannerImageProps>,
) -> CoreResult<()> {
    let page_name = match page.metadata.file_name.rfind('.') {
        Some(i) => &page.metadata.file_name[..i],
        None => page.metadata.file_name.as_str(),
    };
    let filename = external_url
        .rsplit(['/', '\\'])
        .next()
        .unwrap_or(external_url);
    let url = add_image_from_external_url_url(site_url, filename, page_name, external_url);
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.set_banner_image_from_external_url"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;

    #[derive(Deserialize)]
    struct Resp {
        #[serde(rename = "ServerRelativeUrl")]
        server_relative_url: String,
    }
    let resp: Resp = serde_json::from_slice(&bytes)?;
    page.set_banner_image(resp.server_relative_url, props);
    Ok(())
}

/// Recycles this page. A documented simplification of PnPjs's `recycle()`
/// (see [`site_pages_item_url`]'s doc comment).
///
/// Ported from PnPjs `_ClientsidePage.recycle`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn recycle_page(client: &Client<'_>, site_url: &str, id: i64) -> CoreResult<()> {
    let url = format!("{}/recycle", site_pages_item_url(site_url, id));
    client
        .run_ladder(
            entry("sp.clientside_pages.recycle"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

/// Deletes this page outright. A documented simplification of PnPjs's
/// `delete()` (see [`site_pages_item_url`]'s doc comment).
///
/// Ported from PnPjs `_ClientsidePage.delete`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn delete_page(client: &Client<'_>, site_url: &str, id: i64) -> CoreResult<()> {
    let url = site_pages_item_url(site_url, id);
    client
        .run_ladder(
            entry("sp.clientside_pages.delete"),
            "DELETE",
            &url,
            &[("IF-MATCH", "*")],
            None,
        )
        .await?;
    Ok(())
}

// ---------------------------------------------------------------------
// Phase 4: create / save / copy — the canvas write path
// ---------------------------------------------------------------------

#[pnp_operation(id = "sp.clientside_pages.create", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_op() {}

#[pnp_operation(id = "sp.clientside_pages.save_page", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn save_page_op() {}

#[pnp_operation(id = "sp.clientside_pages.publish", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn publish_op() {}

#[pnp_operation(id = "sp.clientside_pages.save_as_template", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn save_as_template_op() {}

/// Persists this page's content changes (canvas + core metadata) to
/// SharePoint, optionally publishing it, then reloads the page from the
/// server response (mirroring PnPjs's own final `await this.load()`) so
/// server-computed fields (URLs, version, etc.) stay current. Returns
/// whether the page ended up published.
///
/// A simplified port: PnPjs's `save()` also has a branch that, when the
/// banner image was just set from a raw uploaded file (not an external
/// URL — i.e. `_bannerImageDirty`), issues a small batch of two GETs to
/// compute a `getpreview.ashx`-based banner URL carrying the image's site/
/// web/file GUIDs. That branch is not reproduced here; banner images set
/// via [`set_banner_image_from_external_url`] (which already produces a
/// plain server-relative URL) are unaffected — this only matters for a
/// banner set from a file uploaded outside this flow.
///
/// Ported from PnPjs `_ClientsidePage.save`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn save_client_side_page(
    client: &Client<'_>,
    site_url: &str,
    page: &mut ClientsidePage,
    publish: bool,
) -> CoreResult<bool> {
    let id = page.metadata.id;
    let canvas_content1 = page.content.to_canvas_content_json()?;
    let layout_webparts_content = to_layout_webparts_content_json(Some(&page.layout))?;
    let save_body = serde_json::to_vec(&serde_json::json!({
        "AuthorByline": page.metadata.author_byline.clone().unwrap_or_default(),
        "CanvasContent1": canvas_content1,
        "Description": page.metadata.description,
        "LayoutWebpartsContent": layout_webparts_content,
        "Title": page.metadata.title,
        "TopicHeader": page.metadata.topic_header,
        "BannerImageUrl": page.metadata.banner_image_url,
    }))?;

    let save_url = action_url(site_url, id, "savepage");
    client
        .run_ladder(
            entry("sp.clientside_pages.save_page"),
            "POST",
            &save_url,
            &[("if-match", "*")],
            Some(&save_body),
        )
        .await?;

    let mut published = true;
    if publish {
        let publish_url = action_url(site_url, id, "publish");
        let bytes = client
            .run_ladder(
                entry("sp.clientside_pages.publish"),
                "POST",
                &publish_url,
                &[],
                None,
            )
            .await?;
        published = serde_json::from_slice(&bytes)?;
        if published {
            page.metadata.is_page_checked_out_to_current_user = false;
        }
    }

    let comments_disabled = page.comments_disabled;
    let reloaded = get_client_side_page(client, site_url, id).await?;
    *page = reloaded;
    page.comments_disabled = comments_disabled;

    Ok(published)
}

/// Creates a new client-side page. `page_name` is the file name (with or
/// without `.aspx`); `page_layout_type` is one of `"Article"`, `"Home"`,
/// `"SingleWebPartAppPage"`, `"RepostPage"`.
///
/// Faithfully reproduces an odd PnPjs quirk: the page is saved once with
/// `title` set to `page_name`, and only afterward is `title` set to the
/// caller's requested `title` — without a second save. The returned
/// page's in-memory `title` is therefore correct, but if you reload it
/// from the server immediately after, you'll see `page_name` as the title
/// until you call [`save_client_side_page`] again yourself.
///
/// Ported from PnPjs `CreateClientsidePage`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn create_client_side_page(
    client: &Client<'_>,
    site_url: &str,
    page_name: &str,
    title: &str,
    page_layout_type: &str,
    promoted_state: i64,
) -> CoreResult<ClientsidePage> {
    let page_name = page_name.strip_suffix(".aspx").unwrap_or(page_name);

    let create_url = format!("{}/_api/sitepages/pages", site_url.trim_end_matches('/'));
    let body = serde_json::to_vec(&serde_json::json!({
        "PageLayoutType": page_layout_type,
        "PromotedState": promoted_state,
    }))?;
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.create"),
            "POST",
            &create_url,
            &[],
            Some(&body),
        )
        .await?;
    let raw: RawPageData = serde_json::from_slice(&bytes)?;
    let mut page = ClientsidePage::from_raw(raw)?;

    page.metadata.title = page_name.to_string();
    save_client_side_page(client, site_url, &mut page, false).await?;
    page.metadata.title = title.to_string();

    Ok(page)
}

/// Saves a copy of this page as a template in the library's Templates
/// folder.
///
/// Ported from PnPjs `_ClientsidePage.saveAsTemplate`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn save_page_as_template(
    client: &Client<'_>,
    site_url: &str,
    page: &ClientsidePage,
    publish: bool,
) -> CoreResult<ClientsidePage> {
    let url = action_url(site_url, page.metadata.id, "SavePageAsTemplate");
    let bytes = client
        .run_ladder(
            entry("sp.clientside_pages.save_as_template"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    let raw: RawPageData = serde_json::from_slice(&bytes)?;
    let mut new_page = ClientsidePage::from_raw(raw)?;
    new_page.metadata.title = page.metadata.title.clone();
    save_client_side_page(client, site_url, &mut new_page, publish).await?;
    Ok(new_page)
}

/// Copies this page's canvas content and select layout properties onto an
/// existing page instance, fully overwriting the destination's previous
/// content, then saves the destination.
///
/// A simplified port: PnPjs's `copyTo` also re-links the banner image to
/// an equivalent file in the destination site via a cross-site GUID
/// lookup parsed out of the source's `getpreview.ashx`-based banner URL
/// (site/web/file GUIDs embedded as query params). That re-linking is not
/// reproduced here — the destination keeps whichever banner
/// [`create_client_side_page`] gave it; set one explicitly afterward if
/// needed.
///
/// Ported from PnPjs `_ClientsidePage.copyTo`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn copy_page_content_to(
    client: &Client<'_>,
    site_url: &str,
    source: &ClientsidePage,
    dest: &mut ClientsidePage,
    publish: bool,
) -> CoreResult<()> {
    dest.content = source.content.clone();

    let source_props = source.layout.properties.clone();
    dest.set_topic_header(source_props.topic_header.clone());
    dest.layout.properties.image_source_type = source_props.image_source_type;
    dest.layout.properties.layout_type = source_props.layout_type;
    dest.layout.properties.text_alignment = source_props.text_alignment;
    dest.layout.properties.show_topic_header = source_props.show_topic_header;
    dest.layout.properties.show_publish_date = source_props.show_publish_date;
    dest.layout.properties.enable_gradient_effect = source_props.enable_gradient_effect;

    save_client_side_page(client, site_url, dest, publish).await?;
    Ok(())
}

/// Creates a new page with all of the content copied from `source`.
///
/// Ported from PnPjs `_ClientsidePage.copy`
/// (`packages/sp/clientside-pages/types.ts`) @ pnpjs
/// `8ee2375dc3ce720926774cc87168e1251c9ed584`.
pub async fn copy_client_side_page(
    client: &Client<'_>,
    site_url: &str,
    source: &ClientsidePage,
    page_name: &str,
    title: &str,
    publish: bool,
    promoted_state: i64,
) -> CoreResult<ClientsidePage> {
    let mut dest = create_client_side_page(
        client,
        site_url,
        page_name,
        title,
        &source.metadata.page_layout_type,
        promoted_state,
    )
    .await?;
    copy_page_content_to(client, site_url, source, &mut dest, publish).await?;
    Ok(dest)
}
