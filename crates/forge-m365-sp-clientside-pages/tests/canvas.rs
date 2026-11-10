use forge_m365_sp_clientside_pages::{
    new_guid, parse_layout_webparts_content, to_layout_webparts_content_json, ColumnControl,
    Emphasis, Page, ServerProcessedContent, WebPartData,
};
use serde_json::json;

/// A realistic (trimmed) `CanvasContent1` payload: one section with a text
/// control and a web part control side by side in two columns, plus the
/// trailing page-settings slice PnPjs always appends.
const CANVAS_JSON: &str = r#"[
    {
        "controlType": 4,
        "displayMode": 2,
        "emphasis": {},
        "id": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        "position": { "zoneIndex": 1, "sectionIndex": 1, "controlIndex": 1, "sectionFactor": 6, "layoutIndex": 1 },
        "anchorComponentId": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
        "editorType": "CKEditor",
        "addedFromPersistedData": true,
        "innerHTML": "<p>hello</p>"
    },
    {
        "controlType": 3,
        "displayMode": 2,
        "emphasis": { "zoneEmphasis": 2 },
        "id": "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
        "position": { "zoneIndex": 1, "sectionIndex": 2, "controlIndex": 1, "sectionFactor": 6, "layoutIndex": 1 },
        "webPartId": "cccccccc-cccc-cccc-cccc-cccccccccccc",
        "reservedHeight": 300,
        "reservedWidth": 500,
        "addedFromPersistedData": true,
        "webPartData": {
            "id": "cccccccc-cccc-cccc-cccc-cccccccccccc",
            "instanceId": "dddddddd-dddd-dddd-dddd-dddddddddddd",
            "title": "Test Web Part",
            "description": "A test web part",
            "dataVersion": "1.0",
            "properties": { "foo": "bar" },
            "serverProcessedContent": {
                "searchablePlainTexts": { "text": "1 < 2" }
            }
        }
    },
    {
        "controlType": 0,
        "displayMode": 2,
        "emphasis": {},
        "position": { "zoneIndex": 2, "sectionIndex": 1, "layoutIndex": 1, "sectionFactor": 12 }
    },
    {
        "controlType": 0,
        "pageSettingsSlice": { "isDefaultDescription": true, "isDefaultThumbnail": false }
    }
]"#;

#[test]
fn parses_into_expected_tree_shape() {
    let page = Page::from_canvas_content_json(CANVAS_JSON).unwrap();

    assert_eq!(page.sections.len(), 2, "two distinct zoneIndex values");
    let section1 = &page.sections[0];
    assert_eq!(section1.order, 1);
    assert_eq!(
        section1.columns.len(),
        2,
        "two distinct sectionIndex values"
    );
    assert_eq!(section1.columns[0].controls.len(), 1);
    assert_eq!(section1.columns[1].controls.len(), 1);

    match &section1.columns[0].controls[0] {
        ColumnControl::Text(t) => assert_eq!(t.inner_html, "<p>hello</p>"),
        other => panic!("expected text control, got {other:?}"),
    }
    match &section1.columns[1].controls[0] {
        ColumnControl::WebPart(w) => {
            assert_eq!(w.web_part_data.title, "Test Web Part");
            assert_eq!(w.emphasis.zone_emphasis, Some(2));
        }
        other => panic!("expected web part control, got {other:?}"),
    }

    let section2 = &page.sections[1];
    assert_eq!(
        section2.columns.len(),
        1,
        "empty column merged into its own section"
    );
    assert!(section2.columns[0].controls.is_empty());

    let settings = page.page_settings.as_ref().unwrap();
    assert!(settings.page_settings_slice.is_default_description);
    assert!(!settings.page_settings_slice.is_default_thumbnail);
}

#[test]
fn round_trips_back_to_equivalent_json() {
    let mut page = Page::from_canvas_content_json(CANVAS_JSON).unwrap();
    let json = page.to_canvas_content_json().unwrap();

    let reparsed: serde_json::Value = serde_json::from_str(&json).unwrap();
    let arr = reparsed.as_array().unwrap();
    // 2 controls + 1 empty column + 1 trailing page-settings slot
    assert_eq!(arr.len(), 4);

    // Escaping from specialSaveHandling: "1 < 2" -> "1 &lt; 2"
    let webpart_entry = arr
        .iter()
        .find(|v| v.get("controlType").and_then(|c| c.as_i64()) == Some(3))
        .unwrap();
    assert_eq!(
        webpart_entry["webPartData"]["serverProcessedContent"]["searchablePlainTexts"]["text"],
        "1 &lt; 2"
    );
}

#[test]
fn reindex_is_1_based_and_contiguous() {
    let mut page = Page::default();
    let s0 = page.add_section();
    let s1 = page.add_section();
    // out-of-order zoneIndex/order values, as parsing an edited page might produce
    page.sections[s0].order = 5;
    page.sections[s1].order = 9;
    page.sections[s0].add_control(ColumnControl::Text(
        forge_m365_sp_clientside_pages_test_support::text("a"),
    ));
    page.sections[s0].add_control(ColumnControl::Text(
        forge_m365_sp_clientside_pages_test_support::text("b"),
    ));

    page.reindex();

    assert_eq!(page.sections[0].order, 1);
    assert_eq!(page.sections[1].order, 2);
    let col = &page.sections[0].columns[0];
    match &col.controls[0] {
        ColumnControl::Text(t) => assert_eq!(t.position.control_index, Some(1)),
        _ => unreachable!(),
    }
    match &col.controls[1] {
        ColumnControl::Text(t) => assert_eq!(t.position.control_index, Some(2)),
        _ => unreachable!(),
    }
}

#[test]
fn emphasis_serializes_as_empty_map_or_zone_emphasis() {
    assert_eq!(serde_json::to_string(&Emphasis::none()).unwrap(), "{}");
    assert_eq!(
        serde_json::to_string(&Emphasis::from_section_value(2)).unwrap(),
        r#"{"zoneEmphasis":2}"#
    );
    // out-of-range values normalize to "no emphasis", matching PnPjs's getEmphasisObj
    assert_eq!(
        serde_json::to_string(&Emphasis::from_section_value(0)).unwrap(),
        "{}"
    );
    assert_eq!(
        serde_json::to_string(&Emphasis::from_section_value(9)).unwrap(),
        "{}"
    );
}

#[test]
fn missing_page_settings_serializes_trailing_null() {
    let mut page = Page::default();
    page.add_section();
    let json = page.to_canvas_content_json().unwrap();
    let arr: serde_json::Value = serde_json::from_str(&json).unwrap();
    assert_eq!(
        *arr.as_array().unwrap().last().unwrap(),
        serde_json::Value::Null
    );
}

#[test]
fn building_a_page_programmatically_round_trips() {
    let mut page = Page::default();
    let section = page.add_section();
    page.sections[section].add_control(ColumnControl::Text(
        forge_m365_sp_clientside_pages_test_support::text("hello world"),
    ));
    page.sections[section].add_control(ColumnControl::WebPart(Box::new(WebPartData::new(
        "cccccccc-cccc-cccc-cccc-cccccccccccc",
        "My Part",
        "desc",
        json!({"foo": "bar"}),
    ))));

    let json = page.to_canvas_content_json().unwrap();
    let reparsed = Page::from_canvas_content_json(&json).unwrap();

    assert_eq!(reparsed.sections.len(), 1);
    assert_eq!(reparsed.sections[0].columns[0].controls.len(), 2);
}

#[test]
fn layout_webparts_content_round_trips_null_and_present() {
    assert!(parse_layout_webparts_content("null").unwrap().is_none());
    assert_eq!(to_layout_webparts_content_json(None).unwrap(), "null");
}

#[test]
fn special_save_handling_only_touches_webpart_controls() {
    let mut page = Page::default();
    let section = page.add_section();
    let mut wp = WebPartData::new("id", "t", "d", json!({}));
    let mut spc = ServerProcessedContent::default();
    let mut texts = std::collections::HashMap::new();
    texts.insert("k".to_string(), "a<b".to_string());
    spc.searchable_plain_texts = Some(texts);
    wp.web_part_data.server_processed_content = Some(spc);
    page.sections[section].add_control(ColumnControl::WebPart(Box::new(wp)));

    let json = page.to_canvas_content_json().unwrap();
    assert!(json.contains("a&lt;b"));
    assert!(!json.contains("a<b"));
}

#[test]
fn new_guid_has_v4_shape() {
    let g = new_guid();
    let parts: Vec<&str> = g.split('-').collect();
    assert_eq!(parts.len(), 5);
    assert_eq!(parts[0].len(), 8);
    assert_eq!(parts[1].len(), 4);
    assert_eq!(parts[2].len(), 4);
    assert!(parts[2].starts_with('4'));
    assert_eq!(parts[3].len(), 4);
    assert!(matches!(
        parts[3].chars().next().unwrap(),
        '8' | '9' | 'a' | 'b'
    ));
    assert_eq!(parts[4].len(), 12);
}

/// Small local helper module (not part of the crate's public API) so the
/// reindex/round-trip tests above don't need to hand-roll a `TextData`
/// literal every time.
mod forge_m365_sp_clientside_pages_test_support {
    pub fn text(s: &str) -> forge_m365_sp_clientside_pages::TextData {
        forge_m365_sp_clientside_pages::TextData::new(s)
    }
}
