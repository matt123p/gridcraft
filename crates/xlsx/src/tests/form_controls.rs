//! Original M7A worksheet Button packages. No Office binary or downloaded macro is executed.
use super::{ROOT_RELS, make_zip};
use crate::{IoError, read_worksheet_controls, read_xlsx, write_xlsx};
use gridcraft_model::{Preserved, Sheet};
use std::sync::Arc;
const SHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PROPS: &str = r#"<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Button"/>"#;
const VML: &str = r#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="Button 10" o:spid="_x0000_s1034" style="position:absolute;visibility:visible"><v:textbox><div>Copy &amp; <b>ingredients</b>&#33;</div></v:textbox><x:ClientData ObjectType="Button"><x:Anchor>1, 0, 2, 0, 3, 0, 4, 0</x:Anchor><x:FmlaMacro>[0]!AuthoredButton</x:FmlaMacro></x:ClientData></v:shape></xml>"#;
fn control(macro_name: &str) -> String {
    format!(
        r#"<control shapeId="1034" r:id="rId2" name="Button 10"><controlPr macro="{macro_name}"><anchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" moveWithCells="1"><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control>"#
    )
}
fn package(control_xml: &str, vml: &str) -> Vec<u8> {
    let body = format!(r#"<sheetData/><legacyDrawing r:id="rId1"/>{control_xml}"#);
    let rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL_NS}/vmlDrawing" Target="../drawings/button.vml"/><Relationship Id="rId2" Type="{REL_NS}/ctrlProp" Target="../ctrlProps/button.xml"/></Relationships>"#
    );
    let ct = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/ctrlProps/button.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/></Types>"#;
    let workbook =
        format!(r#"<workbook xmlns="{SHEET_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#);
    let workbook_rels = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
    );
    let worksheet = format!(r#"<worksheet xmlns="{SHEET_NS}" xmlns:r="{REL_NS}">{body}</worksheet>"#);
    make_zip(&[
        ("[Content_Types].xml", ct.as_bytes()),
        ("_rels/.rels", ROOT_RELS.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", worksheet.as_bytes()),
        ("xl/worksheets/_rels/sheet1.xml.rels", rels.as_bytes()),
        ("xl/drawings/button.vml", vml.as_bytes()),
        ("xl/ctrlProps/button.xml", PROPS.as_bytes()),
    ])
}
fn sheet() -> Sheet {
    read_xlsx(&package(&format!("<controls>{}</controls>", control("")), VML)).unwrap().0.sheet(0).unwrap().clone()
}
fn mutate(sheet: &mut Sheet, change: impl FnOnce(&mut Preserved)) {
    change(Arc::make_mut(sheet.preserved.as_mut().unwrap()));
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_button_identity_caption_and_assignment_decode_without_effects() {
    let s = sheet();
    let original = s.preserved.clone();
    let data = read_worksheet_controls(&s).unwrap();
    assert!(data.unsupported.is_empty());
    assert_eq!(data.buttons.len(), 1);
    let button = &data.buttons[0];
    assert_eq!(button.shape_id, 1034);
    assert_eq!(button.name, "Button 10");
    assert_eq!(button.caption_text, "Copy & ingredients!");
    assert_eq!(button.macro_assignment.as_deref(), Some("[0]!AuthoredButton"));
    assert_eq!(button.vml_anchor.as_deref(), Some("1, 0, 2, 0, 3, 0, 4, 0"));
    assert_eq!(button.vml_style.as_deref(), Some("position:absolute;visibility:visible"));
    assert_eq!(s.preserved, original, "decoding is inert and read-only");
    assert!(s.shapes.is_empty(), "metadata is not a duplicate rendered shape");
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_button_metadata_survives_sheet_rename_and_original_package_save() {
    let input = package(&format!("<controls>{}</controls>", control("[0]!AuthoredButton")), VML);
    let (mut wb, _) = read_xlsx(&input).unwrap();
    let expected = read_worksheet_controls(wb.sheet(0).unwrap()).unwrap();
    wb.sheet_mut(0).unwrap().name = "Renamed".into();
    let (back, _) = read_xlsx(&write_xlsx(&wb).unwrap()).unwrap();
    assert_eq!(back.sheet(0).unwrap().name, "Renamed");
    assert_eq!(read_worksheet_controls(back.sheet(0).unwrap()).unwrap(), expected);
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn markup_compatibility_selects_one_known_choice_or_fallback_without_duplicate_controls() {
    for (prefix, uri) in [("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"), ("future", "urn:future")] {
        let c = control("");
        let controls = format!(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:{prefix}="{uri}"><mc:Choice Requires="{prefix}"><controls>{c}</controls></mc:Choice><mc:Fallback><controls>{c}</controls></mc:Fallback></mc:AlternateContent>"#
        );
        let wb = read_xlsx(&package(&controls, VML)).unwrap().0;
        assert_eq!(read_worksheet_controls(wb.sheet(0).unwrap()).unwrap().buttons.len(), 1);
    }
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn macro_sources_must_agree_and_empty_legacy_source_does_not_mask_native_assignment() {
    let controls = format!("<controls>{}</controls>", control("[0]!Different"));
    let wb = read_xlsx(&package(&controls, VML)).unwrap().0;
    let s = wb.sheet(0).unwrap();
    let original = s.preserved.clone();
    assert!(matches!(read_worksheet_controls(s),Err(IoError::Format(s)) if s.contains("conflicting native macro")));
    assert_eq!(s.preserved, original);
    let empty = VML.replace("[0]!AuthoredButton", "");
    let wb = read_xlsx(&package(&controls, &empty)).unwrap().0;
    assert_eq!(read_worksheet_controls(wb.sheet(0).unwrap()).unwrap().buttons[0].macro_assignment.as_deref(), Some("[0]!Different"));
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn invalid_or_ambiguous_native_graph_refuses_without_mutating_preservation() {
    for case in 0..9 {
        let mut s = sheet();
        mutate(&mut s, |p| match case {
            0 => p.rels.iter_mut().find(|r| r.id == "rId2").unwrap().external = true,
            1 => p.rels.iter_mut().find(|r| r.id == "rId2").unwrap().rel_type = "urn:wrong/control".into(),
            2 => p.parts.clear(),
            3 => p.parts[0].content_type = "application/xml".into(),
            4 => p.rels.push(p.rels[0].clone()),
            5 => p.parts.push(p.parts[0].clone()),
            6 => p.vml.as_mut().unwrap().data = Arc::from(VML.replace("_x0000_s1034", "_x0000_s1035").into_bytes()),
            7 => p.vml.as_mut().unwrap().data = Arc::from(VML.replace("id=\"Button 10\"", "id=\"_x0000_s1035\"").into_bytes()),
            _ => {
                p.elements.iter_mut().find(|e| e.slot == "controls").unwrap().xml =
                    format!(r#"<controls xmlns="{SHEET_NS}" xmlns:r="{REL_NS}">{}{}</controls>"#, control(""), control(""))
            }
        });
        let original = s.preserved.clone();
        assert!(matches!(read_worksheet_controls(&s), Err(IoError::Format(_))), "case {case}");
        assert_eq!(s.preserved, original);
    }
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn namespace_aliases_are_supported_but_spoofed_native_elements_and_duplicate_attributes_refuse() {
    let alias = VML.replace("xmlns:v=", "xmlns:draw=").replace("v:", "draw:").replace("xmlns:x=", "xmlns:data=").replace("x:", "data:");
    let controls = format!("<controls>{}</controls>", control(""));
    let wb = read_xlsx(&package(&controls, &alias)).unwrap().0;
    assert_eq!(read_worksheet_controls(wb.sheet(0).unwrap()).unwrap().buttons[0].caption_text, "Copy & ingredients!");
    for vml in [
        VML.replace("urn:schemas-microsoft-com:office:excel", "urn:spoof"),
        VML.replace("<xml ", "<!DOCTYPE xml [<!ENTITY remote SYSTEM 'https://invalid.test/never-read'>]><xml "),
        VML.replace("</v:shape>", "</v:wrong>"),
        format!("{VML}<second/>"),
        VML.replace("ingredients", "&#0;"),
    ] {
        let mut s = sheet();
        mutate(&mut s, |p| p.vml.as_mut().unwrap().data = Arc::from(vml.into_bytes()));
        assert!(read_worksheet_controls(&s).is_err());
    }
    let mut s = sheet();
    mutate(&mut s, |p| {
        p.elements[0].xml = format!(
            r#"<controls xmlns="{SHEET_NS}" xmlns:r="{REL_NS}" xmlns:q="{REL_NS}"><control shapeId="1034" r:id="rId2" q:id="rId2" name="Button 10"/></controls>"#
        )
    });
    assert!(read_worksheet_controls(&s).is_err());
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn other_native_control_types_are_explicitly_reported_without_button_semantics() {
    let mut s = sheet();
    mutate(&mut s, |p| {
        p.parts[0].data = Arc::from(PROPS.replace("Button", "CheckBox").into_bytes());
        p.vml.as_mut().unwrap().data = Arc::from(VML.replace("ObjectType=\"Button\"", "ObjectType=\"Checkbox\"").into_bytes());
    });
    let data = read_worksheet_controls(&s).unwrap();
    assert!(data.buttons.is_empty());
    assert_eq!(data.unsupported[0].object_type, "CheckBox");
    assert_eq!(data.unsupported[0].shape_id, 1034);
    assert!(read_worksheet_controls(&Sheet::new("Empty")).unwrap().buttons.is_empty());
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_xml_budgets_bound_depth_count_fields_and_entity_fragmentation() {
    let deep = format!("<xml>{}{}</xml>", "<x>".repeat(33), "</x>".repeat(33));
    let nodes = format!("<xml>{}</xml>", "<x/>".repeat(8192));
    let field = VML.replace("ingredients", &"a".repeat(16385));
    let entities = VML.replace("ingredients", &"&#97;".repeat(20000));
    let input = format!("<xml>{}</xml>", " ".repeat(4 * 1024 * 1024));
    let shapes = VML.replace("</xml>", &format!("{}</xml>", (2000..3024).map(|id| format!("<v:shape id=\"_x0000_s{id}\"/>")).collect::<String>()));
    for vml in [deep, nodes, field, entities, input, shapes] {
        let mut s = sheet();
        mutate(&mut s, |p| p.vml.as_mut().unwrap().data = Arc::from(vml.into_bytes()));
        assert!(matches!(read_worksheet_controls(&s), Err(IoError::TooLarge(_))));
    }
}
