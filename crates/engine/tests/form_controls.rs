//! Original native worksheet-control integration cases (no Office-created fixtures).
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
    model::{AnchorMode, FormControlKind, form_controls::effective_control},
};
use serde_json::{Value as Json, json};
fn session() -> Session {
    let mut s = Session::new();
    s.new_workbook();
    s
}
fn add(s: &mut Session, kind: i32, name: &str) -> Json {
    s.execute("control.add", json!({"type":kind,"name":name,"left":140,"top":80,"width":160,"height":60})).unwrap()["shape"].clone()
}
fn state(s: &Session, id: &Json) -> gridcraft_engine::model::form_controls::FormControl {
    let d = s.doc().unwrap();
    let sh = d.wb.sheet(0).unwrap();
    let shape = sh.shapes.iter().find(|sh| json!(sh.id) == *id).unwrap();
    effective_control(&d.wb, 0, shape).unwrap()
}
fn value(s: &Session, cell: &str) -> Value {
    s.doc().unwrap().wb.sheet(0).unwrap().value(CellRef::parse(cell).unwrap())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn all_nine_native_families_save_reopen_edit_and_delete_without_duplicate_shapes() {
    let mut s = session();
    let kinds = [0, 1, 2, 4, 5, 6, 7, 8, 9];
    let ids = kinds.iter().map(|k| add(&mut s, *k, &format!("Control{k}"))).collect::<Vec<_>>();
    for (k, id) in kinds.iter().zip(&ids) {
        if [2, 6].contains(k) {
            s.execute("control.set", json!({"id":id,"items":["One & <two>","漢字"],"value":2})).unwrap();
        }
        if [8, 9].contains(k) {
            s.execute("control.set", json!({"id":id,"min":2,"max":40,"value":7,"smallChange":3,"largeChange":9})).unwrap();
        }
        s.execute("control.set", json!({"id":id,"caption":"A & <B> \"caption\""})).unwrap();
        s.execute("shape.setMacro", json!({"id":id,"entry":"[0]!Actions.Go"})).unwrap();
    }
    let mut wb = (*s.doc().unwrap().wb).clone();
    for pass in 0..3 {
        let (back, warnings) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(&wb).unwrap()).unwrap();
        assert!(warnings.warnings.is_empty(), "{pass}: {warnings:?}");
        let sh = back.sheet(0).unwrap();
        assert_eq!(sh.shapes.len(), 9);
        for (k, shape) in kinds.iter().zip(&sh.shapes) {
            let c = shape.format.form_control.as_ref().unwrap();
            assert_eq!(c.kind, FormControlKind::from_excel(*k).unwrap());
            assert_eq!(shape.text, if pass == 0 { "A & <B> \"caption\"" } else { "Edited native caption" });
            assert_eq!(shape.on_action.as_deref(), Some("[0]!Actions.Go"));
            assert!(shape.format.office_id.is_some());
            if [2, 6].contains(k) {
                assert_eq!(c.items, ["One & <two>", "漢字"]);
                assert_eq!(c.value, 2);
            }
        }
        wb = back;
        if pass == 0 {
            for shape in &mut wb.sheet_mut(0).unwrap().shapes {
                shape.text = "Edited native caption".into();
                shape.anchor.width += 15.;
                shape.format.form_control.as_mut().unwrap().print_object = false;
            }
        }
    }
    wb.sheet_mut(0).unwrap().shapes.clear();
    let (back, _) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(&wb).unwrap()).unwrap();
    assert!(back.sheet(0).unwrap().shapes.is_empty());
    if let Some(p) = back.sheet(0).unwrap().preserved.as_ref() {
        assert!(!p.elements.iter().any(|e| e.slot == "controls"));
        assert!(!p.parts.iter().any(|p| p.content_type == "application/vnd.ms-excel.controlproperties+xml"));
    }
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn linked_cells_and_list_ranges_are_live_recalculate_and_undo_in_one_command() {
    let mut s = session();
    let check = add(&mut s, 1, "Agree");
    let list = add(&mut s, 2, "Choices");
    s.execute("range.setValues", json!({"range":"A1:A3","values":[["Alpha"],["Beta"],["Gamma"]]})).unwrap();
    s.execute("control.set", json!({"id":list,"listFillRange":"$A$1:$A$3","linkedCell":"D1"})).unwrap();
    s.execute("cell.set", json!({"cell":"E1","input":"=D1*10"})).unwrap();
    s.execute("control.activate", json!({"id":list,"item":2})).unwrap();
    assert_eq!(value(&s, "D1"), Value::Number(2.));
    assert_eq!(value(&s, "E1"), Value::Number(20.));
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(value(&s, "D1"), Value::Empty);
    assert_eq!(state(&s, &list).value, 0);
    s.execute("edit.redo", json!({})).unwrap();
    assert_eq!(state(&s, &list).value, 2);
    s.execute("cell.set", json!({"cell":"A2","input":"Changed"})).unwrap();
    assert_eq!(state(&s, &list).items[1], "Changed");
    s.execute("cell.set", json!({"cell":"D1","input":"3"})).unwrap();
    assert_eq!(state(&s, &list).value, 3);
    s.execute("control.set", json!({"id":check,"linkedCell":"C1"})).unwrap();
    s.execute("control.activate", json!({"id":check})).unwrap();
    assert_eq!(value(&s, "C1"), Value::Bool(true));
    s.execute("control.set", json!({"id":check,"value":2})).unwrap();
    assert_eq!(state(&s, &check).value, 2);
    let (back, _) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(&s.doc().unwrap().wb).unwrap()).unwrap();
    let c = effective_control(&back, 0, &back.sheet(0).unwrap().shapes[1]).unwrap();
    assert_eq!(c.value, 3);
    assert_eq!(c.items[1], "Changed");
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn radio_groups_and_multiselection_do_not_leak_state() {
    let mut s = session();
    let a = add(&mut s, 7, "A");
    let b = add(&mut s, 7, "B");
    for id in [&a, &b] {
        s.execute("control.set", json!({"id":id,"linkedCell":"F1"})).unwrap();
    }
    s.execute("control.activate", json!({"id":b})).unwrap();
    assert_eq!(value(&s, "F1"), Value::Number(2.));
    assert_eq!(state(&s, &a).value, -4146);
    assert_eq!(state(&s, &b).value, 1);
    s.execute("control.activate", json!({"id":a})).unwrap();
    assert_eq!(state(&s, &b).value, -4146);
    let list = add(&mut s, 6, "Multi");
    s.execute("control.set", json!({"id":list,"items":["A","B","C"],"multiSelect":true,"extendedSelect":true})).unwrap();
    for n in [1, 3, 1] {
        s.execute("control.activate", json!({"id":list,"item":n})).unwrap();
    }
    assert_eq!(state(&s, &list).selected, [3]);
    assert!(s.execute("control.set", json!({"id":list,"linkedCell":"F1"})).is_err());
    let (back, _) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(&s.doc().unwrap().wb).unwrap()).unwrap();
    let c = back.sheet(0).unwrap().shapes[2].format.form_control.as_ref().unwrap();
    assert!(c.extended_select);
    assert_eq!(c.selected, [3]);
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn invalid_disabled_protected_and_oversized_actions_are_atomic() {
    let mut s = session();
    let button = add(&mut s, 0, "Run");
    let list = add(&mut s, 6, "List");
    for p in [
        json!({"id":button,"min":1}),
        json!({"id":list,"items":["x".repeat(16385)]}),
        json!({"id":list,"selected":[0]}),
        json!({"id":list,"linkedCell":"[other]Sheet1!A1"}),
        json!({"id":button,"name":""}),
    ] {
        let before = s.doc().unwrap().wb.clone();
        assert!(s.execute("control.set", p).is_err());
        assert_eq!(s.doc().unwrap().wb, before);
    }
    s.execute("control.set", json!({"id":button,"enabled":false})).unwrap();
    assert!(s.execute("control.activate", json!({"id":button})).is_err());
    s.execute("review.protectSheet", json!({})).unwrap();
    let before = s.doc().unwrap().wb.clone();
    assert!(s.execute("control.activate", json!({"id":list})).is_err());
    for command in ["object.move", "object.delete", "object.setAnchorMode", "arrange.bringForward"] {
        assert!(s.execute(command, json!({"kind":"shape", "id":list, "width":200, "mode":"absolute"})).is_err(), "{command}");
    }
    assert_eq!(s.doc().unwrap().wb, before);
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn placement_tracks_sizes_and_hiding_without_losing_unhidden_dimensions() {
    let mut s = session();
    let id = add(&mut s, 0, "Sized");
    let wb = std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb);
    let sh = wb.sheet_mut(0).unwrap();
    let shape = sh.shapes.iter_mut().find(|s| json!(s.id) == id).unwrap();
    shape.anchor.mode = AnchorMode::MoveAndSize;
    let original = shape.anchor;
    s.execute("home.rowHeight", json!({"rows":"5:5","height":40})).unwrap();
    let current = s.doc().unwrap().wb.sheet(0).unwrap().shapes[0].anchor;
    assert!(current.height > original.height);
    s.execute("home.hideRows", json!({"rows":"5:7"})).unwrap();
    s.execute("home.unhideRows", json!({"rows":"5:7"})).unwrap();
    assert_eq!(s.doc().unwrap().wb.sheet(0).unwrap().shapes[0].anchor, current);
    s.execute("edit.undo", json!({})).unwrap();
    s.execute("edit.undo", json!({})).unwrap();
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(s.doc().unwrap().wb.sheet(0).unwrap().shapes[0].anchor, original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unchanged_linked_control_parts_are_byte_preserved_and_names_follow_document_edits() {
    let mut s = session();
    s.execute("range.setValues", json!({"range":"A1:A2","values":[["One"],["Two"]]})).unwrap();
    let list = add(&mut s, 6, "NativeList");
    s.execute("control.set", json!({"id":list,"listFillRange":"Sheet1!$A$1:$A$2","linkedCell":"Sheet1!$B$1","value":1})).unwrap();
    let (wb, _) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(&s.doc().unwrap().wb).unwrap()).unwrap();
    let original = wb.sheet(0).unwrap().preserved.as_ref().unwrap();
    let (back, _) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(&wb).unwrap()).unwrap();
    let kept = back.sheet(0).unwrap().preserved.as_ref().unwrap();
    assert_eq!(original.parts, kept.parts);
    assert_eq!(original.vml, kept.vml);
    s.execute("sheet.rename", json!({"name":"Data renamed"})).unwrap();
    assert!(state(&s, &list).list_fill_range.contains("Data renamed"));
    assert_eq!(state(&s, &list).items, ["One", "Two"]);
    s.execute("home.insertRows", json!({"rows":"1:1"})).unwrap();
    assert!(state(&s, &list).list_fill_range.contains("$A$2:$A$3"));
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_ids_and_package_parts_remain_sheet_local() {
    let mut s = session();
    let first = add(&mut s, 0, "First");
    s.execute("shape.setMacro", json!({"id":first,"entry":"FirstAction"})).unwrap();
    s.execute("home.insertSheet", json!({"name":"Other"})).unwrap();
    let second = s.execute("control.add", json!({"type":0,"name":"Second"})).unwrap()["shape"].clone();
    s.execute("shape.setMacro", json!({"id":second,"entry":"SecondAction"})).unwrap();
    let wb = &s.doc().unwrap().wb;
    assert_eq!(wb.sheet(0).unwrap().shapes[0].format.office_id, wb.sheet(1).unwrap().shapes[0].format.office_id);
    let (back, report) = gridcraft_xlsx::read_xlsx(&gridcraft_xlsx::write_xlsx(wb).unwrap()).unwrap();
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    for (i, name, entry) in [(0, "First", "FirstAction"), (1, "Second", "SecondAction")] {
        let sh = back.sheet(i).unwrap();
        assert_eq!(sh.shapes.len(), 1);
        assert_eq!(sh.shapes[0].format.form_control.as_ref().unwrap().name, name);
        assert_eq!(sh.shapes[0].on_action.as_deref(), Some(entry));
    }
}
