//! Authored recorded-macro sequences against the real engine clipboard.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn setup(source: &str) -> R<(Session, Bridge, Ide)> {
    let mut session = Session::new();
    session.new_workbook();
    session.execute("sheet.rename", json!({"sheet":0,"name":"Menu"}))?;
    session.execute("home.insertSheet", json!({"name":"Ingredients"}))?;
    session.execute("sheet.activate", json!({"sheet":0}))?;
    let mut bridge = Bridge::new(&session)?;
    bridge.consent(true);
    let mut ide = Ide::new();
    ide.add_module("Copy.bas", source)?;
    bridge.configure(&mut ide)?;
    Ok((session, bridge, ide))
}
fn run(session: &mut Session, bridge: &mut Bridge, ide: &mut Ide) -> R {
    bridge.dispatch(session, ide, Action::Run)?;
    for _ in 0..2000 {
        bridge.tick(session, ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
fn value(session: &Session, sheet: usize, row: u32, col: u32) -> R<Value> {
    Ok(session.doc()?.wb.sheet(sheet).ok_or("sheet missing")?.value(CellRef::new(row, col)))
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn recorded_ingredient_copy_loop_preserves_values_formulas_formats_and_one_undo() -> R {
    let source = r#"Sub Main()
For Index = 1 To 4
    For d = 1 To 5
        Dim meal
        meal = Sheets("Menu").Cells(d + 2, Index + 1)
        If meal <> Empty Then
            For t = 1 To 40
                Dim check
                check = Sheets("Ingredients").Cells(1, t)
                If StrComp(check, meal) = 0 Then
                    Application.Goto (ActiveWorkbook.Sheets("Ingredients").Range(Sheets("Ingredients").Cells(2, t), Sheets("Ingredients").Cells(4, t)))
                    Selection.Copy
                    Sheets("Menu").Select
                    ActiveSheet.Cells(d * 15 - 4, Index + 1).Select
                    ActiveSheet.Paste
                End If
            Next
        End If
    Next
Next
Sheets("Menu").Select
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    for index in 1..=4 {
        for d in 1..=5 {
            let t = (index - 1) * 5 + d;
            let meal = format!("meal {t}");
            s.execute("cell.set", json!({"sheet":0,"cell":CellRef::new(d+1,index).a1(),"input":meal}))?;
            for (row, input) in [(0, meal), (1, format!("ingredient {t}")), (2, t.to_string()), (3, format!("={}*2", CellRef::new(2, t - 1).a1()))] {
                s.execute("cell.set", json!({"sheet":1,"cell":CellRef::new(row,t-1).a1(),"input":input}))?;
            }
        }
    }
    s.execute("home.bold", json!({"sheet":1,"range":"A2:T4","on":true}))?;
    run(&mut s, &mut b, &mut ide)?;
    for index in 1..=4 {
        for d in 1..=5 {
            let t = (index - 1) * 5 + d;
            let row = d * 15 - 5;
            assert_eq!(value(&s, 0, row, index)?, Value::Text(format!("ingredient {t}").into()));
            assert_eq!(value(&s, 0, row + 1, index)?, Value::Number(f64::from(t)));
            assert_eq!(value(&s, 0, row + 2, index)?, Value::Number(f64::from(t * 2)));
            let at = CellRef::new(row + 2, index);
            let sheet = s.doc()?.wb.sheet(0).ok_or("menu")?;
            assert_eq!(
                sheet.cell(at).and_then(|c| c.formula.as_ref()).map(|f| f.text.as_str()),
                Some(format!("{}*2", CellRef::new(row + 1, index).a1()).as_str())
            );
            assert!(s.doc()?.wb.styles.get(sheet.style_id(at)).font.bold);
        }
    }
    assert_eq!(s.doc()?.sheet_index(), 0);
    // The bounded undo stack is full after populating all 100 fixture cells;
    // one undo must still remove the complete macro's writes as one group.
    s.execute("edit.undo", json!({}))?;
    for index in 1..=4 {
        for d in 1..=5 {
            let row = d * 15 - 5;
            assert_eq!(value(&s, 0, row, index)?, Value::Empty);
            assert_eq!(value(&s, 0, row + 1, index)?, Value::Empty);
            assert_eq!(value(&s, 0, row + 2, index)?, Value::Empty);
        }
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn clipboard_refusals_preserve_cells_selection_and_previous_copy() -> R {
    for (statement, number) in [
        ("Sheets(\"Ingredients\").Range(\"A1\").Select", 1004.),
        ("Sheets(\"Ingredients\").Paste", 1004.),
        ("Range(\"A1\").Select False", 450.),
        ("Columns(1).Copy", 7.),
        ("Range(\"A1,C1\").Copy", 1004.),
        ("Range(\"A1:A2\").Cut\nActiveSheet.Paste Destination:=Range(\"A2\")", 1004.),
        ("ActiveSheet.Paste Destination:=1", 13.),
        ("ActiveSheet.Paste Destination:=Range(\"D2\"), Link:=True", 1004.),
        ("ActiveSheet.PasteSpecial Format:=\"Unicode Text\"", 1004.),
        ("Range(\"A1\").CopyPicture", 1004.),
        ("ActiveSheet.Paste Unknown:=True", 448.),
        ("Range(\"XFD1048576\").Select\nActiveSheet.Paste", 1004.),
        ("Range(\"A1:A65538\").Select\nActiveSheet.Paste", 7.),
    ] {
        let source = format!("Sub Main()\nRange(\"A1:A2\").Copy\nOn Error Resume Next\n{statement}\nRange(\"D1\") = Err.Number\nEnd Sub");
        let (mut s, mut b, mut ide) = setup(&source)?;
        s.execute("cell.set", json!({"cell":"A1","input":"original"}))?;
        run(&mut s, &mut b, &mut ide)?;
        assert_eq!(value(&s, 0, 0, 3)?, Value::Number(number), "{statement}");
        assert_eq!(value(&s, 0, 0, 0)?, Value::Text("original".into()));
        assert_eq!(value(&s, 0, 0, 1)?, Value::Empty);
        assert_eq!(s.doc()?.sheet_index(), 0);
        assert_eq!(s.clipboard.as_ref().ok_or("copy")?.range.a1(), "A1:A2");
        if statement.contains(".Cut") {
            assert!(s.clipboard.as_ref().is_some_and(|clipboard| clipboard.cut));
        }
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_copy_destination_pastes_a_bounded_snapshot_to_another_sheet() -> R {
    let source = r#"Sub Main()
Range("A1:A2").Copy Destination:=Sheets("Ingredients").Range("D2")
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    s.execute("cell.set", json!({"cell":"A2","input":"=A1+6"}))?;
    s.execute("home.bold", json!({"range":"A1:A2","on":true}))?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(1).ok_or("ingredients")?.code_name = Some("IngredientsCode".into());
    ide.add_document_module(
        "Ingredients.cls",
        "IngredientsCode",
        "Attribute VB_Name = \"IngredientsCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nMe.Range(\"Z1\") = Target.Address\nEnd Sub\n",
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, 1, 1, 3)?, Value::Number(17.));
    assert_eq!(value(&s, 1, 2, 3)?, Value::Number(23.));
    assert_eq!(
        s.doc()?.wb.sheet(1).ok_or("ingredients")?.cell(CellRef::new(2, 3)).and_then(|cell| cell.formula.as_ref()).map(|f| f.text.as_str()),
        Some("D2+6")
    );
    assert!(s.doc()?.wb.styles.get(s.doc()?.wb.sheet(1).ok_or("ingredients")?.style_id(CellRef::new(1, 3))).font.bold);
    assert_eq!(value(&s, 1, 0, 25)?, Value::Text("$D$2:$D$3".into()));
    assert_eq!(s.doc()?.sheet_index(), 1);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_paste_link_creates_live_same_workbook_formulas() -> R {
    let source = "Sub Main()\nRange(\"A1:A2\").Copy\nRange(\"C1:C2\").Select\nActiveSheet.Paste Link:=True\nEnd Sub";
    let (mut session, mut bridge, mut ide) = setup(source)?;
    session.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    session.execute("cell.set", json!({"cell":"A2","input":"=A1+6"}))?;
    run(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(value(&session, 0, 0, 2)?, Value::Number(17.));
    assert_eq!(value(&session, 0, 1, 2)?, Value::Number(23.));
    let sheet = session.doc()?.wb.sheet(0).ok_or("menu")?;
    assert_eq!(sheet.cell(CellRef::new(0, 2)).and_then(|cell| cell.formula.as_ref()).map(|formula| formula.text.as_str()), Some("Menu!A1"));
    assert_eq!(sheet.cell(CellRef::new(1, 2)).and_then(|cell| cell.formula.as_ref()).map(|formula| formula.text.as_str()), Some("Menu!A2"));
    session.execute("cell.set", json!({"cell":"A1","input":"19"}))?;
    assert_eq!(value(&session, 0, 0, 2)?, Value::Number(19.));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_paste_destination_uses_the_range_and_changes_active_sheet() -> R {
    let source = r#"Sub Main()
Range("A1:A2").Copy
ActiveSheet.Paste Destination:=Sheets("Ingredients").Range("D2")
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"first"}))?;
    s.execute("cell.set", json!({"cell":"A2","input":"second"}))?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(1).ok_or("ingredients")?.code_name = Some("IngredientsCode".into());
    ide.add_document_module(
        "Ingredients.cls",
        "IngredientsCode",
        "Attribute VB_Name = \"IngredientsCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nMe.Range(\"Z1\") = Target.Address\nEnd Sub\n",
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, 1, 1, 3)?, Value::Text("first".into()));
    assert_eq!(value(&s, 1, 2, 3)?, Value::Text("second".into()));
    assert_eq!(value(&s, 1, 0, 25)?, Value::Text("$D$2:$D$3".into()));
    assert_eq!(s.doc()?.sheet_index(), 1);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_paste_special_binds_excel_enum_values_and_operations() -> R {
    let source = r#"Sub Main()
Range("A1").Copy
Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationMultiply
Range("A2").Copy
Range("B2").PasteSpecial xlPasteValues
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"3"}))?;
    s.execute("cell.set", json!({"cell":"B1","input":"4"}))?;
    s.execute("cell.set", json!({"cell":"A2","input":"=1+2"}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, 0, 0, 1)?, Value::Number(12.));
    assert_eq!(value(&s, 0, 1, 1)?, Value::Number(3.));
    assert!(s.doc()?.wb.sheet(0).ok_or("menu")?.cell(CellRef::new(1, 1)).and_then(|cell| cell.formula.as_ref()).is_none());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_paste_special_refuses_unimplemented_modes_before_mutation() -> R {
    let source = r#"Sub Main()
Range("A1").Copy
On Error Resume Next
Range("B1").PasteSpecial Paste:=xlPasteAllUsingSourceTheme
Range("C1") = Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"source"}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, 0, 0, 1)?, Value::Empty);
    assert_eq!(value(&s, 0, 0, 2)?, Value::Number(1004.));
    assert_eq!(s.clipboard.as_ref().ok_or("clipboard")?.range.a1(), "A1");
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_paste_special_transposes_and_skips_blank_source_cells() -> R {
    let source = r#"Sub Main()
Range("A1:B2").Copy
Range("D1:E2").PasteSpecial Paste:=xlPasteValues, SkipBlanks:=True, Transpose:=True
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    for (cell, input) in [("A1", "1"), ("A2", "3"), ("B2", "4"), ("D1", "9"), ("E1", "7"), ("D2", "8"), ("E2", "6")] {
        s.execute("cell.set", json!({"cell":cell,"input":input}))?;
    }
    run(&mut s, &mut b, &mut ide)?;
    for (cell, expected) in [(CellRef::new(0, 3), 1.), (CellRef::new(0, 4), 3.), (CellRef::new(1, 3), 8.), (CellRef::new(1, 4), 4.)] {
        assert_eq!(s.doc()?.wb.sheet(0).ok_or("menu")?.value(cell), Value::Number(expected));
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_cut_copy_mode_reports_mode_and_cancels_internal_clipboard() -> R {
    let source = r#"Sub Main()
Range("A1").Copy
Range("B1") = Application.CutCopyMode
Application.CutCopyMode = xlCut
Range("C1") = Application.CutCopyMode
Application.CutCopyMode = False
Range("D1") = Application.CutCopyMode
Application.CutCopyMode = True
Range("E1") = Application.CutCopyMode
Application.CutCopyMode = False
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"copy me"}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, 0, 0, 1)?, Value::Number(1.));
    assert_eq!(value(&s, 0, 0, 2)?, Value::Number(2.));
    assert_eq!(value(&s, 0, 0, 3)?, Value::Bool(false));
    assert_eq!(value(&s, 0, 0, 4)?, Value::Number(1.));
    assert!(s.clipboard.is_none());
    assert_eq!(value(&s, 0, 0, 0)?, Value::Text("copy me".into()));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_cut_and_worksheet_paste_move_the_source_atomically() -> R {
    let source = r#"Sub Main()
Range("A1:A2").Cut
ActiveSheet.Paste Destination:=Range("D2")
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    s.execute("cell.set", json!({"cell":"A2","input":"=A1+6"}))?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("menu")?.code_name = Some("MenuCode".into());
    ide.add_document_module(
        "Menu.cls",
        "MenuCode",
        "Attribute VB_Name = \"MenuCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nMe.Range(\"Z1\") = Target.Address\nEnd Sub\n",
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    let history = s.doc()?.undo.len();
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, 0, 0, 0)?, Value::Empty);
    assert_eq!(value(&s, 0, 1, 0)?, Value::Empty);
    assert_eq!(value(&s, 0, 1, 3)?, Value::Number(17.));
    assert_eq!(value(&s, 0, 2, 3)?, Value::Number(23.));
    assert_eq!(
        s.doc()?.wb.sheet(0).ok_or("menu")?.cell(CellRef::new(2, 3)).and_then(|cell| cell.formula.as_ref()).map(|f| f.text.as_str()),
        Some("D2+6")
    );
    assert_eq!(value(&s, 0, 0, 25)?, Value::Text("$A$1:$A$2,$D$2:$D$3".into()));
    assert_eq!(s.doc()?.undo.len(), history + 1);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_cut_destination_moves_cells_and_queues_a_change() -> R {
    let (mut s, mut b, mut ide) = setup("Sub Main()\nRange(\"A1:A2\").Cut Destination:=Range(\"D2\")\nEnd Sub")?;
    s.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    s.execute("cell.set", json!({"cell":"A2","input":"23"}))?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("menu")?.code_name = Some("MenuCode".into());
    ide.add_document_module(
        "Menu.cls",
        "MenuCode",
        "Attribute VB_Name = \"MenuCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Address\nEnd Sub\n",
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, 0, 0, 0)?, Value::Empty);
    assert_eq!(value(&s, 0, 1, 3)?, Value::Number(17.));
    assert_eq!(value(&s, 0, 2, 3)?, Value::Number(23.));
    assert_eq!(ide.output(), &["$A$1:$A$2,$D$2:$D$3"]);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn cross_sheet_cut_paste_queues_source_then_destination_change() -> R {
    let source = r#"Sub Main()
Range("A1:A2").Cut
ActiveSheet.Paste Destination:=Sheets("Ingredients").Range("D2")
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    s.execute("cell.set", json!({"cell":"A2","input":"23"}))?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("menu")?.code_name = Some("MenuCode".into());
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(1).ok_or("ingredients")?.code_name = Some("IngredientsCode".into());
    ide.add_document_module(
        "Menu.cls",
        "MenuCode",
        "Attribute VB_Name = \"MenuCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Address\nEnd Sub\n",
    )?;
    ide.add_document_module(
        "Ingredients.cls",
        "IngredientsCode",
        "Attribute VB_Name = \"IngredientsCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nDebug.Print Target.Address\nEnd Sub\n",
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    run(&mut s, &mut b, &mut ide)?;
    // A completed event handler can leave another queued sheet event; keep ticking until drained.
    for _ in 0..4 {
        if ide.output().len() == 2 {
            break;
        }
        b.tick(&mut s, &mut ide)?;
    }
    assert_eq!(value(&s, 0, 0, 0)?, Value::Empty);
    assert_eq!(value(&s, 1, 1, 3)?, Value::Number(17.));
    assert_eq!(value(&s, 1, 2, 3)?, Value::Number(23.));
    assert_eq!(ide.output(), &["$A$1:$A$2", "$D$2:$D$3"]);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn paste_obeys_protection_and_empty_clipboard_errors() -> R {
    for protected in [false, true] {
        let copy = if protected { "Range(\"A1\").Copy" } else { "" };
        let source = format!(
            "Sub Main()\n{copy}\nRange(\"B2\").Select\nOn Error Resume Next\nActiveSheet.Paste\nSheets(\"Ingredients\").Range(\"D1\") = Err.Number\nEnd Sub"
        );
        let (mut s, mut b, mut ide) = setup(&source)?;
        s.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
        if protected {
            s.execute("review.protectSheet", json!({}))?;
        }
        run(&mut s, &mut b, &mut ide)?;
        assert_eq!(value(&s, 1, 0, 3)?, Value::Number(1004.));
        assert_eq!(value(&s, 0, 1, 1)?, Value::Empty);
        assert_eq!(value(&s, 0, 0, 0)?, Value::Number(17.));
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn tiled_paste_delivers_one_worksheet_change_with_the_full_footprint() -> R {
    let (mut s, mut b, mut ide) = setup("Sub Main()\nRange(\"A1:A2\").Copy\nRange(\"B3:B6\").Select\nActiveSheet.Paste Link:=False\nEnd Sub")?;
    s.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    s.execute("cell.set", json!({"cell":"A2","input":"23"}))?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("menu")?.code_name = Some("MenuCode".into());
    ide.add_document_module("Menu.cls","MenuCode", "Attribute VB_Name = \"MenuCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nApplication.EnableEvents = False\nMe.Range(\"D1\") = Target.Address\nMe.Range(\"D2\") = Me.Range(\"D2\") + 1\nEnd Sub\n")?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    run(&mut s, &mut b, &mut ide)?;
    for (row, n) in [(2, 17.), (3, 23.), (4, 17.), (5, 23.)] {
        assert_eq!(value(&s, 0, row, 1)?, Value::Number(n));
    }
    assert_eq!(value(&s, 0, 0, 3)?, Value::Text("$B$3:$B$6".into()));
    assert_eq!(value(&s, 0, 1, 3)?, Value::Number(1.));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn pause_before_paste_retains_the_copied_snapshot() -> R {
    let source = "Sub Main()\nRange(\"A1\").Copy\nRange(\"B1\").Select\nActiveSheet.Paste\nEnd Sub";
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"17"}))?;
    ide.toggle_breakpoint_at(source.find("ActiveSheet.Paste").ok_or("paste")?)?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..50 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(value(&s, 0, 0, 1)?, Value::Empty);
    s.execute("cell.set", json!({"cell":"A1","input":"99"}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, 0, 0, 1)?, Value::Number(17.));
    Ok(())
}
