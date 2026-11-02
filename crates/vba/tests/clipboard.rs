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
            // Two populated slots exercise both ends of the recorded 4x5 loop;
            // filling every slot exceeds the shared IDE's 1024 host-call limit.
            if t == 1 || t == 20 {
                s.execute("cell.set", json!({"sheet":0,"cell":CellRef::new(d+1,index).a1(),"input":meal}))?;
            }
            for (row, input) in [(0, meal), (1, format!("ingredient {t}")), (2, t.to_string()), (3, format!("={}*2", CellRef::new(2, t - 1).a1()))] {
                s.execute("cell.set", json!({"sheet":1,"cell":CellRef::new(row,t-1).a1(),"input":input}))?;
            }
        }
    }
    s.execute("home.bold", json!({"sheet":1,"range":"A2:T4","on":true}))?;
    let history = s.doc()?.undo.len();
    run(&mut s, &mut b, &mut ide)?;
    for index in 1..=4 {
        for d in 1..=5 {
            let t = (index - 1) * 5 + d;
            let row = d * 15 - 5;
            if t != 1 && t != 20 {
                assert_eq!(value(&s, 0, row, index)?, Value::Empty);
                continue;
            }
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
    assert_eq!(s.doc()?.undo.len(), history + 1);
    s.execute("edit.undo", json!({}))?;
    assert_eq!(value(&s, 0, 10, 1)?, Value::Empty);
    assert_eq!(value(&s, 0, 72, 4)?, Value::Empty);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn clipboard_refusals_preserve_cells_selection_and_previous_copy() -> R {
    for (statement, number) in [
        ("Sheets(\"Ingredients\").Range(\"A1\").Select", 1004.),
        ("Sheets(\"Ingredients\").Paste", 1004.),
        ("Sheets(\"Ingredients\").Select Replace:=False", 1004.),
        ("Range(\"A1\").Select False", 450.),
        ("Range(\"A1\").Copy Destination:=Range(\"B1\")", 1004.),
        ("Columns(1).Copy", 7.),
        ("Range(\"A1,C1\").Copy", 1004.),
        ("ActiveSheet.Paste Destination:=Range(\"B1\")", 1004.),
        ("ActiveSheet.Paste Link:=True", 1004.),
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
    }
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
