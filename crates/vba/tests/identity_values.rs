//! Original Excel-shaped source; all Office definitions and document assertions live here.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value as CellValue},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
type R<T = ()> = Result<T, Box<dyn std::error::Error>>;
fn setup(source: &str) -> R<(Session, Bridge, Ide)> {
    let mut s = Session::new();
    s.new_workbook();
    let mut b = Bridge::new(&s)?;
    b.consent(true);
    let mut ide = Ide::new();
    ide.add_module("Collections.bas", source)?;
    b.configure(&mut ide)?;
    Ok((s, b, ide))
}
fn drain(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    for _ in 0..1000 {
        b.tick(s, ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            return Ok(());
        }
    }
    Err("macro timeout".into())
}
fn run(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    b.dispatch(s, ide, Action::Run)?;
    drain(s, b, ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
fn cell(s: &Session, address: &str) -> R<CellValue> {
    Ok(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::parse(address).ok_or("cell")?))
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn address_optional_arguments_named_styles_and_external_ownership() -> R {
    let source = r#"Sub Main()
Dim r As Range
Set r=Range("C3:D4")
Range("A1")=r.Address(False,False)
Range("A2")=r.Address(ColumnAbsolute:=False)
Range("A3")=r.Address(ReferenceStyle:=xlR1C1)
Range("A4")=r.Address(False,False,xlR1C1,,Range("D4"))
Range("A5")=Range("C3").Address(False,False,xlR1C1)
Range("A6")=r.Address(External:=True)
On Error Resume Next
Range("A7")=r.Address(ReferenceStyle:=123)
Range("A7")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    for (a, t) in [("A1", "C3:D4"), ("A2", "C$3:D$4"), ("A3", "R3C3:R4C4"), ("A4", "R[-1]C[-1]:RC"), ("A5", "R[2]C[2]")] {
        assert_eq!(cell(&s, a)?, CellValue::Text(t.into()));
    }
    let title = &s.doc()?.title;
    assert_eq!(cell(&s, "A6")?, CellValue::Text(format!("[{title}]Sheet1!$C$3:$D$4").into()));
    assert_eq!(cell(&s, "A7")?, CellValue::Number(1004.));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn live_formats_value_tags_mixed_reads_protection_and_undo() -> R {
    let source = r#"Sub Main()
Dim f As Excel.Font
Range("C1").NumberFormat="yyyy-mm-dd"
Range("C1").Value2=45000
Range("A1")=VarType(Range("C1").Value)
Range("A2")=VarType(Range("C1").Value2)
Range("C2").NumberFormat="$0.0000"
Range("C2").Value2=1.23456
Range("A3")=VarType(Range("C2").Value)
Range("A4")=VarType(Range("C2").Value2)
Range("A5")=CDbl(Range("C2").Value)
Set f=Range("C1:C2").Font
f.Bold=True
f.Italic=True
f.Size=14
f.Name="Example Font"
Range("C1:C2").Interior.Color=255
Range("A6")=f.Bold
Range("A7")=f.Size
Range("A8")=Range("C2").Interior.Color
Range("A9")=IsNull(Range("C1:C2").NumberFormat)
On Error Resume Next
f.Size=500
Range("A10")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    for (a, n) in [("A1", 7.), ("A2", 5.), ("A3", 6.), ("A4", 5.), ("A5", 1.2346), ("A7", 14.), ("A8", 255.), ("A10", 1004.)] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n), "{a}");
    }
    assert_eq!(cell(&s, "A6")?, CellValue::Bool(true));
    assert_eq!(cell(&s, "A9")?, CellValue::Bool(true));
    assert_eq!(s.doc()?.undo.len(), 1);
    s.execute("edit.undo", serde_json::json!({}))?;
    assert!(cell(&s, "C1")?.is_empty());
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert!(!s.doc()?.wb.styles.get(sh.style_id(CellRef::default())).font.bold);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn captured_handles_survive_rename_reorder_and_dynamic_roots_follow_selection() -> R {
    let source = r#"Sub Main()
Dim saved As Worksheet,r As Range
Set saved=ActiveSheet
Set r=Cells(2,2)
Stop
saved.Name="Renamed"
r.Value2=42
Cells(1,1)=ActiveSheet.Name
Range("A2")=ActiveCell.Address
Range("A3")=Selection.Address
saved.Range("C1")=saved.Name
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Other", "before":1}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    let owner = s.doc()?.wb.sheet(0).ok_or("sheet")?.identity.clone();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut s, &mut b, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    s.execute("sheet.move", serde_json::json!({"sheet":0,"to":2}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    s.execute("selection.set", serde_json::json!({"range":"D4:E5"}))?;
    run(&mut s, &mut b, &mut ide)?;
    let original = s.doc()?.wb.sheets.iter().find(|sh| sh.identity.same(&owner)).ok_or("original identity")?;
    assert_eq!(original.name, "Renamed");
    assert_eq!(original.value(CellRef::new(1, 1)), CellValue::Number(42.));
    assert_eq!(original.value(CellRef::new(0, 2)), CellValue::Text("Renamed".into()));
    let active = s.doc()?.wb.active().ok_or("active")?;
    assert_eq!(active.value(CellRef::new(0, 0)), CellValue::Text(active.name.clone().into()));
    assert_eq!(active.value(CellRef::new(1, 0)), CellValue::Text("$D$4".into()));
    assert_eq!(active.value(CellRef::new(2, 0)), CellValue::Text("$D$4:$E$5".into()));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn code_owner_and_active_workbook_remain_distinct_across_switch_and_undo() -> R {
    let source = r#"Sub Main()
Dim saved As Range
Set saved=Range("B1")
Stop
saved=41
ThisWorkbook.Worksheets(1).Range("A1")=ThisWorkbook.Name
Range("A1")=ActiveWorkbook.Name
Range("A2")=saved.Parent.Parent.Name
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    let original = b.document();
    let title = s.doc()?.title.clone();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut s, &mut b, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    s.new_workbook();
    let other = s.doc()?.uid;
    let other_title = s.doc()?.title.clone();
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.doc()?.uid, other);
    assert_eq!(cell(&s, "A1")?, CellValue::Text(other_title.into()));
    assert_eq!(cell(&s, "A2")?, CellValue::Text(title.clone().into()));
    let owner = s.document(original).ok_or("owner")?;
    assert_eq!(owner.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 1)), CellValue::Number(41.));
    assert_eq!(owner.wb.sheet(0).ok_or("sheet")?.value(CellRef::default()), CellValue::Text(title.into()));
    assert!(!s.in_group(original));
    assert!(!s.in_group(other));
    s.execute("edit.undo", serde_json::json!({}))?;
    assert!(cell(&s, "A1")?.is_empty());
    assert_eq!(s.document(original).ok_or("owner")?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 1)), CellValue::Number(41.));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn calculation_mode_and_calculate_have_real_engine_effects() -> R {
    let source = r#"Sub Main()
Application.Calculation=xlCalculationManual
Range("B1").Formula="=A1*2"
Range("A1")=21
Application.Calculate
Range("C1")=Range("B1")
Range("C2")=Application.Calculation
Application.Calculation=xlCalculationAutomatic
Range("A1")=22
Range("C3")=Range("B1")
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    for (a, n) in [("C1", 42.), ("C2", -4135.), ("C3", 44.)] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n));
    }
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn arrays_crop_or_fill_errors_and_deleted_sheets_never_retarget() -> R {
    let source = r#"Sub Main()
Dim a(3 To 3,-2 To -1) As Variant,saved As Worksheet
Set saved=ActiveSheet
a(3,-2)=11
a(3,-1)=12
Range("A1:C2").Value2=a
Range("D1").Value2=a
Stop
On Error Resume Next
saved.Range("A1")=999
ActiveSheet.Range("E1")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Other", "before":1}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut s, &mut b, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    for (a, n) in [("A1", 11.), ("B1", 12.), ("D1", 11.)] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n));
    }
    for a in ["C1", "A2", "C2"] {
        assert_eq!(cell(&s, a)?, CellValue::Error(gridcraft_engine::core::CellError::NA));
    }
    s.execute("home.deleteSheet", serde_json::json!({"sheet":0}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "E1")?, CellValue::Number(91.));
    assert!(cell(&s, "A1")?.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn external_structural_edits_refuse_saved_coordinates_and_formats_obey_protection() -> R {
    let source = r#"Sub Main()
Dim r As Range
Set r=Range("B2")
Stop
On Error Resume Next
r=999
Range("A1")=Err.Number
Err.Clear
Range("D1").Font.Bold=True
Range("A2")=Err.Number
Err.Clear
Application.ScreenUpdating=False
Range("A3")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut s, &mut b, &mut ide)?;
    s.execute("home.insertRows", serde_json::json!({"rows":"1:1"}))?;
    // Leave output cells unlocked while formatting remains protected.
    s.execute("home.formatCells", serde_json::json!({"range":"A1:A3","style":{"protection":{"locked":false}}}))?;
    s.execute("review.protectSheet", serde_json::json!({}))?;
    run(&mut s, &mut b, &mut ide)?;
    for (a, n) in [("A1", 1004.), ("A2", 1004.), ("A3", 438.)] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n), "{a}");
    }
    assert!(cell(&s, "B2")?.is_empty());
    assert!(cell(&s, "B3")?.is_empty());
    let catalog = gridcraft_vba::catalog();
    assert!(catalog.iter().any(|m| m.name == "ScreenUpdating" && !m.supported));
    assert!(catalog.iter().any(|m| m.name == "Address" && m.signature.contains("RelativeTo")));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn untrusted_active_target_keeps_consent_scoped_to_code_owner() -> R {
    let source = r#"Sub Main()
Dim saved As Workbook
Set saved=ActiveWorkbook
Stop
On Error Resume Next
Range("A1")=999
ThisWorkbook.Worksheets(1).Range("A1")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut s, &mut b, &mut ide)?;
    s.add_document(gridcraft_engine::DocState::new(
        gridcraft_engine::model::Workbook::new(),
        None,
        "Untrusted".into(),
        gridcraft_engine::Trust::FromFile,
    ));
    run(&mut s, &mut b, &mut ide)?;
    assert!(cell(&s, "A1")?.is_empty());
    let owner = s.document(b.document()).ok_or("owner")?;
    assert_eq!(owner.wb.sheet(0).ok_or("sheet")?.value(CellRef::default()), CellValue::Number(70.));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn closed_saved_workbooks_refuse_even_when_another_document_takes_their_index() -> R {
    let source = r#"Sub Main()
Dim saved As Workbook,n As String
Set saved=ActiveWorkbook
Stop
On Error Resume Next
n=saved.Name
ThisWorkbook.Worksheets(1).Range("A1")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.new_workbook();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut s, &mut b, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    s.close_document(1);
    s.new_workbook();
    run(&mut s, &mut b, &mut ide)?;
    let owner = s.document(b.document()).ok_or("owner")?;
    assert_eq!(owner.wb.sheet(0).ok_or("sheet")?.value(CellRef::default()), CellValue::Number(91.));
    assert!(cell(&s, "A1")?.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn typed_value_writes_and_1904_dates_remain_distinct_from_value2() -> R {
    let source = r#"Sub Main()
Range("B1").Value=CDate(45000.5)
Range("A1")=VarType(Range("B1").Value)
Range("A2")=CDbl(Range("B1").Value)
Range("A3")=Range("B1").Value2
Range("B2").Value=CCur(1.23456)
Range("B3").Value2=CCur(1.23456)
Range("A4")=Range("B2").Value2
Range("A5")=Range("B3").Value2
Range("A6")=VarType(Range("B2").Value)
Range("A7")=VarType(Range("B3").Value)
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).date_system = gridcraft_engine::core::date::DateSystem::D1904;
    run(&mut s, &mut b, &mut ide)?;
    for (a, n) in [("A1", 7.), ("A2", 45000.5), ("A3", 43538.5), ("A4", 1.23), ("A5", 1.2346), ("A6", 6.), ("A7", 5.)] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n), "{a}");
    }
    Ok(())
}
