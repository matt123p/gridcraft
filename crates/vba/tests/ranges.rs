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
fn lightweight_whole_sheet_and_row_column_collections() -> R {
    let source = r#"Option Explicit
Sub Main()
Dim r As Range,c As Range,n As Long
Set r=Range("C3:E4")
Range("A1")=Rows.Count
Range("A2")=Columns.Count
Range("A3")=Cells.CountLarge
Range("A4")=r.Rows.Count
Range("A5")=r.Columns.Count
Range("A6")=r.Rows(2).Address
Range("A7")=r.Columns("B:C").Address
Range("A8")=r.Cells(2, "B").Address
For Each c In r.Rows
n=n+1
c=n
Next
Range("A9")=r.Cells("2.5",1).Address
On Error Resume Next
n=Cells.Count
Range("A10")=Err.Number
Err.Clear
n=Cells.Value2
Range("A11")=Err.Number
Err.Clear
For Each c In Cells
Range("B1")=999
Next
Range("A12")=Err.Number
Err.Clear
Set r=Cells(Row:=1,Column:=1)
Range("A13")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    for (a, n) in [
        ("A1", 1048576.),
        ("A2", 16384.),
        ("A3", 17179869184.),
        ("A4", 2.),
        ("A5", 3.),
        ("A10", 6.),
        ("A11", 7.),
        ("A12", 7.),
        ("A13", 448.),
        ("C3", 1.),
        ("E4", 2.),
    ] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n), "{a}");
    }
    for (a, t) in [("A6", "$C$4:$E$4"), ("A7", "$D$3:$E$4"), ("A8", "$D$4"), ("A9", "$C$4")] {
        assert_eq!(cell(&s, a)?, CellValue::Text(t.into()), "{a}");
    }
    assert!(cell(&s, "B1")?.is_empty());
    assert_eq!(ide.active_module().ok_or("source")?.text, source);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn relative_qualified_names_areas_and_atomic_boundaries() -> R {
    let source = r#"Option Explicit
Sub Main()
Dim r As Range,c As Range,a As Excel.Areas,n As Long
Set r=Range("C3:E5")
Range("A1")=r.Range("$B$2").Address
Range("A2")=r.Range(r.Cells(2,2),r.Cells(3,3)).Address
Range("A3")=r.Offset(ColumnOffset:=-1).Resize(2,2).Address
Range("A4")=Range("LocalBlock").Address
Range("A5")=Range("GlobalBlock").Address
Range("A6")=Range("'Quoted '' Sheet'!B2").Worksheet.Name
Range("A7")=Range("C3:E5 D4:F6").Address
Set r=Range("C10:D10,F10")
Set a=r.Areas
Range("A8")=a.Count
Range("A9")=a(2).Address
Range("A10")=r.CountLarge
For Each c In r
n=n+1
c=n
Next
Range("A11")=r.Rows.Count
On Error Resume Next
r.Value2=99
Range("A12")=Err.Number
Err.Clear
Set r=Range("A1").Offset(-1)
Range("A13")=Err.Number
Err.Clear
Set r=Range("XFD1048576").Resize(2,2)
Range("A14")=Err.Number
Err.Clear
Set r=Range("Cycle")
Range("A15")=Err.Number
Err.Clear
Set r=Worksheets(1).Range("'Quoted '' Sheet'!A1")
Range("A16")=Err.Number
End Sub"#;
    let (mut s, _, _) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Quoted ' Sheet"}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    let mut b = Bridge::new(&s)?;
    b.consent(true);
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).names.extend([
        gridcraft_engine::model::DefinedName {
            name: "LocalBlock".into(),
            scope: Some(0),
            formula: "$D$8:$E$9".into(),
            comment: String::new(),
            hidden: false,
        },
        gridcraft_engine::model::DefinedName {
            name: "LocalBlock".into(),
            scope: None,
            formula: "$Z$1".into(),
            comment: String::new(),
            hidden: false,
        },
        gridcraft_engine::model::DefinedName {
            name: "GlobalBlock".into(),
            scope: None,
            formula: "LocalBlock".into(),
            comment: String::new(),
            hidden: false,
        },
        gridcraft_engine::model::DefinedName { name: "Cycle".into(), scope: None, formula: "Cycle".into(), comment: String::new(), hidden: false },
    ]);
    let mut ide = Ide::new();
    ide.add_module("References.bas", source)?;
    b.configure(&mut ide)?;
    run(&mut s, &mut b, &mut ide)?;
    for (a, t) in [
        ("A1", "$D$4"),
        ("A2", "$D$4:$E$5"),
        ("A3", "$B$3:$C$4"),
        ("A4", "$D$8:$E$9"),
        ("A5", "$D$8:$E$9"),
        ("A6", "Quoted ' Sheet"),
        ("A7", "$D$4:$E$5"),
        ("A9", "$F$10"),
    ] {
        assert_eq!(cell(&s, a)?, CellValue::Text(t.into()), "{a}");
    }
    for (a, n) in [
        ("A8", 2.),
        ("A10", 3.),
        ("A11", 1.),
        ("A12", 1004.),
        ("A13", 1004.),
        ("A14", 1004.),
        ("A15", 7.),
        ("A16", 1004.),
        ("C10", 1.),
        ("D10", 2.),
        ("F10", 3.),
    ] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n), "{a}");
    }
    assert!(cell(&s, "E10")?.is_empty());
    assert_eq!(ide.active_module().ok_or("source")?.text, source);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn last_row_end_find_repeat_and_no_selection_effects() -> R {
    let source = r#"Option Explicit
Sub Main()
Dim c As Range,first As String,direction As Excel.XlDirection
direction=Excel.XlDirection.xlUp
Range("B2")="alpha"
Range("B3")="ALPHABET"
Range("B5")="a*b"
Range("D900000")="needle"
Range("E4").Formula="=2+3"
Range("A1")=Cells(Rows.Count,"B").End(direction).Row
Range("A2")=Range("B2").End(xlDown).Address
Range("A3")=Range("B3").End(xlDown).Address
Range("A4")=Range("Z10").End(xlToRight).Address
Set c=Columns("B").Find(What:="alpha",LookIn:=xlValues,LookAt:=xlPart,SearchOrder:=xlByRows)
Range("A5")=c.Address
Set c=Columns("B").FindNext(c)
Range("A6")=c.Address
Set c=Columns("B").FindPrevious(c)
Range("A7")=c.Address
Set c=Cells.Find(What:="needle",LookIn:=xlValues,LookAt:=xlWhole,SearchDirection:=xlPrevious)
Range("A8")=c.Address
Set c=Columns("B").Find(What:="a~*b",LookAt:=xlWhole)
Range("A9")=c.Address
Set c=Range("E1:E8").Find(What:="=2+3",LookIn:=xlFormulas,LookAt:=xlWhole)
Range("A10")=c.Address
Set c=Range("E1:E8").Find(What:=5,LookIn:=xlValues,LookAt:=xlWhole)
Range("A11")=c.Address
Set c=Range("Z1:Z2").Find(What:="missing")
Range("A12")=(c Is Nothing)
On Error Resume Next
Set c=Range("B1:B6").Find(What:="alpha",After:=Range("D1"))
Range("A13")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    let selection = s.doc()?.selection.clone();
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "A1")?, CellValue::Number(5.));
    for (a, t) in [
        ("A2", "$B$3"),
        ("A3", "$B$5"),
        ("A4", "$XFD$10"),
        ("A5", "$B$2"),
        ("A6", "$B$3"),
        ("A7", "$B$2"),
        ("A8", "$D$900000"),
        ("A9", "$B$5"),
        ("A10", "$E$4"),
        ("A11", "$E$4"),
    ] {
        assert_eq!(cell(&s, a)?, CellValue::Text(t.into()), "{a}");
    }
    assert_eq!(cell(&s, "A12")?, CellValue::Bool(true));
    assert_eq!(cell(&s, "A13")?, CellValue::Number(1004.));
    assert_eq!(s.doc()?.selection, selection);
    assert_eq!(ide.active_module().ok_or("source")?.text, source);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn structural_commands_rebase_ranges_recalculate_and_undo_as_one_macro() -> R {
    let source = r#"Option Explicit
Sub Main()
Dim r As Range,gone As Range
Range("B2")=10
Range("C1").Formula="=B2*2"
Set r=Range("B2")
Rows(2).Insert
r=r+1
Range("D1")=r.Address
Set gone=Range("B2")
Rows(2).Delete
Range("D2")=r.Address
On Error Resume Next
Range("D3")=gone.Address
Range("D4")=Err.Number
Err.Clear
Range("B2").Insert Shift:=xlShiftDown
Range("D5")=r.Address
Range("B2").Delete Shift:=xlShiftUp
Range("D6")=r.Address
Range("E1")=ActiveSheet.UsedRange.Address
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "B2")?, CellValue::Number(11.));
    assert_eq!(cell(&s, "C1")?, CellValue::Number(22.));
    for (a, t) in [("D1", "$B$3"), ("D2", "$B$2"), ("D5", "$B$3"), ("D6", "$B$2"), ("E1", "$B$1:$D$6")] {
        assert_eq!(cell(&s, a)?, CellValue::Text(t.into()), "{a}");
    }
    assert!(cell(&s, "D3")?.is_empty());
    assert_eq!(cell(&s, "D4")?, CellValue::Number(1004.));
    assert_eq!(s.doc()?.undo.len(), 1);
    assert!(s.journal.iter().any(|j| j.0 == "home.insertRows"));
    assert!(s.journal.iter().any(|j| j.0 == "home.deleteCells"));
    s.execute("edit.undo", serde_json::json!({}))?;
    assert!(cell(&s, "B2")?.is_empty());
    s.execute("edit.redo", serde_json::json!({}))?;
    assert_eq!(cell(&s, "C1")?, CellValue::Number(22.));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn structural_refusals_preserve_workbook_and_prior_effects() -> R {
    let source = r#"Sub Main()
On Error Resume Next
Rows(2).Insert
If Err.Number<>1004 Then Err.Raise 5
Err.Clear
Range("A2").Delete Shift:=xlShiftUp
If Err.Number<>1004 Then Err.Raise 5
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", serde_json::json!({"cell":"A2","input":"41"}))?;
    s.execute("review.protectSheet", serde_json::json!({}))?;
    let before = s.doc()?.wb.clone();
    let undo = s.doc()?.undo.len();
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(*before, *s.doc()?.wb);
    assert_eq!(s.doc()?.undo.len(), undo);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn structural_edits_use_the_explicit_sheet_and_refuse_push_off_atomically() -> R {
    let source = r#"Sub Main()
Dim ws As Worksheet,r As Range
Set ws=Worksheets(2)
ws.Range("B2")=41
Set r=ws.Range("B2")
ws.Rows(2).Insert
Range("A1")=r.Address
Range("A2")=ws.Range("B3")
ws.Range("XFD7")=9
On Error Resume Next
ws.Range("XFD7").Insert Shift:=xlShiftToRight
Range("A3")=Err.Number
Err.Clear
Range("B1:C2,D5").Delete Shift:=xlShiftUp
Range("A4")=Err.Number
End Sub"#;
    let (mut s, _, _) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    let mut b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("OtherSheet.bas", source)?;
    b.configure(&mut ide)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "A1")?, CellValue::Text("$B$3".into()));
    for (a, n) in [("A2", 41.), ("A3", 1004.), ("A4", 1004.)] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n));
    }
    assert!(cell(&s, "B3")?.is_empty());
    assert_eq!(s.doc()?.wb.sheet(1).ok_or("other")?.value(CellRef::parse("XFD7").ok_or("cell")?), CellValue::Number(9.));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn sparse_search_budget_and_bad_geometry_leave_no_partial_effects() -> R {
    let source = r#"Sub Main()
Dim r As Range
On Error Resume Next
Set r=Columns(2).Find("absent")
Range("A1")=Err.Number
Err.Clear
Set r=Cells(Rows.Count,2).End(xlUp)
Range("A2")=Err.Number
Err.Clear
Set r=Range("B1").Offset(CDbl("2147483648"))
Range("A3")=Err.Number
Err.Clear
Set r=Range("A1,B3").Areas(CLng("-2147483648"))
Range("A4")=Err.Number
Err.Clear
Set r=Cells(Null,1)
Range("A5")=Err.Number
Err.Clear
Set r=Range("A1,A1")
Range("A6")=Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    for row in 0..65537 {
        std::sync::Arc::make_mut(&mut s.doc_mut()?.wb)
            .sheet_mut(0)
            .ok_or("sheet")?
            .cells
            .set(CellRef::new(row, 1), gridcraft_engine::model::Cell { value: CellValue::Number(1.), ..Default::default() });
    }
    run(&mut s, &mut b, &mut ide)?;
    for (a, n) in [("A1", 7.), ("A2", 7.), ("A3", 6.), ("A4", 1004.), ("A5", 94.), ("A6", 1004.)] {
        assert_eq!(cell(&s, a)?, CellValue::Number(n), "{a}");
    }
    assert_eq!(cell(&s, "B1")?, CellValue::Number(1.));
    Ok(())
}
