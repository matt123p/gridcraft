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
fn range_clear_members_use_native_content_and_format_semantics() -> R {
    let source = r#"Sub Main()
Range("A1").ClearContents
Range("A2").ClearFormats
Range("A3").Clear
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    for address in ["A1", "A2", "A3"] {
        s.execute("cell.set", serde_json::json!({"cell":address,"input":"17"}))?;
        s.execute("home.bold", serde_json::json!({"range":address,"on":true}))?;
    }
    run(&mut s, &mut b, &mut ide)?;
    let sheet = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(cell(&s, "A1")?, CellValue::Empty);
    assert!(s.doc()?.wb.styles.get(sheet.style_id(CellRef::parse("A1").ok_or("cell")?)).font.bold);
    assert_eq!(cell(&s, "A2")?, CellValue::Number(17.));
    assert!(!s.doc()?.wb.styles.get(sheet.style_id(CellRef::parse("A2").ok_or("cell")?)).font.bold);
    assert_eq!(cell(&s, "A3")?, CellValue::Empty);
    assert!(!s.doc()?.wb.styles.get(sheet.style_id(CellRef::parse("A3").ok_or("cell")?)).font.bold);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_clear_notes_comments_and_hyperlinks_use_native_feature_maps() -> R {
    let source = r#"Sub Main()
Range("A1").ClearComments
Range("A2").ClearHyperlinks
Range("A3").ClearNotes
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("review.newComment", serde_json::json!({"cell":"A1","text":"thread"}))?;
    s.execute("insert.link", serde_json::json!({"cell":"A2","target":"https://example.com"}))?;
    s.execute("review.newNote", serde_json::json!({"cell":"A3","text":"note"}))?;
    run(&mut s, &mut b, &mut ide)?;
    let sheet = s.doc()?.wb.active().ok_or("sheet")?;
    assert!(sheet.comments.is_empty(), "remaining comments: {:?}", sheet.comments.keys().collect::<Vec<_>>());
    assert!(sheet.hyperlinks.is_empty(), "remaining hyperlinks: {:?}", sheet.hyperlinks.keys().collect::<Vec<_>>());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_replace_is_limited_to_its_receiver_and_reports_a_match() -> R {
    let source = r#"Sub Main()
Dim changed As Boolean
changed = Range("A1:A3").Replace(What:="cat", Replacement:="dog", LookAt:=xlWhole, MatchCase:=False)
Range("E1") = changed
Range("E2") = Range("A1:A3").Replace(What:="missing", Replacement:="x")
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    for (address, value) in [("A1", "cat"), ("A2", "CAT"), ("A3", "wildcat"), ("A4", "cat")] {
        s.execute("cell.set", serde_json::json!({"cell":address,"input":value}))?;
    }
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "A1")?, CellValue::Text("dog".into()));
    assert_eq!(cell(&s, "A2")?, CellValue::Text("dog".into()));
    assert_eq!(cell(&s, "A3")?, CellValue::Text("wildcat".into()));
    assert_eq!(cell(&s, "A4")?, CellValue::Text("cat".into()));
    assert_eq!(cell(&s, "E1")?, CellValue::Bool(true));
    assert_eq!(cell(&s, "E2")?, CellValue::Bool(false));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_replace_queues_exact_change_targets_and_refuses_service_options_atomically() -> R {
    let source = r#"Sub Main()
Range("A1:A3").Replace What:="cat", Replacement:="dog", LookAt:=xlWhole
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.code_name = Some("MenuCode".into());
    ide.add_document_module(
        "Menu.cls",
        "MenuCode",
        "Attribute VB_Name = \"MenuCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nApplication.EnableEvents = False\nMe.Range(\"Z1\") = Target.Address\nApplication.EnableEvents = True\nEnd Sub\n",
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    for (address, value) in [("A1", "cat"), ("A2", "cat"), ("A3", "x"), ("B1", "cat")] {
        s.execute("cell.set", serde_json::json!({"cell":address,"input":value}))?;
    }
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "Z1")?, CellValue::Text("$A$1:$A$2".into()));
    assert_eq!(cell(&s, "B1")?, CellValue::Text("cat".into()));

    for (options, number) in [
        ("MatchByte:=True", 1004.),
        ("SearchFormat:=True", 1004.),
        ("ReplaceFormat:=True", 1004.),
        ("FormulaVersion:=1", 1004.),
        ("Unknown:=True", 448.),
    ] {
        let code = format!(
            "Sub Main()\nOn Error Resume Next\nRange(\"B1\").Replace What:=\"cat\", Replacement:=\"dog\", {options}\nRange(\"B2\")=Err.Number\nEnd Sub"
        );
        let (mut session, mut bridge, mut interpreter) = setup(&code)?;
        session.execute("cell.set", serde_json::json!({"cell":"B1","input":"cat"}))?;
        run(&mut session, &mut bridge, &mut interpreter)?;
        assert_eq!(cell(&session, "B1")?, CellValue::Text("cat".into()), "{options}");
        assert_eq!(cell(&session, "B2")?, CellValue::Number(number), "{options}");
    }

    let code = r#"Sub Main()
On Error Resume Next
Range("A1:EY1").Replace What:="x", Replacement:="y", LookAt:=xlWhole
Application.EnableEvents = False
Range("E2") = Err.Number
End Sub"#;
    let (mut session, mut bridge, mut interpreter) = setup(code)?;
    std::sync::Arc::make_mut(&mut session.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.code_name = Some("MenuCode".into());
    interpreter.add_document_module(
        "Menu.cls",
        "MenuCode",
        "Attribute VB_Name = \"MenuCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nApplication.EnableEvents = False\nMe.Range(\"Z1\") = Target.Address\nApplication.EnableEvents = True\nEnd Sub\n",
    )?;
    bridge.configure_document(&mut interpreter, &session.doc()?.wb)?;
    for column in (0..129).step_by(2) {
        session.execute("cell.set", serde_json::json!({"cell":gridcraft_engine::core::CellRef::new(0, column).a1(),"input":"x"}))?;
    }
    run(&mut session, &mut bridge, &mut interpreter)?;
    assert_eq!(cell(&session, "E2")?, CellValue::Number(7.));
    assert_eq!(cell(&session, "Z1")?, CellValue::Empty);
    for column in (0..129).step_by(2) {
        assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::new(0, column)), CellValue::Text("x".into()));
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_activate_switches_to_its_sheet_and_sets_the_active_cell() -> R {
    let source = r#"Sub Main()
Dim target As Range
Set target = Sheets("Ingredients").Range("B2:C3")
target.Activate
ActiveCell.Value2 = 29
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Ingredients"}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.doc()?.sheet_index(), 1);
    assert_eq!(s.doc()?.selection.a1(), "B2:C3");
    assert_eq!(s.doc()?.selection.active, CellRef::parse("B2").ok_or("cell")?);
    assert_eq!(s.doc()?.wb.sheet(1).ok_or("sheet")?.value(CellRef::parse("B2").ok_or("cell")?), CellValue::Number(29.));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn current_region_uses_the_engine_contiguous_region() -> R {
    let source = r#"Sub Main()
Range("E1") = Range("B2").CurrentRegion.Address
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    for (address, value) in [("A1", "1"), ("B1", "2"), ("A2", "3"), ("B2", "4"), ("D1", "separate")] {
        s.execute("cell.set", serde_json::json!({"cell":address,"input":value}))?;
    }
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "E1")?, CellValue::Text("$A$1:$B$2".into()));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn current_array_returns_the_full_legacy_array_and_rejects_other_cells() -> R {
    let source = r#"Sub Main()
Range("E1") = Range("C3").CurrentArray.Address
On Error Resume Next
Range("E2") = Range("A1").CurrentArray.Address
Range("E3") = Err.Number
Err.Clear
Range("E4") = Range("C3").HasArray
Range("E5") = Range("A1").HasArray
Range("E6") = IsNull(Range("A1,C3").HasArray)
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("range.setValues", serde_json::json!({"range":"B2","values":[[1],[2]]}))?;
    s.execute("selection.set", serde_json::json!({"range":"C2:C4"}))?;
    s.execute("cell.set", serde_json::json!({"cell":"C2","input":"=B2:B3*2","array":true}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "E1")?, CellValue::Text("$C$2:$C$4".into()));
    assert_eq!(cell(&s, "E3")?, CellValue::Number(1004.));
    assert_eq!(cell(&s, "E4")?, CellValue::Bool(true));
    assert_eq!(cell(&s, "E5")?, CellValue::Bool(false));
    assert_eq!(cell(&s, "E6")?, CellValue::Bool(true));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn spill_properties_and_next_previous_return_ranges_without_selecting() -> R {
    let source = r#"Sub Main()
Range("D1") = Range("A2").HasSpill
Range("D2") = Range("A1:A2").HasSpill
Range("D3") = IsNull(Range("A1:A3").HasSpill)
Range("D4") = Range("A2").SpillParent.Address
Range("D5") = Range("B2").Next.Address
Range("D6") = Range("B2").Previous.Address
Range("D7") = Selection.Address
On Error Resume Next
Range("A3").SpillParent.Address
Range("D8") = Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", serde_json::json!({"cell":"A1","input":"=SEQUENCE(2,1)"}))?;
    s.execute("selection.set", serde_json::json!({"range":"C5"}))?;
    run(&mut s, &mut b, &mut ide)?;
    for (address, expected) in [
        ("D1", CellValue::Bool(true)),
        ("D2", CellValue::Bool(true)),
        ("D3", CellValue::Bool(true)),
        ("D4", CellValue::Text("$A$1".into())),
        ("D5", CellValue::Text("$C$2".into())),
        ("D6", CellValue::Text("$A$2".into())),
        ("D7", CellValue::Text("$C$5".into())),
        ("D8", CellValue::Number(1004.)),
    ] {
        assert_eq!(cell(&s, address)?, expected, "{address}");
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn directional_fill_methods_copy_values_formulas_and_formats_from_the_correct_edge() -> R {
    let source = r#"Sub Main()
Range("A1:B3").FillDown
Range("D1:D3").FillUp
Range("F1:H1").FillRight
Range("J1:L1").FillLeft
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("cell.set", serde_json::json!({"cell":"A1","input":"1"}))?;
    s.execute("cell.set", serde_json::json!({"cell":"B1","input":"=A1+1"}))?;
    s.execute("home.bold", serde_json::json!({"range":"A1:B1","on":true}))?;
    s.execute("cell.set", serde_json::json!({"cell":"D3","input":"5"}))?;
    s.execute("cell.set", serde_json::json!({"cell":"F1","input":"7"}))?;
    s.execute("cell.set", serde_json::json!({"cell":"L1","input":"9"}))?;
    run(&mut s, &mut b, &mut ide)?;
    for row in 0..3 {
        assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(row, 0)), CellValue::Number(1.));
        assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(row, 1)), CellValue::Number(2.));
        assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(row, 3)), CellValue::Number(5.));
    }
    for column in 5..=7 {
        assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, column)), CellValue::Number(7.));
    }
    for column in 9..=11 {
        assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, column)), CellValue::Number(9.));
    }
    let sheet = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(
        sheet.cell(CellRef::parse("B3").ok_or("cell")?).and_then(|cell| cell.formula.as_ref()).map(|formula| formula.text.as_str()),
        Some("A3+1")
    );
    assert!(s.doc()?.wb.styles.get(sheet.style_id(CellRef::parse("B3").ok_or("cell")?)).font.bold);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn autofill_extends_a_series_and_refuses_invalid_destinations_before_mutation() -> R {
    let source = r#"Sub Main()
Range("A1:A2").AutoFill Destination:=Range("A1:A5")
Range("B1").AutoFill Range("B1:B3"), xlFillCopy
On Error Resume Next
Range("C1:C2").AutoFill Destination:=Range("C3:C5")
Range("D1") = Err.Number
Err.Clear
Range("C1:C2").AutoFill Destination:=Sheets("Other").Range("C1:C5")
Range("D2") = Err.Number
Err.Clear
Range("C1:C2").AutoFill Destination:=Range("C1:C5"), Type:=999
Range("D3") = Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Other"}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    s.execute("range.setValues", serde_json::json!({"range":"A1","values":[[1],[3]]}))?;
    s.execute("cell.set", serde_json::json!({"cell":"B1","input":"4"}))?;
    s.execute("range.setValues", serde_json::json!({"range":"C1","values":[[10],[20]]}))?;
    run(&mut s, &mut b, &mut ide)?;
    for (address, value) in [("A3", 5.), ("A4", 7.), ("A5", 9.), ("B2", 4.), ("B3", 4.)] {
        assert_eq!(cell(&s, address)?, CellValue::Number(value), "{address}");
    }
    for (address, number) in [("D1", 1004.), ("D2", 1004.), ("D3", 1004.)] {
        assert_eq!(cell(&s, address)?, CellValue::Number(number), "{address}");
    }
    assert_eq!(cell(&s, "C3")?, CellValue::Empty);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn autofill_queues_only_the_filled_destination_cells() -> R {
    let source = r#"Sub Main()
Range("A1:A2").AutoFill Destination:=Range("A1:A5")
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.code_name = Some("SheetCode".into());
    ide.add_document_module(
        "Sheet.cls",
        "SheetCode",
        "Attribute VB_Name = \"SheetCode\"\nPrivate Sub Worksheet_Change(ByVal Target As Range)\nApplication.EnableEvents = False\nMe.Range(\"Z1\") = Target.Address\nApplication.EnableEvents = True\nEnd Sub\n",
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    s.execute("range.setValues", serde_json::json!({"range":"A1","values":[[1],[3]]}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "Z1")?, CellValue::Text("$A$3:$A$5".into()));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn union_and_intersect_are_bounded_same_sheet_range_operations() -> R {
    let source = r#"Sub Main()
Dim combined As Range
Dim overlap As Range
Dim absent As Range
Set combined = Application.Union(Range("A1:B2"), Range("B2:C3"))
Range("Z1") = combined.Count
Range("Z2") = combined.Areas.Count
Set overlap = Application.Intersect(Range("A1:C3"), Range("B2:D4"))
Range("Z3") = overlap.Address
Set absent = Application.Intersect(Range("A1"), Range("C3"))
Range("Z4") = (absent Is Nothing)
On Error Resume Next
Set overlap = Application.Union(Range("A1"), Sheets("Other").Range("A1"))
Range("Z5") = Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Other"}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "Z1")?, CellValue::Number(7.));
    assert_eq!(cell(&s, "Z2")?, CellValue::Number(3.));
    assert_eq!(cell(&s, "Z3")?, CellValue::Text("$B$2:$C$3".into()));
    assert_eq!(cell(&s, "Z4")?, CellValue::Bool(true));
    assert_eq!(cell(&s, "Z5")?, CellValue::Number(1004.));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn special_cells_filters_constants_formulas_blanks_and_last_cell() -> R {
    let source = r#"Sub Main()
Dim hit As Range
Dim lastAddress As String
Range("A1") = 1
Range("A2") = "text"
Range("A3") = True
Range("A4").Formula = "=1/0"
Range("A5").Formula = "=1+1"
lastAddress = Range("A1:F10").SpecialCells(xlCellTypeLastCell).Address
Set hit = Range("A1:A6").SpecialCells(xlCellTypeConstants, xlNumbers)
Range("Z1") = hit.Address
Set hit = Range("A1:A6").SpecialCells(xlCellTypeConstants, xlTextValues + xlLogical)
Range("Z2") = hit.Address
Set hit = Range("A1:A6").SpecialCells(xlCellTypeFormulas, xlErrors)
Range("Z3") = hit.Address
Set hit = Range("A1:A6").SpecialCells(xlCellTypeFormulas, xlNumbers)
Range("Z4") = hit.Address
Set hit = Range("A1:A6").SpecialCells(xlCellTypeBlanks)
Range("Z5") = hit.Address
Range("Z6") = lastAddress
On Error Resume Next
Set hit = Range("B1:B2").SpecialCells(xlCellTypeConstants)
Range("Z7") = Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    for (address, expected) in [("Z1", "$A$1"), ("Z2", "$A$2:$A$3"), ("Z3", "$A$4"), ("Z4", "$A$5"), ("Z5", "$A$6"), ("Z6", "$A$5")] {
        assert_eq!(cell(&s, address)?, CellValue::Text(expected.into()), "{address}");
    }
    assert_eq!(cell(&s, "Z7")?, CellValue::Number(1004.));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn data_series_uses_native_linear_growth_and_date_fills_atomically() -> R {
    let source = r#"Sub Main()
Range("A1") = 1
Range("A1:A5").DataSeries Type:=xlDataSeriesLinear, Step:=2
Range("B1") = 2
Range("B1:B5").DataSeries Type:=xlGrowth, Step:=3, Stop:=20
Range("C1") = 10
Range("C1:C3").DataSeries Type:=xlChronological, Date:=xlDay, Step:=2
On Error Resume Next
Range("D1") = 4
Range("D1:D3").DataSeries Type:=xlDataSeriesLinear, Trend:=True
Range("E1") = 8
Range("E1:E3").DataSeries Type:=xlAutoFill
Range("Z1") = Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    for (address, expected) in [
        ("A1", 1.),
        ("A2", 3.),
        ("A3", 5.),
        ("A4", 7.),
        ("A5", 9.),
        ("B1", 2.),
        ("B2", 6.),
        ("B3", 18.),
        ("C1", 10.),
        ("C2", 12.),
        ("C3", 14.),
        ("D1", 4.),
        ("E1", 8.),
        ("Z1", 1004.),
    ] {
        assert_eq!(cell(&s, address)?, CellValue::Number(expected), "{address}");
    }
    for address in ["B4", "B5", "D2", "D3", "E2", "E3"] {
        assert_eq!(cell(&s, address)?, CellValue::Empty, "{address}");
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_goto_activates_the_range_sheet_and_preserves_selection_identity() -> R {
    let source = r#"Sub Main()
Dim target As Range
Set target = ActiveWorkbook.Sheets("Ingredients").Range(Sheets("Ingredients").Cells(2, 3), Sheets("Ingredients").Cells(4, 3))
Application.Goto (target)
Selection.Value2 = 17
ActiveSheet.Range("A1") = ActiveCell.Address
Application.Goto Scroll:=False, Reference:=Sheets("Sheet1").Range("B2,D4")
ActiveCell.Value2 = 29
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Ingredients"}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    run(&mut s, &mut b, &mut ide)?;
    let ingredients = s.doc()?.wb.sheet(1).ok_or("ingredients sheet")?;
    for address in ["C2", "C3", "C4"] {
        assert_eq!(ingredients.value(CellRef::parse(address).ok_or("cell")?), CellValue::Number(17.));
    }
    assert_eq!(ingredients.value(CellRef::new(0, 0)), CellValue::Text("$C$2".into()));
    assert_eq!(s.doc()?.sheet_index(), 0);
    assert_eq!(s.doc()?.selection.a1(), "B2,D4");
    assert_eq!(s.doc()?.selection.active, CellRef::new(1, 1));
    assert_eq!(cell(&s, "B2")?, CellValue::Number(29.));
    assert_eq!(cell(&s, "D4")?, CellValue::Empty);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_goto_refuses_invalid_forms_without_changing_selection() -> R {
    for (statement, number) in [
        ("Application.Goto Reference:=ActiveSheet", 1004.),
        ("Application.Goto Reference:=Range(\"B2\"), Scroll:=\"invalid\"", 13.),
        ("Application.Goto Reference:=Range(\"B2\"), Unknown:=False", 448.),
        ("Application.Goto Range(\"B2\"), False, False", 450.),
    ] {
        let source = format!("Sub Main()\nOn Error Resume Next\n{statement}\nRange(\"A5\") = Err.Number\nEnd Sub");
        let (mut s, mut b, mut ide) = setup(&source)?;
        let selection = s.doc()?.selection.clone();
        run(&mut s, &mut b, &mut ide)?;
        assert_eq!(cell(&s, "A5")?, CellValue::Number(number), "{statement}");
        assert_eq!(s.doc()?.selection, selection, "{statement}");
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_goto_records_replays_and_scrolls_native_history() -> R {
    let source = r#"Sub Main()
Application.Goto Reference:="Sheet1!C4", Scroll:=True
Range("A1") = ActiveCell.Address
Application.Goto Reference:=Range("E6")
Application.Goto
Range("A2") = ActiveCell.Address
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "A1")?, CellValue::Text("$C$4".into()));
    assert_eq!(cell(&s, "A2")?, CellValue::Text("$E$6".into()));
    assert_eq!(s.doc()?.wb.active().ok_or("sheet")?.view_top_left, CellRef::parse("C4").ok_or("cell")?);
    assert_eq!(s.doc()?.go_to_recent.first().map(String::as_str), Some("Sheet1!E6"));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_previous_selections_exposes_the_four_displaced_goto_ranges() -> R {
    let source = r#"Sub Main()
Dim previous As Range
Application.Goto Range("C4")
Application.Goto Range("E6")
Application.Goto Range("G8")
Application.Goto Range("I10")
Application.Goto Range("K12")
Set previous = Application.PreviousSelections(1)
Range("A1") = previous.Address
Set previous = Application.PreviousSelections(4)
Range("A2") = previous.Address
Range("A3") = UBound(Application.PreviousSelections)
On Error Resume Next
Set previous = Application.PreviousSelections(5)
Range("A4") = Err.Number
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "A1")?, CellValue::Text("$I$10".into()));
    assert_eq!(cell(&s, "A2")?, CellValue::Text("$C$4".into()));
    assert_eq!(cell(&s, "A3")?, CellValue::Number(4.));
    assert_eq!(cell(&s, "A4")?, CellValue::Number(9.));
    assert_eq!(s.doc()?.previous_selections.len(), 4);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_select_replace_false_tracks_a_real_selected_tab_group() -> R {
    let source = r#"Sub Main()
Worksheets(2).Select
Worksheets(3).Select Replace:=False
End Sub"#;
    let (mut s, _, _) = setup(source)?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Second"}))?;
    s.execute("home.insertSheet", serde_json::json!({"name":"Third"}))?;
    s.execute("sheet.activate", serde_json::json!({"sheet":0}))?;
    let mut b = Bridge::new(&s)?;
    b.consent(true);
    let mut ide = Ide::new();
    ide.add_module("Grouped.bas", source)?;
    b.configure(&mut ide)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.doc()?.sheet_index(), 2);
    assert_eq!(s.doc()?.selected_sheets, vec![1, 2]);
    Ok(())
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
            identity: Default::default(),
            name: "LocalBlock".into(),
            scope: Some(0),
            formula: "$D$8:$E$9".into(),
            comment: String::new(),
            hidden: false,
        },
        gridcraft_engine::model::DefinedName {
            identity: Default::default(),
            name: "LocalBlock".into(),
            scope: None,
            formula: "$Z$1".into(),
            comment: String::new(),
            hidden: false,
        },
        gridcraft_engine::model::DefinedName {
            identity: Default::default(),
            name: "GlobalBlock".into(),
            scope: None,
            formula: "LocalBlock".into(),
            comment: String::new(),
            hidden: false,
        },
        gridcraft_engine::model::DefinedName {
            identity: Default::default(),
            name: "Cycle".into(),
            scope: None,
            formula: "Cycle".into(),
            comment: String::new(),
            hidden: false,
        },
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
fn partial_cell_insert_and_delete_use_excel_vertical_defaults() -> R {
    let source = r#"Sub Main()
Range("A1") = 1
Range("A2") = 2
Range("A2").Insert
Range("B1") = Range("A3")
Range("A2").Delete
Range("B2") = Range("A2")
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    assert_eq!(cell(&s, "B1")?, CellValue::Number(2.));
    assert_eq!(cell(&s, "B2")?, CellValue::Number(2.));
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
