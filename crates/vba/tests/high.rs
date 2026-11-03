use gridcraft_engine::{
    Session,
    core::{CellRef, Value as CellValue},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn session() -> Session {
    let mut s = Session::new();
    s.new_workbook();
    s
}
fn run(s: &mut Session, source: &str) -> R {
    run_locale(s, source, gridcraft_engine::formula::FormulaLocale::En)
}
fn run_locale(s: &mut Session, source: &str, locale: gridcraft_engine::formula::FormulaLocale) -> R {
    let mut bridge = Bridge::new(s)?;
    bridge.set_formula_locale(locale);
    let mut ide = Ide::new();
    ide.add_module("Main.bas", source)?;
    bridge.configure(&mut ide)?;
    bridge.dispatch(s, &mut ide, Action::Run)?;
    for _ in 0..20000 {
        bridge.tick(s, &mut ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
fn value(s: &Session, addr: &str) -> CellValue {
    s.doc().unwrap().wb.sheet(0).unwrap().value(CellRef::parse(addr).unwrap())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn formula_function_proxy_preserves_references_arrays_errors_and_names() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Option Explicit
Sub Main()
Dim n As Object
Dim r As Excel.Range
Dim names As Excel.Names
Dim v As Variant
Range("A1").Value2 = 1
Range("A2").Value2 = 2
Range("A3").Value2 = 3
Set names = ThisWorkbook.Names
Set n = names.Add(Name:="Rate", RefersTo:="=0.2", Visible:=False)
Range("B1").Value2 = Application.WorksheetFunction.Sum(Range("A1:A3"))
Range("B2").Value2 = Evaluate("Rate*2")
Range("B3").Value2 = n.Visible
Set r = Application.Evaluate("A1:A3")
Range("B4").Value2 = r.Count
v = Evaluate("{10,20;30,40}")
Range("D1:E2").Value2 = v
Range("B5").Value2 = Application.WorksheetFunction.Sum(v)
n.Name = "RenamedRate"
Range("B6").Value2 = Evaluate("RenamedRate")
Range("B7").Value2 = names.Count
On Error Resume Next
v = Application.WorksheetFunction.Match(99, Range("A1:A3"), 0)
Range("B8").Value2 = Err.Number
Err.Clear
n.Delete
names.Add Name:="RenamedRate", RefersTo:="=9"
v = n.Name
Range("B9").Value2 = Err.Number
End Sub
"#,
    )?;
    for (addr, n) in [("B1", 6.0), ("B2", 0.4), ("B4", 3.0), ("B5", 100.0), ("B6", 0.2), ("B7", 1.0), ("B8", 1004.0), ("B9", 91.0), ("E2", 40.0)] {
        assert_eq!(value(&s, addr), CellValue::Number(n), "{addr}");
    }
    assert_eq!(value(&s, "B3"), CellValue::Bool(false));
    assert_eq!(s.doc()?.undo.len(), 1);
    s.execute("edit.undo", json!({}))?;
    assert!(value(&s, "A1").is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn formula_r1c1_is_relative_per_destination_and_evaluate_uses_explicit_sheet() -> R {
    let mut s = session();
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    run(
        &mut s,
        r#"Sub Main()
Dim ws As Excel.Worksheet
Set ws = Worksheets(1)
ws.Range("A1").Value2 = 7
ws.Range("B1:B2").FormulaR1C1 = "=RC[-1]+2"
ws.Range("C1").NumberFormat = "@"
ws.Range("C1").Value2 = ws.Range("B1").FormulaR1C1
ws.Range("C2").Value2 = ws.Evaluate("SUM(A1:B1)")
ws.Range("C3").Value2 = ws.Range("B1").HasFormula
ws.Range("D1").Formula2 = "=SEQUENCE(2)"
End Sub
"#,
    )?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.value(CellRef::new(0, 1)), CellValue::Number(9.0));
    assert_eq!(sh.value(CellRef::new(1, 1)), CellValue::Number(2.0));
    assert_eq!(sh.value(CellRef::new(0, 2)), CellValue::Text("=RC[-1]+2".into()));
    assert_eq!(sh.value(CellRef::new(1, 2)), CellValue::Number(16.0));
    assert_eq!(sh.value(CellRef::new(2, 2)), CellValue::Bool(true));
    assert_eq!(sh.value(CellRef::new(1, 3)), CellValue::Number(2.0));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn formatting_uses_native_styles_theme_geometry_borders_merges_and_rules() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
Dim r As Excel.Range
Dim rule As Excel.FormatCondition
Set r = Range("B2:C3")
r.Value2 = 12
r.HorizontalAlignment = xlCenter
r.VerticalAlignment = xlTop
r.WrapText = True
r.Font.Underline = xlUnderlineStyleDouble
r.Font.Strikethrough = True
r.Interior.ThemeColor = xlThemeColorAccent1
r.Interior.TintAndShade = 0.25
r.Borders(xlEdgeBottom).LineStyle = xlContinuous
r.Borders(xlEdgeBottom).Color = 255
r.RowHeight = 30
r.ColumnWidth = 15
Range("E1:F2").Merge
Range("H1").Value2 = Range("F2").MergeArea.Count
Range("H2").Value2 = r.WrapText
Range("H3").Value2 = r.Interior.ThemeColor
Set rule = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:=10)
rule.Font.Bold = True
rule.Interior.Color = 255
rule.StopIfTrue = True
Range("H4").Value2 = r.FormatConditions.Count
Range("H5").Value2 = rule.AppliesTo.Count
Range("H6").Value2 = r.FormatConditions(1).StopIfTrue
rule.ModifyAppliesToRange Range("B2:C4")
Range("H7").Value2 = rule.AppliesTo.Count
End Sub
"#,
    )?;
    for (addr, n) in [("H1", 4.0), ("H3", 5.0), ("H4", 1.0), ("H5", 4.0), ("H7", 6.0)] {
        assert_eq!(value(&s, addr), CellValue::Number(n), "{addr}");
    }
    assert_eq!(value(&s, "H2"), CellValue::Bool(true));
    assert_eq!(value(&s, "H6"), CellValue::Bool(true));
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("sheet")?;
    let style = doc.wb.styles.get(sh.style_id(CellRef::new(1, 1)));
    assert!(style.font.strike);
    assert_eq!(style.align.h, gridcraft_engine::model::HAlign::Center);
    assert_eq!(style.fill.fg, gridcraft_engine::model::Color::Theme(4, 250));
    assert_eq!(doc.undo.len(), 1);
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.cond_formats.is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn persistent_sort_filter_and_validation_use_explicit_inactive_worksheet() -> R {
    let mut s = session();
    s.execute("range.setValues", json!({"range":"A1:B4","values":[["Name","Amount"],["c",3],["a",1],["b",2]]}))?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    run(
        &mut s,
        r#"Sub Main()
Dim ws As Excel.Worksheet
Dim sorter As Excel.Sort
Set ws = ThisWorkbook.Worksheets(1)
Set sorter = ws.Sort
sorter.SortFields.Clear
sorter.SortFields.Add Key:=ws.Range("B2:B4"), Order:=xlDescending
ws.Sort.SetRange ws.Range("A1:B4")
ws.Sort.Header = xlYes
ws.Sort.Apply
ws.Range("D1").Value2 = sorter.SortFields.Count
ws.Range("A1:B4").AutoFilter Field:=2, Criteria1:=">1"
ws.Range("D2").Value2 = ws.AutoFilter.Filters.Item(2).On
ws.Range("D3").Value2 = ws.FilterMode
ws.ShowAllData
ws.Range("D4").Value2 = ws.FilterMode
ws.Range("B2:B4").Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=1, Formula2:=10
ws.Range("B2:B4").Validation.InputTitle = "Amount"
ws.Range("D5").Value2 = ws.Range("B2:B4").Validation.Type
ws.Range("D6").Value2 = ws.Range("B2:B4").Validation.InputTitle
On Error Resume Next
ws.Sort.SortMethod = xlPinYin
ws.Range("D7").Value2 = Err.Number
End Sub
"#,
    )?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    for (addr, n) in [("B2", 3.0), ("B3", 2.0), ("B4", 1.0), ("D1", 1.0), ("D5", 1.0), ("D7", 1004.0)] {
        assert_eq!(sh.value(CellRef::parse(addr).ok_or("address")?), CellValue::Number(n), "{addr}");
    }
    assert_eq!(sh.value(CellRef::new(1, 3)), CellValue::Bool(true));
    assert_eq!(sh.value(CellRef::new(2, 3)), CellValue::Bool(true));
    assert_eq!(sh.value(CellRef::new(3, 3)), CellValue::Bool(false));
    assert_eq!(sh.value(CellRef::new(5, 3)), CellValue::Text("Amount".into()));
    assert_eq!(s.doc()?.sheet_index(), 1);
    assert!(s.take_ui_requests().is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn structured_tables_are_live_across_rename_resize_insert_delete_totals_and_unlist() -> R {
    let mut s = session();
    s.execute("range.setValues", json!({"range":"A1:B3","values":[["Name","Amount"],["a",1],["b",2]]}))?;
    run(
        &mut s,
        r#"Sub Main()
Dim table As Excel.ListObject
Dim col As Excel.ListColumn
Dim row As Excel.ListRow
Set table = Worksheets(1).ListObjects.Add(SourceType:=xlSrcRange, Source:=Range("A1:B3"), XlListObjectHasHeaders:=xlYes)
table.Name = "Sales"
Set col = table.ListColumns("Amount")
Range("E1").Value2 = WorksheetFunction.Sum(Range("Sales[Amount]"))
Set row = table.ListRows.Add(AlwaysInsert:=False)
row.Range.Cells(1, 1).Value2 = "c"
row.Range.Cells(1, 2).Value2 = 3
Range("E2").Value2 = table.ListRows.Count
col.Name = "Total"
Range("E3").Value2 = col.Index
Range("E4").Value2 = WorksheetFunction.Sum(Range("Sales[Total]"))
table.ShowTotals = True
col.TotalsCalculation = xlTotalsCalculationSum
Range("E5").Value2 = table.TotalsRowRange.Cells(1,2).Value2
table.ShowTableStyleRowStripes = False
table.TableStyle = "TableStyleLight1"
row.Delete
Range("E6").Value2 = table.ListRows.Count
table.Unlist
Range("E7").Value2 = Worksheets(1).ListObjects.Count
On Error Resume Next
Range("E8").Value2 = table.Name
Range("E9").Value2 = Err.Number
End Sub
"#,
    )?;
    for (addr, n) in [("E1", 3.0), ("E2", 3.0), ("E3", 2.0), ("E4", 6.0), ("E5", 6.0), ("E6", 2.0), ("E7", 0.0), ("E9", 91.0)] {
        assert_eq!(value(&s, addr), CellValue::Number(n), "{addr}");
    }
    assert_eq!(s.doc()?.undo.len(), 2);
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.tables.is_empty());
    assert_eq!(value(&s, "B2"), CellValue::Number(1.0));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn condition_specific_graphs_follow_native_rules_and_priority_identity() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
Dim scale As Excel.ColorScale
Dim bar As Excel.Databar
Dim icons As Excel.IconSetCondition
Dim top As Excel.Top10
Set scale = Range("A1:A5").FormatConditions.AddColorScale(3)
scale.ColorScaleCriteria(2).Type = xlConditionValueNumber
scale.ColorScaleCriteria(2).Value = 5
scale.ColorScaleCriteria(2).FormatColor.Color = 255
Range("J1").Value2 = scale.ColorScaleCriteria.Count
Range("J2").Value2 = scale.ColorScaleCriteria(2).Value
Set bar = Range("B1:B5").FormatConditions.AddDatabar
bar.MinPoint.Modify xlConditionValueNumber, 0
bar.MaxPoint.Modify xlConditionValueNumber, 100
bar.BarColor.ThemeColor = xlThemeColorAccent2
bar.BarColor.TintAndShade = 0.2
Range("J3").Value2 = bar.MaxPoint.Value
Range("J4").Value2 = bar.BarColor.ThemeColor
Set icons = Range("C1:C5").FormatConditions.AddIconSetCondition
Set icons.IconSet = ThisWorkbook.IconSets(xl5Arrows)
icons.IconCriteria(2).Type = xlConditionValueNumber
icons.IconCriteria(2).Value = 10
Range("J5").Value2 = icons.IconCriteria.Count
Range("J6").Value2 = icons.IconSet.ID
Set top = Range("D1:D5").FormatConditions.AddTop10
top.Rank = 3
top.TopBottom = xlTop10Bottom
top.SetFirstPriority
Range("J7").Value2 = bar.MaxPoint.Value
Range("J8").Value2 = top.TopBottom
scale.Delete
On Error Resume Next
Range("J9").Value2 = scale.ColorScaleCriteria.Count
Range("J10").Value2 = Err.Number
End Sub
"#,
    )?;
    for (a, n) in [("J1", 3.0), ("J2", 5.0), ("J3", 100.0), ("J4", 6.0), ("J5", 5.0), ("J6", 14.0), ("J7", 100.0), ("J8", 0.0), ("J10", 91.0)] {
        assert_eq!(value(&s, a), CellValue::Number(n), "{a}");
    }
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.cond_formats.len(), 3);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn formula_arrays_scoped_names_and_named_styles_use_native_state() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
Dim n As Object
Dim style As Excel.Style
Range("A1:A2").FormulaArray = "=ROW(A1:A2)*2"
Range("J1").Value2 = Range("A2").CurrentArray.Count
Range("J2").Value2 = Range("A1:A2").HasArray
Range("C1").Font.Bold = True
Set style = ThisWorkbook.Styles.Add("Emphasis", Range("C1"))
Range("D1").Style = style
Range("J3").Value2 = Range("D1").Font.Bold
Range("J4").Value2 = ThisWorkbook.Styles("Emphasis").BuiltIn
Set n = Worksheets(1).Names.Add(Name:="LocalRate", RefersToR1C1:="=R1C1")
Range("J5").Value2 = n.RefersToRange.Count
Range("J6").Value2 = Worksheets(1).Evaluate("LocalRate")
Range("J7").NumberFormat = "@"
Range("J7").Value2 = n.RefersToR1C1
End Sub
"#,
    )?;
    assert_eq!(value(&s, "A1"), CellValue::Number(2.0));
    assert_eq!(value(&s, "A2"), CellValue::Number(4.0));
    assert_eq!(value(&s, "J1"), CellValue::Number(2.0));
    assert_eq!(value(&s, "J2"), CellValue::Bool(true));
    assert_eq!(value(&s, "J3"), CellValue::Bool(true));
    assert_eq!(value(&s, "J4"), CellValue::Bool(false));
    assert_eq!(value(&s, "J5"), CellValue::Number(1.0));
    // Evaluate references return Range; implicit Value2 coercion reads the native cell.
    assert_eq!(value(&s, "J6"), CellValue::Number(2.0));
    assert_eq!(value(&s, "J7"), CellValue::Text("=R1C1".into()));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn scenarios_goal_seek_and_data_tables_preserve_typed_values() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
Dim sc As Excel.Scenario
Dim ok As Boolean
Range("A1").Value2 = 1
Range("B1").Formula = "=A1*2"
ok = Range("B1").GoalSeek(10, Range("A1"))
Range("J1").Value2 = ok
Range("J2").Value2 = Range("A1").Value2
Set sc = Worksheets(1).Scenarios.Add("Base", Range("A1"), 7, "baseline")
sc.Name = "Renamed"
sc.Comment = "edited"
sc.Show
Range("J3").Value2 = Range("A1").Value2
Range("J4").Value2 = sc.Values(1)
sc.ChangeScenario Values:="001"
sc.Show
Range("J5").NumberFormat = "@"
Range("J5").Value2 = Range("A1").Value2
sc.ChangeScenario Values:=CVErr(2042)
sc.Show
Range("J11").Value2 = sc.Values(1)
Range("A1").Value2 = 3
Range("D1").Formula = "=A1*2"
Range("C2").Value2 = 4
Range("C3").Value2 = 5
Range("C1:D3").Table ColumnInput:=Range("A1")
Range("J6").Value2 = Range("D2").Value2
Range("J7").Value2 = Range("D3").Value2
Range("J8").Value2 = Range("A1").Value2
sc.Delete
On Error Resume Next
Range("J9").Value2 = sc.Name
Range("J10").Value2 = Err.Number
End Sub
"#,
    )?;
    assert_eq!(value(&s, "J1"), CellValue::Bool(true));
    for (a, n) in [("J2", 5.0), ("J3", 7.0), ("J4", 7.0), ("J6", 8.0), ("J7", 10.0), ("J8", 3.0), ("J10", 91.0)] {
        assert_eq!(value(&s, a), CellValue::Number(n), "{a}");
    }
    assert_eq!(value(&s, "J5"), CellValue::Text("001".into()));
    assert_eq!(value(&s, "J11"), CellValue::Error(gridcraft_engine::core::CellError::NA));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn localized_formulas_round_trip_without_changing_canonical_source() -> R {
    let mut s = session();
    run_locale(
        &mut s,
        r#"Sub Main()
Dim n As Object
Range("A1").Value2 = 2
Range("B1").FormulaLocal = "=SUMA(A1;1,5)"
Range("B2").FormulaR1C1Local = "=SUMA(R[-1]C[-1];2,5)"
Range("J1:J4").NumberFormat = "@"
Range("J1").Value2 = Range("B1").Formula
Range("J2").Value2 = Range("B1").FormulaLocal
Range("J3").Value2 = Range("B2").FormulaR1C1Local
Set n = ThisWorkbook.Names.Add(NameLocal:="Tarifa", RefersToLocal:="=SUMA(1;2,5)")
Range("J4").Value2 = n.RefersToLocal
End Sub
"#,
        gridcraft_engine::formula::FormulaLocale::Es,
    )?;
    assert_eq!(value(&s, "B1"), CellValue::Number(3.5));
    assert_eq!(value(&s, "B2"), CellValue::Number(4.5));
    assert_eq!(value(&s, "J1"), CellValue::Text("=SUM(A1,1.5)".into()));
    assert_eq!(value(&s, "J2"), CellValue::Text("=SUMA(A1;1,5)".into()));
    assert_eq!(value(&s, "J3"), CellValue::Text("=SUMA(R[-1]C[-1];2,5)".into()));
    assert_eq!(value(&s, "J4"), CellValue::Text("=SUMA(1;2,5)".into()));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn table_sort_and_filters_remain_table_owned_and_clear_without_stuck_hidden_rows() -> R {
    let mut s = session();
    s.execute("range.setValues", json!({"range":"A1:B4","values":[["Name","Amount"],["c",3],["a",1],["b",2]]}))?;
    run(
        &mut s,
        r#"Sub Main()
Dim t As Excel.ListObject
Dim sorter As Excel.Sort
Set t = Worksheets(1).ListObjects.Add(xlSrcRange, Range("A1:B4"), , xlYes)
Set sorter = t.Sort
sorter.SortFields.Add Key:=t.ListColumns(2).DataBodyRange, Order:=xlDescending
t.Sort.Apply
Range("J1").Value2 = t.Sort.SortFields.Count
t.Range.AutoFilter Field:=2, Criteria1:=">1"
Range("J2").Value2 = t.AutoFilter.Filters(2).On
Range("J3").Value2 = Range("A4").EntireRow.Hidden
t.AutoFilter.ShowAllData
Range("J4").Value2 = Range("A4").EntireRow.Hidden
Range("J5").Value2 = t.AutoFilter.FilterMode
End Sub
"#,
    )?;
    assert_eq!(value(&s, "B2"), CellValue::Number(3.0));
    assert_eq!(value(&s, "B4"), CellValue::Number(1.0));
    assert_eq!(value(&s, "J1"), CellValue::Number(1.0));
    assert_eq!(value(&s, "J2"), CellValue::Bool(true));
    assert_eq!(value(&s, "J3"), CellValue::Bool(true));
    assert_eq!(value(&s, "J4"), CellValue::Bool(false));
    assert_eq!(value(&s, "J5"), CellValue::Bool(false));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn data_tool_writers_preview_exact_output_and_preserve_native_results() -> R {
    let mut s = session();
    s.execute("range.setValues", json!({"range":"A1:B5","values":[["Name","Amount"],["a",1],["b",2],["a",1],["c",3]]}))?;
    let undo = s.doc()?.undo.len();
    let plan = s.preview_data_mutation("data.textToColumns", &json!({"range":"A1:A5","destination":"G1","delimiters":[","]}))?;
    assert_eq!(plan, Some((0, gridcraft_engine::core::RangeRef::parse("G1:G5").ok_or("range")?)));
    assert_eq!(s.doc()?.undo.len(), undo);
    run(
        &mut s,
        r#"Sub Main()
Range("A1:B5").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("D1"), Unique:=True
Range("G1").Consolidate Sources:=Array("B2:B3", "B4:B5"), Function:=xlSum
Range("A1:B5").RemoveDuplicates Columns:=Array(1,2), Header:=xlYes
Range("J1").Value2 = "1,2,3"
Range("J1").TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, Tab:=False, Comma:=True
End Sub
"#,
    )?;
    for (a, n) in [("E2", 1.0), ("E3", 2.0), ("E4", 3.0), ("G1", 2.0), ("G2", 5.0), ("K1", 1.0), ("L1", 2.0), ("M1", 3.0)] {
        assert_eq!(value(&s, a), CellValue::Number(n), "{a}");
    }
    assert!(value(&s, "A5").is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn invalid_sort_and_protected_high_operations_have_no_partial_writes() -> R {
    let mut s = session();
    s.execute("range.setValues", json!({"range":"A1:A3","values":[[3],[2],[1]]}))?;
    run(
        &mut s,
        r#"Sub Main()
On Error Resume Next
Worksheets(1).Sort.SetRange Range("A1:A3")
Worksheets(1).Sort.Apply
Range("J1").Value2 = Err.Number
Err.Clear
Range("A1:A3").FormulaR1C1 = "=R[-100]C"
Range("J2").Value2 = Err.Number
Err.Clear
Range("A1:A3").FormatConditions.AddColorScale 4
Range("J3").Value2 = Err.Number
Worksheets(1).Protect
Err.Clear
Range("A1").Font.Bold = True
Err.Clear
Worksheets(1).ListObjects.Add Source:=Range("A1:A3")
End Sub
"#,
    )?;
    for (a, n) in [("A1", 3.0), ("A2", 2.0), ("A3", 1.0), ("J1", 1004.0), ("J2", 1004.0), ("J3", 1004.0)] {
        assert_eq!(value(&s, a), CellValue::Number(n), "{a}");
    }
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("sheet")?;
    assert!(sh.tables.is_empty());
    assert!(sh.cond_formats.is_empty());
    assert!(!doc.wb.styles.get(sh.style_id(CellRef::new(0, 0))).font.bold);
    assert!(s.take_ui_requests().is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn high_native_state_survives_xlsx_round_trip() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
Dim t As Excel.ListObject
Dim sc As Excel.Scenario
Dim bar As Excel.Databar
Range("A1").Value2 = "Amount"
Range("A2").Value2 = 7
Set t = Worksheets(1).ListObjects.Add(xlSrcRange, Range("A1:A2"), , xlYes)
t.Name = "Amounts"
t.ShowTotals = True
t.ListColumns(1).TotalsCalculation = xlTotalsCalculationSum
Set sc = Worksheets(1).Scenarios.Add("Base", Range("A2"))
Set bar = Range("D1:D5").FormatConditions.AddDatabar
bar.MaxPoint.Modify xlConditionValueNumber, 20
Range("D1:D5").Interior.ThemeColor = xlThemeColorAccent3
Range("D1:D5").Validation.Add Type:=xlValidateDecimal, Operator:=xlBetween, Formula1:=0, Formula2:=20
ThisWorkbook.Names.Add Name:="Total", RefersTo:="=SUM(Amounts[Amount])"
End Sub
"#,
    )?;
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (book, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let sh = book.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.tables[0].name, "Amounts");
    assert!(sh.tables[0].totals_row);
    assert_eq!(sh.cond_formats.len(), 1);
    assert_eq!(sh.validations.len(), 1);
    assert!(book.name("Total", 0).is_some());
    assert_eq!(gridcraft_engine::cmd::whatif::scenario_objects(&book).len(), 1);
    assert_eq!(book.styles.get(sh.style_id(CellRef::new(0, 3))).fill.fg, gridcraft_engine::model::Color::Theme(6, 0));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn inactive_outline_and_subtotals_use_the_native_owner_and_planner() -> R {
    let mut s = session();
    s.execute("range.setValues", json!({"range":"A1:B4","values":[["Group","Amount"],["a",1],["a",2],["b",3]]}))?;
    let plan = s.preview_data_mutation("data.subtotal", &json!({"range":"A1:B4","groupBy":"A","columns":["B"]}))?;
    assert_eq!(plan, Some((0, gridcraft_engine::core::RangeRef::parse("A1:B7").ok_or("range")?)));
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    run(
        &mut s,
        r#"Sub Main()
Dim ws As Excel.Worksheet
Dim info(1,1) As Variant
info(0,0) = 0
info(0,1) = xlGeneralFormat
info(1,0) = 3
info(1,1) = xlGeneralFormat
Set ws = Worksheets(1)
ws.Rows("2:3").Group
ws.Range("J1").Value2 = ws.Rows("2:3").OutlineLevel
ws.Outline.ShowLevels RowLevels:=1
ws.Range("J2").Value2 = ws.Rows(2).Hidden
ws.Outline.ShowLevels RowLevels:=2
ws.Range("J3").Value2 = ws.Rows(2).Hidden
ws.Rows("2:3").Ungroup
ws.Range("A1:B4").Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(2), Replace:=False, PageBreaks:=False, SummaryBelowData:=xlSummaryBelow
ws.Range("D1").Value2 = "abc123"
ws.Range("D1").TextToColumns Destination:=ws.Range("E1"), DataType:=xlFixedWidth, FieldInfo:=info
End Sub
"#,
    )?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.value(CellRef::parse("J1").ok_or("addr")?), CellValue::Number(2.0));
    // Summary insertions move the original J2/J3 along with their rows.
    assert_eq!(sh.value(CellRef::parse("J2").ok_or("addr")?), CellValue::Bool(true));
    assert_eq!(sh.value(CellRef::parse("J3").ok_or("addr")?), CellValue::Bool(false));
    assert_eq!(sh.value(CellRef::parse("B4").ok_or("addr")?), CellValue::Number(3.0));
    assert_eq!(sh.value(CellRef::parse("B7").ok_or("addr")?), CellValue::Number(6.0));
    assert_eq!(sh.value(CellRef::parse("E1").ok_or("addr")?), CellValue::Text("abc".into()));
    assert_eq!(sh.value(CellRef::parse("F1").ok_or("addr")?), CellValue::Number(123.0));
    assert_eq!(s.doc()?.sheet_index(), 1);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn rectangular_matrices_preserve_vba_subscripts_across_formula_and_value_boundaries() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
Dim data(1 To 2, 1 To 3) As Variant
Dim result As Variant
data(1,1) = 1
data(1,2) = 2
data(1,3) = 3
data(2,1) = 4
data(2,2) = 5
data(2,3) = 6
Range("A1:C2").Value2 = data
result = Range("A1:C2").Value2
Range("J1").Value2 = result(1,3)
Range("J2").Value2 = result(2,1)
result = Evaluate("{10,20,30;40,50,60}")
Range("J3").Value2 = result(1,3)
Range("J4").Value2 = result(2,1)
Range("D1:F2").Value2 = result
result = WorksheetFunction.Transpose(data)
Range("J5").Value2 = result(3,2)
Range("J6").Value2 = result(1,2)
End Sub
"#,
    )?;
    for (a, n) in
        [("C1", 3.0), ("A2", 4.0), ("J1", 3.0), ("J2", 4.0), ("J3", 30.0), ("J4", 40.0), ("F1", 30.0), ("D2", 40.0), ("J5", 6.0), ("J6", 4.0)]
    {
        assert_eq!(value(&s, a), CellValue::Number(n), "{a}");
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_data_writer_events_report_the_computed_destination() -> R {
    for (statement, expected) in [
        ("Range(\"A1:B4\").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range(\"D1\")", "$D$1:$E$4"),
        ("Range(\"G1\").TextToColumns Destination:=Range(\"H1\"), Tab:=False, Comma:=True", "$H$1:$J$1"),
    ] {
        let mut s = session();
        s.execute("range.setValues", json!({"range":"A1:B4","values":[["Name","Amount"],["a",1],["b",2],["c",3]]}))?;
        s.execute("cell.set", json!({"cell":"G1","input":"1,2,3"}))?;
        std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.code_name = Some("SheetCode".into());
        let mut bridge = Bridge::new(&s)?;
        let mut ide = Ide::new();
        ide.add_module("Main.bas", &format!("Sub Main()\n{statement}\nEnd Sub\n"))?;
        ide.add_document_module(
            "Sheet1.cls",
            "SheetCode",
            r#"Attribute VB_Name = "SheetCode"
Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
Me.Range("M1").Value2 = Target.Address
End Sub
"#,
        )?;
        bridge.configure(&mut ide)?;
        bridge.dispatch(&mut s, &mut ide, Action::Run)?;
        for _ in 0..20000 {
            bridge.tick(&mut s, &mut ide)?;
            if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
                break;
            }
        }
        assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
        assert_eq!(value(&s, "M1"), CellValue::Text(expected.into()));
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn table_views_do_not_retarget_when_native_table_ids_are_reused() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
Dim t As Excel.ListObject
Dim oldSort As Excel.Sort
Dim oldField As Excel.SortField
Dim oldFilter As Excel.AutoFilter
Range("A1").Value2 = "Amount"
Range("A2").Value2 = 2
Set t = Worksheets(1).ListObjects.Add(Source:=Range("A1:A2"), XlListObjectHasHeaders:=xlYes)
Set oldSort = t.Sort
Set oldField = oldSort.SortFields.Add(t.DataBodyRange)
Set oldFilter = t.AutoFilter
t.Unlist
Set t = Worksheets(1).ListObjects.Add(Source:=Range("A1:A2"), XlListObjectHasHeaders:=xlYes)
t.Sort.SortFields.Add t.DataBodyRange
On Error Resume Next
Range("J1").Value2 = oldSort.SortFields.Count
Range("J2").Value2 = Err.Number
Err.Clear
Range("J3").Value2 = oldFilter.Filters.Count
Range("J4").Value2 = Err.Number
Err.Clear
Range("J5").Value2 = oldField.Order
Range("J6").Value2 = Err.Number
End Sub
"#,
    )?;
    for a in ["J2", "J4", "J6"] {
        assert_eq!(value(&s, a), CellValue::Number(91.0), "{a}");
    }
    Ok(())
}
