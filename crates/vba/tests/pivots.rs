//! Original Excel-shaped chart sources; all effects use native engine chart commands.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value as CellValue},
    model::{PivotFunc, PivotLayout},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    s.execute(
        "range.setValues",
        json!({"range":"A1:C5","values":[["Region","Product","Sales"],["East","A",10],["West","A",20],["East","B",30],["West","B",40]]}),
    )?;
    s.execute("home.insertSheet", json!({"name":"Report"}))?;
    s.execute("insert.pivotTable", json!({"source":"Sheet1!A1:C5","destination":"Report!A3","name":"Totals","rows":["Region"],"values":["Sales"]}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Charts.bas", source)?;
    b.configure(&mut ide)?;
    Ok((s, b, ide))
}
fn finish(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    for _ in 0..1000 {
        b.tick(s, ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) && !b.has_pending_events() {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn typed_pivot_graph_and_inactive_report_settings_reuse_native_output() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Option Explicit
Sub Main()
Dim pt As PivotTable
Dim cache As PivotCache
Dim axis As PivotField
Dim amount As PivotField
Set pt=Worksheets("Report").PivotTables("Totals")
Set cache=pt.PivotCache
If Not (cache Is ThisWorkbook.PivotCaches(pt.CacheIndex)) Then Err.Raise 5
If Not (cache.Parent Is ThisWorkbook) Then Err.Raise 5
Set axis=pt.PivotFields("Product")
axis.Orientation=xlColumnField
Set amount=pt.DataFields(1)
amount.Name="Amount"
amount.Function=xlAverage
amount.NumberFormat="#,##0.00"
pt.RowAxisLayout xlTabularRow
pt.RowGrand=False
pt.ColumnGrand=True
If Not (amount.Parent Is pt) Then Err.Raise 5
If Not (pt.RowFields(1).Parent Is pt) Then Err.Raise 5
Worksheets(1).Range("H1").Value2=cache.RecordCount
Worksheets(1).Range("H2").Value2=pt.TableRange2.Address
Worksheets(1).Range("H3").Value2=pt.RefreshTable
Worksheets(1).Range("H4").Value2=pt.DataBodyRange.Address
End Sub
"##,
    )?;
    let undo = s.doc()?.undo.len();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let d = s.doc()?;
    assert_eq!(d.sheet_index(), 0);
    assert_eq!(d.undo.len(), undo + 1);
    let sh = d.wb.sheet(1).ok_or("report")?;
    let pt = &sh.pivots[0];
    assert_eq!(pt.layout, PivotLayout::Tabular);
    assert_eq!(pt.columns[0].source_col, "Product");
    assert_eq!(pt.values[0].name, "Amount");
    assert_eq!(pt.values[0].func, PivotFunc::Average);
    assert!(!pt.grand_totals_rows);
    assert!(pt.grand_totals_cols);
    assert_eq!(d.wb.sheet(0).ok_or("source")?.value(CellRef::new(0, 7)), CellValue::Number(4.0));
    let out = gridcraft_engine::pivot::compute(&d.wb, pt, 1)?;
    let area = pt.last_range.ok_or("area")?;
    for (ri, row) in out.grid.iter().enumerate() {
        for (ci, c) in row.iter().enumerate() {
            assert_eq!(sh.value(CellRef::new(area.start.row + ri as u32, area.start.col + ci as u32)), c.value);
        }
    }
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.sheet(1).ok_or("report")?.pivots[0].columns.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn filters_calculated_fields_and_stale_children_use_native_lifetimes() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Sub Main()
Dim pt As PivotTable
Dim region As PivotField
Dim old As PivotField
Dim parked As PivotField
Dim amount As PivotField
Set pt=Worksheets(2).PivotTables(1)
Set region=pt.PivotFields("Region")
region.PivotItems("West").Visible=False
If region.VisibleItems.Count <> 1 Then Err.Raise 5
If region.HiddenItems.Count <> 1 Then Err.Raise 5
If Not (region.VisibleItems(1).Parent Is region) Then Err.Raise 5
Worksheets(1).Range("H1").Value2=region.PivotItems.Count
Set old=pt.CalculatedFields.Add("Tax","=Sales*0.1",True)
Set amount=pt.AddDataField(old,"Tax total",xlSum)
old.Formula="=Sales*0.2"
amount.Position=1
Worksheets(1).Range("H2").Value2=amount.Name
Worksheets(1).Range("H3").Value2=amount.Position
Set parked=old
old.Delete
Set old=pt.CalculatedFields.Add("Tax","=Sales*0.5")
region.ClearAllFilters
region.AutoSort xlDescending,"Region"
Worksheets(1).Range("H4").Value2=pt.CalculatedFields.Count
On Error Resume Next
parked.Formula="=999"
Worksheets(1).Range("H5").Value2=Err.Number
Err.Clear
pt.AddDataField parked
Worksheets(1).Range("H6").Value2=Err.Number
On Error GoTo 0
Dim page As PivotField
Set page=pt.PivotFields("Product")
page.Orientation=xlPageField
page.CurrentPage="a"
If Not (page.CurrentPage Is page.PivotItems("A")) Then Err.Raise 5
page.ClearAllFilters
End Sub
"##,
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let d = s.doc()?;
    let src = d.wb.sheet(0).ok_or("source")?;
    for (row, v) in [(0, CellValue::Number(2.0)), (1, CellValue::text("Tax total")), (2, CellValue::Number(1.0)), (3, CellValue::Number(1.0))] {
        assert_eq!(src.value(CellRef::new(row, 7)), v);
    }
    let pt = &d.wb.sheet(1).ok_or("report")?.pivots[0];
    assert!(pt.filters.iter().all(|f| f.selected.is_none()));
    assert_eq!(pt.calculated_fields[0].formula, "=Sales*0.5");
    assert_eq!(pt.values.len(), 1);
    for row in 4..6 {
        assert_eq!(src.value(CellRef::new(row, 7)), CellValue::Number(91.0));
    }
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn pivot_creation_source_rebind_and_id_reuse_refuse_stale_views() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Sub Main()
Dim pt As PivotTable
Dim old As PivotField
Set pt=Worksheets(2).PivotTables.Add(xlDatabase,Worksheets(1).Range("A1:C5"),Worksheets(2).Range("H3"),"Second")
Set old=pt.PivotFields("Region")
pt.SourceData="Sheet1!A1:C4"
On Error Resume Next
old.Orientation=xlRowField
Worksheets(1).Range("H1").Value2=Err.Number
Err.Clear
Stop
pt.RefreshTable
Worksheets(1).Range("H2").Value2=Err.Number
End Sub
"##,
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break, "{:?}", ide.diagnostics);
    let id = s.doc()?.wb.sheet(1).ok_or("report")?.pivots[1].id;
    s.execute("pivot.delete", json!({"sheet":1,"pivot":id}))?;
    s.execute("insert.pivotTable", json!({"source":"Sheet1!A1:C5","destination":"Report!H3","name":"Replacement","activate":false}))?;
    assert_eq!(s.doc()?.wb.sheet(1).ok_or("report")?.pivots[1].id, id);
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("source")?;
    assert_eq!(sh.value(CellRef::new(0, 7)), CellValue::Number(91.0));
    assert_eq!(sh.value(CellRef::new(1, 7)), CellValue::Number(91.0));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn drilldown_uses_native_records_options_and_structural_undo() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Sub Main()
Dim pt As PivotTable
Dim cell As PivotCell
Set pt=Worksheets("Report").PivotTables(1)
Set cell=Worksheets("Report").Range("B4").PivotCell
If Not (cell.PivotTable Is pt) Then Err.Raise 5
pt.EnableDrilldown=False
On Error Resume Next
Worksheets("Report").Range("B4").ShowDetail=True
Worksheets("Sheet1").Range("H1").Value2=Err.Number
Err.Clear
On Error GoTo 0
pt.EnableDrilldown=True
Worksheets("Report").Range("B4").ShowDetail=True
Worksheets("Sheet1").Range("H2").Value2=pt.Name
End Sub
"##,
    )?;
    let before = s.doc()?.undo.len();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    assert_eq!(doc.wb.sheets.len(), 3);
    let detail = doc.wb.active().ok_or("detail")?;
    assert_eq!(detail.value(CellRef::new(0, 0)), CellValue::text("Region"));
    assert_eq!(detail.value(CellRef::new(1, 0)), CellValue::text("East"));
    assert_eq!(detail.value(CellRef::new(2, 0)), CellValue::text("East"));
    assert_eq!(detail.used_range().ok_or("used")?.height(), 3);
    let src = doc.wb.sheets.iter().find(|sh| sh.name == "Sheet1").ok_or("source")?;
    assert_eq!(src.value(CellRef::new(0, 7)), CellValue::Number(1004.0));
    assert_eq!(src.value(CellRef::new(1, 7)), CellValue::text("Totals"));
    assert_eq!(doc.undo.len(), before + 1);
    s.execute("edit.undo", json!({}))?;
    assert_eq!(s.doc()?.wb.sheets.len(), 2);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn pivotcharts_and_local_caches_persist_and_render_the_refreshed_report() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Sub Main()
Dim pt As PivotTable
Dim chart As Chart
Set pt=Worksheets(2).PivotTables(1)
Set chart=Worksheets(2).ChartObjects(1).Chart
If Not (chart.PivotLayout.PivotTable Is pt) Then Err.Raise 5
If Not (chart.PivotLayout.Parent Is chart) Then Err.Raise 5
Worksheets(1).Range("C2").Value2=15
pt.PivotCache.Refresh
Worksheets(1).Range("H1").Value2=pt.PivotCache.RecordCount
End Sub
"##,
    )?;
    s.execute("insert.pivotChart", json!({"sheet":1,"pivot":"Totals","at":"E3"}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let sh = doc.wb.sheet(1).ok_or("report")?;
    let data = gridcraft_chart::resolve(&doc.wb, 1, &sh.charts[0]);
    assert_eq!(data.series[0].values, vec![Some(45.0), Some(60.0)]);
    let prims = gridcraft_chart::render(&sh.charts[0], &data, 480.0, 288.0, &gridcraft_chart::ApproxMeasure);
    let bytes = gridcraft_xlsx::write_xlsx(&doc.wb)?;
    let (loaded, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let report = loaded.sheet(1).ok_or("report")?;
    assert_eq!(
        gridcraft_engine::pivot::resolve_source(&loaded, &report.pivots[0].source, 1)?,
        gridcraft_engine::pivot::resolve_source(&doc.wb, &sh.pivots[0].source, 1)?
    );
    assert_eq!(report.pivots[0].values[0].func, sh.pivots[0].values[0].func);
    let restored = gridcraft_chart::resolve(&loaded, 1, &report.charts[0]);
    assert_eq!(restored.series[0].values, data.series[0].values);
    let prims2 = gridcraft_chart::render(&report.charts[0], &restored, 480.0, 288.0, &gridcraft_chart::ApproxMeasure);
    assert_eq!(gridcraft_chart::rasterize(&prims, 480, 288, [255; 4]), gridcraft_chart::rasterize(&prims2, 480, 288, [255; 4]));
    assert!(!loaded.sheet(1).ok_or("report")?.pivots[0].identity.same(&sh.pivots[0].identity));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn table_slicer_views_filter_format_and_persist_on_inactive_owners() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Option Explicit
Sub Main()
Dim cache As SlicerCache
Dim widget As Slicer
Dim item As SlicerItem
Dim shape As Shape
Set cache=ThisWorkbook.SlicerCaches(1)
Set widget=cache.Slicers(1)
Set shape=widget.Shape
If Not (widget.SlicerCache Is cache) Then Err.Raise 5
If Not (shape Is Worksheets(1).Shapes(widget.Name)) Then Err.Raise 5
If Not (cache.ListObject Is Worksheets(1).ListObjects(1)) Then Err.Raise 5
widget.Caption="Choose region"
widget.NumberOfColumns=2
widget.Left=150
widget.Top=80
widget.Width=220
widget.Height=260
Set item=cache.SlicerItems("West")
item.Selected=False
Worksheets(2).Range("M1").Value2=cache.FilterCleared
Worksheets(2).Range("M2").Value2=cache.SlicerItems.Count
Worksheets(2).Range("M3").Value2=shape.Left
End Sub
"##,
    )?;
    s.execute("insert.table", json!({"sheet":0,"range":"A1:C5","name":"SalesTable"}))?;
    s.execute("insert.slicer", json!({"table":"SalesTable","column":"Region"}))?;
    s.execute("sheet.activate", json!({"sheet":1}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    assert_eq!(doc.sheet_index(), 1);
    let sh = doc.wb.sheet(0).ok_or("source")?;
    assert_eq!(sh.slicers.len(), 1);
    let sl = &sh.slicers[0];
    assert_eq!(sl.caption, "Choose region");
    assert_eq!(sl.columns, 2);
    assert_eq!(sl.selected.as_deref(), Some(["East".to_string()].as_slice()));
    assert!(sh.rows.get(&2).is_some_and(|row| row.hidden));
    assert!(!sh.rows.get(&1).is_some_and(|row| row.hidden));
    let report = doc.wb.sheet(1).ok_or("report")?;
    assert_eq!(report.value(CellRef::new(0, 12)), CellValue::Bool(false));
    assert_eq!(report.value(CellRef::new(1, 12)), CellValue::Number(2.0));
    assert_eq!(report.value(CellRef::new(2, 12)), CellValue::Number(150.0));
    let bytes = gridcraft_xlsx::write_xlsx(&doc.wb)?;
    let (loaded, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert!(loaded.sheet(0).ok_or("source")?.slicers.is_empty(), "native XLSX slicer codec is a documented product gap");
    let snapshot = serde_json::to_vec(&doc.wb)?;
    let restored_book: gridcraft_engine::model::Workbook = serde_json::from_slice(&snapshot)?;
    let restored = &restored_book.sheet(0).ok_or("source")?.slicers[0];
    assert_eq!(restored.caption, sl.caption);
    assert_eq!(restored.selected, sl.selected);
    assert_eq!(restored.anchor, sl.anchor);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn atomic_area_edits_values_axis_and_clear_table_keep_real_cache_views() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Sub Main()
Dim pt As PivotTable
Dim extra As PivotField
Set pt=Worksheets(2).PivotTables(1)
Set extra=pt.AddDataField(pt.PivotFields("Sales"),"Count",xlCount)
pt.DataPivotField.Orientation=xlRowField
pt.DataPivotField.Position=1
Worksheets(1).Range("H1").Value2=pt.DataPivotField.Position
On Error Resume Next
pt.AddFields RowFields:=Array("Product","Missing"),AddToTable:=False
Worksheets(1).Range("H2").Value2=Err.Number
Err.Clear
On Error GoTo 0
Worksheets(1).Range("H3").Value2=pt.RowFields(2).Name
pt.AddFields RowFields:="Region",ColumnFields:="Product",AddToTable:=False
pt.SubtotalLocation xlAtBottom
pt.ClearAllFilters
pt.ClearTable
If Not (pt.DataBodyRange Is Nothing) Then Err.Raise 5
Worksheets(1).Range("H4").Value2=pt.PivotCache.RecordCount
On Error Resume Next
extra.Name="Gone"
Worksheets(1).Range("H5").Value2=Err.Number
End Sub
"##,
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("source")?;
    for (row, v) in [
        (0, CellValue::Number(1.0)),
        (1, CellValue::Number(1004.0)),
        (2, CellValue::text("Region")),
        (3, CellValue::Number(4.0)),
        (4, CellValue::Number(91.0)),
    ] {
        assert_eq!(sh.value(CellRef::new(row, 7)), v);
    }
    let pt = &doc.wb.sheet(1).ok_or("report")?.pivots[0];
    assert!(pt.rows.is_empty() && pt.columns.is_empty() && pt.values.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn protection_bad_sources_and_provider_refusals_do_not_partially_refresh() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Sub Main()
Dim pt As PivotTable
Set pt=Worksheets(2).PivotTables(1)
On Error Resume Next
ThisWorkbook.PivotCaches.Create xlExternal,"Unavailable"
Worksheets(1).Range("H1").Value2=Err.Number
Err.Clear
pt.RefreshTable
Worksheets(1).Range("H2").Value2=Err.Number
Err.Clear
pt.PivotFields("Region").Orientation=xlColumnField
Worksheets(1).Range("H3").Value2=Err.Number
Err.Clear
pt.SourceData="Missing!A1:C5"
Worksheets(1).Range("H4").Value2=Err.Number
End Sub
"##,
    )?;
    s.execute("sheet.activate", json!({"sheet":1}))?;
    s.execute("review.protectSheet", json!({}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    let before = s.doc()?.wb.sheet(1).ok_or("report")?.clone();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    assert_eq!(doc.wb.sheet(1).ok_or("report")?, &before);
    let sh = doc.wb.sheet(0).ok_or("source")?;
    for row in 0..4 {
        assert_eq!(sh.value(CellRef::new(row, 7)), CellValue::Number(1004.0));
    }
    assert!(s.ui_requests.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn source_events_and_report_owned_caches_do_not_invent_shared_refresh() -> R {
    let (mut s, mut b, _) = setup("")?;
    s.execute(
        "insert.pivotTable",
        json!({"source":"Sheet1!A1:C5","destination":"Report!H3","name":"Second","rows":["Region"],"values":["Sales"],"activate":false}),
    )?;
    {
        let doc = s.doc_mut()?;
        let wb = std::sync::Arc::make_mut(&mut doc.wb);
        wb.sheet_mut(0).ok_or("source")?.code_name = Some("SourceCode".into());
        wb.sheet_mut(1).ok_or("report")?.code_name = Some("ReportCode".into());
    }
    let mut ide = Ide::new();
    ide.add_module(
        "Main.bas",
        r##"Sub Main()
Dim first As PivotCache
Dim second As PivotCache
Set first=Worksheets(2).PivotTables(1).PivotCache
Set second=Worksheets(2).PivotTables(2).PivotCache
If first Is second Then Err.Raise 5
If Not (first Is ThisWorkbook.PivotCaches(1)) Then Err.Raise 5
Worksheets(1).Range("C2").Value2=15
first.Refresh
End Sub
"##,
    )?;
    ide.add_document_module(
        "Source.cls",
        "SourceCode",
        r##"Attribute VB_Name = "SourceCode"
Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents=False
Me.Range("M1").Value2=Me.Range("M1").Value2+1
Application.EnableEvents=True
End Sub
"##,
    )?;
    ide.add_document_module(
        "Report.cls",
        "ReportCode",
        r##"Attribute VB_Name = "ReportCode"
Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents=False
ThisWorkbook.Worksheets("Sheet1").Range("N1").Value2=1
Application.EnableEvents=True
End Sub
"##,
    )?;
    b.configure(&mut ide)?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("source")?;
    assert_eq!(sh.value(CellRef::new(0, 12)), CellValue::Number(1.0));
    assert_eq!(sh.value(CellRef::new(0, 13)), CellValue::Empty);
    let report = doc.wb.sheet(1).ok_or("report")?;
    assert_eq!(report.value(CellRef::new(3, 1)), CellValue::Number(45.0));
    assert_eq!(report.value(CellRef::new(3, 8)), CellValue::Number(40.0));
    let bytes = gridcraft_xlsx::write_xlsx(&doc.wb)?;
    let (loaded, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert_eq!(loaded.sheet(1).ok_or("report")?.pivots.len(), 2);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn slicer_lifetimes_and_protection_prevent_retarged_or_partial_widget_edits() -> R {
    let (mut s, mut b, mut ide) = setup(
        r##"Sub Main()
Dim old As Slicer
Set old=ThisWorkbook.SlicerCaches(1).Slicers(1)
Stop
On Error Resume Next
old.Caption="Must not write"
Worksheets(2).Range("M1").Value2=Err.Number
Err.Clear
ThisWorkbook.SlicerCaches(1).Slicers(1).Caption="Protected"
Worksheets(2).Range("M2").Value2=Err.Number
Err.Clear
ThisWorkbook.SlicerCaches(1).SlicerItems("West").Selected=False
Worksheets(2).Range("M3").Value2=Err.Number
End Sub
"##,
    )?;
    s.execute("insert.table", json!({"sheet":0,"range":"A1:C5","name":"SalesTable"}))?;
    s.execute("insert.slicer", json!({"table":"SalesTable","column":"Region"}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break, "{:?}", ide.diagnostics);
    let id = s.doc()?.wb.sheet(0).ok_or("source")?.slicers[0].id;
    s.execute("slicer.delete", json!({"sheet":0,"slicer":id}))?;
    s.execute("insert.slicer", json!({"table":"SalesTable","column":"Region"}))?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("source")?.slicers[0].id, id);
    s.execute("review.protectSheet", json!({}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("source")?;
    assert_eq!(sh.slicers[0].caption, "Region");
    assert!(sh.slicers[0].selected.is_none());
    let out = doc.wb.sheet(1).ok_or("report")?;
    for (row, value) in [(0, 91.0), (1, 1004.0), (2, 1004.0)] {
        assert_eq!(out.value(CellRef::new(row, 12)), CellValue::Number(value));
    }
    Ok(())
}
