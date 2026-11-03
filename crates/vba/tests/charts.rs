//! Original Excel-shaped chart sources; all effects use native engine chart commands.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value as CellValue},
    model::{AnchorMode, ChartKind, LegendPos},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    s.execute("range.setValues", json!({"range":"A1:B3","values":[["Category","Sales"],["One",10],["Two",20]]}))?;
    s.execute("insert.chart", json!({"range":"A1:B3","at":"D2"}))?;
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Charts.bas", source)?;
    b.configure(&mut ide)?;
    Ok((s, b, ide))
}
fn finish(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    for _ in 0..1000 {
        b.tick(s, ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn typed_chart_graph_edits_inactive_sheet_and_retains_single_undo_group() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Option Explicit
Sub Main()
Dim container As Excel.ChartObject
Dim plot As Excel.Chart
Dim title As ChartTitle
Dim legends As Legend
Set container = ThisWorkbook.Worksheets(1).ChartObjects(1)
Set plot = container.Chart
plot.ChartType = xlLineMarkers
plot.ChartStyle = 6
plot.HasTitle = True
Set title = plot.ChartTitle
title.Text = "Native sales"
plot.HasLegend = True
Set legends = plot.Legend
legends.Position = xlLegendPositionTop
container.Left = 150
container.Top = 75
container.Width = 500
container.Height = 300
container.Placement = xlMove
ThisWorkbook.Worksheets(2).Range("D1").Value2 = container.Parent.Name
ThisWorkbook.Worksheets(2).Range("D2").Value2 = title.Parent.Parent.Name
End Sub
"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":1}))?;
    let before = s.doc()?.undo.len();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    assert_eq!(doc.sheet_index(), 1);
    let chart = doc.wb.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?;
    assert_eq!(chart.kind, ChartKind::LineMarkers);
    assert_eq!(chart.title.as_deref(), Some("Native sales"));
    assert_eq!(chart.legend, LegendPos::Top);
    assert_eq!(chart.anchor.width, 500.0);
    assert_eq!(chart.anchor.mode, AnchorMode::MoveOnly);
    assert_eq!(doc.wb.sheet(0).ok_or("sheet")?.col_left(chart.anchor.cell.col) + f64::from(chart.anchor.dx), 150.0);
    assert_eq!(doc.wb.sheet(1).ok_or("sheet")?.value(CellRef::new(0, 3)), CellValue::text("Sheet1"));
    assert_eq!(doc.undo.len(), before + 1);
    s.execute("edit.undo", json!({}))?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?.kind, ChartKind::ColumnClustered);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn chart_source_and_collection_defaults_use_live_native_state() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim item As ChartObject
Dim total As Long
For Each item In Worksheets(1).ChartObjects
total = total + 1
item.Chart.SetSourceData Source:=Worksheets(2).Range("A1:C3"), PlotBy:=xlColumns
Next item
Worksheets(1).Range("F1").Value2 = total
Worksheets(1).Range("F2").Value2 = Worksheets(1).ChartObjects("Chart 1").Chart.PlotBy
End Sub
"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Data"}))?;
    s.execute("range.setValues", json!({"sheet":1,"range":"A1:C3","values":[["x","A","B"],["one",1,2],["two",3,4]]}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    // An explicit source sheet must not change the native command's default chart owner.
    s.execute("chart.selectData", json!({"chart":1,"range":"Data!A1:C3","byRows":false}))?;

    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let chart = s.doc()?.wb.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?;
    assert_eq!(chart.series.len(), 2);
    assert!(!chart.by_rows);
    assert!(chart.source.as_ref().is_some_and(|s| s.contains("Data!")));
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 5)), CellValue::Number(1.0));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_chart_id_reuse_does_not_retarget_parked_handles() -> R {
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nDim old As ChartObject\nSet old=Worksheets(1).ChartObjects(1)\nStop\nOn Error Resume Next\nold.Width=900\nRange(\"F1\").Value2=Err.Number\nEnd Sub\n",
    )?;
    let id = s.doc()?.wb.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?.id;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break);
    s.execute("chart.delete", json!({"chart":id}))?;
    s.execute("insert.chart", json!({"range":"A1:B3"}))?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?.id, id);
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 5)), CellValue::Number(91.0));
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?.anchor.width, 480.0);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn invalid_chart_options_and_protection_do_not_partially_write() -> R {
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nDim p As Chart\nSet p=Worksheets(1).ChartObjects(1).Chart\nOn Error Resume Next\np.ChartType=xl3DColumn\nRange(\"F1\").Value2=Err.Number\nErr.Clear\np.ChartStyle=999\nRange(\"F2\").Value2=Err.Number\nErr.Clear\nWorksheets(1).ChartObjects.Add 1,1,100,100\nRange(\"F3\").Value2=Err.Number\nEnd Sub\n",
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.charts.first().ok_or("chart")?.kind, ChartKind::ColumnClustered);
    for row in 0..3 {
        assert_eq!(sh.value(CellRef::new(row, 5)), CellValue::Number(1004.0));
    }
    s.execute("review.protectSheet", json!({}))?;
    let mut protected = Ide::new();
    protected.add_module("Protected.bas", "Sub Main()\nWorksheets(1).ChartObjects(1).Chart.HasTitle=False\nEnd Sub\n")?;
    b.configure(&mut protected)?;
    b.dispatch(&mut s, &mut protected, Action::Run)?;
    for _ in 0..100 {
        b.tick(&mut s, &mut protected)?;
        if !matches!(protected.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(protected.mode(), Mode::Failed);
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?.title.is_some());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn chart_edits_persist_and_renderer_reads_the_same_native_model() -> R {
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nDim c As Chart\nSet c=Worksheets(1).ChartObjects(1).Chart\nc.HasTitle=True\nc.ChartTitle.Caption=\"Persisted\"\nc.ChartType=xlLine\nc.HasLegend=False\nWorksheets(1).ChartObjects(1).Placement=xlFreeFloating\nEnd Sub\n",
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (loaded, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let native = s.doc()?.wb.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?;
    let restored = loaded.sheet(0).ok_or("sheet")?.charts.first().ok_or("chart")?;
    assert_eq!(restored.kind, native.kind);
    assert_eq!(restored.title, native.title);
    assert_eq!(restored.legend, LegendPos::None);
    assert_eq!(restored.anchor.mode, AnchorMode::Absolute);
    assert!(!native.identity.same(&restored.identity));
    let data = gridcraft_chart::resolve(&loaded, 0, restored);
    assert_eq!(data.series.first().ok_or("series")?.values, vec![Some(10.0), Some(20.0)]);
    let prims = gridcraft_chart::render(restored, &data, 500.0, 300.0, &gridcraft_chart::ApproxMeasure);
    assert!(prims.iter().any(|p| matches!(p,gridcraft_chart::Prim::Text {text,..} if text == "Persisted")));

    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn deleting_title_and_chart_updates_live_counts_and_invalidates_children() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim holder As ChartObject
Dim oldTitle As ChartTitle
Set holder=Worksheets(1).ChartObjects(1)
Set oldTitle=holder.Chart.ChartTitle
Dim result As Variant
result = oldTitle.Delete()
On Error Resume Next
oldTitle.Text="must not write"
Range("F1").Value2=Err.Number
Err.Clear
holder.Delete
Range("F2").Value2=Worksheets(1).ChartObjects.Count
holder.Width=200
Range("F3").Value2=Err.Number
End Sub
"#,
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert!(sh.charts.is_empty());
    for (row, n) in [(0, 91.0), (1, 0.0), (2, 91.0)] {
        assert_eq!(sh.value(CellRef::new(row, 5)), CellValue::Number(n));
    }
    Ok(())
}
