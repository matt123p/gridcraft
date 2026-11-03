//! OM12 native-backed series, layout and drawing contracts.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value as CellValue},
    model::{AnchorMode, ShapeKind},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn png(w: u32, h: u32) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    use image::ImageEncoder;
    let pixels = gridcraft_chart::rasterize(&[], w, h, [30, 60, 90, 255]);
    let mut out = Vec::new();
    image::codecs::png::PngEncoder::new(&mut out).write_image(&pixels, w, h, image::ExtendedColorType::Rgba8)?;
    Ok(out)
}
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
fn typed_series_arrays_trendlines_and_layout_share_native_state() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Option Explicit
Sub Main()
Dim c As Chart
Dim se As Series
Dim t As Trendline
Dim ax As Axis
Dim area As PlotArea
Set c = Worksheets(1).ChartObjects(1).Chart
Set se = c.SeriesCollection(1)
se.Name = "Revenue"
se.Values = Array(4, 8, 12)
se.XValues = Array("A", "B", "C")
se.Smooth = True
se.Format.Fill.ForeColor.RGB = RGB(20, 60, 120)
Set t = se.Trendlines.Add(Type:=xlLinear, Forward:=2, DisplayEquation:=True, Name:="Forecast")
t.Backward = 1
t.Intercept = 0
t.Format.Line.ForeColor.RGB = RGB(90, 30, 10)
t.Format.Line.Weight = 2
t.Format.Line.DashStyle = msoLineDash
Set ax = c.Axes(xlValue)
ax.HasTitle = True
ax.AxisTitle.Text = "Amount"
If Not (ax.Parent Is c) Then Err.Raise 5
If Not (se.Format.Line.Parent.Parent Is se) Then Err.Raise 5
ax.HasMajorGridlines = False
Set area = c.PlotArea
Dim data As Variant
data = se.Values
Worksheets(1).Range("G1").Value2 = data(1)
Worksheets(1).Range("H1").Value2 = t.Name
Worksheets(1).Range("H2").Value2 = ax.MaximumScale
Worksheets(1).Range("H3").Value2 = area.Width
c.ApplyDataLabels
Worksheets(1).Range("H4").Value2 = se.DataLabels.Count
Worksheets(1).Range("H5").Value2 = se.DataLabels(2).Text
End Sub
"#,
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("sheet")?;
    let c = sh.charts.first().ok_or("chart")?;
    let se = c.series.first().ok_or("series")?;
    assert_eq!(se.values, "{4,8,12}");
    assert!(se.smooth);
    assert_eq!(c.y_title.as_deref(), Some("Amount"));
    assert!(!c.gridlines);
    assert!(c.data_labels);
    assert_eq!(se.trendlines.len(), 1);
    assert_eq!(se.trendlines[0].intercept, Some(0.0));
    assert_eq!(se.trendlines[0].width, Some(2.0));
    assert_eq!(se.trendlines[0].dash, Some(gridcraft_engine::model::LineDash::Dash));
    assert_eq!(se.trendlines[0].backward, 1.0);
    let data = gridcraft_chart::resolve(&doc.wb, 0, c);
    let (prims, layout) = gridcraft_chart::render_layout(c, &data, c.anchor.width, c.anchor.height, &gridcraft_chart::ApproxMeasure);
    assert!(!prims.is_empty());
    assert_eq!(sh.value(CellRef::new(1, 7)), CellValue::Number(layout.value.ok_or("scale")?.max));
    assert_eq!(sh.value(CellRef::new(2, 7)), CellValue::Number(f64::from(layout.plot.ok_or("plot")?[2])));
    assert_eq!(sh.value(CellRef::new(4, 7)), CellValue::Number(8.0));
    let bytes = gridcraft_xlsx::write_xlsx(&doc.wb)?;
    let (loaded, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let re = loaded.sheet(0).ok_or("loaded sheet")?.charts.first().ok_or("loaded chart")?;
    assert_eq!(re.series[0].values, se.values);
    let restored_data = gridcraft_chart::resolve(&loaded, 0, re);
    let restored = gridcraft_chart::render(re, &restored_data, re.anchor.width, re.anchor.height, &gridcraft_chart::ApproxMeasure);
    assert_eq!(gridcraft_chart::rasterize(&prims, 480, 288, [255; 4]), gridcraft_chart::rasterize(&restored, 480, 288, [255; 4]));
    assert_eq!(re.series[0].trendlines[0].name.as_deref(), Some("Forecast"));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn shapes_text_range_format_zorder_and_inactive_sheet_undo() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Option Explicit
Sub Main()
Dim text As Shape
Dim box As Shape
Dim group As ShapeRange
Dim chars As Characters
Dim frame As Excel.TextFrame2
Dim shared As Office.TextFrame2
Dim fill As Excel.FillFormat
Set text = Worksheets(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 150, 80, 180, 90)
Set chars = text.TextFrame.Characters
chars.Text = "Hello 😀"
chars.Font.Bold = True
chars.Font.Size = 18
If Not (text.TextFrame.Characters(1, 2).Parent Is text.TextFrame) Then Err.Raise 5
text.TextFrame.Characters(1, 5).Text = "World"
text.OnAction = "Main"
text.Fill.ForeColor.RGB = RGB(10, 30, 60)
text.Line.Visible = msoFalse
Set frame = text.TextFrame2
Set shared = frame
Set fill = shared.TextRange.Font.Fill
If Not (fill.Parent Is shared.TextRange.Font) Then Err.Raise 5
frame.TextRange.Font.Fill.ForeColor.RGB = RGB(80, 100, 140)
frame.TextRange.Font.Line.ForeColor.RGB = RGB(0, 0, 0)
frame.TextRange.Font.Line.Weight = 0.5
frame.VerticalAnchor = msoAnchorMiddle
Worksheets(2).Range("H1").Value2 = frame.TextRange.Text
Set box = Worksheets(1).Shapes.AddShape(msoShapeOval, 300, 150, 80, 60)
Set group = Worksheets(1).Shapes.Range(Array(text.Name, box.Name))
group.Align msoAlignLefts, msoFalse
text.ZOrder msoBringToFront
text.Placement = xlMove
Worksheets(2).Range("H2").Value2 = group.Count
Worksheets(2).Range("H3").Value2 = text.ZOrderPosition
Worksheets(2).Range("H4").Value2 = text.TopLeftCell.Address
End Sub
"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":1}))?;
    let undo = s.doc()?.undo.len();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    assert_eq!(doc.sheet_index(), 1);
    assert_eq!(doc.undo.len(), undo + 1);
    let sh = doc.wb.sheet(0).ok_or("sheet")?;
    let text = sh.shapes.iter().find(|x| x.kind == ShapeKind::TextBox).ok_or("textbox")?;
    assert_eq!(text.text, "World 😀");
    assert!(text.format.bold);
    assert_eq!(text.format.anchor, Some(gridcraft_engine::model::ShapeTextAnchor::Middle));
    assert_eq!(text.format.text_outline, Some((gridcraft_engine::model::Color::rgb(0, 0, 0), 0.5)));
    assert!(text.format.no_line);
    assert_eq!(text.format.font_size, Some(18.0));
    assert_eq!(text.anchor.mode, AnchorMode::MoveOnly);
    assert_eq!(text.on_action.as_deref(), Some("Main"));
    let other = doc.wb.sheet(1).ok_or("other")?;
    assert_eq!(other.value(CellRef::new(0, 7)), CellValue::text("World 😀"));
    assert_eq!(other.value(CellRef::new(1, 7)), CellValue::Number(2.0));
    assert_eq!(other.value(CellRef::new(2, 7)), CellValue::Number(2.0));
    let bytes = gridcraft_xlsx::write_xlsx(&doc.wb)?;
    let (loaded, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let text2 = loaded.sheet(0).ok_or("sheet")?.shapes.iter().find(|x| x.kind == ShapeKind::TextBox).ok_or("textbox")?;
    assert_eq!(text2.text, text.text);
    assert!(text2.format.bold);
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.shapes.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn series_add_delete_and_rebind_never_retarget_parked_children() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim c As Chart
Dim first As Series
Dim added As Series
Dim fitted As Trendline
Set c=Worksheets(1).ChartObjects(1).Chart
Set first=c.SeriesCollection(1)
Set fitted=first.Trendlines.Add()
Set added=c.SeriesCollection.Add(Worksheets(2).Range("A1:B3"),xlColumns)
added.Name="Second"
added.PlotOrder=1
Worksheets(1).Range("G2").Value2=first.PlotOrder
first.Delete
Worksheets(1).Range("F1").Value2=added.Name
Worksheets(1).Range("F2").Value2=added.PlotOrder
On Error Resume Next
fitted.Forward2=9
Worksheets(1).Range("F3").Value2=Err.Number
Err.Clear
c.SetSourceData Worksheets(1).Range("A1:B3")
added.Name="Must not write"
Worksheets(1).Range("F4").Value2=Err.Number
End Sub
"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Data"}))?;
    s.execute("range.setValues", json!({"sheet":1,"range":"A1:B3","values":[["Category","Other"],["X",9],["Y",12]]}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.value(CellRef::new(0, 5)), CellValue::text("Second"));
    for (row, n) in [(1, 1.0), (2, 91.0), (3, 91.0)] {
        assert_eq!(sh.value(CellRef::new(row, 5)), CellValue::Number(n));
    }
    assert_eq!(sh.value(CellRef::new(1, 6)), CellValue::Number(2.0));
    assert_eq!(sh.charts[0].series.len(), 1);
    assert_ne!(sh.charts[0].series[0].name.as_deref(), Some("Must not write"));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn drawing_ranges_follow_sheet_moves_and_refuse_stale_or_unsupported_bulk_writes() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim a As Shape
Dim b As Shape
Dim r As ShapeRange
Set a=Worksheets(1).Shapes.AddShape(msoShapeRectangle,10,20,40,30)
Set b=Worksheets(1).Shapes.AddShape(msoShapeOval,100,80,50,40)
Set r=Worksheets(1).Shapes.Range(Array(a.Name,b.Name))
Worksheets(1).Move After:=Worksheets(2)
r.Left=20
Worksheets(1).Range("F1").Value2=a.Left
Worksheets(1).Range("F2").Value2=b.Left
On Error Resume Next
r.Group
Worksheets(1).Range("F3").Value2=Err.Number
Err.Clear
r.Width=100
Worksheets(1).Range("F4").Value2=Err.Number
Err.Clear
a.Delete
r.Top=200
Worksheets(1).Range("F5").Value2=Err.Number
End Sub
"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let result = doc.wb.sheet(0).ok_or("result")?;
    for (row, n) in [(0, 20.0), (1, 110.0), (2, 1004.0), (3, 1004.0), (4, 91.0)] {
        assert_eq!(result.value(CellRef::new(row, 5)), CellValue::Number(n));
    }
    let shapes = &doc.wb.sheet(1).ok_or("drawings")?.shapes;
    assert_eq!(shapes.len(), 1);
    assert_eq!(shapes[0].anchor.height, 40.0);
    assert_ne!(shapes[0].anchor.dy, 200.0);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn drawing_protection_and_reused_ids_refuse_before_effects() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim old As Shape
Set old=Worksheets(1).Shapes.AddTextbox(1,10,20,100,40)
Stop
On Error Resume Next
old.Width=900
Worksheets(2).Range("F1").Value2=Err.Number
Err.Clear
Worksheets(1).Shapes(2).Left=100
Worksheets(2).Range("F2").Value2=Err.Number
Err.Clear
Worksheets(1).Shapes.AddShape msoShapeRectangle,0,0,40,40
Worksheets(2).Range("F3").Value2=Err.Number
End Sub
"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break);
    let id = s.doc()?.wb.sheet(0).ok_or("sheet")?.shapes[0].id;
    s.execute("object.delete", json!({"kind":"shape","id":id}))?;
    s.execute("insert.shape", json!({"kind":"rectangle","at":"D2"}))?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.shapes[0].id, id);
    s.execute("review.protectSheet", json!({}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let out = doc.wb.sheet(1).ok_or("sheet")?;
    assert_eq!(out.value(CellRef::new(0, 5)), CellValue::Number(91.0));
    // Shapes(2) is the chart, since native paint order puts ordinary shapes first.
    for row in 1..3 {
        assert_eq!(out.value(CellRef::new(row, 5)), CellValue::Number(1004.0));
    }
    assert_eq!(doc.wb.sheet(0).ok_or("sheet")?.shapes.len(), 1);
    assert_eq!(doc.wb.sheet(0).ok_or("sheet")?.shapes[0].anchor.width, 144.0);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn picture_reads_and_failed_services_preserve_native_drawing_state() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim picture As Shape
Set picture=Worksheets(1).Shapes("Picture 2")
picture.Left=125
picture.Top=50
picture.Width=80
picture.Placement=xlFreeFloating
Worksheets(1).Range("F1").Value2=picture.Type
Worksheets(1).Range("F2").Value2=picture.AlternativeText
On Error Resume Next
Worksheets(1).Shapes.AddPicture "/missing/om12.png",False,True,0,0,20,20
Range("F3").Value2=Err.Number
Err.Clear
Worksheets(1).Shapes.AddPicture "unused.png",True,False,0,0,20,20
Range("F4").Value2=Err.Number
End Sub
"#,
    )?;
    // Original one-pixel RGBA image, synthesized by the native rasterizer.
    let png = png(1, 1)?;
    s.execute("insert.picture", json!({"base64":gridcraft_engine::io::base64_encode(&png),"alt":"Original pixel","width":10,"height":10}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.images.len(), 1);
    let picture = &sh.images[0];
    assert_eq!(picture.anchor.width, 80.0);
    assert_eq!(picture.anchor.mode, AnchorMode::Absolute);
    for (row, value) in
        [(0, CellValue::Number(13.0)), (1, CellValue::text("Original pixel")), (2, CellValue::Number(1004.0)), (3, CellValue::Number(1004.0))]
    {
        assert_eq!(sh.value(CellRef::new(row, 5)), value);
    }
    assert!(s.ui_requests.is_empty());
    let bytes = gridcraft_xlsx::write_xlsx(&doc.wb)?;
    let (loaded, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert_eq!(loaded.sheet(0).ok_or("sheet")?.images[0].data, picture.data);
    Ok(())
}
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn add_picture_embeds_native_file_with_original_dimensions_on_inactive_owner() -> R {
    let path = std::env::temp_dir().join(format!("om12-picture-{}.png", std::process::id()));
    let png = png(3, 2)?;
    std::fs::write(&path, &png)?;
    let source = format!(
        "Sub Main()\nDim p As Shape\nSet p=Worksheets(1).Shapes.AddPicture(\"{}\",False,True,17,29,-1,-1)\nEnd Sub\n",
        path.to_string_lossy()
    );
    let result = (|| -> R {
        let (mut s, mut b, mut ide) = setup(&source)?;
        s.execute("home.insertSheet", json!({"name":"Other"}))?;
        s.execute("sheet.activate", json!({"sheet":1}))?;
        b.dispatch(&mut s, &mut ide, Action::Run)?;
        finish(&mut s, &mut b, &mut ide)?;
        let doc = s.doc()?;
        assert_eq!(doc.sheet_index(), 1);
        let sh = doc.wb.sheet(0).ok_or("sheet")?;
        let p = sh.images.first().ok_or("picture")?;
        assert_eq!((p.anchor.width, p.anchor.height), (3.0, 2.0));
        assert_eq!(p.data.as_ref(), &png);
        assert_eq!(sh.col_left(p.anchor.cell.col) + f64::from(p.anchor.dx), 17.0);
        assert_eq!(sh.row_top(p.anchor.cell.row) + f64::from(p.anchor.dy), 29.0);
        Ok(())
    })();
    std::fs::remove_file(path)?;
    result
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn names_native_anchor_transforms_and_source_deletion_stay_live() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim se As Series
Dim p As Shape
Set se=Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1)
se.Name="Hello!"
Worksheets(1).Range("J1").Value2=se.Name
se.Name="=Sheet1!$B$1"
Worksheets(1).Range("J2").Value2=se.Name
Set p=Worksheets(1).Shapes.AddShape(msoShapeRectangle,220,100,70,40)
p.Placement=xlMove
Stop
Worksheets(1).Range("F3").Value2=p.Top
Worksheets(1).Range("F4").Value2=se.Name
On Error Resume Next
Worksheets(1).Shapes.AddLine 100,100,20,20
Worksheets(1).Range("F5").Value2=Err.Number
End Sub
"#,
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break);
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    let anchor = sh.shapes[0].anchor;
    s.execute("home.insertRows", json!({"rows":"1:1"}))?;
    s.execute("edit.clearContents", json!({"range":"A3:B4"}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("sheet")?;
    // Existing results followed the inserted row; the parked native anchor and series still resolve.
    assert_eq!(sh.value(CellRef::new(1, 9)), CellValue::text("Hello!"));
    assert_eq!(sh.value(CellRef::new(2, 9)), CellValue::text("Sales"));
    let moved = sh.shapes[0].anchor;
    assert_eq!(moved.cell.row, anchor.cell.row + 1);
    assert_eq!(moved.height, anchor.height);
    assert_eq!(sh.value(CellRef::new(2, 5)), CellValue::Number(sh.row_top(moved.cell.row) + f64::from(moved.dy)));
    assert_eq!(sh.value(CellRef::new(3, 5)), CellValue::text("Sales"));
    assert_eq!(sh.value(CellRef::new(4, 5)), CellValue::Number(1004.0));
    let data = gridcraft_chart::resolve(&doc.wb, 0, &sh.charts[0]);
    assert!(data.series[0].values.iter().all(Option::is_none));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn embedded_chart_location_retargets_parked_descendants_without_switching_view() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim c As Chart
Dim se As Series
Dim holder As ChartObject
Dim drawing As ShapeRange
Set holder=Worksheets(1).ChartObjects(1)
Set c=holder.Chart
Set se=c.SeriesCollection(1)
Set drawing=Worksheets(1).ChartObjects.ShapeRange
Set c=c.Location(xlLocationAsObject,"Other")
se.Name="Moved"
drawing.Left=190
Worksheets(1).Range("F1").Value2=holder.Parent.Name
Worksheets(1).Range("F2").Value2=c.Parent.Parent.Name
Worksheets(1).Range("F3").Value2=Worksheets(1).ChartObjects.Count
Worksheets(1).Range("F4").Value2=Worksheets(2).ChartObjects.Count
On Error Resume Next
c.Location xlLocationAsNewSheet,"Unsupported"
Worksheets(1).Range("F5").Value2=Err.Number
End Sub
"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    let undo = s.doc()?.undo.len();
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    assert_eq!(doc.sheet_index(), 0);
    assert_eq!(doc.undo.len(), undo + 1);
    let sh = doc.wb.sheet(0).ok_or("sheet")?;
    assert!(sh.charts.is_empty());
    for row in 0..2 {
        assert_eq!(sh.value(CellRef::new(row, 5)), CellValue::text("Other"));
    }
    for (row, n) in [(2, 0.0), (3, 1.0), (4, 1004.0)] {
        assert_eq!(sh.value(CellRef::new(row, 5)), CellValue::Number(n));
    }
    let other = doc.wb.sheet(1).ok_or("other")?;
    let c = &other.charts[0];
    assert_eq!(c.series[0].name.as_deref(), Some("Moved"));
    assert_eq!(other.col_left(c.anchor.cell.col) + f64::from(c.anchor.dx), 190.0);
    s.execute("edit.undo", json!({}))?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.charts.len(), 1);
    assert!(s.doc()?.wb.sheet(1).ok_or("other")?.charts.is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn external_chart_moves_preserve_existing_graph_lifetimes() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim c As Chart
Dim se As Series
Dim r As ShapeRange
Set c=Worksheets(1).ChartObjects(1).Chart
Set se=c.SeriesCollection(1)
Set r=c.Parent.ShapeRange
Stop
Worksheets(1).Range("F1").Value2=c.Parent.Parent.Name
Worksheets(1).Range("F2").Value2=se.Parent.Parent.Parent.Name
Worksheets(1).Range("F3").Value2=r.Parent.Name
Stop
Worksheets(1).Range("F4").Value2=c.Parent.Parent.Name
Worksheets(1).Range("F5").Value2=se.Parent.Parent.Parent.Name
Worksheets(1).Range("F6").Value2=r.Parent.Name
End Sub
"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    let id = s.doc()?.wb.sheet(0).ok_or("sheet")?.charts[0].id;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break);
    s.execute("chart.move", json!({"chart":id,"sheet":"Other"}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Break {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Break);
    // Native chart moves keep the same lifetime, including transfer back through another command.
    s.execute("chart.move", json!({"chart":id,"sourceSheet":1,"sheet":"Sheet1"}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let doc = s.doc()?;
    let sh = doc.wb.sheet(0).ok_or("sheet")?;
    for row in 0..3 {
        assert_eq!(sh.value(CellRef::new(row, 5)), CellValue::text("Other"));
    }
    for row in 3..6 {
        assert_eq!(sh.value(CellRef::new(row, 5)), CellValue::text("Sheet1"));
    }
    Ok(())
}
