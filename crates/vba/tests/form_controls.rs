//! Native Excel Shape/ControlFormat integration, not a second control runtime.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value as CellValue},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn run(source: &str) -> Result<Session, Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    let mut b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Controls.bas", source)?;
    b.configure(&mut ide)?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?} {:?}", ide.diagnostics, ide.output());
    Ok(s)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn shapes_addformcontrol_and_controlformat_use_native_units_types_lists_and_links() -> R {
    let mut s = run(r#"Option Explicit
Sub Main()
Dim cb As Shape
Dim spin As Shape
Dim lb As Shape
Dim c As ControlFormat
Set cb=ActiveSheet.Shapes.AddFormControl(xlCheckBox, 72, 36, 144, 24)
cb.Name="Agree"
cb.TextFrame.Characters.Text="Agree caption"
cb.OnAction="[0]!Controls.Go"
If cb.Type<>msoFormControl Then Err.Raise 5
If cb.FormControlType<>xlCheckBox Then Err.Raise 5
If TypeName(cb.Left)<>"Single" Then Err.Raise 5
If cb.Left<>72 Then Err.Raise 5
cb.IncrementLeft 3
If cb.Left<>75 Then Err.Raise 5
cb.Width=150
cb.Visible=msoFalse
cb.Visible=msoTrue
Set c=cb.ControlFormat
If Not(c.Parent Is cb) Then Err.Raise 5
If Not(c.Application Is Application) Then Err.Raise 5
c.LinkedCell="A1"
c.Value=xlOn
If Range("A1").Value2<>True Then Err.Raise 5
c.PrintObject=False
Set spin=ActiveSheet.Shapes.AddFormControl(xlSpinner, 300, 36, 24, 48)
spin.ControlFormat.Min=2
spin.ControlFormat.Max=60
spin.ControlFormat.Value=14
spin.ControlFormat.SmallChange=2
spin.ControlFormat.LargeChange=8
spin.ControlFormat.LinkedCell="B1"
spin.ControlFormat.Value=18
If Range("B1").Value2<>18 Then Err.Raise 5
Set lb=ActiveSheet.Shapes.AddFormControl(xlListBox, 100, 100, 180, 90)
Set c=lb.ControlFormat
c.List=Array("Alpha","Beta")
c.AddItem "Middle",2
If c.ListCount<>3 Then Err.Raise 5
If c.List(2)<>"Middle" Then Err.Raise 5
c.RemoveItem 2
c.ListIndex=2
If c.Value<>2 Then Err.Raise 5
Dim items As Variant
items=c.List
If LBound(items)<>1 Then Err.Raise 5
If items(2)<>"Beta" Then Err.Raise 5
c.MultiSelect=xlExtended
If c.MultiSelect<>xlExtended Then Err.Raise 5
On Error Resume Next
Dim invalid As Long
invalid=c.ListIndex
If Err.Number<>1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
Range("C1").Value2=c.ListCount
End Sub
Public Sub Go()
Range("D1").Value2=1
End Sub
"#)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.shapes.len(), 3);
    assert_eq!(sh.shapes[0].anchor.width, 200.);
    assert_eq!(sh.shapes[0].anchor.dx + sh.col_left(sh.shapes[0].anchor.cell.col) as f32, 100.);
    assert_eq!(sh.value(CellRef::parse("C1").ok_or("cell")?), CellValue::Number(2.));
    assert_eq!(sh.value(CellRef::parse("D1").ok_or("cell")?), CellValue::Empty);
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (back, report) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert!(report.warnings.is_empty(), "{:?}", report.warnings);
    assert_eq!(back.sheet(0).ok_or("sheet")?.shapes.len(), 3);
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.shapes.is_empty());
    Ok(())
}
