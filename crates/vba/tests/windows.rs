//! OM08 native window ownership, lifecycle and saved-view acceptance.
use gridcraft_engine::{Session, core::CellRef};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Windows.bas", source)?;
    b.configure(&mut ide)?;
    Ok((s, b, ide))
}
fn finish(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    b.dispatch(s, ide, Action::Run)?;
    for _ in 0..10000 {
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
fn window_graph_keeps_parked_selection_and_explicit_activation() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Option Explicit
Sub Main()
Dim first As Window, second As Window, entry As Window
Set first=ActiveWindow
Range("B2").Select
Set second=first.NewWindow
Range("D4:E5").Select
If Not (Application.ActiveWindow Is second) Then Err.Raise 5
If first.ActiveCell.Address <> "$B$2" Then Err.Raise 5
If Not (Application.ActiveWindow Is second) Then Err.Raise 5
If second.Selection.Address <> "$D$4:$E$5" Then Err.Raise 5
If Not (second.RangeSelection Is second.Selection) Then Err.Raise 5
If Not (second.ActiveSheet Is ActiveSheet) Then Err.Raise 5
If Not (first.Parent Is ThisWorkbook) Then Err.Raise 5
If Not (first.Application Is Application) Then Err.Raise 5
If Not (ThisWorkbook.Windows.Parent Is ThisWorkbook) Then Err.Raise 5
If Not (Windows.Parent Is Application) Then Err.Raise 5
If Windows.Count <> 2 Then Err.Raise 5
If second.Type <> xlWorkbook Then Err.Raise 5
If second.WindowNumber <> 2 Then Err.Raise 5
If second.SelectedSheets.Count <> 1 Then Err.Raise 5
If Not (second.SelectedSheets(1) Is second.ActiveSheet) Then Err.Raise 5
If Not (Windows(second.Caption) Is second) Then Err.Raise 5
For Each entry In ThisWorkbook.Windows
If Not (entry.Parent Is ThisWorkbook) Then Err.Raise 5
Next
first.Activate
If Not (ActiveWindow Is first) Then Err.Raise 5
If ActiveCell.Address <> "$B$2" Then Err.Raise 5
first.ActivateNext
If Not (ActiveWindow Is second) Then Err.Raise 5
second.ActivatePrevious
If Not (ActiveWindow Is first) Then Err.Raise 5
second.Activate
If ActiveCell.Address <> "$D$4" Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.windows().len(), 2);
    assert_eq!(s.doc()?.selection.active, CellRef::parse("D4").ok_or("cell")?);
    assert!(!s.doc()?.is_dirty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn window_close_hides_and_rejects_stale_handles_without_closing_workbook() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim first As Window, second As Window, third As Window, x As String
Set first=ActiveWindow
Set second=ThisWorkbook.NewWindow
second.Visible=False
If second.Visible Then Err.Raise 5
second.Visible=True
If Not second.Visible Then Err.Raise 5
If Not second.Close(True, "ignored.xlsx") Then Err.Raise 5
If ThisWorkbook.Windows.Count <> 1 Then Err.Raise 5
Set third=first.NewWindow
On Error Resume Next
x=second.Caption
If Err.Number <> 91 Then Err.Raise 5
Err.Clear
On Error GoTo 0
third.Close
If Not (ActiveWindow Is first) Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.documents().len(), 1);
    assert_eq!(s.windows().len(), 1);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn global_windows_preserve_cross_workbook_ownership_and_focus() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim book As Workbook, old As Window, other As Window, copy As Window
Set old=ActiveWindow
Set book=Workbooks.Add
Set other=ActiveWindow
Range("F6").Select
If Windows.Count <> 2 Then Err.Raise 5
If ThisWorkbook.Windows.Count <> 1 Then Err.Raise 5
If Not (other.Parent Is book) Then Err.Raise 5
If Not (Windows(1) Is old) Then Err.Raise 5
If old.ActiveCell.Address <> "$A$1" Then Err.Raise 5
If Not (ActiveWindow Is other) Then Err.Raise 5
Set copy=ThisWorkbook.NewWindow
If Not (ActiveWindow Is copy) Then Err.Raise 5
If Not (copy.Parent Is ThisWorkbook) Then Err.Raise 5
other.Activate
If Not (ActiveWorkbook Is book) Then Err.Raise 5
If ActiveCell.Address <> "$F$6" Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.documents().len(), 2);
    assert_eq!(s.windows().len(), 3);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn headless_window_gaps_refuse_before_selection_or_sheet_view_mutation() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim w As Window, x As Variant
Set w=ThisWorkbook.NewWindow
Range("C3").Select
On Error Resume Next
w.Zoom=150
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
w.FreezePanes=True
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
w.DisplayGridlines=False
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
x=w.Width
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
x=w.VisibleRange
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
Application.Goto Range("H8"), True
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If ActiveCell.Address <> "$C$3" Then Err.Raise 5
Application.Goto Range("E5"), False
If ActiveCell.Address <> "$E$5" Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.zoom, 100);
    assert_eq!(sh.freeze, None);
    assert!(sh.show_gridlines);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_saved_view_options_snapshots_and_xlsx_persistence() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim view As CustomView, entry As CustomView
Range("G7").Select
Set view=ThisWorkbook.CustomViews.Add("Detail", False, False)
If Not (view Is ThisWorkbook.CustomViews("detail")) Then Err.Raise 5
If Not (view.Parent Is ThisWorkbook) Then Err.Raise 5
If view.PrintSettings Or view.RowColSettings Then Err.Raise 5
Range("B2").Select
view.Show
If ActiveCell.Address <> "$G$7" Then Err.Raise 5
For Each entry In ThisWorkbook.CustomViews
If entry.Name <> "Detail" Then Err.Raise 5
Next
On Error Resume Next
ThisWorkbook.CustomViews.Add "Unsupported", True, False
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
ThisWorkbook.CustomViews.Add "UnsupportedDefaults"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If ThisWorkbook.CustomViews.Count <> 1 Then Err.Raise 5
End Sub"#,
    )?;
    s.execute("view.zoom", json!({"percent":125}))?;
    s.execute("view.freezeTopRow", json!({}))?;
    finish(&mut s, &mut b, &mut ide)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.zoom, 125);
    assert_eq!(sh.freeze, Some((1, 0)));
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert!(wb.names.iter().any(|n| n.name == "_sc_custom_views"));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn saved_view_handles_invalidate_on_native_collection_replacement() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim old As CustomView, newer As CustomView, x As String
Set old=ThisWorkbook.CustomViews.Add("One", False, False)
old.Delete
Set newer=ThisWorkbook.CustomViews.Add("One", False, False)
On Error Resume Next
x=old.Name
If Err.Number <> 91 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If newer.Name <> "One" Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn parked_window_follows_native_sheet_identity_after_rename_and_reorder() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim first As Window, second As Window
Worksheets(2).Activate
Range("C3").Select
Worksheets(1).Activate
Range("B2").Select
Set first=ActiveWindow
Set second=ThisWorkbook.NewWindow
Worksheets(1).Name="Renamed"
Worksheets(1).Move After:=Worksheets(2)
If first.ActiveSheet.Name <> "Renamed" Then Err.Raise 5
If Not (ActiveWindow Is second) Then Err.Raise 5
If first.SelectedSheets(1).Name <> "Renamed" Then Err.Raise 5
first.Activate
Worksheets("Other").Activate
If ActiveCell.Address <> "$C$3" Then Err.Raise 5
End Sub"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn last_window_close_reuses_workbook_lifecycle_and_boolean_result() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim book As Workbook, w As Window
Set book=Workbooks.Add
Set w=ActiveWindow
Range("B1").Value2=5
If Not w.Close(False) Then Err.Raise 5
If Workbooks.Count <> 1 Then Err.Raise 5
If Not (ActiveWorkbook Is ThisWorkbook) Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.documents().len(), 1);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn last_window_beforeclose_cancellation_returns_false() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
If ActiveWindow.Close(False) Then Err.Raise 5
If Windows.Count <> 1 Then Err.Raise 5
Range("B2").Value2=7
End Sub"#,
    )?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).code_name = Some("BookCode".into());
    ide.add_document_module(
        "BookCode.cls",
        "BookCode",
        "Attribute VB_Name = \"BookCode\"\nPrivate Sub Workbook_BeforeClose(ByRef Cancel As Boolean)\nCancel=True\nEnd Sub",
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.documents().len(), 1);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn watches_aggregate_native_workbooks_and_reject_replaced_handles() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim one As Watch, two As Watch, w As Watch, book As Workbook, x As Variant
Set one=Application.Watches.Add(Range("B2"))
If one.Source.Address <> "$B$2" Then Err.Raise 5
If Not (one.Parent Is Application) Then Err.Raise 5
If Not (one.Application Is Application) Then Err.Raise 5
Set book=Workbooks.Add
Set two=Application.Watches.Add(Range("D4"))
If Application.Watches.Count <> 2 Then Err.Raise 5
If Not (Application.Watches(2) Is two) Then Err.Raise 5
For Each w In Application.Watches
If Not (w.Parent Is Application) Then Err.Raise 5
Next
one.Delete
On Error Resume Next
Set x=one.Source
If Err.Number <> 91 Then Err.Raise 5
Err.Clear
Application.Watches.Add Range("D4:E5")
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If Application.Watches.Count <> 1 Then Err.Raise 5
Application.Watches.Delete
If Application.Watches.Count <> 0 Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn malformed_saved_view_fails_before_focus_and_selection_change() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim v As CustomView
Set v=ThisWorkbook.CustomViews(1)
On Error Resume Next
v.Show
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If ActiveSheet.Name <> "Sheet1" Then Err.Raise 5
If ActiveCell.Address <> "$A$1" Then Err.Raise 5
End Sub"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("view.customViews", json!({"save":"Broken"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    let wb = std::sync::Arc::make_mut(&mut s.doc_mut()?.wb);
    let n = wb.names.iter_mut().find(|n| n.name == "_sc_custom_views").ok_or("view")?;
    let value = json!({"Broken":{"sheet":1,"selection":"A1","zoom":999,"freeze":null}}).to_string();
    n.formula = format!("\"{}\"", value.replace('"', "\"\""));
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn window_targets_enforce_workbook_consent_before_any_effect() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim w As Window, x As String
Set w=Windows(2)
On Error Resume Next
x=w.Caption
If Err.Number <> 70 Then Err.Raise 5
Err.Clear
w.Close False
If Err.Number <> 70 Then Err.Raise 5
Err.Clear
ActiveWindow.ActivateNext
If Err.Number <> 70 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If Windows.Count <> 2 Then Err.Raise 5
If Not (ActiveWorkbook Is ThisWorkbook) Then Err.Raise 5
End Sub"#,
    )?;
    let first = s.active_window().ok_or("window")?;
    s.new_workbook();
    s.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
    s.activate_window(first);
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn last_window_close_waits_for_the_existing_save_changes_prompt() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim book As Workbook
Set book=Workbooks.Add
Range("A1").Value2=8
If Not ActiveWindow.Close Then Err.Raise 5
If Workbooks.Count <> 1 Then Err.Raise 5
End Sub"#,
    )?;
    b.set_prompt_services(gridcraft_vba::PromptServices { close: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut prompts = 0;
    for _ in 0..10000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(request) = b.take_prompt_request() {
            assert!(matches!(request.prompt, gridcraft_vba::LifecyclePrompt::SaveChanges { .. }));
            prompts += 1;
            b.respond_to_prompt(request.token, gridcraft_vba::PromptResponse::Discard).map_err(|e| e.description)?;
        }
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(prompts, 1);
    assert_eq!(s.documents().len(), 1);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn inactive_window_reads_preserve_native_most_recent_live_view() -> R {
    let (mut s, _, mut ide) = setup(
        r#"Sub Main()
Dim parked As Window, live As Window, book As Workbook
Set parked=Windows(1)
Set live=Windows(2)
Set book=live.Parent
If parked.ActiveCell.Address <> "$B$2" Then Err.Raise 5
If Not (ActiveWorkbook Is ThisWorkbook) Then Err.Raise 5
book.Activate
If Not (ActiveWindow Is live) Then Err.Raise 5
If ActiveCell.Address <> "$D$4" Then Err.Raise 5
End Sub"#,
    )?;
    // First two native windows belong to the inactive workbook.
    s.execute("selection.set", json!({"cell":"B2"}))?;
    s.new_window()?;
    s.execute("selection.set", json!({"cell":"D4"}))?;
    s.new_workbook();
    let mut b = Bridge::new(&s)?;
    b.configure(&mut ide)?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn dynamic_window_roots_follow_the_view_currently_shown_by_the_ui() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
If Not (ActiveWindow Is Windows(1)) Then Err.Raise 5
If ActiveCell.Address <> "$B$2" Then Err.Raise 5
End Sub"#,
    )?;
    let first = s.active_window().ok_or("window")?;
    s.execute("selection.set", json!({"cell":"B2"}))?;
    s.new_window()?;
    s.execute("selection.set", json!({"cell":"D4"}))?;
    s.show_window(first);
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_window_and_watch_collections_survive_originating_window_close() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim all As Windows, watches As Watches, book As Workbook
Set book=Workbooks.Add
Set all=Application.Windows
Set watches=Application.Watches
book.Close False
If all.Count <> 1 Then Err.Raise 5
If watches.Count <> 0 Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn last_window_close_prompt_cancellation_returns_false_and_keeps_the_view() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim book As Workbook, w As Window
Set book=Workbooks.Add
Set w=ActiveWindow
Range("C3").Value2=8
If w.Close Then Err.Raise 5
If Workbooks.Count <> 2 Then Err.Raise 5
If Not (ActiveWindow Is w) Then Err.Raise 5
End Sub"#,
    )?;
    b.set_prompt_services(gridcraft_vba::PromptServices { close: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut prompts = 0;
    for _ in 0..10000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(request) = b.take_prompt_request() {
            prompts += 1;
            b.respond_to_prompt(request.token, gridcraft_vba::PromptResponse::Cancel).map_err(|e| e.description)?;
        }
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(prompts, 1);
    assert_eq!(s.documents().len(), 2);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn beforeclose_creating_a_sibling_cannot_close_the_shared_workbook() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim first As Window
Set first=ActiveWindow
Range("B2").Value2=7
If Not first.Close(False) Then Err.Raise 5
If Windows.Count <> 1 Then Err.Raise 5
If Workbooks.Count <> 1 Then Err.Raise 5
If ActiveCell.Address <> "$A$1" Then Err.Raise 5
End Sub"#,
    )?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).code_name = Some("BookCode".into());
    ide.add_document_module("BookCode.cls","BookCode","Attribute VB_Name = \"BookCode\"\nPrivate Sub Workbook_BeforeClose(ByRef Cancel As Boolean)\nDim sibling As Window\nSet sibling=ThisWorkbook.NewWindow\nEnd Sub")?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.documents().len(), 1);
    assert!(s.doc()?.is_dirty());
    Ok(())
}
