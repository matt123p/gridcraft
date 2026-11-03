//! OM15 native and headless service acceptance using actual suspended VBA calls.
use gridcraft_engine::{Session, core::CellRef};
use gridcraft_vba::{Action, Bridge, Ide, InteractionKind, InteractionResponse, InteractionServices, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Interactions.bas", source)?;
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
fn autocomplete_uses_native_contiguous_column_logic_without_changing_selection() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Range("A1").Value2="Über"
Range("A3").Value2="über"
If Range("A2").AutoComplete("ÜB")<>"Über" Then Err.Raise 5
Range("A3").Value2="Übung"
If Range("A2").AutoComplete("üb")<>"" Then Err.Raise 5
Range("A3").ClearContents
If Range("A4").AutoComplete("üb")<>"" Then Err.Raise 5
If Range("A2").AutoComplete("=")<>"" Then Err.Raise 5
If ActiveCell.Address<>"$A$1" Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn checkbox_controls_preserve_values_and_formulas_rebase_and_roundtrip() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim c As CellControl
Range("B1").Value2="text"
Range("C1").Formula="=1+2"
Set c=Range("A1:C1").CellControl
If c.Type<>xlTypeNone Then Err.Raise 5
c.SetCheckbox
If c.Type<>xlTypeCheckbox Then Err.Raise 5
If Range("A1").Value2<>False Then Err.Raise 5
If Range("B1").Value2<>"text" Or Range("C1").Value2<>3 Then Err.Raise 5
If Not IsEmpty(Range("A1:D1").CellControl.Type) Then Err.Raise 5
Rows(1).Insert
If c.Type<>xlTypeCheckbox Then Err.Raise 5
Range("A2:C2").ClearFormats
If c.Type<>xlTypeNone Then Err.Raise 5
If Range("B2").Value2<>"text" Then Err.Raise 5
Range("A3").CellControl.SetCheckbox
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let sh = wb.sheet(0).ok_or("sheet")?;
    assert_eq!(wb.styles.get(sh.style_id(CellRef::new(2, 0))).num_fmt.as_str(), "checkbox");
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.cells.iter().next().is_none());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn headless_interactions_and_unsupported_options_refuse_without_fake_results() -> R {
    let source = r#"Sub Main()
Dim d As Dialog, ds As Dialogs
Set ds=Application.Dialogs
If ds.Count<>7 Then Err.Raise 5
Set d=ds(xlDialogColumnWidth)
If Not (d Is ds.Item(xlDialogColumnWidth)) Then Err.Raise 5
If Not (d.Parent Is Application) Then Err.Raise 5
For Each d In ds
If Not (d.Application Is Application) Then Err.Raise 5
Next
On Error Resume Next
x=Application.GetOpenFilename
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
x=ds(xlDialogColumnWidth).Show
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
x=Application.InputBox("Value",Type:=8)
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
x=Application.GetSaveAsFilename(FileFilter:="Text,*.txt")
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
End Sub"#;
    let (mut s, mut b, mut ide) = setup(source)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(s.ui_requests.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn picker_continuations_return_paths_or_false_without_opening_or_saving() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
If Application.GetOpenFilename(MultiSelect:=False)<>False Then Err.Raise 5
If Application.GetSaveAsFilename("report.xlsx")<>"/chosen/report.xlsx" Then Err.Raise 5
Range("A1").Value2=7
End Sub"#,
    )?;
    b.set_interaction_services(InteractionServices { open_picker: true, save_picker: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut count = 0;
    for _ in 0..10000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(r) = b.take_interaction_request() {
            assert!(b.interaction_permitted(&s, &r));
            assert!(b.take_interaction_request().is_none());
            let response = match r.kind {
                InteractionKind::OpenFilename => InteractionResponse::Path(None),
                InteractionKind::SaveFilename { suggested } => {
                    assert_eq!(suggested, "report.xlsx");
                    InteractionResponse::Path(Some("/chosen/report.xlsx".into()))
                }
                _ => return Err("unexpected request".into()),
            };
            b.respond_to_interaction(r.token, Ok(response)).map_err(|e| e.description)?;
            assert!(b.respond_to_interaction(r.token, Ok(InteractionResponse::Path(None))).is_err());
            count += 1;
        }
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(count, 2);
    assert_eq!(s.documents().len(), 1);
    assert!(s.doc()?.path.is_none());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn reset_and_consent_revocation_cancel_owned_interaction_tokens() -> R {
    for reset in [true, false] {
        let (mut s, mut b, mut ide) = setup("Sub Main()\nx=Application.GetOpenFilename\nRange(\"A1\").Value2=99\nEnd Sub")?;
        s.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
        b.consent(true);
        b.set_interaction_services(InteractionServices { open_picker: true, ..Default::default() });
        b.dispatch(&mut s, &mut ide, Action::Run)?;
        let mut request = None;
        for _ in 0..100 {
            b.tick(&mut s, &mut ide)?;
            if let Some(r) = b.take_interaction_request() {
                request = Some(r);
                break;
            }
        }
        let r = request.ok_or("request missing")?;
        if reset {
            b.dispatch(&mut s, &mut ide, Action::Reset)?;
            assert!(b.respond_to_interaction(r.token, Ok(InteractionResponse::Path(None))).is_err());
        } else {
            b.consent(false);
            assert!(!b.interaction_permitted(&s, &r));
            assert!(b.tick(&mut s, &mut ide).is_err());
            assert_eq!(ide.mode(), Mode::Design);
            assert!(!b.interaction_pending(r.token));
        }
        assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 0)).is_empty());
    }
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn protected_and_oversized_controls_and_unknown_dialog_arguments_fail_before_mutation() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
ActiveSheet.Protect
On Error Resume Next
Range("A1").CellControl.SetCheckbox
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
ActiveSheet.Unprotect
On Error Resume Next
Range("A1:A100000").CellControl.SetCheckbox
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>7 Then Err.Raise 5
On Error Resume Next
x=Application.Dialogs(xlDialogZoom).Show(100)
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
x=Application.GetOpenFilename(Unknown:="not ignored")
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>448 Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.cells.iter().next().is_none());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn checkbox_value_changes_deliver_exact_native_event_targets_once() -> R {
    let (mut s, mut b, mut ide) = setup("Sub Main()\nRange(\"A1:B1\").CellControl.SetCheckbox\nRange(\"A1:B1\").CellControl.SetCheckbox\nEnd Sub")?;
    s.execute("cell.set", json!({"cell":"B1","input":"keep"}))?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.code_name = Some("SheetCode".into());
    ide.add_document_module(
        "SheetCode.cls",
        "SheetCode",
        r#"Attribute VB_Name = "SheetCode"
Private Sub Worksheet_Change(ByVal Target As Range)
If Target.Address="$A$1" Then
Range("Y1").Value2=Target.Address
Range("Z1").Value2=Range("Z1").Value2+1
End If
End Sub"#,
    )?;
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    finish(&mut s, &mut b, &mut ide)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.value(CellRef::new(0, 24)).display(), "$A$1");
    assert_eq!(sh.value(CellRef::new(0, 25)).display(), "1");
    Ok(())
}
#[cfg(not(all(target_arch = "wasm32", target_os = "unknown")))]
#[test]
fn find_file_uses_the_existing_loader_only_after_a_successful_picker_result() -> R {
    let path = std::env::temp_dir().join(format!("gridcraft-om15-find-{}.xlsx", std::process::id()));
    let mut book = gridcraft_engine::model::Workbook::new();
    book.sheet_mut(0).ok_or("sheet")?.set_value(CellRef::new(1, 1), gridcraft_engine::core::Value::Number(42.));
    std::fs::write(&path, gridcraft_xlsx::write_xlsx(&book)?)?;
    let (mut s, mut b, mut ide) = setup("Sub Main()\nIf Not Application.FindFile Then Err.Raise 5\nEnd Sub")?;
    b.set_interaction_services(InteractionServices { open_picker: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut requests = 0;
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(r) = b.take_interaction_request() {
            assert!(matches!(r.kind, InteractionKind::FindFile));
            assert_eq!(s.documents().len(), 1);
            b.respond_to_interaction(r.token, Ok(InteractionResponse::Path(Some(path.to_string_lossy().into_owned())))).map_err(|e| e.description)?;
            requests += 1;
        }
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(requests, 1);
    assert_eq!(s.documents().len(), 2);
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(1, 1)).display(), "42");
    std::fs::remove_file(path)?;
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn malformed_modal_selection_refuses_before_delivering_a_request() -> R {
    for budget in [false, true] {
        let expected = if budget { 7 } else { 1004 };
        let source = format!(
            "Sub Main()\nOn Error Resume Next\nx=Application.GetOpenFilename\nactual=Err.Number\nErr.Clear\nOn Error GoTo 0\nIf actual<>{expected} Then Err.Raise 5\nEnd Sub"
        );
        let (mut s, mut b, mut ide) = setup(&source)?;
        if budget {
            s.doc_mut()?.selection.ranges = vec![gridcraft_engine::core::RangeRef::cell(CellRef::new(0, 0)); 65];
        } else {
            s.doc_mut()?.selection.active = CellRef::new(u32::MAX, 0);
        }
        b.set_interaction_services(InteractionServices { open_picker: true, ..Default::default() });
        finish(&mut s, &mut b, &mut ide)?;
        assert!(b.take_interaction_request().is_none());
        assert!(s.ui_requests.is_empty());
    }
    Ok(())
}
