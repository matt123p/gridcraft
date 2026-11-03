//! High invocation/lifecycle acceptance over the running project and native workbook commands.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value as CellValue},
};
use gridcraft_vba::{Action, Bridge, Ide, LifecyclePrompt, Mode, PromptResponse, PromptServices};
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str, document: Option<&str>) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).code_name = Some("BookCode".into());
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Main.bas", source)?;
    if let Some(source) = document {
        ide.add_document_module("BookCode.cls", "BookCode", &format!("Attribute VB_Name = \"BookCode\"\n{source}"))?;
    }
    b.configure_document(&mut ide, &s.doc()?.wb)?;
    Ok((s, b, ide))
}
fn finish(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    for _ in 0..10000 {
        b.tick(s, ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
fn value(s: &Session, address: &str) -> CellValue {
    s.doc().unwrap().wb.sheet(0).unwrap().value(CellRef::parse(address).unwrap())
}
#[cfg(not(target_arch = "wasm32"))]
fn path(name: &str) -> String {
    std::env::temp_dir().join(format!("gridcraft-continuation-{}-{name}.xlsx", std::process::id())).to_string_lossy().into_owned()
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_run_returns_values_preserves_globals_and_routes_nested_errors() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Public counter As Long
Sub Main()
Dim x As Long
Dim result As Variant
x = 5
result = Application.Run("Bump", x)
Range("A1").Value2 = result
Range("A2").Value2 = x
Range("A3").Value2 = counter
On Error Resume Next
result = Application.Run(Macro:="Twice", Arg1:=2)
Range("A4").Value2 = Err.Number
Err.Clear
result = Application.Run("Missing")
Range("A5").Value2 = Err.Number
Range("A6").Value2 = Application.Run("CountArgs", 1, "text", True)
End Sub
Public Function Bump(ByRef x As Long, Optional stepSize As Long = 2) As Long
counter = counter + 1
x = x + stepSize
Bump = Application.Run("Twice", x)
End Function
Public Function Twice(ByVal x As Long) As Long
counter = counter + 1
Twice = x * 2
End Function
Public Function Bad() As Long
Err.Raise 17
End Function
Public Function CountArgs(ParamArray args() As Variant) As Long
CountArgs = UBound(args) + 1
End Function
"#,
        None,
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    for (a, n) in [("A1", 14.0), ("A2", 5.0), ("A3", 2.0), ("A4", 448.0), ("A5", 1004.0), ("A6", 3.0)] {
        assert_eq!(value(&s, a), CellValue::Number(n), "{a}");
    }
    assert_eq!(s.doc()?.undo.len(), 1);
    Ok(())
}
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
#[cfg(not(all(target_arch = "wasm32", target_os = "unknown")))]
fn before_save_is_cancellable_and_after_save_runs_after_native_persistence() -> R {
    let file = path("before-after");
    let _ = std::fs::remove_file(&file);
    let source = format!("Sub Main()\nThisWorkbook.SaveAs Filename:=\"{file}\"\nEnd Sub\n");
    let doc = r#"Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
Me.Worksheets(1).Range("A1").Value2 = 42
If Me.Worksheets(1).Range("C1").Value2 = True Then Cancel = True
End Sub
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
Me.Worksheets(1).Range("B1").Value2 = Success
End Sub
"#;
    let (mut s, mut b, mut ide) = setup(&source, Some(doc))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    let (saved, _) = gridcraft_xlsx::read_xlsx(&std::fs::read(&file)?)?;
    let sh = saved.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.value(CellRef::parse("A1").ok_or("cell")?), CellValue::Number(42.0));
    assert!(sh.value(CellRef::parse("B1").ok_or("cell")?).is_empty());
    assert_eq!(value(&s, "B1"), CellValue::Bool(true));
    std::fs::remove_file(&file)?;
    s.execute("cell.set", serde_json::json!({"cell":"C1","input":"TRUE"}))?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(!std::path::Path::new(&file).exists());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn before_close_cancel_keeps_the_workbook_and_resumes_the_caller() -> R {
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nThisWorkbook.Close SaveChanges:=False\nRange(\"A1\").Value2 = 9\nEnd Sub\n",
        Some("Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)\nCancel = True\nEnd Sub\n"),
    )?;
    let owner = s.doc()?.uid;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(s.document(owner).is_some());
    assert_eq!(value(&s, "A1"), CellValue::Number(9.0));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn save_prompt_is_delivered_once_and_reset_invalidates_its_token() -> R {
    let (mut s, mut b, mut ide) = setup("Sub Main()\nThisWorkbook.Save\nEnd Sub\n", None)?;
    b.set_prompt_services(PromptServices { save: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut request = None;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if let Some(r) = b.take_prompt_request() {
            request = Some(r);
            break;
        }
    }
    let request = request.ok_or("no save picker request")?;
    assert!(matches!(request.prompt, LifecyclePrompt::SaveFile { .. }));
    assert_eq!(ide.mode(), Mode::Waiting);
    for _ in 0..2000 {
        b.tick(&mut s, &mut ide)?;
    }
    assert!(b.take_prompt_request().is_none());
    assert!(s.doc()?.path.is_none());
    b.dispatch(&mut s, &mut ide, Action::Reset)?;
    assert!(b.respond_to_prompt(request.token, PromptResponse::Path(Some("stale.xlsx".into()))).is_err());
    assert!(s.doc()?.path.is_none());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn dirty_close_prompt_cancel_does_not_close_or_replay() -> R {
    let (mut s, mut b, mut ide) = setup("Sub Main()\nRange(\"A1\").Value2 = 8\nThisWorkbook.Close\nRange(\"A2\").Value2 = 9\nEnd Sub\n", None)?;
    b.set_prompt_services(PromptServices { close: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut token = None;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if let Some(r) = b.take_prompt_request() {
            assert!(matches!(r.prompt, LifecyclePrompt::SaveChanges { .. }));
            token = Some(r.token);
            break;
        }
    }
    b.respond_to_prompt(token.ok_or("no close decision")?, PromptResponse::Cancel).map_err(|e| e.description)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, "A1"), CellValue::Number(8.0));
    assert_eq!(value(&s, "A2"), CellValue::Number(9.0));
    assert_eq!(s.documents().len(), 1);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn lifecycle_preflight_rejects_options_before_callbacks_or_prompts() -> R {
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nOn Error Resume Next\nThisWorkbook.SaveAs Password:=\"secret\"\nRange(\"A1\").Value2 = Err.Number\nEnd Sub\n",
        Some("Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)\nRange(\"B1\").Value2 = 99\nEnd Sub\n"),
    )?;
    b.set_prompt_services(PromptServices { save: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(value(&s, "A1"), CellValue::Number(1004.0));
    assert_eq!(value(&s, "B1"), CellValue::Empty);
    assert!(b.take_prompt_request().is_none());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_sink_receives_document_cancel_in_connection_order_during_wait() -> R {
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nBookCode.Connect\nThisWorkbook.Close SaveChanges:=False\nRange(\"A3\").Value2 = 3\nEnd Sub\n",
        Some(
            r#"Private WithEvents app As Excel.Application
Public Sub Connect()
Set app = Application
End Sub
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
Range("A1").Value2 = 1
Cancel = True
End Sub
Private Sub app_WorkbookBeforeClose(ByVal book As Excel.Workbook, ByRef Cancel As Boolean)
book.Worksheets(1).Range("A2").Value2 = Cancel
End Sub
"#,
        ),
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.documents().len(), 1);
    assert_eq!(value(&s, "A1"), CellValue::Number(1.0));
    assert_eq!(value(&s, "A2"), CellValue::Bool(true));
    assert_eq!(value(&s, "A3"), CellValue::Number(3.0));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn malformed_event_signature_refuses_native_effect_before_any_callback() -> R {
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nOn Error Resume Next\nThisWorkbook.Close SaveChanges:=False\nRange(\"A1\").Value2 = Err.Number\nEnd Sub\n",
        Some("Private Sub Workbook_BeforeClose(ByVal Cancel As Boolean)\nRange(\"B1\").Value2 = 99\nEnd Sub\n"),
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.documents().len(), 1);
    assert_eq!(value(&s, "A1"), CellValue::Number(1004.0));
    assert_eq!(value(&s, "B1"), CellValue::Empty);
    Ok(())
}
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn named_format_survives_save_picker_and_open_import_is_inert() -> R {
    let file = path("picker");
    let _ = std::fs::remove_file(&file);
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nThisWorkbook.SaveAs FileFormat:=51\nDim other As Workbook\nSet other = Workbooks.Open(UpdateLinks:=0)\nThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = 1\nEnd Sub\n",
        None,
    )?;
    b.set_prompt_services(PromptServices { save: true, open: true, close: true });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut prompts = 0;
    let imported = path("imported");
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(request) = b.take_prompt_request() {
            prompts += 1;
            let chosen = if matches!(request.prompt, LifecyclePrompt::OpenFile) {
                std::fs::copy(&file, &imported)?;
                imported.clone()
            } else {
                file.clone()
            };
            b.respond_to_prompt(request.token, PromptResponse::Path(Some(chosen))).map_err(|e| e.description)?;
        }
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(prompts, 2);
    assert_eq!(s.documents().len(), 2);
    assert!(!s.doc()?.trust.allows_code());
    std::fs::remove_file(file)?;
    std::fs::remove_file(imported)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn failed_save_reports_after_save_false_and_does_not_close() -> R {
    let destination = std::env::temp_dir().join(format!("gridcraft-no-parent-{}/book.xlsx", std::process::id())).to_string_lossy().into_owned();
    let main = format!(
        "Sub Main()\nOn Error Resume Next\nThisWorkbook.Close SaveChanges:=True, Filename:=\"{destination}\"\nRange(\"A1\").Value2 = Err.Number\nEnd Sub\n"
    );
    let (mut s, mut b, mut ide) =
        setup(&main, Some("Private Sub Workbook_AfterSave(ByVal Success As Boolean)\nMe.Worksheets(1).Range(\"B1\").Value2 = Success\nEnd Sub\n"))?;
    s.execute("cell.set", serde_json::json!({"cell":"C1","input":"1"}))?;
    let owner = s.doc()?.uid;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(s.document(owner).is_some());
    assert_eq!(value(&s, "A1"), CellValue::Number(1004.0));
    assert_eq!(value(&s, "B1"), CellValue::Bool(false));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn revoking_consent_during_picker_cancels_without_a_write() -> R {
    let (mut s, mut b, mut ide) = setup("Sub Main()\nThisWorkbook.Save\nEnd Sub\n", None)?;
    s.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
    b.consent(true);
    b.set_prompt_services(PromptServices { save: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut token = None;
    for _ in 0..100 {
        b.tick(&mut s, &mut ide)?;
        if let Some(request) = b.take_prompt_request() {
            token = Some(request.token);
            break;
        }
    }
    let token = token.ok_or("no prompt")?;
    b.consent(false);
    assert!(b.tick(&mut s, &mut ide).is_err());
    assert!(b.respond_to_prompt(token, PromptResponse::Path(Some("revoked.xlsx".into()))).is_err());
    assert!(s.doc()?.path.is_none());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unhandled_run_error_stops_the_caller_but_child_handled_errors_return_normally() -> R {
    let (mut s, mut b, mut ide) = setup(
        "Sub Main()\nOn Error Resume Next\nRange(\"A1\").Value2 = Application.Run(\"Handled\")\nApplication.Run \"Bad\"\nRange(\"A2\").Value2 = 99\nEnd Sub\nPublic Function Handled() As Long\nOn Error Resume Next\nErr.Raise 17\nHandled = Err.Number\nEnd Function\nPublic Sub Bad()\nErr.Raise 18\nEnd Sub\n",
        None,
    )?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if ide.mode() == Mode::Failed {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Failed);
    assert_eq!(value(&s, "A1"), CellValue::Number(17.0));
    assert_eq!(value(&s, "A2"), CellValue::Empty);
    assert!(ide.diagnostics.iter().any(|d| d.message.contains("Error 18:")));
    assert_eq!(s.doc()?.undo.len(), 1);
    Ok(())
}
