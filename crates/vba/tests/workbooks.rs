use gridcraft_engine::{
    DocState, Session, Trust,
    core::{CellRef, Value as CellValue},
    model::Workbook,
};
use gridcraft_vba::{Action, Bridge, Host, Ide, Mode, ObjectRef, Value};
use serde_json::json;
use vba_foreign::{HostArgument, HostCall, InvokeKind, Reply};
use vba_values::{ScalarKind, VbaArray, VbaString};
type R = Result<(), Box<dyn std::error::Error>>;

fn session() -> Session {
    let mut s = Session::new();
    s.new_workbook();
    s
}
fn setup(s: &Session, source: &str) -> Result<(Bridge, Ide), Box<dyn std::error::Error>> {
    let bridge = Bridge::new(s)?;
    let mut ide = Ide::new();
    ide.add_module("Test.bas", source)?;
    bridge.configure(&mut ide)?;
    Ok((bridge, ide))
}
fn drain(b: &mut Bridge, s: &mut Session, ide: &mut Ide) -> R {
    for _ in 0..1000 {
        b.tick(s, ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            return Ok(());
        }
    }
    Err("macro did not reach boundary".into())
}
fn value(s: &Session, address: &str) -> CellValue {
    s.doc().unwrap().wb.sheet(0).unwrap().value(CellRef::parse(address).unwrap())
}
fn run(s: &mut Session, source: &str) -> R {
    let (mut b, mut ide) = setup(s, source)?;
    b.dispatch(s, &mut ide, Action::Run)?;
    drain(&mut b, s, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
fn text(s: &str) -> Value {
    Value::String(VbaString::from_utf8(s).unwrap())
}
fn call(b: &mut Bridge, s: &mut Session, target: ObjectRef, kind: InvokeKind, member: &str, args: Vec<Value>) -> Reply {
    let mut host = b.host(s);
    host.invoke(HostCall {
        kind,
        target,
        member: member.into(),
        arguments: args.into_iter().map(|value| HostArgument { name: None, value, by_ref: false }).collect(),
    })
}
fn root(b: &mut Bridge, s: &mut Session) -> ObjectRef {
    let doc = b.document();
    let host = b.host(s);
    ObjectRef { host: host.identity(), object: 0, generation: doc }
}
fn complete(reply: Reply) -> Value {
    match reply {
        Reply::Complete(c) => c.value,
        other => panic!("{other:?}"),
    }
}
fn object(reply: Reply) -> ObjectRef {
    match complete(reply) {
        Value::Object(r) => r,
        other => panic!("{other:?}"),
    }
}
fn range(b: &mut Bridge, s: &mut Session, address: &str) -> ObjectRef {
    let r = root(b, s);
    object(call(b, s, r, InvokeKind::Call, "Range", vec![text(address)]))
}
fn error(reply: Reply, number: i32) {
    assert!(matches!(&reply,Reply::Error(e) if e.number==number), "expected {number}: {reply:?}");
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn fillcells_recalculates_and_has_one_undo_step() -> R {
    let mut s = session();
    run(&mut s, gridcraft_vba::DEMO)?;
    for row in 1..=10 {
        assert_eq!(value(&s, &format!("A{row}")), CellValue::Number(f64::from(row)));
    }
    assert_eq!(value(&s, "B1"), CellValue::Number(55.0));
    assert_eq!(s.doc()?.undo.len(), 1);
    assert_eq!(s.doc()?.undo.first().unwrap().label, "VBA Macro");
    s.execute("edit.undo", json!({}))?;
    assert!(value(&s, "A1").is_empty());
    assert!(value(&s, "B1").is_empty());
    s.execute("edit.redo", json!({}))?;
    assert_eq!(value(&s, "B1"), CellValue::Number(55.0));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn qualified_objects_named_arguments_and_error_handlers() -> R {
    let mut s = session();
    run(
        &mut s,
        "Option Explicit\nSub Main()\nDim ws As Object\nSet ws = ThisWorkbook.Worksheets(1)\nws.Cells(ColumnIndex:=2, RowIndex:=3).Value2 = 42\nOn Error Resume Next\nws.NoSuchMember\nRange(\"A1\").Value2 = Err.Number\nEnd Sub\n",
    )?;
    assert_eq!(value(&s, "B3"), CellValue::Number(42.0));
    assert_eq!(value(&s, "A1"), CellValue::Number(438.0));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn excel_two_endpoint_range_and_ownership_execute_unchanged_source() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Option Explicit
Sub Main()
Dim ws As Object
Dim block As Object
Set ws = ThisWorkbook.Worksheets(1)
Set block = ws.Range(ws.Cells(3, 4), ws.Cells(2, 2))
block.Value2 = 17
Range("A1").Value2 = block.Row
Range("A2").Value2 = block.Column
Range("A3").Value2 = block.Count
Range("A4").Value2 = block.Parent.Name
Range("A5").Value2 = block.Worksheet.Name
Range("A6").Value2 = block.Parent.Parent.Name
Range("A7").Value2 = block.Application.ThisWorkbook.Name
ws.Range(Cell2:="F2", Cell1:="E1").Value2 = 23
On Error Resume Next
ThisWorkbook.Range("A1").Value2 = 999
Range("A8").Value2 = Err.Number
Err.Clear
ThisWorkbook.Cells(1, 1).Value2 = 999
Range("A9").Value2 = Err.Number
Err.Clear
ThisWorkbook.ActiveWorkbook.Name
Range("A10").Value2 = Err.Number
End Sub
"#,
    )?;
    for address in ["B2", "C2", "D2", "B3", "C3", "D3"] {
        assert_eq!(value(&s, address), CellValue::Number(17.0));
    }
    for (address, number) in [("A1", 2.0), ("A2", 2.0), ("A3", 6.0), ("A8", 438.0), ("A9", 438.0), ("A10", 438.0)] {
        assert_eq!(value(&s, address), CellValue::Number(number));
    }
    assert_eq!(value(&s, "A4"), value(&s, "A5"));
    assert_eq!(value(&s, "A6"), value(&s, "A7"));
    assert_eq!(value(&s, "F2"), CellValue::Number(23.0));
    assert!(value(&s, "G2").is_empty());
    assert_eq!(s.doc()?.undo.len(), 1);
    s.execute("edit.undo", json!({}))?;
    assert!(value(&s, "B2").is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_endpoint_validation_preserves_identity_and_budgets() -> R {
    let mut s = session();
    s.execute("home.insertSheet", json!({}))?;
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let app = root(&mut b, &mut s);
    let ws = object(call(&mut b, &mut s, app, InvokeKind::Call, "Worksheets", vec![Value::Long(1)]));
    let other = object(call(&mut b, &mut s, app, InvokeKind::Call, "Worksheets", vec![Value::Long(2)]));
    let local = object(call(&mut b, &mut s, ws, InvokeKind::Call, "Range", vec![text("B2:C3")]));
    let remote = object(call(&mut b, &mut s, other, InvokeKind::Call, "Range", vec![text("A1")]));
    let block = object(call(&mut b, &mut s, ws, InvokeKind::Call, "Range", vec![Value::Object(local), text("A1")]));
    assert_eq!(complete(call(&mut b, &mut s, block, InvokeKind::Get, "Address", vec![])), text("$A$1:$C$3"));
    assert_eq!(complete(call(&mut b, &mut s, block, InvokeKind::Get, "Parent", vec![])), Value::Object(ws));
    assert_eq!(complete(call(&mut b, &mut s, block, InvokeKind::Get, "Application", vec![])), Value::Object(app));
    let before = s.doc()?.wb.clone();
    error(call(&mut b, &mut s, ws, InvokeKind::Call, "Range", vec![Value::Object(local)]), 1004);
    error(call(&mut b, &mut s, ws, InvokeKind::Call, "Range", vec![Value::Object(local), Value::Object(remote)]), 1004);
    error(call(&mut b, &mut s, ws, InvokeKind::Call, "Range", vec![Value::Object(app), text("A1")]), 13);
    for bad in
        [ObjectRef { host: app.host + 1, ..local }, ObjectRef { generation: app.generation + 1, ..local }, ObjectRef { object: u64::MAX, ..local }]
    {
        error(call(&mut b, &mut s, ws, InvokeKind::Call, "Range", vec![Value::Object(bad), text("A1")]), 91);
    }
    let whole = object(call(&mut b, &mut s, ws, InvokeKind::Call, "Range", vec![text("A1"), text("XFD1048576")]));
    error(call(&mut b, &mut s, whole, InvokeKind::Get, "Value2", vec![]), 7);
    error(call(&mut b, &mut s, ws, InvokeKind::Call, "Range", vec![text("A1"), Value::Null]), 94);
    error(call(&mut b, &mut s, block, InvokeKind::Let, "Row", vec![Value::Long(1)]), 438);
    error(call(&mut b, &mut s, block, InvokeKind::Get, "Parent", vec![Value::Long(1)]), 450);
    assert_eq!(*before, *s.doc()?.wb);
    Ok(())
}

// These passing refusal tests document incompatibility; they are not successful
// executions of the corresponding Excel workbook source.
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unsupported_excel_declarations_are_diagnosed() -> R {
    for declaration in ["Dim query As Excel.QueryTable", "Dim missing As MissingLibrary.Item"] {
        let mut s = session();
        let (mut b, mut ide) = setup(&s, &format!("Option Explicit\nSub Main()\n{declaration}\nEnd Sub\n"))?;
        assert!(b.dispatch(&mut s, &mut ide, Action::Run).is_err(), "{declaration}");
        assert!(!ide.diagnostics.is_empty(), "missing diagnostic: {declaration}");
        assert!(value(&s, "A1").is_empty());
        assert!(!s.in_group(b.document()));
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn collections_and_optional_cells_indexes_are_real_objects() -> R {
    let mut s = session();
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let app = root(&mut b, &mut s);
    let collection = object(call(&mut b, &mut s, app, InvokeKind::Get, "Worksheets", vec![]));
    let count = i32::try_from(s.doc()?.wb.sheets.len())?;
    assert_eq!(complete(call(&mut b, &mut s, collection, InvokeKind::Get, "Count", vec![])), Value::Long(count));
    let whole = object(call(&mut b, &mut s, app, InvokeKind::Get, "Cells", vec![]));
    assert_eq!(complete(call(&mut b, &mut s, whole, InvokeKind::Get, "CountLarge", vec![])), Value::Double(17_179_869_184.0));
    error(call(&mut b, &mut s, whole, InvokeKind::Get, "Count", vec![]), 6);
    let one = object(call(&mut b, &mut s, app, InvokeKind::Call, "Cells", vec![Value::Long(1)]));
    assert_eq!(complete(call(&mut b, &mut s, one, InvokeKind::Get, "Address", vec![])), text("$A$1"));
    let two = object(call(&mut b, &mut s, app, InvokeKind::Call, "Cells", vec![Value::Long(1), text("B")]));
    assert_eq!(complete(call(&mut b, &mut s, two, InvokeKind::Get, "Address", vec![])), text("$B$1"));
    let r = range(&mut b, &mut s, "A1:B2");
    let rows = object(call(&mut b, &mut s, r, InvokeKind::Get, "Rows", vec![]));
    assert_eq!(complete(call(&mut b, &mut s, rows, InvokeKind::Get, "Count", vec![])), Value::Long(2));
    let item = object(call(&mut b, &mut s, r, InvokeKind::Call, "Item", vec![Value::Long(1)]));
    assert_eq!(complete(call(&mut b, &mut s, item, InvokeKind::Get, "Address", vec![])), text("$A$1"));
    assert_eq!(complete(call(&mut b, &mut s, r, InvokeKind::Get, "Address", vec![Value::Boolean(false)])), text("$A1:$B2"));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn break_edit_resume_does_not_repeat_writes() -> R {
    let mut s = session();
    let source = "Sub Main()\nDim n As Long\nn = 1\nRange(\"A1\").Value2 = n\nn = n + 1\nRange(\"A2\").Value2 = n\nEnd Sub\n";
    let (mut b, mut ide) = setup(&s, source)?;
    ide.toggle_breakpoint_at(source.find("n = n +").unwrap())?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut b, &mut s, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(value(&s, "A1"), CellValue::Number(1.0));
    assert!(s.in_group(b.document()));
    let writes = s.journal.iter().filter(|j| j.0 == "range.setContents").count();
    ide.edit_local("n", "40")?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut b, &mut s, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed);
    assert_eq!(value(&s, "A2"), CellValue::Number(41.0));
    assert_eq!(s.journal.iter().filter(|j| j.0 == "range.setContents").count(), writes + 1);
    assert_eq!(s.doc()?.undo.len(), 1);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn failure_and_reset_keep_prior_effects_undoable() -> R {
    for stop in [false, true] {
        let mut s = session();
        let source = "Sub Main()\nRange(\"A1\").Value2 = 12\nApplication.Unsupported\nEnd Sub\n";
        let (mut b, mut ide) = setup(&s, source)?;
        if stop {
            ide.toggle_breakpoint_at(source.find("Application.Unsupported").unwrap())?;
        }
        b.dispatch(&mut s, &mut ide, Action::Run)?;
        drain(&mut b, &mut s, &mut ide)?;
        if stop {
            assert_eq!(ide.mode(), Mode::Break);
            b.dispatch(&mut s, &mut ide, Action::Reset)?;
        } else {
            assert_eq!(ide.mode(), Mode::Failed);
        }
        assert_eq!(value(&s, "A1"), CellValue::Number(12.0));
        assert!(!s.in_group(b.document()));
        assert_eq!(s.doc()?.undo.len(), 1);
        s.execute("edit.undo", json!({}))?;
        assert!(value(&s, "A1").is_empty());
    }
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn trust_checked_for_pure_code_and_every_host_call() -> R {
    let mut s = Session::new();
    s.add_document(DocState::new(Workbook::new(), None, "Untrusted".into(), Trust::FromFile));
    let (mut b, mut ide) = setup(&s, "Sub Main()\nDebug.Print 42\nEnd Sub\n")?;
    assert!(b.dispatch(&mut s, &mut ide, Action::Run).is_err());
    assert_eq!(ide.mode(), Mode::Design);
    assert!(!s.in_group(b.document()));
    b.consent(true);
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut b, &mut s, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed);
    assert_eq!(s.doc()?.trust, Trust::FromFile);
    let r = root(&mut b, &mut s);
    b.consent(false);
    error(call(&mut b, &mut s, r, InvokeKind::Call, "Range", vec![text("A1")]), 70);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn scalar_tags_literal_text_and_errors_roundtrip() -> R {
    let mut s = session();
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    for (input, expected) in [
        (Value::Empty, Value::Empty),
        (Value::Boolean(true), Value::Boolean(true)),
        (Value::Long(42), Value::Double(42.0)),
        (Value::Currency(12345), Value::Double(1.2345)),
        (Value::Date(45000.5), Value::Double(45000.5)),
        (text("00123"), Value::Double(123.0)),
        (text("=1+2"), Value::Double(3.0)),
        (text("'quoted"), text("quoted")),
        (text(""), Value::Empty),
        (Value::CvErr(2007), Value::CvErr(2007)),
    ] {
        let r = range(&mut b, &mut s, "A1");
        complete(call(&mut b, &mut s, r, InvokeKind::Let, "Value2", vec![input]));
        assert_eq!(complete(call(&mut b, &mut s, r, InvokeKind::Get, "Value2", vec![])), expected);
    }
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn rectangular_arrays_respect_bounds_and_validate_atomically() -> R {
    let mut s = session();
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let mut a = VbaArray::new(ScalarKind::Variant, vec![(-1, 0), (4, 5)], false)?;
    for (i, v) in [Value::Long(1), text("x"), Value::Boolean(true), Value::CvErr(2042)].into_iter().enumerate() {
        a.set(i, v)?;
    }
    let r = range(&mut b, &mut s, "B2:C3");
    complete(call(&mut b, &mut s, r, InvokeKind::Let, "Value", vec![Value::Array(a.clone())]));
    let Value::Array(read) = complete(call(&mut b, &mut s, r, InvokeKind::Get, "Value", vec![])) else { panic!("array expected") };
    assert_eq!(read.bounds(), &[(1, 2), (1, 2)]);
    assert_eq!(read.elements().first(), Some(&Value::Double(1.0)));
    let before = s.doc()?.wb.clone();
    a.set(3, Value::Null)?;
    error(call(&mut b, &mut s, r, InvokeKind::Let, "Value", vec![Value::Array(a)]), 94);
    assert_eq!(*before, *s.doc()?.wb);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn negative_calls_have_exact_errors_and_no_writes() -> R {
    let mut s = session();
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let root = root(&mut b, &mut s);
    let r = range(&mut b, &mut s, "A1");
    let before = s.doc()?.wb.clone();
    for (value, number) in [
        (Value::Null, 94),
        (Value::Missing, 449),
        (Value::Nothing, 13),
        (Value::LongLong(i64::MAX), 6),
        (Value::Double(f64::INFINITY), 6),
        (Value::CvErr(999), 13),
        (Value::String(VbaString::from_units(vec![0xd800])?), 13),
    ] {
        error(call(&mut b, &mut s, r, InvokeKind::Let, "Value2", vec![value]), number);
    }
    error(call(&mut b, &mut s, r, InvokeKind::Set, "Value2", vec![text("x")]), 13);
    error(call(&mut b, &mut s, root, InvokeKind::Call, "Cells", vec![Value::Long(0), Value::Long(1)]), 1004);
    let column = object(call(&mut b, &mut s, root, InvokeKind::Call, "Range", vec![text("A:A")]));
    error(call(&mut b, &mut s, column, InvokeKind::Get, "Value2", vec![]), 7);
    let qualified = object(call(&mut b, &mut s, root, InvokeKind::Call, "Range", vec![text("Sheet1!A1")]));
    assert_eq!(qualified, r);
    error(call(&mut b, &mut s, r, InvokeKind::Call, "Delete", vec![Value::Long(999)]), 1004);
    let foreign = ObjectRef { host: root.host + 1, ..r };
    error(call(&mut b, &mut s, foreign, InvokeKind::Get, "Value", vec![]), 91);
    assert_eq!(*before, *s.doc()?.wb);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn closing_code_owner_cancels_execution() -> R {
    for change in 1..2 {
        let mut s = session();
        let (mut b, mut ide) = setup(&s, "Sub Main()\nDim i As Long\nFor i = 1 To 100000\nCells(1,1).Value2 = i\nNext i\nEnd Sub\n")?;
        b.dispatch(&mut s, &mut ide, Action::Run)?;
        b.tick(&mut s, &mut ide)?;
        match change {
            0 => {
                s.new_workbook();
            }
            1 => s.close_document(0),
            _ => {
                s.execute("home.insertSheet", json!({}))?;
            }
        }
        assert!(b.tick(&mut s, &mut ide).is_err());
        assert_eq!(ide.mode(), Mode::Design);
        assert!(!s.in_group(b.document()));
        let root = root(&mut b, &mut s);
        error(call(&mut b, &mut s, root, InvokeKind::Get, "ActiveSheet", vec![]), 91);
    }
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn workbook_package_retains_vba_bytes_and_macro_results() -> R {
    let mut s = session();
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).vba_project = Some(b"authored inert project".to_vec().into());
    run(&mut s, gridcraft_vba::DEMO)?;
    let bytes = gridcraft_xlsx::write_xlsx_as(&s.doc()?.wb, gridcraft_xlsx::XlsxKind::MacroWorkbook)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert_eq!(wb.vba_project.as_deref().map(|a| a.as_slice()), Some(b"authored inert project".as_slice()));
    assert_eq!(wb.sheet(0).unwrap().value(CellRef::parse("A10").unwrap()), CellValue::Number(10.0));
    let b = Bridge::new(&s)?;
    assert!(b.import_project(&s).is_err()); // invalid project is inert and refused
    assert_eq!(s.doc()?.wb.vba_project, wb.vba_project);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn workbook_macro_strcomp_uses_the_configured_binary_provider() -> R {
    let mut s = session();
    run(&mut s, "Sub Main()\nRange(\"A1\").Value2 = StrComp(\"A\", \"a\")\nRange(\"A2\").Value2 = StrComp(\"a\", \"B\")\nEnd Sub\n")?;
    assert_eq!(value(&s, "A1"), CellValue::Number(-1.0));
    assert_eq!(value(&s, "A2"), CellValue::Number(1.0));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn protection_is_enforced_and_range_payload_is_bounded() -> R {
    let mut s = session();
    s.execute("review.protectSheet", json!({}))?;
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let r = range(&mut b, &mut s, "A1");
    let before = s.doc()?.wb.clone();
    error(call(&mut b, &mut s, r, InvokeKind::Let, "Value2", vec![Value::Long(7)]), 1004);
    assert_eq!(*before, *s.doc()?.wb);
    let large = range(&mut b, &mut s, "A1:A1000");
    error(call(&mut b, &mut s, large, InvokeKind::Let, "Value2", vec![text(&"x".repeat(32768))]), 7);
    let wrong = VbaArray::new(ScalarKind::Variant, vec![(0, 1)], false)?;
    error(call(&mut b, &mut s, r, InvokeKind::Let, "Value2", vec![Value::Array(wrong)]), 13);
    assert_eq!(*before, *s.doc()?.wb);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn text_format_supports_literal_values_errors_and_explicit_formulas() -> R {
    let mut s = session();
    s.execute("home.numberFormat", json!({"range":"A1","format":"@"}))?;
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let r = range(&mut b, &mut s, "A1");
    for (member, value, expected) in [
        ("Value2", text("=literal"), text("=literal")),
        ("Formula", text("=1+2"), Value::Double(3.0)),
        ("Value", Value::CvErr(2007), Value::CvErr(2007)),
        ("Value2", text("'quoted"), text("'quoted")),
        ("Value2", text("00123"), text("00123")),
        ("Value2", Value::Empty, Value::Empty),
    ] {
        complete(call(&mut b, &mut s, r, InvokeKind::Let, member, vec![value]));
        assert_eq!(complete(call(&mut b, &mut s, r, InvokeKind::Get, "Value2", vec![])), expected);
        let d = s.doc()?;
        let sh = d.wb.sheet(0).ok_or("sheet")?;
        assert_eq!(d.wb.styles.get(sh.style_id(CellRef::new(0, 0))).num_fmt.as_str(), "@");
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn macro_writes_text_format_and_records_replayable_typed_commands() -> R {
    let mut s = session();
    s.execute("home.numberFormat", json!({"range":"A1:C1","format":"@"}))?;
    s.start_recording();
    run(&mut s, "Sub Main()\nRange(\"A1\").Value2 = \"=literal\"\nRange(\"B1\").Formula = \"=1+2\"\nRange(\"C1\").Value2 = CVErr(2007)\nEnd Sub\n")?;
    let recording = s.stop_recording().ok_or("recording")?;
    assert_eq!(recording.steps.len(), 3);
    assert!(recording.steps.iter().all(|step| step.id == "range.setContents"));
    assert_eq!(value(&s, "A1"), CellValue::text("=literal"));
    assert_eq!(value(&s, "B1"), CellValue::Number(3.0));
    assert_eq!(value(&s, "C1"), CellValue::Error(gridcraft_engine::core::CellError::Div0));
    let after = s.doc()?.wb.clone();
    s.execute("edit.undo", json!({}))?;
    assert!(value(&s, "A1").is_empty());
    for step in recording.steps {
        s.execute(&step.id, step.params)?;
    }
    assert_eq!(*after, *s.doc()?.wb);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn workbooks_are_live_nominal_collections_and_activation_preserves_code_owner() -> R {
    let mut s = session();
    let owner = s.doc()?.uid;
    s.prefs.sheets_in_new_workbook = 3;
    run(
        &mut s,
        r#"Option Explicit
Sub Main()
Dim books As Excel.Workbooks
Dim created As Excel.Workbook
Dim book As Excel.Workbook
Dim total As Long
Set books = Application.Workbooks
Set created = Workbooks.Add(xlWBATWorksheet)
ThisWorkbook.Worksheets(1).Range("A1").Value2 = books.Count
ThisWorkbook.Worksheets(1).Range("A2").Value2 = created.Worksheets.Count
ThisWorkbook.Worksheets(1).Range("A3").Value2 = ActiveWorkbook.Name
For Each book In books
 total = total + 1
Next book
ThisWorkbook.Worksheets(1).Range("A4").Value2 = total
ThisWorkbook.Activate
Range("A5").Value2 = ActiveWorkbook.Name
created.Activate
Range("B1").Value2 = 99
created.Close SaveChanges:=False
ThisWorkbook.Worksheets(1).Range("A6").Value2 = books.Count
On Error Resume Next
created.Name
ThisWorkbook.Worksheets(1).Range("A7").Value2 = Err.Number
End Sub
"#,
    )?;
    assert_eq!(s.doc()?.uid, owner);
    assert_eq!(value(&s, "A1"), CellValue::Number(2.0));
    assert_eq!(value(&s, "A2"), CellValue::Number(1.0));
    assert_eq!(value(&s, "A3"), CellValue::Text("Book2".into()));
    assert_eq!(value(&s, "A4"), CellValue::Number(2.0));
    assert_eq!(value(&s, "A5"), CellValue::Text("Book1".into()));
    assert_eq!(value(&s, "A6"), CellValue::Number(1.0));
    assert_eq!(value(&s, "A7"), CellValue::Number(91.0));
    assert!(value(&s, "B1").is_empty());
    assert_eq!(s.doc()?.undo.len(), 1);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn sheet_lifecycle_remaps_handles_and_preserves_deleted_tombstones() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Option Explicit
Sub Main()
Dim original As Excel.Worksheet
Dim added As Excel.Worksheet
Dim saved As Excel.Range
Set original = Worksheets(1)
Set saved = original.Range("B1")
Set added = Worksheets.Add(Before:=original, Count:=2, Type:=xlWorksheet)
added.Name = "Added"
saved.Value2 = 19
original.Range("A1").Value2 = original.Index
added.Move After:=original
original.Range("A2").Value2 = added.Index
added.Copy Before:=original
original.Range("A3").Value2 = Worksheets.Count
added.Visible = xlSheetVeryHidden
original.Range("A4").Value2 = added.Visible
added.Visible = xlSheetVisible
added.Delete
original.Range("A5").Value2 = Worksheets.Count
On Error Resume Next
added.Name
original.Range("A6").Value2 = Err.Number
Worksheets.Add
Err.Clear
added.Name
original.Range("A7").Value2 = Err.Number
End Sub
"#,
    )?;
    let original = s.doc()?.wb.sheet_index("Sheet1").ok_or("missing original")?;
    let sh = s.doc()?.wb.sheet(original).ok_or("sheet")?;
    for (addr, expected) in [("A1", 3.0), ("A2", 3.0), ("A3", 4.0), ("A4", 2.0), ("A5", 3.0), ("A6", 91.0), ("A7", 91.0), ("B1", 19.0)] {
        assert_eq!(sh.value(CellRef::parse(addr).ok_or("address")?), CellValue::Number(expected), "{addr}");
    }
    assert_eq!(s.doc()?.undo.len(), 1);
    s.execute("edit.undo", json!({}))?;
    assert_eq!(s.doc()?.wb.sheets.len(), 1);
    assert!(value(&s, "B1").is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn lifecycle_rejects_invalid_options_and_protected_structure_without_changes() -> R {
    let mut s = session();
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let app = root(&mut b, &mut s);
    let books = object(call(&mut b, &mut s, app, InvokeKind::Get, "Workbooks", vec![]));
    let sheets = object(call(&mut b, &mut s, app, InvokeKind::Get, "Worksheets", vec![]));
    let sheet = object(call(&mut b, &mut s, sheets, InvokeKind::DefaultGet, "", vec![Value::Long(1)]));
    let wb = object(call(&mut b, &mut s, books, InvokeKind::DefaultGet, "", vec![text("bOoK1")]));
    let before = s.doc()?.wb.clone();
    error(call(&mut b, &mut s, books, InvokeKind::Get, "Item", vec![Value::Long(0)]), 9);
    error(call(&mut b, &mut s, books, InvokeKind::Call, "Add", vec![Value::Long(-4109)]), 1004);
    error(call(&mut b, &mut s, sheets, InvokeKind::Call, "Add", vec![Value::Object(sheet), Value::Object(sheet)]), 1004);
    error(call(&mut b, &mut s, sheets, InvokeKind::Call, "Add", vec![Value::Missing, Value::Missing, Value::Long(0)]), 1004);
    error(call(&mut b, &mut s, wb, InvokeKind::Call, "Save", vec![]), 1004);
    error(call(&mut b, &mut s, wb, InvokeKind::Call, "Close", vec![]), 1004);
    assert_eq!(*before, *s.doc()?.wb);
    assert!(s.take_ui_requests().is_empty());
    s.execute("review.protectWorkbook", json!({"on":true}))?;
    let protected = s.doc()?.wb.clone();
    error(call(&mut b, &mut s, sheets, InvokeKind::Call, "Add", vec![]), 1004);
    error(call(&mut b, &mut s, sheet, InvokeKind::Call, "Delete", vec![]), 1004);
    error(call(&mut b, &mut s, sheet, InvokeKind::Call, "Copy", vec![Value::Object(sheet)]), 1004);
    error(call(&mut b, &mut s, sheet, InvokeKind::Let, "Visible", vec![Value::Long(2)]), 1004);
    assert_eq!(*protected, *s.doc()?.wb);
    b.finish(&mut s)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn workbook_files_roundtrip_copy_keeps_saved_state_and_open_never_replaces_owner() -> R {
    let folder = std::env::temp_dir().join(format!("gridcraft-om04-{}", std::process::id()));
    std::fs::create_dir_all(&folder)?;
    let path = folder.join("original.xlsm").to_string_lossy().into_owned();
    let copy = folder.join("copy.xlsm").to_string_lossy().into_owned();
    let mut s = session();
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).vba_project = Some(b"authored inert OM04 project".to_vec().into());
    let owner = s.doc()?.uid;
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let app = root(&mut b, &mut s);
    let wb = object(call(&mut b, &mut s, app, InvokeKind::Get, "ThisWorkbook", vec![]));
    let cell = range(&mut b, &mut s, "A1");
    complete(call(&mut b, &mut s, cell, InvokeKind::Let, "Value2", vec![Value::Long(12)]));
    complete(call(&mut b, &mut s, wb, InvokeKind::Call, "SaveAs", vec![text(&path), Value::Long(52)]));
    assert_eq!(complete(call(&mut b, &mut s, wb, InvokeKind::Get, "Saved", vec![])), Value::Boolean(true));
    assert_eq!(complete(call(&mut b, &mut s, wb, InvokeKind::Get, "Name", vec![])), text("original.xlsm"));
    complete(call(&mut b, &mut s, cell, InvokeKind::Let, "Value2", vec![Value::Long(23)]));
    let saved = s.doc()?.saved.clone();
    complete(call(&mut b, &mut s, wb, InvokeKind::Call, "SaveCopyAs", vec![text(&copy)]));
    assert_eq!(s.doc()?.saved, saved);
    assert!(s.doc()?.is_dirty());
    assert_eq!(s.doc()?.path.as_deref(), Some(path.as_str()));
    let (copy_wb, _) = gridcraft_engine::io::open_bytes("copy.xlsm", &std::fs::read(&copy)?)?;
    assert_eq!(copy_wb.vba_project.as_deref().map(|v| v.as_slice()), Some(b"authored inert OM04 project".as_slice()));
    assert_eq!(copy_wb.sheet(0).ok_or("sheet")?.value(CellRef::parse("A1").ok_or("cell")?), CellValue::Number(23.0));
    let books = object(call(&mut b, &mut s, app, InvokeKind::Get, "Workbooks", vec![]));
    let opened = object(call(&mut b, &mut s, books, InvokeKind::Call, "Open", vec![text(&copy), Value::Long(0), Value::Boolean(false)]));
    assert_eq!(s.documents().len(), 2);
    assert!(s.document(owner).is_some());
    assert_eq!(s.doc()?.trust, Trust::FromFile);
    // Open never grants execution consent in the imported document.
    error(call(&mut b, &mut s, opened, InvokeKind::Get, "Name", vec![]), 70);
    s.set_active(s.document_index(owner).ok_or("owner")?);
    let before = s.documents().len();
    error(call(&mut b, &mut s, books, InvokeKind::Call, "Open", vec![text(&folder.join("missing.xlsx").to_string_lossy())]), 1004);
    error(call(&mut b, &mut s, books, InvokeKind::Call, "Open", vec![text(&path), Value::Missing, Value::Boolean(true)]), 1004);
    assert_eq!(s.documents().len(), before);
    b.finish(&mut s)?;
    // Open must retain even a pristine blank code-owner workbook.
    let mut blank = session();
    let blank_id = blank.doc()?.uid;
    let (mut bridge, _) = setup(&blank, "Sub Main()\nEnd Sub\n")?;
    let app = root(&mut bridge, &mut blank);
    let books = object(call(&mut bridge, &mut blank, app, InvokeKind::Get, "Workbooks", vec![]));
    object(call(&mut bridge, &mut blank, books, InvokeKind::Call, "Open", vec![text(&path)]));
    assert_eq!(blank.documents().len(), 2);
    assert!(blank.document(blank_id).is_some());
    bridge.finish(&mut blank)?;
    std::fs::remove_dir_all(folder)?;
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_calculation_settings_are_live_validated_and_grouped_across_documents() -> R {
    let mut s = session();
    let first = s.doc()?.uid;
    s.new_workbook();
    run(
        &mut s,
        r#"Sub Main()
Application.Iteration = True
Application.MaxIterations = 42
Application.MaxChange = 0.005
Range("A1").Value2 = Application.MaxIterations
Application.CalculateFullRebuild
On Error Resume Next
Application.MaxIterations = 0
Range("A2").Value2 = Err.Number
Err.Clear
Application.MaxChange = -1
Range("A3").Value2 = Err.Number
End Sub
"#,
    )?;
    for doc in s.documents() {
        assert!(doc.wb.calc.iterative);
        assert_eq!(doc.wb.calc.max_iterations, 42);
        assert_eq!(doc.wb.calc.max_change, 0.005);
        assert_eq!(doc.undo.len(), 1);
    }
    assert_eq!(value(&s, "A1"), CellValue::Number(42.0));
    assert_eq!(value(&s, "A2"), CellValue::Number(1004.0));
    assert_eq!(value(&s, "A3"), CellValue::Number(1004.0));
    s.execute_in(gridcraft_engine::Target::Doc(first), "edit.undo", json!({}))?;
    assert!(!s.document(first).ok_or("first")?.wb.calc.iterative);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn workbook_structure_protection_uses_native_passwords_and_blocks_lifecycle() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
ThisWorkbook.Protect Password:="secret", Structure:=True
Range("A1").Value2 = ThisWorkbook.ProtectStructure
On Error Resume Next
Worksheets.Add
Range("A2").Value2 = Err.Number
Err.Clear
ThisWorkbook.Unprotect "wrong"
Range("A3").Value2 = Err.Number
Range("A4").Value2 = ThisWorkbook.ProtectStructure
ThisWorkbook.Unprotect "secret"
Range("A5").Value2 = ThisWorkbook.ProtectStructure
End Sub
"#,
    )?;
    assert_eq!(value(&s, "A1"), CellValue::Bool(true));
    assert_eq!(value(&s, "A2"), CellValue::Number(1004.0));
    assert_eq!(value(&s, "A3"), CellValue::Number(1004.0));
    assert_eq!(value(&s, "A4"), CellValue::Bool(true));
    assert_eq!(value(&s, "A5"), CellValue::Bool(false));
    assert_eq!(s.doc()?.wb.sheets.len(), 1);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn application_wide_calculation_refuses_untrusted_participants_before_mutation() -> R {
    let mut s = session();
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let owner = s.doc()?.uid;
    s.add_document(DocState::new(Workbook::new(), None, "External".into(), Trust::FromFile));
    s.set_active(s.document_index(owner).ok_or("owner")?);
    let app = root(&mut b, &mut s);
    let before = s.documents().iter().map(|d| d.wb.clone()).collect::<Vec<_>>();
    for (kind, member, args) in [
        (InvokeKind::Let, "Iteration", vec![Value::Boolean(true)]),
        (InvokeKind::Let, "Calculation", vec![Value::Long(-4135)]),
        (InvokeKind::Call, "CalculateFull", vec![]),
        (InvokeKind::Call, "CalculateFullRebuild", vec![]),
        (InvokeKind::Call, "Calculate", vec![]),
    ] {
        error(call(&mut b, &mut s, app, kind, member, args), 70);
    }
    for (old, doc) in before.iter().zip(s.documents()) {
        assert_eq!(**old, *doc.wb);
    }
    b.finish(&mut s)?;
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn sheet_transfers_refuse_cross_document_and_document_module_copies() -> R {
    let mut s = session();
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.code_name = Some("SheetCode".into());
    let owner = s.doc()?.uid;
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let app = root(&mut b, &mut s);
    let original = object(call(&mut b, &mut s, app, InvokeKind::Get, "ActiveSheet", vec![]));
    s.new_workbook();
    let foreign = object(call(&mut b, &mut s, app, InvokeKind::Get, "ActiveSheet", vec![]));
    let before = s.documents().iter().map(|d| d.wb.clone()).collect::<Vec<_>>();
    error(call(&mut b, &mut s, original, InvokeKind::Call, "Move", vec![Value::Object(foreign)]), 1004);
    error(call(&mut b, &mut s, original, InvokeKind::Call, "Copy", vec![Value::Object(original)]), 1004);
    error(call(&mut b, &mut s, original, InvokeKind::Call, "Copy", vec![]), 1004);
    for (old, doc) in before.iter().zip(s.documents()) {
        assert_eq!(**old, *doc.wb);
    }
    assert_ne!(s.doc()?.uid, owner);
    b.finish(&mut s)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn failed_save_and_close_preserve_document_then_explicit_save_close_succeeds() -> R {
    let mut s = session();
    let owner = s.doc()?.uid;
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let app = root(&mut b, &mut s);
    let books = object(call(&mut b, &mut s, app, InvokeKind::Get, "Workbooks", vec![]));
    let created = object(call(&mut b, &mut s, books, InvokeKind::Call, "Add", vec![]));
    let added = s.doc()?.uid;
    let cell = range(&mut b, &mut s, "A1");
    complete(call(&mut b, &mut s, cell, InvokeKind::Let, "Value2", vec![Value::Long(42)]));
    let before = s.doc()?.wb.clone();
    let missing = std::env::temp_dir().join(format!("gridcraft-no-dir-{}", added)).join("book.xlsx").to_string_lossy().into_owned();
    error(call(&mut b, &mut s, created, InvokeKind::Call, "SaveAs", vec![text(&missing)]), 1004);
    error(call(&mut b, &mut s, created, InvokeKind::Call, "Close", vec![Value::Boolean(true), text(&missing)]), 1004);
    assert_eq!(*before, *s.doc()?.wb);
    assert_eq!(s.doc()?.path, None);
    assert!(s.take_ui_requests().is_empty());
    let path = std::env::temp_dir().join(format!("gridcraft-save-close-{}-{}.xlsx", std::process::id(), added)).to_string_lossy().into_owned();
    complete(call(&mut b, &mut s, created, InvokeKind::Call, "Close", vec![Value::Boolean(true), text(&path)]));
    assert_eq!(s.doc()?.uid, owner);
    assert!(s.document(added).is_none());
    let (wb, _) = gridcraft_engine::io::open_bytes("book.xlsx", &std::fs::read(&path)?)?;
    assert_eq!(wb.sheet(0).ok_or("sheet")?.value(CellRef::parse("A1").ok_or("cell")?), CellValue::Number(42.0));
    std::fs::remove_file(path)?;
    b.finish(&mut s)?;
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn copying_a_table_sheet_refuses_instead_of_silently_dropping_the_table() -> R {
    let mut s = session();
    s.execute("range.setValues", json!({"range":"A1:B2","values":[["Name","Value"],["Example",12]]}))?;
    s.execute("insert.table", json!({"range":"A1:B2","header":true}))?;
    let (mut b, _) = setup(&s, "Sub Main()\nEnd Sub\n")?;
    let app = root(&mut b, &mut s);
    let sheet = object(call(&mut b, &mut s, app, InvokeKind::Get, "ActiveSheet", vec![]));
    let before = s.doc()?.wb.clone();
    error(call(&mut b, &mut s, sheet, InvokeKind::Call, "Copy", vec![Value::Object(sheet)]), 1004);
    assert_eq!(*before, *s.doc()?.wb);
    assert_eq!(s.doc()?.wb.sheets.len(), 1);
    b.finish(&mut s)?;
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn lifecycle_function_expressions_support_omitted_parentheses() -> R {
    let mut s = session();
    run(
        &mut s,
        r#"Sub Main()
Dim added As Excel.Worksheet
Dim removed As Boolean
Set added = Worksheets.Add
removed = added.Delete
Range("A1").Value2 = removed
Range("A2").Value2 = Workbooks.Count
On Error Resume Next
Range("A3").Value2 = Application.Run("MissingMacro")
Range("A4").Value2 = Err.Number
End Sub
"#,
    )?;
    assert_eq!(s.doc()?.wb.sheets.len(), 1);
    assert_eq!(value(&s, "A1"), CellValue::Bool(true));
    assert_eq!(value(&s, "A2"), CellValue::Number(1.0));
    assert_eq!(value(&s, "A4"), CellValue::Number(1004.0));
    Ok(())
}
