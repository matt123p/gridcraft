//! Original Excel-shaped source; all Office definitions and document assertions live here.
use gridcraft_engine::{
    Session,
    core::{CellRef, Value as CellValue},
};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use vba_foreign::{Host, HostArgument, HostCall, InvokeKind, Reply};
use vba_values::{ObjectRef, Value, VbaString};
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
fn invoke(b: &mut Bridge, s: &mut Session, target: ObjectRef, kind: InvokeKind, member: &str, args: Vec<Value>) -> Reply {
    b.host(s).invoke(HostCall {
        target,
        kind,
        member: member.into(),
        arguments: args.into_iter().map(|value| HostArgument { name: None, value, by_ref: false }).collect(),
    })
}
fn object(reply: Reply) -> R<ObjectRef> {
    match reply {
        Reply::Complete(c) => match c.value {
            Value::Object(o) => Ok(o),
            v => Err(format!("expected object: {v:?}").into()),
        },
        r => Err(format!("host error {r:?}").into()),
    }
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unchanged_typed_collections_count_item_names_and_sheet_order() -> R {
    let source = "Option Explicit\nSub Main()\nDim group As Excel.Worksheets,all As Sheets,ws As Worksheet,i As Long\nSet group=ThisWorkbook.Worksheets\nSet all=Sheets\nFor Each ws In group\ni=i+1\nws.Range(\"A1\")=i\nNext ws\nRange(\"B1\")=Worksheets.Count\nRange(\"B2\")=all.Item(1).Name\nRange(\"B3\")=Worksheets(Index:=1).Name\nRange(\"B4\")=Application.Sheets.Count\nEnd Sub";
    let (mut s, mut b, mut ide) = setup(source)?;
    let n = s.doc()?.wb.sheets.len();
    let first = s.doc()?.wb.sheet(0).ok_or("first")?.name.clone();
    run(&mut s, &mut b, &mut ide)?;
    for (i, sheet) in s.doc()?.wb.sheets.iter().enumerate() {
        assert_eq!(sheet.value(CellRef::parse("A1").ok_or("A1")?), CellValue::Number((i + 1) as f64));
    }
    assert_eq!(cell(&s, "B1")?, CellValue::Number(n as f64));
    assert_eq!(cell(&s, "B4")?, CellValue::Number(n as f64));
    assert_eq!(cell(&s, "B2")?, CellValue::Text(first.clone().into()));
    assert_eq!(cell(&s, "B3")?, CellValue::Text(first.into()));
    assert_eq!(ide.active_module().ok_or("source")?.text, source);
    assert!(!s.in_group(b.document()));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_defaults_indexed_item_and_cell_enumeration_mutate_engine() -> R {
    let source = "Sub Main()\nDim r As Excel.Range,c As Range\nSet r=Range(\"C1:D2\")\nr=0\nr(1)=10\nr.Item(RowIndex:=2,ColumnIndex:=\"B\")=20\nFor Each c In r\nc=c+1\nNext c\nRange(\"E1\")=r(1)\nOn Error Resume Next\nSet r(1)=Range(\"F1\")\nRange(\"E2\")=Err.Number\nEnd Sub";
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    for (address, expected) in [("C1", 11.0), ("D1", 1.0), ("C2", 1.0), ("D2", 21.0), ("E1", 11.0), ("E2", 13.0)] {
        assert_eq!(cell(&s, address)?, CellValue::Number(expected), "{address}");
    }
    assert!(cell(&s, "F1")?.is_empty());
    s.execute("edit.undo", serde_json::json!({}))?;
    assert!(cell(&s, "C1")?.is_empty());
    assert!(cell(&s, "D2")?.is_empty());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn collection_errors_and_stale_handles_are_explicit() -> R {
    let (mut s, mut b, _) = setup("Sub Main()\nEnd Sub")?;
    let root = ObjectRef { host: b.host(&mut s).identity(), object: 0, generation: b.document() };
    let collection = object(invoke(&mut b, &mut s, root, InvokeKind::Get, "Worksheets", vec![]))?;
    for value in [Value::Integer(0), Value::Integer(-1), Value::String(VbaString::from_utf8("Missing sheet")?)] {
        assert!(matches!(invoke(&mut b,&mut s,collection,InvokeKind::DefaultGet,"",vec![value]),Reply::Error(e) if e.number==9));
    }
    assert!(matches!(invoke(&mut b,&mut s,collection,InvokeKind::DefaultGet,"",vec![]),Reply::Error(e) if e.number==449));
    assert!(
        matches!(invoke(&mut b,&mut s,ObjectRef{generation:collection.generation+1,..collection},InvokeKind::Enumerate,"",vec![]),Reply::Error(e) if e.number==91)
    );
    assert!(matches!(invoke(&mut b,&mut s,collection,InvokeKind::Let,"Count",vec![Value::Integer(1)]),Reply::Error(e) if e.number==438));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn enumeration_preflights_handle_budget_without_partial_registry_exhaustion() -> R {
    let (mut s, mut b, _) = setup("Sub Main()\nEnd Sub")?;
    let app = ObjectRef { host: b.host(&mut s).identity(), object: 0, generation: b.document() };
    let large = object(invoke(&mut b, &mut s, app, InvokeKind::Get, "Range", vec![Value::String(VbaString::from_utf8("A1:A8190")?)]))?;
    assert!(matches!(invoke(&mut b,&mut s,large,InvokeKind::Enumerate,"",vec![]),Reply::Error(e) if e.number==7));
    let small = object(invoke(&mut b, &mut s, app, InvokeKind::Get, "Range", vec![Value::String(VbaString::from_utf8("B1:B2")?)]))?;
    assert!(
        matches!(invoke(&mut b,&mut s,small,InvokeKind::Enumerate,"",vec![]),Reply::Complete(c) if matches!(&c.value,Value::Array(a) if a.elements().len()==2))
    );
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_iteration_order_and_relative_item_offsets_match_public_contract() -> R {
    let source = "Sub Main()\nDim r As Range,c As Range,n As Long\nSet r=Range(\"B2:C3\")\nFor Each c In r\nn=n+1\nc=n\nNext\nRange(\"A1\")=r.Item(3,3).Address\nRange(\"A2\")=r.Item(0).Address\nEnd Sub";
    let (mut s, mut b, mut ide) = setup(source)?;
    run(&mut s, &mut b, &mut ide)?;
    for (address, expected) in [("B2", 1.0), ("C2", 2.0), ("B3", 3.0), ("C3", 4.0)] {
        assert_eq!(cell(&s, address)?, CellValue::Number(expected));
    }
    assert_eq!(cell(&s, "A1")?, CellValue::Text("$D$4".into()));
    assert_eq!(cell(&s, "A2")?, CellValue::Text("$C$1".into()));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn pause_safe_watches_and_trust_revocation_do_not_replay_loop_effects() -> R {
    let source = "Sub Main()\nDim ws As Worksheet\nFor Each ws In Worksheets\nws.Range(\"A1\")=7\nDebug.Print 1\nws.Range(\"A2\")=9\nNext\nEnd Sub";
    let (mut s, mut b, mut ide) = setup(source)?;
    ide.toggle_breakpoint_at(source.find("Debug.Print").ok_or("breakpoint")?)?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    drain(&mut s, &mut b, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(cell(&s, "A1")?, CellValue::Number(7.0));
    assert!(ide.evaluate("ws.Range(\"A1\")").is_err());
    s.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
    b.consent(false);
    assert!(b.tick(&mut s, &mut ide).is_err());
    assert!(!matches!(ide.mode(), Mode::Running | Mode::Waiting | Mode::Break));
    assert!(cell(&s, "A2")?.is_empty());
    assert!(!s.in_group(b.document()));
    Ok(())
}
