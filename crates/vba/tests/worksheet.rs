//! Independently authored worksheet procedures; public documentation supplies behavior only.
use gridcraft_engine::{
    Session,
    calc::{CustomFunctionContext, CustomFunctions},
    core::{CellError, CellRef, Value},
};
use gridcraft_vba::{
    Bridge, Ide,
    worksheet::{Functions, Project},
};
use serde_json::json;
use std::sync::{Arc, atomic::Ordering};
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<(Session, Project, Arc<Functions>, Bridge), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Numbers.bas", source)?;
    let project = Project::from_ide(&ide, None)?;
    project.store(&mut s)?;
    let functions = Arc::new(Functions::compile(&project, &b, true)?);
    s.doc_mut()?.calc.custom_functions = Some(functions.clone());
    Ok((s, project, functions, b))
}
fn cell(s: &Session, address: &str) -> Value {
    s.doc().unwrap().wb.sheet(0).unwrap().value(CellRef::parse(address).unwrap())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_arguments_dependencies_and_builtins() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Public Function Scale(ByVal x As Double, Optional ByVal factor As Double = 3) As Double\nScale = x * factor\nEnd Function\nPrivate Function Hidden() As Long\nHidden = 99\nEnd Function\nSub NeedsArgument(x As Long)\nEnd Sub\nPublic Sub Fill()\nRange(\"D1\").Value2 = 8\nEnd Sub\n",
    )?;
    s.execute("cell.set", json!({"cell":"A1","input":"7"}))?;
    s.execute("cell.set", json!({"cell":"B1","input":"=Scale(A1)"}))?;
    assert_eq!(cell(&s, "B1"), Value::Number(21.));
    s.execute("cell.set", json!({"cell":"A1","input":"9"}))?;
    assert_eq!(cell(&s, "B1"), Value::Number(27.));
    s.execute("cell.set", json!({"cell":"B2","input":"=sCaLe(A1,2)+SUM(1,2)"}))?;
    assert_eq!(cell(&s, "B2"), Value::Number(21.));
    for (a, f) in [("B3", "=Hidden()"), ("B4", "=Scale()"), ("B5", "=Scale(1,2,3)"), ("B6", "=Scale(\"x\")")] {
        s.execute("cell.set", json!({"cell":a,"input":f}))?;
        assert!(matches!(cell(&s, a), Value::Error(_)), "{f}: {:?}", cell(&s, a));
    }
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn functions_return_exact_cell_types_and_reference_values() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Function Echo(ByVal x As Variant) As Variant\nEcho = x\nEnd Function\nFunction Fault() As Variant\nFault = CVErr(2042)\nEnd Function\nFunction First(ByVal x As Variant) As Variant\nFirst = x(1,1)\nEnd Function\n",
    )?;
    for (at, input, expected) in [
        ("A1", "=Echo(TRUE)", Value::Bool(true)),
        ("A2", "=Echo(\"猫\")", Value::text("猫")),
        ("A3", "=Fault()", Value::Error(CellError::NA)),
        ("A4", "=Echo(1/0)", Value::Error(CellError::Div0)),
        ("A5", "=First(B1:C2)", Value::Number(6.)),
    ] {
        s.execute("cell.set", json!({"cell":"B1","input":"6"}))?;
        s.execute("cell.set", json!({"cell":at,"input":input}))?;
        assert_eq!(cell(&s, at), expected, "{input}");
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_typed_arguments_preserve_reference_identity_and_caller_context() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Public Function RangeProbe(ByVal source As Range) As Double\nRangeProbe = source.Row + source.Column + source.Value2 + Application.Caller.Row\nEnd Function\n",
    )?;
    s.execute("cell.set", json!({"cell":"B1","input":"9"}))?;
    s.execute("cell.set", json!({"cell":"D4","input":"=RangeProbe(B1)"}))?;
    assert_eq!(cell(&s, "D4"), Value::Number(16.0));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_aware_udf_supports_address_offset_and_resize() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Public Function RangeAddress(ByVal source As Range) As String\nRangeAddress = source.Address\nEnd Function\nPublic Function OffsetValue(ByVal source As Range) As Double\nOffsetValue = source.Offset(1,1).Value2\nEnd Function\nPublic Function ResizedCount(ByVal source As Range) As Long\nResizedCount = source.Resize(2,2).Count\nEnd Function\n",
    )?;
    s.execute("cell.set", json!({"cell":"B1","input":"9"}))?;
    s.execute("cell.set", json!({"cell":"C2","input":"12"}))?;
    s.execute("cell.set", json!({"cell":"D4","input":"=RangeAddress(B1)"}))?;
    s.execute("cell.set", json!({"cell":"D5","input":"=OffsetValue(B1)"}))?;
    s.execute("cell.set", json!({"cell":"D6","input":"=ResizedCount(B1)"}))?;
    assert_eq!(cell(&s, "D4"), Value::text("$B$1"));
    assert_eq!(cell(&s, "D5"), Value::Number(12.0));
    assert_eq!(cell(&s, "D6"), Value::Number(4.0));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn range_aware_udf_preserves_typed_value_and_value2_distinction() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Public Function RangeKinds(ByVal source As Range) As Long\nRangeKinds = VarType(source.Value) * 10 + VarType(source.Value2)\nEnd Function\nPublic Function RangeRead(ByVal source As Range, ByVal typed As Boolean) As Double\nIf typed Then\nRangeRead = CDbl(source.Value)\nElse\nRangeRead = CDbl(source.Value2)\nEnd If\nEnd Function\n",
    )?;
    s.execute("cell.set", json!({"cell":"A1","input":"45000.5"}))?;
    s.execute("cell.set", json!({"cell":"A2","input":"1.23456"}))?;
    s.execute("home.numberFormat", json!({"range":"A1","format":"m/d/yyyy"}))?;
    s.execute("home.numberFormat", json!({"range":"A2","format":"$#,##0.00"}))?;
    s.execute("cell.set", json!({"cell":"C1","input":"=RangeKinds(A1)"}))?;
    s.execute("cell.set", json!({"cell":"C2","input":"=RangeKinds(A2)"}))?;
    s.execute("cell.set", json!({"cell":"D1","input":"=RangeRead(A1,TRUE)"}))?;
    s.execute("cell.set", json!({"cell":"D2","input":"=RangeRead(A1,FALSE)"}))?;
    s.execute("cell.set", json!({"cell":"E1","input":"=RangeRead(A2,TRUE)"}))?;
    s.execute("cell.set", json!({"cell":"E2","input":"=RangeRead(A2,FALSE)"}))?;
    assert_eq!(cell(&s, "C1"), Value::Number(75.0));
    assert_eq!(cell(&s, "C2"), Value::Number(65.0));
    assert_eq!(cell(&s, "D1"), Value::Number(45000.5));
    assert_eq!(cell(&s, "D2"), Value::Number(45000.5));
    assert_eq!(cell(&s, "E1"), Value::Number(1.2346));
    assert_eq!(cell(&s, "E2"), Value::Number(1.23456));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn static_udf_state_and_volatile_setting_follow_the_caller_formula() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Public Function CallerProbe(ByVal source As Range, ByVal makeVolatile As Boolean) As Long\nStatic calls As Long\ncalls = calls + 1\nApplication.Volatile makeVolatile\nCallerProbe = Application.Caller.Row + source.Row + calls\nEnd Function\n",
    )?;
    s.execute("cell.set", json!({"cell":"B1","input":"1"}))?;
    s.execute("cell.set", json!({"cell":"D4","input":"=CallerProbe(B1,TRUE)"}))?;
    assert_eq!(cell(&s, "D4"), Value::Number(6.0));
    s.execute("cell.set", json!({"cell":"E1","input":"2"}))?;
    assert_eq!(cell(&s, "D4"), Value::Number(7.0));
    s.execute("cell.set", json!({"cell":"D4","input":"=CallerProbe(B1,FALSE)"}))?;
    assert_eq!(cell(&s, "D4"), Value::Number(8.0));
    s.execute("cell.set", json!({"cell":"E1","input":"3"}))?;
    assert_eq!(cell(&s, "D4"), Value::Number(8.0));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn functions_cannot_mutate_and_work_is_bounded() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Function WriteCell() As Long\nRange(\"Z1\").Value2 = 40\nWriteCell = 8\nEnd Function\nFunction Schedule() As Long\nApplication.OnTime 1, \"DoNothing\"\nSchedule = 8\nEnd Function\nFunction Forever() As Long\nDim i As Long\nFor i = 1 To 10000000\nNext i\nForever = i\nEnd Function\n",
    )?;
    for (at, input) in [("A1", "=WriteCell()"), ("A2", "=Schedule()"), ("A3", "=Forever()")] {
        s.execute("cell.set", json!({"cell":at,"input":input}))?;
        assert_eq!(cell(&s, at), Value::Error(CellError::Value));
    }
    assert_eq!(cell(&s, "Z1"), Value::Empty);
    assert!(!s.in_group(s.doc()?.uid));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn worksheet_function_for_each_reads_value_arrays_and_recalculates() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Function TotalItems(ByVal items As Variant) As Double\nDim item As Variant,total As Double\nFor Each item In items\ntotal=total+item\nNext item\nTotalItems=total\nEnd Function\n",
    )?;
    for (address, input) in [("A1", "2"), ("B1", "3"), ("A2", "4"), ("B2", "5")] {
        s.execute("cell.set", json!({"cell":address,"input":input}))?;
    }
    s.execute("cell.set", json!({"cell":"C1","input":"=TotalItems(A1:B2)"}))?;
    assert_eq!(cell(&s, "C1"), Value::Number(14.0));
    s.execute("cell.set", json!({"cell":"B2","input":"10"}))?;
    assert_eq!(cell(&s, "C1"), Value::Number(19.0));
    assert!(!s.in_group(s.doc()?.uid));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn revoked_or_replaced_source_never_executes_old_functions() -> R {
    let (mut s, mut p, f, _bridge) = setup("Function ConstantValue() As Long\nConstantValue = 6\nEnd Function\n")?;
    let context = CustomFunctionContext { sheet: 0, caller: CellRef::new(0, 0) };
    assert_eq!(f.call(&s.doc()?.wb, "CONSTANTVALUE", context, &[]).map(|result| result.value), Some(Value::Number(6.)));
    f.allowed.store(false, Ordering::Release);
    assert_eq!(f.call(&s.doc()?.wb, "CONSTANTVALUE", context, &[]).map(|result| result.value), Some(Value::Error(CellError::Name)));
    f.allowed.store(true, Ordering::Release);
    p.modules[0].1 = p.modules[0].1.replace("= 6", "= 7");
    p.store(&mut s)?;
    assert_eq!(f.call(&s.doc()?.wb, "CONSTANTVALUE", context, &[]).map(|result| result.value), Some(Value::Error(CellError::Name)));
    s.execute("edit.undo", json!({}))?;
    assert_eq!(f.call(&s.doc()?.wb, "CONSTANTVALUE", context, &[]).map(|result| result.value), Some(Value::Number(6.)));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn discovery_filters_private_modules_parameters_and_duplicate_entries() -> R {
    let (_, p, _, _bridge) = setup(
        "Public Sub Fill()\nEnd Sub\nPrivate Sub Secret()\nEnd Sub\nSub WithArg(x As Long)\nEnd Sub\nPublic Function Twice(x As Double) As Double\nTwice=x*2\nEnd Function",
    )?;
    let entries = p.entries()?;
    assert_eq!(entries.iter().map(|e| e.name.as_str()).collect::<Vec<_>>(), ["Fill", "Twice"]);
    let mut private = p.clone();
    private.modules[0].1 = format!("Option Private Module\n{}", private.modules[0].1);
    assert!(private.entries()?.is_empty());
    let mut duplicate = p.clone();
    duplicate.modules.push(("Other.bas".into(), p.modules[0].1.clone()));
    assert!(duplicate.entries().is_err());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn source_and_assignments_round_trip_inertly_through_xlsx() -> R {
    let (mut s, mut p, _, _bridge) = setup("Sub Fill()\nRange(\"A1\").Value2 = 70\nEnd Sub")?;
    let shape = s.execute("insert.shape", json!({"text":"Fill"}))?["shape"].clone();
    s.execute("shape.setMacro", json!({"id":shape,"entry":"Fill"}))?;
    p.toolbar.push("Fill".into());
    p.store(&mut s)?;
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let p2 = Project::read(&wb)?.ok_or("project absent")?;
    assert_eq!(p2.modules, p.modules);
    assert_eq!(wb.sheet(0).unwrap().shapes[0].on_action.as_deref(), Some("Fill"));
    assert_eq!(p2.toolbar, ["Fill"]);
    assert_eq!(wb.sheet(0).unwrap().value(CellRef::new(0, 0)), Value::Empty);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn array_results_spill_and_oversized_references_fail_before_materializing() -> R {
    let (mut s, _, _, _bridge) = setup(
        "Function Pair(ByVal n As Double) As Variant\nDim a(1 To 1, 1 To 2) As Variant\na(1,1)=n\na(1,2)=n+1\nPair=a\nEnd Function\nFunction Scalar(ByVal x As Variant) As Double\nScalar=1\nEnd Function",
    )?;
    s.execute("cell.set", json!({"cell":"A1","input":"=Pair(8)"}))?;
    assert_eq!(cell(&s, "A1"), Value::Number(8.));
    assert_eq!(cell(&s, "B1"), Value::Number(9.));
    s.execute("cell.set", json!({"cell":"D1","input":"=Scalar(A1:A65537)"}))?;
    assert_eq!(cell(&s, "D1"), Value::Error(CellError::Value));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn private_names_in_other_modules_do_not_shadow_the_public_function() -> R {
    let mut s = Session::new();
    s.new_workbook();
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Private.bas", "Private Function Score() As Long\nScore=1\nEnd Function")?;
    ide.add_module("Public.bas", "Public Function Score() As Long\nScore=8\nEnd Function")?;
    let p = Project::from_ide(&ide, None)?;
    p.store(&mut s)?;
    let f = Functions::compile(&p, &b, true)?;
    assert_eq!(
        f.call(&s.doc()?.wb, "SCORE", CustomFunctionContext { sheet: 0, caller: CellRef::new(0, 0) }, &[]).map(|result| result.value),
        Some(Value::Number(8.))
    );
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn default_byref_parameters_use_owned_values_without_writing_precedents() -> R {
    let (mut s, _, _, _bridge) = setup("Function Bump(n As Double) As Double\nn=n+1\nBump=n\nEnd Function")?;
    s.execute("cell.set", json!({"cell":"A1","input":"7"}))?;
    s.execute("cell.set", json!({"cell":"B1","input":"=Bump(A1)"}))?;
    assert_eq!(cell(&s, "A1"), Value::Number(7.));
    assert_eq!(cell(&s, "B1"), Value::Number(8.));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn repeated_shared_strings_exceeding_32bit_payloads_are_refused_before_vba_allocation() -> R {
    let (mut s, _, f, _bridge) = setup("Function Scalar(ByVal x As Variant) As Long\nScalar=1\nEnd Function")?;
    // Each cell fits Excel's text limit; Arc sharing represents >4 GiB of logical UTF-8 data cheaply.
    let text = Value::text("é".repeat(32_767));
    let array = gridcraft_engine::core::Array::new(65_536, 1, vec![text.clone(); 65_536]).ok_or("array")?;
    assert_eq!(
        f.call(
            &s.doc()?.wb,
            "SCALAR",
            CustomFunctionContext { sheet: 0, caller: CellRef::new(0, 0) },
            &[gridcraft_engine::calc::CustomFunctionArgument { value: Value::from(array), references: vec![] }]
        )
        .map(|result| result.value),
        Some(Value::Error(CellError::Value))
    );
    let doc = s.doc_mut()?;
    let wb = Arc::make_mut(&mut doc.wb);
    let sh = wb.sheet_mut(0).ok_or("sheet")?;
    for row in 0..65_536 {
        sh.set_value(CellRef::new(row, 0), text.clone());
    }
    s.execute("cell.set", json!({"cell":"B1","input":"=Scalar(A1:A65536)"}))?;
    assert_eq!(cell(&s, "B1"), Value::Error(CellError::Value));
    Ok(())
}
