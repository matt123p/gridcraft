//! OM20 metadata regressions: the declared profile is distinct from full Excel parity.
use gridcraft_engine::Session;
use gridcraft_vba::{Action, Bridge, Ide, Mode};
type R = Result<(), Box<dyn std::error::Error>>;
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn reference_metadata_uses_the_live_nominal_set() -> R {
    let metadata = gridcraft_vba::reference_metadata()?;
    let libraries = metadata.as_array().ok_or("metadata")?;
    for (library, expected) in [("Excel", "Worksheet"), ("Excel", "CellControl"), ("Office", "DocumentProperties"), ("Office", "TextRange2")] {
        let entry = libraries.iter().find(|l| l["library"] == library).ok_or("library")?;
        assert_eq!(entry["available"], true);
        assert!(entry["types"].as_array().ok_or("types")?.iter().any(|t| t["name"] == expected && t["kind"] == "object"));
    }
    let replace = gridcraft_vba::member_descriptors().into_iter().find(|d| d.owner == "Range" && d.name == "Replace").ok_or("Replace descriptor")?;
    assert!(replace.signature.contains("FormulaVersion As Variant"));
    assert!(!replace.signature.contains("XlReplaceFormula"));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn forecast_metadata_and_replace_variant_binding_are_usable() -> R {
    let mut s = Session::new();
    s.new_workbook();
    let mut b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module(
        "Profile.bas",
        r#"Option Explicit
Sub Main()
Dim aggregation As XlForecastAggregation
Dim chartType As XlForecastChartType
Dim completion As XlForecastDataCompletion
Dim actual As Long, ignored As Boolean
aggregation=xlForecastAggregationMin
chartType=xlForecastChartTypeColumn
completion=xlForecastDataCompletionZeros
If aggregation<>6 Or chartType<>1 Or completion<>0 Then Err.Raise 5
If xlForecastAggregationSum<>7 Or xlForecastDataCompletionInterpolate<>1 Then Err.Raise 5
On Error Resume Next
ignored=Range("A1").Replace("x","y",FormulaVersion:=1)
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
End Sub"#,
    )?;
    b.configure(&mut ide)?;
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
