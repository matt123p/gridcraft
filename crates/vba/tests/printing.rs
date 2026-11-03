//! OM10 native print setup, pagination and PDF export acceptance.
use gridcraft_engine::Session;
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Print.bas", source)?;
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
fn typed_page_setup_targets_native_settings_and_roundtrips_in_xlsx() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim p As PageSetup
Set p=Worksheets("Other").PageSetup
If Not (p.Parent Is Worksheets("Other")) Then Err.Raise 5
If Not (p.Application Is Application) Then Err.Raise 5
p.Orientation=xlLandscape
p.PaperSize=xlPaperA4
p.LeftMargin=36
p.RightMargin=40
p.TopMargin=42
p.BottomMargin=44
p.HeaderMargin=18
p.FooterMargin=20
p.PrintArea="'Other'!$A$1:$F$80"
p.PrintTitleRows="$1:$2"
p.PrintTitleColumns="$A:$B"
p.PrintGridlines=True
p.PrintHeadings=True
p.CenterHorizontally=True
p.CenterVertically=True
p.Zoom=False
p.FitToPagesWide=1
p.FitToPagesTall=False
If p.Orientation <> xlLandscape Or p.PaperSize <> xlPaperA4 Then Err.Raise 5
If Abs(p.LeftMargin-36)>0.0001 Or Abs(p.FooterMargin-20)>0.0001 Then Err.Raise 5
If p.PrintArea <> "$A$1:$F$80" Then Err.Raise 5
If p.PrintTitleRows <> "$1:$2" Or p.PrintTitleColumns <> "$A:$B" Then Err.Raise 5
If p.Zoom <> False Or p.FitToPagesWide <> 1 Or p.FitToPagesTall <> False Then Err.Raise 5
If p.Order <> xlDownThenOver Then Err.Raise 5
If Not p.PrintGridlines Or Not p.PrintHeadings Or Not p.CenterHorizontally Or Not p.CenterVertically Then Err.Raise 5
If ActiveSheet.Name <> "Sheet1" Then Err.Raise 5
End Sub"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    finish(&mut s, &mut b, &mut ide)?;
    let p = s.doc()?.wb.sheet(1).ok_or("sheet")?.print.clone();
    assert_eq!(p.fit_to, Some((1, 0)));
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert_eq!(wb.sheet(1).ok_or("sheet")?.print, p);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn header_footer_definitions_and_numbered_pages_use_native_variants() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim p As PageSetup, first As Page, header As HeaderFooter, pages As Pages
Set p=ActiveSheet.PageSetup
p.PrintArea="A1:C120"
p.LeftHeader="Native"
p.CenterHeader="Page &P of &N"
p.RightFooter="&A"
p.DifferentFirstPageHeaderFooter=True
p.OddAndEvenPagesHeaderFooter=True
p.ScaleWithDocHeaderFooter=False
p.AlignMarginsHeaderFooter=False
Set first=p.FirstPage
Set header=first.CenterHeader
header.Text="Cover"
p.EvenPage.CenterHeader.Text="Even &P"
If header.Text <> "Cover" Then Err.Raise 5
If p.CenterHeader <> "Page &P of &N" Then Err.Raise 5
If p.EvenPage.CenterHeader.Text <> "Even &P" Then Err.Raise 5
If Not p.DifferentFirstPageHeaderFooter Or Not p.OddAndEvenPagesHeaderFooter Then Err.Raise 5
If p.ScaleWithDocHeaderFooter Or p.AlignMarginsHeaderFooter Then Err.Raise 5
Set pages=p.Pages
If pages.Count < 2 Then Err.Raise 5
If pages(1).CenterHeader.Text <> "Cover" Then Err.Raise 5
If pages(2).CenterHeader.Text <> "Even &P" Then Err.Raise 5
On Error Resume Next
pages(1).CenterHeader.Text="No per-page state"
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
On Error GoTo 0
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    let json = s.execute("file.exportPdf", json!({}))?;
    let bytes = gridcraft_engine::io::base64_decode(json["base64"].as_str().ok_or("pdf")?).ok_or("invalid base64")?;
    let pdf = String::from_utf8_lossy(&bytes);
    assert!(pdf.contains("(Cover) Tj"));
    assert!(pdf.contains("(Even 2) Tj"));
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn page_setup_invalid_options_refuse_before_native_edits() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim p As PageSetup
Set p=ActiveSheet.PageSetup
On Error Resume Next
p.LeftMargin=-1
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
p.LeftMargin=2000
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
p.Zoom=401
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
p.FitToPagesWide=2
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
p.PrintArea="A1,B2"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
p.PrintTitleRows="A1:B2"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
p.CenterHeader=String(256,"X")
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If p.Zoom <> 100 Or p.PrintArea <> "" Or p.CenterHeader <> "" Then Err.Raise 5
End Sub"#,
    )?;
    let before = s.doc()?.wb.sheet(0).ok_or("sheet")?.print.clone();
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.print, before);
    assert!(!s.doc()?.is_dirty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn manual_page_breaks_are_command_backed_and_stale_snapshots_refuse() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim h As HPageBreak, v As VPageBreak, x As Long
Set h=ActiveSheet.HPageBreaks.Add(Range("B20"))
If h.Location.Address <> "$A$20" Or h.Type <> xlPageBreakManual Then Err.Raise 5
If h.Extent <> xlPageBreakFull Then Err.Raise 5
If Not (h.Parent Is ActiveSheet) Or Not (h.Application Is Application) Then Err.Raise 5
Set v=ActiveSheet.VPageBreaks.Add(Range("D3"))
If v.Location.Address <> "$D$1" Then Err.Raise 5
On Error Resume Next
x=h.Type
If Err.Number <> 91 Then Err.Raise 5
Err.Clear
On Error GoTo 0
Set h=ActiveSheet.HPageBreaks(1)
Set h.Location=Range("A25")
If Rows(25).PageBreak <> xlPageBreakManual Then Err.Raise 5
Columns("F").PageBreak=xlPageBreakManual
If Columns("F").PageBreak <> xlPageBreakManual Then Err.Raise 5
ActiveSheet.VPageBreaks(1).Delete
Rows(25).PageBreak=xlPageBreakNone
ActiveSheet.ResetAllPageBreaks
If ActiveSheet.HPageBreaks.Count <> 0 Or ActiveSheet.VPageBreaks.Count <> 0 Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.print.row_breaks.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn automatic_breaks_are_derived_and_cannot_be_removed_as_manual_markers() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim h As HPageBreak
ActiveSheet.PageSetup.PrintArea="A1:C120"
If ActiveSheet.HPageBreaks.Count < 1 Then Err.Raise 5
For Each h In ActiveSheet.HPageBreaks
If h.Type <> xlPageBreakAutomatic Then Err.Raise 5
Next
Set h=ActiveSheet.HPageBreaks(1)
On Error Resume Next
h.Delete
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If h.Type <> xlPageBreakAutomatic Then Err.Raise 5
h.Type=xlPageBreakManual
If ActiveSheet.HPageBreaks(1).Type <> xlPageBreakManual Then Err.Raise 5
Cells.PageBreak=xlPageBreakNone
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unavailable_print_services_and_export_options_have_no_effects() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
On Error Resume Next
ActiveSheet.PrintOut Preview:=True
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
ThisWorkbook.PrintPreview
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
Application.ActivePrinter="No such printer"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
ThisWorkbook.ExportAsFixedFormat xlTypePDF
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
ActiveSheet.ExportAsFixedFormat Type:=xlTypeXPS,Filename:="unavailable.xps"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF,Filename:="unavailable.pdf",OpenAfterPublish:=True
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF,Filename:="unavailable.pdf",From:=1
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF,Filename:="unavailable.pdf",IncludeDocProperties:=False
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(!s.doc()?.is_dirty());
    assert!(s.ui_requests.is_empty());
    Ok(())
}
#[cfg(not(target_arch = "wasm32"))]
fn path(label: &str) -> String {
    std::env::temp_dir().join(format!("gridcraft-om10-{}-{label}.pdf", std::process::id())).to_string_lossy().replace('"', "\"\"")
}
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn pdf_export_renders_receiver_scope_without_focus_or_workbook_file_changes() -> R {
    let sheet = path("inspection");
    let workbook = path("workbook");
    let range = path("range");
    let source = format!(
        r#"Sub Main()
Dim p As PageSetup
Set p=Worksheets("Report").PageSetup
p.PaperSize=xlPaperA4
p.Orientation=xlLandscape
p.Zoom=False
p.FitToPagesWide=1
p.FitToPagesTall=1
p.CenterHeader="Native PDF report"
p.CenterFooter="Page &P of &N"
p.PrintGridlines=True
p.PrintArea="A1:C8"
Worksheets("Report").ExportAsFixedFormat Type:=xlTypePDF,Filename:="{sheet}"
ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF,Filename:="{workbook}"
Worksheets("Report").Range("B2:C8").ExportAsFixedFormat Type:=xlTypePDF,Filename:="{range}",IgnorePrintAreas:=True
If ActiveSheet.Name <> "Sheet1" Then Err.Raise 5
End Sub"#
    );
    let (mut s, mut b, mut ide) = setup(&source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"Other sheet marker"}))?;
    s.execute("home.insertSheet", json!({"name":"Report"}))?;
    s.execute("range.setValues", json!({"range":"A1:C3","values":[["Report marker","Item","Amount"],["Excluded","East",123],["","West",456]]}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    let original_path = s.doc()?.path.clone();
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.doc()?.path, original_path);
    let read = |p: &str| -> Result<String, Box<dyn std::error::Error>> { Ok(String::from_utf8_lossy(&std::fs::read(p)?).into_owned()) };
    let pdf = read(&sheet)?;
    assert!(pdf.starts_with("%PDF-1.7"));
    assert!(pdf.contains("/MediaBox [0 0 841.89 595.28]"));
    assert!(pdf.contains("(Native PDF report) Tj"));
    assert!(pdf.contains("(Report marker) Tj"));
    assert!(!pdf.contains("(Other sheet marker) Tj"));
    let all = read(&workbook)?;
    assert!(all.contains("(Other sheet marker) Tj") && all.contains("(Report marker) Tj"));
    let cropped = read(&range)?;
    assert!(!cropped.contains("(Report marker) Tj"));
    assert!(cropped.contains("(East) Tj"));
    Ok(())
}
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn pdf_area_override_and_io_failures_preserve_native_state() -> R {
    let keep = path("keep-area");
    let ignore = path("ignore-area");
    let missing = std::env::temp_dir().join(format!("gridcraft-om10-missing-{}", std::process::id())).join("out.pdf").to_string_lossy().into_owned();
    let source = format!(
        r#"Sub Main()
ActiveSheet.PageSetup.PrintArea="A1"
ActiveSheet.ExportAsFixedFormat xlTypePDF,"{keep}"
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF,Filename:="{ignore}",IgnorePrintAreas:=True
On Error Resume Next
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF,Filename:="{missing}"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If ActiveSheet.PageSetup.PrintArea <> "$A$1" Then Err.Raise 5
End Sub"#
    );
    let (mut s, mut b, mut ide) = setup(&source)?;
    s.execute("cell.set", json!({"cell":"A1","input":"Kept"}))?;
    s.execute("cell.set", json!({"cell":"E5","input":"Outside marker"}))?;
    finish(&mut s, &mut b, &mut ide)?;
    let a = std::fs::read(keep)?;
    let z = std::fs::read(ignore)?;
    assert!(!String::from_utf8_lossy(&a).contains("(Outside marker) Tj"));
    assert!(String::from_utf8_lossy(&z).contains("(Outside marker) Tj"));
    assert!(!std::path::Path::new(&missing).exists());
    Ok(())
}
#[cfg(all(target_arch = "wasm32", target_os = "unknown"))]
#[wasm_bindgen_test::wasm_bindgen_test]
fn browser_export_requires_a_file_service_without_claiming_a_successful_download() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
On Error Resume Next
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF,Filename:="browser.pdf"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(!s.doc()?.is_dirty());
    assert!(s.ui_requests.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn numbered_page_views_expire_but_native_header_definitions_remain_live() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim p As PageSetup, rendered As HeaderFooter, definition As HeaderFooter, x As String
Set p=ActiveSheet.PageSetup
Set definition=p.FirstPage.CenterHeader
Set rendered=p.Pages(1).CenterHeader
p.Zoom=125
On Error Resume Next
x=rendered.Text
If Err.Number <> 91 Then Err.Raise 5
Err.Clear
On Error GoTo 0
definition.Text="Live first header"
If p.FirstPage.CenterHeader.Text <> "Live first header" Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn malformed_native_print_bounds_refuse_instead_of_overflowing() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim x As Variant
On Error Resume Next
x=ActiveSheet.PageSetup.PrintTitleRows
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
x=ActiveSheet.HPageBreaks.Count
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
Rows(2).PageBreak=xlPageBreakManual
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
End Sub"#,
    )?;
    let sh = std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?;
    sh.print.title_rows = Some((u32::MAX, u32::MAX));
    sh.print.row_breaks = vec![u32::MAX];
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn print_receivers_keep_consent_and_sheet_lifetimes() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim p As PageSetup, h As HPageBreak, x As Variant
Set p=Worksheets("Other").PageSetup
Set h=Worksheets("Other").HPageBreaks.Add(Worksheets("Other").Range("A25"))
Worksheets("Other").Delete
On Error Resume Next
x=p.Orientation
If Err.Number <> 91 Then Err.Raise 5
Err.Clear
Set x=h.Location
If Err.Number <> 91 Then Err.Raise 5
Err.Clear
Workbooks(2).ExportAsFixedFormat Type:=xlTypePDF,Filename:="untrusted.pdf"
If Err.Number <> 70 Then Err.Raise 5
Err.Clear
On Error GoTo 0
End Sub"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    let first = s.active_window().ok_or("window")?;
    s.new_workbook();
    s.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
    s.activate_window(first);
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn malformed_print_areas_are_validated_before_direct_layout_queries() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim n As Long
On Error Resume Next
n=ActiveSheet.PageSetup.Pages.Count
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
n=ActiveSheet.HPageBreaks.Count
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
End Sub"#,
    )?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.print.print_area =
        Some(gridcraft_engine::core::RangeRef::cell(gridcraft_engine::core::CellRef::new(u32::MAX, 0)));
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn owned_print_requests_suspend_once_and_propagate_real_submission_errors() -> R {
    use gridcraft_vba::{PrintOperation as Op, PrintResponse, PrintServices};
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Application.ActivePrinter="Office queue"
If Application.ActivePrinter <> "Office queue" Then Err.Raise 5
Range("A1").Value2="Printed marker"
On Error Resume Next
ActiveSheet.PrintOut Copies:=2,ActivePrinter:="Office queue"
Range("B1").Value2=Err.Number
Err.Clear
On Error GoTo 0
End Sub"#,
    )?;
    b.set_print_services(PrintServices { native: true, printer: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut requests = 0;
    for _ in 0..10000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(req) = b.take_print_request() {
            requests += 1;
            assert!(b.print_request_permitted(&s, req.token));
            let reply = match req.kind {
                Op::SetPrinter(name) => {
                    assert_eq!(name, "Office queue");
                    Ok(PrintResponse::Done)
                }
                Op::GetPrinter => Ok(PrintResponse::Printer("Office queue".into())),
                Op::Submit { pdf, copies, printer } => {
                    assert_eq!(copies, 2);
                    assert_eq!(printer.as_deref(), Some("Office queue"));
                    assert!(String::from_utf8_lossy(&pdf).contains("(Printed marker) Tj"));
                    for _ in 0..3 {
                        b.tick(&mut s, &mut ide)?;
                        assert_eq!(ide.mode(), Mode::Waiting);
                        assert!(b.take_print_request().is_none());
                    }
                    Err("printer is offline".into())
                }
                _ => return Err("unexpected print request".into()),
            };
            b.respond_to_print(req.token, reply).map_err(|e| e.description)?;
        }
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(requests, 3);
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::new(0, 1)).display(), "1004");
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn preview_waits_for_service_completion_and_rechecks_consent() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
ActiveSheet.PrintPreview
Range("A1").Value2="After preview"
End Sub"#,
    )?;
    b.set_print_services(gridcraft_vba::PrintServices { preview: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut request = None;
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(r) = b.take_print_request() {
            request = Some(r);
            break;
        }
    }
    let req = request.ok_or("preview request")?;
    assert!(matches!(req.kind, gridcraft_vba::PrintOperation::Preview { sheet: Some(_), .. }));
    for _ in 0..3 {
        b.tick(&mut s, &mut ide)?;
        assert_eq!(ide.mode(), Mode::Waiting);
    }
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.cell(gridcraft_engine::core::CellRef::new(0, 0)).is_none());
    s.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
    assert!(!b.print_request_permitted(&s, req.token));
    b.respond_to_print(req.token, Ok(gridcraft_vba::PrintResponse::Done)).map_err(|e| e.description)?;
    for _ in 0..10 {
        let _ = b.tick(&mut s, &mut ide);
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_ne!(ide.mode(), Mode::Completed);
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.cell(gridcraft_engine::core::CellRef::new(0, 0)).is_none());
    Ok(())
}
#[cfg(all(target_arch = "wasm32", target_os = "unknown"))]
#[wasm_bindgen_test::wasm_bindgen_test]
fn connected_browser_download_receives_owned_native_pdf_once() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF,Filename:="browser.pdf"
End Sub"#,
    )?;
    b.set_print_services(gridcraft_vba::PrintServices { download: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut downloads = 0;
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(req) = b.take_print_request() {
            let gridcraft_vba::PrintOperation::Download { pdf, filename } = req.kind else {
                return Err("download request".into());
            };
            assert_eq!(filename, "browser.pdf");
            assert!(pdf.starts_with(b"%PDF"));
            downloads += 1;
            b.respond_to_print(req.token, Ok(gridcraft_vba::PrintResponse::Done)).map_err(|e| e.description)?;
        }
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(downloads, 1);
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert!(!s.doc()?.is_dirty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn window_printout_uses_the_real_selected_sheet_group() -> R {
    let (mut s, mut b, mut ide) = setup("Sub Main()\nActiveWindow.PrintOut\nEnd Sub")?;
    s.execute("cell.set", json!({"cell":"A1","input":"Selected one"}))?;
    s.execute("home.insertSheet", json!({"name":"Second"}))?;
    s.execute("cell.set", json!({"cell":"A1","input":"Selected two"}))?;
    s.execute("home.insertSheet", json!({"name":"Excluded"}))?;
    s.execute("cell.set", json!({"cell":"A1","input":"Unselected marker"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    s.execute("sheet.select", json!({"sheet":1,"replace":false}))?;
    b.set_print_services(gridcraft_vba::PrintServices { native: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut jobs = 0;
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(req) = b.take_print_request() {
            let gridcraft_vba::PrintOperation::Submit { pdf, .. } = req.kind else {
                return Err("print request".into());
            };
            let data = String::from_utf8_lossy(&pdf);
            assert!(data.contains("(Selected one) Tj"));
            assert!(data.contains("(Selected two) Tj"));
            assert!(!data.contains("(Unselected marker) Tj"));
            jobs += 1;
            b.respond_to_print(req.token, Ok(gridcraft_vba::PrintResponse::Done)).map_err(|e| e.description)?;
        }
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(jobs, 1);
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn reset_cancels_unsubmitted_print_requests_without_replay() -> R {
    let (mut s, mut b, mut ide) = setup("Sub Main()\nActiveSheet.PrintOut\nRange(\"A1\").Value2=99\nEnd Sub")?;
    b.set_print_services(gridcraft_vba::PrintServices { native: true, ..Default::default() });
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    let mut request = None;
    for _ in 0..1000 {
        b.tick(&mut s, &mut ide)?;
        if let Some(req) = b.take_print_request() {
            request = Some(req);
            break;
        }
    }
    let request = request.ok_or("print request")?;
    b.dispatch(&mut s, &mut ide, Action::Reset)?;
    assert!(!b.print_request_pending(request.token));
    assert!(b.respond_to_print(request.token, Ok(gridcraft_vba::PrintResponse::Done)).is_err());
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.cell(gridcraft_engine::core::CellRef::new(0, 0)).is_none());
    Ok(())
}
