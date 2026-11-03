//! Existing PDF export service. Printer, preview UI and post-export viewer need host services.
use super::features::reads;
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure};
use serde_json::json;
#[derive(Clone, Copy, Debug, Default)]
pub struct PrintServices {
    pub native: bool,
    pub browser: bool,
    pub download: bool,
    pub preview: bool,
    pub printer: bool,
}
pub struct PrintRequest {
    pub token: u64,
    pub doc: u64,
    pub kind: PrintOperation,
}
pub enum PrintOperation {
    Submit { pdf: Vec<u8>, copies: u32, printer: Option<String> },
    Download { pdf: Vec<u8>, filename: String },
    Preview { window: u64, sheet: Option<gridcraft_engine::model::sheet::SheetIdentity> },
    ExportWithPicker { params: serde_json::Value, suggested: String, sheet: Option<gridcraft_engine::model::sheet::SheetIdentity>, range_epoch: u64 },
    GetPrinter,
    SetPrinter(String),
}
pub enum PrintResponse {
    Done,
    Printer(String),
}
fn pdf_bytes(s: &mut Session, p: serde_json::Value) -> HostResult<Vec<u8>> {
    let data = s.execute("file.exportPdf", p).map_err(|e| failure(1004, e.to_string()))?;
    let encoded = data["base64"].as_str().ok_or_else(|| failure(1004, "native export did not return PDF bytes"))?;
    if encoded.len() > 90_000_000 {
        return Err(failure(7, "print service payload budget"));
    }
    let bytes = gridcraft_engine::io::base64_decode(encoded).ok_or_else(|| failure(1004, "invalid PDF payload"))?;
    if bytes.len() > 64 * 1024 * 1024 {
        return Err(failure(7, "print service payload budget"));
    }
    Ok(bytes)
}
impl Bridge {
    pub(crate) fn print_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        let m = c.member.to_ascii_lowercase();
        if *o == Object::Application && matches!(m.as_str(), "activeprinter" | "printcommunication") {
            return Some(Err(failure(1004, "printer selection and print batching have no connected host service/state")));
        }
        let print_owner = matches!(o, Object::Workbook | Object::Sheet(_) | Object::Range(..) | Object::Worksheets | Object::Sheets)
            || matches!(o,Object::Feature(f) if matches!(f.kind,"Window"|"WindowSheets"|"Chart"));
        if matches!(o,Object::Feature(f) if f.kind=="Chart") && matches!(m.as_str(), "printout" | "printpreview") {
            return Some(Err(failure(438, "independent chart print/page layout has no native backend")));
        }
        if print_owner && matches!(m.as_str(), "printout" | "printpreview") {
            return Some((|| {
                if c.kind != InvokeKind::Call {
                    return Err(failure(438, "printing requires a method call"));
                }
                arguments(
                    c,
                    if m == "printpreview" {
                        &["EnableChanges"]
                    } else {
                        &["From", "To", "Copies", "Preview", "ActivePrinter", "PrintToFile", "Collate", "PrToFileName", "IgnorePrintAreas"]
                    },
                )?;
                Err(failure(1004, "native printer/preview presentation service is not connected; no job was submitted"))
            })());
        }
        if matches!(o, Object::Workbook | Object::Sheet(_) | Object::Range(..)) && m == "exportasfixedformat" {
            return Some(self.export_fixed(s, o, c));
        }
        None
    }
    fn export_fixed(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> HostResult<Value> {
        let p = self.export_params(s, o, c, false)?;
        if cfg!(target_arch = "wasm32") {
            return Err(failure(1004, "browser file/download service is not connected"));
        }
        s.execute("file.exportPdf", p).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    fn printer_request(&self, c: &HostCall) -> HostResult<PrintOperation> {
        if reads(c) {
            arguments(c, &[])?;
            return Ok(PrintOperation::GetPrinter);
        }
        if c.kind != InvokeKind::Let {
            return Err(failure(438, "ActivePrinter requires Get or Let"));
        }
        let a = arguments(c, &["Value"])?;
        let name = string(required(&a, 0)?)?;
        if name.is_empty() || name.len() > 1024 {
            return Err(failure(1004, "invalid printer name"));
        }
        Ok(PrintOperation::SetPrinter(name))
    }
    fn export_request(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> HostResult<Option<PrintOperation>> {
        let services = self.operations.print_services;
        let m = c.member.to_ascii_lowercase();
        if matches!(o, Object::Workbook | Object::Sheet(_) | Object::Range(..))
            && m == "exportasfixedformat"
            && !cfg!(target_arch = "wasm32")
            && self.operations.services.save
        {
            let a = arguments(
                c,
                &[
                    "Type",
                    "Filename",
                    "Quality",
                    "IncludeDocProperties",
                    "IgnorePrintAreas",
                    "From",
                    "To",
                    "OpenAfterPublish",
                    "FixedFormatExtClassPtr",
                ],
            )?;
            if matches!(a[1], Value::Missing) {
                let mut params = self.export_params(s, o, c, true)?;
                let suggested = params["path"].as_str().ok_or_else(|| failure(1004, "PDF filename"))?.to_owned();
                if let Some(map) = params.as_object_mut() {
                    map.remove("path");
                }
                let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
                let sheet = match o {
                    Object::Sheet(si) | Object::Range(si, _) => {
                        Some(d.wb.sheet(*si).ok_or_else(|| failure(91, "worksheet deleted"))?.identity.clone())
                    }
                    _ => None,
                };
                return Ok(Some(PrintOperation::ExportWithPicker { params, suggested, sheet, range_epoch: d.range_epoch }));
            }
        }
        if matches!(o, Object::Workbook | Object::Sheet(_) | Object::Range(..))
            && m == "exportasfixedformat"
            && cfg!(target_arch = "wasm32")
            && services.download
        {
            let mut p = self.export_params(s, o, c, true)?;
            let filename = p["path"].as_str().ok_or_else(|| failure(1004, "Filename required"))?.to_owned();
            if let Some(map) = p.as_object_mut() {
                map.remove("path");
            }
            return Ok(Some(PrintOperation::Download { pdf: pdf_bytes(s, p)?, filename }));
        }
        Ok(None)
    }
    pub(crate) fn defer_print_call(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> HostResult<Option<PrintOperation>> {
        let services = self.operations.print_services;
        let m = c.member.to_ascii_lowercase();
        if *o == Object::Application && m == "activeprinter" && services.printer {
            return self.printer_request(c).map(Some);
        }

        if m == "exportasfixedformat" {
            return self.export_request(s, o, c);
        }
        let printable = matches!(o, Object::Workbook | Object::Sheet(_) | Object::Range(..) | Object::Worksheets | Object::Sheets)
            || matches!(o,Object::Feature(f) if matches!(f.kind,"Window"|"WindowSheets"));
        if !printable || !matches!(m.as_str(), "printout" | "printpreview") {
            return Ok(None);
        }
        if c.kind != InvokeKind::Call {
            return Err(failure(438, "printing requires a method call"));
        }
        let mut preview = m == "printpreview";
        let mut copies = 1;
        let mut printer = None;
        let mut ignore = false;
        if preview {
            let a = arguments(c, &["EnableChanges"])?;
            if !matches!(a[0], Value::Missing) && !super::address::boolean(&a[0])? {
                return Err(failure(1004, "native preview has no read-only controls mode"));
            }
        } else {
            let a =
                arguments(c, &["From", "To", "Copies", "Preview", "ActivePrinter", "PrintToFile", "Collate", "PrToFileName", "IgnorePrintAreas"])?;
            if [0, 1, 6, 7].iter().any(|i| !matches!(a[*i], Value::Missing)) || !matches!(a[5], Value::Missing) && super::address::boolean(&a[5])? {
                return Err(failure(1004, "native print service does not expose page intervals, collation or print-to-file"));
            }
            if !matches!(a[2], Value::Missing) {
                let n = numeric_index(&a[2])?;
                if !(1..=999).contains(&n) {
                    return Err(failure(1004, "Copies must be 1..999"));
                }
                copies = n as u32;
            }
            if !matches!(a[3], Value::Missing) {
                preview = super::address::boolean(&a[3])?;
            }
            if !matches!(a[4], Value::Missing) {
                let name = string(&a[4])?;
                if name.is_empty() || name.len() > 1024 {
                    return Err(failure(1004, "invalid printer name"));
                }
                printer = Some(name);
            }
            if !matches!(a[8], Value::Missing) {
                ignore = super::address::boolean(&a[8])?;
            }
        }
        if preview {
            if copies != 1 || printer.is_some() {
                return Err(failure(1004, "preview cannot represent per-job copies/printer overrides"));
            }
            if !services.preview {
                return Ok(None);
            }
            if matches!(o, Object::Range(..)) || ignore {
                return Err(failure(1004, "native preview UI cannot represent range/ignore-area scope"));
            }
            let si = match o {
                Object::Sheet(si) => Some(*si),
                Object::Feature(f) if matches!(f.kind, "Window" | "WindowSheets") => {
                    let selected = &s.doc().map_err(|e| failure(91, e.to_string()))?.selected_sheets;
                    if selected.len() != 1 {
                        return Err(failure(1004, "native preview UI has no selected-group scope"));
                    }
                    selected.first().copied()
                }
                _ => None,
            };
            let sheet = si
                .map(|si| {
                    s.doc()
                        .map_err(|e| failure(91, e.to_string()))?
                        .wb
                        .sheet(si)
                        .map(|sh| sh.identity.clone())
                        .ok_or_else(|| failure(91, "sheet deleted"))
                })
                .transpose()?;
            return Ok(Some(PrintOperation::Preview { window: s.active_window().ok_or_else(|| failure(91, "window closed"))?, sheet }));
        }
        if !services.native && !services.browser {
            return Ok(None);
        }
        if !services.native && (copies != 1 || printer.is_some()) {
            return Err(failure(1004, "browser print flow cannot configure copies or printer"));
        }
        let mut p = json!({"ignorePrintAreas":ignore});
        match o {
            Object::Workbook | Object::Worksheets | Object::Sheets => p["sheets"] = json!("all"),
            Object::Feature(f) if matches!(f.kind, "Window" | "WindowSheets") => {
                p["sheetIndices"] = json!(s.doc().map_err(|e| failure(91, e.to_string()))?.selected_sheets)
            }
            Object::Sheet(si) => p["sheet"] = json!(si),
            Object::Range(si, r) => {
                let mut rect = r.rectangle()?;
                if !ignore && let Some(area) = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(*si).and_then(|sh| sh.print.print_area) {
                    rect = rect.intersection(&area).ok_or_else(|| failure(1004, "print range does not intersect print area"))?;
                }
                p["sheet"] = json!(si);
                p["range"] = json!(rect.a1());
            }
            _ => return Ok(None),
        }
        Ok(Some(PrintOperation::Submit { pdf: pdf_bytes(s, p)?, copies, printer }))
    }

    fn export_params(&self, s: &Session, o: &Object, c: &HostCall, allow_missing: bool) -> HostResult<serde_json::Value> {
        if !reads(c) {
            return Err(failure(438, "export requires a method invocation"));
        }
        let a = arguments(
            c,
            &["Type", "Filename", "Quality", "IncludeDocProperties", "IgnorePrintAreas", "From", "To", "OpenAfterPublish", "FixedFormatExtClassPtr"],
        )?;
        if numeric_index(required(&a, 0)?)? != 0 {
            return Err(failure(1004, "only the native PDF exporter is available"));
        }
        if !matches!(a[2], Value::Missing) && numeric_index(&a[2])? != 0 {
            return Err(failure(1004, "native PDF has no alternative quality mode"));
        }
        if !matches!(a[3], Value::Missing) && !super::address::boolean(&a[3])? {
            return Err(failure(1004, "native PDF metadata suppression is unavailable"));
        }
        if !matches!(a[5], Value::Missing) || !matches!(a[6], Value::Missing) {
            return Err(failure(1004, "native PDF service has no page-interval export option"));
        }
        if !matches!(a[7], Value::Missing) && super::address::boolean(&a[7])? {
            return Err(failure(1004, "post-export viewer service is not connected"));
        }
        if !matches!(a[8], Value::Missing) {
            return Err(failure(1004, "external fixed-format class pointers are unavailable"));
        }
        let path = if matches!(a[1], Value::Missing) && allow_missing {
            std::path::Path::new(&s.doc().map_err(|e| failure(91, e.to_string()))?.display_title())
                .with_extension("pdf")
                .to_string_lossy()
                .into_owned()
        } else {
            string(required(&a, 1).map_err(|_| failure(1004, "export requires Filename; no picker service is connected"))?)?
        };
        if path.is_empty() || path.len() > 4096 || path.contains('\0') {
            return Err(failure(1004, "invalid export Filename"));
        }
        let ignore = !matches!(a[4], Value::Missing) && super::address::boolean(&a[4])?;
        let mut p = json!({"path":path,"ignorePrintAreas":ignore});
        match o {
            Object::Workbook => p["sheets"] = json!("all"),
            Object::Sheet(si) => p["sheet"] = json!(si),
            Object::Range(si, r) => {
                let mut rect = r.rectangle()?;
                if !ignore && let Some(area) = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(*si).and_then(|sh| sh.print.print_area) {
                    rect = rect.intersection(&area).ok_or_else(|| failure(1004, "export range does not intersect the print area"))?;
                }
                p["sheet"] = json!(si);
                p["range"] = json!(rect.a1());
            }
            _ => return Err(failure(438, "unsupported export receiver")),
        }
        Ok(p)
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    let mut out = Vec::new();
    for owner in ["Workbook", "Worksheet", "Range"] {
        out.push(super::MemberDescriptor{owner,name:"ExportAsFixedFormat",signature:"Sub ExportAsFixedFormat(Type As XlFixedFormatType, Optional Filename As Variant, Optional Quality As Variant, Optional IncludeDocProperties As Variant, Optional IgnorePrintAreas As Variant, Optional From As Variant, Optional To As Variant, Optional OpenAfterPublish As Variant, Optional FixedFormatExtClassPtr As Variant)",availability:super::MemberAvailability::Implemented,implementation_symbol:"crates/vba/src/print_api.rs::Bridge::print_api"});
    }
    for owner in ["Workbook", "Worksheet", "Range", "Worksheets", "Sheets", "Window"] {
        for (name, signature) in [
            (
                "PrintOut",
                "Sub PrintOut(Optional From As Variant, Optional To As Variant, Optional Copies As Variant, Optional Preview As Variant, Optional ActivePrinter As Variant, Optional PrintToFile As Variant, Optional Collate As Variant, Optional PrToFileName As Variant, Optional IgnorePrintAreas As Variant)",
            ),
            ("PrintPreview", "Sub PrintPreview(Optional EnableChanges As Variant)"),
        ] {
            out.push(super::MemberDescriptor {
                owner,
                name,
                signature,
                availability: if name == "PrintOut" || owner != "Range" {
                    super::MemberAvailability::Implemented
                } else {
                    super::MemberAvailability::ServiceRequired
                },
                implementation_symbol: "crates/vba/src/print_api.rs::Bridge::print_api",
            });
        }
    }
    out.push(super::MemberDescriptor {
        owner: "Application",
        name: "ActivePrinter",
        signature: "Property ActivePrinter As String",
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/print_api.rs::Bridge::print_api",
    });
    out.push(super::MemberDescriptor {
        owner: "Application",
        name: "PrintCommunication",
        signature: "Property PrintCommunication As Boolean",
        availability: super::MemberAvailability::OutOfScope,
        implementation_symbol: "crates/vba/src/print_api.rs::Bridge::print_api",
    });
    out
}
