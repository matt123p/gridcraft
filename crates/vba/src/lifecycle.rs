//! Workbook and worksheet projections over existing session commands and identities.
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, MAX_OBJECTS, Object, Session, Value, failure, values};
use gridcraft_engine::model::Visibility;
use serde_json::json;

fn reads(call: &HostCall) -> bool {
    matches!(call.kind, InvokeKind::Get | InvokeKind::Call)
}
fn supplied(value: &Value) -> bool {
    !matches!(value, Value::Missing)
}
fn unsupported(args: &[Value], start: usize) -> HostResult<()> {
    if args.iter().skip(start).any(supplied) {
        return Err(failure(1004, "this lifecycle option is outside GridCraft's supported profile"));
    }
    Ok(())
}
pub(crate) fn file_name(value: &Value) -> HostResult<String> {
    let name = string(value)?;
    if name.is_empty() || name.len() > 4096 || name.contains('\0') {
        return Err(failure(1004, "invalid workbook file name"));
    }
    Ok(name)
}
pub(crate) fn workbook_format(path: &str) -> HostResult<()> {
    if !matches!(gridcraft_engine::io::FileKind::from_path(path), Some(gridcraft_engine::io::FileKind::Xlsx)) {
        return Err(failure(1004, "VBA workbook persistence currently requires an .xlsx/.xlsm/.xltx/.xltm destination"));
    }
    Ok(())
}
pub(crate) const SAVE_AS_ARGS: &[&str] = &[
    "Filename",
    "FileFormat",
    "Password",
    "WriteResPassword",
    "ReadOnlyRecommended",
    "CreateBackup",
    "AccessMode",
    "ConflictResolution",
    "AddToMru",
    "TextCodepage",
    "TextVisualLayout",
    "Local",
];
pub(crate) const OPEN_ARGS: &[&str] = &[
    "Filename",
    "UpdateLinks",
    "ReadOnly",
    "Format",
    "Password",
    "WriteResPassword",
    "IgnoreReadOnlyRecommended",
    "Origin",
    "Delimiter",
    "Editable",
    "Notify",
    "Converter",
    "AddToMru",
    "Local",
    "CorruptLoad",
];
/// Validate the complete call before exposing a picker or executing user callbacks.
/// A missing path is deliberately retained for the configured prompt service.
pub(crate) fn preflight(call: &HostCall) -> HostResult<Vec<Value>> {
    let member = call.member.to_ascii_lowercase();
    let names = match member.as_str() {
        "save" => &[][..],
        "saveas" => SAVE_AS_ARGS,
        "savecopyas" => &["Filename"][..],
        "close" => &["SaveChanges", "Filename", "RouteWorkbook"][..],
        "open" => OPEN_ARGS,
        _ => return Err(failure(438, "unsupported lifecycle operation")),
    };
    let args = arguments(call, names)?;
    match member.as_str() {
        "open" => {
            if supplied(&args[0]) {
                file_name(&args[0])?;
            }
            if supplied(&args[1]) && numeric_index(&args[1])? != 0 {
                return Err(failure(1004, "automatic link updating is unavailable"));
            }
            if supplied(&args[2]) && super::address::boolean(&args[2])? {
                return Err(failure(1004, "read-only documents are unavailable"));
            }
            if supplied(&args[3]) {
                return Err(failure(1004, "text format options are unavailable"));
            }
            if supplied(&args[4]) {
                string(&args[4])?;
            }
            unsupported(&args, 5)?;
        }
        "saveas" | "savecopyas" => {
            if member == "saveas" {
                unsupported(&args, 2)?;
                if supplied(&args[1]) && !matches!(numeric_index(&args[1])?, 51..=54) {
                    return Err(failure(1004, "unsupported FileFormat"));
                }
            }
            if supplied(&args[0]) {
                let path = file_name(&args[0])?;
                workbook_format(&path)?;
                if member == "saveas" && supplied(&args[1]) {
                    let extension = std::path::Path::new(&path).extension().and_then(|v| v.to_str()).unwrap_or("");
                    let expected = if extension.eq_ignore_ascii_case("xlsm") {
                        52
                    } else if extension.eq_ignore_ascii_case("xltm") {
                        53
                    } else if extension.eq_ignore_ascii_case("xltx") {
                        54
                    } else {
                        51
                    };
                    if numeric_index(&args[1])? != expected {
                        return Err(failure(1004, "FileFormat must match destination extension"));
                    }
                }
            }
        }
        "close" => {
            unsupported(&args, 2)?;
            if supplied(&args[0]) {
                super::address::boolean(&args[0])?;
            }
            if supplied(&args[1]) {
                workbook_format(&file_name(&args[1])?)?;
            }
        }
        _ => {}
    }
    Ok(args)
}
impl Bridge {
    pub(crate) fn lifecycle(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        let member = call.member.to_ascii_lowercase();
        let supported = match object {
            Object::Application => member == "workbooks" && reads(call),
            Object::Workbooks => reads(call) && matches!(member.as_str(), "count" | "item" | "add" | "open"),
            Object::Workbook => {
                (reads(call) && matches!(member.as_str(), "path" | "fullname" | "saved" | "codename" | "protectstructure"))
                    || (call.kind == InvokeKind::Call
                        && matches!(member.as_str(), "activate" | "save" | "saveas" | "savecopyas" | "close" | "protect" | "unprotect"))
            }
            Object::Worksheets | Object::Sheets => member == "add" && reads(call),
            Object::Sheet(_) => {
                (reads(call) && matches!(member.as_str(), "index" | "codename" | "visible" | "delete"))
                    || (call.kind == InvokeKind::Let && member == "visible")
                    || (call.kind == InvokeKind::Call && matches!(member.as_str(), "delete" | "copy" | "move"))
            }
            _ => false,
        };
        if !supported {
            return None;
        }
        Some((|| match object {
            Object::Application => {
                if call.arguments.is_empty() {
                    self.intern(Object::Workbooks)
                } else {
                    self.workbook_item(session, call)
                }
            }
            Object::Workbooks => match member.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    Ok(Value::Long(i32::try_from(session.documents().len()).map_err(|_| failure(6, "workbook Count overflow"))?))
                }
                "item" => self.workbook_item(session, call),
                "add" => {
                    let args = arguments(call, &["Template"])?;
                    if supplied(&args[0]) && numeric_index(&args[0])? != -4167 {
                        return Err(failure(1004, "only the xlWBATWorksheet template is available"));
                    }
                    if session.documents().len() >= 64 {
                        return Err(failure(7, "workbook budget exceeded"));
                    }
                    // Reserve the returned handle before creating a document.
                    self.objects.reserve(1)?;
                    if supplied(&args[0]) {
                        session.new_single_sheet_workbook().map_err(|e| failure(1004, e.to_string()))?;
                    } else {
                        session.execute("file.new", json!({})).map_err(|e| failure(1004, e.to_string()))?;
                    }
                    self.sync_sheets(session)?;
                    self.intern(Object::Workbook)
                }
                "open" => {
                    let args = arguments(
                        call,
                        &[
                            "Filename",
                            "UpdateLinks",
                            "ReadOnly",
                            "Format",
                            "Password",
                            "WriteResPassword",
                            "IgnoreReadOnlyRecommended",
                            "Origin",
                            "Delimiter",
                            "Editable",
                            "Notify",
                            "Converter",
                            "AddToMru",
                            "Local",
                            "CorruptLoad",
                        ],
                    )?;
                    let path = file_name(required(&args, 0)?)?;
                    // Native open preserves links; it does not perform provider refresh.
                    if supplied(&args[1]) && numeric_index(&args[1])? != 0 {
                        return Err(failure(1004, "automatic link updating is unavailable"));
                    }
                    if supplied(&args[2]) && super::address::boolean(&args[2])? {
                        return Err(failure(1004, "GridCraft has no read-only document mode"));
                    }
                    if supplied(&args[3]) {
                        return Err(failure(1004, "text format conversion options are unavailable"));
                    }
                    unsupported(&args, 5)?;
                    let password = supplied(&args[4]).then(|| string(&args[4])).transpose()?;
                    if session.documents().len() >= 64 {
                        return Err(failure(7, "workbook budget exceeded"));
                    }
                    self.objects.reserve(1)?;
                    session.open_workbook(&path, password.as_deref()).map_err(|e| failure(1004, e.to_string()))?;
                    // Imported code is never enabled or executed by Open.
                    self.sync_sheets(session)?;
                    self.intern(Object::Workbook)
                }
                _ => Err(failure(438, "unsupported Workbooks member")),
            },
            Object::Workbook => self.workbook_lifecycle(session, call, &member),
            Object::Worksheets | Object::Sheets => self.add_sheets(session, call),
            Object::Sheet(sheet) => self.sheet_lifecycle(session, *sheet, call, &member),
            _ => Err(failure(438, "unsupported lifecycle receiver")),
        })())
    }
    fn workbook_item(&mut self, session: &Session, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Index"])?;
        let value = required(&args, 0)?;
        let doc = if matches!(value, Value::String(_)) {
            let name = string(value)?;
            session.documents().iter().find(|doc| {
                let filename = doc.path.as_deref().and_then(|p| std::path::Path::new(p).file_name()).and_then(|p| p.to_str()).unwrap_or(&doc.title);
                filename.eq_ignore_ascii_case(&name)
            })
        } else {
            let index = usize::try_from(numeric_index(value)?.checked_sub(1).ok_or_else(|| failure(9, "workbook index out of range"))?)
                .map_err(|_| failure(9, "workbook index out of range"))?;
            session.documents().get(index)
        }
        .ok_or_else(|| failure(9, "workbook not found"))?;
        self.intern(Object::Document(doc.uid, Box::new(Object::Workbook)))
    }
    fn workbook_lifecycle(&mut self, session: &mut Session, call: &HostCall, member: &str) -> HostResult<Value> {
        match member {
            "path" | "fullname" | "saved" | "codename" | "protectstructure" => {
                arguments(call, &[])?;
                let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
                match member {
                    "saved" => Ok(Value::Boolean(!doc.is_dirty())),
                    "protectstructure" => Ok(Value::Boolean(doc.wb.protected_structure)),
                    "codename" => values::vba_text(doc.wb.code_name.as_deref().unwrap_or("")),
                    "fullname" => values::vba_text(doc.path.as_deref().unwrap_or(&doc.title)),
                    _ => values::vba_text(doc.path.as_deref().and_then(|p| std::path::Path::new(p).parent()).and_then(|p| p.to_str()).unwrap_or("")),
                }
            }
            "protect" | "unprotect" => {
                let args = arguments(call, if member == "protect" { &["Password", "Structure", "Windows"] } else { &["Password"] })?;
                let password = supplied(&args[0]).then(|| string(&args[0])).transpose()?.unwrap_or_default();
                if member == "protect" {
                    if !supplied(&args[1]) || !super::address::boolean(&args[1])? {
                        return Err(failure(1004, "Workbook.Protect currently requires Structure:=True"));
                    }
                    if supplied(&args[2]) && super::address::boolean(&args[2])? {
                        return Err(failure(1004, "window protection is unavailable"));
                    }
                }
                session
                    .execute("review.protectWorkbook", json!({"on":member == "protect","password":password}))
                    .map_err(|e| failure(1004, e.to_string()))?;
                Ok(Value::Empty)
            }
            "activate" => {
                arguments(call, &[])?;
                // WorkbookHost commits this activation after leaving its temporary target.
                Ok(Value::Empty)
            }
            "save" => {
                arguments(call, &[])?;
                let path = session
                    .doc()
                    .map_err(|e| failure(91, e.to_string()))?
                    .path
                    .as_deref()
                    .ok_or_else(|| failure(1004, "unsaved workbook requires SaveAs with Filename"))?;
                workbook_format(path)?;
                session.execute("file.save", json!({})).map_err(|e| failure(1004, e.to_string()))?;
                Ok(Value::Empty)
            }
            "saveas" | "savecopyas" => {
                let args = arguments(
                    call,
                    if member == "savecopyas" {
                        &["Filename"]
                    } else {
                        &[
                            "Filename",
                            "FileFormat",
                            "Password",
                            "WriteResPassword",
                            "ReadOnlyRecommended",
                            "CreateBackup",
                            "AccessMode",
                            "ConflictResolution",
                            "AddToMru",
                            "TextCodepage",
                            "TextVisualLayout",
                            "Local",
                        ]
                    },
                )?;
                let path = file_name(required(&args, 0)?)?;
                workbook_format(&path)?;
                if member == "saveas" {
                    if supplied(&args[1]) {
                        let format = numeric_index(&args[1])?;
                        let extension = std::path::Path::new(&path).extension().and_then(|p| p.to_str()).unwrap_or("");
                        let expected = if extension.eq_ignore_ascii_case("xlsm") {
                            52
                        } else if extension.eq_ignore_ascii_case("xltm") {
                            53
                        } else if extension.eq_ignore_ascii_case("xltx") {
                            54
                        } else {
                            51
                        };
                        if format != expected {
                            return Err(failure(1004, "FileFormat must match the supported destination extension"));
                        }
                    }
                    unsupported(&args, 2)?;
                }
                session
                    .execute(if member == "saveas" { "file.saveAs" } else { "file.saveCopy" }, json!({"path":path}))
                    .map_err(|e| failure(1004, e.to_string()))?;
                Ok(Value::Empty)
            }
            "close" => {
                let args = arguments(call, &["SaveChanges", "Filename", "RouteWorkbook"])?;
                unsupported(&args, 2)?;
                let dirty = session.doc().map_err(|e| failure(91, e.to_string()))?.is_dirty();
                let save = supplied(&args[0]).then(|| super::address::boolean(&args[0])).transpose()?;
                if dirty && save.is_none() {
                    return Err(failure(1004, "dirty workbook Close requires explicit SaveChanges; no dialog is opened"));
                }
                if save == Some(true) && dirty {
                    let path = if supplied(&args[1]) {
                        file_name(&args[1])?
                    } else {
                        session
                            .doc()
                            .map_err(|e| failure(91, e.to_string()))?
                            .path
                            .clone()
                            .ok_or_else(|| failure(1004, "unsaved workbook Close requires Filename"))?
                    };
                    workbook_format(&path)?;
                    session.execute("file.save", json!({"path":path})).map_err(|e| failure(1004, e.to_string()))?;
                }
                session.execute("file.close", json!({"force":true,"document":true})).map_err(|e| failure(1004, e.to_string()))?;
                Ok(Value::Empty)
            }
            _ => Err(failure(438, "unsupported Workbook member")),
        }
    }
    fn destination(&self, session: &Session, args: &[Value]) -> HostResult<Option<(usize, bool)>> {
        let before = args.first().filter(|v| supplied(v));
        let after = args.get(1).filter(|v| supplied(v));
        if before.is_some() && after.is_some() {
            return Err(failure(1004, "Before and After are mutually exclusive"));
        }
        let Some(value) = before.or(after) else { return Ok(None) };
        let Value::Object(reference) = value else { return Err(failure(13, "worksheet destination must be an object")) };
        let Object::Sheet(sheet) = self.object(session, *reference)? else { return Err(failure(13, "worksheet destination required")) };
        if session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).is_none() {
            return Err(failure(91, "destination worksheet was deleted"));
        }
        Ok(Some((sheet, after.is_some())))
    }
    fn add_sheets(&mut self, session: &mut Session, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Before", "After", "Count", "Type"])?;
        if supplied(&args[3]) && numeric_index(&args[3])? != -4167 {
            return Err(failure(1004, "only worksheet creation is available; chartsheets and templates are unavailable"));
        }
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        if doc.wb.protected_structure {
            return Err(failure(1004, "workbook structure is protected"));
        }
        let count = if supplied(&args[2]) { numeric_index(&args[2])? } else { 1 };
        let count = usize::try_from(count).ok().filter(|n| (1..=255).contains(n)).ok_or_else(|| failure(1004, "Count must be between 1 and 255"))?;
        if doc.wb.sheets.len().checked_add(count).is_none_or(|n| n > MAX_OBJECTS) {
            return Err(failure(7, "worksheet budget exceeded"));
        }
        let before = self.destination(session, &args)?.map(|(i, after)| i + usize::from(after)).unwrap_or(doc.sheet_index());
        self.objects.reserve(1)?;
        for offset in 0..count {
            session.execute("home.insertSheet", json!({"before":before + offset})).map_err(|e| failure(1004, e.to_string()))?;
        }
        session.execute("sheet.activate", json!({"sheet":before})).map_err(|e| failure(1004, e.to_string()))?;
        // Existing handles must be remapped before interning the newly inserted index.
        self.sync_sheets(session)?;
        self.intern(Object::Sheet(before))
    }
    fn sheet_lifecycle(&mut self, session: &mut Session, sheet: usize, call: &HostCall, member: &str) -> HostResult<Value> {
        if matches!(member, "index" | "codename") || (member == "visible" && reads(call)) {
            arguments(call, &[])?;
            let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
            return match member {
                "index" => Ok(Value::Long(i32::try_from(sheet + 1).map_err(|_| failure(6, "Index overflow"))?)),
                "codename" => values::vba_text(sh.code_name.as_deref().unwrap_or("")),
                _ => Ok(Value::Long(match sh.visibility {
                    Visibility::Visible => -1,
                    Visibility::Hidden => 0,
                    Visibility::VeryHidden => 2,
                })),
            };
        }
        if member == "visible" {
            let args = arguments(call, &["Value"])?;
            let visibility = match numeric_index(required(&args, 0)?)? {
                -1 => Visibility::Visible,
                0 => Visibility::Hidden,
                2 => Visibility::VeryHidden,
                _ => return Err(failure(1004, "invalid worksheet visibility")),
            };
            session.set_worksheet_visibility(sheet, visibility).map_err(|e| failure(1004, e.to_string()))?;
        } else if member == "delete" {
            arguments(call, &[])?;
            session.execute("home.deleteSheet", json!({"sheet":sheet})).map_err(|e| failure(1004, e.to_string()))?;
            self.sync_sheets(session)?;
            return Ok(Value::Boolean(true));
        } else {
            let args = arguments(call, &["Before", "After"])?;
            let Some((dest, after)) = self.destination(session, &args)? else {
                return Err(failure(1004, "new-workbook sheet transfers require a native transfer service; supply Before or After in this workbook"));
            };
            let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
            if member == "copy" && doc.wb.sheet(sheet).is_some_and(|sh| sh.code_name.is_some() || !sh.tables.is_empty() || sh.preserved.is_some()) {
                return Err(failure(1004, "copying document modules, tables or preserved package parts requires a native transfer service"));
            }
            if doc.wb.sheets.len() >= MAX_OBJECTS && member == "copy" {
                return Err(failure(7, "worksheet budget exceeded"));
            }
            let insertion = dest + usize::from(after);
            let to = if member == "move" && sheet < insertion { insertion - 1 } else { insertion };
            session.execute("sheet.move", json!({"sheet":sheet,"to":to,"copy":member == "copy"})).map_err(|e| failure(1004, e.to_string()))?;
            self.sync_sheets(session)?;
        }
        Ok(Value::Empty)
    }
}
