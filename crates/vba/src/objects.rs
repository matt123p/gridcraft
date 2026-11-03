use super::ranges::{DataSeriesOptions, Range, Unit};
use super::{Bridge, HostCall, HostResult, InvokeKind, MAX_CELLS, MAX_OBJECTS, Session, Value, failure, ranges, values};
use gridcraft_engine::core::RangeRef;
use std::collections::BTreeSet;
use vba_foreign::HostArgument;

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Object {
    Application,
    Document(u64, Box<Object>),
    Workbook,
    Worksheets,
    Sheets,
    Sheet(usize),
    Range(usize, Range),
    Areas(usize, Range),
    DeletedRange(bool),
    Format(usize, Range, bool),
}

impl Object {
    fn receiver_name(&self) -> &'static str {
        match self {
            Self::Application => "Application",
            Self::Document(_, object) => object.receiver_name(),
            Self::Workbook => "Workbook",
            Self::Worksheets => "Worksheets",
            Self::Sheets => "Sheets",
            Self::Sheet(_) => "Worksheet",
            Self::Range(_, _) | Self::DeletedRange(_) => "Range",
            Self::Areas(_, _) => "Areas",
            Self::Format(_, _, true) => "Font",
            Self::Format(_, _, false) => "Interior",
        }
    }
}

/// Bind names before dispatch; never silently drop named/omitted/extra arguments.
pub(crate) fn arguments(call: &HostCall, names: &[&str]) -> HostResult<Vec<Value>> {
    let mut result = vec![Value::Missing; names.len()];
    let mut assigned = vec![false; names.len()];
    let mut named = false;
    for (position, HostArgument { name, value, .. }) in call.arguments.iter().enumerate() {
        let index = match name {
            Some(name) => {
                named = true;
                names
                    .iter()
                    .position(|n| n.eq_ignore_ascii_case(name) || (name == "$value" && *n == "Value"))
                    .ok_or_else(|| failure(448, "unknown named argument"))?
            }
            None if !named => position,
            None => return Err(failure(5, "positional argument after named argument")),
        };
        let slot = result.get_mut(index).ok_or_else(|| failure(450, "wrong number of arguments"))?;
        let used = assigned.get_mut(index).ok_or_else(|| failure(450, "wrong number of arguments"))?;
        if *used {
            return Err(failure(450, "duplicate argument"));
        }
        *used = true;
        *slot = value.clone();
    }
    Ok(result)
}
pub(crate) fn required(args: &[Value], index: usize) -> HostResult<&Value> {
    args.get(index).filter(|v| !matches!(v, Value::Missing)).ok_or_else(|| failure(449, "argument not optional"))
}
pub(crate) fn string(value: &Value) -> HostResult<String> {
    super::values::text(value)
}
fn read_kind(kind: InvokeKind) -> bool {
    matches!(kind, InvokeKind::Call | InvokeKind::Get)
}
pub(crate) fn numeric_index(value: &Value) -> HostResult<i32> {
    let value = vba_runtime::bind(vba_values::ScalarKind::Long, value).map_err(|error| match error {
        vba_runtime::Error::Vba { number, message } => failure(number, &message),
        vba_runtime::Error::Resource(message) => failure(7, &message),
        vba_runtime::Error::Unsupported(message) => failure(13, &message),
    })?;
    let Value::Long(index) = value else {
        return Err(failure(13, "index type mismatch"));
    };
    Ok(index)
}

impl Bridge {
    fn collection_item(&mut self, session: &Session, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Index"])?;
        let value = required(&args, 0)?;
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let sheet = if matches!(value, Value::String(_)) {
            let name = string(value)?;
            doc.wb.sheets.iter().position(|s| s.name.eq_ignore_ascii_case(&name)).ok_or_else(|| failure(9, "worksheet not found"))?
        } else {
            let index = numeric_index(value)?;
            usize::try_from(index.checked_sub(1).ok_or_else(|| failure(9, "worksheet index out of range"))?)
                .ok()
                .filter(|i| *i < doc.wb.sheets.len())
                .ok_or_else(|| failure(9, "worksheet index out of range"))?
        };
        self.intern(Object::Sheet(sheet))
    }
    fn index(value: &Value, column: bool) -> HostResult<i32> {
        if column && let Value::String(_) = value {
            let text = string(value)?;
            if text.bytes().all(|b| b.is_ascii_alphabetic()) && !text.is_empty() {
                return gridcraft_engine::core::letters_to_col(&text)
                    .and_then(|c| i32::try_from(c + 1).ok())
                    .ok_or_else(|| failure(1004, "invalid column address"));
            }
        }
        numeric_index(value)
    }
    fn range_item(range: &Range, call: &HostCall, write: bool) -> HostResult<Range> {
        let args = arguments(call, if write { &["RowIndex", "ColumnIndex", "Value"] } else { &["RowIndex", "ColumnIndex"] })?;
        if range.unit != Unit::Cells && args.get(1).is_some_and(|v| !matches!(v, Value::Missing)) {
            return Err(failure(1004, "Rows/Columns Item takes one index"));
        }
        if range.unit != Unit::Cells
            && matches!(args.get(1), None | Some(Value::Missing))
            && let Value::String(_) = required(&args, 0)?
        {
            let text = string(required(&args, 0)?)?;
            if let Some((a, b)) = text.split_once(':') {
                return range.slice(
                    Self::index(&values::vba_text(a)?, range.unit == Unit::Columns)?,
                    Self::index(&values::vba_text(b)?, range.unit == Unit::Columns)?,
                );
            }
        }
        let row = Self::index(required(&args, 0)?, range.unit == Unit::Columns)?;
        let column = args.get(1).filter(|v| !matches!(v, Value::Missing)).map(|v| Self::index(v, true)).transpose()?;
        range.item(row, column)
    }
    fn subranges(&mut self, sheet: usize, range: &Range, unit: Unit, call: &HostCall) -> HostResult<Value> {
        let range = range.collection(unit)?;
        if call.arguments.is_empty() {
            return self.intern(Object::Range(sheet, range));
        }
        if unit != Unit::Cells {
            let args = arguments(call, &["RowIndex", "ColumnIndex"])?;
            if matches!(args.get(1), None | Some(Value::Missing))
                && let Value::String(_) = required(&args, 0)?
            {
                let text = string(required(&args, 0)?)?;
                if let Some((a, b)) = text.split_once(':') {
                    let a = Self::index(&values::vba_text(a)?, unit == Unit::Columns)?;
                    let b = Self::index(&values::vba_text(b)?, unit == Unit::Columns)?;
                    return self.intern(Object::Range(sheet, range.slice(a, b)?));
                }
            }
        }
        self.intern(Object::Range(sheet, Self::range_item(&range, call, false)?))
    }
    fn enumerate(&mut self, session: &Session, object: Object, call: &HostCall) -> HostResult<Value> {
        arguments(call, &[])?;
        let objects = match object {
            Object::Worksheets | Object::Sheets => {
                let n = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheets.len();
                if n > MAX_OBJECTS {
                    return Err(failure(7, "collection enumeration budget"));
                }
                (0..n).map(Object::Sheet).collect::<Vec<_>>()
            }
            Object::Areas(sheet, range) => range.areas.into_iter().map(|r| Object::Range(sheet, r.into())).collect(),
            Object::Range(sheet, range) => range.entries()?.into_iter().map(|r| Object::Range(sheet, r)).collect(),
            _ => return Err(failure(438, "object is not enumerable")),
        };
        let objects = objects.into_iter().map(|o| Object::Document(self.context, Box::new(o))).collect();
        let handles = self.objects.intern_many(objects)?;
        let mut array = if handles.is_empty() {
            vba_values::VbaArray::empty(vba_values::ScalarKind::Object, 1)
        } else {
            vba_values::VbaArray::new(
                vba_values::ScalarKind::Object,
                vec![(1, i32::try_from(handles.len()).map_err(|_| failure(7, "enumeration budget"))?)],
                false,
            )
            .map_err(|e| failure(7, e.to_string()))?
        };
        for (i, handle) in handles.into_iter().enumerate() {
            array.set(i, Value::Object(self.reference(handle))).map_err(|e| failure(7, e.to_string()))?;
        }
        Ok(Value::Array(array))
    }
    /// Single dispatcher over the Excel object profile: each branch mirrors one member name or
    /// object kind. Splitting the ladder into helpers would separate the mapping from VBA's own
    /// member table, so the size is deliberate.
    #[allow(clippy::cognitive_complexity)]
    pub(crate) fn invoke_object(&mut self, session: &mut Session, object: Object, call: HostCall) -> HostResult<Value> {
        if let Object::Sheet(s) | Object::Range(s, _) | Object::Areas(s, _) = &object
            && session.doc().map_err(|_| failure(91, "workbook closed"))?.wb.sheet(*s).is_none()
        {
            return Err(failure(91, "worksheet was deleted"));
        }
        if let Object::Format(sheet, range, font) = &object {
            return self.format(session, *sheet, range, *font, &call);
        }
        if matches!(object, Object::DeletedRange(_)) {
            return Err(failure(1004, "Range was deleted by a structural edit"));
        }
        if call.member.len() > 256 || call.arguments.len() > 256 {
            return Err(failure(7, "host call budget exceeded"));
        }
        let mut call = call;
        if call.kind == InvokeKind::Enumerate {
            return self.enumerate(session, object, &call);
        }
        match call.kind {
            InvokeKind::DefaultGet => {
                call.kind = InvokeKind::Get;
                call.member = match &object {
                    Object::Range(..) if call.arguments.is_empty() => "Value",
                    Object::Range(..) | Object::Areas(..) | Object::Worksheets | Object::Sheets => "Item",
                    _ => return Err(failure(438, "object has no default property")),
                }
                .into();
            }
            InvokeKind::DefaultLet => {
                call.kind = InvokeKind::Let;
                call.member = match &object {
                    Object::Range(..) if call.arguments.len() == 1 => "Value",
                    Object::Range(..) => "Item",
                    _ => return Err(failure(438, "default property is not writable")),
                }
                .into();
            }
            InvokeKind::DefaultSet => return Err(failure(13, "default Range/collection property cannot be assigned with Set")),
            _ => {}
        }
        let member = call.member.to_ascii_lowercase();
        let receiver = object.receiver_name();
        if object == Object::Application && matches!(member.as_str(), "union" | "intersect") && read_kind(call.kind) {
            let names = (1..=30).map(|index| format!("Arg{index}")).collect::<Vec<_>>();
            let name_refs = names.iter().map(String::as_str).collect::<Vec<_>>();
            let args = arguments(&call, &name_refs)?;
            required(&args, 0)?;
            required(&args, 1)?;
            let mut sheet = None;
            let mut ranges = Vec::new();
            for value in args.iter().filter(|value| !matches!(value, Value::Missing)) {
                let Value::Object(reference) = value else {
                    return Err(failure(13, format!("Application.{} arguments must be Range objects", call.member)));
                };
                let Object::Range(candidate_sheet, range) = self.object(session, *reference)? else {
                    return Err(failure(13, format!("Application.{} arguments must be Range objects", call.member)));
                };
                if sheet.is_some_and(|existing| existing != candidate_sheet) {
                    return Err(failure(1004, format!("Application.{} requires ranges on one worksheet", call.member)));
                }
                sheet = Some(candidate_sheet);
                ranges.push(range);
            }
            let sheet = sheet.ok_or_else(|| failure(449, format!("Application.{} requires two ranges", call.member)))?;
            let mut cells = BTreeSet::new();
            if member == "union" {
                for range in ranges {
                    range.validate()?;
                    for area in range.areas {
                        for cell in area.iter() {
                            cells.insert(cell);
                            if cells.len() > MAX_CELLS {
                                return Err(failure(7, "Application.Union exceeds VBA cell budget"));
                            }
                        }
                    }
                }
            } else {
                let first = ranges.first().ok_or_else(|| failure(449, "Application.Intersect requires two ranges"))?;
                first.validate()?;
                for area in &first.areas {
                    for cell in area.iter() {
                        cells.insert(cell);
                        if cells.len() > MAX_CELLS {
                            return Err(failure(7, "Application.Intersect exceeds VBA cell budget"));
                        }
                    }
                }
                for range in ranges.iter().skip(1) {
                    range.validate()?;
                    cells.retain(|cell| range.areas.iter().any(|area| area.contains(*cell)));
                }
                if cells.is_empty() {
                    return Ok(Value::Nothing);
                }
            }
            return self.intern(Object::Range(sheet, ranges::from_cells(cells.into_iter().collect(), &format!("Application.{}", call.member))?));
        }
        if object == Object::Application && member == "goto" && call.kind == InvokeKind::Call {
            let args = arguments(&call, &["Reference", "Scroll"])?;
            let (reference, preferred_active) = match args.first().filter(|value| !matches!(value, Value::Missing)) {
                Some(Value::Object(reference)) => {
                    let Object::Range(sheet, range) = self.object(session, *reference)? else {
                        return Err(failure(1004, "Application.Goto Reference must be a Range or reference text"));
                    };
                    let name = session
                        .doc()
                        .map_err(|error| failure(91, error.to_string()))?
                        .wb
                        .sheet(sheet)
                        .map(|sheet| gridcraft_engine::core::quote_sheet(&sheet.name))
                        .ok_or_else(|| failure(91, "worksheet gone"))?;
                    (range.areas.iter().map(|area| format!("{name}!{}", area.a1())).collect::<Vec<_>>().join(","), Some(range.first()?.start.a1()))
                }
                Some(value @ Value::String(_)) => (string(value)?, None),
                Some(_) => return Err(failure(1004, "Application.Goto Reference must be a Range or reference text")),
                None => (
                    session
                        .doc()
                        .map_err(|error| failure(91, error.to_string()))?
                        .go_to_recent
                        .first()
                        .cloned()
                        .ok_or_else(|| failure(1004, "Application.Goto history is empty"))?,
                    None,
                ),
            };
            let scroll = args.get(1).filter(|value| !matches!(value, Value::Missing)).map(super::address::boolean).transpose()?.unwrap_or(false);
            let mut params = serde_json::json!({"reference":reference,"record":true});
            if let Some(active) = preferred_active {
                params["active"] = serde_json::Value::String(active);
            }
            session.execute("edit.goTo", params).map_err(|error| failure(1004, error.to_string()))?;
            if scroll {
                let doc = session.doc_mut().map_err(|error| failure(91, error.to_string()))?;
                let sheet = doc.sheet_index();
                let active = doc.selection.active;
                doc.store_view(&[(sheet, active)]);
            }
            return Ok(Value::Empty);
        }
        if object == Object::Application && member == "previousselections" && read_kind(call.kind) {
            let args = arguments(&call, &["Index"])?;
            let selections = session.doc().map_err(|error| failure(91, error.to_string()))?.previous_selections.clone();
            let objects = selections
                .into_iter()
                .map(|(sheet, selection)| Object::Range(sheet, Range { areas: selection.ranges, unit: Unit::Cells }))
                .collect::<Vec<_>>();
            if let Some(index) = args.first().filter(|value| !matches!(value, Value::Missing)) {
                let index = usize::try_from(numeric_index(index)?.checked_sub(1).ok_or_else(|| failure(9, "PreviousSelections index out of range"))?)
                    .map_err(|_| failure(9, "PreviousSelections index out of range"))?;
                let object = objects.get(index).cloned().ok_or_else(|| failure(9, "PreviousSelections index out of range"))?;
                return self.intern(object);
            }
            let handles = self.objects.intern_many(objects.into_iter().map(|object| Object::Document(self.context, Box::new(object))).collect())?;
            let mut array = if handles.is_empty() {
                vba_values::VbaArray::empty(vba_values::ScalarKind::Object, 1)
            } else {
                vba_values::VbaArray::new(
                    vba_values::ScalarKind::Object,
                    vec![(1, i32::try_from(handles.len()).map_err(|_| failure(7, "PreviousSelections budget"))?)],
                    false,
                )
                .map_err(|error| failure(7, error.to_string()))?
            };
            for (index, handle) in handles.into_iter().enumerate() {
                array.set(index, Value::Object(self.reference(handle))).map_err(|error| failure(7, error.to_string()))?;
            }
            return Ok(Value::Array(array));
        }
        if matches!(call.kind, InvokeKind::Let | InvokeKind::Set)
            && matches!(object, Object::Application | Object::Sheet(_))
            && matches!(member.as_str(), "range" | "cells" | "rows" | "columns")
        {
            if call.kind == InvokeKind::Set {
                return Err(failure(13, "Range assignment requires Let"));
            }
            let value = call
                .arguments
                .pop()
                .filter(|a| a.name.as_deref() == Some("$value") || a.name.is_none())
                .ok_or_else(|| failure(450, "missing assignment value"))?
                .value;
            call.kind = InvokeKind::Get;
            let Value::Object(reference) = self.invoke_object(session, object.clone(), call)? else {
                return Err(failure(13, "Range required"));
            };
            let Object::Range(sheet, range) = self.object(session, reference)? else {
                return Err(failure(13, "Range required"));
            };
            values::write(session, sheet, range.rectangle()?, &value, false, true)?;
            return Ok(Value::Empty);
        }
        if read_kind(call.kind) && member == "application" {
            arguments(&call, &[])?;
            return self.intern(Object::Application);
        }
        if read_kind(call.kind) && member == "parent" {
            arguments(&call, &[])?;
            return self.intern(match object {
                Object::Application | Object::Workbook => Object::Application,
                Object::Document(..) => return Err(failure(91, "unresolved document owner")),
                Object::Sheet(_) | Object::Worksheets | Object::Sheets => Object::Workbook,
                Object::Range(sheet, _) | Object::Areas(sheet, _) => Object::Sheet(sheet),
                Object::Format(sheet, range, _) => Object::Range(sheet, range),
                Object::DeletedRange(_) => return Err(failure(1004, "deleted Range")),
            });
        }
        if object == Object::Application && member == "calculation" {
            if read_kind(call.kind) {
                arguments(&call, &[])?;
                let mode = session.application_calculation_mode().map_err(|e| failure(91, e.to_string()))?;
                return Ok(Value::Long(match mode {
                    gridcraft_engine::model::CalcMode::Manual => -4135,
                    gridcraft_engine::model::CalcMode::AutomaticExceptTables => 2,
                    _ => -4105,
                }));
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(13, "Calculation requires Let"));
            }
            let args = arguments(&call, &["Value"])?;
            let mode = match numeric_index(required(&args, 0)?)? {
                -4135 => gridcraft_engine::model::CalcMode::Manual,
                -4105 => gridcraft_engine::model::CalcMode::Automatic,
                2 => gridcraft_engine::model::CalcMode::AutomaticExceptTables,
                _ => return Err(failure(1004, "invalid calculation mode")),
            };
            let documents: Vec<_> = session.documents().iter().map(|d| d.uid).collect();
            for doc in documents {
                self.ensure_group(session, doc).map_err(|e| failure(7, e.to_string()))?;
            }
            session.set_application_calculation_mode(mode).map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if object == Object::Application && member == "enableevents" {
            if call.kind == InvokeKind::Get {
                arguments(&call, &[])?;
                return Ok(Value::Boolean(self.enable_events));
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(13, "EnableEvents requires Let"));
            }
            let args = arguments(&call, &["Value"])?;
            let Value::Boolean(enabled) = vba_runtime::bind(vba_values::ScalarKind::Boolean, required(&args, 0)?).map_err(values::runtime_error)?
            else {
                return Err(failure(13, "EnableEvents requires Boolean"));
            };
            self.enable_events = enabled;
            if !enabled {
                self.pending_events.clear();
            }
            return Ok(Value::Empty);
        }
        if object == Object::Application && member == "cutcopymode" {
            if call.kind == InvokeKind::Get {
                arguments(&call, &[])?;
                return Ok(match session.clipboard.as_ref() {
                    Some(clipboard) if clipboard.cut => Value::Long(2),
                    Some(_) => Value::Long(1),
                    None => Value::Boolean(false),
                });
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(13, "CutCopyMode requires Let"));
            }
            let args = arguments(&call, &["Value"])?;
            let value = required(&args, 0)?;
            let command = match value {
                Value::Boolean(false) => "edit.clearClipboard",
                Value::Boolean(true) | Value::Long(1) => "edit.copy",
                Value::Long(2) => "edit.cut",
                _ if vba_runtime::bind(vba_values::ScalarKind::Boolean, value).is_ok_and(|v| v == Value::Boolean(false)) => "edit.clearClipboard",
                _ => return Err(failure(1004, "CutCopyMode accepts False, True, xlCopy or xlCut")),
            };
            let params = serde_json::json!({});
            if command == "edit.clearClipboard" {
                session.execute(command, params).map_err(|e| failure(1004, e.to_string()))?;
            } else {
                let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
                if doc.selection.ranges.len() != 1 || doc.selection.current().count() > super::MAX_CELLS as u64 {
                    return Err(failure(1004, "CutCopyMode requires one bounded selected area"));
                }
                session.execute(command, params).map_err(|e| failure(1004, e.to_string()))?;
            }
            return Ok(Value::Empty);
        }
        if member == "calculate" && matches!(object, Object::Application | Object::Workbook) && call.kind == InvokeKind::Call {
            arguments(&call, &[])?;
            if object == Object::Application {
                session.calculate_application().map_err(|e| failure(1004, e.to_string()))?;
            } else {
                session.calculate_workbook().map_err(|e| failure(1004, e.to_string()))?;
            }
            return Ok(Value::Empty);
        }
        if let Object::Sheet(sheet) = object
            && member == "name"
            && call.kind == InvokeKind::Let
        {
            let args = arguments(&call, &["Value"])?;
            let name = string(required(&args, 0)?)?;
            session.rename_worksheet(sheet, &name).map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        match object {
            Object::Application | Object::Workbook if read_kind(call.kind) => {
                let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
                let active = doc.sheet_index();
                let implicit_sheet = match self.event_owner_sheet {
                    Some(sheet) if doc.wb.sheet(sheet).is_some() => sheet,
                    Some(_) => return Err(failure(91, "worksheet event owner was deleted")),
                    None => active,
                };
                match member.as_str() {
                    "thisworkbook" | "activeworkbook" if object == Object::Application => {
                        arguments(&call, &[])?;
                        if member == "thisworkbook" {
                            let handle = self.objects.intern(Object::Document(self.doc, Box::new(Object::Workbook)))?;
                            Ok(Value::Object(self.reference(handle)))
                        } else {
                            self.intern(Object::Workbook)
                        }
                    }
                    "activecell" | "selection" if object == Object::Application => {
                        arguments(&call, &[])?;
                        let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
                        let range = if member == "activecell" {
                            RangeRef::cell(doc.selection.active).into()
                        } else {
                            ranges::Range { areas: doc.selection.ranges.clone(), unit: Unit::Cells }
                        };
                        self.intern(Object::Range(active, range))
                    }
                    "activesheet" => {
                        arguments(&call, &[])?;
                        self.intern(Object::Sheet(active))
                    }
                    "worksheets" | "sheets" => {
                        if call.arguments.is_empty() {
                            self.intern(if member == "worksheets" { Object::Worksheets } else { Object::Sheets })
                        } else {
                            self.collection_item(session, &call)
                        }
                    }
                    "range" if object == Object::Application => self.construct_range(session, implicit_sheet, None, true, &call),
                    "cells" | "rows" | "columns" if object == Object::Application => self.invoke_object(session, Object::Sheet(implicit_sheet), call),
                    "name" if object == Object::Workbook => {
                        arguments(&call, &[])?;
                        values::vba_text(&session.doc().map_err(|e| failure(91, e.to_string()))?.title)
                    }
                    _ => Err(failure(438, format!("{receiver}.{} is not a supported member", call.member))),
                }
            }
            Object::Worksheets | Object::Sheets if read_kind(call.kind) => match member.as_str() {
                "count" => {
                    arguments(&call, &[])?;
                    Ok(Value::Long(
                        i32::try_from(session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheets.len())
                            .map_err(|_| failure(6, "collection Count overflow"))?,
                    ))
                }
                "item" => self.collection_item(session, &call),
                _ => Err(failure(438, format!("{receiver}.{} is not a supported member", call.member))),
            },
            Object::Sheet(sheet) if read_kind(call.kind) => match member.as_str() {
                "select" | "activate" if call.kind == InvokeKind::Call => self.select_sheet(session, sheet, &call),
                "paste" if call.kind == InvokeKind::Call => self.paste_sheet(session, sheet, &call),
                "pastespecial" if call.kind == InvokeKind::Call => {
                    arguments(&call, &["Format", "Link", "DisplayAsIcon", "IconFileName", "IconIndex", "IconLabel", "NoHTMLFormatting"])?;
                    Err(failure(1004, "Worksheet.PasteSpecial requires an external clipboard-format service"))
                }
                "name" => {
                    arguments(&call, &[])?;
                    values::vba_text(
                        &session
                            .doc()
                            .map_err(|_| failure(91, "workbook closed"))?
                            .wb
                            .sheet(sheet)
                            .ok_or_else(|| failure(91, "worksheet gone"))?
                            .name,
                    )
                }
                "usedrange" => {
                    arguments(&call, &[])?;
                    let r = session
                        .doc()
                        .map_err(|e| failure(91, e.to_string()))?
                        .wb
                        .sheet(sheet)
                        .ok_or_else(|| failure(91, "worksheet gone"))?
                        .used_range()
                        .unwrap_or_default();
                    self.intern(Object::Range(sheet, r.into()))
                }
                "range" => self.construct_range(session, sheet, None, false, &call),
                "cells" | "rows" | "columns" => self.subranges(
                    sheet,
                    &RangeRef::all().into(),
                    match member.as_str() {
                        "rows" => Unit::Rows,
                        "columns" => Unit::Columns,
                        _ => Unit::Cells,
                    },
                    &call,
                ),
                _ => Err(failure(438, format!("Worksheet.{} is not a supported member", call.member))),
            },
            Object::Areas(sheet, range) if read_kind(call.kind) => match member.as_str() {
                "count" => {
                    arguments(&call, &[])?;
                    Ok(Value::Long(range.areas.len() as i32))
                }
                "item" => {
                    let args = arguments(&call, &["Index"])?;
                    let i = numeric_index(required(&args, 0)?)?;
                    let rect = i
                        .checked_sub(1)
                        .and_then(|i| usize::try_from(i).ok())
                        .and_then(|i| range.areas.get(i))
                        .copied()
                        .ok_or_else(|| failure(1004, "area index"))?;
                    self.intern(Object::Range(sheet, rect.into()))
                }
                _ => Err(failure(438, format!("Areas.{} is not a supported member", call.member))),
            },
            Object::Range(sheet, range) => match member.as_str() {
                "select" if call.kind == InvokeKind::Call => self.select_range(session, sheet, &range, &call),
                "activate" if call.kind == InvokeKind::Call => self.activate_range(session, sheet, &range, &call),
                "copy" if call.kind == InvokeKind::Call => self.copy_range(session, sheet, &range, &call),
                "cut" if call.kind == InvokeKind::Call => self.cut_range(session, sheet, &range, &call),
                "pastespecial" if call.kind == InvokeKind::Call => self.paste_special(session, sheet, &range, &call),
                "copypicture" if call.kind == InvokeKind::Call => {
                    arguments(&call, &["Appearance", "Format"])?;
                    Err(failure(1004, "Range.CopyPicture requires a clipboard image service"))
                }
                "font" | "interior" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    self.intern(Object::Format(sheet, range, member == "font"))
                }
                "numberformat" => self.number_format(session, sheet, &range, &call),
                "end" if read_kind(call.kind) => self.end(session, sheet, &range, &call),
                "find" | "findnext" | "findprevious" if read_kind(call.kind) => self.find(session, sheet, &range, &call),
                "replace" if read_kind(call.kind) => self.replace_range(session, sheet, &range, &call),
                "filldown" | "fillup" | "fillleft" | "fillright" if call.kind == InvokeKind::Call => {
                    arguments(&call, &[])?;
                    self.fill_direction(session, sheet, &range, member.trim_start_matches("fill"))
                }
                "autofill" if read_kind(call.kind) => {
                    let args = arguments(&call, &["Destination", "Type"])?;
                    let Value::Object(reference) = required(&args, 0)? else {
                        return Err(failure(13, "Range.AutoFill Destination must be a Range"));
                    };
                    let Object::Range(destination_sheet, destination) = self.object(session, *reference)? else {
                        return Err(failure(13, "Range.AutoFill Destination must be a Range"));
                    };
                    let fill_type = args.get(1).filter(|value| !matches!(value, Value::Missing)).map(numeric_index).transpose()?.unwrap_or(0);
                    self.auto_fill(session, sheet, &range, destination_sheet, &destination, fill_type)
                }
                "specialcells" if read_kind(call.kind) => {
                    let args = arguments(&call, &["Type", "Value"])?;
                    let cell_type = numeric_index(required(&args, 0)?)?;
                    let value_mask = args.get(1).filter(|value| !matches!(value, Value::Missing)).map(numeric_index).transpose()?.unwrap_or(23);
                    if value_mask <= 0 || value_mask & !23 != 0 {
                        return Err(failure(1004, format!("Range.SpecialCells Value mask {value_mask} is invalid")));
                    }
                    self.special_cells(session, sheet, &range, cell_type, value_mask)
                }
                "dataseries" if call.kind == InvokeKind::Call => {
                    let args = arguments(&call, &["Rowcol", "Type", "Date", "Step", "Stop", "Trend"])?;
                    let rowcol = args.first().filter(|value| !matches!(value, Value::Missing)).map(numeric_index).transpose()?;
                    let rows = match rowcol {
                        Some(1) => Some(true),
                        Some(2) => Some(false),
                        Some(value) => return Err(failure(1004, format!("Range.DataSeries Rowcol {value} is invalid"))),
                        None => None,
                    };
                    let series_type = args.get(1).filter(|value| !matches!(value, Value::Missing)).map(numeric_index).transpose()?.unwrap_or(-4132);
                    let date_unit = args.get(2).filter(|value| !matches!(value, Value::Missing)).map(numeric_index).transpose()?.unwrap_or(1);
                    let number = |index: usize, default: Option<f64>| -> HostResult<Option<f64>> {
                        let Some(value) = args.get(index).filter(|value| !matches!(value, Value::Missing)) else { return Ok(default) };
                        match vba_runtime::bind(vba_values::ScalarKind::Double, value).map_err(values::runtime_error)? {
                            Value::Double(value) if value.is_finite() => Ok(Some(value)),
                            _ => Err(failure(13, "Range.DataSeries numeric option type mismatch")),
                        }
                    };
                    let step = number(3, Some(1.0))?.ok_or_else(|| failure(13, "Range.DataSeries Step is required"))?;
                    let stop = number(4, None)?;
                    let trend =
                        args.get(5).filter(|value| !matches!(value, Value::Missing)).map(super::address::boolean).transpose()?.unwrap_or(false);
                    self.data_series(session, sheet, &range, DataSeriesOptions { rows, series_type, date_unit, step, stop, trend })
                }
                "insert" | "delete" if call.kind == InvokeKind::Call => self.structural(session, sheet, &range, &call),
                "entirerow" | "entirecolumn" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    let mut r = range.clone();
                    r.unit = Unit::Cells;
                    r.areas = r
                        .areas
                        .iter()
                        .map(|r| if member == "entirerow" { RangeRef::rows(r.start.row, r.end.row) } else { RangeRef::cols(r.start.col, r.end.col) })
                        .collect();
                    r.areas.dedup();
                    self.intern(Object::Range(sheet, r))
                }
                "range" if read_kind(call.kind) => self.construct_range(session, sheet, Some(range.first()?.start), false, &call),
                "areas" if read_kind(call.kind) => {
                    if call.arguments.is_empty() {
                        self.intern(Object::Areas(sheet, range))
                    } else {
                        self.invoke_object(session, Object::Areas(sheet, range), HostCall { member: "Item".into(), ..call })
                    }
                }
                "offset" | "resize" if read_kind(call.kind) => {
                    let args = arguments(&call, if member == "offset" { &["RowOffset", "ColumnOffset"] } else { &["RowSize", "ColumnSize"] })?;
                    let number = |i| args.get(i).filter(|v| !matches!(v, Value::Missing)).map(numeric_index).transpose();
                    let r = if member == "offset" {
                        range.offset(i64::from(number(0)?.unwrap_or(0)), i64::from(number(1)?.unwrap_or(0)))?
                    } else {
                        range.resize(number(0)?, number(1)?)?
                    };
                    self.intern(Object::Range(sheet, r))
                }
                "currentregion" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    let start = range.rectangle()?.start;
                    let doc = session.doc().map_err(|error| failure(91, error.to_string()))?;
                    let worksheet = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
                    let region = gridcraft_engine::selection::current_region(worksheet, start);
                    if region.count() > MAX_CELLS as u64 {
                        return Err(failure(7, "Range.CurrentRegion exceeds VBA cell budget"));
                    }
                    self.intern(Object::Range(sheet, region.into()))
                }
                "currentarray" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    let active = range.rectangle()?.start;
                    let doc = session.doc().map_err(|error| failure(91, error.to_string()))?;
                    let worksheet = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
                    let anchor = worksheet.spill_anchor(active).unwrap_or(active);
                    let array = worksheet.cell(anchor).and_then(|cell| cell.formula.as_ref()).and_then(|formula| formula.array);
                    let Some(array) = array else {
                        return Err(failure(1004, "Range.CurrentArray requires a cell in a legacy array formula"));
                    };
                    self.intern(Object::Range(sheet, array.into()))
                }
                "hasarray" | "hasspill" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    if range.count()? > MAX_CELLS as u64 {
                        return Err(failure(7, format!("Range.{} exceeds VBA cell budget", call.member)));
                    }
                    let doc = session.doc().map_err(|error| failure(91, error.to_string()))?;
                    let worksheet = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
                    let mut seen_true = false;
                    let mut seen_false = false;
                    for cell in range.areas.iter().flat_map(RangeRef::iter) {
                        let member = if member == "hasarray" {
                            let anchor = worksheet.spill_anchor(cell).unwrap_or(cell);
                            worksheet
                                .cell(anchor)
                                .and_then(|entry| entry.formula.as_ref())
                                .and_then(|formula| formula.array)
                                .is_some_and(|array| array.contains(cell))
                        } else {
                            worksheet.spill_ranges.values().any(|spill| spill.contains(cell))
                        };
                        seen_true |= member;
                        seen_false |= !member;
                    }
                    Ok(if seen_true && seen_false { Value::Null } else { Value::Boolean(seen_true) })
                }
                "spillparent" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    let cell = range.first()?.start;
                    let doc = session.doc().map_err(|error| failure(91, error.to_string()))?;
                    let worksheet = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
                    let anchor = worksheet
                        .spill_ranges
                        .iter()
                        .find(|(_, spill)| spill.contains(cell))
                        .map(|(anchor, _)| *anchor)
                        .ok_or_else(|| failure(1004, "Range.SpillParent requires a cell in a dynamic spill"))?;
                    self.intern(Object::Range(sheet, RangeRef::cell(anchor).into()))
                }
                "next" | "previous" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    let start = range.first()?.start;
                    let doc = session.doc().map_err(|error| failure(91, error.to_string()))?;
                    let worksheet = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
                    let forward = member == "next";
                    let mut candidate = start;
                    loop {
                        candidate = if forward {
                            candidate.offset(0, 1).ok_or_else(|| failure(1004, "Range.Next is outside the worksheet"))?
                        } else {
                            candidate.offset(0, -1).ok_or_else(|| failure(1004, "Range.Previous is outside the worksheet"))?
                        };
                        if worksheet.protection.is_none() || !doc.wb.styles.get(worksheet.style_id(candidate)).protection.locked {
                            break;
                        }
                    }
                    self.intern(Object::Range(sheet, RangeRef::cell(candidate).into()))
                }
                "item" => {
                    if !read_kind(call.kind) && call.kind != InvokeKind::Let {
                        return Err(failure(13, "Range.Item assignment requires Let"));
                    }
                    let write = call.kind == InvokeKind::Let;
                    let cell = Self::range_item(&range, &call, write)?;
                    if write {
                        let args = arguments(&call, &["RowIndex", "ColumnIndex", "Value"])?;
                        values::write(session, sheet, cell.rectangle()?, required(&args, 2)?, false, true)?;
                        Ok(Value::Empty)
                    } else {
                        self.intern(Object::Range(sheet, cell))
                    }
                }
                "worksheet" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    self.intern(Object::Sheet(sheet))
                }
                "cells" | "rows" | "columns" if read_kind(call.kind) => self.subranges(
                    sheet,
                    &range,
                    match member.as_str() {
                        "rows" => Unit::Rows,
                        "columns" => Unit::Columns,
                        _ => Unit::Cells,
                    },
                    &call,
                ),
                "row" | "column" | "count" | "countlarge" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    let first = range.first()?;
                    let count = match member.as_str() {
                        "row" => u64::from(first.start.row) + 1,
                        "column" => u64::from(first.start.col) + 1,
                        _ => range.count()?,
                    };
                    if member == "countlarge" {
                        Ok(Value::Double(count as f64))
                    } else {
                        Ok(Value::Long(i32::try_from(count).map_err(|_| failure(6, "Range.Count overflow"))?))
                    }
                }
                "value" | "value2" | "formula" => {
                    if read_kind(call.kind) {
                        arguments(&call, &[])?;
                        values::read(session, sheet, range.rectangle()?, member == "formula", member == "value")
                    } else if call.kind == InvokeKind::Let {
                        let args = arguments(&call, &["Value"])?;
                        values::write(session, sheet, range.rectangle()?, required(&args, 0)?, member == "formula", member == "value")?;
                        Ok(Value::Empty)
                    } else {
                        Err(failure(13, "Range value requires Let, not Set"))
                    }
                }
                "clearcontents" if call.kind == InvokeKind::Call => {
                    arguments(&call, &[])?;
                    self.clear_range(session, sheet, &range, "contents")?;
                    Ok(Value::Empty)
                }
                "clear" if call.kind == InvokeKind::Call => {
                    arguments(&call, &[])?;
                    self.clear_range(session, sheet, &range, "all")?;
                    Ok(Value::Empty)
                }
                "clearformats" if call.kind == InvokeKind::Call => {
                    arguments(&call, &[])?;
                    self.clear_range(session, sheet, &range, "formats")?;
                    Ok(Value::Empty)
                }
                "clearcomments" | "clearnotes" if call.kind == InvokeKind::Call => {
                    arguments(&call, &[])?;
                    self.clear_range(session, sheet, &range, "comments")?;
                    Ok(Value::Empty)
                }
                "clearhyperlinks" if call.kind == InvokeKind::Call => {
                    arguments(&call, &[])?;
                    self.clear_range(session, sheet, &range, "hyperlinks")?;
                    Ok(Value::Empty)
                }
                "address" if read_kind(call.kind) => self.address(session, sheet, &range, &call),
                _ => Err(failure(438, format!("Range.{} is not a supported member", call.member))),
            },
            _ => Err(failure(438, format!("{receiver}.{} is not available for this invocation kind", call.member))),
        }
    }
}
