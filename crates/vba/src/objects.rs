use super::ranges::{Range, Unit};
use super::*;
use gridcraft_engine::core::RangeRef;
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
                let active = session.doc().map_err(|e| failure(91, e.to_string()))?.sheet_index();
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
                    "range" if object == Object::Application => self.construct_range(session, active, None, true, &call),
                    "cells" | "rows" | "columns" if object == Object::Application => self.invoke_object(session, Object::Sheet(active), call),
                    "name" if object == Object::Workbook => {
                        arguments(&call, &[])?;
                        values::vba_text(&session.doc().map_err(|e| failure(91, e.to_string()))?.title)
                    }
                    _ => Err(failure(438, "unsupported Application/Workbook member")),
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
                _ => Err(failure(438, "unsupported collection member")),
            },
            Object::Sheet(sheet) if read_kind(call.kind) => match member.as_str() {
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
                _ => Err(failure(438, "unsupported Worksheet member")),
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
                _ => Err(failure(438, "unsupported Areas member")),
            },
            Object::Range(sheet, range) => match member.as_str() {
                "font" | "interior" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    self.intern(Object::Format(sheet, range, member == "font"))
                }
                "numberformat" => self.number_format(session, sheet, &range, &call),
                "end" if read_kind(call.kind) => self.end(session, sheet, &range, &call),
                "find" | "findnext" | "findprevious" if read_kind(call.kind) => self.find(session, sheet, &range, &call),
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
                    values::write(session, sheet, range.rectangle()?, &Value::Empty, false, false)?;
                    Ok(Value::Empty)
                }
                "address" if read_kind(call.kind) => self.address(session, sheet, &range, &call),
                _ => Err(failure(438, "unsupported Range member")),
            },
            _ => Err(failure(438, "unsupported property assignment")),
        }
    }
}
