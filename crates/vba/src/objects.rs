use super::*;
use gridcraft_engine::core::{CellRef, MAX_COLS, MAX_ROWS, RangeRef};
use vba_foreign::HostArgument;

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Object {
    Application,
    Workbook,
    Worksheets,
    Sheets,
    Sheet(usize),
    Range(usize, RangeRef),
}

/// Bind names before dispatch; never silently drop named/omitted/extra arguments.
fn arguments(call: &HostCall, names: &[&str]) -> HostResult<Vec<Value>> {
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
fn required(args: &[Value], index: usize) -> HostResult<&Value> {
    args.get(index).filter(|v| !matches!(v, Value::Missing)).ok_or_else(|| failure(449, "argument not optional"))
}
fn string(value: &Value) -> HostResult<String> {
    super::values::text(value)
}
fn coordinate(value: &Value, max: u32) -> HostResult<u32> {
    let number = match value {
        Value::Byte(n) => i64::from(*n),
        Value::Integer(n) => i64::from(*n),
        Value::Long(n) => i64::from(*n),
        Value::LongLong(n) => *n,
        _ => return Err(failure(13, "index must be an integer")),
    };
    if number < 1 || number > i64::from(max) {
        return Err(failure(1004, "cell index outside worksheet"));
    }
    u32::try_from(number - 1).map_err(|_| failure(6, "index overflow"))
}
fn read_kind(kind: InvokeKind) -> bool {
    matches!(kind, InvokeKind::Call | InvokeKind::Get)
}
fn numeric_index(value: &Value) -> HostResult<i32> {
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
    /// Resolve only local rectangular references. Range endpoints are host handles,
    /// not serialized addresses; validate their owner before inspecting the object.
    fn endpoint(&self, sheet: usize, value: &Value) -> HostResult<RangeRef> {
        match value {
            Value::Object(reference) => {
                if reference.host != self.identity || reference.generation != self.doc {
                    return Err(failure(91, "stale or foreign range endpoint"));
                }
                match self.objects.get(*reference)? {
                    Object::Range(owner, range) if owner == sheet => Ok(range),
                    Object::Range(..) => Err(failure(1004, "range endpoint belongs to another worksheet")),
                    _ => Err(failure(13, "range endpoint must be a Range")),
                }
            }
            _ => {
                let address = string(value)?;
                if address.len() > 64 || address.contains(['!', ',']) {
                    return Err(failure(1004, "only a local rectangular A1 range is supported"));
                }
                RangeRef::parse(&address).ok_or_else(|| failure(1004, "invalid range"))
            }
        }
    }
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
    fn range_item(range: RangeRef, call: &HostCall, write: bool) -> HostResult<RangeRef> {
        let args = arguments(call, if write { &["RowIndex", "ColumnIndex", "Value"] } else { &["RowIndex", "ColumnIndex"] })?;
        let index = |v: &Value| -> HostResult<i64> {
            let n = numeric_index(v)?;
            Ok(i64::from(n) - 1)
        };
        let first = index(required(&args, 0)?)?;
        let (row, col) = if let Some(second) = args.get(1).filter(|v| !matches!(v, Value::Missing)) {
            let col = if let Value::String(_) = second {
                let text = string(second)?;
                if text.is_empty() || text.len() > 3 || !text.bytes().all(|b| b.is_ascii_alphabetic()) {
                    return Err(failure(1004, "invalid column index"));
                }
                let mut n = 0i64;
                for b in text.bytes() {
                    n = n * 26 + i64::from(b.to_ascii_uppercase() - b'A' + 1);
                }
                n - 1
            } else {
                index(second)?
            };
            (first, col)
        } else {
            let width = i64::from(range.end.col - range.start.col + 1);
            (first.div_euclid(width), first.rem_euclid(width))
        };
        let row = i64::from(range.start.row) + row;
        let col = i64::from(range.start.col) + col;
        if row < 0 || col < 0 || row >= i64::from(MAX_ROWS) || col >= i64::from(MAX_COLS) {
            return Err(failure(1004, "Range.Item outside worksheet"));
        }
        Ok(RangeRef::cell(CellRef::new(
            u32::try_from(row).map_err(|_| failure(6, "row overflow"))?,
            u32::try_from(col).map_err(|_| failure(6, "column overflow"))?,
        )))
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
            Object::Range(sheet, range) => {
                values::dimensions(range)?;
                let n = range.iter().count();
                if n > MAX_OBJECTS {
                    return Err(failure(7, "range enumeration budget"));
                }
                range.iter().map(|c| Object::Range(sheet, RangeRef::cell(c))).collect()
            }
            _ => return Err(failure(438, "object is not enumerable")),
        };
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
                call.member = match object {
                    Object::Range(..) if call.arguments.is_empty() => "Value",
                    Object::Range(..) | Object::Worksheets | Object::Sheets => "Item",
                    _ => return Err(failure(438, "object has no default property")),
                }
                .into();
            }
            InvokeKind::DefaultLet => {
                call.kind = InvokeKind::Let;
                call.member = match object {
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
            && matches!(member.as_str(), "range" | "cells")
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
            let Value::Object(reference) = self.invoke_object(session, object, call)? else {
                return Err(failure(13, "Range required"));
            };
            let Object::Range(sheet, range) = self.objects.get(reference)? else {
                return Err(failure(13, "Range required"));
            };
            values::write(session, sheet, range, &value, false)?;
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
                Object::Sheet(_) | Object::Worksheets | Object::Sheets => Object::Workbook,
                Object::Range(sheet, _) => Object::Sheet(sheet),
            });
        }
        match object {
            Object::Application | Object::Workbook if read_kind(call.kind) => {
                let active = session.doc().map_err(|e| failure(91, e.to_string()))?.sheet_index();
                match member.as_str() {
                    "thisworkbook" | "activeworkbook" if object == Object::Application => {
                        arguments(&call, &[])?;
                        self.intern(Object::Workbook)
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
                    "range" | "cells" if object == Object::Application => self.invoke_object(session, Object::Sheet(active), call),
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
                    values::vba_text(self.sheets.get(sheet).ok_or_else(|| failure(91, "worksheet gone"))?)
                }
                "range" => {
                    let args = arguments(&call, &["Cell1", "Cell2"])?;
                    let first = required(&args, 0)?;
                    let mut range = self.endpoint(sheet, first)?;
                    if let Some(second) = args.get(1).filter(|v| !matches!(v, Value::Missing)) {
                        let other = self.endpoint(sheet, second)?;
                        range = RangeRef {
                            start: CellRef::new(range.start.row.min(other.start.row), range.start.col.min(other.start.col)),
                            end: CellRef::new(range.end.row.max(other.end.row), range.end.col.max(other.end.col)),
                        };
                    } else if matches!(first, Value::Object(_)) {
                        return Err(failure(1004, "single-argument Worksheet.Range requires an address"));
                    }
                    values::dimensions(range)?;
                    self.intern(Object::Range(sheet, range))
                }
                "cells" => {
                    let args = arguments(&call, &["Row", "Column"])?;
                    let at = CellRef::new(coordinate(required(&args, 0)?, MAX_ROWS)?, coordinate(required(&args, 1)?, MAX_COLS)?);
                    self.intern(Object::Range(sheet, RangeRef::cell(at)))
                }
                _ => Err(failure(438, "unsupported Worksheet member")),
            },
            Object::Range(sheet, range) => match member.as_str() {
                "item" => {
                    if !read_kind(call.kind) && call.kind != InvokeKind::Let {
                        return Err(failure(13, "Range.Item assignment requires Let"));
                    }
                    let write = call.kind == InvokeKind::Let;
                    let cell = Self::range_item(range, &call, write)?;
                    if write {
                        let args = arguments(&call, &["RowIndex", "ColumnIndex", "Value"])?;
                        values::write(session, sheet, cell, required(&args, 2)?, false)?;
                        Ok(Value::Empty)
                    } else {
                        self.intern(Object::Range(sheet, cell))
                    }
                }
                "worksheet" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    self.intern(Object::Sheet(sheet))
                }
                "row" | "column" | "count" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    let number = match member.as_str() {
                        "row" => u64::from(range.start.row) + 1,
                        "column" => u64::from(range.start.col) + 1,
                        _ => u64::from(range.end.row - range.start.row + 1) * u64::from(range.end.col - range.start.col + 1),
                    };
                    Ok(Value::Long(i32::try_from(number).map_err(|_| failure(6, "Range.Count overflow"))?))
                }
                "value" | "value2" | "formula" => {
                    if read_kind(call.kind) {
                        arguments(&call, &[])?;
                        values::read(session, sheet, range, member == "formula")
                    } else if call.kind == InvokeKind::Let {
                        let args = arguments(&call, &["Value"])?;
                        values::write(session, sheet, range, required(&args, 0)?, member == "formula")?;
                        Ok(Value::Empty)
                    } else {
                        Err(failure(13, "Range value requires Let, not Set"))
                    }
                }
                "clearcontents" if call.kind == InvokeKind::Call => {
                    arguments(&call, &[])?;
                    values::write(session, sheet, range, &Value::Empty, false)?;
                    Ok(Value::Empty)
                }
                "address" if read_kind(call.kind) => {
                    arguments(&call, &[])?;
                    let absolute = |c: CellRef| format!("${}${}", gridcraft_engine::core::col_to_letters(c.col), c.row + 1);
                    values::vba_text(&if range.is_single() {
                        absolute(range.start)
                    } else {
                        format!("{}:{}", absolute(range.start), absolute(range.end))
                    })
                }
                _ => Err(failure(438, "unsupported Range member")),
            },
            _ => Err(failure(438, "unsupported property assignment")),
        }
    }
}
