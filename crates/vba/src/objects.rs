use super::*;
use gridcraft_engine::core::{CellRef, MAX_COLS, MAX_ROWS, RangeRef};
use vba_foreign::HostArgument;

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Object {
    Application,
    Workbook,
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
                names.iter().position(|n| n.eq_ignore_ascii_case(name)).ok_or_else(|| failure(448, "unknown named argument"))?
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

impl Bridge {
    /// Resolve only local rectangular references. Range endpoints are host handles,
    /// not serialized addresses; validate their owner before inspecting the object.
    fn endpoint(&self, sheet: usize, value: &Value) -> HostResult<RangeRef> {
        match value {
            Value::Object(reference) => {
                if reference.host != self.identity || reference.generation != self.doc {
                    return Err(failure(91, "stale or foreign range endpoint"));
                }
                let index = usize::try_from(reference.object).map_err(|_| failure(91, "invalid range endpoint"))?;
                match self.objects.get(index) {
                    Some(Object::Range(owner, range)) if *owner == sheet => Ok(*range),
                    Some(Object::Range(..)) => Err(failure(1004, "range endpoint belongs to another worksheet")),
                    Some(_) => Err(failure(13, "range endpoint must be a Range")),
                    None => Err(failure(91, "invalid range endpoint")),
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
    pub(crate) fn invoke_object(&mut self, session: &mut Session, object: Object, call: HostCall) -> HostResult<Value> {
        if call.member.len() > 256 || call.arguments.len() > 256 {
            return Err(failure(7, "host call budget exceeded"));
        }
        let member = call.member.to_ascii_lowercase();
        if read_kind(call.kind) && member == "application" {
            arguments(&call, &[])?;
            return self.intern(Object::Application);
        }
        if read_kind(call.kind) && member == "parent" {
            arguments(&call, &[])?;
            return self.intern(match object {
                Object::Application | Object::Workbook => Object::Application,
                Object::Sheet(_) => Object::Workbook,
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
                        let args = arguments(&call, &["Index"])?;
                        let value = required(&args, 0)?;
                        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
                        let sheet = if matches!(value, Value::String(_)) {
                            doc.wb.sheet_index(&string(value)?).ok_or_else(|| failure(9, "worksheet not found"))?
                        } else {
                            coordinate(value, u32::try_from(doc.wb.sheets.len()).map_err(|_| failure(7, "worksheet budget"))?)? as usize
                        };
                        self.intern(Object::Sheet(sheet))
                    }
                    "range" | "cells" if object == Object::Application => self.invoke_object(session, Object::Sheet(active), call),
                    "name" if object == Object::Workbook => {
                        arguments(&call, &[])?;
                        values::vba_text(&session.doc().map_err(|e| failure(91, e.to_string()))?.title)
                    }
                    _ => Err(failure(438, "unsupported Application/Workbook member")),
                }
            }
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
