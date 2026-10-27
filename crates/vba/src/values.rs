use super::*;
use gridcraft_engine::{
    CellContent,
    core::{CellError, RangeRef, Value as CellValue},
};
use vba_values::{ScalarKind, VbaArray, VbaString};

pub(crate) fn text(value: &Value) -> HostResult<String> {
    match value {
        Value::String(s) => String::from_utf16(s.units()).map_err(|_| failure(13, "worksheet text cannot contain unpaired UTF-16 surrogates")),
        Value::Null => Err(failure(94, "invalid use of Null")),
        _ => Err(failure(13, "string required")),
    }
}
pub(crate) fn vba_text(s: &str) -> HostResult<Value> {
    VbaString::from_utf8(s).map(Value::String).map_err(|e| failure(7, e.to_string()))
}
pub(crate) fn dimensions(range: RangeRef) -> HostResult<(usize, usize)> {
    let rows = range.end.row.checked_sub(range.start.row).and_then(|v| v.checked_add(1)).ok_or_else(|| failure(1004, "invalid range"))? as usize;
    let cols = range.end.col.checked_sub(range.start.col).and_then(|v| v.checked_add(1)).ok_or_else(|| failure(1004, "invalid range"))? as usize;
    if rows.checked_mul(cols).is_none_or(|n| n > MAX_CELLS) {
        return Err(failure(7, "range exceeds VBA cell budget"));
    }
    Ok((rows, cols))
}
const ERRORS: &[(u16, CellError)] = &[
    (2000, CellError::Null),
    (2007, CellError::Div0),
    (2015, CellError::Value),
    (2023, CellError::Ref),
    (2029, CellError::Name),
    (2036, CellError::Num),
    (2042, CellError::NA),
    (2043, CellError::GettingData),
];
fn from_cell(value: CellValue) -> HostResult<Value> {
    match value {
        CellValue::Empty => Ok(Value::Empty),
        CellValue::Number(n) => Ok(Value::Double(n)),
        CellValue::Bool(b) => Ok(Value::Boolean(b)),
        CellValue::Text(s) => vba_text(&s),
        CellValue::Error(e) => ERRORS
            .iter()
            .find(|(_, error)| *error == e)
            .map(|(n, _)| Value::CvErr(*n))
            .ok_or_else(|| failure(13, "GridCraft-specific error has no supported VBA CVErr")),
        CellValue::Array(_) => Err(failure(13, "nested cell array unsupported")),
    }
}
pub(crate) fn read(session: &Session, sheet: usize, range: RangeRef, formula: bool) -> HostResult<Value> {
    let (rows, cols) = dimensions(range)?;
    let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
    let mut values = Vec::with_capacity(rows * cols);
    let mut payload = 0usize;
    for at in range.iter() {
        let value = if formula {
            if let Some(f) = sh.cell(at).and_then(|c| c.formula.as_ref()) {
                vba_text(&format!("={}", f.text.trim_start_matches('=')))?
            } else {
                from_cell(sh.value(at))?
            }
        } else {
            from_cell(sh.value(at))?
        };
        payload = payload.saturating_add(64).saturating_add(value.payload_bytes());
        if payload > 8 * 1024 * 1024 {
            return Err(failure(7, "range payload exceeds VBA budget"));
        }
        values.push(value);
    }
    if range.is_single() {
        return values.into_iter().next().ok_or_else(|| failure(1004, "empty range"));
    }
    let mut array = VbaArray::new(ScalarKind::Variant, vec![(1, rows as i32), (1, cols as i32)], false).map_err(|e| failure(7, e.to_string()))?;
    for (index, value) in values.into_iter().enumerate() {
        array.set(index, value).map_err(|e| failure(7, e.to_string()))?;
    }
    Ok(Value::Array(array))
}
// Conversion happens only at the workbook storage/engine-command boundary. VBA values
// retain their exact tags in the VM and host protocol; a worksheet stores IEEE doubles.
fn content(value: &Value, formula: bool) -> HostResult<CellContent> {
    let number = match value {
        Value::Empty => return Ok(CellContent::Value(CellValue::Empty)),
        Value::Boolean(b) => return Ok(CellContent::Value(CellValue::Bool(*b))),
        Value::Byte(n) => f64::from(*n),
        Value::Integer(n) => f64::from(*n),
        Value::Long(n) => f64::from(*n),
        Value::LongLong(n) if (-9_007_199_254_740_992..=9_007_199_254_740_992).contains(n) => *n as f64,
        Value::LongLong(_) => return Err(failure(6, "integer cannot be stored exactly in a worksheet number")),
        Value::Single(n) => f64::from(*n),
        Value::Double(n) | Value::Date(n) => *n,
        Value::Currency(n) => *n as f64 / 10000.0,
        Value::String(_) => {
            let s = text(value)?;
            return Ok(if formula && s.starts_with('=') { CellContent::Formula(s) } else { CellContent::Value(CellValue::text(s)) });
        }
        Value::CvErr(n) => {
            return ERRORS
                .iter()
                .find(|(code, _)| code == n)
                .map(|(_, e)| CellContent::Value(CellValue::Error(*e)))
                .ok_or_else(|| failure(13, "unsupported worksheet CVErr"));
        }
        Value::Null => return Err(failure(94, "invalid use of Null")),
        Value::Missing => return Err(failure(449, "argument not optional")),
        _ => return Err(failure(13, "value cannot be stored in a worksheet")),
    };
    if !number.is_finite() {
        return Err(failure(6, "non-finite worksheet number"));
    }
    Ok(CellContent::Value(CellValue::Number(number)))
}
pub(crate) fn write(session: &mut Session, sheet: usize, range: RangeRef, value: &Value, formula: bool) -> HostResult<()> {
    let (rows, cols) = dimensions(range)?;
    if let Value::Array(_) = value {
        value.validate().map_err(|e| failure(13, e.to_string()))?;
    }
    if !matches!(value, Value::Array(_))
        && rows.checked_mul(cols).and_then(|n| n.checked_mul(value.payload_bytes().saturating_add(64))).is_none_or(|n| n > 8 * 1024 * 1024)
    {
        return Err(failure(7, "range payload exceeds VBA budget"));
    }
    let array = if let Value::Array(a) = value {
        let shape: Vec<usize> =
            a.bounds().iter().map(|(lo, hi)| i64::from(*hi) - i64::from(*lo) + 1).map(|n| usize::try_from(n).unwrap_or(0)).collect();
        if shape != [rows, cols] {
            return Err(failure(13, "array must have two dimensions matching the destination"));
        }
        Some(a.elements())
    } else {
        None
    };
    let mut data = Vec::with_capacity(rows * cols);
    // Validate every conversion before the engine commits any part of the range.
    for row in 0..rows {
        for col in 0..cols {
            let element = match array {
                Some(a) => a.get(row * cols + col).ok_or_else(|| failure(9, "invalid array storage"))?,
                None => value,
            };
            data.push(content(element, formula)?);
        }
    }
    session.set_range_contents(sheet, range, &data).map_err(|e| failure(1004, e.to_string()))?;
    Ok(())
}
