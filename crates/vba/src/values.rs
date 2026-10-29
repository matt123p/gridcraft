use super::*;
use gridcraft_engine::{
    CellContent,
    core::{CellError, RangeRef, Value as CellValue},
};
use vba_values::{ScalarKind, VbaArray, VbaString};

pub(crate) fn runtime_error(error: vba_runtime::Error) -> ForeignError {
    match error {
        vba_runtime::Error::Vba { number, message } => failure(number, message),
        vba_runtime::Error::Resource(message) => failure(7, message),
        vba_runtime::Error::Unsupported(message) => failure(13, message),
    }
}
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
pub(crate) fn from_cell(value: CellValue) -> HostResult<Value> {
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
pub(crate) fn read(session: &Session, sheet: usize, range: RangeRef, formula: bool, typed: bool) -> HostResult<Value> {
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
        let value = if typed && let Value::Double(number) = value {
            let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
            let format = gridcraft_engine::numfmt::NumberFormat::parse(doc.wb.styles.get(sh.style_id(at)).num_fmt.as_str());
            if format.is_date() {
                let number = if doc.wb.date_system == gridcraft_engine::core::date::DateSystem::D1904 {
                    number + 1462.0
                } else if (0.0..60.0).contains(&number) {
                    number + 1.0
                } else if (60.0..61.0).contains(&number) {
                    return Err(failure(1004, "Excel's fictitious leap day has no VBA Date"));
                } else {
                    number
                };
                vba_runtime::bind(ScalarKind::Date, &Value::Double(number)).map_err(runtime_error)?
            } else if matches!(format.kind(), gridcraft_engine::numfmt::FormatKind::Currency | gridcraft_engine::numfmt::FormatKind::Accounting) {
                vba_runtime::bind(ScalarKind::Currency, &Value::Double(number)).map_err(runtime_error)?
            } else {
                Value::Double(number)
            }
        } else {
            value
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
pub(crate) fn content(value: &Value, formula: bool) -> HostResult<CellContent> {
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
pub(crate) fn write(session: &mut Session, sheet: usize, range: RangeRef, value: &Value, formula: bool, typed: bool) -> HostResult<()> {
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
        if shape.len() != 2 || shape.contains(&0) {
            return Err(failure(13, "two-dimensional array required"));
        }
        Some((a.elements(), shape))
    } else {
        None
    };
    let mut data = Vec::with_capacity(rows * cols);
    let mut formats = Vec::with_capacity(rows * cols);
    let mut payload = 0usize;
    // Validate every conversion before the engine commits any part of the range.
    for row in 0..rows {
        for col in 0..cols {
            let missing = Value::CvErr(2042);
            let element = match &array {
                Some((a, shape)) if row < shape.first().copied().unwrap_or(0) && col < shape.get(1).copied().unwrap_or(0) => {
                    a.get(row * shape.get(1).copied().unwrap_or(0) + col).ok_or_else(|| failure(9, "invalid array storage"))?
                }
                Some(_) => &missing,
                None => value,
            };
            payload = payload.saturating_add(64).saturating_add(element.payload_bytes());
            if payload > 8 * 1024 * 1024 {
                return Err(failure(7, "range payload exceeds VBA budget"));
            }
            let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
            let at = gridcraft_engine::core::CellRef::new(range.start.row + row as u32, range.start.col + col as u32);
            let sheet = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
            let general = doc.wb.styles.get(sheet.style_id(at)).num_fmt.as_str() == "General";
            let (stored, format) = if typed {
                match element {
                    Value::Date(n) => {
                        let number = if doc.wb.date_system == gridcraft_engine::core::date::DateSystem::D1904 {
                            *n - 1462.0
                        } else if (1.0..61.0).contains(n) {
                            *n - 1.0
                        } else {
                            *n
                        };
                        (Value::Double(number), general.then(|| "m/d/yyyy".to_string()))
                    }
                    Value::Currency(n) => {
                        let cents = n / 100 + if (n % 100).abs() >= 50 { n.signum() } else { 0 };
                        (Value::Double(cents as f64 / 100.0), general.then(|| "$#,##0.00".to_string()))
                    }
                    _ => (element.clone(), None),
                }
            } else {
                (element.clone(), None)
            };
            let (cell, format) = if let Value::String(_) = &stored {
                let text = text(&stored)?;
                let text_format = gridcraft_engine::numfmt::NumberFormat::parse(doc.wb.styles.get(sheet.style_id(at)).num_fmt.as_str()).is_text();
                if text.starts_with('=') && (formula || !text_format) {
                    (CellContent::Formula(text), format)
                } else if text_format {
                    (content(&stored, false)?, format)
                } else {
                    let parsed = gridcraft_engine::core::parse::parse_input(&text, doc.wb.date_system);
                    (CellContent::Value(parsed.value), if general { parsed.format.map(str::to_string).or(format) } else { format })
                }
            } else {
                (content(&stored, formula)?, format)
            };
            formats.push(format);
            data.push(cell);
        }
    }
    session.set_range_contents_with_formats(sheet, range, &data, &formats).map_err(|e| failure(1004, e.to_string()))?;
    Ok(())
}
