//! Address rendering never materializes worksheet cells.
use super::objects::{arguments, numeric_index};
use super::*;
use gridcraft_engine::core::{CellRef, col_to_letters};

pub(crate) fn boolean(value: &Value) -> HostResult<bool> {
    match vba_runtime::bind(vba_values::ScalarKind::Boolean, value).map_err(|e| match e {
        vba_runtime::Error::Vba { number, message } => failure(number, message),
        vba_runtime::Error::Resource(message) => failure(7, message),
        vba_runtime::Error::Unsupported(message) => failure(13, message),
    })? {
        Value::Boolean(b) => Ok(b),
        _ => Err(failure(13, "Boolean required")),
    }
}
impl Bridge {
    pub(crate) fn address(&self, session: &Session, sheet: usize, range: &ranges::Range, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["RowAbsolute", "ColumnAbsolute", "ReferenceStyle", "External", "RelativeTo"])?;
        let flag = |i: usize, default| args.get(i).filter(|v| !matches!(v, Value::Missing)).map(boolean).transpose().map(|v| v.unwrap_or(default));
        let row = flag(0, true)?;
        let col = flag(1, true)?;
        let external = flag(3, false)?;
        let style = args.get(2).filter(|v| !matches!(v, Value::Missing)).map(numeric_index).transpose()?.unwrap_or(1);
        if !matches!(style, 1 | -4150) {
            return Err(failure(1004, "unsupported ReferenceStyle"));
        }
        let origin = match args.get(4) {
            None | Some(Value::Missing) => CellRef::default(),
            Some(Value::Object(reference)) => match self.object(session, *reference)? {
                Object::Range(owner, r) if owner == sheet => r.rectangle()?.start,
                _ => return Err(failure(1004, "RelativeTo must be a Range on the same worksheet")),
            },
            _ => return Err(failure(13, "RelativeTo must be a Range")),
        };
        let cell = |at: CellRef| {
            if style == 1 {
                format!("{}{}{}{}", if col { "$" } else { "" }, col_to_letters(at.col), if row { "$" } else { "" }, at.row + 1)
            } else {
                let component = |label: &str, value: u32, absolute, base: u32| {
                    if absolute {
                        format!("{label}{}", value + 1)
                    } else if value == base {
                        label.to_string()
                    } else {
                        format!("{label}[{}]", i64::from(value) - i64::from(base))
                    }
                };
                format!("{}{}", component("R", at.row, row, origin.row), component("C", at.col, col, origin.col))
            }
        };
        range.validate()?;
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let prefix = if external {
            let name = &doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?.name;
            let owner = format!("[{}]{}", doc.title, name);
            if owner.chars().all(|c| c.is_ascii_alphanumeric() || matches!(c, '_' | '.' | '[' | ']')) {
                format!("{owner}!")
            } else {
                format!("'{}'!", owner.replace('\'', "''"))
            }
        } else {
            String::new()
        };
        let mut areas = Vec::new();
        for r in &range.areas {
            let address = if r.is_full_cols() {
                if style == 1 {
                    format!(
                        "{}{}:{}{}",
                        if col { "$" } else { "" },
                        col_to_letters(r.start.col),
                        if col { "$" } else { "" },
                        col_to_letters(r.end.col)
                    )
                } else {
                    let a = cell(r.start);
                    let b = cell(r.end);
                    format!(
                        "{}:{}",
                        a.split_once('C').map(|(_, c)| format!("C{c}")).unwrap_or_default(),
                        b.split_once('C').map(|(_, c)| format!("C{c}")).unwrap_or_default()
                    )
                }
            } else if r.is_full_rows() {
                if style == 1 {
                    format!("{}{}:{}{}", if row { "$" } else { "" }, r.start.row + 1, if row { "$" } else { "" }, r.end.row + 1)
                } else {
                    let a = cell(r.start);
                    let b = cell(r.end);
                    format!("{}:{}", a.split('C').next().unwrap_or_default(), b.split('C').next().unwrap_or_default())
                }
            } else if r.is_single() {
                cell(r.start)
            } else {
                format!("{}:{}", cell(r.start), cell(r.end))
            };
            areas.push(format!("{prefix}{address}"));
        }
        values::vba_text(&areas.join(","))
    }
}
