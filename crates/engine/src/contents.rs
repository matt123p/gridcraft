//! Explicit cell contents for native hosts and the equivalent recorded command.
//! Kept separate from typed input: text never becomes a number, date or formula.
use crate::{EngineError, Result, Session, cmd};
use gridcraft_core::{RangeRef, Value};
use gridcraft_model::{Cell, Formula};
use serde::{Deserialize, Serialize};
use serde_json::{Value as Json, json};

pub const MAX_CONTENT_CELLS: usize = 65_536;
pub const MAX_CONTENT_BYTES: usize = 8 * 1024 * 1024;
const COMMAND: &str = "range.setContents";

/// One scalar constant or an explicit formula; no host-language or Office types.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum CellContent {
    Value(Value),
    Formula(String),
}

fn bad(message: impl Into<String>) -> EngineError {
    EngineError::BadParams { cmd: COMMAND.into(), msg: message.into() }
}

fn formula_bytes(text: &str) -> Result<usize> {
    let max = gridcraft_formula::MAX_FORMULA_LEN;
    let body = text.strip_prefix('=').unwrap_or(text);
    if body.is_empty() || body.len() > max * 4 || (body.len() > max && body.chars().count() > max) {
        return Err(bad("formula body must contain 1–8192 characters"));
    }
    Ok(text.len())
}

fn validate(range: RangeRef, contents: &[CellContent]) -> Result<()> {
    if !range.start.is_valid() || !range.end.is_valid() || range.start.row > range.end.row || range.start.col > range.end.col {
        return Err(bad("invalid destination rectangle"));
    }
    if range.count() > MAX_CONTENT_CELLS as u64 || contents.len() != range.count() as usize {
        return Err(bad("contents must exactly fill the rectangle, at most 65536 cells"));
    }
    let mut bytes = 0usize;
    for content in contents {
        let payload = match content {
            CellContent::Value(Value::Number(n)) if !n.is_finite() => return Err(bad("non-finite number")),
            CellContent::Value(Value::Array(_)) => return Err(bad("cell constants must be scalar")),
            CellContent::Value(Value::Text(s)) => s.len(),
            CellContent::Formula(s) => formula_bytes(s)?,
            _ => 16,
        };
        bytes = bytes.saturating_add(64).saturating_add(payload);
        if bytes > MAX_CONTENT_BYTES {
            return Err(bad("cell content payload exceeds 8 MiB"));
        }
    }
    Ok(())
}

impl Session {
    /// Write a bounded row-major rectangle through the ordinary command lifecycle.
    /// Constants retain their types even in Text-formatted cells; formulas are explicit.
    /// Recording/journaling serialize the command after success, never for dispatch.
    pub fn set_range_contents(&mut self, sheet: usize, range: RangeRef, contents: &[CellContent]) -> Result<Json> {
        self.set_range_contents_with_formats(sheet, range, contents, &[])
    }
    /// Optional per-cell number formats commit atomically with values.
    pub fn set_range_contents_with_formats(
        &mut self,
        sheet: usize,
        range: RangeRef,
        contents: &[CellContent],
        formats: &[Option<String>],
    ) -> Result<Json> {
        self.execute_with(
            COMMAND,
            || std::borrow::Cow::Owned(json!({"sheet":sheet,"range":range.a1(),"contents":contents,"formats":formats})),
            |session, _| apply(session, sheet, range, contents, formats),
        )
    }
}

pub(crate) fn run(session: &mut Session, p: &Json) -> Result<Json> {
    let range_text = p.get("range").and_then(Json::as_str).ok_or_else(|| bad("missing range"))?;
    if range_text.len() > 512 || range_text.contains(',') {
        return Err(bad("one rectangular range is required"));
    }
    let range = cmd::target_range(session, p)?;
    let sheet = match p.get("sheet") {
        Some(Json::Number(n)) => {
            usize::try_from(n.as_u64().ok_or_else(|| bad("sheet must be a nonnegative index"))?).map_err(|_| bad("sheet index overflow"))?
        }
        Some(Json::String(name)) if name.len() > 128 => return Err(bad("sheet name exceeds address budget")),
        None | Some(Json::Null) | Some(Json::String(_)) => cmd::target_sheet(session, p)?,
        _ => return Err(bad("sheet must be an index or name")),
    };
    if let (Some(name), _) = gridcraft_core::split_sheet(range_text)
        && session.doc()?.wb.sheet_index(&name) != Some(sheet)
    {
        return Err(bad("range and sheet identify different worksheets"));
    }
    let raw = p.get("contents").and_then(Json::as_array).ok_or_else(|| bad("contents must be a row-major array"))?;
    // Bound shape and serialized payload before deserializing untrusted data.
    if raw.len() > MAX_CONTENT_CELLS || range.count() != raw.len() as u64 {
        return Err(bad("contents shape or cell budget exceeded"));
    }
    let mut contents = Vec::with_capacity(raw.len());
    let mut bytes = 0usize;
    for item in raw {
        // Reject recursive Array values before serde can allocate their nested payloads.
        let object = item.as_object().ok_or_else(|| bad("content must have exactly one value or formula field"))?;
        if object.len() != 1 {
            return Err(bad("content must have exactly one value or formula field"));
        }
        bytes = bytes.saturating_add(64);
        let content = if let Some(text) = object.get("formula").and_then(Json::as_str) {
            bytes = bytes.saturating_add(formula_bytes(text)?);
            if bytes > MAX_CONTENT_BYTES {
                return Err(bad("cell content payload exceeds 8 MiB"));
            }
            CellContent::Formula(text.to_string())
        } else if let Some(value) = object.get("value").and_then(Json::as_object) {
            let tag = value.get("t").and_then(Json::as_str).ok_or_else(|| bad("missing scalar value tag"))?;
            if !["Empty", "Number", "Text", "Bool", "Error"].contains(&tag) {
                return Err(bad("unsupported scalar value tag"));
            }
            if value.len() > 2 || value.keys().any(|k| k != "t" && k != "v") {
                return Err(bad("unknown scalar value field"));
            }
            let payload = value.get("v");
            if tag != "Text" {
                bytes = bytes.saturating_add(16);
            }
            let scalar = match (tag, payload) {
                ("Empty", None) => Value::Empty,
                ("Number", Some(Json::Number(n))) => Value::Number(n.as_f64().ok_or_else(|| bad("invalid number"))?),
                ("Bool", Some(Json::Bool(b))) => Value::Bool(*b),
                ("Text", Some(Json::String(s))) => {
                    bytes = bytes.saturating_add(s.len());
                    if bytes > MAX_CONTENT_BYTES {
                        return Err(bad("cell content payload exceeds 8 MiB"));
                    }
                    Value::text(s.as_str())
                }
                ("Error", Some(Json::String(s))) => {
                    if s.len() > 32 {
                        return Err(bad("invalid cell error tag"));
                    }
                    // Serde's unit error enum is bounded; no recursively shaped values.
                    Value::Error(serde_json::from_value(Json::String(s.clone())).map_err(|e| bad(e.to_string()))?)
                }
                _ => return Err(bad("scalar payload does not match its tag")),
            };
            if bytes > MAX_CONTENT_BYTES {
                return Err(bad("cell content payload exceeds 8 MiB"));
            }
            CellContent::Value(scalar)
        } else {
            return Err(bad("invalid value or formula content"));
        };
        contents.push(content);
    }
    let formats = match p.get("formats") {
        None => Vec::new(),
        Some(Json::Array(raw)) if raw.len() <= MAX_CONTENT_CELLS => {
            let mut formats = Vec::with_capacity(raw.len());
            let mut bytes = 0usize;
            for format in raw {
                let format = match format {
                    Json::Null => None,
                    Json::String(s) if s.len() <= 4096 => Some(s),
                    _ => return Err(bad("invalid number format")),
                };
                bytes = bytes.saturating_add(format.map_or(0, |s| s.len()));
                if bytes > MAX_CONTENT_BYTES {
                    return Err(bad("format payload exceeds budget"));
                }
                formats.push(format.cloned());
            }
            formats
        }
        _ => return Err(bad("invalid formats")),
    };
    apply(session, sheet, range, &contents, &formats)
}

fn apply(session: &mut Session, sheet: usize, range: RangeRef, contents: &[CellContent], formats: &[Option<String>]) -> Result<Json> {
    validate(range, contents)?;
    if !formats.is_empty() && formats.len() != contents.len() {
        return Err(bad("format shape mismatch"));
    }
    let mut bytes = 0usize;
    for format in formats.iter().flatten() {
        bytes = bytes.saturating_add(format.len());
        if format.len() > 4096 || bytes > MAX_CONTENT_BYTES {
            return Err(bad("format payload exceeds budget"));
        }
    }
    if session.doc()?.wb.sheet(sheet).is_none() {
        return Err(bad("worksheet not found"));
    }
    if formats.iter().any(Option::is_some) {
        cmd::refuse_unless_allowed(session, sheet, "Formatting", |p| p.format_cells)?;
    }
    cmd::edit(session, |cx| {
        for (index, (at, content)) in range.iter().zip(contents).enumerate() {
            let style = cx.wb.sheet(sheet).map(|s| s.style_id(at)).unwrap_or_default();
            let style = if let Some(Some(format)) = formats.get(index) {
                cx.wb.styles.derive(style, |s| s.num_fmt = gridcraft_model::NumFmt::new(format))
            } else {
                style
            };
            let cell = match content {
                CellContent::Value(value) => Cell { value: value.clone(), formula: None, style },
                CellContent::Formula(text) => {
                    let body = text.strip_prefix('=').unwrap_or(text);
                    let expr = cx.wb.name_call_case(cmd::edit::parse_typed_formula(body)?);
                    cmd::edit::check_arg_counts(&expr)?;
                    Cell { style, ..Cell::formula(Formula::from_expr(expr)) }
                }
            };
            cx.sheet_mut(sheet)?.set_cell(at, cell);
            cmd::data::grow_autofilter(cx.sheet_mut(sheet)?, at);
            cx.touch(sheet, at);
        }
        Ok(json!({"cells":contents.len()}))
    })
}
