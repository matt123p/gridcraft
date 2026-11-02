//! Recorded-macro selection and bounded native workbook clipboard operations.
use super::objects::arguments;
use super::ranges::Range;
use super::{Bridge, HostCall, HostResult, MAX_CELLS, Session, Value, failure, values};
use gridcraft_engine::core::RangeRef;
use serde_json::json;

impl Bridge {
    pub(crate) fn select_range(&mut self, session: &mut Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        arguments(call, &[])?;
        if session.doc().map_err(|e| failure(91, e.to_string()))?.sheet_index() != sheet {
            return Err(failure(1004, "Range.Select requires its worksheet to be active"));
        }
        let first = range.first()?;
        let address = range.areas.iter().map(RangeRef::a1).collect::<Vec<_>>().join(",");
        session.execute("selection.set", json!({"range":address,"active":first.start.a1()})).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }

    pub(crate) fn select_sheet(&mut self, session: &mut Session, sheet: usize, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, if call.member.eq_ignore_ascii_case("Select") { &["Replace"] } else { &[] })?;
        if let Some(replace) = args.first().filter(|v| !matches!(v, Value::Missing)) {
            let replace = vba_runtime::bind(vba_values::ScalarKind::Boolean, replace).map_err(values::runtime_error)?;
            if replace == Value::Boolean(false) {
                return Err(failure(1004, "Worksheet.Select Replace:=False (sheet grouping) is unsupported"));
            }
        }
        session.execute("sheet.activate", json!({"sheet":sheet})).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }

    pub(crate) fn copy_range(&mut self, session: &mut Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Destination"])?;
        if args.first().is_some_and(|v| !matches!(v, Value::Missing)) {
            return Err(failure(1004, "Range.Copy with Destination is unsupported; use Copy then Worksheet.Paste"));
        }
        let rect = range.rectangle()?;
        if rect.count() > MAX_CELLS as u64 {
            return Err(failure(7, "clipboard cell budget exceeded"));
        }
        session.execute("edit.copy", json!({"sheet":sheet,"range":rect.a1()})).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }

    pub(crate) fn paste_sheet(&mut self, session: &mut Session, sheet: usize, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Destination", "Link"])?;
        if args.first().is_some_and(|v| !matches!(v, Value::Missing)) {
            return Err(failure(1004, "Worksheet.Paste with Destination is unsupported; select the destination first"));
        }
        if let Some(link) = args.get(1).filter(|v| !matches!(v, Value::Missing)) {
            let link = vba_runtime::bind(vba_values::ScalarKind::Boolean, link).map_err(values::runtime_error)?;
            if link == Value::Boolean(true) {
                return Err(failure(1004, "Worksheet.Paste Link:=True is unsupported"));
            }
        }
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        if doc.sheet_index() != sheet {
            return Err(failure(1004, "Worksheet.Paste requires its worksheet to be active"));
        }
        if doc.selection.ranges.len() != 1 {
            return Err(failure(1004, "Worksheet.Paste requires one destination area"));
        }
        let selected = doc.selection.current();
        let clip = session.clipboard.as_ref().ok_or_else(|| failure(1004, "Nothing to paste"))?;
        if clip.cut || !clip.areas.is_empty() || clip.doc_uid != doc.uid {
            return Err(failure(1004, "VBA paste requires a single-area copy from the current workbook"));
        }
        let src = clip.range;
        // Mirror the engine's rectangular tiling before invoking it, so both
        // materialization and the Worksheet_Change target have a bounded footprint.
        let tiled = selected.height() % src.height() == 0
            && selected.width() % src.width() == 0
            && !selected.is_single()
            && !selected.is_full_cols()
            && !selected.is_full_rows();
        let (height, width) = if tiled { (selected.height(), selected.width()) } else { (src.height(), src.width()) };
        if src.count() > MAX_CELLS as u64 || u64::from(height) * u64::from(width) > MAX_CELLS as u64 {
            return Err(failure(7, "clipboard cell budget exceeded"));
        }
        let end =
            selected.start.offset(i64::from(height) - 1, i64::from(width) - 1).ok_or_else(|| failure(1004, "Paste extends outside the worksheet"))?;
        let target = RangeRef::new(selected.start, end);
        let event = self.preflight_change(session, sheet)?;
        session.execute("edit.paste", json!({})).map_err(|e| failure(1004, e.to_string()))?;
        if let Some(owner) = event {
            self.queue_change(owner, sheet, target.into())?;
        }
        Ok(Value::Empty)
    }
}
