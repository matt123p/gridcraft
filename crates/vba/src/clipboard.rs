//! Recorded-macro selection and bounded native workbook clipboard operations.
use super::objects::{Object, arguments};
use super::ranges::Range;
use super::{Bridge, HostCall, HostResult, MAX_CELLS, Session, Value, failure, values};
use gridcraft_engine::core::{CellRef, RangeRef};
use serde_json::json;
use vba_values::{ScalarKind, VbaArray};

impl Bridge {
    pub(crate) fn paste_special(&mut self, session: &mut Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Paste", "Operation", "SkipBlanks", "Transpose"])?;
        let clip = session.clipboard.clone().ok_or_else(|| failure(1004, "Nothing to paste"))?;
        if clip.cut || clip.doc_uid != session.doc().map_err(|e| failure(91, e.to_string()))?.uid || !clip.areas.is_empty() {
            return Err(failure(1004, "PasteSpecial requires a single-area copy from the current workbook"));
        }
        let selected = range.rectangle()?;
        if selected.count() > MAX_CELLS as u64 || clip.range.count() > MAX_CELLS as u64 {
            return Err(failure(7, "clipboard cell budget exceeded"));
        }
        let paste_type = enum_value(args.first(), -4104, "Paste")?;
        let what = match paste_type {
            -4104 => "all",
            7 => "allExceptBorders",
            8 => "columnWidths",
            -4144 => "comments",
            -4122 => "formats",
            -4123 => "formulas",
            11 => "formulasAndNumberFormats",
            6 => "validation",
            -4163 => "values",
            12 => "valuesAndNumberFormats",
            _ => return Err(failure(1004, format!("unsupported Range.PasteSpecial Paste value: {paste_type}"))),
        };
        let operation_value = enum_value(args.get(1), -4142, "Operation")?;
        let operation = match operation_value {
            -4142 => "none",
            2 => "add",
            3 => "subtract",
            4 => "multiply",
            5 => "divide",
            _ => return Err(failure(1004, format!("invalid Range.PasteSpecial Operation value: {operation_value}"))),
        };
        let skip_blanks = boolean_option(args.get(2), "SkipBlanks")?;
        let transpose = boolean_option(args.get(3), "Transpose")?;
        let h = if transpose { clip.range.width() } else { clip.range.height() };
        let w = if transpose { clip.range.height() } else { clip.range.width() };
        let tiled = selected.height() % h == 0
            && selected.width() % w == 0
            && !selected.is_single()
            && !selected.is_full_cols()
            && !selected.is_full_rows()
            && selected.count() <= MAX_CELLS as u64;
        let (height, width) = if tiled { (selected.height(), selected.width()) } else { (h, w) };
        let end =
            selected.start.offset(i64::from(height) - 1, i64::from(width) - 1).ok_or_else(|| failure(1004, "PasteSpecial exceeds the worksheet"))?;
        let target = RangeRef::new(selected.start, end);
        let event = if matches!(what, "all" | "allExceptBorders" | "formulas" | "formulasAndNumberFormats" | "values" | "valuesAndNumberFormats") {
            self.prepare_change(session, sheet, target.into())?
        } else {
            super::PreparedChanges::default()
        };
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        if doc.sheet_index() != sheet {
            session.execute("sheet.activate", json!({"sheet":sheet})).map_err(|e| failure(1004, e.to_string()))?;
        }
        session.execute("selection.set", json!({"range":selected.a1(),"active":selected.start.a1()})).map_err(|e| failure(1004, e.to_string()))?;
        session
            .execute("edit.pasteSpecial", json!({"sheet":sheet,"what":what,"operation":operation,"skipBlanks":skip_blanks,"transpose":transpose}))
            .map_err(|e| failure(1004, e.to_string()))?;
        self.queue_changes(event);
        Ok(Value::Empty)
    }

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

    pub(crate) fn activate_range(&mut self, session: &mut Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        arguments(call, &[])?;
        let first = range.first()?;
        let address = range.areas.iter().map(RangeRef::a1).collect::<Vec<_>>().join(",");
        session
            .execute("selection.set", json!({"sheet":sheet,"range":address,"active":first.start.a1()}))
            .map_err(|error| failure(1004, error.to_string()))?;
        Ok(Value::Empty)
    }

    pub(crate) fn select_sheet(&mut self, session: &mut Session, sheet: usize, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, if call.member.eq_ignore_ascii_case("Select") { &["Replace"] } else { &[] })?;
        let select = call.member.eq_ignore_ascii_case("Select");
        let replace = args
            .first()
            .filter(|value| !matches!(value, Value::Missing))
            .map(|value| vba_runtime::bind(vba_values::ScalarKind::Boolean, value).map_err(values::runtime_error))
            .transpose()?
            .map(|value| value == Value::Boolean(true))
            .unwrap_or(true);
        session
            .execute(
                if select { "sheet.select" } else { "sheet.activate" },
                if select { json!({"sheet":sheet,"replace":replace}) } else { json!({"sheet":sheet}) },
            )
            .map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }

    pub(crate) fn copy_range(&mut self, session: &mut Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Destination"])?;
        let rect = range.rectangle()?;
        if rect.count() > MAX_CELLS as u64 {
            return Err(failure(7, "clipboard cell budget exceeded"));
        }
        if let Some(Value::Object(reference)) = args.first().filter(|v| !matches!(v, Value::Missing)) {
            let Object::Range(destination_sheet, destination) = self.object(session, *reference)? else {
                return Err(failure(13, "Range.Copy Destination must be a Range"));
            };
            if destination.areas.len() != 1 {
                return Err(failure(1004, "Range.Copy Destination must be one area"));
            }
            let start = destination.first()?.start;
            let end = start
                .offset(i64::from(rect.height()) - 1, i64::from(rect.width()) - 1)
                .ok_or_else(|| failure(1004, "Copy destination exceeds the worksheet"))?;
            let target = RangeRef::new(start, end);
            let event = self.prepare_change(session, destination_sheet, target.into())?;
            session.execute("edit.copy", json!({"sheet":sheet,"range":rect.a1()})).map_err(|e| failure(1004, e.to_string()))?;
            session.execute("sheet.activate", json!({"sheet":destination_sheet})).map_err(|e| failure(1004, e.to_string()))?;
            session.execute("selection.set", json!({"range":start.a1(),"active":start.a1()})).map_err(|e| failure(1004, e.to_string()))?;
            session.execute("edit.paste", json!({})).map_err(|e| failure(1004, e.to_string()))?;
            self.queue_changes(event);
            return Ok(Value::Empty);
        }
        if args.first().is_some_and(|v| !matches!(v, Value::Missing)) {
            return Err(failure(13, "Range.Copy Destination must be a Range"));
        }
        session.execute("edit.copy", json!({"sheet":sheet,"range":rect.a1()})).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }

    pub(crate) fn cut_range(&mut self, session: &mut Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Destination"])?;
        let rect = range.rectangle()?;
        if rect.count() > MAX_CELLS as u64 {
            return Err(failure(7, "clipboard cell budget exceeded"));
        }
        if let Some(Value::Object(reference)) = args.first().filter(|v| !matches!(v, Value::Missing)) {
            let Object::Range(destination_sheet, destination) = self.object(session, *reference)? else {
                return Err(failure(13, "Range.Cut Destination must be a Range"));
            };
            if destination.areas.len() != 1 {
                return Err(failure(1004, "Range.Cut Destination must be one area"));
            }
            let start = destination.first()?.start;
            let end = start
                .offset(i64::from(rect.height()) - 1, i64::from(rect.width()) - 1)
                .ok_or_else(|| failure(1004, "Cut destination exceeds the worksheet"))?;
            let target = RangeRef::new(start, end);
            if sheet == destination_sheet && rect.intersects(&target) {
                return Err(failure(1004, "overlapping cut moves are unsupported"));
            }
            let events = if sheet == destination_sheet {
                let mut changed: Range = rect.into();
                changed.areas.push(target);
                changed.areas.sort_by_key(|area| (area.start.row, area.start.col));
                self.prepare_change(session, sheet, changed)?
            } else {
                self.prepare_changes(session, &[(sheet, rect.into()), (destination_sheet, target.into())])?
            };
            session.execute("edit.cut", json!({"sheet":sheet,"range":rect.a1()})).map_err(|e| failure(1004, e.to_string()))?;
            session.execute("sheet.activate", json!({"sheet":destination_sheet})).map_err(|e| failure(1004, e.to_string()))?;
            session.execute("selection.set", json!({"range":start.a1(),"active":start.a1()})).map_err(|e| failure(1004, e.to_string()))?;
            session.execute("edit.paste", json!({})).map_err(|e| failure(1004, e.to_string()))?;
            self.queue_changes(events);
            return Ok(Value::Empty);
        }
        if args.first().is_some_and(|v| !matches!(v, Value::Missing)) {
            return Err(failure(13, "Range.Cut Destination must be a Range"));
        }
        session.execute("edit.cut", json!({"sheet":sheet,"range":rect.a1()})).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }

    pub(crate) fn paste_sheet(&mut self, session: &mut Session, sheet: usize, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Destination", "Link"])?;
        let link = boolean_option(args.get(1), "Link")?;
        if link {
            if args.first().is_some_and(|value| !matches!(value, Value::Missing)) {
                return Err(failure(1004, "Worksheet.Paste cannot combine Destination and Link"));
            }
            return self.paste_sheet_link(session, sheet);
        }
        let destination = match args.first().filter(|v| !matches!(v, Value::Missing)) {
            Some(Value::Object(reference)) => {
                let Object::Range(destination_sheet, range) = self.object(session, *reference)? else {
                    return Err(failure(13, "Worksheet.Paste Destination must be a Range"));
                };
                if range.areas.len() != 1 {
                    return Err(failure(1004, "Worksheet.Paste Destination must be one area"));
                }
                Some((destination_sheet, range.first()?.start))
            }
            Some(_) => return Err(failure(13, "Worksheet.Paste Destination must be a Range")),
            None => None,
        };
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        if destination.is_none() && doc.sheet_index() != sheet {
            return Err(failure(1004, "Worksheet.Paste requires its worksheet to be active"));
        }
        if destination.is_none() && doc.selection.ranges.len() != 1 {
            return Err(failure(1004, "Worksheet.Paste requires one destination area"));
        }
        let selected = destination.map_or_else(|| doc.selection.current(), |(_, start)| RangeRef::cell(start));
        let (destination_sheet, start) = destination.unwrap_or((sheet, selected.start));
        let clip = session.clipboard.clone().ok_or_else(|| failure(1004, "Nothing to paste"))?;
        if !clip.areas.is_empty() || clip.doc_uid != doc.uid {
            return Err(failure(1004, "VBA paste requires a single-area clipboard from the current workbook"));
        }
        let src = clip.range;
        // Mirror the engine's rectangular tiling before invoking it, so both
        // materialization and the Worksheet_Change target have a bounded footprint.
        let tiled = !clip.cut
            && selected.height() % src.height() == 0
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
        let events = if clip.cut {
            if clip.sheet == destination_sheet && clip.range.intersects(&target) {
                return Err(failure(1004, "overlapping cut moves are unsupported"));
            }
            if clip.sheet == destination_sheet {
                let mut changed: Range = clip.range.into();
                changed.areas.push(target);
                changed.areas.sort_by_key(|area| (area.start.row, area.start.col));
                self.prepare_change(session, destination_sheet, changed)?
            } else {
                self.prepare_changes(session, &[(clip.sheet, clip.range.into()), (destination_sheet, target.into())])?
            }
        } else {
            self.prepare_change(session, destination_sheet, target.into())?
        };
        if doc.sheet_index() != destination_sheet {
            session.execute("sheet.activate", json!({"sheet":destination_sheet})).map_err(|e| failure(1004, e.to_string()))?;
        }
        if destination.is_some() || clip.cut {
            session.execute("selection.set", json!({"range":start.a1(),"active":start.a1()})).map_err(|e| failure(1004, e.to_string()))?;
        }
        session.execute("edit.paste", json!({})).map_err(|e| failure(1004, e.to_string()))?;
        self.queue_changes(events);
        Ok(Value::Empty)
    }

    fn paste_sheet_link(&mut self, session: &mut Session, sheet: usize) -> HostResult<Value> {
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        if doc.sheet_index() != sheet {
            return Err(failure(1004, "Worksheet.Paste Link requires its worksheet to be active"));
        }
        if doc.selection.ranges.len() != 1 {
            return Err(failure(1004, "Worksheet.Paste Link requires one selected area"));
        }
        let clipboard = session.clipboard.clone().ok_or_else(|| failure(1004, "Nothing to paste"))?;
        if clipboard.cut || clipboard.doc_uid != doc.uid || !clipboard.areas.is_empty() {
            return Err(failure(1004, "Worksheet.Paste Link requires a copied single-area range in this workbook"));
        }
        let source = clipboard.range;
        let selected = doc.selection.current();
        let tiled = !selected.is_single() && selected.height() % source.height() == 0 && selected.width() % source.width() == 0;
        let (height, width) = if tiled { (selected.height(), selected.width()) } else { (source.height(), source.width()) };
        let count = u64::from(height).checked_mul(u64::from(width)).ok_or_else(|| failure(7, "linked paste cell count overflow"))?;
        if source.count() > MAX_CELLS as u64 || count > MAX_CELLS as u64 {
            return Err(failure(7, "linked paste exceeds VBA cell budget"));
        }
        let end =
            selected.start.offset(i64::from(height) - 1, i64::from(width) - 1).ok_or_else(|| failure(1004, "linked paste exceeds the worksheet"))?;
        let target = RangeRef::new(selected.start, end);
        let source_sheet = doc.wb.sheet(clipboard.sheet).ok_or_else(|| failure(91, "clipboard worksheet was deleted"))?;
        let qualifier = format!("'{}'!", source_sheet.name.replace('\'', "''"));
        let rows = i32::try_from(height).map_err(|_| failure(7, "linked paste row budget"))?;
        let columns = i32::try_from(width).map_err(|_| failure(7, "linked paste column budget"))?;
        let mut formulas = VbaArray::new(ScalarKind::Variant, vec![(1, rows), (1, columns)], false).map_err(|error| failure(7, error.to_string()))?;
        for row in 0..height {
            for column in 0..width {
                let source_row = row % source.height();
                let source_column = column % source.width();
                let source_cell = CellRef::new(source.start.row + source_row, source.start.col + source_column);
                let formula = format!("={qualifier}{}", source_cell.a1());
                formulas.set((row * width + column) as usize, values::vba_text(&formula)?).map_err(|error| failure(7, error.to_string()))?;
            }
        }
        let event = self.prepare_change(session, sheet, target.into())?;
        let linked_values = Value::Array(formulas);
        values::write(session, sheet, target, &linked_values, true, false)?;
        session
            .execute("selection.set", json!({"range":target.a1(),"active":target.start.a1()}))
            .map_err(|error| failure(1004, error.to_string()))?;
        self.queue_changes(event);
        Ok(Value::Empty)
    }
}

fn enum_value(value: Option<&Value>, default: i32, name: &str) -> HostResult<i32> {
    match value.filter(|value| !matches!(value, Value::Missing)) {
        None => Ok(default),
        Some(value) => match vba_runtime::bind(vba_values::ScalarKind::Long, value).map_err(values::runtime_error)? {
            Value::Long(number) => Ok(number),
            _ => Err(failure(13, format!("{name} type mismatch"))),
        },
    }
}

fn boolean_option(value: Option<&Value>, name: &str) -> HostResult<bool> {
    match value.filter(|value| !matches!(value, Value::Missing)) {
        None => Ok(false),
        Some(value) => match vba_runtime::bind(vba_values::ScalarKind::Boolean, value).map_err(values::runtime_error)? {
            Value::Boolean(value) => Ok(value),
            _ => Err(failure(13, format!("{name} type mismatch"))),
        },
    }
}
