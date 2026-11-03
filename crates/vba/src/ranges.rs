//! Lightweight Excel references; materialization belongs to values, never construction.
use super::{Bridge, HostResult, MAX_CELLS, MAX_OBJECTS, Session, Value, failure};
use gridcraft_engine::core::{CellRef, MAX_COLS, MAX_ROWS, RangeRef};
use serde_json::json;
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Unit {
    Cells,
    Rows,
    Columns,
}
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct Range {
    pub(crate) areas: Vec<RangeRef>,
    pub(crate) unit: Unit,
}

pub(crate) struct DataSeriesOptions {
    pub(crate) rows: Option<bool>,
    pub(crate) series_type: i32,
    pub(crate) date_unit: i32,
    pub(crate) step: f64,
    pub(crate) stop: Option<f64>,
    pub(crate) trend: bool,
}
impl From<RangeRef> for Range {
    fn from(rect: RangeRef) -> Self {
        Self { areas: vec![rect], unit: Unit::Cells }
    }
}
impl Range {
    pub(crate) fn validate(&self) -> HostResult<()> {
        if self.areas.is_empty() || self.areas.len() > 64 {
            return Err(failure(1004, "range area budget"));
        }
        for (index, r) in self.areas.iter().enumerate() {
            if !r.start.is_valid() || !r.end.is_valid() || r.start.row > r.end.row || r.start.col > r.end.col {
                return Err(failure(1004, "range outside worksheet"));
            }
            if self.areas.iter().take(index).any(|a| a.intersects(r)) {
                return Err(failure(1004, "overlapping areas are unsupported"));
            }
        }
        Ok(())
    }
    pub(crate) fn first(&self) -> HostResult<RangeRef> {
        self.validate()?;
        self.areas.first().copied().ok_or_else(|| failure(1004, "empty range"))
    }
    pub(crate) fn rectangle(&self) -> HostResult<RangeRef> {
        if self.areas.len() != 1 {
            return Err(failure(1004, "this operation requires one area"));
        }
        self.first()
    }
    pub(crate) fn count(&self) -> HostResult<u64> {
        self.validate()?;
        self.areas.iter().try_fold(0u64, |n, r| {
            n.checked_add(match self.unit {
                Unit::Cells => r.count(),
                Unit::Rows => u64::from(r.height()),
                Unit::Columns => u64::from(r.width()),
            })
            .ok_or_else(|| failure(6, "range count overflow"))
        })
    }
    pub(crate) fn collection(&self, unit: Unit) -> HostResult<Self> {
        Ok(Self { areas: if unit == Unit::Cells { self.areas.clone() } else { vec![self.first()?] }, unit })
    }
    pub(crate) fn item(&self, row: i32, column: Option<i32>) -> HostResult<Self> {
        let r = self.first()?;
        if self.unit != Unit::Cells && column.is_some() {
            return Err(failure(1004, "Rows/Columns Item takes one index; use Cells for cell indexes"));
        }
        let offset = i64::from(row) - 1;
        let (dr, dc) = match (self.unit, column) {
            (Unit::Cells, Some(col)) => (offset, i64::from(col) - 1),
            (Unit::Cells, None) => (offset.div_euclid(i64::from(r.width())), offset.rem_euclid(i64::from(r.width()))),
            (Unit::Rows, None) => (offset, 0),
            (Unit::Columns, None) => (0, offset),
            _ => return Err(failure(1004, "invalid Item indexes")),
        };
        let row = i64::from(r.start.row).checked_add(dr).ok_or_else(|| failure(6, "row overflow"))?;
        let col = i64::from(r.start.col).checked_add(dc).ok_or_else(|| failure(6, "column overflow"))?;
        if row < 0 || row >= i64::from(MAX_ROWS) || col < 0 || col >= i64::from(MAX_COLS) {
            return Err(failure(1004, "Item outside worksheet"));
        }
        let start =
            CellRef::new(u32::try_from(row).map_err(|_| failure(6, "row overflow"))?, u32::try_from(col).map_err(|_| failure(6, "column overflow"))?);
        let end = match self.unit {
            Unit::Cells => start,
            Unit::Rows => CellRef::new(start.row, r.end.col),
            Unit::Columns => CellRef::new(r.end.row, start.col),
        };
        let result = Self { areas: vec![RangeRef { start, end }], unit: self.unit };
        result.validate()?;
        Ok(result)
    }
    pub(crate) fn slice(&self, first: i32, last: i32) -> HostResult<Self> {
        if self.unit == Unit::Cells {
            return Err(failure(1004, "a row/column selector is required"));
        }
        let a = self.item(first, None)?.first()?;
        let b = self.item(last, None)?.first()?;
        let result = Self { areas: vec![a.union(&b)], unit: self.unit };
        result.validate()?;
        Ok(result)
    }
    pub(crate) fn entries(&self) -> HostResult<Vec<Self>> {
        let n = self.count()?;
        if n > MAX_OBJECTS as u64 {
            return Err(failure(7, "range enumeration budget"));
        }
        let mut out = Vec::new();
        out.try_reserve(n as usize).map_err(|_| failure(7, "enumeration allocation"))?;
        for r in &self.areas {
            match self.unit {
                Unit::Cells => out.extend(r.iter().map(|c| Self::from(RangeRef::cell(c)))),
                Unit::Rows => {
                    for row in r.start.row..=r.end.row {
                        out.push(Self {
                            areas: vec![RangeRef { start: CellRef::new(row, r.start.col), end: CellRef::new(row, r.end.col) }],
                            unit: Unit::Rows,
                        });
                    }
                }
                Unit::Columns => {
                    for col in r.start.col..=r.end.col {
                        out.push(Self {
                            areas: vec![RangeRef { start: CellRef::new(r.start.row, col), end: CellRef::new(r.end.row, col) }],
                            unit: Unit::Columns,
                        });
                    }
                }
            }
        }
        Ok(out)
    }
    pub(crate) fn offset(&self, dr: i64, dc: i64) -> HostResult<Self> {
        let areas = self
            .areas
            .iter()
            .map(|r| {
                let shift = |c: CellRef| -> HostResult<CellRef> { c.offset(dr, dc).ok_or_else(|| failure(1004, "Offset outside worksheet")) };
                Ok(RangeRef { start: shift(r.start)?, end: shift(r.end)? })
            })
            .collect::<HostResult<Vec<_>>>()?;
        let result = Self { areas, unit: Unit::Cells };
        result.validate()?;
        Ok(result)
    }
    pub(crate) fn resize(&self, rows: Option<i32>, cols: Option<i32>) -> HostResult<Self> {
        let r = self.rectangle()?;
        let rows = rows.map(i64::from).unwrap_or(i64::from(r.height()));
        let cols = cols.map(i64::from).unwrap_or(i64::from(r.width()));
        if rows <= 0 || cols <= 0 {
            return Err(failure(1004, "Resize requires positive dimensions"));
        }
        let end = r.start.offset(rows - 1, cols - 1).ok_or_else(|| failure(1004, "Resize outside worksheet"))?;
        Ok(RangeRef { start: r.start, end }.into())
    }
}

pub(crate) fn from_cells(mut cells: Vec<CellRef>, operation: &str) -> HostResult<Range> {
    cells.sort_unstable();
    cells.dedup();
    if cells.is_empty() {
        return Err(failure(1004, format!("{operation} found no cells")));
    }
    if cells.len() > MAX_CELLS {
        return Err(failure(7, format!("{operation} exceeds VBA cell budget")));
    }
    let mut areas = Vec::<RangeRef>::new();
    let mut index = 0;
    while index < cells.len() {
        let start = cells[index];
        let mut end = start;
        index += 1;
        while index < cells.len() && cells[index].row == start.row && end.col.checked_add(1) == Some(cells[index].col) {
            end = cells[index];
            index += 1;
        }
        areas.push(RangeRef::new(start, end));
    }
    areas.sort_by_key(|area| (area.start.col, area.end.col, area.start.row));
    let mut merged = Vec::<RangeRef>::new();
    for area in areas {
        if let Some(previous) = merged.last_mut()
            && previous.start.col == area.start.col
            && previous.end.col == area.end.col
            && previous.end.row.checked_add(1) == Some(area.start.row)
        {
            previous.end.row = area.end.row;
        } else {
            merged.push(area);
        }
    }
    if merged.len() > 64 {
        return Err(failure(7, format!("{operation} exceeds 64 disjoint areas")));
    }
    merged.sort_by_key(|area| (area.start.row, area.start.col));
    Ok(Range { areas: merged, unit: Unit::Cells })
}

impl Bridge {
    pub(crate) fn data_series(&mut self, session: &mut Session, sheet: usize, range: &Range, options: DataSeriesOptions) -> HostResult<Value> {
        let DataSeriesOptions { rows, series_type, date_unit, step, stop, trend } = options;
        let rect = range.rectangle()?;
        if rect.count() > MAX_CELLS as u64 {
            return Err(failure(7, "Range.DataSeries exceeds VBA cell budget"));
        }
        if trend {
            return Err(failure(1004, "Range.DataSeries Trend:=True requires a regression service"));
        }
        let kind = match series_type {
            -4132 => "linear",
            2 => "growth",
            3 => "date",
            4 => return Err(failure(1004, "Range.DataSeries xlAutoFill requires AutoFill pattern inference")),
            _ => return Err(failure(1004, format!("Range.DataSeries Type {series_type} is invalid"))),
        };
        let unit = match date_unit {
            1 => "day",
            2 => "weekday",
            3 => "month",
            4 => "year",
            _ => return Err(failure(1004, format!("Range.DataSeries Date {date_unit} is invalid"))),
        };
        let event = self.prepare_change(session, sheet, range.clone())?;
        let mut params = json!({"sheet":sheet,"range":rect.a1(),"type":kind,"dateUnit":unit,"step":step});
        if let Some(rows) = rows {
            params["direction"] = serde_json::Value::String(if rows { "rows" } else { "columns" }.into());
        }
        if let Some(stop) = stop {
            params["stop"] = json!(stop);
        }
        session.execute("edit.fillSeries", params).map_err(|error| failure(1004, error.to_string()))?;
        self.queue_changes(event);
        Ok(Value::Empty)
    }

    pub(crate) fn special_cells(&mut self, session: &Session, sheet: usize, range: &Range, cell_type: i32, value_mask: i32) -> HostResult<Value> {
        range.validate()?;
        let worksheet =
            session.doc().map_err(|error| failure(91, error.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
        let contains = |cell: CellRef| range.areas.iter().any(|area| area.contains(cell));
        let cells = if cell_type == 11 {
            worksheet.used_range().map(|used| used.end).filter(|cell| contains(*cell)).into_iter().collect::<Vec<_>>()
        } else {
            if range.count()? > MAX_CELLS as u64 {
                return Err(failure(7, "Range.SpecialCells exceeds VBA cell budget"));
            }
            let value_matches = |value: &gridcraft_engine::core::Value| match value {
                gridcraft_engine::core::Value::Number(_) => value_mask & 1 != 0,
                gridcraft_engine::core::Value::Text(_) => value_mask & 2 != 0,
                gridcraft_engine::core::Value::Bool(_) => value_mask & 4 != 0,
                gridcraft_engine::core::Value::Error(_) => value_mask & 16 != 0,
                gridcraft_engine::core::Value::Empty | gridcraft_engine::core::Value::Array(_) => false,
            };
            let reference = range.first()?.start;
            range
                .areas
                .iter()
                .flat_map(RangeRef::iter)
                .filter(|cell| {
                    let stored = worksheet.cell(*cell);
                    let has_formula = stored.is_some_and(|entry| entry.formula.is_some());
                    match cell_type {
                        2 => !has_formula && value_matches(&worksheet.value(*cell)),
                        -4123 => has_formula && value_matches(&worksheet.value(*cell)),
                        4 => !has_formula && worksheet.value(*cell).is_empty(),
                        -4144 => worksheet.comments.contains_key(cell),
                        12 => !worksheet.is_row_hidden(cell.row) && !worksheet.is_col_hidden(cell.col),
                        -4172 => worksheet.cond_formats.iter().any(|format| format.ranges.iter().any(|area| area.contains(*cell))),
                        -4173 => {
                            let reference_formats = worksheet
                                .cond_formats
                                .iter()
                                .enumerate()
                                .filter(|(_, format)| format.ranges.iter().any(|area| area.contains(reference)))
                                .map(|(index, _)| index)
                                .collect::<Vec<_>>();
                            !reference_formats.is_empty()
                                && reference_formats.iter().all(|index| worksheet.cond_formats[*index].ranges.iter().any(|area| area.contains(*cell)))
                        }
                        -4174 => worksheet.validations.iter().any(|validation| validation.ranges.iter().any(|area| area.contains(*cell))),
                        -4175 => {
                            let reference_validations = worksheet
                                .validations
                                .iter()
                                .enumerate()
                                .filter(|(_, validation)| validation.ranges.iter().any(|area| area.contains(reference)))
                                .map(|(index, _)| index)
                                .collect::<Vec<_>>();
                            !reference_validations.is_empty()
                                && reference_validations
                                    .iter()
                                    .all(|index| worksheet.validations[*index].ranges.iter().any(|area| area.contains(*cell)))
                        }
                        _ => false,
                    }
                })
                .collect::<Vec<_>>()
        };
        if !matches!(cell_type, 2 | -4123 | 4 | -4144 | 11 | 12 | -4172 | -4173 | -4174 | -4175) {
            return Err(failure(1004, format!("Range.SpecialCells Type {cell_type} is invalid")));
        }
        let result = from_cells(cells, "Range.SpecialCells")?;
        self.intern(super::objects::Object::Range(sheet, result))
    }

    pub(crate) fn clear_range(&mut self, session: &mut Session, sheet: usize, range: &Range, what: &str) -> HostResult<Value> {
        range.validate()?;
        let rect_count = range
            .areas
            .iter()
            .try_fold(0u64, |count, area| count.checked_add(area.count()).ok_or_else(|| failure(7, "clear range cell count overflow")))?;
        if rect_count > MAX_CELLS as u64 {
            return Err(failure(7, "clear range exceeds VBA cell budget"));
        }
        let changes_contents = matches!(what, "all" | "contents");
        let event = if changes_contents { self.prepare_change(session, sheet, range.clone())? } else { super::PreparedChanges::default() };
        let address = range.areas.iter().map(RangeRef::a1).collect::<Vec<_>>().join(",");
        let command = match what {
            "all" => "edit.clearAll",
            "contents" => "edit.clearContents",
            "formats" => "edit.clearFormats",
            "comments" => "edit.clearComments",
            "hyperlinks" => "edit.clearHyperlinks",
            _ => return Err(failure(438, "unsupported Range clear operation")),
        };
        session.execute(command, json!({"sheet":sheet,"range":address})).map_err(|e| failure(1004, e.to_string()))?;
        self.queue_changes(event);
        Ok(Value::Empty)
    }

    pub(crate) fn fill_direction(&mut self, session: &mut Session, sheet: usize, range: &Range, direction: &str) -> HostResult<Value> {
        let rect = range.rectangle()?;
        if rect.count() > MAX_CELLS as u64 {
            return Err(failure(7, "Range fill exceeds VBA cell budget"));
        }
        let changed = match direction {
            "down" if rect.height() == 1 && rect.start.row > 0 => Some(rect),
            "down" if rect.height() > 1 => Some(RangeRef::new(CellRef::new(rect.start.row + 1, rect.start.col), rect.end)),
            "up" if rect.height() > 1 => Some(RangeRef::new(rect.start, CellRef::new(rect.end.row - 1, rect.end.col))),
            "right" if rect.width() == 1 && rect.start.col > 0 => Some(rect),
            "right" if rect.width() > 1 => Some(RangeRef::new(CellRef::new(rect.start.row, rect.start.col + 1), rect.end)),
            "left" if rect.width() > 1 => Some(RangeRef::new(rect.start, CellRef::new(rect.end.row, rect.end.col - 1))),
            "down" | "up" | "right" | "left" => None,
            _ => return Err(failure(438, "unsupported Range fill direction")),
        };
        let event = match changed {
            Some(target) => self.prepare_change(session, sheet, target.into())?,
            None => super::PreparedChanges::default(),
        };
        let command = match direction {
            "down" => "edit.fillDown",
            "up" => "edit.fillUp",
            "left" => "edit.fillLeft",
            "right" => "edit.fillRight",
            _ => return Err(failure(438, "unsupported Range fill direction")),
        };
        session.execute(command, json!({"sheet":sheet,"range":rect.a1()})).map_err(|error| failure(1004, error.to_string()))?;
        self.queue_changes(event);
        Ok(Value::Empty)
    }

    pub(crate) fn auto_fill(
        &mut self,
        session: &mut Session,
        sheet: usize,
        source: &Range,
        destination_sheet: usize,
        destination: &Range,
        fill_type: i32,
    ) -> HostResult<Value> {
        if sheet != destination_sheet {
            return Err(failure(1004, "Range.AutoFill Destination must be on the source worksheet"));
        }
        let source = source.rectangle()?;
        let destination = destination.rectangle()?;
        if source.count() > MAX_CELLS as u64 || destination.count() > MAX_CELLS as u64 {
            return Err(failure(7, "Range.AutoFill exceeds VBA cell budget"));
        }
        if !destination.contains_range(&source) {
            return Err(failure(1004, "Range.AutoFill Destination must contain the source"));
        }
        let vertical = source.width() == destination.width() && source.height() <= destination.height();
        let horizontal = source.height() == destination.height() && source.width() <= destination.width();
        if !vertical && !horizontal {
            return Err(failure(1004, "Range.AutoFill Destination must extend the source in one direction"));
        }
        let changed = if destination == source {
            None
        } else if vertical && destination.start.row == source.start.row {
            Some(RangeRef::new(CellRef::new(source.end.row + 1, destination.start.col), destination.end))
        } else if vertical && destination.end.row == source.end.row {
            Some(RangeRef::new(destination.start, CellRef::new(source.start.row - 1, destination.end.col)))
        } else if horizontal && destination.start.col == source.start.col {
            Some(RangeRef::new(CellRef::new(destination.start.row, source.end.col + 1), destination.end))
        } else if horizontal && destination.end.col == source.end.col {
            Some(RangeRef::new(destination.start, CellRef::new(destination.end.row, source.start.col - 1)))
        } else {
            return Err(failure(1004, "Range.AutoFill Destination must contain the source at an edge"));
        };
        let mode = match fill_type {
            0 => "series",
            1 => "copy",
            2 => "extend",
            3 => "formats",
            4 => "values",
            5 => "days",
            6 => "weekdays",
            7 => "months",
            8 => "years",
            9..=11 => return Err(failure(1004, "Range.AutoFill type requires a trend or FlashFill service")),
            _ => return Err(failure(1004, "invalid Range.AutoFill Type")),
        };
        let event = match changed {
            Some(target) => self.prepare_change(session, sheet, target.into())?,
            None => super::PreparedChanges::default(),
        };
        session
            .execute("edit.autoFill", json!({"sheet":sheet,"source":source.a1(),"target":destination.a1(),"mode":mode}))
            .map_err(|error| failure(1004, error.to_string()))?;
        self.queue_changes(event);
        Ok(Value::Empty)
    }
}
