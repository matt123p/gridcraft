//! Lightweight Excel references; materialization belongs to values, never construction.
use super::*;
use gridcraft_engine::core::{CellRef, MAX_COLS, MAX_ROWS, RangeRef};
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
