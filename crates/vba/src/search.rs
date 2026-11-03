//! Sparse, bounded worksheet navigation. Never iterate the logical worksheet grid.
use super::objects::{arguments, numeric_index, required};
use super::ranges::Range;
use super::{Bridge, HostCall, HostResult, MAX_CELLS, Object, Session, Value, failure, values};
use gridcraft_engine::{
    core::{CellRef, MAX_COLS, MAX_ROWS, RangeRef},
    model::Cell,
};
#[derive(Clone, Debug)]
pub(crate) struct FindState {
    what: Option<Value>,
    look_in: i32,
    look_at: i32,
    order: i32,
    case: bool,
}
impl Default for FindState {
    fn default() -> Self {
        Self { what: None, look_in: -4123, look_at: 2, order: 1, case: false }
    }
}
fn boolean(v: &Value) -> HostResult<bool> {
    match vba_runtime::bind(vba_values::ScalarKind::Boolean, v).map_err(|e| match e {
        vba_runtime::Error::Vba { number, message } => failure(number, message),
        _ => failure(13, "Boolean conversion failed"),
    })? {
        Value::Boolean(b) => Ok(b),
        _ => Err(failure(13, "Boolean required")),
    }
}
fn occupied(cell: &Cell) -> bool {
    cell.formula.is_some() || !cell.value.is_empty()
}
fn scalar_text(value: &Value) -> HostResult<String> {
    match values::content(value, false)? {
        gridcraft_engine::CellContent::Value(v) => Ok(v.display()),
        _ => Err(failure(13, "scalar required")),
    }
}
#[derive(Clone, Copy)]
enum Token {
    Star,
    Any,
    Literal(char),
}
fn pattern(text: &str, whole: bool) -> HostResult<Vec<Token>> {
    if text.len() > 4096 {
        return Err(failure(7, "Find pattern budget"));
    }
    let mut tokens = Vec::new();
    if !whole {
        tokens.push(Token::Star);
    }
    let mut chars = text.chars();
    while let Some(c) = chars.next() {
        tokens.push(match c {
            '*' => Token::Star,
            '?' => Token::Any,
            '~' => Token::Literal(chars.next().unwrap_or('~')),
            c => Token::Literal(c),
        });
    }
    if !whole {
        tokens.push(Token::Star);
    }
    Ok(tokens)
}
fn matches(tokens: &[Token], text: &str, work: &mut usize) -> HostResult<bool> {
    if text.len() > 32768 {
        return Err(failure(7, "Find text budget"));
    }
    let mut previous = vec![false; tokens.len() + 1];
    if let Some(first) = previous.first_mut() {
        *first = true;
    }
    for (i, token) in tokens.iter().enumerate() {
        if matches!(token, Token::Star)
            && previous.get(i).copied().unwrap_or(false)
            && let Some(v) = previous.get_mut(i + 1)
        {
            *v = true;
        }
    }
    let mut next = vec![false; tokens.len() + 1];
    for c in text.chars() {
        *work = work.saturating_add(tokens.len());
        if *work > 8_000_000 {
            return Err(failure(7, "Find matching budget"));
        }
        next.fill(false);
        for (i, token) in tokens.iter().enumerate() {
            let diagonal = previous.get(i).copied().unwrap_or(false);
            let value = match token {
                Token::Star => next.get(i).copied().unwrap_or(false) || previous.get(i + 1).copied().unwrap_or(false),
                Token::Any => diagonal,
                Token::Literal(t) => diagonal && *t == c,
            };
            if let Some(slot) = next.get_mut(i + 1) {
                *slot = value;
            }
        }
        std::mem::swap(&mut previous, &mut next);
    }
    Ok(previous.last().copied().unwrap_or(false))
}
impl Bridge {
    fn scan(&mut self, n: usize) -> HostResult<()> {
        self.scan_work = self.scan_work.saturating_add(n);
        if self.scan_work > 1_000_000 {
            return Err(failure(7, "macro sparse scan budget"));
        }
        Ok(())
    }
    pub(crate) fn end(&mut self, session: &Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        let args = arguments(call, &["Direction"])?;
        let direction = numeric_index(required(&args, 0)?)?;
        let (vertical, forward) = match direction {
            -4162 => (true, false),
            -4121 => (true, true),
            -4159 => (false, false),
            -4161 => (false, true),
            _ => return Err(failure(1004, "invalid End direction")),
        };
        let from = range.first()?.start;
        let lane = if vertical { RangeRef::cols(from.col, from.col) } else { RangeRef::rows(from.row, from.row) };
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
        let mut filled = Vec::new();
        let mut scanned = 0;
        for (at, cell) in sh.cells.iter_range(lane) {
            self.scan(1)?;
            scanned += 1;
            if scanned > MAX_CELLS {
                return Err(failure(7, "End sparse scan budget"));
            }
            if occupied(cell) {
                filled.push(if vertical { at.row } else { at.col });
            }
        }

        filled.sort_unstable();
        let pos = if vertical { from.row } else { from.col };
        let limit = if vertical { MAX_ROWS } else { MAX_COLS };
        let neighbor = if forward { pos.checked_add(1).filter(|n| *n < limit) } else { pos.checked_sub(1) };
        let contains = |n| filled.binary_search(&n).is_ok();
        let destination = if let Some(next) = neighbor {
            if contains(pos) && contains(next) {
                let mut end = next;
                if forward {
                    for &n in filled.iter().filter(|n| **n > next) {
                        if n != end + 1 {
                            break;
                        }
                        end = n;
                    }
                } else {
                    for &n in filled.iter().rev().filter(|n| **n < next) {
                        if n + 1 != end {
                            break;
                        }
                        end = n;
                    }
                }
                end
            } else if forward {
                filled.iter().copied().find(|n| *n > pos).unwrap_or(limit - 1)
            } else {
                filled.iter().rev().copied().find(|n| *n < pos).unwrap_or(0)
            }
        } else {
            pos
        };
        self.intern(Object::Range(
            sheet,
            RangeRef::cell(if vertical { CellRef::new(destination, from.col) } else { CellRef::new(from.row, destination) }).into(),
        ))
    }
    pub(crate) fn find(&mut self, session: &Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        let rect = range.rectangle()?;
        let repeat = call.member.eq_ignore_ascii_case("FindNext") || call.member.eq_ignore_ascii_case("FindPrevious");
        let args = arguments(
            call,
            if repeat {
                &["After"]
            } else {
                &["What", "After", "LookIn", "LookAt", "SearchOrder", "SearchDirection", "MatchCase", "MatchByte", "SearchFormat"]
            },
        )?;
        let mut state = self.find.clone();
        let supplied = |i| args.get(i).filter(|v| !matches!(v, Value::Missing));
        let direction = if repeat {
            if call.member.eq_ignore_ascii_case("FindPrevious") { 2 } else { 1 }
        } else {
            state.what = Some(required(&args, 0)?.clone());
            if let Some(v) = supplied(2) {
                state.look_in = numeric_index(v)?;
            }
            if let Some(v) = supplied(3) {
                state.look_at = numeric_index(v)?;
            }
            if let Some(v) = supplied(4) {
                state.order = numeric_index(v)?;
            }
            state.case = supplied(6).map(boolean).transpose()?.unwrap_or(false);
            for i in [7, 8] {
                if supplied(i).map(boolean).transpose()?.unwrap_or(false) {
                    return Err(failure(1004, "MatchByte/SearchFormat unsupported"));
                }
            }
            supplied(5).map(numeric_index).transpose()?.unwrap_or(1)
        };
        if ![-4123, -4163].contains(&state.look_in)
            || ![1, 2].contains(&state.look_at)
            || ![1, 2].contains(&state.order)
            || ![1, 2].contains(&direction)
        {
            return Err(failure(1004, "unsupported Find option"));
        }
        let what = state.what.as_ref().ok_or_else(|| failure(1004, "FindNext requires a prior Find"))?;
        let mut needle = match what {
            Value::String(_) => values::text(what)?,
            Value::Null => return Err(failure(94, "invalid use of Null")),
            Value::Boolean(b) => if *b { "TRUE" } else { "FALSE" }.into(),
            Value::CvErr(_) => scalar_text(what)?,
            Value::Empty | Value::Missing | Value::Array(_) | Value::Object(_) | Value::Nothing => {
                return Err(failure(1004, "Find requires a nonempty scalar"));
            }
            _ => scalar_text(what)?,
        };
        if needle.is_empty() {
            return Err(failure(1004, "empty-cell Find is unsupported"));
        }
        if !state.case {
            needle = needle.to_lowercase();
        }
        let tokens = pattern(&needle, state.look_at == 1)?;
        let after = match supplied(if repeat { 0 } else { 1 }) {
            None => rect.start,
            Some(Value::Object(reference)) => match self.object(session, *reference)? {
                Object::Range(owner, r) if owner == sheet => {
                    let r = r.rectangle()?;
                    if !r.is_single() || !rect.contains(r.start) {
                        return Err(failure(1004, "After must be a single cell in the search range"));
                    }
                    r.start
                }
                _ => return Err(failure(1004, "After has another owner or type")),
            },
            Some(_) => return Err(failure(13, "After must be a Range")),
        };
        let key = |c: CellRef| if state.order == 1 { (c.row, c.col) } else { (c.col, c.row) };
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
        let mut hits = Vec::new();
        let mut scanned = 0;

        for (at, cell) in sh.cells.iter_range(rect) {
            self.scan(1)?;
            scanned += 1;
            if scanned > MAX_CELLS {
                return Err(failure(7, "Find sparse scan budget"));
            }
            if !occupied(cell) {
                continue;
            }
            let text = if state.look_in == -4123 {
                cell.formula.as_ref().map(|f| format!("={}", f.text.trim_start_matches('='))).unwrap_or_else(|| sh.value(at).display())
            } else {
                sh.value(at).display()
            };
            let text = if state.case { text } else { text.to_lowercase() };
            if matches(&tokens, &text, &mut self.match_work)? {
                hits.push(at);
            }
        }

        hits.sort_unstable_by_key(|c| key(*c));
        if direction == 2 {
            hits.reverse();
        }
        let found =
            hits.iter().copied().find(|c| if direction == 1 { key(*c) > key(after) } else { key(*c) < key(after) }).or_else(|| hits.first().copied());
        let result = match found {
            Some(c) => self.intern(Object::Range(sheet, RangeRef::cell(c).into()))?,
            None => Value::Nothing,
        };
        self.find = state;
        Ok(result)
    }

    pub(crate) fn replace_range(&mut self, session: &mut Session, sheet: usize, range: &Range, call: &HostCall) -> HostResult<Value> {
        let args = arguments(
            call,
            &["What", "Replacement", "LookAt", "SearchOrder", "MatchCase", "MatchByte", "SearchFormat", "ReplaceFormat", "FormulaVersion"],
        )?;
        let what = replacement_text(required(&args, 0)?, false)?;
        let replacement = replacement_text(required(&args, 1)?, true)?;
        if what.is_empty() {
            return Err(failure(1004, "Range.Replace What cannot be empty"));
        }
        let look_at = args.get(2).filter(|v| !matches!(v, Value::Missing)).map(numeric_index).transpose()?.unwrap_or(2);
        let search_order = args.get(3).filter(|v| !matches!(v, Value::Missing)).map(numeric_index).transpose()?.unwrap_or(1);
        if ![1, 2].contains(&look_at) || ![1, 2].contains(&search_order) {
            return Err(failure(1004, "Range.Replace LookAt or SearchOrder is invalid"));
        }
        let match_case = args.get(4).filter(|v| !matches!(v, Value::Missing)).map(boolean).transpose()?.unwrap_or(false);
        if args.get(5).filter(|v| !matches!(v, Value::Missing)).map(boolean).transpose()?.unwrap_or(false) {
            return Err(failure(1004, "Range.Replace MatchByte requires double-byte language support"));
        }
        for (index, name) in [(6, "SearchFormat"), (7, "ReplaceFormat")] {
            if args.get(index).filter(|v| !matches!(v, Value::Missing)).map(boolean).transpose()?.unwrap_or(false) {
                return Err(failure(1004, format!("Range.Replace {name} requires the format search service")));
            }
        }
        if args.get(8).is_some_and(|value| !matches!(value, Value::Missing)) {
            return Err(failure(1004, "Range.Replace FormulaVersion is not supported"));
        }
        let rect = range.rectangle()?;
        if rect.count() > MAX_CELLS as u64 {
            return Err(failure(7, "Range.Replace exceeds VBA cell budget"));
        }
        let params = serde_json::json!({
            "what":what,"with":replacement,"wholeCell":look_at == 1,"matchCase":match_case,
            "all":true,"sheet":sheet,"range":rect.a1()
        });
        let preview = session.execute("edit.replace", serde_json::json!({"preview":true,"what":params["what"],"with":params["with"],"wholeCell":params["wholeCell"],"matchCase":params["matchCase"],"all":true,"sheet":sheet,"range":rect.a1()}))
            .map_err(|error| failure(1004, error.to_string()))?;
        let changed = preview
            .get("cells")
            .and_then(serde_json::Value::as_array)
            .ok_or_else(|| failure(1004, "replace preview did not return changed cells"))?;
        if changed.len() > MAX_CELLS {
            return Err(failure(7, "Range.Replace change target exceeds VBA cell budget"));
        }
        if changed.is_empty() {
            return Ok(Value::Boolean(false));
        }
        let event = self.preflight_change(session, sheet)?;
        let target = if event.is_some() { Some(replace_target(changed)?) } else { None };
        if let Some(target) = &target {
            self.intern(Object::Range(sheet, target.clone()))?;
            self.pending_events.try_reserve(1).map_err(|_| failure(7, "worksheet event queue allocation"))?;
        }
        let result = session.execute("edit.replace", params).map_err(|error| failure(1004, error.to_string()))?;
        if result.get("replaced").and_then(serde_json::Value::as_u64) != Some(changed.len() as u64) {
            return Err(failure(1004, "Range.Replace result changed after preview"));
        }
        if let (Some(owner), Some(target)) = (event, target) {
            self.queue_change(owner, sheet, target)?;
        }
        Ok(Value::Boolean(true))
    }
}

fn replacement_text(value: &Value, allow_empty: bool) -> HostResult<String> {
    match value {
        Value::Null => Err(failure(94, "invalid use of Null")),
        Value::Missing => Err(failure(449, "argument not optional")),
        Value::Nothing | Value::Object(_) | Value::Array(_) => Err(failure(13, "Replace arguments must be scalar values")),
        Value::Empty if allow_empty => Ok(String::new()),
        Value::Empty => Err(failure(1004, "Range.Replace What cannot be empty")),
        _ => match values::content(value, false)? {
            gridcraft_engine::CellContent::Value(value) => Ok(value.display()),
            _ => Err(failure(13, "Replace arguments must be scalar values")),
        },
    }
}

fn replace_target(changed: &[serde_json::Value]) -> HostResult<Range> {
    let cells = changed
        .iter()
        .map(|entry| {
            let address = entry.get("cell").and_then(serde_json::Value::as_str).ok_or_else(|| failure(1004, "invalid replace preview cell"))?;
            CellRef::parse(address).ok_or_else(|| failure(1004, "invalid replace preview address"))
        })
        .collect::<HostResult<Vec<_>>>()?;
    super::ranges::from_cells(cells, "Range.Replace event target")
}
