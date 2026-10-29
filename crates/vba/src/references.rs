//! Bounded reference grammar: geometry and names, never formula evaluation.
use super::objects::{arguments, required, string};
use super::ranges::{Range, Unit};
use super::*;
use gridcraft_engine::core::{CellRef, RangeRef, split_sheet};

// Split operators outside quoted worksheet names, including escaped apostrophes.
fn parts(text: &str, separator: char) -> HostResult<Vec<&str>> {
    let mut quoted = false;
    let mut start = 0;
    let mut out = Vec::new();
    let mut chars = text.char_indices().peekable();
    while let Some((i, c)) = chars.next() {
        if c == '\'' {
            if quoted && chars.peek().is_some_and(|(_, next)| *next == '\'') {
                chars.next();
            } else {
                quoted = !quoted;
            }
        } else if c == separator && !quoted {
            let part = text.get(start..i).unwrap_or_default().trim();
            if part.is_empty() && separator != ' ' {
                return Err(failure(1004, "empty range operand"));
            }
            if !part.is_empty() {
                out.push(part);
            }
            start = i + c.len_utf8();
        }
        if out.len() >= 64 {
            return Err(failure(7, "reference operator budget"));
        }
    }
    if quoted {
        return Err(failure(1004, "unclosed worksheet name"));
    }
    let last = text.get(start..).unwrap_or_default().trim();
    if last.is_empty() && separator != ' ' {
        return Err(failure(1004, "empty range operand"));
    }
    if !last.is_empty() {
        out.push(last);
    }
    if out.is_empty() {
        return Err(failure(1004, "empty reference"));
    }
    Ok(out)
}
impl Bridge {
    fn resolve_text(
        session: &Session,
        owner: usize,
        text: &str,
        relative: Option<CellRef>,
        depth: usize,
        work: &mut usize,
    ) -> HostResult<(usize, Range)> {
        *work = work.saturating_add(1);
        if depth >= 32 || text.len() > 4096 || *work > 1024 {
            return Err(failure(7, "reference/name budget"));
        }
        let text = text.trim().strip_prefix('=').unwrap_or(text.trim());
        if text.contains(['[', ']', '(', ')']) {
            return Err(failure(1004, "external or computed reference unsupported"));
        }
        let union = parts(text, ',')?;
        if union.len() > 1 {
            let mut sheet = None;
            let mut areas = Vec::new();
            for part in union {
                let (s, r) = Self::resolve_text(session, owner, part, relative, depth + 1, work)?;
                if sheet.is_some_and(|prior| prior != s) {
                    return Err(failure(1004, "areas must share a worksheet"));
                }
                sheet = Some(s);
                if areas.len() + r.areas.len() > 64 {
                    return Err(failure(7, "area budget"));
                }
                areas.extend(r.areas);
            }
            let r = Range { areas, unit: Unit::Cells };
            r.validate()?;
            return Ok((sheet.unwrap_or(owner), r));
        }
        let intersection = parts(text, ' ')?;
        if intersection.len() > 1 {
            let (sheet, mut result) =
                Self::resolve_text(session, owner, intersection.first().copied().unwrap_or_default(), relative, depth + 1, work)?;
            for part in intersection.iter().skip(1) {
                let (other, r) = Self::resolve_text(session, owner, part, relative, depth + 1, work)?;
                if other != sheet {
                    return Err(failure(1004, "intersection across worksheets"));
                }
                let mut areas = Vec::new();
                for a in &result.areas {
                    for b in &r.areas {
                        if let Some(i) = a.intersection(b) {
                            if !areas.contains(&i) {
                                areas.push(i);
                            }
                            if areas.len() > 64 {
                                return Err(failure(7, "area budget"));
                            }
                        }
                    }
                }
                result = Range { areas, unit: Unit::Cells };
                result.validate()?;
            }
            return Ok((sheet, result));
        }
        let (qualifier, body) = split_sheet(text);
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let sheet = match qualifier.as_deref() {
            Some(name) => doc.wb.sheet_index(name).ok_or_else(|| failure(1004, "worksheet not found"))?,
            None => owner,
        };
        if let Some(rect) = RangeRef::parse(body) {
            let range: Range = rect.into();
            let range = if qualifier.is_none() {
                if let Some(origin) = relative { range.offset(i64::from(origin.row), i64::from(origin.col))? } else { range }
            } else {
                range
            };
            range.validate()?;
            return Ok((sheet, range));
        }
        if let Some(name) = doc
            .wb
            .names
            .iter()
            .find(|n| n.scope == Some(sheet) && n.name.eq_ignore_ascii_case(body))
            .or_else(|| doc.wb.names.iter().find(|n| n.scope.is_none() && n.name.eq_ignore_ascii_case(body)))
        {
            return Self::resolve_text(session, sheet, &name.formula, None, depth + 1, work);
        }
        let ends = parts(body, ':')?;
        if ends.len() == 2 {
            let (a, first) = Self::resolve_text(session, sheet, ends.first().copied().unwrap_or_default(), relative, depth + 1, work)?;
            let (b, last) = Self::resolve_text(session, sheet, ends.get(1).copied().unwrap_or_default(), relative, depth + 1, work)?;
            if a != b {
                return Err(failure(1004, "endpoints on different worksheets"));
            }
            return Ok((a, first.rectangle()?.union(&last.rectangle()?).into()));
        }
        Err(failure(1004, "invalid or unresolved range"))
    }
    fn resolve_value(&self, session: &Session, sheet: usize, value: &Value, relative: Option<CellRef>) -> HostResult<(usize, Range)> {
        if let Value::Object(reference) = value {
            if reference.host != self.identity || reference.generation != self.doc {
                return Err(failure(91, "stale or foreign endpoint"));
            }
            match self.object(session, *reference)? {
                Object::Range(s, r) => Ok((s, r)),
                Object::DeletedRange(_) => Err(failure(1004, "deleted Range endpoint")),
                _ => Err(failure(13, "Range endpoint required")),
            }
        } else {
            Self::resolve_text(session, sheet, &string(value)?, relative, 0, &mut 0)
        }
    }
    pub(crate) fn construct_range(
        &mut self,
        session: &Session,
        sheet: usize,
        relative: Option<CellRef>,
        qualified: bool,
        call: &HostCall,
    ) -> HostResult<Value> {
        let args = arguments(call, &["Cell1", "Cell2"])?;
        let first = required(&args, 0)?;
        let (owner, mut range) = self.resolve_value(session, sheet, first, relative)?;
        if !qualified && owner != sheet {
            return Err(failure(1004, "range belongs to another worksheet"));
        }
        if let Some(second) = args.get(1).filter(|v| !matches!(v, Value::Missing)) {
            let (other, last) = self.resolve_value(session, sheet, second, relative)?;
            if other != owner {
                return Err(failure(1004, "range endpoints have different owners"));
            }
            range = range.rectangle()?.union(&last.rectangle()?).into();
        } else if matches!(first, Value::Object(_)) {
            return Err(failure(1004, "single Range endpoint requires an address"));
        }
        self.intern(Object::Range(owner, range))
    }
}
