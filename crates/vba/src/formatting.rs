//! Live formatting objects use the same range identity, protection and command history.
use super::objects::{arguments, numeric_index, required};
use super::*;
use gridcraft_engine::{StyleChange, model::Style};
impl Bridge {
    fn styles<'a>(&self, session: &'a Session, sheet: usize, range: &ranges::Range) -> HostResult<Vec<&'a Style>> {
        let rect = range.rectangle()?;
        values::dimensions(rect)?;
        let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
        let sheet = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet gone"))?;
        Ok(rect.iter().map(|c| doc.wb.styles.get(sheet.style_id(c))).collect())
    }
    fn style_write(&self, session: &mut Session, sheet: usize, range: &ranges::Range, change: StyleChange) -> HostResult<Value> {
        values::dimensions(range.rectangle()?)?;
        session.set_range_style(sheet, range.rectangle()?, change).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    pub(crate) fn number_format(&self, session: &mut Session, sheet: usize, range: &ranges::Range, call: &HostCall) -> HostResult<Value> {
        if matches!(call.kind, InvokeKind::Get | InvokeKind::Call) {
            arguments(call, &[])?;
            let values = self.styles(session, sheet, range)?;
            let Some(first) = values.first() else {
                return Err(failure(1004, "empty format range"));
            };
            if values.iter().any(|s| s.num_fmt != first.num_fmt) {
                return Ok(Value::Null);
            }
            values::vba_text(first.num_fmt.as_str())
        } else if call.kind == InvokeKind::Let {
            let args = arguments(call, &["Value"])?;
            let text = values::text(required(&args, 0)?)?;
            self.style_write(session, sheet, range, StyleChange::NumberFormat(text))
        } else {
            Err(failure(13, "NumberFormat requires Let"))
        }
    }
    pub(crate) fn format(&self, session: &mut Session, sheet: usize, range: &ranges::Range, font: bool, call: &HostCall) -> HostResult<Value> {
        let member = call.member.to_ascii_lowercase();
        if member == "parent" && matches!(call.kind, InvokeKind::Get | InvokeKind::Call) {
            return Err(failure(438, "format Parent is not supported"));
        }
        let color = |color: gridcraft_engine::model::Color| -> HostResult<Value> {
            let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
            let [r, g, b] = color.resolve(&doc.wb.theme).unwrap_or([0, 0, 0]);
            Ok(Value::Long(i32::from(r) | (i32::from(g) << 8) | (i32::from(b) << 16)))
        };
        let get = |style: &Style| match (font, member.as_str()) {
            (true, "bold") => Ok(Value::Boolean(style.font.bold)),
            (true, "italic") => Ok(Value::Boolean(style.font.italic)),
            (true, "name") => values::vba_text(&style.font.name),
            (true, "size") => Ok(Value::Double(f64::from(style.font.size))),
            (true, "color") => color(style.font.color),
            (false, "color") => color(style.fill.fg),
            _ => Err(failure(438, "unsupported formatting member")),
        };
        if matches!(call.kind, InvokeKind::Get | InvokeKind::Call) {
            arguments(call, &[])?;
            let styles = self.styles(session, sheet, range)?;
            let first = get(styles.first().ok_or_else(|| failure(1004, "empty format range"))?)?;
            for style in styles.iter().skip(1) {
                if get(style)? != first {
                    return Ok(Value::Null);
                }
            }
            return Ok(first);
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(13, "formatting assignment requires Let"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        let change = match (font, member.as_str()) {
            (true, "bold") => StyleChange::Bold(address::boolean(value)?),
            (true, "italic") => StyleChange::Italic(address::boolean(value)?),
            (true, "name") => StyleChange::FontName(values::text(value)?),
            (true, "size") => {
                let Value::Double(n) = vba_runtime::bind(vba_values::ScalarKind::Double, value).map_err(values::runtime_error)? else {
                    return Err(failure(13, "size required"));
                };
                if !(1.0..=409.0).contains(&n) {
                    return Err(failure(1004, "invalid font size"));
                }
                StyleChange::FontSize(n as f32)
            }
            (_, "color") => {
                let n = u32::try_from(numeric_index(value)?).map_err(|_| failure(1004, "invalid RGB color"))?;
                if font { StyleChange::FontColor(n) } else { StyleChange::InteriorColor(n) }
            }
            _ => return Err(failure(438, "unsupported formatting member")),
        };
        self.style_write(session, sheet, range, change)
    }
}
