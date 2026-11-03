//! OM05: style/geometry projections use the workbook's rendered and persisted state.
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required};
use super::{Bridge, HostCall, HostResult, InvokeKind, MAX_CELLS, Object, Session, Value, failure, ranges, values};
use gridcraft_engine::{
    core::{CellRef, MAX_COLS, MAX_ROWS, RangeRef},
    model::{Color, HAlign, PatternType, Style, Underline, VAlign, VertAlign},
};
use serde_json::{Value as Json, json};

pub(crate) fn number(value: &Value) -> HostResult<f64> {
    let Value::Double(n) = vba_runtime::bind(vba_values::ScalarKind::Double, value).map_err(values::runtime_error)? else {
        return Err(failure(13, "numeric value required"));
    };
    if !n.is_finite() {
        return Err(failure(1004, "finite number required"));
    }
    Ok(n)
}
pub(crate) fn rgb(value: &Value) -> HostResult<Color> {
    let n = u32::try_from(numeric_index(value)?).ok().filter(|n| *n <= 0xffffff).ok_or_else(|| failure(1004, "invalid RGB color"))?;
    Ok(Color::rgb((n & 255) as u8, ((n >> 8) & 255) as u8, ((n >> 16) & 255) as u8))
}
pub(crate) fn ole(color: Color, session: &Session) -> HostResult<Value> {
    let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
    let [r, g, b] = color.resolve(&doc.wb.theme).unwrap_or([0, 0, 0]);
    Ok(Value::Long(i32::from(r) | (i32::from(g) << 8) | (i32::from(b) << 16)))
}
pub(crate) fn theme_index(n: i32) -> HostResult<u8> {
    match n {
        1 => Ok(1),
        2 => Ok(0),
        3 => Ok(3),
        4 => Ok(2),
        5..=12 => Ok((n - 1) as u8),
        _ => Err(failure(1004, "invalid ThemeColor")),
    }
}
pub(crate) fn theme_number(i: u8) -> i32 {
    match i {
        0 => 2,
        1 => 1,
        2 => 4,
        3 => 3,
        _ => i32::from(i) + 1,
    }
}
fn pattern(n: i32) -> HostResult<PatternType> {
    Ok(match n {
        -4142 => PatternType::None,
        1 => PatternType::Solid,
        -4125 => PatternType::Gray50,
        -4126 => PatternType::Gray75,
        -4124 => PatternType::Gray25,
        17 => PatternType::Gray125,
        18 => PatternType::Gray0625,
        -4128 => PatternType::DarkHorizontal,
        -4166 => PatternType::DarkVertical,
        -4121 => PatternType::DarkDown,
        -4162 => PatternType::DarkUp,
        9 => PatternType::DarkGrid,
        10 => PatternType::DarkTrellis,
        11 => PatternType::LightHorizontal,
        12 => PatternType::LightVertical,
        13 => PatternType::LightDown,
        14 => PatternType::LightUp,
        15 => PatternType::LightGrid,
        16 => PatternType::LightTrellis,
        _ => return Err(failure(1004, "unsupported fill pattern")),
    })
}
fn pattern_number(p: PatternType) -> i32 {
    match p {
        PatternType::None => -4142,
        PatternType::Solid => 1,
        PatternType::Gray50 => -4125,
        PatternType::Gray75 => -4126,
        PatternType::Gray25 => -4124,
        PatternType::Gray125 => 17,
        PatternType::Gray0625 => 18,
        PatternType::DarkHorizontal => -4128,
        PatternType::DarkVertical => -4166,
        PatternType::DarkDown => -4121,
        PatternType::DarkUp => -4162,
        PatternType::DarkGrid => 9,
        PatternType::DarkTrellis => 10,
        PatternType::LightHorizontal => 11,
        PatternType::LightVertical => 12,
        PatternType::LightDown => 13,
        PatternType::LightUp => 14,
        PatternType::LightGrid => 15,
        PatternType::LightTrellis => 16,
    }
}
impl Bridge {
    pub(crate) fn style_patch(&self, session: &mut Session, sheet: usize, rect: RangeRef, patch: Json) -> HostResult<Value> {
        values::dimensions(rect)?;
        session.execute("home.formatCells", json!({"sheet":sheet,"range":rect.a1(),"style":patch})).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    pub(crate) fn extended_font(
        &mut self,
        session: &mut Session,
        sheet: usize,
        range: &ranges::Range,
        font: bool,
        call: &HostCall,
    ) -> Option<HostResult<Value>> {
        let member = call.member.to_ascii_lowercase();
        if !matches!(
            member.as_str(),
            "parent"
                | "application"
                | "underline"
                | "strikethrough"
                | "superscript"
                | "subscript"
                | "themecolor"
                | "tintandshade"
                | "colorindex"
                | "pattern"
                | "patterncolor"
                | "patterncolorthemecolor"
                | "patterncolorindex"
        ) {
            return None;
        }
        Some((|| {
            if reads(call) {
                arguments(call, &[])?;
                if member == "parent" {
                    return self.intern(Object::Range(sheet, range.clone()));
                }
                if member == "application" {
                    return self.intern(Object::Application);
                }
            }
            let styles = self.styles(session, sheet, range)?;
            let color = |s: &Style| {
                if font {
                    s.font.color
                } else if member.starts_with("patterncolor") {
                    s.fill.bg
                } else {
                    s.fill.fg
                }
            };
            let get = |s: &Style| -> HostResult<Value> {
                Ok(match member.as_str() {
                    "underline" if font => Value::Long(match s.font.underline {
                        Underline::None => -4142,
                        Underline::Single => 2,
                        Underline::Double => -4119,
                        Underline::SingleAccounting => 4,
                        Underline::DoubleAccounting => 5,
                    }),
                    "strikethrough" if font => Value::Boolean(s.font.strike),
                    "superscript" if font => Value::Boolean(s.font.vert == VertAlign::Superscript),
                    "subscript" if font => Value::Boolean(s.font.vert == VertAlign::Subscript),
                    "pattern" if !font => Value::Long(pattern_number(s.fill.pattern)),
                    "patterncolor" if !font => return ole(s.fill.bg, session),
                    "themecolor" | "patterncolorthemecolor" => match color(s) {
                        Color::Theme(i, _) => Value::Long(theme_number(i)),
                        _ => return Err(failure(1004, "color has no theme identity")),
                    },
                    "tintandshade" => Value::Double(match color(s) {
                        Color::Theme(_, t) => f64::from(t) / 1000.0,
                        _ => 0.0,
                    }),
                    "colorindex" | "patterncolorindex" => {
                        if color(s) == Color::Auto {
                            Value::Long(if font { -4105 } else { -4142 })
                        } else {
                            return Err(failure(1004, "indexed palettes are not represented by GridCraft"));
                        }
                    }
                    _ => return Err(failure(438, "unsupported format member")),
                })
            };
            if reads(call) {
                let first = get(styles.first().ok_or_else(|| failure(1004, "empty style range"))?)?;
                for s in styles.iter().skip(1) {
                    if get(s)? != first {
                        return Ok(Value::Null);
                    }
                }
                return Ok(first);
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(13, "format assignment requires Let"));
            }
            let args = arguments(call, &["Value"])?;
            let value = required(&args, 0)?;
            let patch = match member.as_str() {
                "underline" if font => {
                    json!({"font":{"underline":match numeric_index(value)?{-4142=>Underline::None,2=>Underline::Single,-4119=>Underline::Double,4=>Underline::SingleAccounting,5=>Underline::DoubleAccounting,_=>return Err(failure(1004,"invalid underline"))}}})
                }
                "strikethrough" if font => json!({"font":{"strike":super::address::boolean(value)?}}),
                "superscript" | "subscript" if font => {
                    json!({"font":{"vert":if super::address::boolean(value)?{if member=="superscript"{VertAlign::Superscript}else{VertAlign::Subscript}}else{VertAlign::Baseline}}})
                }
                "pattern" if !font => json!({"fill":{"pattern":pattern(numeric_index(value)?)?}}),
                "patterncolor" if !font => json!({"fill":{"bg":rgb(value)?}}),
                "colorindex" | "patterncolorindex" => {
                    let n = numeric_index(value)?;
                    if !matches!(n, -4142 | -4105) {
                        return Err(failure(1004, "indexed color palettes are unavailable"));
                    }
                    if font {
                        json!({"font":{"color":Color::Auto}})
                    } else if member == "patterncolorindex" {
                        json!({"fill":{"bg":Color::Auto}})
                    } else {
                        json!({"fill":{"fg":Color::Auto,"pattern":PatternType::None}})
                    }
                }
                "themecolor" | "tintandshade" | "patterncolorthemecolor" => {
                    let current = color(styles.first().ok_or_else(|| failure(1004, "empty style range"))?);
                    if styles.iter().any(|s| color(s) != current) {
                        return Err(failure(1004, "theme modification requires one color identity"));
                    }
                    let next = if member == "tintandshade" {
                        let tint = number(value)?;
                        if !(-1.0..=1.0).contains(&tint) {
                            return Err(failure(1004, "invalid tint"));
                        }
                        let Color::Theme(i, _) = current else {
                            return Err(failure(1004, "RGB tint identity is unavailable"));
                        };
                        Color::Theme(i, (tint * 1000.0).round() as i16)
                    } else {
                        Color::Theme(
                            theme_index(numeric_index(value)?)?,
                            match current {
                                Color::Theme(_, t) => t,
                                _ => 0,
                            },
                        )
                    };
                    if font {
                        json!({"font":{"color":next}})
                    } else if member.starts_with("patterncolor") {
                        json!({"fill":{"bg":next}})
                    } else {
                        json!({"fill":{"fg":next,"pattern":PatternType::Solid}})
                    }
                }
                _ => return Err(failure(438, "unsupported format assignment")),
            };
            self.style_patch(session, sheet, range.rectangle()?, patch)
        })())
    }
    pub(crate) fn format_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if let Object::Feature(f) = object
            && matches!(f.kind, "Borders" | "Border")
        {
            return Some(self.borders_api(session, f, call));
        }
        let Object::Range(sheet, range) = object else {
            return None;
        };
        let sheet = *sheet;
        let member = call.member.to_ascii_lowercase();
        if member == "borders" && reads(call) {
            return Some({
                let mut f = Feature::new("Borders", Some(sheet));
                f.range = Some(range.clone());
                if call.arguments.is_empty() { self.feature(f) } else { self.borders_api(session, &f, call) }
            });
        }
        if matches!(member.as_str(), "rowheight" | "columnwidth" | "height" | "width" | "hidden" | "autofit") {
            return Some(self.geometry_api(session, sheet, range, call, &member));
        }
        if matches!(member.as_str(), "merge" | "unmerge" | "mergecells" | "mergearea") {
            return Some(self.merge_api(session, sheet, range, call, &member));
        }
        if !matches!(
            member.as_str(),
            "horizontalalignment"
                | "verticalalignment"
                | "wraptext"
                | "shrinktofit"
                | "indentlevel"
                | "orientation"
                | "locked"
                | "formulahidden"
                | "numberformatlocal"
        ) {
            return None;
        }
        Some((|| {
            if member == "numberformatlocal" {
                if self.formula_locale != gridcraft_engine::formula::FormulaLocale::En {
                    return Err(failure(1004, "localized number-format codes are unavailable"));
                }
                return self.number_format(session, sheet, range, call);
            }
            let get = |s: &Style| match member.as_str() {
                "horizontalalignment" => Value::Long(match s.align.h {
                    HAlign::General => 1,
                    HAlign::Left => -4131,
                    HAlign::Center => -4108,
                    HAlign::Right => -4152,
                    HAlign::Fill => 5,
                    HAlign::Justify => -4130,
                    HAlign::CenterAcross => 7,
                    HAlign::Distributed => -4117,
                }),
                "verticalalignment" => Value::Long(match s.align.v {
                    VAlign::Top => -4160,
                    VAlign::Center => -4108,
                    VAlign::Bottom => -4107,
                    VAlign::Justify => -4130,
                    VAlign::Distributed => -4117,
                }),
                "wraptext" => Value::Boolean(s.align.wrap),
                "shrinktofit" => Value::Boolean(s.align.shrink),
                "indentlevel" => Value::Long(i32::from(s.align.indent)),
                "orientation" => Value::Long(i32::from(s.align.rotation)),
                "locked" => Value::Boolean(s.protection.locked),
                _ => Value::Boolean(s.protection.hidden),
            };
            if reads(call) {
                arguments(call, &[])?;
                let styles = self.styles(session, sheet, range)?;
                let first = get(styles.first().ok_or_else(|| failure(1004, "empty range"))?);
                return Ok(if styles.iter().all(|s| get(s) == first) { first } else { Value::Null });
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(13, "format assignment requires Let"));
            }
            let args = arguments(call, &["Value"])?;
            let value = required(&args, 0)?;
            let patch = match member.as_str() {
                "horizontalalignment" => {
                    json!({"align":{"h":match numeric_index(value)?{1=>HAlign::General,-4131=>HAlign::Left,-4108=>HAlign::Center,-4152=>HAlign::Right,5=>HAlign::Fill,-4130=>HAlign::Justify,7=>HAlign::CenterAcross,-4117=>HAlign::Distributed,_=>return Err(failure(1004,"invalid horizontal alignment"))}}})
                }
                "verticalalignment" => {
                    json!({"align":{"v":match numeric_index(value)?{-4160=>VAlign::Top,-4108=>VAlign::Center,-4107=>VAlign::Bottom,-4130=>VAlign::Justify,-4117=>VAlign::Distributed,_=>return Err(failure(1004,"invalid vertical alignment"))}}})
                }
                "wraptext" => json!({"align":{"wrap":super::address::boolean(value)?}}),
                "shrinktofit" => json!({"align":{"shrink":super::address::boolean(value)?}}),
                "indentlevel" => {
                    let n = numeric_index(value)?;
                    if !(0..=250).contains(&n) {
                        return Err(failure(1004, "invalid indent"));
                    }
                    json!({"align":{"indent":n}})
                }
                "orientation" => {
                    let n = numeric_index(value)?;
                    if !(-90..=90).contains(&n) && n != 255 {
                        return Err(failure(1004, "invalid orientation"));
                    }
                    json!({"align":{"rotation":n}})
                }
                "locked" => json!({"protection":{"locked":super::address::boolean(value)?}}),
                _ => json!({"protection":{"hidden":super::address::boolean(value)?}}),
            };
            self.style_patch(session, sheet, range.rectangle()?, patch)
        })())
    }
    fn geometry_api(&self, session: &mut Session, sheet: usize, range: &ranges::Range, call: &HostCall, member: &str) -> HostResult<Value> {
        let rect = range.rectangle()?;
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let rows =
            member == "rowheight" || member == "height" || range.unit == ranges::Unit::Rows || rect.start.col == 0 && rect.end.col == MAX_COLS - 1;
        let whole = if rows { rect.start.col == 0 && rect.end.col == MAX_COLS - 1 } else { rect.start.row == 0 && rect.end.row == MAX_ROWS - 1 };
        if matches!(member, "hidden" | "autofit") && !whole {
            return Err(failure(1004, "Hidden/AutoFit requires entire rows or columns"));
        }
        let (a, b) = if rows { (rect.start.row, rect.end.row) } else { (rect.start.col, rect.end.col) };
        if u64::from(b - a) + 1 > MAX_CELLS as u64 {
            return Err(failure(7, "line geometry budget"));
        }
        if reads(call) && member != "autofit" {
            arguments(call, &[])?;
            let items = (a..=b)
                .map(|i| {
                    if member == "hidden" {
                        Value::Boolean(if rows { sh.rows.get(&i).is_some_and(|x| x.hidden) } else { sh.cols.get(&i).is_some_and(|x| x.hidden) })
                    } else {
                        let px = f64::from(if rows { sh.row_height(i) } else { sh.col_width(i) });
                        Value::Double(if member == "columnwidth" { if px < 12.0 { px / 12.0 } else { (px - 5.0) / 7.0 } } else { px * 0.75 })
                    }
                })
                .collect::<Vec<_>>();
            if matches!(member, "width" | "height") {
                return Ok(Value::Double(items.iter().filter_map(|v| if let Value::Double(n) = v { Some(n) } else { None }).sum()));
            }
            let first = items.first().cloned().ok_or_else(|| failure(1004, "empty geometry"))?;
            return Ok(if items.iter().all(|v| *v == first) { first } else { Value::Null });
        }
        let lines = if rows {
            format!("{}:{}", a + 1, b + 1)
        } else {
            format!("{}:{}", gridcraft_engine::core::col_to_letters(a), gridcraft_engine::core::col_to_letters(b))
        };
        let mut params = json!({"sheet":sheet});
        params[if rows { "rows" } else { "cols" }] = json!(lines);
        let command = if member == "autofit" {
            arguments(call, &[])?;
            if rows { "home.autofitRowHeight" } else { "home.autofitColumnWidth" }
        } else {
            if call.kind != InvokeKind::Let {
                return Err(failure(438, "read-only geometry"));
            }
            let args = arguments(call, &["Value"])?;
            let v = required(&args, 0)?;
            match member {
                "rowheight" => {
                    params["points"] = json!(number(v)?);
                    "home.rowHeight"
                }
                "columnwidth" => {
                    params["chars"] = json!(number(v)?);
                    "home.columnWidth"
                }
                "hidden" => match (rows, super::address::boolean(v)?) {
                    (true, true) => "home.hideRows",
                    (true, false) => "home.unhideRows",
                    (false, true) => "home.hideColumns",
                    (false, false) => "home.unhideColumns",
                },
                _ => return Err(failure(438, "read-only geometry")),
            }
        };
        session.execute(command, params).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    fn merge_api(&mut self, session: &mut Session, sheet: usize, range: &ranges::Range, call: &HostCall, member: &str) -> HostResult<Value> {
        let rect = range.rectangle()?;
        values::dimensions(rect)?;
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if member == "mergearea" && reads(call) {
            arguments(call, &[])?;
            if !rect.is_single() {
                return Err(failure(1004, "MergeArea requires one cell"));
            }
            let area = sh.merges.iter().find(|r| r.contains(rect.start)).copied().unwrap_or(rect);
            return self.intern(Object::Range(sheet, area.into()));
        }
        if member == "mergecells" && reads(call) {
            arguments(call, &[])?;
            let flags = rect.iter().map(|at| sh.merges.iter().any(|r| r.contains(at))).collect::<Vec<_>>();
            return Ok(if flags.iter().all(|v| *v) {
                Value::Boolean(true)
            } else if flags.iter().all(|v| !*v) {
                Value::Boolean(false)
            } else {
                Value::Null
            });
        }
        let command = match member {
            "unmerge" => {
                arguments(call, &[])?;
                "home.unmergeCells"
            }
            "merge" => {
                let args = arguments(call, &["Across"])?;
                if !super::features::absent(&args[0]) && super::address::boolean(&args[0])? { "home.mergeAcross" } else { "home.mergeCells" }
            }
            "mergecells" if call.kind == InvokeKind::Let => {
                let args = arguments(call, &["Value"])?;
                if super::address::boolean(required(&args, 0)?)? { "home.mergeCells" } else { "home.unmergeCells" }
            }
            _ => return Err(failure(438, "unsupported merge operation")),
        };
        let changes = self.prepare_change(session, sheet, range.clone())?;
        session.execute(command, json!({"sheet":sheet,"range":rect.a1()})).map_err(|e| failure(1004, e.to_string()))?;
        self.queue_changes(changes);
        Ok(Value::Empty)
    }
    fn borders_api(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "border owner deleted"))?;
        let range = f.range.as_ref().ok_or_else(|| failure(91, "border range deleted"))?;
        let rect = range.rectangle()?;
        values::dimensions(rect)?;
        let member = call.member.to_ascii_lowercase();
        if f.kind == "Borders" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(
                    (5..=12)
                        .map(|part| {
                            let mut f = f.clone();
                            f.kind = "Border";
                            f.part = part;
                            f.object()
                        })
                        .collect(),
                );
            }
            if member == "count" && reads(call) {
                arguments(call, &[])?;
                return Ok(Value::Long(8));
            }
            if reads(call) && matches!(member.as_str(), "" | "item" | "borders") {
                let args = arguments(call, &["Index"])?;
                let part = numeric_index(required(&args, 0)?)?;
                if !(5..=12).contains(&part) {
                    return Err(failure(1004, "invalid border index"));
                }
                let mut f = f.clone();
                f.kind = "Border";
                f.part = part;
                return self.feature(f);
            }
        }
        if reads(call) && member == "parent" {
            arguments(call, &[])?;
            return self.intern(Object::Range(sheet, range.clone()));
        }
        if reads(call) && member == "application" {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        let parts = if f.kind == "Borders" { (5..=12).collect::<Vec<_>>() } else { vec![f.part] };
        let targets = parts
            .into_iter()
            .filter_map(|part| {
                let (a, b, side) = match part {
                    5 => (rect.start, rect.end, "diag_down"),
                    6 => (rect.start, rect.end, "diag_up"),
                    7 => (rect.start, CellRef::new(rect.end.row, rect.start.col), "left"),
                    8 => (rect.start, CellRef::new(rect.start.row, rect.end.col), "top"),
                    9 => (CellRef::new(rect.end.row, rect.start.col), rect.end, "bottom"),
                    10 => (CellRef::new(rect.start.row, rect.end.col), rect.end, "right"),
                    11 if rect.width() > 1 => (rect.start, CellRef::new(rect.end.row, rect.end.col - 1), "right"),
                    12 if rect.height() > 1 => (rect.start, CellRef::new(rect.end.row - 1, rect.end.col), "bottom"),
                    _ => return None,
                };
                Some((RangeRef::new(a, b), side))
            })
            .collect::<Vec<_>>();
        if reads(call) {
            arguments(call, &[])?;
            let mut result = None;
            for (target, side) in &targets {
                for style in self.styles(session, sheet, &(*target).into())? {
                    let line = match *side {
                        "left" => style.borders.left,
                        "right" => style.borders.right,
                        "top" => style.borders.top,
                        "bottom" => style.borders.bottom,
                        "diag_down" => style.borders.diag_down,
                        _ => style.borders.diag_up,
                    };
                    let value = match member.as_str() {
                        "color" => ole(line.color, session)?,
                        "linestyle" => Value::Long(line_style_number(line.style)),
                        "weight" => Value::Long(weight_number(line.style)),
                        _ => return Err(failure(438, "unsupported border member")),
                    };
                    if result.as_ref().is_some_and(|v| *v != value) {
                        return Ok(Value::Null);
                    }
                    result = Some(value);
                }
            }
            return Ok(result.unwrap_or(Value::Null));
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(13, "border assignment requires Let"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        let patch = match member.as_str() {
            "color" => json!({"color":rgb(value)?}),
            "linestyle" => json!({"style":line_style(numeric_index(value)?)?}),
            "weight" => {
                json!({"style":match numeric_index(value)?{1=>gridcraft_engine::model::BorderStyle::Hair,2=>gridcraft_engine::model::BorderStyle::Thin,-4138=>gridcraft_engine::model::BorderStyle::Medium,4=>gridcraft_engine::model::BorderStyle::Thick,_=>return Err(failure(1004,"invalid border weight"))}})
            }
            _ => return Err(failure(438, "unsupported border assignment")),
        };
        // Preflight protection for the entire footprint before issuing boundary patches.
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        if doc.wb.sheet(sheet).is_some_and(|sh| sh.is_protected() && sh.protection.as_ref().is_some_and(|p| !p.format_cells)) {
            return Err(failure(1004, "sheet protection forbids formatting"));
        }
        for (target, side) in targets {
            self.style_patch(session, sheet, target, json!({"borders":{side:patch}}))?;
        }
        Ok(Value::Empty)
    }
}
fn line_style(n: i32) -> HostResult<gridcraft_engine::model::BorderStyle> {
    use gridcraft_engine::model::BorderStyle as B;
    Ok(match n {
        -4142 => B::None,
        1 => B::Thin,
        -4115 => B::Dashed,
        -4118 => B::Dotted,
        -4119 => B::Double,
        4 => B::DashDot,
        5 => B::DashDotDot,
        13 => B::SlantDashDot,
        _ => return Err(failure(1004, "invalid border line style")),
    })
}
fn line_style_number(style: gridcraft_engine::model::BorderStyle) -> i32 {
    use gridcraft_engine::model::BorderStyle as B;
    match style {
        B::None => -4142,
        B::Dashed | B::MediumDashed => -4115,
        B::Dotted => -4118,
        B::Double => -4119,
        B::DashDot | B::MediumDashDot => 4,
        B::DashDotDot | B::MediumDashDotDot => 5,
        B::SlantDashDot => 13,
        _ => 1,
    }
}
fn weight_number(style: gridcraft_engine::model::BorderStyle) -> i32 {
    use gridcraft_engine::model::BorderStyle as B;
    match style {
        B::Hair => 1,
        B::Medium | B::MediumDashed | B::MediumDashDot | B::MediumDashDotDot => -4138,
        B::Thick => 4,
        _ => 2,
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    vec![
        super::MemberDescriptor {
            owner: "Range",
            name: "HorizontalAlignment",
            signature: "Property HorizontalAlignment As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "VerticalAlignment",
            signature: "Property VerticalAlignment As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "WrapText",
            signature: "Property WrapText As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "ShrinkToFit",
            signature: "Property ShrinkToFit As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "IndentLevel",
            signature: "Property IndentLevel As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Orientation",
            signature: "Property Orientation As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Locked",
            signature: "Property Locked As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "FormulaHidden",
            signature: "Property FormulaHidden As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "NumberFormatLocal",
            signature: "Property NumberFormatLocal As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "RowHeight",
            signature: "Property RowHeight As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "ColumnWidth",
            signature: "Property ColumnWidth As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Hidden",
            signature: "Property Hidden As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "MergeCells",
            signature: "Property MergeCells As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Width",
            signature: "Property Get Width() As Double",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Height",
            signature: "Property Get Height() As Double",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "MergeArea",
            signature: "Property Get MergeArea() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Borders",
            signature: "Property Get Borders() As Borders",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "Merge",
            signature: "Sub Merge(Optional Across As Variant)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "UnMerge",
            signature: "Sub UnMerge()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Range",
            name: "AutoFit",
            signature: "Sub AutoFit()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Borders",
            name: "Item",
            signature: "Property Get Item(Index As Long) As Border",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Borders",
            name: "Count",
            signature: "Property Get Count() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "Parent",
            signature: "Property Get Parent() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "ThemeColor",
            signature: "Property ThemeColor As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "TintAndShade",
            signature: "Property TintAndShade As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "ColorIndex",
            signature: "Property ColorIndex As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Interior",
            name: "Parent",
            signature: "Property Get Parent() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Interior",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Interior",
            name: "ThemeColor",
            signature: "Property ThemeColor As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Interior",
            name: "TintAndShade",
            signature: "Property TintAndShade As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Interior",
            name: "ColorIndex",
            signature: "Property ColorIndex As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "Underline",
            signature: "Property Underline As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "Strikethrough",
            signature: "Property Strikethrough As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "Superscript",
            signature: "Property Superscript As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Font",
            name: "Subscript",
            signature: "Property Subscript As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Interior",
            name: "Pattern",
            signature: "Property Pattern As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Interior",
            name: "PatternColor",
            signature: "Property PatternColor As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Interior",
            name: "PatternColorIndex",
            signature: "Property PatternColorIndex As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Borders",
            name: "Color",
            signature: "Property Color As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Borders",
            name: "LineStyle",
            signature: "Property LineStyle As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Borders",
            name: "Weight",
            signature: "Property Weight As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Borders",
            name: "Parent",
            signature: "Property Get Parent() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Borders",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Border",
            name: "Color",
            signature: "Property Color As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Border",
            name: "LineStyle",
            signature: "Property LineStyle As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Border",
            name: "Weight",
            signature: "Property Weight As Variant",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Border",
            name: "Parent",
            signature: "Property Get Parent() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
        super::MemberDescriptor {
            owner: "Border",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/format_api.rs::Bridge::format_api",
        },
    ]
}
