//! Native worksheet Form Control commands. All linked writes use the ordinary edit/recalc path.
use super::{
    CellRef, CommandSpec, EngineError, Json, Result, Session, bad, cell_param_on, cmd, edit, f64_param, has_doc, str_param, target_sheet, u32_param,
};
use gridcraft_model::form_controls::{effective_control, linked_write, radio_group, resolve_binding};
use gridcraft_model::{Anchor, AnchorMode, Color, FormControl, FormControlKind, Shape, ShapeFormat, ShapeKind};
use serde_json::json;
use std::collections::BTreeSet;
pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            "control.add",
            "Insert Form Control",
            ["Developer", "Controls"],
            None,
            "{sheet?, type: Excel FormControl enum, name?, at?, left?, top?, width?, height?, caption?}",
            has_doc,
            add
        ),
        cmd!(
            "control.set",
            "Format Control",
            [],
            None,
            "{sheet?, id, name?, caption?, value?, enabled?, visible?, printObject?, min?, max?, smallChange?, largeChange?, linkedCell?, listFillRange?, items?, selected?, multiSelect?, horizontal?, dropLines?}",
            has_doc,
            set
        ),
        cmd!(
            "control.activate",
            "Activate Form Control",
            [],
            None,
            "{sheet?, id, value?, item?, delta?}; changes state/linked cells; returns inert macro assignment",
            has_doc,
            activate
        ),
    ]
}
fn err(s: impl Into<String>) -> EngineError {
    bad("control", s)
}
fn integer(p: &Json, key: &str) -> Result<Option<i32>> {
    p.get(key).map(|v| v.as_i64().and_then(|n| i32::try_from(n).ok()).ok_or_else(|| err(format!("invalid {key}")))).transpose()
}
fn add(s: &mut Session, p: &Json) -> Result<Json> {
    let si = target_sheet(s, p)?;
    let wb = &s.doc()?.wb;
    let sh = wb.sheet(si).ok_or_else(|| err("sheet"))?;
    if sh.is_protected() {
        return Err(err("drawing objects are protected"));
    }
    if sh.shapes.iter().filter(|s| s.format.form_control.is_some()).count() >= 1024 {
        return Err(err("control count limit"));
    }
    let kind = integer(p, "type")?.and_then(FormControlKind::from_excel).ok_or_else(|| err("unsupported worksheet control type"))?;
    let id = wb.next_object_id();
    let office = sh
        .shapes
        .iter()
        .filter_map(|s| s.format.office_id)
        .max()
        .unwrap_or(1024)
        .checked_add(1)
        .ok_or_else(|| err("native control identity exhausted"))?;
    let name = str_param(p, "name").map(str::to_owned).unwrap_or_else(|| format!("{} {office}", kind.xml()));
    if sh.shapes.iter().any(|s| s.format.office_name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(&name))) {
        return Err(err("duplicate shape/control name"));
    }
    let c = FormControl::new(kind, name.clone());
    c.validate().map_err(err)?;
    let at = cell_param_on(s, p, "at", si)?.unwrap_or(s.doc()?.selection.active);
    let left = f64_param(p, "left").unwrap_or(sh.col_left(at.col));
    let top = f64_param(p, "top").unwrap_or(sh.row_top(at.row));
    let width = f64_param(p, "width").unwrap_or(120.);
    let height = f64_param(p, "height").unwrap_or(24.);
    if [left, top, width, height].iter().any(|v| !v.is_finite() || *v < 0. || *v > 1_000_000.) {
        return Err(err("control geometry bounds"));
    }
    if width > 10000. || height > 10000. {
        return Err(err("control size bounds"));
    }
    let cell = CellRef::new(sh.row_at(top), sh.col_at(left));
    let caption = str_param(p, "caption").unwrap_or(&name);
    if caption.len() > 16384 || !gridcraft_model::form_controls::valid_control_text(caption) {
        return Err(err("caption bounds"));
    }
    let shape = Shape {
        identity: Default::default(),
        id,
        on_action: None,
        kind: ShapeKind::Rectangle,
        anchor: Anchor {
            cell,
            dx: (left - sh.col_left(cell.col)) as f32,
            dy: (top - sh.row_top(cell.row)) as f32,
            width: width as f32,
            height: height as f32,
            mode: AnchorMode::MoveOnly,
        },
        fill: Color::rgb(240, 240, 240),
        line: Color::rgb(100, 100, 100),
        text: caption.into(),
        format: ShapeFormat { office_id: Some(office), office_name: Some(name), form_control: Some(c), ..Default::default() },
    };
    edit(s, |cx| {
        cx.sheet_mut(si)?.shapes.push(shape);
        Ok(json!({"shape":id}))
    })
}
fn set(s: &mut Session, p: &Json) -> Result<Json> {
    change(s, p, false)
}
fn activate(s: &mut Session, p: &Json) -> Result<Json> {
    change(s, p, true)
}
fn change(s: &mut Session, p: &Json, activation: bool) -> Result<Json> {
    let si = target_sheet(s, p)?;
    let id = u32_param(p, "id").ok_or_else(|| err("control id required"))?;
    let wb = s.doc()?.wb.clone();
    let sh = wb.sheet(si).ok_or_else(|| err("sheet"))?;
    if sh.is_protected() {
        return Err(err("drawing objects are protected"));
    }
    let shape = sh.shapes.iter().find(|s| s.id == id && s.format.form_control.is_some()).ok_or_else(|| err("not a worksheet Form Control"))?.clone();
    let mut c = match effective_control(&wb, si, &shape) {
        Ok(c) => c,
        Err(_) if !activation && (p.get("linkedCell").is_some() || p.get("listFillRange").is_some() || p.get("value").is_some()) => {
            shape.format.form_control.clone().ok_or_else(|| err("control"))?
        }
        Err(e) => return Err(err(e)),
    };
    if activation && (!c.visible || !c.enabled || sh.is_protected()) {
        return Err(err("control is hidden, disabled or protected"));
    }
    let keys = p.as_object().ok_or_else(|| err("object parameters required"))?;
    let allowed = [
        "sheet",
        "id",
        "name",
        "caption",
        "value",
        "enabled",
        "visible",
        "printObject",
        "min",
        "max",
        "smallChange",
        "largeChange",
        "linkedCell",
        "listFillRange",
        "items",
        "selected",
        "multiSelect",
        "extendedSelect",
        "horizontal",
        "dropLines",
        "item",
        "delta",
    ];
    if keys.keys().any(|k| !allowed.contains(&k.as_str())) {
        return Err(err("unsupported control property"));
    }
    if activation && keys.keys().any(|k| !["sheet", "id", "item", "value", "delta", "selected"].contains(&k.as_str())) {
        return Err(err("activation cannot edit control properties"));
    }
    for (key, valid) in [
        ("min", c.kind.is_numeric()),
        ("max", c.kind.is_numeric()),
        ("smallChange", c.kind.is_numeric()),
        ("largeChange", c.kind.is_numeric()),
        ("horizontal", c.kind == FormControlKind::ScrollBar),
        ("items", c.kind.is_list()),
        ("item", c.kind.is_list()),
        ("selected", c.kind == FormControlKind::ListBox),
        ("multiSelect", c.kind == FormControlKind::ListBox),
        ("extendedSelect", c.kind == FormControlKind::ListBox),
        ("listFillRange", c.kind.is_list()),
        ("dropLines", c.kind == FormControlKind::DropDown),
        ("delta", c.kind.is_numeric()),
    ] {
        if p.get(key).is_some() && !valid {
            return Err(err(format!("{key} is not applicable to this control")));
        }
    }
    let link_field = if c.kind == FormControlKind::GroupBox { &mut c.group_link } else { &mut c.linked_cell };
    for (key, field) in [("name", &mut c.name), ("linkedCell", link_field), ("listFillRange", &mut c.list_fill_range)] {
        if let Some(value) = p.get(key) {
            *field = value.as_str().ok_or_else(|| err("control property requires text"))?.into();
        }
    }
    for (key, field) in [
        ("enabled", &mut c.enabled),
        ("visible", &mut c.visible),
        ("printObject", &mut c.print_object),
        ("multiSelect", &mut c.multi_select),
        ("extendedSelect", &mut c.extended_select),
        ("horizontal", &mut c.horizontal),
    ] {
        if let Some(value) = p.get(key) {
            *field = value.as_bool().ok_or_else(|| err("control property requires Boolean"))?;
        }
    }
    for (key, field) in [
        ("value", &mut c.value),
        ("min", &mut c.min),
        ("max", &mut c.max),
        ("smallChange", &mut c.small_change),
        ("largeChange", &mut c.large_change),
    ] {
        if let Some(value) = integer(p, key)? {
            *field = value;
        }
    }
    if let Some(lines) = integer(p, "dropLines")? {
        c.drop_lines = u32::try_from(lines).map_err(|_| err("dropLines"))?;
    }
    if !c.multi_select {
        c.extended_select = false;
    }
    if c.multi_select && !c.linked_cell.is_empty() && p.get("linkedCell").is_some() {
        return Err(err("multiselect lists cannot set LinkedCell"));
    }
    if !c.list_fill_range.is_empty() {
        if !c.kind.is_list() {
            return Err(err("ListFillRange requires a list control"));
        }
        let (_, r) = resolve_binding(&wb, si, &c.list_fill_range).map_err(err)?;
        if r.width() != 1 || r.count() > 1024 {
            return Err(err("list range bounds"));
        }
    }
    if p.get("listFillRange").is_some() && !c.list_fill_range.is_empty() {
        let (source, r) = resolve_binding(&wb, si, &c.list_fill_range).map_err(err)?;
        c.items = r.iter().map(|cell| wb.sheet(source).map(|s| s.value(cell).display()).unwrap_or_default()).collect();
        if c.value as usize > c.items.len() {
            c.value = 0;
        }
    }
    if let Some(items) = p.get("items") {
        if !c.kind.is_list() {
            return Err(err("List requires list control"));
        }
        let items = items.as_array().filter(|a| a.len() <= 1024).ok_or_else(|| err("list item count"))?;
        c.items = items
            .iter()
            .map(|v| v.as_str().filter(|s| s.len() <= 16384).map(str::to_owned).ok_or_else(|| err("list item text bounds")))
            .collect::<Result<Vec<_>>>()?;
        c.list_fill_range.clear();
        if c.value as usize > c.items.len() {
            c.value = 0;
        }
        c.selected.retain(|n| *n as usize <= c.items.len());
    }
    if let Some(selected) = p.get("selected") {
        c.selected = selected
            .as_array()
            .filter(|a| a.len() <= 1024)
            .ok_or_else(|| err("selected item count"))?
            .iter()
            .map(|v| v.as_u64().and_then(|v| u32::try_from(v).ok()).ok_or_else(|| err("selection index")))
            .collect::<Result<Vec<_>>>()?;
    }
    if activation {
        activate_state(&mut c, p)?;
    }
    if c.kind.is_numeric() && p.get("value").is_none() && c.min <= c.max {
        c.value = c.value.clamp(c.min, c.max);
    }
    c.validate().map_err(err)?;
    if c.kind.is_list() && (c.value as usize > c.items.len() || c.selected.iter().any(|n| *n == 0 || *n as usize > c.items.len())) {
        return Err(err("selection outside list"));
    }
    if c.selected.iter().copied().collect::<BTreeSet<_>>().len() != c.selected.len() {
        return Err(err("duplicate selected item"));
    }
    if sh.shapes.iter().any(|s| s.id != id && s.format.office_name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(&c.name))) {
        return Err(err("duplicate control name"));
    }
    let caption = p
        .get("caption")
        .map(|v| {
            v.as_str().filter(|s| s.len() <= 16384 && gridcraft_model::form_controls::valid_control_text(s)).ok_or_else(|| err("caption bounds"))
        })
        .transpose()?;
    let write_value = activation || p.get("value").is_some();
    if !c.group_link.is_empty() {
        let (_, r) = resolve_binding(&wb, si, &c.group_link).map_err(err)?;
        if r.count() != 1 {
            return Err(err("group link requires one cell"));
        }
    }
    if !c.linked_cell.is_empty() {
        let (_, r) = resolve_binding(&wb, si, &c.linked_cell).map_err(err)?;
        if r.count() != 1 {
            return Err(err("LinkedCell requires one cell"));
        }
    }
    let total = sh
        .shapes
        .iter()
        .filter_map(|s| s.format.form_control.as_ref())
        .map(|c| c.items.iter().map(String::len).sum::<usize>())
        .sum::<usize>()
        .saturating_add(c.items.iter().map(String::len).sum::<usize>());
    if total > 4 * 1024 * 1024 {
        return Err(err("worksheet control text budget"));
    }
    let link = if write_value { linked_write(&wb, si, &shape, &c).map_err(err)? } else { None };
    let group = if c.kind == FormControlKind::OptionButton && c.value == 1 { radio_group(sh, id) } else { vec![] };
    let entry = shape.on_action.clone();
    edit(s, |cx| {
        let sh = cx.sheet_mut(si)?;
        for other in &mut sh.shapes {
            if other.id != id
                && group.contains(&other.id)
                && let Some(c) = other.format.form_control.as_mut()
            {
                c.value = -4146;
            }
        }
        let target = sh.shapes.iter_mut().find(|s| s.id == id).ok_or_else(|| err("control replaced"))?;
        target.format.office_name = Some(c.name.clone());
        target.format.form_control = Some(c.clone());
        if let Some(caption) = caption {
            target.text = caption.into();
        }
        if let Some((sheet, cell, value)) = link {
            cx.sheet_mut(sheet)?.set_value(cell, value);
            cx.touch(sheet, cell);
        }
        Ok(json!({"shape":id,"value":c.value,"entry":entry}))
    })
}

/// Maintain Office placement against the final (including automatic heights) grid.
/// Explicit drawing edits and structural shifts already compute their own anchors.
pub(super) fn resize_attachments(before: &gridcraft_model::Workbook, after: &mut gridcraft_model::Workbook) {
    use gridcraft_model::form_controls::ControlAttachment;
    for si in 0..after.sheets.len() {
        // An explicit object edit invalidates previously captured cell-relative endpoints.
        if let Some(new) = after.sheet(si)
            && let Some(old) = before.sheets.iter().find(|s| s.identity.same(&new.identity))
        {
            let changed = new
                .shapes
                .iter()
                .filter(|s| s.format.control_attachment.is_some() && old.shapes.iter().any(|o| o.identity.same(&s.identity) && o.anchor != s.anchor))
                .map(|s| s.id)
                .collect::<Vec<_>>();
            if !changed.is_empty()
                && let Some(sh) = after.sheet_mut(si)
            {
                for s in &mut sh.shapes {
                    if changed.contains(&s.id) {
                        s.format.control_attachment = None;
                    }
                }
            }
        }
        let Some(new) = after.sheet(si) else { continue };
        let Some(old) = before.sheets.iter().find(|s| s.identity.same(&new.identity)) else { continue };
        if old.rows == new.rows && old.cols == new.cols {
            continue;
        }
        let mut updates = Vec::new();
        for shape in &new.shapes {
            if shape.format.form_control.is_none() {
                continue;
            }
            let Some(prior) = old.shapes.iter().find(|s| s.identity.same(&shape.identity)) else { continue };
            if prior.anchor != shape.anchor {
                continue;
            }
            let a = shape.anchor;
            if a.mode == AnchorMode::Absolute {
                let x = old.col_left(a.cell.col) + f64::from(a.dx);
                let y = old.row_top(a.cell.row) + f64::from(a.dy);
                let cell = CellRef::new(new.row_at(y), new.col_at(x));
                updates.push((shape.id, Anchor { cell, dx: (x - new.col_left(cell.col)) as f32, dy: (y - new.row_top(cell.row)) as f32, ..a }, None));
                continue;
            }
            let attachment = prior.format.control_attachment.clone().unwrap_or_else(|| {
                let x = old.col_left(a.cell.col) + f64::from(a.dx + a.width);
                let y = old.row_top(a.cell.row) + f64::from(a.dy + a.height);
                let to = CellRef::new(old.row_at(y), old.col_at(x));
                let ratio = |offset: f64, size: f32| if size > 0. { (offset / f64::from(size)) as f32 } else { 0. };
                ControlAttachment {
                    from: a.cell,
                    to,
                    fractions: [
                        ratio(f64::from(a.dx), old.col_width(a.cell.col)),
                        ratio(f64::from(a.dy), old.row_height(a.cell.row)),
                        ratio(x - old.col_left(to.col), old.col_width(to.col)),
                        ratio(y - old.row_top(to.row), old.row_height(to.row)),
                    ],
                }
            });
            let [dx, dy, ex, ey] = attachment.fractions;
            let dx = dx * new.col_width(attachment.from.col);
            let dy = dy * new.row_height(attachment.from.row);
            let mut a = Anchor { dx, dy, ..a };
            if a.mode == AnchorMode::MoveAndSize {
                a.width =
                    (new.col_left(attachment.to.col) + f64::from(ex * new.col_width(attachment.to.col)) - new.col_left(a.cell.col) - f64::from(dx))
                        .max(0.) as f32;
                a.height =
                    (new.row_top(attachment.to.row) + f64::from(ey * new.row_height(attachment.to.row)) - new.row_top(a.cell.row) - f64::from(dy))
                        .max(0.) as f32;
            }
            updates.push((shape.id, a, Some(attachment)));
        }
        if !updates.is_empty()
            && let Some(sh) = after.sheet_mut(si)
        {
            for (id, a, attachment) in updates {
                if let Some(s) = sh.shapes.iter_mut().find(|s| s.id == id) {
                    s.anchor = a;
                    s.format.control_attachment = attachment;
                }
            }
        }
    }
}

fn activate_state(c: &mut FormControl, p: &Json) -> Result<()> {
    match c.kind {
        FormControlKind::CheckBox => c.value = if c.value == 1 { -4146 } else { 1 },
        FormControlKind::OptionButton => c.value = 1,
        k if k.is_numeric() => {
            if let Some(delta) = integer(p, "delta")? {
                c.value = c.value.saturating_add(delta).clamp(c.min, c.max);
            }
        }
        k if k.is_list() => {
            if let Some(item) = integer(p, "item")? {
                if c.multi_select {
                    let item = u32::try_from(item).map_err(|_| err("item"))?;
                    if c.selected.contains(&item) {
                        c.selected.retain(|n| *n != item)
                    } else {
                        c.selected.push(item);
                    }
                } else {
                    c.value = item;
                }
            }
        }
        FormControlKind::Label | FormControlKind::GroupBox => return Err(err("labels/group boxes are presentation controls")),
        _ => {}
    }

    Ok(())
}
