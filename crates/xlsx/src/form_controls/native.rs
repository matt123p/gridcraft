use super::{
    BTreeSet, Budget, EXCEL, IoError, Node, SHEET, Sheet, VML, X14, bad, control_properties, normalize, optional_text, positive,
    read_worksheet_controls, shape_map, xml,
};
use gridcraft_core::{CellRef, MAX_COLS, MAX_ROWS};
use gridcraft_model::form_controls::FormControl;
use gridcraft_model::{Anchor, AnchorMode, Color, FormControlKind, Shape, ShapeFormat, ShapeKind};

pub(crate) fn decode(sheet: &Sheet) -> Result<Vec<Shape>, IoError> {
    let Some(kept) = sheet.preserved.as_deref() else { return Ok(vec![]) };
    // Reuse strict graph/identity/namespace/budget verification even for non-Button families.
    let _ = read_worksheet_controls(sheet)?;
    let Some(fragment) = kept.elements.iter().find(|e| e.slot == "controls") else { return Ok(vec![]) };
    let mut budget = Budget::default();
    let controls = xml::parse(fragment.xml.as_bytes(), &mut budget)?;
    let vml = xml::parse(&kept.vml.as_ref().ok_or_else(|| bad("missing VML"))?.data, &mut budget)?;
    let shapes = shape_map(&vml)?;
    let relationships = kept.rels.iter().map(|r| (r.id.as_str(), r)).collect();
    let parts = kept.parts.iter().map(|p| (normalize(&p.name), p)).collect();
    let mut nodes = vec![];
    controls.collect(SHEET, "control", &mut nodes)?;
    let mut result = vec![];
    let mut used_parts = BTreeSet::new();
    for node in nodes {
        let id = positive(node.attr("shapeId").ok_or_else(|| bad("shape ID"))?)?;
        let v = shapes.get(&id).ok_or_else(|| bad("VML identity"))?;
        let client = v.child(EXCEL, "ClientData")?.ok_or_else(|| bad("ClientData"))?;
        let properties = control_properties(&relationships, &parts, node)?;
        if !used_parts.insert(normalize(&properties.name)) {
            return Err(bad("aliased control property part"));
        }
        let props = xml::parse(&properties.data, &mut budget)?;
        let Some(kind) = FormControlKind::from_xml(props.attr("objectType").or_else(|| client.attr("ObjectType")).unwrap_or("")) else { continue };
        let mut control = FormControl::new(kind, node.attr("name").unwrap_or("").into());
        if props.attr("fmlaTxbx").is_some_and(|s| !s.is_empty()) || optional_text(client.child(EXCEL, "FmlaTxbx")?)?.is_some() {
            return Err(bad("label/edit-box text binding is outside the worksheet control profile"));
        }
        if kind == FormControlKind::DropDown
            && (props.attr("dropStyle").is_some_and(|s| s != "combo")
                || optional_text(client.child(EXCEL, "DropStyle")?)?.is_some_and(|s| !s.eq_ignore_ascii_case("Combo")))
        {
            return Err(bad("drop-down style is outside the standard combo profile"));
        }

        let number = |attr: &str, legacy: &str, default: i32| -> Result<i32, IoError> {
            let text = props.attr(attr).map(str::to_owned).or(optional_text(client.child(EXCEL, legacy)?)?);
            text.map(|s| s.trim().parse::<i32>().map_err(|_| bad(format!("invalid {attr}")))).unwrap_or(Ok(default))
        };
        let flag = |n: &str, default: bool| -> Result<bool, IoError> {
            match props.attr(n) {
                None => Ok(default),
                Some("true" | "1") => Ok(true),
                Some("false" | "0") => Ok(false),
                _ => Err(bad(format!("invalid {n}"))),
            }
        };
        if kind.is_numeric()
            && (props.attr("max").is_none() && client.child(EXCEL, "Max")?.is_none()
                || props.attr("page").is_none() && client.child(EXCEL, "Page")?.is_none())
        {
            return Err(bad("numeric Max/Page defaults are outside the declared native profile"));
        }
        control.min = number("min", "Min", 0)?;
        control.max = number("max", "Max", 100)?;
        control.small_change = number("inc", "Inc", 1)?;
        control.large_change = number("page", "Page", 10)?;
        control.drop_lines = u32::try_from(number("dropLines", "DropLines", 8)?).map_err(|_| bad("dropLines"))?;
        control.horizontal = kind == FormControlKind::ScrollBar && flag("horiz", legacy_flag(client, "Horiz", false)?)?;
        control.first_button = flag("firstButton", legacy_flag(client, "FirstButton", false)?)?;
        control.linked_cell = props.attr("fmlaLink").map(str::to_owned).or(optional_text(client.child(EXCEL, "FmlaLink")?)?).unwrap_or_default();
        control.list_fill_range =
            props.attr("fmlaRange").map(str::to_owned).or(optional_text(client.child(EXCEL, "FmlaRange")?)?).unwrap_or_default();
        control.group_link = props.attr("fmlaGroup").map(str::to_owned).or(optional_text(client.child(EXCEL, "FmlaGroup")?)?).unwrap_or_default();
        control.multi_select = match props.attr("seltype").unwrap_or("single") {
            "single" => false,
            "multi" | "extended" => true,
            _ => return Err(bad("invalid native selection type")),
        };
        control.extended_select = props.attr("seltype") == Some("extended");
        let selected = props.attr("multiSel").map(str::to_owned).or(optional_text(client.child(EXCEL, "MultiSel")?)?).unwrap_or_default();
        control.selected = selected
            .split(',')
            .filter(|s| !s.trim().is_empty())
            .map(|s| s.trim().parse::<u32>().map_err(|_| bad("multiselection index")))
            .collect::<Result<Vec<_>, _>>()?;
        control.value = if matches!(kind, FormControlKind::CheckBox | FormControlKind::OptionButton) {
            match props.attr("checked") {
                Some("Checked") => 1,
                Some("Mixed") => 2,
                Some("Unchecked") => -4146,
                None => match number("", "Checked", 0)? {
                    0 => -4146,
                    1 => 1,
                    2 => 2,
                    _ => return Err(bad("check state")),
                },
                _ => return Err(bad("checked state")),
            }
        } else if kind.is_list() {
            number("sel", "Sel", 0)?
        } else {
            number("val", "Val", 0)?
        };
        if let Some(items) = props.child(X14, "itemLst")? {
            for item in items.children() {
                if item.name == "extLst" && [X14, SHEET].contains(&item.ns.as_str()) {
                    continue;
                }
                if item.ns != X14 || item.name != "item" {
                    return Err(bad("list item type"));
                }
                control.items.push(item.attr("val").ok_or_else(|| bad("list item val"))?.into());
            }
        }
        let style = v.attr("style").unwrap_or("");
        control.visible = !style.split(';').any(|s| s.split_once(':').is_some_and(|(k, v)| k.trim() == "visibility" && v.trim() == "hidden"))
            && legacy_flag(client, "Visible", true)?;
        control.enabled = !legacy_flag(client, "Disabled", false)?;
        control.print_object = legacy_flag(client, "PrintObject", true)?;
        control.validate().map_err(bad)?;
        let mut anchor = anchor(sheet, client)?;
        if let Some(pr) = node.child(SHEET, "controlPr")?
            && let Some(native) = pr.child(SHEET, "anchor")?
        {
            let modern = worksheet_anchor(sheet, native)?;
            if !same_geometry(sheet, &anchor, &modern) {
                return Err(bad("conflicting worksheet/VML geometry"));
            }
            anchor = modern;
        }
        let legacy = optional_text(client.child(EXCEL, "FmlaMacro")?)?;
        let current = node.child(SHEET, "controlPr")?.and_then(|p| p.attr("macro")).filter(|s| !s.trim().is_empty()).map(str::to_owned);
        if legacy.as_ref().zip(current.as_ref()).is_some_and(|(a, b)| a != b) {
            return Err(bad("conflicting native macro assignments"));
        }
        let drawing = sheet.shapes.iter().filter(|s| s.format.office_id == Some(id)).collect::<Vec<_>>();
        if drawing.len() > 1 {
            return Err(bad("ambiguous DrawingML identity"));
        }
        let macro_assignment = current.or(legacy);
        if let Some(d) = drawing.first() {
            if d.format.office_name.as_ref().is_some_and(|n| n != &control.name) {
                return Err(bad("conflicting DrawingML control name"));
            }
            if d.on_action.as_ref().zip(macro_assignment.as_ref()).is_some_and(|(a, b)| a != b) {
                return Err(bad("conflicting DrawingML macro assignment"));
            }
        }
        let text = v.child(VML, "textbox")?.map(Node::text).transpose()?.unwrap_or_default();
        if let Some(d) = drawing.first() {
            if !same_geometry(sheet, &anchor, &d.anchor) {
                return Err(bad("conflicting DrawingML control geometry"));
            }
            if !d.text.is_empty() && d.text != text {
                return Err(bad("conflicting DrawingML control caption"));
            }
        }
        result.push(Shape {
            identity: Default::default(),
            math: Default::default(),
            id: 0,
            anchor,
            kind: ShapeKind::Rectangle,
            fill: Color::rgb(240, 240, 240),
            line: Color::rgb(100, 100, 100),
            text,
            on_action: macro_assignment.or_else(|| drawing.first().and_then(|s| s.on_action.clone())),
            format: ShapeFormat {
                office_id: Some(id),
                office_name: Some(control.name.clone()),
                control_attachment: Some(attachment(client)?),
                form_control: Some(control),
                ..Default::default()
            },
        });
    }
    Ok(result)
}
fn anchor(sheet: &Sheet, client: &Node) -> Result<Anchor, IoError> {
    let text = client.child(EXCEL, "Anchor")?.ok_or_else(|| bad("native anchor missing"))?.text()?;
    let values = text.split(',').map(|s| s.trim().parse::<u32>().map_err(|_| bad("native anchor component"))).collect::<Result<Vec<_>, _>>()?;
    let [col, dx, row, dy, end_col, end_dx, end_row, end_dy]: [u32; 8] = values.try_into().map_err(|_| bad("native anchor arity"))?;
    if col >= MAX_COLS || end_col >= MAX_COLS || row >= MAX_ROWS || end_row >= MAX_ROWS || dx > 1023 || end_dx > 1023 || dy > 255 || end_dy > 255 {
        return Err(bad("native anchor bounds"));
    }
    let dx = sheet.col_width(col) * dx as f32 / 1024.;
    let dy = sheet.row_height(row) * dy as f32 / 256.;
    let x = sheet.col_left(col) + f64::from(dx);
    let y = sheet.row_top(row) + f64::from(dy);
    let width = (sheet.col_left(end_col) + f64::from(sheet.col_width(end_col)) * f64::from(end_dx) / 1024. - x) as f32;
    let height = (sheet.row_top(end_row) + f64::from(sheet.row_height(end_row)) * f64::from(end_dy) / 256. - y) as f32;
    if width < 0. || height < 0. || !width.is_finite() || !height.is_finite() {
        return Err(bad("negative native anchor"));
    }
    let mode = if legacy_flag(client, "SizeWithCells", false)? {
        AnchorMode::MoveAndSize
    } else if legacy_flag(client, "MoveWithCells", false)? {
        AnchorMode::MoveOnly
    } else {
        AnchorMode::Absolute
    };
    Ok(Anchor { cell: CellRef::new(row, col), dx, dy, width, height, mode })
}

fn legacy_flag(client: &Node, name: &str, default: bool) -> Result<bool, IoError> {
    let Some(n) = client.child(EXCEL, name)? else { return Ok(default) };
    match n.text()?.trim() {
        "" | "True" | "true" | "1" | "t" => Ok(true),
        "False" | "false" | "0" | "f" => Ok(false),
        _ => Err(bad(format!("invalid VML {name}"))),
    }
}
fn same_geometry(sheet: &Sheet, a: &Anchor, b: &Anchor) -> bool {
    let values = |a: &Anchor| {
        [sheet.col_left(a.cell.col) + f64::from(a.dx), sheet.row_top(a.cell.row) + f64::from(a.dy), f64::from(a.width), f64::from(a.height)]
    };
    let av = values(a);
    let bv = values(b);
    let end = |v: [f64; 4]| CellRef::new(sheet.row_at(v[1] + v[3]), sheet.col_at(v[0] + v[2]));
    let cells = [a.cell, b.cell, end(av), end(bv)];
    let xt = 1. + 2. * f64::from(cells.iter().map(|c| sheet.col_width(c.col)).fold(0., f32::max)) / 1024.;
    let yt = 1. + 2. * f64::from(cells.iter().map(|c| sheet.row_height(c.row)).fold(0., f32::max)) / 256.;
    av.into_iter().zip(bv).zip([xt, yt, xt, yt]).all(|((a, b), t)| (a - b).abs() <= t)
}

fn worksheet_anchor(sheet: &Sheet, node: &Node) -> Result<Anchor, IoError> {
    const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    let point = |tag: &str| -> Result<(CellRef, f32, f32), IoError> {
        let p = node.child(SHEET, tag)?.ok_or_else(|| bad("control anchor endpoint"))?;
        let component = |name: &str| -> Result<u32, IoError> {
            p.child(XDR, name)?.ok_or_else(|| bad("anchor coordinate"))?.text()?.parse::<u32>().map_err(|_| bad("anchor integer"))
        };
        let cell = CellRef::new(component("row")?, component("col")?);
        if !cell.is_valid() {
            return Err(bad("control anchor bounds"));
        }
        let dx = component("colOff")? as f32 / 9525.;
        let dy = component("rowOff")? as f32 / 9525.;
        Ok((cell, dx, dy))
    };
    let (cell, dx, dy) = point("from")?;
    let (end, ex, ey) = point("to")?;
    let width = (sheet.col_left(end.col) + f64::from(ex) - sheet.col_left(cell.col) - f64::from(dx)) as f32;
    let height = (sheet.row_top(end.row) + f64::from(ey) - sheet.row_top(cell.row) - f64::from(dy)) as f32;
    if width < 0. || height < 0. {
        return Err(bad("negative control anchor"));
    }
    let flag = |name: &str| -> Result<bool, IoError> {
        match node.attr(name) {
            None | Some("0" | "false") => Ok(false),
            Some("1" | "true") => Ok(true),
            _ => Err(bad("anchor movement flag")),
        }
    };
    let mode = if flag("sizeWithCells")? {
        AnchorMode::MoveAndSize
    } else if flag("moveWithCells")? {
        AnchorMode::MoveOnly
    } else {
        AnchorMode::Absolute
    };
    Ok(Anchor { cell, dx, dy, width, height, mode })
}

fn attachment(client: &Node) -> Result<gridcraft_model::form_controls::ControlAttachment, IoError> {
    let text = client.child(EXCEL, "Anchor")?.ok_or_else(|| bad("control anchor"))?.text()?;
    let values = text.split(',').map(|s| s.trim().parse::<u32>().map_err(|_| bad("anchor component"))).collect::<Result<Vec<_>, _>>()?;
    let [col, dx, row, dy, end_col, end_dx, end_row, end_dy]: [u32; 8] = values.try_into().map_err(|_| bad("anchor arity"))?;
    Ok(gridcraft_model::form_controls::ControlAttachment {
        from: CellRef::new(row, col),
        to: CellRef::new(end_row, end_col),
        fractions: [dx as f32 / 1024., dy as f32 / 256., end_dx as f32 / 1024., end_dy as f32 / 256.],
    })
}

pub(crate) fn attach(cx: &mut crate::read::Ctx<'_>, sheet: &mut Sheet) {
    let name = sheet.name.clone();
    // Identify retained native shapes before joining. Unknown classes and failed joins
    // must never fall through to the ordinary Shape.OnAction execution route.
    let declared = read_worksheet_controls(sheet);
    for shape in &mut sheet.shapes {
        if shape.format.office_id.is_some_and(|id| match &declared {
            Ok(m) => m.buttons.iter().any(|c| c.shape_id == id) || m.unsupported.iter().any(|c| c.shape_id == id),
            Err(_) => true,
        }) {
            shape.format.control_inert = true;
        }
    }
    match decode(sheet) {
        Ok(controls) => {
            if let Ok(metadata) = &declared
                && metadata.buttons.len() + metadata.unsupported.len() > controls.len()
            {
                cx.warn(format!("sheet \"{name}\": unsupported native controls were retained without activation"));
            }
            for control in controls {
                if let Some(existing) = sheet.shapes.iter_mut().find(|s| s.format.office_id == control.format.office_id) {
                    existing.on_action = control.on_action;
                    existing.text = control.text;
                    existing.format.form_control = control.format.form_control;
                    existing.format.control_attachment = control.format.control_attachment;
                    existing.format.control_inert = false;
                    existing.anchor = control.anchor;
                } else {
                    sheet.shapes.push(gridcraft_model::Shape { id: cx.next_object_id(), ..control });
                }
            }
        }
        Err(e) => cx.warn(format!("worksheet controls were retained inertly: {e}")),
    }
}
