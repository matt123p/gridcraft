use super::{
    BTreeSet, Budget, CTRL_CT, CTRL_REL, EXCEL, IoError, MAX_CONTROLS, Node, OFFICE, PreservedPart, PreservedRel, REL, SHEET, Sheet, VML, X14, bad,
    decode_native_controls, limited, normalize, patch, shape_map, xml,
};
use crate::xml::{esc, esc_attr};
use gridcraft_model::{FormControl, FormControlKind, Preserved, PreservedXml, Shape, Workbook};
use std::sync::Arc;

fn bool_text(v: bool) -> String {
    if v { "True" } else { "False" }.into()
}
fn checked(v: i32) -> &'static str {
    match v {
        1 => "Checked",
        2 => "Mixed",
        _ => "Unchecked",
    }
}
fn attrs(c: &FormControl) -> Vec<(&'static str, String)> {
    vec![
        ("objectType", c.kind.xml().into()),
        ("checked", checked(c.value).into()),
        ("min", c.min.to_string()),
        ("max", c.max.to_string()),
        ("inc", c.small_change.to_string()),
        ("page", c.large_change.to_string()),
        ("val", c.value.max(0).to_string()),
        ("sel", c.value.max(0).to_string()),
        ("fmlaLink", c.linked_cell.clone()),
        ("fmlaRange", c.list_fill_range.clone()),
        ("fmlaGroup", c.group_link.clone()),
        ("horiz", c.horizontal.to_string()),
        ("firstButton", c.first_button.to_string()),
        ("dropLines", c.drop_lines.to_string()),
        (
            "seltype",
            if c.extended_select {
                "extended"
            } else if c.multi_select {
                "multi"
            } else {
                "single"
            }
            .into(),
        ),
        ("multiSel", c.selected.iter().map(u32::to_string).collect::<Vec<_>>().join(",")),
    ]
}
fn applicable(c: &FormControl, key: &str) -> bool {
    match key {
        "checked" => matches!(c.kind, FormControlKind::CheckBox | FormControlKind::OptionButton),
        "min" | "max" | "inc" | "page" => c.kind.is_numeric(),
        "horiz" => c.kind == FormControlKind::ScrollBar,
        "sel" | "fmlaRange" => c.kind.is_list(),
        "dropLines" => c.kind == FormControlKind::DropDown,
        "seltype" => c.kind == FormControlKind::ListBox,
        "multiSel" => c.kind == FormControlKind::ListBox && c.multi_select && !c.extended_select,
        "val" => c.kind.is_numeric(),
        "firstButton" => c.kind == FormControlKind::OptionButton,
        "fmlaGroup" => c.kind == FormControlKind::GroupBox,
        "fmlaLink" => !matches!(c.kind, FormControlKind::Button | FormControlKind::Label | FormControlKind::GroupBox),
        _ => true,
    }
}
fn items(c: &FormControl) -> String {
    format!("<itemLst xmlns=\"{X14}\">{}</itemLst>", c.items.iter().map(|s| format!("<item val=\"{}\"/>", esc_attr(s))).collect::<String>())
}
fn anchor(sheet: &Sheet, shape: &Shape) -> String {
    if shape.anchor.mode == gridcraft_model::AnchorMode::MoveAndSize
        && let Some(a) = &shape.format.control_attachment
    {
        let [dx, dy, ex, ey] = a.fractions;
        return format!(
            "{}, {}, {}, {}, {}, {}, {}, {}",
            a.from.col,
            (dx * 1024.).round().clamp(0., 1023.) as u32,
            a.from.row,
            (dy * 256.).round().clamp(0., 255.) as u32,
            a.to.col,
            (ex * 1024.).round().clamp(0., 1023.) as u32,
            a.to.row,
            (ey * 256.).round().clamp(0., 255.) as u32
        );
    }
    let a = shape.anchor;
    let x = sheet.col_left(a.cell.col) + f64::from(a.dx + a.width);
    let y = sheet.row_top(a.cell.row) + f64::from(a.dy + a.height);
    let col = sheet.col_at(x);
    let row = sheet.row_at(y);
    let fraction = |offset: f64, size: f32, units: f64| {
        if size > 0. { (offset / f64::from(size) * units).round().clamp(0., units - 1.) as u32 } else { 0 }
    };
    format!(
        "{}, {}, {}, {}, {}, {}, {}, {}",
        a.cell.col,
        fraction(f64::from(a.dx), sheet.col_width(a.cell.col), 1024.),
        a.cell.row,
        fraction(f64::from(a.dy), sheet.row_height(a.cell.row), 256.),
        col,
        fraction(x - sheet.col_left(col), sheet.col_width(col), 1024.),
        row,
        fraction(y - sheet.row_top(row), sheet.row_height(row), 256.)
    )
}
fn worksheet_anchor(sheet: &Sheet, shape: &Shape) -> String {
    let a = shape.anchor;
    let x = sheet.col_left(a.cell.col) + f64::from(a.dx + a.width);
    let y = sheet.row_top(a.cell.row) + f64::from(a.dy + a.height);
    let col = sheet.col_at(x);
    let row = sheet.row_at(y);
    let (cell, dx, dy, col, row, ex, ey) = if shape.anchor.mode == gridcraft_model::AnchorMode::MoveAndSize
        && let Some(attachment) = &shape.format.control_attachment
    {
        let [dx, dy, ex, ey] = attachment.fractions;
        (
            attachment.from,
            dx * sheet.col_width(attachment.from.col),
            dy * sheet.row_height(attachment.from.row),
            attachment.to.col,
            attachment.to.row,
            ex * sheet.col_width(attachment.to.col),
            ey * sheet.row_height(attachment.to.row),
        )
    } else {
        (a.cell, a.dx, a.dy, col, row, (x - sheet.col_left(col)) as f32, (y - sheet.row_top(row)) as f32)
    };
    format!(
        "<anchor xmlns=\"{SHEET}\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" moveWithCells=\"{}\" sizeWithCells=\"{}\"><from><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></from><to><xdr:col>{col}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{}</xdr:rowOff></to></anchor>",
        shape.anchor.mode != gridcraft_model::AnchorMode::Absolute,
        shape.anchor.mode == gridcraft_model::AnchorMode::MoveAndSize,
        cell.col,
        (f64::from(dx) * 9525.).round() as i64,
        cell.row,
        (f64::from(dy) * 9525.).round() as i64,
        (f64::from(ex) * 9525.).round() as i64,
        (f64::from(ey) * 9525.).round() as i64
    )
}
fn client_values(sheet: &Sheet, shape: &Shape, c: &FormControl) -> Vec<(&'static str, String)> {
    vec![
        ("Anchor", anchor(sheet, shape)),
        ("FmlaMacro", shape.on_action.clone().unwrap_or_default()),
        ("PrintObject", bool_text(c.print_object)),
        ("Visible", bool_text(c.visible)),
        ("FmlaLink", c.linked_cell.clone()),
        ("FmlaRange", c.list_fill_range.clone()),
        ("FmlaGroup", c.group_link.clone()),
        ("MultiSel", c.selected.iter().map(u32::to_string).collect::<Vec<_>>().join(",")),
        (
            "SelType",
            if c.extended_select {
                "Extend"
            } else if c.multi_select {
                "Multi"
            } else {
                "Single"
            }
            .into(),
        ),
        (
            "Checked",
            match c.value {
                1 => "1",
                2 => "2",
                _ => "0",
            }
            .into(),
        ),
        ("Val", c.value.max(0).to_string()),
        ("Sel", c.value.max(0).to_string()),
        ("Min", c.min.to_string()),
        ("Max", c.max.to_string()),
        ("Inc", c.small_change.to_string()),
        ("Page", c.large_change.to_string()),
        ("DropLines", c.drop_lines.to_string()),
    ]
    .into_iter()
    .filter(|(k, _)| {
        matches!(*k, "Anchor" | "FmlaMacro" | "Visible" | "PrintObject")
            || applicable(
                c,
                match *k {
                    "Checked" => "checked",
                    "FmlaLink" => "fmlaLink",
                    "FmlaRange" => "fmlaRange",
                    "FmlaGroup" => "fmlaGroup",
                    "Val" => "val",
                    "Sel" => "sel",
                    "Min" => "min",
                    "Max" => "max",
                    "Inc" => "inc",
                    "Page" => "page",
                    "DropLines" => "dropLines",
                    "MultiSel" | "SelType" => "seltype",
                    _ => "",
                },
            )
    })
    .collect()
}
fn new_vml(sheet: &Sheet, shape: &Shape, c: &FormControl, id: u32) -> String {
    let children = client_values(sheet, shape, c)
        .into_iter()
        .filter(|(k, _)| {
            matches!(*k, "Anchor" | "FmlaMacro" | "Visible" | "PrintObject")
                || applicable(
                    c,
                    match *k {
                        "Checked" => "checked",
                        "FmlaLink" => "fmlaLink",
                        "FmlaRange" => "fmlaRange",
                        "FmlaGroup" => "fmlaGroup",
                        "Val" => "val",
                        "Sel" => "sel",
                        "Min" => "min",
                        "Max" => "max",
                        "Inc" => "inc",
                        "Page" => "page",
                        "DropLines" => "dropLines",
                        "MultiSel" | "SelType" => "seltype",
                        _ => "",
                    },
                )
        })
        .map(|(k, v)| format!("<x:{k}>{}</x:{k}>", esc(&v)))
        .collect::<String>();
    format!(
        "<v:shape xmlns:v=\"{VML}\" xmlns:o=\"{OFFICE}\" xmlns:x=\"{EXCEL}\" id=\"_x0000_s{id}\" o:spid=\"_x0000_s{id}\" style=\"position:absolute;visibility:{}\"><v:textbox><div>{}</div></v:textbox><x:ClientData ObjectType=\"{}\">{children}{}{}{}</x:ClientData></v:shape>",
        if c.visible { "visible" } else { "hidden" },
        esc(&shape.text),
        c.kind.xml(),
        if c.enabled { "" } else { "<x:Disabled/>" },
        if shape.anchor.mode != gridcraft_model::AnchorMode::Absolute { "<x:MoveWithCells/>" } else { "" },
        if shape.anchor.mode == gridcraft_model::AnchorMode::MoveAndSize { "<x:SizeWithCells/>" } else { "" }
    )
}
fn walk<'a>(node: &'a Node, ns: &str, name: &str, out: &mut Vec<&'a Node>) {
    if node.ns == ns && node.name == name {
        out.push(node);
    }
    for n in node.children() {
        walk(n, ns, name, out);
    }
}
pub(crate) fn prepare(wb: &Workbook) -> Result<Workbook, IoError> {
    let mut out = wb.clone();
    for si in 0..wb.sheets.len() {
        let sheet = wb.sheet(si).ok_or_else(|| bad("sheet"))?;
        if !sheet.shapes.iter().any(|s| s.format.form_control.is_some())
            && sheet.preserved.as_ref().is_none_or(|p| !p.elements.iter().any(|e| e.slot == "controls"))
        {
            continue;
        }
        let mut original_sheet = sheet.clone();
        original_sheet.shapes.clear();
        let mut original = match decode_native_controls(&original_sheet) {
            Ok(v) => v,
            Err(_) if !sheet.shapes.iter().any(|s| s.format.form_control.is_some()) => continue,
            Err(e) => return Err(e),
        };
        for old in &mut original {
            if old.on_action.is_none() {
                old.on_action = sheet.shapes.iter().find(|s| s.format.office_id == old.format.office_id).and_then(|s| s.format.source_macro.clone());
            }
        }
        let live = sheet.shapes.iter().filter(|s| s.format.form_control.is_some()).collect::<Vec<_>>();
        if live.len() > MAX_CONTROLS {
            return Err(limited("worksheet controls"));
        }
        let mut total = 0usize;
        let mut ids = BTreeSet::new();
        let mut names = BTreeSet::new();
        for shape in &live {
            let c = shape.format.form_control.as_ref().ok_or_else(|| bad("control"))?;
            c.validate().map_err(bad)?;
            if !shape.anchor.cell.is_valid()
                || [shape.anchor.dx, shape.anchor.dy, shape.anchor.width, shape.anchor.height]
                    .iter()
                    .any(|n| !n.is_finite() || *n < 0. || *n > 1_000_000.)
                || shape.text.len() > 16384
                || !gridcraft_model::form_controls::valid_control_text(&shape.text)
            {
                return Err(bad("native control geometry/caption bounds"));
            }
            let id = shape.format.office_id.ok_or_else(|| bad("native control shape ID"))?;
            if id == 0 || !ids.insert(id) || !names.insert(c.name.to_ascii_lowercase()) {
                return Err(bad("duplicate native control identity"));
            }
            total = total
                .saturating_add(shape.text.len() + c.linked_cell.len() + c.list_fill_range.len() + c.items.iter().map(String::len).sum::<usize>());
        }
        if total > 4 * 1024 * 1024 {
            return Err(limited("native control aggregate text"));
        }
        let mut preserved = sheet.preserved.as_deref().cloned().unwrap_or_default();
        update(sheet, &live, &original, &mut preserved)?;
        let target = out.sheet_mut(si).ok_or_else(|| bad("sheet"))?;
        target.preserved = Some(Arc::new(preserved));
    }
    Ok(out)
}
fn update(sheet: &Sheet, live: &[&Shape], original: &[Shape], p: &mut Preserved) -> Result<(), IoError> {
    let old_controls = p
        .elements
        .iter()
        .find(|e| e.slot == "controls")
        .map(|e| e.xml.clone())
        .unwrap_or_else(|| format!("<controls xmlns=\"{SHEET}\" xmlns:r=\"{REL}\"></controls>"));
    let old_vml = p
        .vml
        .as_ref()
        .map(|p| String::from_utf8(p.data.to_vec()).map_err(|_| bad("VML encoding")))
        .transpose()?
        .unwrap_or_else(|| format!("<xml xmlns:v=\"{VML}\" xmlns:x=\"{EXCEL}\"></xml>"));
    let mut budget = Budget::default();
    let controls = xml::parse(old_controls.as_bytes(), &mut budget)?;
    let vml = xml::parse(old_vml.as_bytes(), &mut budget)?;
    let mut nodes = vec![];
    walk(&controls, SHEET, "control", &mut nodes);
    let mut cp = vec![];
    let mut vp = vec![];
    let mut append_controls = String::new();
    let mut append_vml = String::new();
    let vm = shape_map(&vml)?;
    let mut removed_rels = BTreeSet::new();
    for old in original {
        if !live.iter().any(|s| s.format.office_id == old.format.office_id) {
            let id = old.format.office_id.ok_or_else(|| bad("old control ID"))?;
            for n in nodes.iter().filter(|n| n.attr("shapeId").and_then(|s| s.parse::<u32>().ok()) == Some(id)) {
                cp.push((n.start, n.end, String::new()));
                if let Some(rid) = n.attr_ns(REL, "id") {
                    removed_rels.insert(rid.to_owned());
                }
            }
            if let Some(n) = vm.get(&id) {
                vp.push((n.start, n.end, String::new()));
            }
        }
    }
    let removed_parts = p.rels.iter().filter(|r| removed_rels.contains(&r.id)).map(|r| normalize(&r.target)).collect::<BTreeSet<_>>();
    p.rels.retain(|r| !removed_rels.contains(&r.id));
    p.parts.retain(|part| !removed_parts.contains(&normalize(&part.name)));
    for shape in live {
        let id = shape.format.office_id.ok_or_else(|| bad("native ID"))?;
        let c = shape.format.form_control.as_ref().ok_or_else(|| bad("control"))?;
        let old = original.iter().find(|s| s.format.office_id == Some(id));
        if old.is_some_and(|old| {
            old.anchor == shape.anchor
                && old.on_action == shape.on_action
                && old.text == shape.text
                && old.format.form_control == shape.format.form_control
        }) {
            continue;
        }
        let references = nodes.iter().filter(|n| n.attr("shapeId").and_then(|s| s.parse::<u32>().ok()) == Some(id)).copied().collect::<Vec<_>>();
        if references.is_empty() {
            let rid = format!("rIdFormControl{id}");
            let name = format!("xl/ctrlProps/control{id}.xml");
            if p.rels.iter().any(|r| r.id == rid) || p.parts.iter().any(|p| p.name == name) {
                return Err(bad("new control package identity collision"));
            }
            let a = attrs(c).into_iter().filter(|(k, _)| applicable(c, k)).map(|(k, v)| format!(" {k}=\"{}\"", esc_attr(&v))).collect::<String>();
            p.parts.push(PreservedPart {
                name: name.clone(),
                content_type: CTRL_CT.into(),
                data: format!("<formControlPr xmlns=\"{X14}\"{a}>{}</formControlPr>", if c.kind.is_list() { items(c) } else { String::new() })
                    .into_bytes()
                    .into(),
                rels: vec![],
            });
            p.rels.push(PreservedRel { id: rid.clone(), rel_type: CTRL_REL.into(), target: name, external: false });
            append_controls.push_str(&format!("<control xmlns=\"{SHEET}\" xmlns:r=\"{REL}\" shapeId=\"{id}\" name=\"{}\" r:id=\"{rid}\"><controlPr macro=\"{}\">{}</controlPr></control>",esc_attr(&c.name),esc_attr(shape.on_action.as_deref().unwrap_or("")),worksheet_anchor(sheet,shape)));
            append_vml.push_str(&new_vml(sheet, shape, c, id));
            continue;
        }
        for n in references {
            let rid = n.attr_ns(REL, "id").ok_or_else(|| bad("control relation"))?;
            let target = p.rels.iter().find(|r| r.id == rid && !r.external).map(|r| r.target.clone()).ok_or_else(|| bad("control relation"))?;
            let part = p.parts.iter_mut().find(|p| normalize(&p.name) == normalize(&target)).ok_or_else(|| bad("control property part"))?;
            let source = std::str::from_utf8(&part.data).map_err(|_| bad("control property encoding"))?;
            let root = xml::parse(source.as_bytes(), &mut Budget::default())?;
            let updates = attrs(c).into_iter().filter(|(k, v)| applicable(c, k) && root.attr(k) != Some(v.as_str())).collect::<Vec<_>>();
            let mut children = vec![];
            let mut append = String::new();
            if c.kind.is_list() {
                if let Some(item) = root.child(X14, "itemLst")? {
                    let fresh = c.items.iter().map(|s| format!("<item xmlns=\"{X14}\" val=\"{}\"/>", esc_attr(s))).collect::<String>();
                    let existing = item.children().filter(|n| n.ns == X14 && n.name == "item").collect::<Vec<_>>();
                    let mut edits = existing
                        .iter()
                        .enumerate()
                        .map(|(i, n)| (n.start, n.end, if i == 0 { fresh.clone() } else { String::new() }))
                        .collect::<Vec<_>>();
                    let mut append_items = String::new();
                    if existing.is_empty() {
                        if let Some(ext) = item.children().find(|n| n.name == "extLst") {
                            edits.push((ext.start, ext.start, fresh));
                        } else {
                            append_items = fresh;
                        }
                    }
                    let updated = patch::element(source, item, &[], edits, &append_items)?;
                    children.push((item.start, item.end, updated));
                } else {
                    append = items(c);
                }
            }
            part.data = patch::element_removing(
                source,
                &root,
                &updates,
                children,
                &append,
                if c.multi_select && !c.extended_select { &[] } else { &["multiSel"] },
            )?
            .into_bytes()
            .into();
            let anchor = worksheet_anchor(sheet, shape);
            let pr = if let Some(pr) = n.child(SHEET, "controlPr")? {
                let child = pr.child(SHEET, "anchor")?;
                patch::element(
                    &old_controls,
                    pr,
                    &[("macro", shape.on_action.clone().unwrap_or_default())],
                    child.map(|a| vec![(a.start, a.end, anchor.clone())]).unwrap_or_default(),
                    if child.is_none() { &anchor } else { "" },
                )?
            } else {
                format!("<controlPr xmlns=\"{SHEET}\" macro=\"{}\">{anchor}</controlPr>", esc_attr(shape.on_action.as_deref().unwrap_or("")))
            };
            let children = n.child(SHEET, "controlPr")?.map(|child| vec![(child.start, child.end, pr.clone())]).unwrap_or_default();
            let text = patch::element(
                &old_controls,
                n,
                &[("name", c.name.clone())],
                children,
                if n.child(SHEET, "controlPr")?.is_none() { &pr } else { "" },
            )?;
            cp.push((n.start, n.end, text));
        }
        let v = vm.get(&id).ok_or_else(|| bad("control VML shape"))?;
        let client = v.child(EXCEL, "ClientData")?.ok_or_else(|| bad("ClientData"))?;
        let values = client_values(sheet, shape, c);
        let (mut changes, mut append) = patch::child_changes(&old_vml, client, EXCEL, &values)?;
        let q = patch::qname(&old_vml, client)?;
        let prefix = q.rsplit_once(':').map(|(p, _)| format!("{p}:")).unwrap_or_default();
        for (name, enabled) in [
            ("Disabled", !c.enabled),
            ("Horiz", c.horizontal),
            ("FirstButton", c.first_button),
            ("MoveWithCells", shape.anchor.mode != gridcraft_model::AnchorMode::Absolute),
            ("SizeWithCells", shape.anchor.mode == gridcraft_model::AnchorMode::MoveAndSize),
        ] {
            if let Some(n) = client.child(EXCEL, name)? {
                if !enabled {
                    changes.push((n.start, n.end, String::new()));
                }
            } else if enabled {
                append.push_str(&format!("<{prefix}{name}/>"));
            }
        }
        let client_text = patch::element(&old_vml, client, &[], changes, &append)?;
        let mut children = vec![(client.start, client.end, client_text)];
        let mut append = String::new();
        let textbox = format!("<v:textbox xmlns:v=\"{VML}\"><div>{}</div></v:textbox>", esc(&shape.text));
        if let Some(tb) = v.child(VML, "textbox")? {
            if old.is_none_or(|o| o.text != shape.text) {
                children.push((tb.start, tb.end, textbox));
            }
        } else {
            append = textbox;
        }
        let old_style = v.attr("style").unwrap_or("");
        let mut style = old_style.split(';').filter(|s| !s.trim().starts_with("visibility:")).collect::<Vec<_>>().join(";");
        style.push_str(if c.visible { ";visibility:visible" } else { ";visibility:hidden" });
        vp.push((v.start, v.end, patch::element(&old_vml, v, &[("style", style)], children, &append)?));
    }
    let mut controls = patch::apply(&old_controls, cp)?;
    if !append_controls.is_empty() {
        let root = xml::parse(controls.as_bytes(), &mut Budget::default())?;
        let mut containers = vec![];
        walk(&root, SHEET, "controls", &mut containers);
        let edits = containers
            .iter()
            .map(|n| patch::element(&controls, n, &[], vec![], &append_controls).map(|s| (n.start, n.end, s)))
            .collect::<Result<Vec<_>, _>>()?;
        controls = patch::apply(&controls, edits)?;
    }
    p.elements.retain(|e| e.slot != "controls");
    let root = xml::parse(controls.as_bytes(), &mut Budget::default())?;
    let mut remaining = vec![];
    walk(&root, SHEET, "control", &mut remaining);
    if !remaining.is_empty() {
        p.elements.push(PreservedXml { slot: "controls".into(), xml: controls });
    }
    let mut vml = patch::apply(&old_vml, vp)?;
    if !append_vml.is_empty() {
        let root = xml::parse(vml.as_bytes(), &mut Budget::default())?;
        vml = patch::element(&vml, &root, &[], vec![], &append_vml)?;
    }
    if !live.is_empty() || p.vml.is_some() {
        let mut part = p.vml.clone().unwrap_or_else(|| PreservedPart {
            name: "xl/drawings/controls.vml".into(),
            content_type: "application/vnd.openxmlformats-officedocument.vmlDrawing".into(),
            ..Default::default()
        });
        part.data = vml.into_bytes().into();
        p.vml = Some(part);
    }
    Ok(())
}
