//! Shared native drawing graph. ShapeRange snapshots live handles, never a second drawing store.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::{
    core::{CellRef, RangeRef},
    model::{Anchor, AnchorMode, Color, ShapeFormat, ShapeKind, ShapeTextAlign, ShapeTextAnchor, Sheet},
};
use serde_json::{Value as Json, json};
#[derive(Clone)]
struct View {
    id: u32,
    tag: &'static str,
    name: String,
    anchor: Anchor,
    text: String,
    format: ShapeFormat,
    fill: Color,
    line: Color,
    on_action: Option<String>,
    shape_kind: Option<ShapeKind>,
    index: usize,
}
fn root(sheet: usize, id: u32, identity: &gridcraft_engine::model::sheet::SheetIdentity, tag: &str) -> Feature {
    let mut f = Feature::new("Shape", Some(sheet));
    f.key = u64::from(id);
    f.identity = Some(identity.clone());
    f.text = Some(tag.into());
    f
}
fn child(f: &Feature, kind: &'static str, part: i32) -> Feature {
    let mut n = f.clone();
    n.kind = kind;
    n.part = part;
    n
}
fn roots(sh: &Sheet, sheet: usize) -> HostResult<Vec<Feature>> {
    let count = sh.shapes.len().saturating_add(sh.images.len()).saturating_add(sh.charts.len()).saturating_add(sh.slicers.len());
    if count > super::MAX_OBJECTS {
        return Err(failure(7, "drawing collection budget"));
    }
    let mut out = Vec::new();
    out.try_reserve(count).map_err(|_| failure(7, "drawing allocation"))?;
    out.extend(sh.shapes.iter().map(|x| root(sheet, x.id, &x.identity, "shape")));
    out.extend(sh.images.iter().map(|x| root(sheet, x.id, &x.identity, "image")));
    out.extend(sh.charts.iter().map(|x| root(sheet, x.id, &x.identity, "chart")));
    out.extend(sh.slicers.iter().map(|x| root(sheet, x.id, &x.identity, "slicer")));
    Ok(out)
}
fn view(sh: &Sheet, f: &Feature) -> HostResult<View> {
    let identity = f.identity.as_ref().ok_or_else(|| failure(91, "drawing identity"))?;
    let duplicates = match f.text.as_deref().unwrap_or("shape") {
        "chart" => sh.charts.iter().filter(|x| u64::from(x.id) == f.key).count(),
        "image" => sh.images.iter().filter(|x| u64::from(x.id) == f.key).count(),
        "slicer" => sh.slicers.iter().filter(|x| u64::from(x.id) == f.key).count(),
        _ => sh.shapes.iter().filter(|x| u64::from(x.id) == f.key).count(),
    };
    if duplicates > 1 {
        return Err(failure(1004, "native drawing ID is ambiguous"));
    }

    let mut v = View {
        id: 0,
        tag: "shape",
        name: String::new(),
        anchor: Anchor::default(),
        text: String::new(),
        format: ShapeFormat::default(),
        fill: Color::rgb(255, 255, 255),
        line: Color::rgb(0, 0, 0),
        on_action: None,
        shape_kind: None,
        index: 0,
    };
    match f.text.as_deref().unwrap_or("shape") {
        "shape" => {
            let (i, x) = sh
                .shapes
                .iter()
                .enumerate()
                .find(|(_, x)| u64::from(x.id) == f.key && identity.same(&x.identity))
                .ok_or_else(|| failure(91, "shape deleted or replaced"))?;
            if x.text.len() > 131072 || x.format.font.as_ref().is_some_and(|s| s.len() > 256) {
                return Err(failure(7, "shape metadata budget"));
            }
            v.id = x.id;
            v.index = i;
            v.name = format!("{:?} {}", x.kind, x.id);
            v.anchor = x.anchor;
            v.text = x.text.clone();
            v.format = x.format.clone();
            v.fill = x.fill;
            v.line = x.line;
            v.on_action = x.on_action.clone();
            v.shape_kind = Some(x.kind);
        }
        "slicer" => {
            let (i, x) = sh
                .slicers
                .iter()
                .enumerate()
                .find(|(_, x)| u64::from(x.id) == f.key && identity.same(&x.identity))
                .ok_or_else(|| failure(91, "slicer deleted or replaced"))?;
            v.id = x.id;
            v.index = i;
            v.tag = "slicer";
            v.name = format!("Slicer{}", x.id);
            v.anchor = x.anchor;
        }
        "image" => {
            let (i, x) = sh
                .images
                .iter()
                .enumerate()
                .find(|(_, x)| u64::from(x.id) == f.key && identity.same(&x.identity))
                .ok_or_else(|| failure(91, "picture deleted or replaced"))?;
            v.id = x.id;
            v.index = i;
            v.tag = "image";
            v.name = format!("Picture {}", x.id);
            v.anchor = x.anchor;
            v.text = x.alt.clone();
        }
        "chart" => {
            let (i, x) = sh
                .charts
                .iter()
                .enumerate()
                .find(|(_, x)| u64::from(x.id) == f.key && identity.same(&x.identity))
                .ok_or_else(|| failure(91, "chart deleted or replaced"))?;
            v.id = x.id;
            v.index = i;
            v.tag = "chart";
            v.name = format!("Chart {}", x.id);
            v.anchor = x.anchor;
        }
        _ => return Err(failure(91, "invalid drawing owner")),
    }
    Ok(v)
}
impl Bridge {
    pub(crate) fn drawing_api(&mut self, s: &mut Session, o: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if let Object::Sheet(sheet) = o
            && call.member.eq_ignore_ascii_case("Shapes")
            && reads(call)
        {
            let f = Feature::new("Shapes", Some(*sheet));
            return Some(if call.arguments.is_empty() { self.feature(f) } else { self.drawing_dispatch(s, &f, call) });
        }
        let Object::Feature(f) = o else {
            return None;
        };
        let m = call.member.to_ascii_lowercase();
        if f.kind == "ChartObjects" && m == "shaperange" && reads(call) {
            return Some((|| {
                arguments(call, &[])?;
                let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
                let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?;
                if sh.charts.is_empty() || sh.charts.len() > 256 {
                    return Err(failure(7, "ChartObjects.ShapeRange requires 1..256 charts"));
                }
                let mut n = Feature::new("ShapeRange", Some(sheet));
                n.items = sh.charts.iter().map(|c| root(sheet, c.id, &c.identity, "chart")).collect();
                self.feature(n)
            })());
        }
        if f.kind == "ChartObject" && matches!(m.as_str(), "shaperange" | "zorder" | "bottomrightcell" | "bringtofront" | "sendtoback") {
            let mut n = f.clone();
            n.kind = "Shape";
            n.text = Some("chart".into());
            if m == "shaperange" {
                let mut range = Feature::new("ShapeRange", f.sheet);
                range.items.push(n);
                return Some(self.feature(range));
            }
            if matches!(m.as_str(), "bringtofront" | "sendtoback") {
                return Some((|| {
                    arguments(call, &[])?;
                    let sh = s
                        .doc()
                        .map_err(|e| failure(91, e.to_string()))?
                        .wb
                        .sheet(f.sheet.ok_or_else(|| failure(91, "sheet"))?)
                        .ok_or_else(|| failure(91, "sheet"))?;
                    let v = view(sh, &n)?;
                    self.zorder(s, f.sheet.ok_or_else(|| failure(91, "sheet"))?, &v, if m == "bringtofront" { 0 } else { 1 })
                })());
            }
            let mut call = call.clone();
            if m == "zorder" {
                call.member = "ZOrderPosition".into();
            }
            return Some(self.drawing_dispatch(s, &n, &call));
        }
        if matches!(f.kind, "Series" | "Trendline") && m == "format" && reads(call) {
            return Some(self.feature({
                let mut n = f.clone();
                n.kind = "SeriesFormat";
                n.text = Some(if f.kind == "Trendline" { "trendline" } else { "series" }.into());
                n
            }));
        }
        if matches!(f.kind, "FillFormat" | "LineFormat" | "ColorFormat" | "SeriesFormat") && matches!(f.text.as_deref(), Some("series" | "trendline"))
        {
            return Some(self.series_format(s, f, call));
        }
        if !matches!(
            f.kind,
            "Shapes"
                | "Shape"
                | "ShapeRange"
                | "TextFrame"
                | "Characters"
                | "TextFrame2"
                | "TextRange2"
                | "ShapeFont"
                | "ShapeFont2"
                | "FillFormat"
                | "LineFormat"
                | "ColorFormat"
        ) {
            return None;
        }
        Some(self.drawing_dispatch(s, f, call))
    }
    fn drawing_dispatch(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "drawing worksheet deleted"))?;
        let wb = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.clone();
        let sh = wb.sheet(sheet).ok_or_else(|| failure(91, "drawing worksheet deleted"))?;
        let nodes = roots(sh, sheet)?;
        if f.kind == "Shapes" {
            return self.drawing_collection(s, f, &nodes, call);
        }
        if f.kind == "ShapeRange" {
            return self.drawing_range(s, f, call);
        }
        let v = view(sh, f)?;
        let m = call.member.to_ascii_lowercase();
        if m == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if f.kind != "Shape" {
            return self.drawing_child(s, f, &v, call);
        }
        if reads(call) {
            if m == "parent" {
                arguments(call, &[])?;
                return self.intern(Object::Sheet(sheet));
            }
            if m == "shaperange" {
                arguments(call, &[])?;
                let mut n = Feature::new("ShapeRange", Some(sheet));
                n.items.push(f.clone());
                return self.feature(n);
            }
            if m == "textframe" || m == "textframe2" {
                arguments(call, &[])?;
                if v.tag != "shape" {
                    return Err(failure(1004, "drawing has no native text frame"));
                }
                return self.feature(child(f, if m == "textframe" { "TextFrame" } else { "TextFrame2" }, 0));
            }
            if matches!(m.as_str(), "fill" | "line") {
                arguments(call, &[])?;
                if v.tag != "shape" {
                    return Err(failure(1004, "picture/chart fill is not an independent native shape field"));
                }
                return self.feature(child(f, if m == "fill" { "FillFormat" } else { "LineFormat" }, if m == "fill" { 1 } else { 2 }));
            }
            if m == "chart" {
                arguments(call, &[])?;
                if v.tag != "chart" {
                    return Err(failure(1004, "shape is not a chart"));
                }
                let mut n = f.clone();
                n.kind = "Chart";
                n.text = None;
                return self.feature(n);
            }
            if m == "delete" {
                arguments(call, &[])?;
                return if v.tag == "slicer" {
                    self.drawing_command(s, sheet, "slicer.delete", json!({"slicer":v.id})).map(|_| Value::Empty)
                } else {
                    self.drawing_command(s, sheet, "object.delete", json!({"kind":v.tag,"id":v.id})).map(|_| Value::Empty)
                };
            }
            if m == "zorder" {
                let a = arguments(call, &["ZOrderCmd"])?;
                return self.zorder(s, sheet, &v, numeric_index(required(&a, 0)?)?);
            }
            if m == "incrementleft" || m == "incrementtop" {
                let a = arguments(call, &["Increment"])?;
                let n = super::format_api::number(required(&a, 0)?)?;
                let pos = if m == "incrementleft" {
                    sh.col_left(v.anchor.cell.col) + f64::from(v.anchor.dx)
                } else {
                    sh.row_top(v.anchor.cell.row) + f64::from(v.anchor.dy)
                };
                return self.set_geometry(s, sheet, &v, if m == "incrementleft" { "left" } else { "top" }, pos + n);
            }
            if m == "pickupe" || matches!(m.as_str(), "group" | "ungroup" | "copy" | "cut" | "copypicture" | "duplicate") {
                return Err(failure(1004, "native drawings have no persistent group or object clipboard service"));
            }
            arguments(call, &[])?;
            return match m.as_str() {
                "name" => values::vba_text(&v.name),
                "id" => Ok(Value::Long(i32::try_from(v.id).map_err(|_| failure(6, "drawing ID overflow"))?)),
                "type" => Ok(Value::Long(match (v.tag, v.shape_kind) {
                    ("chart", _) => 3,
                    ("image", _) => 13,
                    ("slicer", _) => 25,
                    (_, Some(ShapeKind::TextBox)) => 17,
                    (_, Some(ShapeKind::Line | ShapeKind::Arrow)) => 9,
                    _ => 1,
                })),
                "autoshapetype" => Ok(Value::Long(match v.shape_kind {
                    Some(ShapeKind::Rectangle) => 1,
                    Some(ShapeKind::RoundedRectangle) => 5,
                    Some(ShapeKind::Ellipse) => 9,
                    Some(ShapeKind::Triangle) => 7,
                    _ => return Err(failure(1004, "no native AutoShape variant")),
                })),
                "width" => Ok(Value::Double(f64::from(v.anchor.width))),
                "height" => Ok(Value::Double(f64::from(v.anchor.height))),
                "left" => Ok(Value::Double(sh.col_left(v.anchor.cell.col) + f64::from(v.anchor.dx))),
                "top" => Ok(Value::Double(sh.row_top(v.anchor.cell.row) + f64::from(v.anchor.dy))),
                "placement" => Ok(Value::Long(match v.anchor.mode {
                    AnchorMode::MoveAndSize => 1,
                    AnchorMode::MoveOnly => 2,
                    AnchorMode::Absolute => 3,
                })),
                "topleftcell" => self.intern(Object::Range(sheet, RangeRef::cell(v.anchor.cell).into())),
                "bottomrightcell" => {
                    let at = CellRef::new(
                        sh.row_at(sh.row_top(v.anchor.cell.row) + f64::from(v.anchor.dy + v.anchor.height)),
                        sh.col_at(sh.col_left(v.anchor.cell.col) + f64::from(v.anchor.dx + v.anchor.width)),
                    );
                    self.intern(Object::Range(sheet, RangeRef::cell(at).into()))
                }
                "onaction" if v.tag == "shape" => values::vba_text(v.on_action.as_deref().unwrap_or("")),
                "alternativetext" if v.tag == "image" => values::vba_text(&v.text),
                "haschart" => Ok(Value::Long(if v.tag == "chart" { -1 } else { 0 })),
                "hastextframe" => Ok(Value::Long(if v.tag == "shape" { -1 } else { 0 })),
                "zorderposition" => Ok(Value::Long((v.index + 1) as i32)),
                _ => Err(failure(1004, "drawing field is outside the native profile")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "unsupported drawing invocation"));
        }
        let a = arguments(call, &["Value"])?;
        let value = required(&a, 0)?;
        match m.as_str() {
            "width" | "height" | "left" | "top" => self.set_geometry(s, sheet, &v, &m, super::format_api::number(value)?),
            "placement" => {
                let mode = match numeric_index(value)? {
                    1 => "moveAndSize",
                    2 => "moveOnly",
                    3 => "absolute",
                    _ => return Err(failure(1004, "invalid Placement")),
                };
                self.drawing_command(s, sheet, "object.setAnchorMode", json!({"kind":v.tag,"id":v.id,"mode":mode}))?;
                Ok(Value::Empty)
            }
            "onaction" if v.tag == "shape" => {
                self.drawing_command(s, sheet, "shape.setMacro", json!({"id":v.id,"entry":string(value)?}))?;
                Ok(Value::Empty)
            }
            _ => Err(failure(1004, "native drawing has no writable field for this property")),
        }
    }
    fn drawing_collection(&mut self, s: &mut Session, f: &Feature, nodes: &[Feature], call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let m = call.member.to_ascii_lowercase();
        if call.kind == InvokeKind::Enumerate {
            arguments(call, &[])?;
            return self.object_array(nodes.iter().cloned().map(Feature::object).collect());
        }
        if !reads(call) {
            return Err(failure(438, "drawing collection is read-only"));
        }
        match m.as_str() {
            "count" => {
                arguments(call, &[])?;
                Ok(Value::Long(nodes.len() as i32))
            }
            "parent" => {
                arguments(call, &[])?;
                self.intern(Object::Sheet(sheet))
            }
            "application" => {
                arguments(call, &[])?;
                self.intern(Object::Application)
            }
            "item" | "" | "shapes" => {
                let a = arguments(call, &["Index"])?;
                let n = self.drawing_index(s, nodes, required(&a, 0)?)?;
                self.feature(n)
            }
            "range" => {
                let a = arguments(call, &["Index"])?;
                let mut n = Feature::new("ShapeRange", Some(sheet));
                match required(&a, 0)? {
                    Value::Array(a) => {
                        if a.elements().is_empty() || a.elements().len() > 256 {
                            return Err(failure(7, "ShapeRange requires 1..256 items"));
                        }
                        for v in a.elements() {
                            n.items.push(self.drawing_index(s, nodes, v)?);
                        }
                    }
                    v => n.items.push(self.drawing_index(s, nodes, v)?),
                }
                let mut unique = Vec::new();
                for item in n.items {
                    if !unique.contains(&item) {
                        unique.push(item);
                    }
                }
                n.items = unique;
                self.feature(n)
            }
            "addshape" | "addtextbox" | "addline" | "addpicture" => self.add_drawing(s, sheet, &m, call),
            _ => Err(failure(1004, "native drawing operation unavailable; groups/connectors and object clipboard are not simulated")),
        }
    }
    fn drawing_index(&self, s: &Session, nodes: &[Feature], v: &Value) -> HostResult<Feature> {
        if matches!(v, Value::String(_)) {
            let name = string(v)?;
            let sheet = nodes.first().and_then(|n| n.sheet).ok_or_else(|| failure(9, "no drawing"))?;
            let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?;
            for n in nodes {
                if view(sh, n)?.name.eq_ignore_ascii_case(&name) {
                    return Ok(n.clone());
                }
            }
            return Err(failure(9, "drawing name not found"));
        }
        let index =
            usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "drawing index"))?).map_err(|_| failure(9, "drawing index"))?;
        nodes.get(index).cloned().ok_or_else(|| failure(9, "drawing index"))
    }
    fn drawing_range(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "range owner"))?;
        if f.items.is_empty() || f.items.len() > 256 {
            return Err(failure(7, "ShapeRange budget"));
        }
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?;
        for n in &f.items {
            view(sh, n)?;
        }
        let m = call.member.to_ascii_lowercase();
        if call.kind == InvokeKind::Enumerate {
            arguments(call, &[])?;
            return self.object_array(f.items.iter().cloned().map(Feature::object).collect());
        }
        match m.as_str() {
            "count" => {
                arguments(call, &[])?;
                return Ok(Value::Long(f.items.len() as i32));
            }
            "item" | "" => {
                let a = arguments(call, &["Index"])?;
                let item = self.drawing_index(s, &f.items, required(&a, 0)?)?;
                return self.feature(item);
            }
            "parent" => {
                arguments(call, &[])?;
                return self.intern(Object::Sheet(sheet));
            }
            "application" => {
                arguments(call, &[])?;
                return self.intern(Object::Application);
            }
            "group" | "ungroup" => return Err(failure(1004, "arrange.group is alignment, not native persistent grouping")),
            "align" => {
                let a = arguments(call, &["AlignCmd", "RelativeTo"])?;
                if !absent(&a[1]) && super::address::boolean(&a[1])? {
                    return Err(failure(1004, "sheet-relative alignment unavailable"));
                }
                let how = match numeric_index(required(&a, 0)?)? {
                    0 => "left",
                    1 => "center",
                    2 => "right",
                    3 => "top",
                    4 => "middle",
                    5 => "bottom",
                    _ => return Err(failure(1004, "alignment option")),
                };
                let tag = f.items.first().and_then(|n| n.text.clone()).ok_or_else(|| failure(91, "range"))?;
                if f.items.iter().any(|n| n.text.as_ref() != Some(&tag)) {
                    return Err(failure(1004, "native alignment requires one drawing layer"));
                }
                self.drawing_command(
                    s,
                    sheet,
                    "arrange.align",
                    json!({"kind":tag,"ids":f.items.iter().map(|n|n.key).collect::<Vec<_>>(),"align":how}),
                )?;
                return Ok(Value::Empty);
            }
            _ => {}
        }
        if matches!(m.as_str(), "left" | "top" | "width" | "height") {
            let views = f.items.iter().map(|n| view(sh, n)).collect::<HostResult<Vec<_>>>()?;
            let left = views.iter().map(|v| sh.col_left(v.anchor.cell.col) + f64::from(v.anchor.dx)).fold(f64::INFINITY, f64::min);
            let top = views.iter().map(|v| sh.row_top(v.anchor.cell.row) + f64::from(v.anchor.dy)).fold(f64::INFINITY, f64::min);
            let right = views
                .iter()
                .map(|v| sh.col_left(v.anchor.cell.col) + f64::from(v.anchor.dx) + f64::from(v.anchor.width))
                .fold(f64::NEG_INFINITY, f64::max);
            let bottom = views
                .iter()
                .map(|v| sh.row_top(v.anchor.cell.row) + f64::from(v.anchor.dy) + f64::from(v.anchor.height))
                .fold(f64::NEG_INFINITY, f64::max);
            if reads(call) {
                arguments(call, &[])?;
                return Ok(Value::Double(match m.as_str() {
                    "left" => left,
                    "top" => top,
                    "width" => right - left,
                    _ => bottom - top,
                }));
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(438, "geometry requires Let"));
            }
            if f.items.len() > 1 && matches!(m.as_str(), "width" | "height") {
                return Err(failure(1004, "native drawings have no collective scaling transform"));
            }
            let a = arguments(call, &["Value"])?;
            let value = super::format_api::number(required(&a, 0)?)?;
            validate_geometry(&m, value)?;
            let origin = if m == "left" { left } else { top };
            let mut targets = Vec::new();
            for v in views {
                let next = if m == "left" {
                    sh.col_left(v.anchor.cell.col) + f64::from(v.anchor.dx) + value - origin
                } else if m == "top" {
                    sh.row_top(v.anchor.cell.row) + f64::from(v.anchor.dy) + value - origin
                } else {
                    value
                };
                validate_geometry(&m, next)?;
                targets.push((v, next));
            }
            if sh.is_protected() {
                return Err(failure(1004, "drawing objects are protected"));
            }
            for (v, next) in targets {
                self.set_geometry(s, sheet, &v, &m, next)?;
            }
            return Ok(Value::Empty);
        }
        if reads(call) && m != "delete" {
            let mut first = None;
            for n in &f.items {
                let v = self.drawing_dispatch(s, n, call)?;
                if let Some(old) = &first {
                    if old != &v {
                        return Ok(Value::Null);
                    }
                } else {
                    first = Some(v);
                }
            }
            return first.ok_or_else(|| failure(91, "range empty"));
        }
        // Preflight all identities, bounds and drawing protection before any bulk mutation.
        if sh.is_protected() {
            return Err(failure(1004, "drawing objects are protected"));
        }
        if call.kind == InvokeKind::Let {
            let a = arguments(call, &["Value"])?;
            let value = required(&a, 0)?;
            if matches!(m.as_str(), "left" | "top" | "width" | "height") {
                validate_geometry(&m, super::format_api::number(value)?)?;
            } else if m == "placement" {
                if !matches!(numeric_index(value)?, 1..=3) {
                    return Err(failure(1004, "placement"));
                }
            } else {
                return Err(failure(1004, "bulk property is not supported"));
            }
        } else if m != "delete" {
            return Err(failure(1004, "unsupported bulk drawing operation"));
        }
        for n in &f.items {
            self.drawing_dispatch(s, n, call)?;
        }
        Ok(Value::Empty)
    }
    fn drawing_command(&mut self, s: &mut Session, sheet: usize, command: &str, mut p: Json) -> HostResult<Json> {
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet deleted"))?;
        if sh.is_protected() {
            return Err(failure(1004, "drawing objects are protected"));
        }
        p["sheet"] = json!(sheet);
        s.execute(command, p).map_err(|e| failure(1004, e.to_string()))
    }
    fn set_geometry(&mut self, s: &mut Session, sheet: usize, v: &View, member: &str, n: f64) -> HostResult<Value> {
        validate_geometry(member, n)?;
        let mut p = json!({"kind":v.tag,"id":v.id});
        if matches!(member, "width" | "height") {
            p[member] = json!(n);
        } else {
            let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?;
            let mut c = v.anchor.cell;
            let offset = if member == "left" {
                c.col = sh.col_at(n);
                sh.col_left(c.col)
            } else {
                c.row = sh.row_at(n);
                sh.row_top(c.row)
            };
            p["at"] = json!(c.a1());
            p[if member == "left" { "dx" } else { "dy" }] = json!(n - offset);
        }
        self.drawing_command(s, sheet, "object.move", p)?;
        Ok(Value::Empty)
    }
    fn zorder(&mut self, s: &mut Session, sheet: usize, v: &View, n: i32) -> HostResult<Value> {
        let (id, p) = match n {
            0 => ("arrange.bringForward", json!({"toFront":true})),
            1 => ("arrange.sendBackward", json!({"toBack":true})),
            2 => ("arrange.bringForward", json!({})),
            3 => ("arrange.sendBackward", json!({})),
            _ => return Err(failure(1004, "unsupported z-order command")),
        };
        let mut p = p;
        p["kind"] = json!(v.tag);
        p["id"] = json!(v.id);
        self.drawing_command(s, sheet, id, p)?;
        Ok(Value::Empty)
    }
    fn add_drawing(&mut self, s: &mut Session, sheet: usize, member: &str, call: &HostCall) -> HostResult<Value> {
        let names = match member {
            "addshape" => &["Type", "Left", "Top", "Width", "Height"][..],
            "addtextbox" => &["Orientation", "Left", "Top", "Width", "Height"][..],
            "addline" => &["BeginX", "BeginY", "EndX", "EndY"][..],
            _ => &["Filename", "LinkToFile", "SaveWithDocument", "Left", "Top", "Width", "Height"][..],
        };
        let a = arguments(call, names)?;
        let (kind, left, top, width, height) = match member {
            "addshape" | "addtextbox" => {
                let ty = numeric_index(required(&a, 0)?)?;
                let kind = if member == "addtextbox" {
                    if ty != 1 {
                        return Err(failure(1004, "native text boxes are horizontal"));
                    }
                    "textBox"
                } else {
                    match ty {
                        1 => "rectangle",
                        5 => "roundedRectangle",
                        7 => "triangle",
                        9 => "ellipse",
                        _ => return Err(failure(1004, "AutoShape type is not represented natively")),
                    }
                };
                (
                    kind,
                    super::format_api::number(required(&a, 1)?)?,
                    super::format_api::number(required(&a, 2)?)?,
                    super::format_api::number(required(&a, 3)?)?,
                    super::format_api::number(required(&a, 4)?)?,
                )
            }
            "addline" => {
                let x = super::format_api::number(required(&a, 0)?)?;
                let y = super::format_api::number(required(&a, 1)?)?;
                ("line", x, y, super::format_api::number(required(&a, 2)?)? - x, super::format_api::number(required(&a, 3)?)? - y)
            }
            _ => {
                if super::address::boolean(required(&a, 1)?)? || !super::address::boolean(required(&a, 2)?)? {
                    return Err(failure(1004, "pictures must be embedded, not linked"));
                }
                (
                    "image",
                    super::format_api::number(required(&a, 3)?)?,
                    super::format_api::number(required(&a, 4)?)?,
                    super::format_api::number(required(&a, 5)?)?,
                    super::format_api::number(required(&a, 6)?)?,
                )
            }
        };
        validate_geometry("left", left)?;
        validate_geometry("top", top)?;
        if kind != "image" || width != -1.0 {
            validate_geometry("width", width)?;
        }
        if kind != "image" || height != -1.0 {
            validate_geometry("height", height)?;
        }
        self.objects.reserve(1)?;
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?;
        if sh.shapes.len().saturating_add(sh.images.len()).saturating_add(sh.charts.len()).saturating_add(sh.slicers.len()) >= super::MAX_OBJECTS {
            return Err(failure(7, "drawing collection budget"));
        }
        let at = CellRef::new(sh.row_at(top), sh.col_at(left));
        let dx = left - sh.col_left(at.col);
        let dy = top - sh.row_top(at.row);
        let p = json!({"kind":kind,"at":at.a1(),"width":width,"height":height});
        let (command, mut p) = if kind == "image" {
            ("insert.picture", json!({"path":string(required(&a,0)?)?,"at":at.a1(),"width":width,"height":height}))
        } else {
            ("insert.shape", p)
        };
        // Creation commands carry the exact offsets, so handle/allocation failure cannot leave a partially positioned drawing.
        p["dx"] = json!(dx);
        p["dy"] = json!(dy);
        let result = self.drawing_command(s, sheet, command, p)?;
        let key = if kind == "image" { "image" } else { "shape" };
        let id = result.get(key).and_then(Json::as_u64).ok_or_else(|| failure(1004, "drawing creation did not return identity"))?;
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?;
        let item = roots(sh, sheet)?
            .into_iter()
            .find(|n| n.key == id && n.text.as_deref() == Some(if kind == "image" { "image" } else { "shape" }))
            .ok_or_else(|| failure(91, "created drawing vanished"))?;
        self.feature(item)
    }
    fn drawing_child(&mut self, s: &mut Session, f: &Feature, v: &View, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let m = call.member.to_ascii_lowercase();
        if matches!(f.kind, "FillFormat" | "LineFormat" | "ColorFormat") {
            return self.shape_color(s, f, v, call);
        }
        if m == "parent" && reads(call) {
            arguments(call, &[])?;
            let mut parent = child(
                f,
                match f.kind {
                    "Characters" => "TextFrame",
                    "TextRange2" => "TextFrame2",
                    "ShapeFont" => "Characters",
                    "ShapeFont2" => "TextRange2",
                    _ => "Shape",
                },
                0,
            );
            if matches!(parent.kind, "Shape" | "TextFrame" | "TextFrame2") {
                parent.span = None;
            }
            return self.feature(parent);
        }
        if matches!(f.kind, "TextFrame" | "TextFrame2") {
            if (m == "characters" || m == "textrange") && reads(call) {
                let a = arguments(call, if m == "characters" { &["Start", "Length"] } else { &[] })?;
                let len = u32::try_from(v.text.encode_utf16().count()).map_err(|_| failure(7, "text budget"))?;
                let start = if a.is_empty() || absent(&a[0]) {
                    0
                } else {
                    u32::try_from(numeric_index(&a[0])?.checked_sub(1).ok_or_else(|| failure(5, "character index"))?)
                        .map_err(|_| failure(5, "character index"))?
                };
                let n = if a.len() < 2 || absent(&a[1]) {
                    len.saturating_sub(start)
                } else {
                    u32::try_from(numeric_index(&a[1])?).map_err(|_| failure(5, "character length"))?
                };
                if start > len || n > len.saturating_sub(start) {
                    return Err(failure(5, "text range bounds"));
                }
                let mut nnode = child(f, if f.kind == "TextFrame" { "Characters" } else { "TextRange2" }, 0);
                nnode.span = if a.iter().all(absent) { None } else { Some((start, n)) };
                return self.feature(nnode);
            }
            if m == "hastext" && reads(call) {
                arguments(call, &[])?;
                return Ok(Value::Long(if v.text.is_empty() { 0 } else { -1 }));
            }
            if reads(call) {
                arguments(call, &[])?;
                return match m.as_str() {
                    "horizontalalignment" => Ok(Value::Long(match v.format.align.unwrap_or_default() {
                        ShapeTextAlign::Left => -4131,
                        ShapeTextAlign::Center => -4108,
                        ShapeTextAlign::Right => -4152,
                    })),
                    "verticalanchor" if f.kind == "TextFrame2" => Ok(Value::Long(match v.format.anchor.unwrap_or_default() {
                        ShapeTextAnchor::Top => 1,
                        ShapeTextAnchor::Middle => 3,
                        ShapeTextAnchor::Bottom => 4,
                    })),
                    "readingorder" if f.kind == "TextFrame" => Ok(Value::Long(if v.format.rtl { -5004 } else { -5003 })),
                    "verticalalignment" => Ok(Value::Long(match v.format.anchor.unwrap_or_default() {
                        ShapeTextAnchor::Top => -4160,
                        ShapeTextAnchor::Middle => -4108,
                        ShapeTextAnchor::Bottom => -4107,
                    })),
                    _ => Err(failure(1004, "native text frame field unavailable")),
                };
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(438, "property requires Let"));
            }
            let a = arguments(call, &["Value"])?;
            let n = numeric_index(required(&a, 0)?)?;
            let patch = match m.as_str() {
                "verticalanchor" if f.kind == "TextFrame2" => {
                    json!({"anchor":match n{1=>"Top",3=>"Middle",4=>"Bottom",_=>return Err(failure(1004,"vertical anchor unavailable"))}})
                }
                "readingorder" if f.kind == "TextFrame" => {
                    json!({"rtl":match n{-5004=>true,-5003=>false,_=>return Err(failure(1004,"reading order unavailable"))}})
                }
                "horizontalalignment" => {
                    json!({"align":match n{-4131=>"Left",-4108=>"Center",-4152=>"Right",_=>return Err(failure(1004,"text alignment"))}})
                }
                "verticalalignment" => {
                    json!({"anchor":match n{-4160=>"Top",-4108=>"Middle",-4107=>"Bottom",_=>return Err(failure(1004,"text anchor"))}})
                }
                _ => return Err(failure(1004, "unsupported text frame write")),
            };
            self.drawing_command(s, sheet, "shape.setText", json!({"id":v.id,"format":patch}))?;
            return Ok(Value::Empty);
        }
        if matches!(f.kind, "Characters" | "TextRange2") {
            let units = v.text.encode_utf16().collect::<Vec<_>>();
            let (start, len) = f.span.unwrap_or((0, units.len() as u32));
            let a = usize::try_from(start).map_err(|_| failure(5, "text index"))?;
            let b = a.checked_add(len as usize).ok_or_else(|| failure(5, "text range"))?;
            let slice = units.get(a..b).ok_or_else(|| failure(91, "text range no longer exists"))?;
            if m == "font" && reads(call) {
                arguments(call, &[])?;
                return self.feature(child(f, if f.kind == "Characters" { "ShapeFont" } else { "ShapeFont2" }, 0));
            }
            if m == "count" || m == "length" {
                arguments(call, &[])?;
                return Ok(Value::Long(len as i32));
            }
            if reads(call) && m == "text" {
                arguments(call, &[])?;
                return Ok(Value::String(vba_values::VbaString::from_units(slice.to_vec()).map_err(|e| failure(7, e.to_string()))?));
            }
            let text = if m == "delete" && reads(call) {
                arguments(call, &[])?;
                String::new()
            } else {
                if call.kind != InvokeKind::Let {
                    return Err(failure(438, "property requires Let"));
                }
                let a = arguments(call, &["Value"])?;
                if call.kind != InvokeKind::Let || m != "text" {
                    return Err(failure(1004, "unsupported text operation"));
                }
                string(required(&a, 0)?)?
            };
            let mut replacement = units[..a].to_vec();
            replacement.extend(text.encode_utf16());
            replacement.extend_from_slice(&units[b..]);
            let text = String::from_utf16(&replacement).map_err(|_| failure(13, "text write splits a surrogate pair"))?;
            self.drawing_command(s, sheet, "shape.setText", json!({"id":v.id,"text":text}))?;
            return Ok(Value::Empty);
        }
        self.drawing_font(s, f, v, call)
    }
    fn drawing_font(&mut self, s: &mut Session, f: &Feature, v: &View, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let m = call.member.to_ascii_lowercase();
        if matches!(f.kind, "ShapeFont" | "ShapeFont2") {
            if f.kind == "ShapeFont2" && reads(call) && matches!(m.as_str(), "fill" | "line") {
                arguments(call, &[])?;
                return self.feature(child(f, if m == "fill" { "FillFormat" } else { "LineFormat" }, if m == "fill" { 3 } else { 4 }));
            }
            if reads(call) {
                arguments(call, &[])?;
                return match m.as_str() {
                    "name" => values::vba_text(v.format.font.as_deref().unwrap_or("Calibri")),
                    "size" => Ok(Value::Double(f64::from(v.format.font_size.unwrap_or(11.0)))),
                    "bold" => Ok(if f.kind == "ShapeFont" { Value::Boolean(v.format.bold) } else { Value::Long(if v.format.bold { -1 } else { 0 }) }),
                    "italic" => {
                        Ok(if f.kind == "ShapeFont" { Value::Boolean(v.format.italic) } else { Value::Long(if v.format.italic { -1 } else { 0 }) })
                    }
                    "color" => super::format_api::ole(v.format.text_color.unwrap_or(Color::rgb(0, 0, 0)), s),
                    _ => Err(failure(1004, "font field has no native shape text state")),
                };
            }
            if f.span.is_some_and(|(start, len)| start != 0 || len as usize != v.text.encode_utf16().count()) {
                return Err(failure(1004, "native shapes store one text format for the entire body"));
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(438, "property requires Let"));
            }
            let a = arguments(call, &["Value"])?;
            let value = required(&a, 0)?;
            let patch = match m.as_str() {
                "name" => json!({"font":string(value)?}),
                "size" => {
                    let n = super::format_api::number(value)?;
                    if !(1.0..=4000.0).contains(&n) {
                        return Err(failure(1004, "font size bounds"));
                    }
                    json!({"font_size":n})
                }
                "bold" => json!({"bold":super::address::boolean(value)?}),
                "italic" => json!({"italic":super::address::boolean(value)?}),
                "color" => json!({"text_color":super::format_api::rgb(value)?}),
                _ => return Err(failure(1004, "unsupported text font option")),
            };
            self.drawing_command(s, sheet, "shape.setText", json!({"id":v.id,"format":patch}))?;
            return Ok(Value::Empty);
        }
        Err(failure(438, "unsupported drawing child"))
    }
    fn shape_color(&mut self, s: &mut Session, f: &Feature, v: &View, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let m = call.member.to_ascii_lowercase();
        let color = match f.part {
            1 => v.fill,
            2 => v.line,
            3 => v.format.text_color.unwrap_or(Color::rgb(0, 0, 0)),
            _ => v.format.text_outline.map(|v| v.0).unwrap_or(Color::rgb(0, 0, 0)),
        };
        if m == "parent" && reads(call) {
            arguments(call, &[])?;
            return self.feature(child(
                f,
                if f.kind == "ColorFormat" {
                    if matches!(f.part, 1 | 3) { "FillFormat" } else { "LineFormat" }
                } else if f.part >= 3 {
                    "ShapeFont2"
                } else {
                    "Shape"
                },
                if f.kind == "ColorFormat" { f.part } else { 0 },
            ));
        }
        if m == "forecolor" && reads(call) {
            arguments(call, &[])?;
            return self.feature(child(f, "ColorFormat", f.part));
        }
        if reads(call) {
            arguments(call, &[])?;
            return match m.as_str() {
                "rgb" => super::format_api::ole(color, s),
                "objectthemecolor" => Ok(Value::Long(if let Color::Theme(i, _) = color { super::format_api::theme_number(i) } else { 0 })),
                "tintandshade" => Ok(Value::Double(if let Color::Theme(_, t) = color { f64::from(t) / 1000.0 } else { 0.0 })),
                "visible" => Ok(Value::Long(if f.part == 2 && v.format.no_line || f.part == 4 && v.format.text_outline.is_none() { 0 } else { -1 })),
                "weight" if f.part == 4 => Ok(Value::Double(f64::from(v.format.text_outline.map(|v| v.1).unwrap_or(1.0)))),
                "type" => Ok(Value::Long(1)),
                "solid" if matches!(f.part, 1 | 3) => Ok(Value::Empty),
                _ => Err(failure(1004, "no native independent fill/line option")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "property requires Let"));
        }
        let a = arguments(call, &["Value"])?;
        let value = required(&a, 0)?;
        if f.part >= 3 {
            if f.span.is_some_and(|(start, len)| start != 0 || len as usize != v.text.encode_utf16().count()) {
                return Err(failure(1004, "native text has one whole-body format"));
            }
            let current_width = v.format.text_outline.map(|v| v.1).unwrap_or(1.0);
            let patch = match m.as_str() {
                "visible" if f.part == 4 => json!({"text_outline":if super::address::boolean(value)?{json!([color,current_width])}else{Json::Null}}),
                "weight" if f.part == 4 => {
                    let n = super::format_api::number(value)?;
                    if !(0.0..=100.0).contains(&n) {
                        return Err(failure(1004, "outline width bounds"));
                    }
                    json!({"text_outline":[color,n]})
                }
                "rgb" | "objectthemecolor" | "tintandshade" => {
                    let col = match m.as_str() {
                        "rgb" => super::format_api::rgb(value)?,
                        "objectthemecolor" => Color::Theme(super::format_api::theme_index(numeric_index(value)?)?, 0),
                        _ => {
                            let Color::Theme(i, _) = color else {
                                return Err(failure(1004, "tint requires a theme color"));
                            };
                            let n = super::format_api::number(value)?;
                            if !(-1.0..=1.0).contains(&n) {
                                return Err(failure(1004, "tint bounds"));
                            }
                            Color::Theme(i, (n * 1000.0).round() as i16)
                        }
                    };
                    if f.part == 3 { json!({"text_color":col}) } else { json!({"text_outline":[col,current_width]}) }
                }
                _ => return Err(failure(1004, "text fill/outline option unavailable")),
            };
            self.drawing_command(s, sheet, "shape.setText", json!({"id":v.id,"format":patch}))?;
            return Ok(Value::Empty);
        }
        if m == "visible" {
            if f.part != 2 {
                return Err(failure(1004, "native shapes have no independent no-fill state"));
            }
            self.drawing_command(s, sheet, "shape.setText", json!({"id":v.id,"format":{"no_line":!super::address::boolean(value)?}}))?;
            return Ok(Value::Empty);
        }
        let p = match m.as_str() {
            "rgb" => {
                let col = super::format_api::rgb(value)?;
                let wb = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb;
                let [r, g, b] = col.resolve(&wb.theme).ok_or_else(|| failure(13, "color"))?;
                json!(format!("#{r:02X}{g:02X}{b:02X}"))
            }
            "objectthemecolor" => json!({"theme":super::format_api::theme_index(numeric_index(value)?)?,"tint":0.0}),
            "tintandshade" => {
                let Color::Theme(i, _) = color else {
                    return Err(failure(1004, "tint requires a native theme color"));
                };
                let t = super::format_api::number(value)?;
                if !(-1.0..=1.0).contains(&t) {
                    return Err(failure(1004, "tint bounds"));
                }
                json!({"theme":i,"tint":t})
            }
            _ => return Err(failure(1004, "fill or line option unavailable")),
        };
        self.drawing_command(
            s,
            sheet,
            if f.part == 1 { "chart.shapeFill" } else { "chart.shapeOutline" },
            json!({"kind":"shape","id":v.id,"color":p}),
        )?;
        Ok(Value::Empty)
    }
    fn series_format(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let m = call.member.to_ascii_lowercase();
        if f.kind == "SeriesFormat" {
            arguments(call, &[])?;
            return match m.as_str() {
                "fill" if f.text.as_deref() != Some("trendline") => self.feature(child(f, "FillFormat", 1)),
                "line" => self.feature(child(f, "LineFormat", 2)),
                "parent" => {
                    let mut n = f.clone();
                    n.kind = if f.text.as_deref() == Some("trendline") { "Trendline" } else { "Series" };
                    n.text = None;
                    n.part = 0;
                    self.feature(n)
                }
                _ => Err(failure(1004, "chart format field unavailable")),
            };
        }
        if m == "forecolor" && reads(call) {
            arguments(call, &[])?;
            return self.feature(child(f, "ColorFormat", f.part));
        }
        if m == "parent" && reads(call) {
            arguments(call, &[])?;
            return self.feature(child(
                f,
                if f.kind == "ColorFormat" { if f.part == 1 { "FillFormat" } else { "LineFormat" } } else { "SeriesFormat" },
                if f.kind == "ColorFormat" { f.part } else { 0 },
            ));
        }
        if m == "rgb" || f.text.as_deref() == Some("trendline") && matches!(m.as_str(), "weight" | "dashstyle") {
            let mut n = f.clone();
            n.kind = if f.text.as_deref() == Some("trendline") {
                "TrendlineBorder"
            } else if f.part == 1 {
                "SeriesInterior"
            } else {
                "SeriesBorder"
            };
            let mut call = call.clone();
            call.member = if m == "rgb" {
                "Color"
            } else if m == "dashstyle" {
                "DashStyle"
            } else {
                "Weight"
            }
            .into();
            return self.series_api(s, &Object::Feature(n), &call).ok_or_else(|| failure(438, "series format"))?;
        }
        Err(failure(1004, "native series have one shared fill/line color"))
    }
}
fn validate_geometry(member: &str, n: f64) -> HostResult<()> {
    let valid = n.is_finite()
        && match member {
            "width" => (1.0..=10000.0).contains(&n),
            "height" => (0.0..=10000.0).contains(&n),
            _ => (0.0..=1.0e9).contains(&n),
        };
    if valid { Ok(()) } else { Err(failure(1004, "drawing geometry outside native bounds")) }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("ChartObjects","ShapeRange","Property Get ShapeRange() As ShapeRange",super::MemberAvailability::Implemented),
        ("Trendline","Format","Property Get Format() As ChartFormat",super::MemberAvailability::Implemented),
        ("Trendline","Border","Property Get Border() As Border",super::MemberAvailability::Implemented),
        ("TextFrame2","VerticalAnchor","Property VerticalAnchor As MsoVerticalAnchor",super::MemberAvailability::Implemented),
        ("TextFrame","ReadingOrder","Property ReadingOrder As Long",super::MemberAvailability::Implemented),
        ("Font2","Fill","Property Get Fill() As Office.FillFormat",super::MemberAvailability::Implemented),
        ("Font2","Line","Property Get Line() As Office.LineFormat",super::MemberAvailability::Implemented),
        ("LineFormat","Weight","Property Weight As Single",super::MemberAvailability::Implemented),
        ("LineFormat","DashStyle","Property DashStyle As MsoLineDashStyle",super::MemberAvailability::Implemented),
        ("Worksheet","Shapes","Property Get Shapes() As Shapes",super::MemberAvailability::Implemented),
        ("Shapes","Parent","Property Get Parent() As Worksheet",super::MemberAvailability::Implemented),
        ("Shapes","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Shapes","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("Shapes","Item","Function Item(Index As Variant) As Shape",super::MemberAvailability::Implemented),
        ("Shapes","Range","Property Get Range(Index As Variant) As ShapeRange",super::MemberAvailability::Implemented),
        ("Shapes","AddShape","Function AddShape(Type As MsoAutoShapeType, Left As Single, Top As Single, Width As Single, Height As Single) As Shape",super::MemberAvailability::Implemented),
        ("Shapes","AddTextbox","Function AddTextbox(Orientation As MsoTextOrientation, Left As Single, Top As Single, Width As Single, Height As Single) As Shape",super::MemberAvailability::Implemented),
        ("Shapes","AddLine","Function AddLine(BeginX As Single, BeginY As Single, EndX As Single, EndY As Single) As Shape",super::MemberAvailability::Implemented),
        ("Shapes","AddPicture","Function AddPicture(Filename As String, LinkToFile As MsoTriState, SaveWithDocument As MsoTriState, Left As Single, Top As Single, Width As Single, Height As Single) As Shape",super::MemberAvailability::Implemented),
        ("Shape","Parent","Property Get Parent() As Worksheet",super::MemberAvailability::Implemented),
        ("Shape","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Shape","Left","Property Left As Single",super::MemberAvailability::Implemented),
        ("Shape","Top","Property Top As Single",super::MemberAvailability::Implemented),
        ("Shape","Width","Property Width As Single",super::MemberAvailability::Implemented),
        ("Shape","Height","Property Height As Single",super::MemberAvailability::Implemented),
        ("Shape","Placement","Property Placement As XlPlacement",super::MemberAvailability::Implemented),
        ("Shape","Name","Property Get Name() As String",super::MemberAvailability::Implemented),
        ("Shape","Type","Property Get Type() As MsoShapeType",super::MemberAvailability::Implemented),
        ("Shape","AutoShapeType","Property Get AutoShapeType() As MsoAutoShapeType",super::MemberAvailability::Implemented),
        ("Shape","ZOrderPosition","Property Get ZOrderPosition() As Long",super::MemberAvailability::Implemented),
        ("Shape","OnAction","Property OnAction As String",super::MemberAvailability::Implemented),
        ("Shape","Fill","Property Get Fill() As Office.FillFormat",super::MemberAvailability::Implemented),
        ("Shape","Line","Property Get Line() As Office.LineFormat",super::MemberAvailability::Implemented),
        ("Shape","TextFrame","Property Get TextFrame() As TextFrame",super::MemberAvailability::Implemented),
        ("Shape","TextFrame2","Property Get TextFrame2() As Office.TextFrame2",super::MemberAvailability::Implemented),
        ("Shape","Delete","Sub Delete()",super::MemberAvailability::Implemented),
        ("Shape","ZOrder","Sub ZOrder(ZOrderCmd As MsoZOrderCmd)",super::MemberAvailability::Implemented),
        ("Shape","Ungroup","Function Ungroup() As Variant",super::MemberAvailability::OutOfScope),
        ("Shape","Copy","Function Copy() As Variant",super::MemberAvailability::OutOfScope),
        ("Shape","Cut","Function Cut() As Variant",super::MemberAvailability::OutOfScope),
        ("Shape","CopyPicture","Function CopyPicture() As Variant",super::MemberAvailability::OutOfScope),
        ("Shape","Duplicate","Function Duplicate() As Variant",super::MemberAvailability::OutOfScope),
        ("ShapeRange","Parent","Property Get Parent() As Worksheet",super::MemberAvailability::Implemented),
        ("ShapeRange","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("ShapeRange","Left","Property Left As Single",super::MemberAvailability::Implemented),
        ("ShapeRange","Top","Property Top As Single",super::MemberAvailability::Implemented),
        ("ShapeRange","Width","Property Width As Single",super::MemberAvailability::Implemented),
        ("ShapeRange","Height","Property Height As Single",super::MemberAvailability::Implemented),
        ("ShapeRange","Name","Property Get Name() As String",super::MemberAvailability::Implemented),
        ("ShapeRange","Type","Property Get Type() As MsoShapeType",super::MemberAvailability::Implemented),
        ("ShapeRange","AutoShapeType","Property Get AutoShapeType() As MsoAutoShapeType",super::MemberAvailability::Implemented),
        ("ShapeRange","ZOrderPosition","Property Get ZOrderPosition() As Long",super::MemberAvailability::Implemented),
        ("ShapeRange","Fill","Property Get Fill() As Office.FillFormat",super::MemberAvailability::Implemented),
        ("ShapeRange","Line","Property Get Line() As Office.LineFormat",super::MemberAvailability::Implemented),
        ("ShapeRange","TextFrame","Property Get TextFrame() As TextFrame",super::MemberAvailability::Implemented),
        ("ShapeRange","TextFrame2","Property Get TextFrame2() As Office.TextFrame2",super::MemberAvailability::Implemented),
        ("ShapeRange","Delete","Sub Delete()",super::MemberAvailability::Implemented),
        ("ShapeRange","ZOrder","Sub ZOrder(ZOrderCmd As MsoZOrderCmd)",super::MemberAvailability::Implemented),
        ("ShapeRange","Group","Function Group() As Variant",super::MemberAvailability::OutOfScope),
        ("ShapeRange","Ungroup","Function Ungroup() As Variant",super::MemberAvailability::OutOfScope),
        ("ShapeRange","Duplicate","Function Duplicate() As Variant",super::MemberAvailability::OutOfScope),
        ("Shape","ID","Property Get ID() As Long",super::MemberAvailability::Implemented),
        ("Shape","Chart","Property Get Chart() As Chart",super::MemberAvailability::Implemented),
        ("Shape","HasChart","Property Get HasChart() As MsoTriState",super::MemberAvailability::Implemented),
        ("Shape","AlternativeText","Property Get AlternativeText() As String",super::MemberAvailability::Implemented),
        ("Shape","TopLeftCell","Property Get TopLeftCell() As Range",super::MemberAvailability::Implemented),
        ("Shape","BottomRightCell","Property Get BottomRightCell() As Range",super::MemberAvailability::Implemented),
        ("Shape","IncrementLeft","Sub IncrementLeft(Increment As Single)",super::MemberAvailability::Implemented),
        ("Shape","IncrementTop","Sub IncrementTop(Increment As Single)",super::MemberAvailability::Implemented),
        ("ShapeRange","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("ShapeRange","Item","Function Item(Index As Variant) As Shape",super::MemberAvailability::Implemented),
        ("ShapeRange","Align","Sub Align(AlignCmd As MsoAlignCmd, RelativeTo As MsoTriState)",super::MemberAvailability::Implemented),
        ("ChartObject","ShapeRange","Property Get ShapeRange() As ShapeRange",super::MemberAvailability::Implemented),
        ("ChartObject","ZOrder","Property Get ZOrder() As Variant",super::MemberAvailability::Implemented),
        ("ChartObject","BottomRightCell","Property Get BottomRightCell() As Range",super::MemberAvailability::Implemented),
        ("Series","Format","Property Get Format() As ChartFormat",super::MemberAvailability::Implemented),
        ("ChartFormat","Parent","Property Get Parent() As Series",super::MemberAvailability::Implemented),
        ("ChartFormat","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("ChartFormat","Fill","Property Get Fill() As Office.FillFormat",super::MemberAvailability::Implemented),
        ("ChartFormat","Line","Property Get Line() As Office.LineFormat",super::MemberAvailability::Implemented),
        ("TextFrame","Parent","Property Get Parent() As Shape",super::MemberAvailability::Implemented),
        ("TextFrame","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("TextFrame","Characters","Property Get Characters(Optional Start As Variant, Optional Length As Variant) As Characters",super::MemberAvailability::Implemented),
        ("TextFrame","HorizontalAlignment","Property HorizontalAlignment As XlHAlign",super::MemberAvailability::Implemented),
        ("TextFrame","VerticalAlignment","Property VerticalAlignment As XlVAlign",super::MemberAvailability::Implemented),
        ("Characters","Parent","Property Get Parent() As TextFrame",super::MemberAvailability::Implemented),
        ("Characters","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Characters","Text","Property Text As String",super::MemberAvailability::Implemented),
        ("Characters","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("Characters","Font","Property Get Font() As Font",super::MemberAvailability::Implemented),
        ("Characters","Delete","Function Delete() As Variant",super::MemberAvailability::Implemented),
        ("TextFrame2","Parent","Property Get Parent() As Shape",super::MemberAvailability::Implemented),
        ("TextFrame2","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("TextFrame2","TextRange","Property Get TextRange() As Office.TextRange2",super::MemberAvailability::Implemented),
        ("TextFrame2","HasText","Property Get HasText() As MsoTriState",super::MemberAvailability::Implemented),
        ("TextRange2","Parent","Property Get Parent() As Office.TextFrame2",super::MemberAvailability::Implemented),
        ("TextRange2","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("TextRange2","Text","Property Text As String",super::MemberAvailability::Implemented),
        ("TextRange2","Length","Property Get Length() As Long",super::MemberAvailability::Implemented),
        ("TextRange2","Font","Property Get Font() As Office.Font2",super::MemberAvailability::Implemented),
        ("TextRange2","Delete","Sub Delete()",super::MemberAvailability::Implemented),
        ("Font2","Parent","Property Get Parent() As Office.TextRange2",super::MemberAvailability::Implemented),
        ("Font2","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Font2","Name","Property Name As String",super::MemberAvailability::Implemented),
        ("Font2","Size","Property Size As Single",super::MemberAvailability::Implemented),
        ("Font2","Bold","Property Bold As MsoTriState",super::MemberAvailability::Implemented),
        ("Font2","Italic","Property Italic As MsoTriState",super::MemberAvailability::Implemented),
        ("FillFormat","Parent","Property Get Parent() As Object",super::MemberAvailability::Implemented),
        ("FillFormat","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("FillFormat","ForeColor","Property Get ForeColor() As Office.ColorFormat",super::MemberAvailability::Implemented),
        ("FillFormat","Visible","Property Get Visible() As MsoTriState",super::MemberAvailability::Implemented),
        ("LineFormat","Parent","Property Get Parent() As Object",super::MemberAvailability::Implemented),
        ("LineFormat","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("LineFormat","ForeColor","Property Get ForeColor() As Office.ColorFormat",super::MemberAvailability::Implemented),
        ("LineFormat","Visible","Property Visible As MsoTriState",super::MemberAvailability::Implemented),
        ("FillFormat","Type","Property Get Type() As MsoFillType",super::MemberAvailability::Implemented),
        ("FillFormat","Solid","Sub Solid()",super::MemberAvailability::Implemented),
        ("ColorFormat","Parent","Property Get Parent() As Object",super::MemberAvailability::Implemented),
        ("ColorFormat","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("ColorFormat","RGB","Property RGB As Long",super::MemberAvailability::Implemented),
        ("ColorFormat","ObjectThemeColor","Property ObjectThemeColor As MsoThemeColorIndex",super::MemberAvailability::Implemented),
        ("ColorFormat","TintAndShade","Property TintAndShade As Single",super::MemberAvailability::Implemented),
    ].into_iter().map(|(owner,name,signature,availability)|super::MemberDescriptor{owner,name,signature,availability,implementation_symbol:"crates/vba/src/drawing_api.rs::Bridge::drawing_api"}).collect()
}
