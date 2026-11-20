use super::{
    Anchor, CellPicture, CellRef, Color, EngineError, Image, Json, MAX_OBJECT_OFFSET, ObjKind, ObjRef, PROTECTED, Result, Session, Shape, ShapeKind,
    Sparkline, SparklineKind, anchor_mode_name, anchor_mode_param, bad, bool_param, cell_param, cell_param_on, edit, f64_param, json,
    locked_cell_error, ok, size_cap, str_param, target_sheet, u32_param,
};

#[cfg(test)]
pub(super) fn compiled() -> bool {
    true
}

pub(super) fn insert_sparkline(s: &mut Session, p: &Json) -> Result<Json> {
    let src = str_param(p, "range").ok_or_else(|| bad("insert.sparkline", "missing `range`"))?.to_string();
    let loc = cell_param(s, p, "location")?.ok_or_else(|| bad("insert.sparkline", "missing `location`"))?;
    let kind = match str_param(p, "type") {
        Some("column") => SparklineKind::Column,
        Some("winLoss") => SparklineKind::WinLoss,
        _ => SparklineKind::Line,
    };
    let markers = bool_param(p, "markers").unwrap_or(false);
    let sheet = s.doc()?.wb.active_sheet();
    edit(s, |cx| {
        let sh = cx.sheet_mut(sheet)?;
        sh.sparklines.retain(|x| x.cell != loc);
        sh.sparklines.push(Sparkline { cell: loc, source: src.clone(), kind, color: Color::Theme(4, 0), markers });
        Ok(Json::Null)
    })
}

pub(super) fn insert_picture(s: &mut Session, p: &Json) -> Result<Json> {
    let placement = str_param(p, "placement").unwrap_or("overCells");
    if !matches!(placement, "cell" | "overCells") {
        return Err(bad("insert.picture", "placement must be cell or overCells"));
    }
    let data = if let Some(b) = str_param(p, "base64") {
        crate::io::base64_decode(b).ok_or_else(|| bad("insert.picture", "invalid base64"))?
    } else if let Some(path) = str_param(p, "path") {
        crate::io::read_file(path)?
    } else {
        s.ui_requests.push(crate::UiRequest::Dialog("insertPicture".into(), json!({"placement": placement})));
        return ok();
    };
    if placement == "cell" {
        let mime = validate_cell_picture(&data)?;
        let d = s.doc()?;
        let sheet = target_sheet(s, p)?;
        let sh = d.wb.sheet(sheet).ok_or(EngineError::NoDocument)?;
        let at = match str_param(p, "at") {
            Some(a) => CellRef::parse(a).ok_or_else(|| bad("insert.picture", "invalid cell address"))?,
            None => d.selection.active,
        };
        let at = sh.merge_at(at).map(|r| r.start).unwrap_or(at);
        let style = sh.style_id(at);
        if sh.is_protected() && d.wb.styles.get(style).protection.locked && locked_cell_error(sh, at, &d.unlocked_ranges).is_some() {
            return Err(EngineError::Other("The cell you're trying to change is on a protected sheet.".into()));
        }
        let picture = CellPicture { data, mime: mime.into(), alt: str_param(p, "alt").unwrap_or("").into() };
        return edit(s, |cx| {
            cx.sheet_mut(sheet)?.set_picture(at, std::sync::Arc::new(picture));
            cx.touch(sheet, at);
            Ok(json!({"cell": at.a1(), "placement": "cell"}))
        });
    }
    if data.len() > 64 * 1024 * 1024 {
        return Err(bad("insert.picture", "image larger than 64 MB"));
    }
    let mime = str_param(p, "mime")
        .map(str::to_string)
        .unwrap_or_else(|| if data.starts_with(&[0x89, b'P', b'N', b'G']) { "image/png".into() } else { "image/jpeg".into() });
    let (w, h) = crate::io::image_size(&data).unwrap_or((320, 240));
    let d = s.doc()?;
    // Sheet protection locks drawing objects, as for a text box's text.
    if d.wb.active().is_some_and(|sh| sh.is_protected()) {
        return Err(EngineError::Other(PROTECTED.into()));
    }
    let id = d.wb.next_object_id();
    let at = cell_param_on(s, p, "at", target_sheet(s, p)?)?.unwrap_or(d.selection.active);
    let scale = (640.0 / w.max(1) as f32).min(1.0);
    let dimension = |key, original: u32, default: f32| -> Result<f32> {
        let n = f64_param(p, key).unwrap_or(f64::from(default));
        let n = if n == -1.0 { f64::from(original) } else { n };
        if !n.is_finite() || !(0.0..=10000.0).contains(&n) {
            return Err(bad("insert.picture", "picture dimensions exceed native bounds"));
        }
        Ok(n as f32)
    };
    let width = dimension("width", w, w as f32 * scale)?;
    let height = dimension("height", h, h as f32 * scale)?;
    let img = Image {
        identity: Default::default(),
        id,
        anchor: Anchor {
            cell: at,
            dx: f64_param(p, "dx").unwrap_or(0.0) as f32,
            dy: f64_param(p, "dy").unwrap_or(0.0) as f32,
            width,
            height,
            // Excel inserts pictures as "Move but don't size with cells" (`editAs="oneCell"`).
            mode: anchor_mode_param(p).unwrap_or(gridcraft_model::AnchorMode::MoveOnly),
        },
        data: data.into(),
        mime,
        alt: str_param(p, "alt").unwrap_or("").into(),
    };
    let sheet = target_sheet(s, p)?;
    edit(s, |cx| {
        cx.sheet_mut(sheet)?.images.push(img);
        Ok(json!({"image": id}))
    })
}

/// Validate before mutating the workbook; header dimensions alone do not prove the pixels decode.
fn validate_cell_picture(data: &[u8]) -> Result<&'static str> {
    if data.len() > 16 * 1024 * 1024 {
        return Err(bad("insert.picture", "cell pictures must be at most 16 MiB"));
    }
    let format = image::guess_format(data).map_err(|_| bad("insert.picture", "choose a valid PNG or JPEG picture"))?;
    let mime = match format {
        image::ImageFormat::Png => "image/png",
        image::ImageFormat::Jpeg => "image/jpeg",
        _ => return Err(bad("insert.picture", "cell pictures support PNG and JPEG")),
    };
    let (width, height) = image::ImageReader::with_format(std::io::Cursor::new(data), format)
        .into_dimensions()
        .map_err(|_| bad("insert.picture", "invalid image dimensions"))?;
    if width == 0 || height == 0 || width > 8192 || height > 8192 || u64::from(width) * u64::from(height) > 16_000_000 {
        return Err(bad("insert.picture", "cell pictures are limited to 8192 pixels per side and 16 million pixels"));
    }
    let mut reader = image::ImageReader::with_format(std::io::Cursor::new(data), format);
    let mut limits = image::Limits::default();
    limits.max_image_width = Some(8192);
    limits.max_image_height = Some(8192);
    limits.max_alloc = Some(64 * 1024 * 1024);
    reader.limits(limits);
    reader.decode().map_err(|_| bad("insert.picture", "picture cannot be decoded within the 64 MiB limit"))?;
    Ok(mime)
}

pub(super) fn insert_shape(s: &mut Session, p: &Json) -> Result<Json> {
    let kind = match str_param(p, "kind").unwrap_or("rectangle") {
        "roundedRectangle" => ShapeKind::RoundedRectangle,
        "ellipse" | "oval" => ShapeKind::Ellipse,
        "triangle" => ShapeKind::Triangle,
        "line" => ShapeKind::Line,
        "arrow" => ShapeKind::Arrow,
        "textBox" => ShapeKind::TextBox,
        _ => ShapeKind::Rectangle,
    };
    let d = s.doc()?;
    let id = d.wb.next_object_id();
    let at = cell_param_on(s, p, "at", target_sheet(s, p)?)?.unwrap_or(d.selection.active);
    let fill = crate::cmd::format::color_param(p.get("fill")).unwrap_or(if kind == ShapeKind::TextBox {
        Color::rgb(255, 255, 255)
    } else {
        Color::Theme(4, 0)
    });
    let line = crate::cmd::format::color_param(p.get("line")).unwrap_or(if kind == ShapeKind::TextBox {
        Color::rgb(0x80, 0x80, 0x80)
    } else {
        Color::Theme(4, -250)
    });
    let shape = Shape {
        identity: Default::default(),
        on_action: None,
        id,
        kind,
        anchor: Anchor {
            cell: at,
            dx: f64_param(p, "dx").unwrap_or(0.0) as f32,
            dy: f64_param(p, "dy").unwrap_or(0.0) as f32,
            width: f64_param(p, "width").unwrap_or(144.0) as f32,
            height: f64_param(p, "height").unwrap_or(if kind == ShapeKind::Line { 0.0 } else { 96.0 }) as f32,
            mode: anchor_mode_param(p).unwrap_or_default(),
        },
        fill,
        line,
        text: str_param(p, "text").unwrap_or("").into(),
        format: Default::default(),
        rotation: 0.0,
        math: Default::default(),
    };
    let sheet = target_sheet(s, p)?;
    edit(s, |cx| {
        cx.sheet_mut(sheet)?.shapes.push(shape);
        Ok(json!({"shape": id}))
    })
}

/// A text box holds at most as many characters as a cell.
const SHAPE_TEXT_MAX_CHARS: usize = 32_767;

pub(super) fn shape_set_text(s: &mut Session, p: &Json) -> Result<Json> {
    let id = u32_param(p, "id").ok_or_else(|| bad("shape.setText", "missing shape id"))?;
    let sheet = target_sheet(s, p)?;
    let shape = s.doc()?.wb.sheet(sheet).and_then(|sh| sh.shapes.iter().find(|x| x.id == id)).ok_or_else(|| bad("shape.setText", "no such shape"))?;
    if shape.format.control_inert {
        return Err(bad("shape.setText", "unsupported native control is retained without editing"));
    }
    let text = str_param(p, "text");
    if text.is_some() && !matches!(shape.kind, ShapeKind::TextBox | ShapeKind::WordArt) && shape.format.form_control.is_none() {
        return Err(bad("shape.setText", "the shape is not a text box"));
    }
    if shape.format.form_control.is_some() && text.is_some_and(|t| t.len() > 16384 || !gridcraft_model::form_controls::valid_control_text(t)) {
        return Err(bad("shape.setText", "native control caption budget"));
    }
    if text.is_some_and(|t| t.chars().nth(SHAPE_TEXT_MAX_CHARS).is_some()) {
        return Err(bad("shape.setText", "text is longer than 32767 characters"));
    }
    if shape.kind == ShapeKind::WordArt {
        if text.is_some_and(str::is_empty) {
            return Err(bad("shape.setText", "WordArt text must not be empty"));
        }
        if text.is_some_and(|t| t.chars().nth(crate::cmd::wordart::MAX_CHARS).is_some()) {
            return Err(bad("shape.setText", format!("`text` is longer than {} characters", crate::cmd::wordart::MAX_CHARS)));
        }
    }
    let mut format = shape.format.clone();
    if let Some(patch) = p.get("format") {
        let fields = patch.as_object().ok_or_else(|| bad("shape.setText", "format must be an object"))?;
        let mut merged = serde_json::to_value(&format).map_err(|e| bad("shape.setText", e.to_string()))?;
        for (k, v) in fields {
            if !["no_line", "font_size", "bold", "italic", "font", "text_color", "text_outline", "align", "anchor", "rtl"].contains(&k.as_str()) {
                return Err(bad("shape.setText", "unsupported native text format"));
            }
            merged[k] = v.clone();
        }
        format = serde_json::from_value(merged).map_err(|e| bad("shape.setText", e.to_string()))?;
        if format.font_size.is_some_and(|v| !v.is_finite() || !(1.0..=4000.0).contains(&v))
            || format.font.as_ref().is_some_and(|n| n.len() > 256)
            || format.text_outline.is_some_and(|(_, w)| !w.is_finite() || !(0.0..=100.0).contains(&w))
        {
            return Err(bad("shape.setText", "text format bounds"));
        }
    } else if text.is_none() {
        return Err(bad("shape.setText", "text or format required"));
    }
    if s.doc()?.wb.sheet(sheet).is_some_and(|sh| sh.is_protected()) {
        return Err(EngineError::Other(PROTECTED.into()));
    }
    if text.is_none_or(|t| t == shape.text) && format == shape.format {
        return ok();
    }
    edit(s, |cx| {
        let shape = cx.sheet_mut(sheet)?.shapes.iter_mut().find(|x| x.id == id).ok_or_else(|| bad("shape.setText", "shape deleted"))?;
        if let Some(t) = text {
            shape.text = t.into();
        }
        shape.format = format;
        Ok(Json::Null)
    })
}

/// The floating object a `{kind, id}` command names on the active sheet. Unknown kinds and
/// objects are errors, so a stale id never succeeds silently.
fn object_param(s: &Session, p: &Json, cmd: &str) -> Result<(usize, ObjRef)> {
    let id = u32_param(p, "id").ok_or_else(|| bad(cmd, "missing `id`"))?;
    let kind =
        str_param(p, "kind").and_then(ObjKind::parse).ok_or_else(|| bad(cmd, "kind must be chart, image, shape, slicer, equation or control"))?;
    let d = s.doc()?;
    let sheet = target_sheet(s, p)?;
    let o = ObjRef::new(kind, id);
    if !d.wb.sheet(sheet).is_some_and(|sh| sh.has_object(o)) {
        return Err(bad(cmd, "no such object"));
    }
    // A locked form control on a protected sheet stays as it is.
    crate::cmd::controls::check_objects_unlocked(s, sheet, kind, &[id])?;
    Ok((sheet, o))
}

pub(super) fn delete_object(s: &mut Session, p: &Json) -> Result<Json> {
    let (sheet, o) = object_param(s, p, "object.delete")?;
    edit(s, |cx| {
        cx.sheet_mut(sheet)?.remove_object(o);
        Ok(json!({"deleted": o}))
    })
}

pub(super) fn move_object(s: &mut Session, p: &Json) -> Result<Json> {
    let (sheet, o) = object_param(s, p, "object.move")?;
    let id = o.id;
    let kind = o.kind.name();
    let at = cell_param_on(s, p, "at", sheet)?;
    let caption = match p.get("caption") {
        None => None,
        Some(Json::String(t)) if t.len() <= 65536 => Some(t.clone()),
        _ => return Err(bad("object.move", "caption bounds")),
    };
    let columns = match p.get("columns") {
        None => None,
        Some(v) => Some(v.as_u64().filter(|n| (1..=10).contains(n)).ok_or_else(|| bad("object.move", "columns bounds"))? as u32),
    };
    if kind != "slicer" && (caption.is_some() || columns.is_some()) {
        return Err(bad("object.move", "caption/columns require a slicer"));
    }
    let rotation = f64_param(p, "rotation").filter(|v| v.is_finite());
    if rotation.is_some() {
        let ink = s.doc()?.wb.sheet(sheet).is_some_and(|sh| sh.shapes.iter().any(|x| x.id == o.id && x.kind == ShapeKind::Ink));
        if o.kind != ObjKind::Shape || ink {
            return Err(bad("object.move", "only shapes other than ink rotate"));
        }
        if s.doc()?.wb.sheet(sheet).is_some_and(|sh| sh.is_protected()) {
            return Err(EngineError::Other(PROTECTED.into()));
        }
    }
    edit(s, |cx| {
        let sh = cx.sheet_mut(sheet)?;
        let Some(a) = sh.object_anchor_mut(o) else { return Err(bad("object.move", "no such object")) };
        if let Some(at) = at {
            a.cell = at;
        }
        if let Some(v) = f64_param(p, "dx") {
            a.dx = v.clamp(-MAX_OBJECT_OFFSET, MAX_OBJECT_OFFSET) as f32;
        }
        if let Some(v) = f64_param(p, "dy") {
            a.dy = v.clamp(-MAX_OBJECT_OFFSET, MAX_OBJECT_OFFSET) as f32;
        }
        if let Some(v) = f64_param(p, "width") {
            a.width = v.clamp(1.0, size_cap(a.width)) as f32;
        }
        if let Some(v) = f64_param(p, "height") {
            a.height = v.clamp(0.0, size_cap(a.height)) as f32;
        }
        if let Some(sl) = sh.slicers.iter_mut().find(|sl| kind == "slicer" && sl.id == id) {
            if let Some(caption) = caption {
                sl.caption = caption;
            }
            if let Some(columns) = columns {
                sl.columns = columns;
            }
        }
        if let Some(r) = rotation
            && let Some(shape) = sh.shapes.iter_mut().find(|x| x.id == o.id)
        {
            shape.rotation = r.rem_euclid(360.0) as f32;
        }
        Ok(Json::Null)
    })
}

/// Sets an object's anchor mode (Excel's "Move and size with cells" / "Move but don't size with
/// cells" / "Don't move or size with cells"), so it follows — or ignores — its rows and columns
/// when the sheet is sorted or edited structurally.
pub(super) fn set_anchor_mode(s: &mut Session, p: &Json) -> Result<Json> {
    let (sheet, o) = object_param(s, p, "object.setAnchorMode")?;
    let mode = anchor_mode_param(p).ok_or_else(|| bad("object.setAnchorMode", "mode must be moveAndSize, moveOnly or absolute"))?;
    edit(s, |cx| {
        let Some(a) = cx.sheet_mut(sheet)?.object_anchor_mut(o) else { return Err(bad("object.setAnchorMode", "no such object")) };
        a.mode = mode;
        Ok(json!({"mode": anchor_mode_name(mode)}))
    })
}

pub(super) fn shape_set_macro(s: &mut Session, p: &Json) -> Result<Json> {
    let id = p.get("id").and_then(Json::as_u64).and_then(|n| u32::try_from(n).ok()).ok_or_else(|| bad("shape.setMacro", "missing shape id"))?;
    let entry = match p.get("entry") {
        Some(Json::Null) => None,
        Some(Json::String(s)) if s.len() <= 256 && !s.chars().any(char::is_control) => {
            if s.is_empty() {
                None
            } else {
                Some(s.clone())
            }
        }
        _ => return Err(bad("shape.setMacro", "entry must be text up to 256 bytes or null")),
    };
    let sheet = target_sheet(s, p)?;
    edit(s, |cx| {
        let sh = cx.sheet_mut(sheet)?;
        if sh.is_protected() {
            return Err(EngineError::Other(PROTECTED.into()));
        }
        let shape = sh.shapes.iter_mut().find(|shape| shape.id == id).ok_or_else(|| bad("shape.setMacro", "no such shape"))?;
        if shape.format.control_inert {
            return Err(bad("shape.setMacro", "unsupported native control is retained without editing"));
        }
        if matches!(shape.kind, ShapeKind::Ink | ShapeKind::Icon) {
            return Err(bad("shape.setMacro", "assign macros to an ordinary worksheet shape"));
        }
        shape.on_action = entry;
        Ok(Json::Null)
    })
}
