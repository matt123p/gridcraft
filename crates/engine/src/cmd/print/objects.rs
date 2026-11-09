use super::draw::{Block, PdfMeasure, Xf, draw_prims, rgb_of};
use super::text::wrap_lines;
use super::{Anchor, Font, Page, RangeRef, Rect, ShapeKind, Sheet, Value, Workbook};

/// Cell pictures fit the current cell (or merged-cell) rectangle without changing aspect ratio.
pub(super) fn draw_cell_picture(pg: &mut Page, data: &[u8], rect: (f32, f32, f32, f32), k: f32) {
    let (x, y, w, h) = rect;
    let pad = 2.0 * k;
    let (w, h) = ((w - 2.0 * pad).max(0.0), (h - 2.0 * pad).max(0.0));
    if w <= 0.0 || h <= 0.0 || data.len() > 16 * 1024 * 1024 {
        return;
    }
    let Some((iw, ih)) = crate::io::image_size(data) else { return };
    if iw == 0 || ih == 0 || iw > 8192 || ih > 8192 || u64::from(iw) * u64::from(ih) > 16_000_000 {
        return;
    }
    let Ok(mut reader) = image::ImageReader::new(std::io::Cursor::new(data)).with_guessed_format() else { return };
    let mut limits = image::Limits::default();
    limits.max_image_width = Some(8192);
    limits.max_image_height = Some(8192);
    limits.max_alloc = Some(64 * 1024 * 1024);
    reader.limits(limits);
    let Ok(image) = reader.decode() else { return };
    let image =
        image.thumbnail(((w / 72.0) * 150.0).ceil().clamp(1.0, 4000.0) as u32, ((h / 72.0) * 150.0).ceil().clamp(1.0, 4000.0) as u32).to_rgba8();
    let fit = (w / iw as f32).min(h / ih as f32);
    let (pw, ph) = (iw as f32 * fit, ih as f32 * fit);
    pg.image_rgba(Rect::new(x + pad + (w - pw) / 2.0, y + pad + (h - ph) / 2.0, pw, ph), image.as_raw(), [image.width(), image.height()]);
}

/// Sheet pixel rectangle of an anchor.
/// Most source lines printed for an equation that has no layout.
pub(super) const MAX_FALLBACK_LINES: usize = 500;

/// The lines of `text` printed in place of an equation that has no layout: those of its first
/// `MAX_SOURCE_BYTES` (a longer source never lays out, and a file can carry megabytes) that fit a
/// frame `h` points tall at `line_h` points per line (at least the first), and never more than
/// [`MAX_FALLBACK_LINES`]. A single `pg.text` call would print every newline as `?`.
pub(super) fn fallback_lines(text: &str, h: f32, line_h: f32) -> impl Iterator<Item = &str> {
    let text = crate::cmd::equation::source_excerpt(text);
    let fit = h / line_h;
    let n = if fit >= MAX_FALLBACK_LINES as f32 {
        MAX_FALLBACK_LINES
    } else if fit >= 1.0 {
        fit as usize
    } else {
        1
    };
    text.lines().take(n)
}

fn anchor_rect(sh: &Sheet, a: &Anchor) -> (f64, f64, f64, f64) {
    (sh.col_left(a.cell.col) + a.dx as f64, sh.row_top(a.cell.row) + a.dy as f64, a.width.max(0.0) as f64, a.height.max(0.0) as f64)
}

/// Decodes an embedded picture to RGBA, downscaled to at most `max_w`×`max_h` pixels.
fn decode_image(data: &[u8], max_w: u32, max_h: u32) -> Option<(Vec<u8>, u32, u32)> {
    let (w, h) = crate::io::image_size(data)?;
    if w == 0 || h == 0 || w as u64 * h as u64 > 60_000_000 {
        return None;
    }
    let img = image::load_from_memory(data).ok()?;
    let img = if img.width() > max_w.max(1) || img.height() > max_h.max(1) { img.thumbnail(max_w.max(1), max_h.max(1)) } else { img };
    let rgba = img.to_rgba8();
    let (pw, ph) = (rgba.width(), rgba.height());
    Some((rgba.into_raw(), pw, ph))
}

/// Draws charts, pictures, shapes and sparklines overlapping the body block of a page.
/// `pics` memoizes decoded pictures for the whole render (one page at a time draws each
/// picture again, but the decode happens once).
pub(super) fn draw_objects(pg: &mut Page, wb: &Workbook, si: usize, sh: &Sheet, b: &Block, pics: &mut super::PicCache) {
    let block = b.range();
    let (sx0, sy0) = (sh.col_left(block.start.col), sh.row_top(block.start.row));
    let (sx1, sy1) = (sh.col_left(block.end.col.saturating_add(1)), sh.row_top(block.end.row.saturating_add(1)));
    let (bx, by) = (b.xs.first().copied().unwrap_or(0.0), b.ys.first().copied().unwrap_or(0.0));
    let (bw, bh) = (b.xs.last().copied().unwrap_or(bx) - bx, b.ys.last().copied().unwrap_or(by) - by);
    let visible = |r: (f64, f64, f64, f64)| r.2 > 0.0 && r.3 > 0.0 && r.0 < sx1 && r.0 + r.2 > sx0 && r.1 < sy1 && r.1 + r.3 > sy0;
    let place = |r: (f64, f64, f64, f64)| (bx + ((r.0 - sx0) as f32) * b.k, by + ((r.1 - sy0) as f32) * b.k, r.2 as f32 * b.k, r.3 as f32 * b.k);
    pg.save();
    pg.clip_rect(bx, by, bw, bh);
    draw_sparklines(pg, wb, si, sh, b, block);
    for (imi, im) in sh.images.iter().enumerate() {
        let r = anchor_rect(sh, &im.anchor);
        if !visible(r) {
            continue;
        }
        let (x, y, w, h) = place(r);
        // About 150 dpi is plenty for print.
        let (mw, mh) = (((w / 72.0) * 150.0).ceil().clamp(1.0, 4000.0) as u32, ((h / 72.0) * 150.0).ceil().clamp(1.0, 4000.0) as u32);
        let decoded = pics.entry((si, imi)).or_insert_with(|| decode_image(&im.data, mw, mh));
        if let Some((px, pw, ph)) = decoded.as_ref() {
            pg.image_rgba(Rect::new(x, y, w, h), px, [*pw, *ph]);
        }
    }
    for shp in &sh.shapes {
        let r = anchor_rect(sh, &shp.anchor);
        if !visible(r) {
            continue;
        }
        let (x, y, w, h) = place(r);
        if shp.kind == ShapeKind::Equation {
            let ink = rgb_of(shp.line, wb).unwrap_or([0, 0, 0]);
            match shp.math.get(&shp.text) {
                Some(Ok(list)) => crate::cmd::equation::draw_pdf(pg, list, (x, y, w, h), crate::cmd::equation::PADDING * b.k, ink),
                _ => {
                    // No layout (the source does not parse, or the file was past the layout
                    // budget): the source, one text run per line, as many as the frame holds.
                    let line_h = 12.0 * b.scale;
                    // A long line stays inside the frame; this clip nests in the block's.
                    pg.save();
                    pg.clip_rect(x, y, w, h);
                    for (i, line) in fallback_lines(&shp.text, h, line_h).enumerate() {
                        pg.text(x + 2.0 * b.k, y + line_h * (i + 1) as f32, 10.0 * b.scale, Font::Courier, ink, line);
                    }
                    pg.restore();
                }
            }
            continue;
        }
        if shp.format.form_control.is_some() {
            if let Ok(c) = gridcraft_model::form_controls::effective_control(wb, si, shp)
                && c.visible
                && c.print_object
            {
                draw_control(pg, shp, &c, b, (x, y, w, h));
            }
        } else {
            draw_shape(pg, wb, shp, b, (x, y, w, h));
        }
    }
    for ch in &sh.charts {
        let r = anchor_rect(sh, &ch.anchor);
        if !visible(r) {
            continue;
        }
        let (x, y, _, _) = place(r);
        let data = gridcraft_chart::resolve(wb, si, ch);
        let prims = gridcraft_chart::render(ch, &data, r.2 as f32, r.3 as f32, &PdfMeasure);
        draw_prims(pg, &prims, Xf { ox: x, oy: y, k: b.k });
    }
    pg.restore();
}

fn draw_sparklines(pg: &mut Page, wb: &Workbook, si: usize, sh: &Sheet, b: &Block, block: RangeRef) {
    for sp in &sh.sparklines {
        if !block.contains(sp.cell) {
            continue;
        }
        let vals: Vec<Option<f64>> = match gridcraft_calc::evaluate(wb, si, sp.cell, &sp.source) {
            Value::Array(a) => a.iter().take(10_000).map(Value::as_f64).collect(),
            v => vec![v.as_f64()],
        };
        let color = rgb_of(sp.color, wb).unwrap_or([0x15, 0x60, 0x82]);
        let (x, y, w, h) = b.rect(RangeRef::cell(sp.cell));
        let (wpx, hpx) = (w / b.k, h / b.k);
        let prims = gridcraft_chart::render_sparkline(sp.kind, &vals, [color[0], color[1], color[2], 255], sp.markers, wpx, hpx);
        draw_prims(pg, &prims, Xf { ox: x, oy: y, k: b.k });
    }
}

fn draw_shape(pg: &mut Page, wb: &Workbook, shp: &gridcraft_model::Shape, b: &Block, (x, y, w, h): (f32, f32, f32, f32)) {
    let fill = rgb_of(shp.fill, wb);
    let line = rgb_of(shp.line, wb).filter(|_| !shp.format.no_line).map(|c| (c, (0.75 * b.k).max(0.25)));
    match shp.kind {
        ShapeKind::Ellipse => {
            let pts: Vec<[f32; 2]> = (0..72)
                .map(|i| {
                    let a = i as f32 / 72.0 * std::f32::consts::TAU;
                    [x + w / 2.0 + w / 2.0 * a.cos(), y + h / 2.0 + h / 2.0 * a.sin()]
                })
                .collect();
            pg.polygon(&pts, fill, line);
        }
        ShapeKind::Triangle => pg.polygon(&[[x + w / 2.0, y], [x + w, y + h], [x, y + h]], fill, line),
        ShapeKind::Line | ShapeKind::Arrow => {
            let c = line.map(|l| l.0).or(fill).unwrap_or([0, 0, 0]);
            pg.stroke_line(x, y, x + w, y + h, (1.0 * b.k).max(0.3), c);
            if shp.kind == ShapeKind::Arrow {
                let ang = h.atan2(w);
                let al = 8.0 * b.k;
                let p = |da: f32| [x + w - al * (ang + da).cos(), y + h - al * (ang + da).sin()];
                pg.polygon(&[[x + w, y + h], p(0.4), p(-0.4)], Some(c), None);
            }
        }
        _ => {
            pg.polygon(&[[x, y], [x + w, y], [x + w, y + h], [x, y + h]], fill, line);
        }
    }
    if shp.text.is_empty() {
        return;
    }
    use gridcraft_model::{ShapeTextAlign, ShapeTextAnchor};
    let fm = &shp.format;
    let text_box = shp.kind == ShapeKind::TextBox;
    let size = fm.font_size.unwrap_or(11.0).clamp(1.0, 4000.0) * b.scale;
    let font = Font::styled(Font::family_for(fm.font.as_deref().unwrap_or("")), fm.bold, fm.italic);
    let color = fm.text_color.and_then(|c| rgb_of(c, wb)).unwrap_or([0, 0, 0]);
    let align = fm.align.unwrap_or(if text_box { ShapeTextAlign::Left } else { ShapeTextAlign::Center });
    let anchor = fm.anchor.unwrap_or(if text_box { ShapeTextAnchor::Top } else { ShapeTextAnchor::Middle });
    let lines = wrap_lines(&shp.text, font, size, (w - 4.0 * b.k).max(size));
    let block_h = lines.len().min(500) as f32 * size * 1.2;
    let mut yy = match anchor {
        ShapeTextAnchor::Top => y + 2.0 * b.k,
        ShapeTextAnchor::Middle => y + (h - block_h) / 2.0,
        ShapeTextAnchor::Bottom => y + h - 2.0 * b.k - block_h,
    } + size * 0.8;
    for l in lines.iter().take(500) {
        let lw = gridcraft_pdf::text_width(font, size, l);
        let lx = match align {
            ShapeTextAlign::Left => x + 2.0 * b.k,
            ShapeTextAlign::Center => x + (w - lw) / 2.0,
            ShapeTextAlign::Right => x + w - 2.0 * b.k - lw,
        };
        pg.text(lx, yy, size, font, color, l);
        yy += size * 1.2;
    }
}

/// Print native control state, independent of the interactive egui surface.
fn draw_control(
    pg: &mut Page,
    shape: &gridcraft_model::Shape,
    c: &gridcraft_model::form_controls::FormControl,
    b: &Block,
    (x, y, w, h): (f32, f32, f32, f32),
) {
    use gridcraft_model::FormControlKind as K;
    let k = b.k;
    let size = 11. * b.scale;
    let ink = if c.enabled { [0, 0, 0] } else { [128, 128, 128] };
    let font = Font::styled(Font::family_for(shape.format.font.as_deref().unwrap_or("")), shape.format.bold, shape.format.italic);
    let rect = |pg: &mut Page, x: f32, y: f32, w: f32, h: f32, fill: Option<[u8; 3]>| {
        pg.polygon(&[[x, y], [x + w, y], [x + w, y + h], [x, y + h]], fill, Some((ink, 0.5)))
    };
    pg.save();
    pg.clip_rect(x, y, w, h);
    match c.kind {
        K::Label => pg.text(x, y + size, size, font, ink, &shape.text),
        K::CheckBox | K::OptionButton => {
            let d = (12. * k).min(h);
            let cy = y + (h - d) / 2.;
            if c.kind == K::CheckBox {
                rect(pg, x, cy, d, d, None);
            } else {
                let pts = (0..24)
                    .map(|i| {
                        let a = i as f32 * std::f32::consts::TAU / 24.;
                        [x + d / 2. + d / 2. * a.cos(), cy + d / 2. + d / 2. * a.sin()]
                    })
                    .collect::<Vec<_>>();
                pg.polygon(&pts, None, Some((ink, 0.5)));
            }
            if c.value == 1 {
                pg.stroke_line(x + 2. * k, cy + 2. * k, x + d - 2. * k, cy + d - 2. * k, 1., ink);
                pg.stroke_line(x + d - 2. * k, cy + 2. * k, x + 2. * k, cy + d - 2. * k, 1., ink);
            } else if c.value == 2 {
                rect(pg, x + 2. * k, cy + 2. * k, (d - 4. * k).max(0.), (d - 4. * k).max(0.), Some(ink));
            }
            pg.text(x + d + 4. * k, y + h / 2. + size * 0.3, size, font, ink, &shape.text);
        }
        K::ListBox | K::DropDown => {
            rect(pg, x, y, w, h, Some([255, 255, 255]));
            if c.kind == K::DropDown {
                if let Some(item) = c.value.checked_sub(1).and_then(|n| usize::try_from(n).ok()).and_then(|n| c.items.get(n)) {
                    pg.text(x + 2. * k, y + size, size, font, ink, item);
                }
                pg.polygon(&[[x + w - 12. * k, y + h / 3.], [x + w - 4. * k, y + h / 3.], [x + w - 8. * k, y + h * 2. / 3.]], Some(ink), None);
            } else {
                for (i, item) in c.items.iter().take(1024).enumerate() {
                    let yy = y + i as f32 * size * 1.2;
                    if yy >= y + h {
                        break;
                    }
                    let selected = if c.multi_select { c.selected.contains(&((i + 1) as u32)) } else { c.value == (i + 1) as i32 };
                    if selected {
                        rect(pg, x, yy, w, size * 1.2, Some([220, 230, 245]));
                    }
                    pg.text(x + 2. * k, yy + size, size, font, ink, item);
                }
            }
        }
        K::ScrollBar | K::Spinner => {
            rect(pg, x, y, w, h, Some([240, 240, 240]));
            let fraction = if c.max > c.min { (c.value - c.min) as f32 / (c.max - c.min) as f32 } else { 0. };
            if c.kind == K::ScrollBar {
                if c.horizontal {
                    rect(pg, x + fraction * (w - 8. * k).max(0.), y, 8. * k, h, Some([150, 150, 150]));
                } else {
                    rect(pg, x, y + fraction * (h - 8. * k).max(0.), w, 8. * k, Some([150, 150, 150]));
                }
            } else {
                pg.stroke_line(x, y + h / 2., x + w, y + h / 2., 0.5, ink);
                pg.polygon(&[[x + w / 4., y + h * 0.35], [x + w * 0.75, y + h * 0.35], [x + w / 2., y + h * 0.15]], Some(ink), None);
                pg.polygon(&[[x + w / 4., y + h * 0.65], [x + w * 0.75, y + h * 0.65], [x + w / 2., y + h * 0.85]], Some(ink), None);
            }
        }
        K::Button | K::GroupBox => {
            rect(pg, x, y, w, h, if c.kind == K::Button { Some([240, 240, 240]) } else { None });
            pg.text(x + 4. * k, if c.kind == K::Button { y + h / 2. + size * 0.3 } else { y + size }, size, font, ink, &shape.text);
        }
    }
    pg.restore();
}
