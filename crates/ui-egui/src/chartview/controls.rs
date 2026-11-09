//! Native worksheet controls use ordinary egui input/focus and engine commands.
use super::{Geo, Rect, Sheet, SheetApp, Stroke, Workbook, anchor_rect, objects_hidden, vec2};
use gridcraft_engine::model::{FormControlKind, form_controls::effective_control};
#[cfg(not(feature = "vba"))]
use serde_json::Value as Json;
use serde_json::json;

pub(crate) fn show(app: &mut SheetApp, ui: &mut egui::Ui, geo: &Geo, wb: &Workbook, si: usize, sh: &Sheet) -> bool {
    if objects_hidden(wb) {
        return false;
    }
    #[cfg(feature = "vba")]
    let design = app.vba.as_ref().is_some_and(|e| e.design_mode);
    #[cfg(not(feature = "vba"))]
    let design = false;
    let mut actions = vec![];
    let mut consumed = false;
    let doc = app.session.active().map_or(0, |d| d.uid);
    for shape in &sh.shapes {
        if shape.format.form_control.is_none() {
            continue;
        }
        let rect = geo.mirror_rect(anchor_rect(geo, sh, &shape.anchor));
        if !geo.cells.intersects(rect) || rect.width() <= 0. || rect.height() <= 0. {
            continue;
        }
        let (c, binding_error) = match effective_control(wb, si, shape) {
            Ok(c) => (c, None),
            Err(error) => {
                let Some(c) = shape.format.form_control.clone() else { continue };
                (c, Some(error))
            }
        };
        if !c.visible {
            continue;
        }
        let id = shape.id;
        ui.push_id(("worksheet-control", doc, si, id), |ui| {
            // Controls are document objects: dark application chrome must not turn
            // their captions white against the worksheet's white page.
            *ui.visuals_mut() = egui::Visuals::light();
            ui.visuals_mut().override_text_color =
                Some(shape.format.text_color.and_then(|c| c.resolve(&wb.theme)).map(crate::theme::color32).unwrap_or(egui::Color32::BLACK));
            ui.style_mut().override_font_id = Some(crate::theme::ui_font((shape.format.font_size.unwrap_or(11.) * 4. / 3. * geo.z).clamp(1., 2048.)));
            let previous = ui.clip_rect();
            ui.set_clip_rect(previous.intersect(geo.cells));
            ui.add_enabled_ui(c.enabled && binding_error.is_none() && !design && !sh.is_protected(), |ui| {
                let mut action = None;
                match c.kind {
                    FormControlKind::Button => {
                        let r = ui.put(rect, egui::Button::new(&shape.text));
                        if r.clicked() {
                            r.request_focus();
                            action = Some(json!({}));
                        }
                    }
                    FormControlKind::CheckBox => {
                        let mut value = c.value == 1;
                        let r = ui.put(rect, egui::Checkbox::new(&mut value, &shape.text).indeterminate(c.value == 2));
                        if r.changed() {
                            r.request_focus();
                            action = Some(json!({}));
                        }
                    }
                    FormControlKind::OptionButton => {
                        let r = ui.put(rect, egui::RadioButton::new(c.value == 1, &shape.text));
                        if r.clicked() {
                            r.request_focus();
                            action = Some(json!({}));
                        }
                    }
                    FormControlKind::Label => {
                        ui.put(rect, egui::Label::new(&shape.text));
                    }
                    FormControlKind::GroupBox => {
                        ui.painter().rect_stroke(
                            rect,
                            2.,
                            Stroke::new(1., ui.visuals().widgets.noninteractive.bg_stroke.color),
                            egui::StrokeKind::Inside,
                        );
                        ui.put(Rect::from_min_size(rect.min, vec2(rect.width(), 20. * geo.z)), egui::Label::new(&shape.text));
                    }
                    FormControlKind::ScrollBar => {
                        ui.spacing_mut().slider_width = if c.horizontal { rect.width() } else { rect.height() };
                        let mut value = c.value;
                        let mut slider = egui::Slider::new(&mut value, c.min..=c.max).step_by(f64::from(c.small_change.max(1))).show_value(false);
                        if !c.horizontal {
                            slider = slider.vertical();
                        }
                        let r = ui.put(rect, slider);
                        r.widget_info(|| egui::WidgetInfo::labeled(egui::Role::Slider, ui.is_enabled(), &shape.text));
                        if r.changed() {
                            action = Some(json!({"value":value}));
                        }
                        if r.has_focus() {
                            let keys = ui.input(|i| {
                                (
                                    i.key_pressed(egui::Key::PageUp),
                                    i.key_pressed(egui::Key::PageDown),
                                    i.key_pressed(egui::Key::Home),
                                    i.key_pressed(egui::Key::End),
                                    i.key_pressed(egui::Key::ArrowLeft) || i.key_pressed(egui::Key::ArrowUp),
                                    i.key_pressed(egui::Key::ArrowRight) || i.key_pressed(egui::Key::ArrowDown),
                                )
                            });
                            if keys.0 {
                                action = Some(json!({"delta":-c.large_change}));
                            } else if keys.1 {
                                action = Some(json!({"delta":c.large_change}));
                            } else if keys.2 {
                                action = Some(json!({"value":c.min}));
                            } else if keys.3 {
                                action = Some(json!({"value":c.max}));
                            } else if keys.4 {
                                action = Some(json!({"delta":-c.small_change}));
                            } else if keys.5 {
                                action = Some(json!({"delta":c.small_change}));
                            }
                        }
                    }
                    FormControlKind::Spinner => {
                        ui.scope_builder(egui::UiBuilder::new().max_rect(rect).layout(egui::Layout::top_down(egui::Align::Center)), |ui| {
                            if ui
                                .add_sized(vec2(rect.width(), rect.height() / 2.), egui::Button::new("▲"))
                                .on_hover_text(format!("Increase {}", shape.text))
                                .clicked()
                            {
                                action = Some(json!({"delta":c.small_change}));
                            }
                            if ui
                                .add_sized(vec2(rect.width(), rect.height() / 2.), egui::Button::new("▼"))
                                .on_hover_text(format!("Decrease {}", shape.text))
                                .clicked()
                            {
                                action = Some(json!({"delta":-c.small_change}));
                            }
                        });
                    }
                    FormControlKind::DropDown => {
                        ui.scope_builder(egui::UiBuilder::new().max_rect(rect), |ui| {
                            let selected =
                                usize::try_from(c.value.saturating_sub(1)).ok().and_then(|n| c.items.get(n)).map(String::as_str).unwrap_or("");
                            let response = craft_ui::ComboBox::from_id_salt("list")
                                .width(rect.width())
                                .height((c.drop_lines.min(100) as f32 * 20.).max(20.))
                                .selected_text(selected)
                                .show_ui(ui, |ui| {
                                    for (i, text) in c.items.iter().enumerate() {
                                        if ui.selectable_label(c.value == (i + 1) as i32, text).clicked() {
                                            action = Some(json!({"item":i+1}));
                                        }
                                    }
                                });
                            response.response.widget_info(|| egui::WidgetInfo::labeled(egui::Role::ComboBox, ui.is_enabled(), &shape.text));
                        });
                    }
                    FormControlKind::ListBox => {
                        ui.painter().rect_filled(rect, 0., ui.visuals().extreme_bg_color);
                        ui.painter().rect_stroke(rect, 0., ui.visuals().widgets.noninteractive.bg_stroke, egui::StrokeKind::Inside);
                        ui.scope_builder(egui::UiBuilder::new().max_rect(rect), |ui| {
                            egui::ScrollArea::vertical().max_height(rect.height()).show_rows(ui, 20. * geo.z, c.items.len(), |ui, rows| {
                                for i in rows {
                                    if let Some(text) = c.items.get(i) {
                                        let selected =
                                            if c.multi_select { c.selected.contains(&((i + 1) as u32)) } else { c.value == (i + 1) as i32 };
                                        if ui.selectable_label(selected, text).clicked() {
                                            if c.extended_select {
                                                let mods = ui.input(|i| i.modifiers);
                                                let anchor_id = ui.scope_id().with("selection-anchor");
                                                let mut selected = if mods.ctrl || mods.command { c.selected.clone() } else { vec![] };
                                                if mods.shift {
                                                    let anchor = ui.ctx().data_mut(|d| d.get_temp::<usize>(anchor_id)).unwrap_or(i);
                                                    selected.extend((anchor.min(i)..=anchor.max(i)).map(|n| (n + 1) as u32));
                                                    selected.sort_unstable();
                                                    selected.dedup();
                                                } else {
                                                    let item = (i + 1) as u32;
                                                    if selected.contains(&item) {
                                                        selected.retain(|n| *n != item)
                                                    } else {
                                                        selected.push(item);
                                                    }
                                                    ui.ctx().data_mut(|d| d.insert_temp(anchor_id, i));
                                                }
                                                action = Some(json!({"selected":selected}));
                                            } else {
                                                action = Some(json!({"item":i+1}));
                                            }
                                        }
                                    }
                                }
                            });
                        });
                    }
                }
                if let Some(mut action) = action {
                    action["id"] = json!(id);
                    action["sheet"] = json!(si);
                    actions.push(action);
                    consumed = true;
                }
            });
            if design {
                let r = ui.interact(rect, ui.scope_id().with("design-select"), egui::Sense::click_and_drag());
                if r.clicked() || r.drag_started() {
                    app.selected_object = Some(gridcraft_engine::model::ObjRef::new(gridcraft_engine::model::ObjKind::Shape, id));
                }
                if r.drag_started()
                    && let Some(pos) = ui.input(|i| i.pointer.press_origin())
                {
                    let start = geo.mirror_pos(pos);
                    let logical = anchor_rect(geo, sh, &shape.anchor);
                    let a = shape.anchor;
                    app.grid.drag = crate::grid::Drag::Object {
                        obj: gridcraft_engine::model::ObjRef::new(gridcraft_engine::model::ObjKind::Shape, id),
                        start,
                        orig: (a.cell, a.dx, a.dy, a.width, a.height),
                        handle: crate::objects::handle_at(logical, start),
                        stepped: false,
                    };
                }
                if app.selected_object == Some(gridcraft_engine::model::ObjRef::new(gridcraft_engine::model::ObjKind::Shape, id)) {
                    super::selection_frame(ui.painter(), rect);
                }
            }
            if let Some(error) = binding_error {
                ui.interact(rect, ui.scope_id().with("binding-error"), egui::Sense::hover()).on_hover_text(error);
            }
            ui.set_clip_rect(previous);
        });
    }
    for action in actions {
        #[cfg(feature = "vba")]
        {
            app.run_or_alert("vba.control", action);
        }
        #[cfg(not(feature = "vba"))]
        {
            let assigned = action
                .get("id")
                .and_then(Json::as_u64)
                .and_then(|id| sh.shapes.iter().find(|s| u64::from(s.id) == id))
                .is_some_and(|s| s.on_action.is_some());
            if assigned {
                app.run_or_alert("vba.control", action);
            } else {
                app.run_or_alert("control.activate", action);
            }
        }
    }
    consumed
}
