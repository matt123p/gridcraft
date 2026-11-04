use super::controls::sep;
use super::{Icon, SheetApp, Ui, big_button, json};
use crate::widgets::big_toggle_button;

pub(super) fn developer(app: &mut SheetApp, ui: &mut Ui) {
    let visual = big_button(ui, Icon::Script, "Visual\nBasic", "Visual Basic (Alt+F11)", false);
    visual.widget_info(|| egui::WidgetInfo::labeled(egui::Role::Button, ui.is_enabled(), "Visual Basic"));
    if visual.clicked() {
        app.run_or_alert("vba.open", json!({}));
    }
    let macros = big_button(ui, Icon::Script, "Macros", "Macros (Alt+F8)", false);
    macros.widget_info(|| egui::WidgetInfo::labeled(egui::Role::Button, ui.is_enabled(), "Macros (Alt+F8)"));
    if macros.clicked() {
        app.open_dialog("macros", json!({}));
    }
    sep(ui);
    let insert = big_button(ui, Icon::Shapes, "Insert", "Insert worksheet button", true);
    insert.widget_info(|| egui::WidgetInfo::labeled(egui::Role::Button, ui.is_enabled(), "Insert worksheet button"));
    egui::Popup::menu(&insert).show(|ui| {
        if ui.button("Button (Assign Macro?)").clicked() {
            app.open_dialog("macros", json!({"assign":true}));
            ui.close();
        }
    });
    let on = app.vba.as_ref().is_some_and(|e| e.design_mode);
    if big_toggle_button(ui, Icon::Shapes, "Design\nMode", "Design Mode", on).clicked() {
        app.run_or_alert("vba.designMode", json!({"on":!on}));
    }
}
