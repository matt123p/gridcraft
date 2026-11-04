//! Original native worksheet controls, rendered with the production widget path.
use gridcraft_engine::Session;
use gridcraft_ui_egui::{Services, SheetApp, theme};
use serde_json::json;
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let folder = std::env::args().nth(1).unwrap_or_else(|| "target/worksheet-controls".into());
    std::fs::create_dir_all(&folder)?;
    for (name, dark, width, scale) in [("light", false, 1200., 1.), ("dark", true, 1200., 1.), ("scaled", false, 900., 1.5)] {
        let mut session = Session::new();
        session.new_workbook();
        let mut app = SheetApp::new(session, Services::default());
        app.use_prepared_fonts();
        app.ui.dark = dark;
        app.ui.system_theme = false;
        app.run("range.setValues", json!({"range":"A1:A3","values":[["Alpha"],["Beta"],["Gamma"]]}))?;
        app.run("cell.set", json!({"cell":"B1","input":"Native worksheet Form Controls"}))?;
        for (i, (kind, caption)) in [
            (4, "Options"),
            (0, "Run report"),
            (1, "Include totals"),
            (7, "Detailed report"),
            (5, "Choose a category"),
            (2, "Category"),
            (6, "Categories"),
            (8, "Scroll value"),
            (9, "Spin value"),
        ]
        .into_iter()
        .enumerate()
        {
            let id=app.run("control.add",json!({"type":kind,"name":format!("Control{i}"),"caption":caption,"left":90+(i%3)*240,"top":60+(i/3)*140,"width":if kind==9{28}else{190},"height":if [4,6].contains(&kind){105}else if [8,9].contains(&kind){60}else{28}}))?["shape"].clone();
            if [2, 6].contains(&kind) {
                app.run("control.set", json!({"id":id,"listFillRange":"$A$1:$A$3","value":2}))?;
            }
            if kind == 1 || kind == 7 {
                app.run("control.set", json!({"id":id,"value":1}))?;
            }
            if kind == 8 {
                app.run("control.set", json!({"id":id,"horizontal":true,"value":35}))?;
            }
        }
        let mut h = egui_kittest::Harness::builder()
            .with_size(egui::vec2(width, 850.))
            .with_pixels_per_point(scale)
            .with_context_setup(move |ctx| {
                ctx.set_fonts(theme::pinned_font_definitions());
                theme::apply(ctx, dark);
            })
            .wgpu()
            .build_ui_state(
                |ui, app: &mut SheetApp| {
                    app.logic(ui.ctx());
                    app.ui(ui);
                },
                app,
            );
        h.run_steps(5);
        h.render()?.save(format!("{folder}/{name}.png"))?;
    }
    Ok(())
}
