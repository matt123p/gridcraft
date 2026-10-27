//! Original actual-paused-state GridCraft/tiny-vba integration screenshots.
use gridcraft_engine::Session;
use gridcraft_ui_egui::{Services, SheetApp, theme};
use serde_json::json;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let folder = std::env::args().nth(1).unwrap_or_else(|| "target/vba-ui".into());
    std::fs::create_dir_all(&folder)?;
    for (name, dark, width, scale) in [("light", false, 1440., 1.), ("dark", true, 1440., 1.), ("narrow", false, 700., 1.5)] {
        let mut session = Session::new();
        session.new_workbook();
        let mut app = SheetApp::new(session, Services::default());
        app.use_prepared_fonts();
        app.ui.dark = dark;
        app.ui.system_theme = false;
        app.run("vba.open", json!({})).map_err(std::io::Error::other)?;
        if let Some(editor) = app.vba.as_mut() {
            let byte = gridcraft_vba::DEMO.find("Range(\"B1\")").ok_or("breakpoint not found")?;
            editor.ide.toggle_breakpoint_at(byte)?;
        }
        app.run("vba.run", json!({})).map_err(std::io::Error::other)?;
        let mut h = egui_kittest::Harness::builder()
            .with_size(egui::vec2(width, 900.))
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
        h.run_steps(10);
        if h.state().vba.as_ref().is_none_or(|e| e.ide.mode() != gridcraft_vba::Mode::Break) {
            return Err("expected real debugger break".into());
        }
        h.render()?.save(format!("{folder}/{name}.png"))?;
    }
    Ok(())
}
