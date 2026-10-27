//! Optional shell glue; the editor/debugger itself is rendered by tiny-vba.
use crate::{SheetApp, theme};
use gridcraft_vba::{Action, Bridge, Ide, Mode, vba_ide::HostRequest};
use serde_json::{Value as Json, json};
use vba_ui_egui::{Command, EditorUi, Theme};

pub struct Editor {
    pub ide: Ide,
    pub bridge: Bridge,
    pub view: EditorUi,
    pub visible: bool,
    confirm_close: bool,
}
type Result<T> = std::result::Result<T, String>;
fn error(e: impl std::fmt::Display) -> String {
    e.to_string()
}

pub fn commands(app: &SheetApp) -> Vec<Json> {
    let mut commands = vec![
        json!({"id":"vba.open","label":"VBA Editor","params":"{source?: string, project?: bool}","enabled":app.session.active().is_some(),"ui":true}),
    ];
    for (id, label, params, action) in [
        ("vba.compile", "Compile VBA", "{}", Some(Action::Compile)),
        ("vba.run", "Run / Continue VBA", "{entry?: string}", Some(Action::Run)),
        ("vba.stepInto", "Step Into VBA", "{}", Some(Action::StepInto)),
        ("vba.stepOver", "Step Over VBA", "{}", Some(Action::StepOver)),
        ("vba.stepOut", "Step Out VBA", "{}", Some(Action::StepOut)),
        ("vba.pause", "Pause VBA", "{}", Some(Action::Pause)),
        ("vba.reset", "Reset VBA", "{}", Some(Action::Reset)),
        ("vba.status", "VBA Status", "{}", None),
        ("vba.trust", "VBA Consent", "{allow: bool}", None),
        ("vba.return", "Return to Workbook", "{}", None),
        ("vba.close", "Close VBA Project", "{discard?: bool}", None),
    ] {
        let enabled = app.vba.as_ref().is_some_and(|e| {
            action.is_none_or(|a| {
                e.ide.enabled(a)
                    && (!matches!(a, Action::Run | Action::StepInto | Action::StepOver | Action::StepOut) || e.bridge.permitted(&app.session))
            })
        });
        commands.push(json!({"id":id,"label":label,"params":params,"enabled":enabled,"ui":true}));
    }
    commands
}

pub fn before_command(app: &mut SheetApp, id: &str) -> Result<()> {
    if matches!(id, "file.close" | "window.close" | "app.quit")
        && let Some(editor) = app.vba.as_mut()
    {
        if editor.ide.modules.iter().any(|m| m.dirty()) {
            return Err("Save VBA source or explicitly discard it with vba.close before closing the workbook.".into());
        }
        if editor.ide.enabled(Action::Reset) {
            editor.bridge.dispatch(&mut app.session, &mut editor.ide, Action::Reset).map_err(error)?;
        }
    }
    Ok(())
}

/// UI-only commands are also reachable through SheetApp::run/control, with no dialogs.
pub fn run(app: &mut SheetApp, id: &str, p: &Json) -> Result<Json> {
    if id == "vba.open" {
        if let Some(editor) = app.vba.as_mut() {
            if editor.bridge.document() == app.session.doc().map_err(error)?.uid
                && p.get("source").is_none()
                && !p.get("project").and_then(Json::as_bool).unwrap_or(false)
            {
                editor.visible = true;
                return Ok(json!({"mode":editor.ide.mode().label()}));
            }
            return Err("Close the existing VBA project before opening another.".into());
        }
        app.settle_drafts()?;
        let bridge = Bridge::new(&app.session).map_err(error)?;
        let ide = if p.get("project").and_then(Json::as_bool).unwrap_or(false) {
            bridge.import_project(&app.session).map_err(error)?
        } else {
            let mut ide = Ide::new();
            let source = match p.get("source") {
                None => gridcraft_vba::DEMO,
                Some(Json::String(s)) => s,
                Some(_) => return Err("source must be VBA text".into()),
            };
            ide.add_module("Module1.bas", source).map_err(error)?;
            bridge.configure(&mut ide).map_err(error)?;
            ide
        };
        app.vba = Some(Editor { ide, bridge, view: EditorUi::default(), visible: true, confirm_close: false });
        return Ok(json!({"opened":true,"executed":false}));
    }
    let editor = app.vba.as_mut().ok_or("VBA editor is not open")?;
    match id {
        "vba.status" => {
            return Ok(
                json!({"mode":editor.ide.mode().label(),"output":editor.ide.output(),"diagnostics":editor.ide.diagnostics.iter().map(ToString::to_string).collect::<Vec<_>>()}),
            );
        }
        "vba.trust" => {
            let allow = p.get("allow").and_then(Json::as_bool).ok_or("allow must be true or false")?;
            editor.bridge.consent(allow);
        }
        "vba.return" => editor.visible = false,
        "vba.close" => {
            let discard = p.get("discard").and_then(Json::as_bool).unwrap_or(false);
            editor.bridge.close(&mut app.session, &mut editor.ide, discard).map_err(error)?;
            app.vba = None;
        }
        _ => {
            let action = match id {
                "vba.run" => Action::Run,
                "vba.compile" => Action::Compile,
                "vba.pause" => Action::Pause,
                "vba.reset" => Action::Reset,
                "vba.stepInto" => Action::StepInto,
                "vba.stepOver" => Action::StepOver,
                "vba.stepOut" => Action::StepOut,
                _ => return Err(format!("unknown VBA command {id}")),
            };
            if let Some(entry) = p.get("entry") {
                editor.ide.entry = entry.as_str().ok_or("entry must be a procedure name")?.into();
            }
            editor.bridge.dispatch(&mut app.session, &mut editor.ide, action).map_err(error)?;
        }
    }
    Ok(Json::Null)
}
pub fn tick(app: &mut SheetApp, ctx: &egui::Context) {
    let Some(editor) = app.vba.as_mut() else { return };
    if editor.visible && app.dialog.as_ref().is_some_and(|d| d.name == "macros") {
        app.dialog = None;
    }
    if let Err(e) = editor.bridge.tick(&mut app.session, &mut editor.ide) {
        editor.ide.message(e.to_string());
    }
    if matches!(editor.ide.mode(), Mode::Running | Mode::Waiting) {
        ctx.request_repaint();
    }
}
pub fn host_theme(tokens: theme::Tokens) -> Theme {
    Theme {
        dark: tokens.dark,
        chrome: tokens.window,
        editor: tokens.input_bg,
        text: tokens.text,
        muted: tokens.text_dim,
        border: tokens.separator,
        accent: tokens.accent,
        keyword: tokens.syntax.keyword,
        comment: tokens.syntax.comment,
        string: tokens.syntax.string,
        number: tokens.syntax.number,
        ..if tokens.dark { Theme::dark() } else { Theme::light() }
    }
}
pub fn show(app: &mut SheetApp, ui: &mut egui::Ui) {
    let Some(mut editor) = app.vba.take() else { return };
    let mut commands = Vec::new();
    egui::Panel::top("gridcraft_vba_host").show(ui, |ui| {
        ui.horizontal(|ui| {
            ui.label("GridCraft VBA");
            if ui.button("Return to workbook").clicked() {
                editor.visible = false;
            }
            if ui.button("Allow VBA for this editor session").clicked() {
                editor.bridge.consent(true);
            }
            if ui.button("Close project").clicked() {
                commands.push(Command::Action(Action::CloseProject));
            }
        });
    });
    commands.extend(editor.view.show(ui, &editor.ide, host_theme(theme::Tokens::get(ui.ctx()))));
    for command in commands {
        let result = match command {
            Command::Action(a) => editor.bridge.dispatch(&mut app.session, &mut editor.ide, a),
            Command::RunToCursor => {
                editor.bridge.begin(&mut app.session).and_then(|()| editor.ide.run_to_cursor(&mut editor.bridge.host(&mut app.session)))
            }
            other => vba_ui_egui::apply(&mut editor.ide, other, &mut editor.bridge.host(&mut app.session)),
        };
        if let Err(e) = result {
            editor.ide.message(e.to_string());
        }
    }
    let mut closed = false;
    for request in editor.ide.take_requests() {
        let result: Result<()> = match request {
            HostRequest::ReturnToHost => {
                editor.visible = false;
                Ok(())
            }
            HostRequest::ConfirmClose => {
                if editor.ide.modules.iter().any(|m| m.dirty()) {
                    editor.confirm_close = true;
                    if editor.ide.enabled(Action::Reset) {
                        editor.bridge.dispatch(&mut app.session, &mut editor.ide, Action::Reset).map_err(error)
                    } else {
                        Ok(())
                    }
                } else {
                    editor.bridge.close(&mut app.session, &mut editor.ide, false).map(|()| closed = true).map_err(error)
                }
            }
            HostRequest::Save { id, name, text, revision } => {
                save_source(app, &name, &text).and_then(|saved| if saved { editor.ide.mark_saved(id, revision).map_err(error) } else { Ok(()) })
            }
            HostRequest::Export { name, text } => save_source(app, &name, &text).map(|_| ()),
            HostRequest::Import => import_source(app, &mut editor.ide),
        };
        if let Err(e) = result {
            editor.ide.message(e);
        }
    }
    if editor.confirm_close {
        egui::Window::new("Unsaved VBA source").collapsible(false).resizable(false).show(ui.ctx(), |ui| {
            ui.label("Export source before closing to keep your edits.");
            ui.horizontal(|ui| {
                if ui.button("Keep editing").clicked() {
                    editor.confirm_close = false;
                }
                if ui.button("Discard source changes").clicked() {
                    match editor.bridge.close(&mut app.session, &mut editor.ide, true) {
                        Ok(()) => closed = true,
                        Err(e) => editor.ide.message(e.to_string()),
                    }
                }
            });
        });
    }
    if !closed {
        app.vba = Some(editor);
    }
}
fn save_source(app: &SheetApp, name: &str, text: &str) -> Result<bool> {
    if let Some(pick) = &app.services.pick_save {
        if let Some(path) = pick(name) {
            gridcraft_engine::io::write_file(&path, text.as_bytes()).map_err(error)?;
            return Ok(true);
        }
        return Ok(false);
    }
    if let Some(download) = &app.services.download {
        download(name, text.as_bytes());
        return Ok(true);
    }
    Err("No source export service is available.".into())
}
fn import_source(app: &SheetApp, ide: &mut Ide) -> Result<()> {
    let Some(pick) = &app.services.pick_open else { return Err("Use vba.open with source text to import VBA in this host.".into()) };
    let Some(path) = pick() else { return Ok(()) };
    let bytes = read_source(&path)?;
    let text = String::from_utf8(bytes).map_err(error)?;
    let name = std::path::Path::new(&path).file_name().and_then(|n| n.to_str()).ok_or("invalid module name")?;
    ide.add_module(name, &text).map_err(error)?;
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn read_source(path: &str) -> Result<Vec<u8>> {
    use std::io::Read;
    let file = std::fs::File::open(path).map_err(error)?;
    if !file.metadata().map_err(error)?.is_file() {
        return Err("Only regular source files can be imported.".into());
    }
    let limit = gridcraft_vba::vba_ide::MAX_PROJECT_BYTES;
    let mut bytes = Vec::new();
    file.take(limit as u64 + 1).read_to_end(&mut bytes).map_err(error)?;
    if bytes.len() > limit {
        return Err("VBA source exceeds project budget".into());
    }
    Ok(bytes)
}
#[cfg(target_arch = "wasm32")]
fn read_source(_path: &str) -> Result<Vec<u8>> {
    Err("Use vba.open with source text in the browser.".into())
}
