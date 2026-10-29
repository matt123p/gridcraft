//! Optional shell glue; the editor/debugger itself is rendered by tiny-vba.
use crate::{SheetApp, theme};
use gridcraft_vba::worksheet::{Functions, Project};
use gridcraft_vba::{Action, Bridge, Ide, Mode, vba_ide::HostRequest};
use serde_json::{Value as Json, json};
use std::sync::{Arc, atomic::Ordering};
use vba_ui_egui::{Command, EditorUi, Theme};

pub struct Editor {
    pub ide: Ide,
    pub bridge: Bridge,
    pub view: EditorUi,
    pub visible: bool,
    confirm_close: bool,
    pub design_mode: bool,
    functions: Option<Arc<Functions>>,
    pub separate_window: bool,
    focus_window: bool,
    #[cfg(all(feature = "vba-window", not(target_arch = "wasm32")))]
    owner_viewport: egui::ViewportId,
}
impl Drop for Editor {
    fn drop(&mut self) {
        if let Some(f) = &self.functions {
            f.allowed.store(false, Ordering::Release);
        }
    }
}
type Result<T> = std::result::Result<T, String>;
fn error(e: impl std::fmt::Display) -> String {
    e.to_string()
}
fn placement(app: &SheetApp, p: &Json) -> Result<bool> {
    let available = cfg!(all(feature = "vba-window", not(target_arch = "wasm32"))) && app.native_windows;
    match p.get("separateWindow") {
        None => Ok(available),
        Some(Json::Bool(false)) => Ok(false),
        Some(Json::Bool(true)) if available => Ok(true),
        Some(Json::Bool(true)) => Err("A separate VBA window requires the desktop vba-window feature and a native window host.".into()),
        Some(_) => Err("separateWindow must be true or false".into()),
    }
}
fn navigate(editor: &mut Editor, entry: &str) -> Result<()> {
    for (index, module) in editor.ide.modules.iter_mut().enumerate() {
        if let Some(p) = module.procedures.iter().find(|p| p.name.eq_ignore_ascii_case(entry)) {
            editor.ide.active = index;
            module.caret = p.start;
            editor.ide.entry = p.name.clone();
            return Ok(());
        }
    }
    Err("No such VBA procedure".into())
}

/// One immediate native viewport, rendered once by the desktop shell after workbook windows.
/// Execution still advances only in the parent SheetApp::logic; no second VM or host is created.
#[cfg(all(feature = "vba-window", not(target_arch = "wasm32")))]
pub fn show_desktop(app: &mut SheetApp, ctx: &egui::Context, owner: egui::ViewportId) {
    let Some(editor) = app.vba.as_mut().filter(|e| e.visible && e.separate_window) else { return };
    let id = egui::ViewportId::from_hash_of(("gridcraft-vba-editor", editor.bridge.document()));
    let focus = editor.focus_window;
    if focus {
        editor.owner_viewport = owner;
        editor.focus_window = false;
    }
    let return_to = editor.owner_viewport;
    let title = format!(
        "{} — Visual Basic",
        app.session.documents().iter().find(|d| d.uid == editor.bridge.document()).map(|d| d.title.as_str()).unwrap_or("GridCraft")
    );
    ctx.show_viewport_immediate(
        id,
        egui::ViewportBuilder::default().with_title(title).with_inner_size([1200., 800.]).with_min_inner_size([640., 420.]).with_decorations(true),
        |ui, _class| {
            let ctx = ui.ctx().clone();
            if ctx.input(|i| i.viewport().close_requested()) {
                ctx.send_viewport_cmd(egui::ViewportCommand::CancelClose);
                // Closing VBE hides its window; the source/project remains owned by the workbook.
                app.run_or_alert("vba.return", json!({}));
            } else {
                if focus {
                    ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
                }
                shortcuts(app, &ctx);
                if app.vba.as_ref().is_some_and(|e| e.visible) {
                    show(app, ui);
                }
            }
            if app.vba.as_ref().is_none_or(|e| !e.visible) {
                ctx.send_viewport_cmd_to(return_to, egui::ViewportCommand::Focus);
            }
            if app.vba.as_ref().is_some_and(|e| matches!(e.ide.mode(), Mode::Running | Mode::Waiting)) {
                ctx.request_repaint();
                ctx.request_repaint_of(egui::ViewportId::ROOT);
            }
        },
    );
}

pub fn commands(app: &SheetApp) -> Vec<Json> {
    let mut commands = vec![
        json!({"id":"vba.showDeveloper","label":"Show Developer Tab","params":"{on: bool}","enabled":true,"ui":true}),
        json!({"id":"vba.open","label":"Visual Basic","params":"{source?: string, project?: bool, separateWindow?: bool, entry?: string}","enabled":app.session.active().is_some(),"ui":true}),
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
        ("vba.apply", "Save VBA Source to Workbook", "{}", None),
        ("vba.entries", "VBA Procedures", "{}", None),
        ("vba.invoke", "Run Assigned Macro", "{entry: string}", None),
        ("vba.assign", "Assign Macro", "{entry: string, id?: shape, at?: cell, caption?: string}", None),
        ("vba.pin", "Assign Toolbar Macro", "{entry: string, remove?: bool}", None),
        ("vba.designMode", "Design Mode", "{on: bool}", None),
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
        if editor.ide.modules.iter().any(|m| m.dirty()) || editor.ide.has_modified_forms() {
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
    if id == "vba.showDeveloper" {
        app.ui.developer = p.get("on").and_then(Json::as_bool).ok_or("on must be true or false")?;
        if !app.ui.developer && app.ui.ribbon_tab == "Developer" {
            app.ui.ribbon_tab = "Home".into();
        }
        return Ok(Json::Null);
    }
    if id == "vba.open" {
        let separate = placement(app, p)?;
        if let Some(editor) = app.vba.as_mut() {
            if editor.bridge.document() == app.session.doc().map_err(error)?.uid
                && p.get("source").is_none()
                && !p.get("project").and_then(Json::as_bool).unwrap_or(false)
            {
                if let Some(entry) = p.get("entry").and_then(Json::as_str) {
                    navigate(editor, entry)?;
                }
                if p.get("separateWindow").is_some() {
                    editor.separate_window = separate;
                }
                editor.visible = true;
                editor.focus_window = true;
                return Ok(json!({"mode":editor.ide.mode().label()}));
            }
            return Err("Close the existing VBA project before opening another.".into());
        }
        app.settle_drafts()?;
        let bridge = Bridge::new(&app.session).map_err(error)?;
        let stored = Project::read(&app.session.doc().map_err(error)?.wb).map_err(error)?;
        let binary = app.session.doc().map_err(error)?.wb.vba_project.is_some();
        let mut ide = if p.get("source").is_none() && !p.get("project").and_then(Json::as_bool).unwrap_or(false) && stored.is_some() {
            stored.ok_or("missing source")?.ide(&bridge).map_err(error)?
        } else if p.get("project").and_then(Json::as_bool).unwrap_or(false) || (p.get("source").is_none() && binary) {
            bridge.import_project(&app.session).map_err(error)?
        } else {
            let mut ide = Ide::new();
            let source = match p.get("source") {
                None => "Option Explicit\n",
                Some(Json::String(s)) => s,
                Some(_) => return Err("source must be VBA text".into()),
            };
            ide.add_module("Module1.bas", source).map_err(error)?;
            bridge.configure(&mut ide).map_err(error)?;
            ide
        };
        ide.title = format!("VBAProject ({})", app.session.doc().map_err(error)?.title);
        let mut editor = Editor {
            ide,
            bridge,
            view: EditorUi::default(),
            visible: true,
            confirm_close: false,
            design_mode: false,
            functions: None,
            separate_window: separate,
            focus_window: true,
            #[cfg(all(feature = "vba-window", not(target_arch = "wasm32")))]
            owner_viewport: egui::ViewportId::ROOT,
        };
        if let Some(entry) = p.get("entry").and_then(Json::as_str) {
            navigate(&mut editor, entry)?;
        }
        app.vba = Some(editor);
        return Ok(json!({"opened":true,"executed":false}));
    }
    if matches!(id, "vba.invoke" | "vba.assign" | "vba.pin" | "vba.entries" | "vba.designMode") && app.vba.is_none() {
        if id == "vba.invoke" {
            let doc = app.session.doc().map_err(error)?;
            if doc.wb.vba_project.is_none() && Project::read(&doc.wb).map_err(error)?.is_none() {
                return Err("This workbook has no VBA source for its assigned macro.".into());
            }
        }
        run(app, "vba.open", &json!({}))?;
        if let Some(editor) = app.vba.as_mut() {
            editor.visible = false;
        }
    }
    if matches!(id, "vba.apply" | "vba.assign" | "vba.pin") {
        return publish(app, id, p);
    }
    let editor = app.vba.as_mut().ok_or("VBA editor is not open")?;
    match id {
        "vba.status" => {
            return Ok(
                json!({"mode":editor.ide.mode().label(),"separateWindow":editor.separate_window,"visible":editor.visible,"output":editor.ide.output(),"diagnostics":editor.ide.diagnostics.iter().map(ToString::to_string).collect::<Vec<_>>()}),
            );
        }
        "vba.trust" => {
            let allow = p.get("allow").and_then(Json::as_bool).ok_or("allow must be true or false")?;
            editor.bridge.consent(allow);
            if let Some(f) = &editor.functions {
                f.allowed.store(editor.bridge.permitted(&app.session), Ordering::Release);
            }
            recalculate(&mut app.session)?;
            if allow && editor.functions.is_none() && Project::read(&app.session.doc().map_err(error)?.wb).map_err(error)?.is_some() {
                return publish(app, "vba.apply", &json!({}));
            }
        }
        "vba.designMode" => editor.design_mode = p.get("on").and_then(Json::as_bool).ok_or("on must be true or false")?,
        "vba.entries" => {
            let project = Project::from_ide(&editor.ide, None).map_err(error)?;
            return Ok(json!(
                project
                    .entries()
                    .map_err(error)?
                    .iter()
                    .map(|e| json!({"name":e.name,"signature":e.signature,"function":e.function}))
                    .collect::<Vec<_>>()
            ));
        }
        "vba.invoke" => {
            let entry = p.get("entry").and_then(Json::as_str).ok_or("entry is required")?;
            let entries = Project::from_ide(&editor.ide, None).map_err(error)?.entries().map_err(error)?;
            if !entries.iter().any(|e| !e.function && e.name.eq_ignore_ascii_case(entry)) {
                return Err("Select a public Sub with no parameters.".into());
            }
            if matches!(editor.ide.mode(), Mode::Running | Mode::Waiting | Mode::Break | Mode::Stale) {
                return Err("Finish or reset the current macro first.".into());
            }
            editor.ide.entry = entry.into();
            editor.bridge.dispatch(&mut app.session, &mut editor.ide, Action::Run).map_err(error)?;
            editor.visible = false;
        }
        "vba.return" => editor.visible = false,
        "vba.close" => {
            let discard = p.get("discard").and_then(Json::as_bool).unwrap_or(false);
            editor.bridge.close(&mut app.session, &mut editor.ide, discard).map_err(error)?;
            if let Some(f) = &editor.functions {
                f.allowed.store(false, Ordering::Release);
            }
            let d = app.session.doc_mut().map_err(error)?;
            d.calc.custom_functions = None;
            recalculate(&mut app.session)?;
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
    let was_running = matches!(editor.ide.mode(), Mode::Running | Mode::Waiting);
    if let Some(f) = &editor.functions {
        let allowed = editor.bridge.permitted(&app.session);
        if f.allowed.swap(allowed, Ordering::AcqRel) != allowed {
            let _ = recalculate(&mut app.session);
        }
    }
    if let Err(e) = editor.bridge.tick(&mut app.session, &mut editor.ide) {
        editor.ide.message(e.to_string());
    }
    if was_running && matches!(editor.ide.mode(), Mode::Break | Mode::Failed) {
        editor.visible = true;
        editor.focus_window = true;
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
fn recalculate(session: &mut gridcraft_engine::Session) -> Result<()> {
    let doc = session.doc_mut().map_err(error)?;
    doc.calc.recalc_all(Arc::make_mut(&mut doc.wb));
    Ok(())
}
fn publish(app: &mut SheetApp, id: &str, p: &Json) -> Result<Json> {
    let editor = app.vba.as_ref().ok_or("VBA editor is not open")?;
    if !editor.bridge.current(&app.session) {
        return Err("Workbook or sheets changed; close and reopen the VBA project.".into());
    }
    if matches!(editor.ide.mode(), Mode::Running | Mode::Waiting | Mode::Break | Mode::Stale) {
        return Err("Finish or reset VBA before saving or assigning macros.".into());
    }
    let old = Project::read(&app.session.doc().map_err(error)?.wb).map_err(error)?;
    let mut project = Project::from_ide(&editor.ide, old).map_err(error)?;
    let entry = p.get("entry").and_then(Json::as_str).unwrap_or("");
    if id != "vba.apply" && !project.entries().map_err(error)?.iter().any(|e| !e.function && e.name.eq_ignore_ascii_case(entry)) {
        return Err("Select a public Sub with no parameters.".into());
    }
    let functions = Arc::new(Functions::compile(&project, &editor.bridge, editor.bridge.permitted(&app.session)).map_err(error)?);
    let mut shape = None;
    if id == "vba.assign" {
        let doc = app.session.doc().map_err(error)?;
        let sheet = doc.wb.active().ok_or("no sheet")?;
        if sheet.is_protected() {
            return Err("Unprotect the sheet before assigning a macro.".into());
        }
        if let Some(value) = p.get("id") {
            let id = value.as_u64().and_then(|n| u32::try_from(n).ok()).ok_or("id must be a shape id")?;
            if !sheet.shapes.iter().any(|s| s.id == id) {
                return Err("no such worksheet shape".into());
            }
            shape = Some(id);
        }
        if sheet.shapes.iter().filter(|s| s.on_action.is_some()).count() >= 256
            && !sheet.shapes.iter().any(|s| Some(s.id) == shape && s.on_action.is_some())
        {
            return Err("worksheet button budget exceeded".into());
        }
    }
    if id == "vba.pin" {
        project.toolbar.retain(|e| !e.eq_ignore_ascii_case(entry));
        if !p.get("remove").and_then(Json::as_bool).unwrap_or(false) {
            if project.toolbar.len() >= 32 {
                return Err("toolbar macro budget exceeded".into());
            }
            project.toolbar.push(entry.into());
        }
    }
    // Both edits use the normal command lifecycle and make one undo step.
    let token = app.session.group_begin("VBA source and assignment").map_err(error)?;
    let result: Result<()> = (|| {
        if id == "vba.assign" {
            let assigned = if let Some(id) = shape {
                id
            } else {
                let value = app.session.execute("insert.shape",json!({"kind":"roundedRectangle","at":p.get("at").cloned().unwrap_or(Json::Null),"width":144,"height":32,"text":p.get("caption").and_then(Json::as_str).unwrap_or(entry)})).map_err(error)?;
                value.get("shape").and_then(Json::as_u64).and_then(|n| u32::try_from(n).ok()).ok_or("shape was not created")?
            };
            shape = Some(assigned);
            app.session.execute("shape.setMacro", json!({"id":assigned,"entry":entry})).map_err(error)?;
        }
        project.store(&mut app.session).map_err(error)
    })();
    if let Err(e) = result {
        app.session.group_rollback(token).map_err(error)?;
        return Err(e);
    }
    app.session.group_end(token).map_err(error)?;
    // Assignment metadata is part of the provider's source fingerprint too.
    let editor = app.vba.as_mut().ok_or("VBA editor closed")?;
    let functions = if id == "vba.apply" {
        functions
    } else {
        Arc::new(Functions::compile(&project, &editor.bridge, editor.bridge.permitted(&app.session)).map_err(error)?)
    };
    if let Some(old) = editor.functions.replace(functions.clone()) {
        old.allowed.store(false, Ordering::Release);
    }
    app.session.doc_mut().map_err(error)?.calc.custom_functions = Some(functions);
    let revisions: Vec<_> = editor.ide.modules.iter().map(|m| (m.id, m.revision)).collect();
    for (id, revision) in revisions {
        editor.ide.mark_saved(id, revision).map_err(error)?;
    }
    recalculate(&mut app.session)?;
    Ok(json!({"saved":true,"shape":shape}))
}

pub fn assigned(app: &SheetApp, shape: u32) -> Option<String> {
    let doc = app.session.doc().ok()?;
    let sheet = doc.wb.active()?;
    let entry = sheet.shapes.iter().find(|s| s.id == shape)?.on_action.as_deref()?;
    let entry = entry.strip_prefix("[0]!").unwrap_or(entry);
    (entry.len() <= 256).then(|| entry.to_string())
}

pub fn developer(app: &mut SheetApp, ui: &mut egui::Ui) {
    use crate::{icons::Icon, widgets::big_button};
    ui.vertical(|ui| {
        ui.horizontal(|ui| {
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
        });
        ui.label(egui::RichText::new("Code").small());
    });
    ui.separator();
    ui.vertical(|ui| {
        ui.horizontal(|ui| {
            let insert = big_button(ui, Icon::Shapes, "Insert", "Insert worksheet button", true);
            insert.widget_info(|| egui::WidgetInfo::labeled(egui::Role::Button, ui.is_enabled(), "Insert worksheet button"));
            egui::Popup::menu(&insert).show(|ui| {
                if ui.button("Button (Assign Macro?)").clicked() {
                    app.open_dialog("macros", json!({"assign":true}));
                    ui.close();
                }
            });
            let on = app.vba.as_ref().is_some_and(|e| e.design_mode);
            ui.vertical(|ui| {
                if crate::widgets::toggle_button(ui, Icon::Shapes, on, "Design Mode").clicked() {
                    app.run_or_alert("vba.designMode", json!({"on":!on}));
                }
                ui.label("Design Mode");
            });
        });
        ui.label(egui::RichText::new("Controls").small());
    });
}

pub fn toolbar(app: &mut SheetApp, ui: &mut egui::Ui) {
    let pins = app.session.doc().ok().and_then(|d| Project::read(&d.wb).ok().flatten()).map(|p| p.toolbar).unwrap_or_default();
    for entry in pins {
        if ui.button(&entry).on_hover_text("Run assigned VBA macro").clicked() {
            app.run_or_alert("vba.invoke", json!({"entry":entry}));
        }
    }
}
pub fn shortcuts(app: &mut SheetApp, ctx: &egui::Context) {
    if ctx.input_mut(|i| i.consume_key(egui::Modifiers::ALT, egui::Key::F11)) {
        let id = if app.vba.as_ref().is_some_and(|e| e.visible) { "vba.return" } else { "vba.open" };
        app.run_or_alert(id, json!({}));
    }
    if ctx.input_mut(|i| i.consume_key(egui::Modifiers::ALT, egui::Key::F8)) {
        if let Some(e) = app.vba.as_mut() {
            e.visible = false;
        }
        app.open_dialog("macros", json!({}));
    }
}

pub fn macro_picker(app: &mut SheetApp, ui: &mut egui::Ui, values: &mut serde_json::Map<String, Json>) -> bool {
    if app.vba.is_none() {
        if let Err(e) = run(app, "vba.open", &json!({})) {
            ui.label(e);
            return false;
        }
        if let Some(e) = app.vba.as_mut() {
            e.visible = false;
        }
    }
    let entries = app.vba.as_ref().and_then(|e| Project::from_ide(&e.ide, None).and_then(|p| p.entries()).ok()).unwrap_or_default();
    let names: Vec<_> = entries.into_iter().filter(|e| !e.function).map(|e| e.name).collect();
    let mut selected = values.get("macro").and_then(Json::as_str).map(str::to_string).unwrap_or_else(|| names.first().cloned().unwrap_or_default());
    let label = ui.label("Macro name:");
    ui.text_edit_singleline(&mut selected).labelled_by(label.id);
    let mut edit = false;
    egui::ScrollArea::vertical().max_height(200.).show(ui, |ui| {
        for name in &names {
            let r = ui.selectable_label(selected.eq_ignore_ascii_case(name), name);
            if r.clicked() {
                selected = name.clone();
            }
            if r.double_clicked() {
                selected = name.clone();
                edit = true;
            }
        }
        if names.is_empty() {
            ui.weak("No public macros in this workbook.");
        }
    });
    values.insert("macro".into(), json!(selected));
    let valid = names.iter().any(|n| n.eq_ignore_ascii_case(&selected));
    let assign = values.get("assign").and_then(Json::as_bool).unwrap_or(false) || values.get("shape").is_some();
    let mut action = None;
    ui.horizontal(|ui| {
        if assign {
            if ui.add_enabled(valid, egui::Button::new("Assign")).clicked() {
                action = Some("vba.assign");
            }
        } else {
            if ui.add_enabled(valid, egui::Button::new("Run")).clicked() {
                action = Some("vba.invoke");
            }
            if ui.add_enabled(valid, egui::Button::new("Step Into")).clicked() {
                action = Some("vba.stepInto");
            }
            if ui.add_enabled(valid, egui::Button::new("Edit")).clicked() {
                edit = true;
            }
        }
    });
    if !assign {
        ui.collapsing("GridCraft customizations", |ui| {
            if ui.add_enabled(valid, egui::Button::new("Assign Button")).clicked() {
                action = Some("vba.assign");
            }
            if ui.add_enabled(valid, egui::Button::new("Add to Toolbar")).clicked() {
                action = Some("vba.pin");
            }
        });
    }
    if edit && valid {
        action = Some("vba.open");
    }
    if let Some(id) = action {
        let mut p = json!({"entry":selected});
        if let Some(shape) = values.get("shape") {
            p["id"] = shape.clone();
        }
        match app.run(id, p) {
            Ok(_) => {
                if id == "vba.stepInto"
                    && let Some(e) = app.vba.as_mut()
                {
                    e.visible = true;
                    e.focus_window = true;
                }
                return id != "vba.pin";
            }
            Err(e) => app.message = Some(("VBA".into(), e)),
        }
    }
    false
}

pub fn show(app: &mut SheetApp, ui: &mut egui::Ui) {
    let Some(mut editor) = app.vba.take() else { return };
    let mut commands = Vec::new();
    let mut save_workbook = false;
    let mut allow = false;
    egui::Panel::top("gridcraft_vba_host").show(ui, |ui| {
        ui.horizontal(|ui| {
            ui.label("GridCraft VBA");
            if ui.button("Save Source to Workbook").clicked() {
                save_workbook = true;
            }
            if ui.button("Return to workbook").clicked() {
                editor.visible = false;
            }
            if ui.button("Allow VBA for this editor session").clicked() {
                allow = true;
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
                if editor.ide.modules.iter().any(|m| m.dirty()) || editor.ide.has_modified_forms() {
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
        if allow {
            app.run_or_alert("vba.trust", json!({"allow":true}));
        }
        if save_workbook {
            app.run_or_alert("vba.apply", json!({}));
        }
    } else if let Ok(doc) = app.session.doc_mut() {
        doc.calc.custom_functions = None;
        let _ = recalculate(&mut app.session);
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
