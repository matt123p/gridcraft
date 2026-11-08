//! Optional shell glue; the editor/debugger itself is rendered by tiny-vba.
#[path = "vba_interaction.rs"]
mod interaction;
#[path = "vba_refedit.rs"]
mod refedit;
use crate::{SheetApp, theme};
use gridcraft_vba::worksheet::{Functions, Project};
use gridcraft_vba::{
    Action, Bridge, Ide, LifecycleEvent, LifecycleEventBatch, LifecyclePrompt, Mode, PromptRequest, PromptResponse, PromptServices,
    vba_ide::HostRequest,
};
use serde_json::{Value as Json, json};
use std::sync::{Arc, atomic::Ordering};
use vba_ui_egui::{Command, EditorUi, Theme};

pub struct Editor {
    pub ide: Ide,
    pub bridge: Bridge,
    pub view: EditorUi,
    pub visible: bool,
    confirm_close: bool,
    pub(crate) pending_prompt: Option<PromptRequest>,
    pending_print: Option<PendingPrint>,
    pending_interaction: Option<interaction::Pending>,
    refedit_picker: Option<refedit::Picker>,
    external_lifecycle: Option<ExternalLifecycle>,
    lifecycle_commit: bool,
    approved_close: Option<u64>,
    close_after_save: Option<(u64, bool)>,
    lifecycle_result: Option<Json>,
    saved_events: std::collections::VecDeque<(u64, bool)>,
    pub design_mode: bool,
    functions: Option<Arc<Functions>>,
    pub separate_window: bool,
    focus_window: bool,
    #[cfg(all(feature = "vba-window", not(target_arch = "wasm32")))]
    owner_viewport: egui::ViewportId,
}
enum PendingPrint {
    Worker { token: u64 },
    Preview { token: u64, editor_visible: bool },
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
    let Some(editor) = app.vba.as_mut().filter(|e| {
        e.visible
            && e.separate_window
            && (e.ide.form_range_selection().ok().flatten().is_none() || matches!(e.ide.mode(), Mode::Break | Mode::Stale | Mode::Failed))
    }) else {
        return;
    };
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
        ("vba.exportNativePackage", "Export Edited XLSM Copy", "{}", None),
        ("vba.apply", "Save VBA Source to Workbook", "{}", None),
        ("vba.entries", "VBA Procedures", "{}", None),
        ("vba.invoke", "Run Assigned Macro", "{entry: string}", None),
        ("vba.control", "Activate Form Control", "{sheet?, id, value?, item?, delta?}", None),
        ("vba.workbookOpen", "Run Workbook_Open Event", "{}", None),
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
        let enabled = if id == "vba.control" {
            app.session.active().is_some()
                && app
                    .vba
                    .as_ref()
                    .is_none_or(|e| !e.design_mode && !matches!(e.ide.mode(), Mode::Running | Mode::Waiting | Mode::Break | Mode::Stale))
        } else {
            enabled
        };
        commands.push(json!({"id":id,"label":label,"params":params,"enabled":enabled,"ui":true}));
    }
    commands
}

struct ExternalLifecycle {
    doc: u64,
    id: String,
    params: Json,
    batch: Option<LifecycleEventBatch>,
    event: LifecycleEvent,
    after: bool,
    save: bool,
}
pub fn before_command(app: &mut SheetApp, id: &str, params: &Json) -> Result<bool> {
    interaction::before_command(app, id, params)?;
    let Some(editor) = app.vba.as_mut() else {
        return Ok(true);
    };
    if editor.lifecycle_commit {
        return Ok(true);
    }
    refedit::before_command(editor, id)?;
    if id == "ui.dialog.cancel" && editor.pending_prompt.as_ref().is_some_and(|p| matches!(p.prompt, LifecyclePrompt::SaveChanges { .. })) {
        let token = editor.pending_prompt.as_ref().map(|p| p.token).ok_or("missing save prompt")?;
        editor.bridge.respond_to_prompt(token, PromptResponse::Cancel).map_err(|e| e.description)?;
        editor.pending_prompt = None;
    }
    if id == "ui.dialog.cancel" {
        editor.approved_close = None;
        editor.close_after_save = None;
    }
    if !matches!(id, "file.save" | "file.saveAs" | "file.close" | "window.close") {
        return Ok(true);
    }
    if editor.external_lifecycle.is_some() {
        return Err("A workbook lifecycle event is still in progress.".into());
    }
    let close = matches!(id, "file.close" | "window.close");
    if close && (editor.ide.modules.iter().any(|m| m.dirty()) || editor.ide.has_modified_forms()) {
        return Err("Save VBA source or explicitly discard it with vba.close before closing the workbook.".into());
    }
    let doc = if id == "window.close" {
        let target = gridcraft_engine::Target::from_json(params).map_err(error)?;
        let window = app.session.target_window(target).map_err(error)?;
        let uid = app.session.window(window).ok_or("window closed")?.doc_uid;
        app.session.document(uid)
    } else {
        app.session.active()
    };
    let Some(doc) = doc else {
        return Ok(true);
    };
    let doc_id = doc.uid;
    if close && editor.approved_close == Some(doc_id) && params.get("force").and_then(Json::as_bool) == Some(true) {
        editor.approved_close = None;
        return Ok(true);
    }
    if !close && editor.saved_events.len() >= 256 {
        return Err("VBA lifecycle event queue limit reached; finish queued events before saving.".into());
    }
    if editor.bridge.permitted(&app.session) && editor.bridge.has_lifecycle_handlers(&editor.ide) {
        if !matches!(editor.ide.mode(), Mode::Design | Mode::Completed) {
            return Err("Finish or reset VBA before a native save/close; its lifecycle callbacks cannot run during another invocation.".into());
        }
        if ((id == "window.close") || (id == "file.close" && params.get("document").and_then(Json::as_bool) != Some(true)))
            && app.session.windows().iter().filter(|w| w.doc_uid == doc_id).count() > 1
        {
            return Ok(true);
        }
        let event =
            if close { LifecycleEvent::BeforeClose } else { LifecycleEvent::BeforeSave { save_as_ui: id == "file.saveAs" || doc.path.is_none() } };
        let mut command_params = params.clone();
        if id == "window.close" {
            let window = app.session.target_window(gridcraft_engine::Target::from_json(params).map_err(error)?).map_err(error)?;
            if let Some(map) = command_params.as_object_mut() {
                map.remove("doc");
                map.insert("window".into(), json!(window));
            } else {
                command_params = json!({"window":window});
            }
        }
        editor.external_lifecycle =
            Some(ExternalLifecycle { doc: doc_id, id: id.into(), params: command_params, batch: None, event, after: false, save: !close });
        editor.lifecycle_result = Some(json!({"state":"pending","command":id}));
        return Ok(false);
    }
    if close && editor.ide.enabled(Action::Reset) {
        editor.bridge.dispatch(&mut app.session, &mut editor.ide, Action::Reset).map_err(error)?;
    }
    Ok(true)
}

/// Capture supported cell/range content edits after the engine commits them. VBA-authored
/// writes queue events at the native host boundary instead, avoiding dispatch before commit.
pub fn after_engine_command(app: &mut SheetApp, id: &str, params: &Json, result: &Json) {
    if matches!(id, "file.save" | "file.saveAs")
        && result.get("path").and_then(Json::as_str).is_some()
        && result.get("bytes").and_then(Json::as_u64).is_some()
        && let Some(editor) = app.vba.as_mut()
        && editor.bridge.permitted(&app.session)
        && editor.bridge.has_lifecycle_handlers(&editor.ide)
        && let Some(doc) = app.session.active()
    {
        if editor.saved_events.len() < 256 {
            editor.saved_events.push_back((doc.uid, true));
        } else {
            editor.ide.message("VBA lifecycle event queue limit reached");
        }
    }
    if !matches!(
        id,
        "cell.set"
            | "range.setContents"
            | "range.setValues"
            | "range.fill"
            | "edit.clearContents"
            | "edit.clearAll"
            | "edit.cut"
            | "edit.paste"
            | "edit.pasteSpecial"
            | "edit.autoFill"
            | "edit.fillDown"
            | "edit.fillRight"
            | "edit.fillUp"
            | "edit.fillLeft"
            | "edit.fillSeries"
            | "edit.flashFill"
            | "data.flashFill"
            | "data.sortAscending"
            | "data.sortDescending"
            | "data.sort"
            | "data.removeDuplicates"
            | "data.textToColumns"
    ) {
        return;
    }
    if id == "edit.pasteSpecial"
        && params.get("link").and_then(Json::as_bool).is_none_or(|link| !link)
        && params
            .get("what")
            .and_then(Json::as_str)
            .is_some_and(|what| matches!(what, "formats" | "comments" | "notes" | "validation" | "columnWidths"))
    {
        return;
    }
    if id == "edit.autoFill" && params.get("mode").and_then(Json::as_str).is_some_and(|mode| mode.eq_ignore_ascii_case("formats")) {
        return;
    }
    if id == "data.removeDuplicates" && result.get("removed").and_then(Json::as_u64) == Some(0) {
        return;
    }
    let Some(editor) = app.vba.as_mut() else { return };
    if !editor.bridge.current(&app.session) {
        return;
    }
    let Ok(doc) = app.session.doc() else { return };
    let default_sheet = doc.wb.active_sheet();
    let sheet = match params.get("sheet") {
        Some(Json::Number(number)) => number.as_u64().and_then(|index| usize::try_from(index).ok()),
        Some(Json::String(name)) => doc.wb.sheet_index(name),
        None | Some(Json::Null) => Some(default_sheet),
        _ => None,
    };
    let Some(sheet) = sheet.filter(|index| *index < doc.wb.sheets.len()) else { return };
    let target = if id == "cell.set" {
        params.get("cell").and_then(Json::as_str).and_then(gridcraft_engine::core::CellRef::parse).map(gridcraft_engine::core::RangeRef::cell)
    } else if matches!(id, "edit.paste" | "edit.pasteSpecial")
        && let Some(pasted) = result.get("pasted").and_then(Json::as_str).and_then(gridcraft_engine::core::RangeRef::parse)
    {
        Some(pasted)
    } else if id == "edit.paste"
        && let (Some(rows), Some(cols)) = (result.get("rows").and_then(Json::as_u64), result.get("cols").and_then(Json::as_u64))
        && rows > 0
        && cols > 0
    {
        let start =
            params.get("at").and_then(Json::as_str).and_then(gridcraft_engine::core::CellRef::parse).unwrap_or_else(|| doc.selection.current().start);
        let row_offset = i64::try_from(rows.saturating_sub(1)).ok();
        let col_offset = i64::try_from(cols.saturating_sub(1)).ok();
        row_offset.zip(col_offset).map(|(row_offset, col_offset)| {
            let end = start.offset_clamped(row_offset, col_offset);
            gridcraft_engine::core::RangeRef::new(start, end)
        })
    } else if id == "edit.autoFill" {
        params.get("target").and_then(Json::as_str).and_then(gridcraft_engine::core::RangeRef::parse)
    } else if id == "data.textToColumns" {
        result.get("range").and_then(Json::as_str).and_then(gridcraft_engine::core::RangeRef::parse)
    } else {
        params
            .get("range")
            .and_then(Json::as_str)
            .and_then(gridcraft_engine::core::RangeRef::parse)
            .or_else(|| {
                params.get("at").and_then(Json::as_str).and_then(gridcraft_engine::core::CellRef::parse).map(gridcraft_engine::core::RangeRef::cell)
            })
            .or_else(|| Some(doc.selection.current()))
    };
    if let Some(target) = target
        && let Err(error) = editor.bridge.notify_sheet_change(&app.session, sheet, target)
    {
        editor.ide.message(error.to_string());
    }
}

/// UI-only commands are also reachable through SheetApp::run/control, with no dialogs.
pub fn run(app: &mut SheetApp, id: &str, p: &Json) -> Result<Json> {
    if id == "vba.control" {
        return activate_control(app, p);
    }
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
        // Document context is metadata only, but it must see the live workbook's
        // code names so worksheet modules resolve implicit Me members and the IDE
        // completes host objects without ever invoking them.
        bridge.configure_document(&mut ide, &app.session.doc().map_err(error)?.wb).map_err(error)?;
        let mut editor = Editor {
            ide,
            bridge,
            view: EditorUi::default(),
            visible: true,
            confirm_close: false,
            pending_prompt: None,
            pending_print: None,
            pending_interaction: None,
            refedit_picker: None,
            external_lifecycle: None,
            lifecycle_commit: false,
            approved_close: None,
            close_after_save: None,
            lifecycle_result: None,
            saved_events: Default::default(),
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
    if id == "vba.importSource" {
        let mut editor = app.vba.take().ok_or("VBA editor is not open")?;
        let imported = import_source(app, &mut editor.ide);
        app.vba = Some(editor);
        return imported.map(|imported| json!({"imported":imported}));
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
    if id == "vba.exportNativePackage" {
        return export_native_package_copy(app);
    }
    if matches!(id, "vba.apply" | "vba.assign" | "vba.pin") {
        return publish(app, id, p);
    }
    let editor = app.vba.as_mut().ok_or("VBA editor is not open")?;
    match id {
        "vba.workbookOpen" => {
            editor.bridge.queue_workbook_open(&app.session).map_err(error)?;
            editor.visible = true;
            return Ok(json!({"queued":true,"executed":false}));
        }
        "vba.status" => {
            return Ok(
                json!({"mode":editor.ide.mode().label(),"separateWindow":editor.separate_window,"visible":editor.visible,"output":editor.ide.output(),"lifecycle":editor.lifecycle_result,"diagnostics":editor.ide.diagnostics.iter().map(ToString::to_string).collect::<Vec<_>>()}),
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
            if let (Some(sheet), Some(control)) = (
                p.get("controlSheet").and_then(Json::as_u64).and_then(|n| usize::try_from(n).ok()),
                p.get("controlId").and_then(Json::as_u64).and_then(|n| u32::try_from(n).ok()),
            ) {
                editor.bridge.bind_control_invocation(&app.session, sheet, control).map_err(error)?;
            }
            editor.bridge.set_prompt_services(prompt_services(&app.services));
            editor.bridge.set_print_services(print_services(&app.services, app.chrome.file_tab));
            editor.bridge.set_interaction_services(gridcraft_vba::InteractionServices {
                dialogs: true,
                open_picker: app.services.pick_open.is_some(),
                save_picker: app.services.pick_save.is_some(),
                status_visibility: true,
            });
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
            if editor.pending_interaction.take().is_some() && app.dialog.as_ref().is_some_and(|d| d.values.contains_key("om15Token")) {
                app.dialog = None;
            }
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
            if matches!(action, Action::Reset) {
                editor.external_lifecycle = None;
                editor.refedit_picker = None;
                editor.approved_close = None;
                editor.close_after_save = None;
                editor.saved_events.clear();
                editor.pending_prompt = None;
                editor.lifecycle_result = Some(json!({"state":"cancelled"}));
            }
            editor.bridge.set_prompt_services(prompt_services(&app.services));
            editor.bridge.set_print_services(print_services(&app.services, app.chrome.file_tab));
            editor.bridge.set_interaction_services(gridcraft_vba::InteractionServices {
                dialogs: true,
                open_picker: app.services.pick_open.is_some(),
                save_picker: app.services.pick_save.is_some(),
                status_visibility: true,
            });
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
    editor.bridge.set_prompt_services(prompt_services(&app.services));
    editor.bridge.set_print_services(print_services(&app.services, app.chrome.file_tab));
    editor.bridge.set_interaction_services(gridcraft_vba::InteractionServices {
        dialogs: true,
        open_picker: app.services.pick_open.is_some(),
        save_picker: app.services.pick_save.is_some(),
        status_visibility: true,
    });
    editor.bridge.set_formula_locale(app.ui.language.formula_locale());
    let progress = if editor.external_lifecycle.as_ref().is_some_and(|p| p.batch.is_some()) {
        editor.bridge.tick_lifecycle(&mut app.session, &mut editor.ide)
    } else {
        editor.bridge.tick(&mut app.session, &mut editor.ide)
    };
    if let Err(e) = progress {
        editor.ide.message(e.to_string());
    }
    if !matches!(editor.ide.mode(), Mode::Running | Mode::Waiting | Mode::Break | Mode::Stale) {
        editor.pending_prompt = None;
    }
    if editor.pending_prompt.is_none() && app.dialog.as_ref().is_some_and(|d| d.values.contains_key("vbaToken")) {
        app.dialog = None;
    }
    if editor.pending_prompt.is_none() {
        editor.pending_prompt = editor.bridge.take_prompt_request();
    }
    if let Some(request) = editor.pending_prompt.clone() {
        let response = match request.prompt {
            LifecyclePrompt::OpenFile => Some(PromptResponse::Path(app.services.pick_open.as_ref().and_then(|pick| pick()))),
            LifecyclePrompt::SaveFile { suggested } => Some(PromptResponse::Path(app.services.pick_save.as_ref().and_then(|pick| pick(&suggested)))),
            LifecyclePrompt::SaveChanges { workbook } => {
                if app.dialog.is_none() {
                    app.dialog =
                        Some(crate::dialogs::Dialog::custom("saveChanges", "Save Changes", json!({"title":workbook,"vbaToken":request.token})));
                }
                None
            }
        };
        if let Some(response) = response {
            if let Err(e) = editor.bridge.respond_to_prompt(request.token, response) {
                editor.ide.message(e.description);
            }
            editor.pending_prompt = None;
        }
    }
    if was_running && matches!(editor.ide.mode(), Mode::Break | Mode::Failed) {
        editor.visible = true;
        editor.focus_window = true;
    }
    if matches!(editor.ide.mode(), Mode::Running | Mode::Waiting) || editor.external_lifecycle.is_some() {
        ctx.request_repaint();
    }
    if editor.external_lifecycle.is_none()
        && matches!(editor.ide.mode(), Mode::Design | Mode::Completed)
        && let Some((doc, success)) = editor.saved_events.pop_front()
    {
        editor.external_lifecycle = Some(ExternalLifecycle {
            doc,
            id: "file.save".into(),
            params: Json::Null,
            batch: None,
            event: LifecycleEvent::AfterSave { success },
            after: true,
            save: false,
        });
    }
    if let Err(e) = refedit::poll(editor, &mut app.session) {
        editor.ide.message(e);
    }
    let commit = poll_external_lifecycle(editor, &mut app.session);
    if let Some((doc, id, params, save)) = commit {
        commit_lifecycle(app, ctx, doc, id, params, save);
    }
    finish_native_close(app, ctx);
    poll_print_io(app, ctx);
    interaction::poll(app, ctx);
}
/// Owns only service tokens/worker state across frames, never a document borrow.
fn poll_print_io(app: &mut SheetApp, ctx: &egui::Context) {
    use gridcraft_vba::{PrintOperation as Op, PrintResponse};
    let Some(mut editor) = app.vba.take() else {
        return;
    };
    if let Some(pending) = editor.pending_print.take() {
        let restore = match &pending {
            PendingPrint::Preview { editor_visible, .. } => Some(*editor_visible),
            _ => None,
        };
        let token = match pending {
            PendingPrint::Worker { token } | PendingPrint::Preview { token, .. } => token,
        };
        if editor.bridge.print_request_pending(token) {
            if !editor.bridge.print_request_permitted(&app.session, token) {
                let _ = editor.bridge.respond_to_print(token, Err("Print request canceled: workbook closed or consent revoked.".into()));
            } else {
                let result = match pending {
                    PendingPrint::Preview { .. } => {
                        if app.backstage.as_ref().is_some_and(|b| b.page == crate::backstage::Page::Print) {
                            None
                        } else {
                            Some(Ok(PrintResponse::Done))
                        }
                    }
                    PendingPrint::Worker { .. } => {
                        app.printing.poll(ctx);
                        match app.printing.message.as_ref() {
                            Some(crate::printing::Feedback::Submitted(_)) => Some(Ok(PrintResponse::Done)),
                            Some(crate::printing::Feedback::Failed(e)) => Some(Err(e.clone())),
                            Some(crate::printing::Feedback::WorkerStopped) => Some(Err("printing worker stopped".into())),
                            _ => None,
                        }
                    }
                };
                if let Some(result) = result {
                    let _ = editor.bridge.respond_to_print(token, result);
                } else {
                    editor.pending_print = Some(pending);
                }
            }
        }
        if editor.pending_print.is_none()
            && let Some(visible) = restore
        {
            editor.visible = visible || matches!(editor.ide.mode(), Mode::Break | Mode::Failed);
            editor.focus_window = editor.visible;
            if app.backstage.as_ref().is_some_and(|b| b.page == crate::backstage::Page::Print) {
                app.backstage = None;
            }
        }
    }
    if editor.pending_print.is_none()
        && let Some(request) = editor.bridge.take_print_request()
    {
        let token = request.token;
        let result = (|| -> std::result::Result<Option<PrintResponse>, String> {
            if !editor.bridge.print_request_permitted(&app.session, token) {
                return Err("workbook closed or consent revoked".into());
            }
            match request.kind {
                Op::ExportWithPicker { mut params, suggested, sheet, range_epoch } => {
                    let pick = app.services.pick_save.as_ref().ok_or("PDF save picker became unavailable")?;
                    let path = pick(&suggested).ok_or("PDF export canceled")?;
                    if path.is_empty() || path.len() > 4096 || path.contains('\0') {
                        return Err("invalid PDF path".into());
                    }
                    if !editor.bridge.print_request_permitted(&app.session, token) {
                        return Err("workbook closed or consent revoked".into());
                    }
                    let d = app.session.document(request.doc).ok_or("workbook closed")?;
                    if params.get("range").is_some() && d.range_epoch != range_epoch {
                        return Err("export range changed during picker".into());
                    }
                    if let Some(id) = sheet {
                        let si = d.wb.sheets.iter().position(|sh| sh.identity.same(&id)).ok_or("export worksheet deleted")?;
                        params["sheet"] = json!(si);
                    }
                    let path = std::path::Path::new(&path).with_extension("pdf").to_string_lossy().into_owned();
                    params["path"] = json!(path);
                    app.session.execute_in(gridcraft_engine::Target::Doc(request.doc), "file.exportPdf", params).map_err(|e| e.to_string())?;
                    Ok(Some(PrintResponse::Done))
                }
                Op::GetPrinter => {
                    if app.printing.settings.printer.is_empty() {
                        return Err("No printer is selected in File > Print.".into());
                    }
                    Ok(Some(PrintResponse::Printer(app.printing.settings.printer.clone())))
                }
                Op::SetPrinter(name) => {
                    crate::printing::settings(&mut app.printing, &json!({"printer":name}), app.services.printers.as_ref())?;
                    Ok(Some(PrintResponse::Done))
                }
                Op::Download { pdf, filename } => {
                    let download = app.services.download.as_ref().ok_or("download service became unavailable")?;
                    download(&filename, &pdf);
                    Ok(Some(PrintResponse::Done))
                }
                Op::Submit { pdf, copies, printer } => {
                    if let Some(submit) = &app.services.print_native {
                        if app.printing.busy() {
                            return Err("A print operation is already active.".into());
                        }
                        let mut settings = app.printing.settings.clone();
                        settings.copies = copies;
                        if let Some(printer) = printer {
                            settings.printer = printer;
                        }
                        if !app.printing.printers.iter().any(|p| p.name == settings.printer) {
                            return Err("Select an available printer in File > Print before submitting.".into());
                        }
                        let receiver = submit(pdf, settings.clone());
                        app.printing.begin(receiver).map_err(|_| "A print operation is already active.".to_owned())?;
                        app.printing.settings = settings;
                        editor.pending_print = Some(PendingPrint::Worker { token });
                        Ok(None)
                    } else {
                        let print = app.services.print.as_ref().ok_or("print service became unavailable")?;
                        print(&pdf);
                        Ok(Some(PrintResponse::Done))
                    }
                }
                Op::Preview { window, sheet } => {
                    if !app.chrome.file_tab || app.session.window(window).is_none_or(|w| w.doc_uid != request.doc) {
                        return Err("preview window is no longer available".into());
                    }
                    let si = sheet
                        .as_ref()
                        .map(|id| {
                            app.session
                                .document(request.doc)
                                .and_then(|d| d.wb.sheets.iter().position(|sh| sh.identity.same(id)))
                                .ok_or("preview worksheet was deleted")
                        })
                        .transpose()?;
                    app.enter_window(window);
                    if let Some(si) = si {
                        app.session.execute("sheet.activate", json!({"sheet":si})).map_err(|e| e.to_string())?;
                    }
                    let mut view = crate::backstage::Backstage::new(crate::backstage::Page::Print);
                    view.print_all = sheet.is_none();
                    app.backstage = Some(view);
                    let editor_visible = editor.visible;
                    editor.visible = false;
                    editor.focus_window = false;
                    editor.pending_print = Some(PendingPrint::Preview { token, editor_visible });
                    Ok(None)
                }
            }
        })();
        match result {
            Ok(Some(reply)) => {
                let _ = editor.bridge.respond_to_print(token, Ok(reply));
            }
            Err(error) => {
                let _ = editor.bridge.respond_to_print(token, Err(error));
            }
            Ok(None) => {}
        }
    }
    if editor.pending_print.is_some() {
        ctx.request_repaint();
    }
    app.vba = Some(editor);
}
fn commit_lifecycle(app: &mut SheetApp, ctx: &egui::Context, doc: u64, id: String, params: Json, save: bool) {
    let quitting = params.get("vbaQuitting").and_then(Json::as_bool) == Some(true);
    let result = match app.session.enter_target(gridcraft_engine::Target::Doc(doc)) {
        Ok(guard) => {
            let result = app.run(&id, params);
            app.session.leave_target(guard);
            result
        }
        Err(e) => Err(e.to_string()),
    };
    if let Some(editor) = app.vba.as_mut() {
        editor.lifecycle_commit = false;
        if !save && result.as_ref().is_ok_and(|v| v.get("dirty").and_then(Json::as_bool) == Some(true)) {
            editor.approved_close = Some(doc);
        }
        editor.lifecycle_result = Some(match &result {
            Ok(_) => json!({"state":"completed","command":id}),
            Err(e) => json!({"state":"failed","command":id,"error":e}),
        });
        if save && result.is_err() && app.session.document(doc).is_some() {
            editor.saved_events.push_back((doc, false));
        }
    }
    if quitting {
        if result.as_ref().is_ok_and(|v| v.get("closed").and_then(Json::as_bool) == Some(true)) {
            app.continue_quit(ctx);
        } else if result.as_ref().is_ok_and(|v| v.get("dirty").and_then(Json::as_bool) == Some(true))
            && let Some(dialog) = app.dialog.as_mut().filter(|d| d.name() == "saveChanges")
        {
            dialog.values.insert("quit".into(), json!(true));
        }
    }
    if let Err(e) = result {
        app.message = Some(("VBA".into(), e));
    }
}
fn finish_native_close(app: &mut SheetApp, ctx: &egui::Context) {
    let close = app.vba.as_mut().and_then(|editor| {
        if editor.external_lifecycle.is_some() || !editor.saved_events.is_empty() || !matches!(editor.ide.mode(), Mode::Design | Mode::Completed) {
            return None;
        }
        editor.close_after_save.take()
    });
    if let Some((doc, quitting)) = close {
        if app.session.document(doc).is_some_and(|d| !d.is_dirty()) {
            if let Ok(guard) = app.session.enter_target(gridcraft_engine::Target::Doc(doc)) {
                let result = app.run("file.close", json!({"force":true,"document":true}));
                app.session.leave_target(guard);
                if let Err(e) = result {
                    app.message = Some(("VBA".into(), e));
                }
            }
            if quitting {
                app.continue_quit(ctx);
            }
        } else if let Some(editor) = app.vba.as_mut() {
            editor.approved_close = None;
        }
    }
}
fn poll_external_lifecycle(editor: &mut Editor, session: &mut gridcraft_engine::Session) -> Option<(u64, String, Json, bool)> {
    let mut pending = editor.external_lifecycle.take()?;
    if !editor.bridge.permitted(session) || session.document(pending.doc).is_none() {
        editor.lifecycle_result = Some(json!({"state":"cancelled","command":pending.id}));
        return None;
    }
    if pending.batch.is_none() {
        if editor.bridge.has_pending_events() || matches!(editor.ide.mode(), Mode::Running | Mode::Waiting | Mode::Break) {
            editor.external_lifecycle = Some(pending);
            return None;
        }
        match editor.bridge.prepare_lifecycle_events(session, &mut editor.ide, pending.doc, pending.event) {
            Ok(batch) => pending.batch = Some(batch),
            Err(e) => {
                editor.lifecycle_result = Some(json!({"state":"failed","command":pending.id,"error":e.description}));
                editor.ide.message(e.description);
                return None;
            }
        }
    }
    let batch = pending.batch.as_mut()?;
    match batch.poll(&mut editor.bridge, session, &mut editor.ide) {
        Ok(None) => editor.external_lifecycle = Some(pending),
        Ok(Some(cancelled)) => {
            if pending.after {
                if !editor.lifecycle_result.as_ref().is_some_and(|r| r.get("state").and_then(Json::as_str) == Some("failed")) {
                    editor.lifecycle_result = Some(json!({"state":"completed","command":pending.id}));
                }
            } else if cancelled {
                editor.lifecycle_result = Some(json!({"state":"cancelled","command":pending.id}));
            } else {
                editor.lifecycle_commit = true;
                return Some((pending.doc, pending.id, pending.params, pending.save));
            }
        }
        Err(e) => {
            editor.ide.message(e.description.clone());
            editor.lifecycle_result = Some(json!({"state":"failed","command":pending.id,"error":e.description}));
        }
    }
    None
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

/// Validate the captured owner and assignment before any linked-cell mutation. No click
/// is queued while a VM is paused/waiting, and the existing bridge alone starts execution.
fn activate_control(app: &mut SheetApp, p: &Json) -> Result<Json> {
    let doc = app.session.doc().map_err(error)?;
    let owner = doc.uid;
    let si = match p.get("sheet") {
        None => doc.wb.active_sheet(),
        Some(Json::String(s)) => doc.wb.sheet_index(s).ok_or("control sheet missing")?,
        Some(n) => n.as_u64().and_then(|n| usize::try_from(n).ok()).ok_or("control sheet index")?,
    };
    let id = p.get("id").and_then(Json::as_u64).and_then(|n| u32::try_from(n).ok()).ok_or("control id")?;
    let sh = doc.wb.sheet(si).ok_or("control sheet missing")?;
    let shape = sh.shapes.iter().find(|s| s.id == id && !s.format.control_inert).ok_or("control missing or inert")?;
    let state = gridcraft_engine::model::form_controls::effective_control(&doc.wb, si, shape)?;
    if !state.enabled || !state.visible || sh.is_protected() {
        return Err("Control is disabled, hidden or protected.".into());
    }
    let identity = shape.identity.clone();
    let assignment = shape.on_action.clone().filter(|s| !s.trim().is_empty());
    if app.vba.as_ref().is_some_and(|e| e.design_mode || matches!(e.ide.mode(), Mode::Running | Mode::Waiting | Mode::Break | Mode::Stale)) {
        return Err("Finish or reset the current macro and leave Design Mode first.".into());
    }
    let entry = if let Some(assignment) = assignment {
        if app.vba.is_none() {
            if doc.wb.vba_project.is_none() && Project::read(&doc.wb).map_err(error)?.is_none() {
                return Err("This workbook has no VBA source for its assigned macro.".into());
            }
            run(app, "vba.open", &json!({}))?;
            if let Some(e) = app.vba.as_mut() {
                e.visible = false;
            }
        }
        let editor = app.vba.as_ref().ok_or("VBA project missing")?;
        if editor.bridge.document() != owner || !editor.bridge.current(&app.session) || !editor.bridge.permitted(&app.session) {
            return Err("VBA consent is required for this workbook.".into());
        }
        let target = assignment.strip_prefix("[0]!").unwrap_or(&assignment);
        let target = if let Some((book, proc)) = target.split_once('!') {
            let book = book.trim_matches('\'').replace("''", "'");
            let d = app.session.doc().map_err(error)?;
            let filename = d.path.as_ref().and_then(|p| std::path::Path::new(p).file_name()).and_then(|s| s.to_str());
            if !filename.is_some_and(|n| n.eq_ignore_ascii_case(&book)) && !d.title.eq_ignore_ascii_case(&book) {
                return Err("Assigned macro belongs to another workbook.".into());
            }
            proc
        } else {
            target
        };
        let project = Project::from_ide(&editor.ide, None).map_err(error)?;
        let entry = if let Some((module, proc)) = target.split_once('.') {
            let selected = project
                .modules
                .iter()
                .find(|(n, _)| n.trim_end_matches(".bas").eq_ignore_ascii_case(module))
                .ok_or("Assigned macro module is missing.")?;
            let local = Project { modules: vec![selected.clone()], toolbar: vec![] };
            if !local.entries().map_err(error)?.iter().any(|e| !e.function && e.name.eq_ignore_ascii_case(proc)) {
                return Err("Assigned procedure is missing from its module.".into());
            }
            proc
        } else {
            target
        };
        if !project.entries().map_err(error)?.iter().any(|e| !e.function && e.name.eq_ignore_ascii_case(entry)) {
            return Err("Select a public Sub with no parameters.".into());
        }
        Some(entry.to_owned())
    } else {
        None
    };
    let d = app.session.doc().map_err(error)?;
    if d.uid != owner || !d.wb.sheet(si).is_some_and(|s| s.shapes.iter().any(|s| s.id == id && identity.same(&s.identity))) {
        return Err("Control owner changed.".into());
    }
    let result = app.run("control.activate", p.clone())?;
    if let Some(entry) = entry {
        run(app, "vba.invoke", &json!({"entry":entry,"controlSheet":si,"controlId":id}))?;
    }
    Ok(result)
}

pub fn assigned(app: &SheetApp, shape: u32) -> Option<String> {
    let doc = app.session.doc().ok()?;
    let sheet = doc.wb.active()?;
    let shape = sheet.shapes.iter().find(|s| s.id == shape && !s.format.control_inert)?;
    let entry = shape.on_action.as_deref()?;
    let entry = entry.strip_prefix("[0]!").unwrap_or(entry);
    (entry.len() <= 256).then(|| entry.to_string())
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

/// Shared editor actions keep the workbook's existing execution-consent policy.
pub fn dispatch_editor_command(editor: &mut Editor, session: &mut gridcraft_engine::Session, command: Command) -> gridcraft_vba::Result<()> {
    match command {
        Command::Action(a) => {
            if a == Action::Reset {
                editor.external_lifecycle = None;
                editor.refedit_picker = None;
                editor.saved_events.clear();
                editor.pending_prompt = None;
                editor.approved_close = None;
                editor.close_after_save = None;
                editor.lifecycle_result = Some(json!({"state":"cancelled"}));
            }
            editor.bridge.dispatch(session, &mut editor.ide, a)
        }
        Command::RunToCursor => editor.bridge.begin(session).and_then(|()| editor.ide.run_to_cursor(&mut editor.bridge.host(session))),
        other => {
            if other.executes_code() {
                editor.bridge.begin(session)?;
            }
            vba_ui_egui::apply(&mut editor.ide, other, &mut editor.bridge.host(session))
        }
    }
}

pub fn show(app: &mut SheetApp, ui: &mut egui::Ui) {
    let Some(mut editor) = app.vba.take() else { return };
    editor.bridge.set_prompt_services(prompt_services(&app.services));
    editor.bridge.set_print_services(print_services(&app.services, app.chrome.file_tab));
    editor.bridge.set_interaction_services(gridcraft_vba::InteractionServices {
        dialogs: true,
        open_picker: app.services.pick_open.is_some(),
        save_picker: app.services.pick_save.is_some(),
        status_visibility: true,
    });
    let mut commands = Vec::new();
    let mut save_workbook = false;
    let mut export_package = false;
    let mut allow = false;
    egui::Panel::top("gridcraft_vba_host").show(ui, |ui| {
        ui.horizontal(|ui| {
            ui.label("GridCraft VBA");
            if ui.button("Save Source to Workbook").clicked() {
                save_workbook = true;
            }
            if ui.button("Export Edited XLSM Copy").clicked() {
                export_package = true;
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
        if !editor.bridge.permitted(&app.session) {
            ui.label("VBA is blocked. Click Allow VBA for this editor session to enable Run and Into.");
        }
        // Host failures can happen before Ide::dispatch, so they only reach the
        // message log. Keep its latest entry visible even when Locals is selected.
        if let Some(message) = editor.ide.messages.last() {
            ui.horizontal_wrapped(|ui| {
                ui.label(format!("VBA: {message}"));
            });
        }
    });
    editor.view.range_selection_enabled = editor.bridge.permitted(&app.session) && editor.bridge.current(&app.session);
    commands.extend(editor.view.show(ui, &editor.ide, host_theme(theme::Tokens::get(ui.ctx()))));
    for command in commands {
        let result = dispatch_editor_command(&mut editor, &mut app.session, command);
        if let Err(e) = result {
            editor.ide.message(e.to_string());
        }
    }
    let mut closed = false;
    for request in editor.ide.take_requests() {
        let result: Result<()> = match request {
            HostRequest::NativeForm { snapshot, save } => {
                editor.ide.validate_native_form_save(&snapshot).map_err(error).and_then(|()| export_native_form(app, &snapshot)).and_then(
                    |committed| {
                        if committed && save {
                            editor.ide.acknowledge_native_form_save(&snapshot).map_err(error)
                        } else {
                            if !committed {
                                editor.ide.message("Native files were offered for download; save remains unacknowledged.");
                            }
                            Ok(())
                        }
                    },
                )
            }
            HostRequest::RenameDocument { .. } => {
                Err("Workbook component identities are host-owned; code-name changes require a project writeback contract.".into())
            }
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
            HostRequest::Import => import_source(app, &mut editor.ide).map(|_| ()),
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
        if export_package {
            app.run_or_alert("vba.exportNativePackage", json!({}));
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
fn import_source(app: &SheetApp, ide: &mut Ide) -> Result<bool> {
    let Some(pick) = &app.services.pick_open else { return Err("Use vba.open with source text to import VBA in this host.".into()) };
    let Some(path) = pick() else { return Ok(false) };
    let bytes = read_source(&path)?;
    let name = std::path::Path::new(&path).file_name().and_then(|n| n.to_str()).ok_or("invalid module name")?;
    let mut files = vec![gridcraft_vba::vba_ide::NativeFormResource { name: name.into(), bytes }];
    if name.to_ascii_lowercase().ends_with(".frm") {
        let file = files.first().ok_or("source missing")?;
        let companion =
            gridcraft_vba::vba_ide::native_form_companion_name(&file.bytes, 1252, &gridcraft_vba::vba_ide::ProjectImportLimits::default())
                .map_err(error)?;
        if let Some(name) = companion {
            let companion_path = std::path::Path::new(&path).with_file_name(&name);
            let bytes = read_source(companion_path.to_str().ok_or("invalid companion filename")?)?;
            files.push(gridcraft_vba::vba_ide::NativeFormResource { name, bytes });
        }
    }
    ide.import_source_bundle(&files, 1252).map_err(error)?;
    Ok(true)
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

fn print_services(services: &crate::services::Services, preview: bool) -> gridcraft_vba::PrintServices {
    gridcraft_vba::PrintServices {
        native: services.print_native.is_some(),
        browser: services.print.is_some(),
        download: services.download.is_some(),
        preview,
        printer: services.print_native.is_some(),
    }
}
fn prompt_services(services: &crate::services::Services) -> PromptServices {
    PromptServices { open: services.pick_open.is_some(), save: services.pick_save.is_some(), close: true }
}
pub fn has_quit_hooks(app: &SheetApp) -> bool {
    app.vba.as_ref().is_some_and(|e| e.bridge.permitted(&app.session) && e.bridge.has_lifecycle_handlers(&e.ide))
}
/// A native quit first closes the explicitly attached project through its ordinary command.
/// Other packages remain inert; they do not acquire a project runtime during shutdown.
pub fn begin_quit(app: &mut SheetApp) -> bool {
    if !has_quit_hooks(app) {
        return false;
    }
    let owner = app.vba.as_ref().map(|e| e.bridge.document());
    if let Some(window) = owner.and_then(|doc| app.session.doc_window(doc)) {
        app.enter_window(window);
    }
    app.run_or_alert("file.close", json!({"document":true,"vbaQuitting":true}));
    true
}
/// Retain the one approved BeforeClose across the native Save Changes dialog.
pub fn answer_native_close(app: &mut SheetApp, answer: crate::SaveAnswer, quitting: bool, ctx: &egui::Context) -> bool {
    let Some(doc) = app.vba.as_ref().and_then(|e| e.approved_close) else {
        return false;
    };
    let Ok(guard) = app.session.enter_target(gridcraft_engine::Target::Doc(doc)) else {
        if let Some(e) = app.vba.as_mut() {
            e.approved_close = None;
            e.close_after_save = None;
        }
        return true;
    };
    match answer {
        crate::SaveAnswer::Cancel => {
            if let Some(e) = app.vba.as_mut() {
                e.approved_close = None;
                e.close_after_save = None;
            }
        }
        crate::SaveAnswer::DontSave => {
            app.run_or_alert("file.close", json!({"force":true,"document":true}));
            if quitting {
                app.continue_quit(ctx);
            }
        }
        crate::SaveAnswer::Save => {
            if let Some(e) = app.vba.as_mut() {
                e.close_after_save = Some((doc, quitting));
            }
            if let Err(error) = app.run("file.save", json!({})) {
                if let Some(e) = app.vba.as_mut() {
                    e.close_after_save = None;
                    e.approved_close = None;
                }
                app.message = Some(("VBA".into(), error));
            }
        }
    }
    app.session.leave_target(guard);
    true
}
pub fn answer_save_prompt(app: &mut SheetApp, answer: crate::SaveAnswer) -> bool {
    let Some(editor) = app.vba.as_mut() else {
        return false;
    };
    let Some(request) = editor.pending_prompt.as_ref().filter(|r| matches!(r.prompt, LifecyclePrompt::SaveChanges { .. })) else {
        return false;
    };
    let response = match answer {
        crate::SaveAnswer::Save => PromptResponse::Save,
        crate::SaveAnswer::DontSave => PromptResponse::Discard,
        crate::SaveAnswer::Cancel => PromptResponse::Cancel,
    };
    if let Err(error) = editor.bridge.respond_to_prompt(request.token, response) {
        editor.ide.message(error.description);
    }
    editor.pending_prompt = None;
    true
}

pub(crate) fn interaction_dialog_closed(app: &mut SheetApp, token: u64, accepted: bool) -> Result<()> {
    interaction::complete(app, token, accepted)
}

/// Render form windows over the workbook while VBE is hidden or RefEdit is collapsed.
pub fn show_form_windows(app: &mut SheetApp, ctx: &egui::Context) {
    let Some(editor) = app.vba.as_mut() else {
        return;
    };
    if editor.visible && editor.ide.form_range_selection().ok().flatten().is_none() {
        return;
    }
    editor.view.range_selection_enabled = editor.bridge.permitted(&app.session) && editor.bridge.current(&app.session);
    let commands = editor.view.show_forms(ctx, &editor.ide, host_theme(theme::Tokens::get(ctx)));
    for command in commands {
        if let Err(e) = dispatch_editor_command(editor, &mut app.session, command) {
            editor.ide.message(e.to_string());
        }
    }
}

/// RefEdit temporarily lends worksheet selection input while retaining modal ownership.
pub(crate) fn selecting_range(app: &SheetApp) -> bool {
    app.vba.as_ref().is_some_and(|e| e.ide.form_range_selection().ok().flatten().is_some())
}

/// Document gestures cross the authorized VBA boundary before egui sees them.
pub(crate) fn document_raw_input(app: &mut SheetApp, ctx: &egui::Context, raw: &mut egui::RawInput) {
    let Some(editor) = app.vba.as_mut() else { return };
    let document = app.session.active_window().unwrap_or(0);
    if editor.bridge.permitted(&app.session) && editor.view.document_input_ready(&editor.ide) {
        editor.visible = false;
    }
    let visible = !editor.visible;
    let area = (editor.bridge.permitted(&app.session)
        && editor.bridge.current(&app.session)
        && app.backstage.is_none()
        && editor.ide.form_range_selection().ok().flatten().is_none())
    .then_some(app.grid.cells_rect)
    .flatten();
    if let Some(command) = editor.view.document_raw_input(ctx, &editor.ide, document, area, visible, raw)
        && let Err(error) = dispatch_editor_command(editor, &mut app.session, command)
    {
        editor.view.cancel_document_input(ctx, raw);
        editor.ide.message(error.to_string());
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[path = "vba_native_files.rs"]
mod native_files;
fn export_native_form(app: &SheetApp, snapshot: &gridcraft_vba::vba_ide::NativeFormSave) -> Result<bool> {
    #[cfg(not(target_arch = "wasm32"))]
    if let Some(pick) = &app.services.pick_save {
        let name = snapshot.files().first().ok_or("Native form source missing")?.name.as_str();
        let Some(path) = pick(name) else { return Ok(false) };
        native_files::write_native_form(std::path::Path::new(&path), snapshot).map_err(error)?;
        return Ok(true);
    }
    if let Some(download) = &app.services.download {
        for file in snapshot.files() {
            download(&file.name, &file.bytes);
        }
        return Ok(false);
    }
    Err("No native form export service is available".into())
}

fn export_native_package_copy(app: &SheetApp) -> Result<Json> {
    let editor = app.vba.as_ref().ok_or("VBA editor is not open")?;
    let snapshot = editor.bridge.prepare_native_project_export(&app.session, &editor.ide).map_err(error)?;
    let bytes = gridcraft_engine::io::encode_edited_vba_package_copy(snapshot.workbook(), snapshot.original_project(), snapshot.encoded_project())
        .map_err(error)?;
    editor.bridge.validate_native_project_export(&app.session, &editor.ide, &snapshot).map_err(error)?;
    #[cfg(not(target_arch = "wasm32"))]
    if let Some(pick) = &app.services.pick_save {
        let Some(path) = pick("EditedWorkbook.xlsm") else {
            return Ok(json!({"cancelled":true}));
        };
        if !std::path::Path::new(&path).extension().is_some_and(|s| s.eq_ignore_ascii_case("xlsm")) {
            return Err("Choose an .xlsm copy".into());
        }
        if let Some(original) = app.session.doc().map_err(error)?.path.as_deref()
            && (path == original || std::fs::canonicalize(&path).ok().zip(std::fs::canonicalize(original).ok()).is_some_and(|(a, b)| a == b))
        {
            return Err("Choose a different path for the edited workbook copy".into());
        }
        editor.bridge.validate_native_project_export(&app.session, &editor.ide, &snapshot).map_err(error)?;
        native_files::write_package(std::path::Path::new(&path), &bytes).map_err(error)?;
        return Ok(json!({"exported":true,"path":path}));
    }
    if let Some(download) = &app.services.download {
        download("EditedWorkbook.xlsm", &bytes);
        return Ok(json!({"downloadRequested":true}));
    }
    Err("No workbook copy export service is available".into())
}
