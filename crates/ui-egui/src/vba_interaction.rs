//! UI glue for existing dialogs and file pickers; never owns an engine borrow across frames.
use super::{Json, Mode, Result, json};
use crate::SheetApp;
use gridcraft_vba::{InteractionKind as Kind, InteractionRequest, InteractionResponse as Response};
pub(super) struct Pending {
    request: InteractionRequest,
    editor_visible: bool,
    command: &'static str,
}
pub(super) fn before_command(app: &SheetApp, id: &str, p: &Json) -> Result<()> {
    let Some(e) = app.vba.as_ref() else {
        return Ok(());
    };
    let Some(wait) = e.pending_interaction.as_ref() else {
        return Ok(());
    };
    if matches!(id, "vba.reset" | "vba.close" | "ui.dialog.cancel") {
        return Ok(());
    }
    if !e.bridge.interaction_permitted(&app.session, &wait.request) {
        return Err("Interaction owner changed or consent was revoked; cancel the dialog or reset VBA.".into());
    }
    if id == "ui.dialog.set" && p.get("field").and_then(Json::as_str) == Some("om15Token") {
        return Err("Interaction ownership cannot be edited.".into());
    }
    if matches!(id, "ui.dialog.set" | "ui.dialog.confirm" | "vba.status")
        || id == wait.command
        || wait.command == "view.zoom" && id == "view.zoomToSelection"
    {
        return Ok(());
    }
    Err("A VBA dialog is waiting for its result; complete or cancel it before running another command.".into())
}
pub(super) fn complete(app: &mut SheetApp, token: u64, accepted: bool) -> Result<()> {
    let e = app.vba.as_mut().ok_or("VBA editor closed")?;
    if e.pending_interaction.as_ref().is_none_or(|p| p.request.token != token) {
        return Err("stale dialog completion".into());
    }
    e.bridge.respond_to_interaction(token, Ok(Response::Accepted(accepted))).map_err(|e| e.description)
}
pub(super) fn poll(app: &mut SheetApp, ctx: &egui::Context) {
    let Some(mut e) = app.vba.take() else {
        return;
    };
    if let Some(wait) = e.pending_interaction.take() {
        let token = wait.request.token;
        let editor_visible = wait.editor_visible;
        let owned = app.dialog.as_ref().is_some_and(|d| d.values.get("om15Token").and_then(Json::as_u64) == Some(token));
        if e.bridge.interaction_pending(token) {
            if !e.bridge.interaction_permitted(&app.session, &wait.request) {
                let _ = e.bridge.respond_to_interaction(token, Err("Interaction owner changed or consent was revoked.".into()));
            } else if !owned {
                let _ = e.bridge.respond_to_interaction(token, Ok(Response::Accepted(false)));
            } else {
                e.pending_interaction = Some(wait);
            }
        }
        if e.pending_interaction.is_none() {
            if owned {
                app.dialog = None;
            }
            e.visible = editor_visible || matches!(e.ide.mode(), Mode::Break | Mode::Failed);
            e.focus_window = e.visible;
        }
    }
    if e.pending_interaction.is_none()
        && let Some(request) = e.bridge.take_interaction_request()
    {
        let token = request.token;
        let result = if !e.bridge.interaction_permitted(&app.session, &request) {
            Some(Err("Interaction owner changed or consent was revoked.".into()))
        } else {
            match &request.kind {
                Kind::OpenFilename | Kind::FindFile => Some(match app.services.pick_open.as_ref() {
                    Some(f) => Ok(Response::Path(f())),
                    None => Err("Open picker service disconnected.".into()),
                }),
                Kind::SaveFilename { suggested } => Some(match app.services.pick_save.as_ref() {
                    Some(f) => Ok(Response::Path(f(suggested))),
                    None => Err("Save picker service disconnected.".into()),
                }),
                Kind::StatusVisibility { value } => {
                    if let Some(v) = value {
                        match app.run("view.statusBarVisible", json!({"on":v})) {
                            Ok(_) => Some(Ok(Response::Visible(app.ui.status_bar))),
                            Err(e) => Some(Err(e)),
                        }
                    } else {
                        Some(Ok(Response::Visible(app.ui.status_bar)))
                    }
                }
                Kind::Dialog { name } => {
                    if app.dialog.is_some() || app.backstage.is_some() {
                        Some(Err("An existing modal or Backstage view owns the UI.".into()))
                    } else {
                        let command = match name.as_str() {
                            "columnWidth" => "home.columnWidth",
                            "rowHeight" => "home.rowHeight",
                            "zoom" => "view.zoom",
                            "insertCells" => "home.insertCells",
                            "deleteCells" => "home.deleteCells",
                            "defineName" => "formulas.defineName",
                            "insertLink" => "insert.link",
                            _ => "",
                        };
                        app.follow_active_window();
                        app.open_dialog(name, json!({}));
                        if let Some(d) = app.dialog.as_mut().filter(|d| d.name() == name && d.command == Some(command)) {
                            d.values.insert("om15Token".into(), json!(token));
                            e.pending_interaction = Some(Pending { request: request.clone(), editor_visible: e.visible, command });
                            e.visible = false;
                            None
                        } else {
                            Some(Err("Native dialog could not be opened.".into()))
                        }
                    }
                }
            }
        };
        if let Some(result) = result {
            let _ = e.bridge.respond_to_interaction(token, result);
        }
        ctx.request_repaint();
    }
    app.vba = Some(e);
}
