//! Owned host continuations. Workbook effects occur only after callback/prompt completion.
use super::objects::{arguments, required, string};
use super::{Bridge, HostCall, HostResult, Ide, InvokeKind, Mode, Object, Session, Value, failure, values};
use std::collections::{BTreeMap, VecDeque};
use vba_foreign::{Completion, Reply};
use vba_ide::{HostEventParameter, SuspendedCall};
use vba_values::ScalarKind;
#[derive(Clone, Debug)]
pub enum LifecyclePrompt {
    OpenFile,
    SaveFile { suggested: String },
    SaveChanges { workbook: String },
}
#[derive(Clone, Debug)]
pub struct PromptRequest {
    pub token: u64,
    pub prompt: LifecyclePrompt,
}
#[derive(Clone, Debug)]
pub enum PromptResponse {
    Path(Option<String>),
    Save,
    Discard,
    Cancel,
}
#[derive(Clone, Copy, Debug, Default)]
pub struct PromptServices {
    pub open: bool,
    pub save: bool,
    pub close: bool,
}
#[derive(Default)]
pub(crate) struct Operations {
    next: u64,
    pub pending: BTreeMap<u64, Operation>,
    pub services: PromptServices,
    pub print_services: super::print_api::PrintServices,
}
pub(crate) struct Operation {
    doc: u64,
    actions: VecDeque<Action>,
    callback: Option<HostResult<Completion>>,
    started: bool,
    result: HostResult<Value>,
    cancel: bool,
    window_close: Option<u64>,
}
enum Action {
    Print { request: Option<super::print_api::PrintOperation>, response: Option<HostResult<Value>> },
    Procedure(SuspendedCall),
    Event { event: &'static str, save_ui: bool, success: bool },
    Callbacks { calls: VecDeque<SuspendedCall>, cancel: bool },
    Prompt { kind: LifecyclePrompt, delivered: bool, response: Option<PromptResponse>, call: HostCall },
    PlanClose(HostCall),
    PlanSave(HostCall),
    CommitSave(HostCall),
    CommitClose,
}
fn empty() -> HostResult<Value> {
    Ok(Value::Empty)
}
fn call(member: &str, args: Vec<Value>) -> HostCall {
    HostCall {
        kind: InvokeKind::Call,
        target: super::ObjectRef { host: 0, object: 0, generation: 0 },
        member: member.into(),
        arguments: args.into_iter().map(|value| vba_foreign::HostArgument { name: None, value, by_ref: false }).collect(),
    }
}
impl Bridge {
    pub(crate) fn refresh_lifecycle_handlers(&self, ide: &Ide) {
        self.lifecycle_handlers.store(
            ide.modules.iter().flat_map(|m| m.procedures.iter()).any(|p| {
                let n = p.name.to_ascii_lowercase();
                matches!(n.as_str(), "workbook_beforesave" | "workbook_beforeclose" | "workbook_aftersave")
                    || ["_workbookbeforesave", "_workbookbeforeclose", "_workbookaftersave"].iter().any(|s| n.ends_with(s))
            }),
            std::sync::atomic::Ordering::Relaxed,
        );
    }
    pub fn has_lifecycle_handlers(&self, ide: &Ide) -> bool {
        self.refresh_lifecycle_handlers(ide);
        self.enable_events && self.lifecycle_handlers.load(std::sync::atomic::Ordering::Relaxed)
    }
    pub fn has_pending_events(&self) -> bool {
        !self.pending_events.is_empty()
    }
    pub fn set_print_services(&mut self, services: super::print_api::PrintServices) {
        self.operations.print_services = services;
    }
    pub fn print_request_pending(&self, token: u64) -> bool {
        self.operations.pending.get(&token).is_some_and(|op| matches!(op.actions.front(), Some(Action::Print { response: None, .. })))
    }
    pub fn print_request_permitted(&self, session: &Session, token: u64) -> bool {
        self.permitted(session)
            && self
                .operations
                .pending
                .get(&token)
                .is_some_and(|op| session.document(op.doc).is_some_and(|d| d.trust.allows_code() || d.uid == self.doc && self.consent))
    }
    pub fn take_print_request(&mut self) -> Option<super::print_api::PrintRequest> {
        self.operations.pending.iter_mut().find_map(|(token, op)| {
            if let Some(Action::Print { request, response: None }) = op.actions.front_mut() {
                request.take().map(|kind| super::print_api::PrintRequest { token: *token, doc: op.doc, kind })
            } else {
                None
            }
        })
    }
    pub fn respond_to_print(&mut self, token: u64, response: std::result::Result<super::print_api::PrintResponse, String>) -> HostResult<()> {
        let op = self.operations.pending.get_mut(&token).ok_or_else(|| failure(91, "stale print request"))?;
        let Some(Action::Print { response: slot, .. }) = op.actions.front_mut() else {
            return Err(failure(5, "not waiting for print service"));
        };
        if slot.is_some() {
            return Err(failure(5, "print service already completed"));
        }
        *slot = Some(match response {
            Ok(super::print_api::PrintResponse::Done) => Ok(Value::Empty),
            Ok(super::print_api::PrintResponse::Printer(name)) => {
                if name.len() > 1024 {
                    Err(failure(7, "printer name budget"))
                } else {
                    values::vba_text(&name)
                }
            }
            Err(message) => Err(failure(1004, message)),
        });
        Ok(())
    }
    pub fn set_prompt_services(&mut self, services: PromptServices) {
        self.operations.services = services;
    }
    pub fn take_prompt_request(&mut self) -> Option<PromptRequest> {
        self.operations.pending.iter_mut().find_map(|(token, op)| {
            if let Some(Action::Prompt { kind, delivered, response: None, .. }) = op.actions.front_mut()
                && !*delivered
            {
                *delivered = true;
                Some(PromptRequest { token: *token, prompt: kind.clone() })
            } else {
                None
            }
        })
    }
    pub fn respond_to_prompt(&mut self, token: u64, response: PromptResponse) -> HostResult<()> {
        let op = self.operations.pending.get_mut(&token).ok_or_else(|| failure(91, "stale lifecycle prompt"))?;
        let Some(Action::Prompt { response: slot, .. }) = op.actions.front_mut() else {
            return Err(failure(5, "operation is not waiting for a prompt"));
        };
        if slot.is_some() {
            return Err(failure(5, "prompt already completed"));
        }
        *slot = Some(response);
        Ok(())
    }
    pub(crate) fn window_close_result(&mut self, token: u64, window: u64) {
        if let Some(op) = self.operations.pending.get_mut(&token) {
            op.window_close = Some(window);
        }
    }
    pub(crate) fn cancel_operation(&mut self, token: u64) {
        self.operations.pending.remove(&token);
    }
    fn add_operation(&mut self, doc: u64, actions: VecDeque<Action>) -> HostResult<u64> {
        if self.operations.pending.len() >= 8 {
            return Err(failure(28, "nested host operation depth exceeded"));
        }
        let token = self.operations.next.checked_add(1).ok_or_else(|| failure(7, "host token exhaustion"))?;
        self.operations.next = token;
        self.operations
            .pending
            .insert(token, Operation { doc, actions, callback: None, started: false, result: empty(), cancel: false, window_close: None });
        Ok(token)
    }
    pub(crate) fn defer_call(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> HostResult<Option<u64>> {
        if let Some(request) = self.defer_print_call(session, object, call)? {
            let doc = session.doc().map_err(|e| failure(91, e.to_string()))?.uid;
            return self.add_operation(doc, VecDeque::from([Action::Print { request: Some(request), response: None }])).map(Some);
        }
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let member = call.member.to_ascii_lowercase();
        if *object == Object::Application && member == "run" && matches!(call.kind, InvokeKind::Call | InvokeKind::Get) {
            if call.arguments.iter().any(|a| a.name.is_some()) {
                return Err(failure(448, "Application.Run accepts positional arguments only"));
            }
            let names = std::iter::once("Macro".to_string()).chain((1..=30).map(|n| format!("Arg{n}"))).collect::<Vec<_>>();
            let mut args = arguments(call, &names.iter().map(String::as_str).collect::<Vec<_>>())?;
            let name = string(required(&args, 0)?)?;
            if name.len() > 256 {
                return Err(failure(1004, "macro name budget"));
            }
            let name = if let Some((book, name)) = name.rsplit_once('!') {
                let book = book.trim_matches('\'').replace("''", "'");
                let owner = session.document(self.doc).ok_or_else(|| failure(91, "code owner closed"))?;
                let filename =
                    owner.path.as_deref().and_then(|p| std::path::Path::new(p).file_name()).and_then(|p| p.to_str()).unwrap_or(&owner.title);
                if !book.eq_ignore_ascii_case(filename) && !book.eq_ignore_ascii_case(&owner.title) {
                    return Err(failure(1004, "macro must belong to the running project; another workbook needs its own project host"));
                }
                name.to_owned()
            } else {
                name
            };
            args.remove(0);
            while args.last().is_some_and(|v| matches!(v, Value::Missing)) {
                args.pop();
            }
            return self.add_operation(self.doc, VecDeque::from([Action::Procedure(SuspendedCall::Procedure { name, arguments: args })])).map(Some);
        }
        if !matches!(call.kind, InvokeKind::Call | InvokeKind::Get) {
            return Ok(None);
        }
        let needs_prompt =
            if *object == Object::Workbooks && member == "open" { matches!(super::lifecycle::preflight(call)?[0], Value::Missing) } else { false };
        if needs_prompt {
            if !self.operations.services.open {
                return Err(failure(1004, "workbook open picker is unavailable on this host"));
            }
            return self
                .add_operation(
                    doc.uid,
                    VecDeque::from([Action::Prompt { kind: LifecyclePrompt::OpenFile, delivered: false, response: None, call: call.clone() }]),
                )
                .map(Some);
        }
        if *object != Object::Workbook || !matches!(member.as_str(), "save" | "saveas" | "savecopyas" | "close") {
            return Ok(None);
        }
        let args = super::lifecycle::preflight(call)?;
        if member == "save"
            && let Some(path) = doc.path.as_deref()
        {
            super::lifecycle::workbook_format(path)?;
        }
        let hooks = self.lifecycle_handlers.load(std::sync::atomic::Ordering::Relaxed) && self.enable_events;
        let missing_path = match member.as_str() {
            "save" => doc.path.is_none(),
            "saveas" | "savecopyas" => matches!(args[0], Value::Missing),
            "close" => doc.is_dirty() && matches!(args[0], Value::Missing),
            _ => false,
        };
        if !hooks && !missing_path {
            return Ok(None);
        }
        if !hooks && missing_path && !(if member == "close" { self.operations.services.close } else { self.operations.services.save }) {
            return Ok(None);
        }
        // Existing argument/format validation is repeated by the native projection before commit.
        if member == "save" {
            arguments(call, &[])?;
        }
        let actions = if member == "close" {
            VecDeque::from([Action::Event { event: "BeforeClose", save_ui: false, success: true }, Action::PlanClose(call.clone())])
        } else {
            VecDeque::from([Action::PlanSave(call.clone())])
        };
        self.add_operation(doc.uid, actions).map(Some)
    }
    pub(crate) fn operation_reply(&mut self, token: u64) -> Reply {
        let complete = self.operations.pending.get(&token).is_some_and(|op| op.actions.is_empty() && !op.started);
        if !complete {
            return if self.operations.pending.contains_key(&token) {
                Reply::Wait(token)
            } else {
                Reply::Error(failure(91, "host continuation was cancelled"))
            };
        }
        match self.operations.pending.remove(&token).map(|op| op.result) {
            Some(Ok(value)) => Reply::Complete(Completion { value, writes: vec![] }),
            Some(Err(e)) => Reply::Error(e),
            None => Reply::Error(failure(91, "host continuation vanished")),
        }
    }
    pub(crate) fn drain_host_results(&mut self, session: &mut Session, ide: &mut Ide) -> super::Result<()> {
        while let Some((token, result)) = ide.take_host_call_result() {
            if let Err(error) = &result
                && self.operations.pending.get(&token).is_some_and(|op| matches!(op.actions.front(), Some(Action::Procedure(_))))
            {
                ide.abort_host_wait(token, error.clone(), &mut self.host(session))?;
                self.cancel_pending_events();
                break;
            }
            if let Some(op) = self.operations.pending.get_mut(&token) {
                op.callback = Some(result);
                op.started = false;
            }
        }
        Ok(())
    }
    pub(crate) fn prepare_operation(&mut self, session: &mut Session, ide: &mut Ide) -> super::Result<bool> {
        if ide.mode() != Mode::Waiting {
            return Ok(true);
        }
        let Some(token) = ide.pending_host_token() else {
            return Ok(true);
        };
        let Some(mut op) = self.operations.pending.remove(&token) else {
            return Ok(true);
        };
        let permitted = if op.actions.is_empty() {
            Ok(())
        } else {
            session.document(op.doc).ok_or_else(|| failure(91, "workbook closed during host wait")).and_then(|doc| {
                if doc.trust.allows_code() || (doc.uid == self.doc && self.consent) {
                    Ok(())
                } else {
                    Err(failure(70, "workbook consent revoked during host wait"))
                }
            })
        };
        let result = permitted.and_then(|_| self.advance_operation(session, ide, token, &mut op));
        if let Err(error) = result {
            op.result = Err(error);
            op.actions.clear();
            op.started = false;
        }
        if op.actions.is_empty()
            && (op.result.is_ok() || op.result.as_ref().is_err_and(|e| e.number == 1004 && e.description == "file operation cancelled"))
            && let Some(window) = op.window_close
        {
            op.result = Ok(Value::Boolean(session.window(window).is_none()));
        }
        let ready = op.actions.is_empty() || op.started;
        self.operations.pending.insert(token, op);
        Ok(ready)
    }
    fn advance_operation(&mut self, session: &mut Session, ide: &mut Ide, token: u64, op: &mut Operation) -> HostResult<()> {
        for _ in 0..32 {
            let Some(action) = op.actions.pop_front() else {
                return Ok(());
            };
            match action {
                Action::Print { request, response } => {
                    if let Some(result) = response {
                        op.result = result;
                    } else {
                        op.actions.push_front(Action::Print { request, response: None });
                        return Ok(());
                    }
                }
                Action::Procedure(request) => {
                    if let Some(result) = op.callback.take() {
                        op.result = result.map(|r| r.value);
                        return Ok(());
                    }
                    if !op.started {
                        ide.start_suspended_call(token, request.clone(), &mut self.host(session)).map_err(|e| failure(1004, e.message))?;
                        op.started = true;
                    }
                    op.actions.push_front(Action::Procedure(request));
                    return Ok(());
                }
                Action::Event { event, save_ui, success } => {
                    let calls = self.lifecycle_event_calls(session, ide, op.doc, event, save_ui, success)?;
                    if !calls.is_empty() {
                        op.actions.push_front(Action::Callbacks { calls, cancel: matches!(event, "BeforeSave" | "BeforeClose") });
                    }
                }
                Action::Callbacks { mut calls, cancel } => {
                    if let Some(result) = op.callback.take() {
                        let result = result?;
                        if cancel && let Some(active) = calls.front() {
                            let args = callback_args(active);
                            let index = args.len().saturating_sub(1);
                            if let Some((_, value)) = result.writes.iter().find(|(i, _)| *i == index) {
                                op.cancel = super::address::boolean(value)?;
                            }
                        }
                        calls.pop_front();
                    }
                    let Some(request) = calls.front_mut() else {
                        if cancel && op.cancel {
                            op.actions.clear();
                            op.result = empty();
                        }
                        continue;
                    };
                    if cancel && let Some(last) = callback_args_mut(request).last_mut() {
                        *last = Value::Boolean(op.cancel);
                    }
                    if !op.started {
                        ide.start_suspended_call(token, request.clone(), &mut self.host(session)).map_err(|e| failure(1004, e.message))?;
                        op.started = true;
                    }
                    op.actions.push_front(Action::Callbacks { calls, cancel });
                    return Ok(());
                }
                Action::PlanClose(original) => {
                    if op.window_close.is_some() && session.windows().iter().filter(|w| w.doc_uid == op.doc).count() > 1 {
                        op.actions.push_front(Action::CommitClose);
                        continue;
                    }
                    let args = arguments(&original, &["SaveChanges", "Filename", "RouteWorkbook"])?;
                    if !matches!(args[2], Value::Missing) {
                        return Err(failure(1004, "workbook routing is unavailable"));
                    }
                    let doc = session.document(op.doc).ok_or_else(|| failure(91, "workbook closed"))?;
                    if doc.is_dirty() && matches!(args[0], Value::Missing) {
                        if !self.operations.services.close {
                            return Err(failure(1004, "dirty Close requires SaveChanges or a prompt service"));
                        }
                        op.actions.push_front(Action::Prompt {
                            kind: LifecyclePrompt::SaveChanges { workbook: doc.title.clone() },
                            delivered: false,
                            response: None,
                            call: original,
                        });
                        continue;
                    }
                    op.actions.push_front(Action::CommitClose);
                    if doc.is_dirty() && !matches!(args[0], Value::Missing) && super::address::boolean(&args[0])? {
                        let save = if !matches!(args[1], Value::Missing) { call("SaveAs", vec![args[1].clone()]) } else { call("Save", vec![]) };
                        op.actions.push_front(Action::PlanSave(save));
                    }
                }
                Action::PlanSave(original) => {
                    let doc = session.document(op.doc).ok_or_else(|| failure(91, "workbook closed"))?;
                    let is_save = original.member.eq_ignore_ascii_case("save");
                    let copy = original.member.eq_ignore_ascii_case("savecopyas");
                    let needs_path = if is_save { doc.path.is_none() } else { matches!(super::lifecycle::preflight(&original)?[0], Value::Missing) };
                    let next = if needs_path {
                        if !self.operations.services.save {
                            return Err(failure(1004, "SaveAs requires Filename or a save picker service"));
                        }
                        Action::Prompt {
                            kind: LifecyclePrompt::SaveFile { suggested: format!("{}.xlsx", doc.title) },
                            delivered: false,
                            response: None,
                            call: original,
                        }
                    } else {
                        Action::CommitSave(original)
                    };
                    op.actions.push_front(next);
                    if !copy {
                        op.actions.push_front(Action::Event { event: "BeforeSave", save_ui: needs_path, success: true });
                    }
                }
                Action::Prompt { kind, delivered, response, call: original } => {
                    let Some(response) = response else {
                        op.actions.push_front(Action::Prompt { kind, delivered, response: None, call: original });
                        return Ok(());
                    };
                    match (kind, response) {
                        (LifecyclePrompt::SaveChanges { .. }, choice) => {
                            let mut args = arguments(&original, &["SaveChanges", "Filename", "RouteWorkbook"])?;
                            args[0] = match choice {
                                PromptResponse::Save => Value::Boolean(true),
                                PromptResponse::Discard => Value::Boolean(false),
                                PromptResponse::Cancel => {
                                    op.actions.clear();
                                    return Ok(());
                                }
                                _ => return Err(failure(13, "save decision required")),
                            };
                            op.actions.push_front(Action::PlanClose(call("Close", args)));
                        }
                        (LifecyclePrompt::OpenFile, PromptResponse::Path(Some(path))) => {
                            let mut args = super::lifecycle::preflight(&original)?;
                            args[0] = values::vba_text(&path)?;
                            let original = call("Open", args);
                            super::lifecycle::preflight(&original)?;
                            op.result = session
                                .with_target(gridcraft_engine::Target::Doc(op.doc), |s| {
                                    self.sync_sheets(s)?;
                                    self.lifecycle(s, &Object::Workbooks, &original).ok_or_else(|| failure(438, "Open unavailable"))?
                                })
                                .map_err(|e| failure(91, e.to_string()))?;
                        }
                        (LifecyclePrompt::SaveFile { .. }, PromptResponse::Path(Some(path))) => {
                            let mut original = original;
                            if original.member.eq_ignore_ascii_case("save") {
                                original = call("SaveAs", vec![values::vba_text(&path)?]);
                            } else {
                                let mut args = super::lifecycle::preflight(&original)?;
                                args[0] = values::vba_text(&path)?;
                                original = call(&original.member, args);
                            }
                            super::lifecycle::preflight(&original)?;
                            op.actions.push_front(Action::CommitSave(original));
                        }
                        (_, PromptResponse::Cancel | PromptResponse::Path(None)) => {
                            op.result = Err(failure(1004, "file operation cancelled"));
                            op.actions.clear();
                        }
                        _ => return Err(failure(13, "file picker result required")),
                    }
                }
                Action::CommitSave(original) => {
                    let copy = original.member.eq_ignore_ascii_case("savecopyas");
                    op.result = session
                        .with_target(gridcraft_engine::Target::Doc(op.doc), |s| {
                            self.sync_sheets(s)?;
                            self.lifecycle(s, &Object::Workbook, &original).ok_or_else(|| failure(438, "Save unavailable"))?
                        })
                        .map_err(|e| failure(91, e.to_string()))?;
                    if op.result.is_err() {
                        op.actions.clear();
                    }
                    if !copy {
                        op.actions.push_front(Action::Event { event: "AfterSave", save_ui: false, success: op.result.is_ok() });
                    }
                }
                Action::CommitClose => {
                    // A callback/prompt can create a sibling after the original last-window check.
                    if let Some(window) = op.window_close
                        && session.windows().iter().filter(|w| w.doc_uid == op.doc).count() > 1
                    {
                        op.result = session
                            .execute("window.close", serde_json::json!({"window":window}))
                            .map(|_| Value::Empty)
                            .map_err(|e| failure(1004, e.to_string()));
                        continue;
                    }
                    op.result = session
                        .with_target(gridcraft_engine::Target::Doc(op.doc), |s| {
                            self.sync_sheets(s)?;
                            self.lifecycle(s, &Object::Workbook, &call("Close", vec![Value::Boolean(false)]))
                                .ok_or_else(|| failure(438, "Close unavailable"))?
                        })
                        .map_err(|e| failure(91, e.to_string()))?;
                }
            }
        }
        Err(failure(28, "host continuation phase budget"))
    }
    pub(crate) fn lifecycle_event_calls(
        &mut self,
        session: &mut Session,
        ide: &mut Ide,
        doc: u64,
        event: &'static str,
        save_ui: bool,
        success: bool,
    ) -> HostResult<VecDeque<SuspendedCall>> {
        if !self.enable_events {
            return Ok(VecDeque::new());
        }
        let target = session.document(doc).ok_or_else(|| failure(91, "event workbook closed"))?;
        let code = if doc == self.doc { target.wb.code_name.clone() } else { None };
        let workbook = self.intern(Object::Document(doc, Box::new(Object::Workbook)))?;
        let doc_name = format!("Workbook_{event}");
        let mut calls = VecDeque::new();
        if let Some(code) = code
            && ide.modules.iter().any(|m| {
                m.code_name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(&code))
                    && m.procedures.iter().any(|p| p.name.eq_ignore_ascii_case(&doc_name))
            })
        {
            super::lifecycle_events::validate_document_signature(ide, &code, &doc_name, event)?;
            let arguments = match event {
                "BeforeSave" => vec![Value::Boolean(save_ui), Value::Boolean(false)],
                "AfterSave" => vec![Value::Boolean(success)],
                _ => vec![Value::Boolean(false)],
            };
            let Value::Object(owner) = workbook.clone() else {
                return Err(failure(91, "workbook owner handle"));
            };
            calls.push_back(SuspendedCall::DocumentEvent { code_name: code, handler: doc_name, arguments, owner: Some(owner) });
        }
        let app_event = format!("Workbook{event}");
        let mut args = vec![workbook];
        if event == "BeforeSave" {
            args.push(Value::Boolean(save_ui));
        }
        args.push(Value::Boolean(if event == "AfterSave" { success } else { false }));
        let parameters = args
            .iter()
            .enumerate()
            .map(|(i, _)| HostEventParameter {
                kind: if i == 0 { ScalarKind::Object } else { ScalarKind::Boolean },
                by_ref: event != "AfterSave" && i == args.len() - 1,
                nominal: if i == 0 { vba_values::NominalType::new("Excel", "GridCraft.Excel.v1", "Workbook").ok() } else { None },
            })
            .collect::<Vec<_>>();
        let Value::Object(publisher) = self.intern(Object::Application)? else {
            return Err(failure(91, "application owner handle"));
        };
        let handlers = ide.host_event_handlers(publisher, &app_event, &parameters, &self.host(session)).map_err(|e| failure(1004, e.message))?;
        calls.extend(handlers.into_iter().map(|handler| SuspendedCall::HostEvent { handler, arguments: args.clone() }));
        Ok(calls)
    }
}
pub(crate) fn callback_args(call: &SuspendedCall) -> &[Value] {
    match call {
        SuspendedCall::Procedure { arguments, .. } | SuspendedCall::DocumentEvent { arguments, .. } | SuspendedCall::HostEvent { arguments, .. } => {
            arguments
        }
    }
}
pub(crate) fn callback_args_mut(call: &mut SuspendedCall) -> &mut Vec<Value> {
    match call {
        SuspendedCall::Procedure { arguments, .. } | SuspendedCall::DocumentEvent { arguments, .. } | SuspendedCall::HostEvent { arguments, .. } => {
            arguments
        }
    }
}
