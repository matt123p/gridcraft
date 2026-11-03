//! Precommit event batches for native shell commands; the caller retains commit ownership.
use super::{Bridge, HostResult, Ide, Mode, Session, Value, failure};
use std::collections::VecDeque;
use vba_ide::SuspendedCall;
#[derive(Clone, Copy, Debug)]
pub enum LifecycleEvent {
    BeforeSave { save_as_ui: bool },
    BeforeClose,
    AfterSave { success: bool },
}
pub struct LifecycleEventBatch {
    calls: VecDeque<SuspendedCall>,
    started: bool,
    cancel: bool,
    cancellable: bool,
}
impl LifecycleEventBatch {
    pub fn is_empty(&self) -> bool {
        self.calls.is_empty()
    }
    /// None means the event is paused/waiting/running. A completed result carries cancellation.
    /// Drive an active batch with `Bridge::tick_lifecycle`, collecting outputs here before
    /// permitting ordinary queued mutation callbacks to start.
    pub fn poll(&mut self, bridge: &mut Bridge, session: &mut Session, ide: &mut Ide) -> HostResult<Option<bool>> {
        for _ in 0..16 {
            if self.started {
                match ide.mode() {
                    Mode::Running | Mode::Waiting | Mode::Break | Mode::Stale => return Ok(None),
                    Mode::Failed | Mode::Design => {
                        return Err(failure(1004, "lifecycle event did not complete; native operation remains uncommitted"));
                    }
                    Mode::Completed => {}
                }
                if self.cancellable
                    && let Some(call) = self.calls.front()
                {
                    let index = super::host_ops::callback_args(call).len().saturating_sub(1);
                    if let Some(outputs) = ide.take_callback_outputs()
                        && let Some((_, value)) = outputs.iter().find(|(i, _)| *i == index)
                    {
                        self.cancel = super::address::boolean(value)?;
                    }
                }
                self.calls.pop_front();
                self.started = false;
            }
            let Some(call) = self.calls.front_mut() else {
                return Ok(Some(self.cancel));
            };
            if self.cancellable
                && let Some(last) = super::host_ops::callback_args_mut(call).last_mut()
            {
                *last = Value::Boolean(self.cancel);
            }
            ide.clear_callback_outputs();
            match call.clone() {
                SuspendedCall::DocumentEvent { code_name, handler, arguments, owner } => {
                    ide.dispatch_document_event_for_owner(&code_name, &handler, arguments, owner, &mut bridge.host(session))
                }
                SuspendedCall::HostEvent { handler, arguments } => ide.dispatch_host_event_handler(&handler, arguments, &mut bridge.host(session)),
                SuspendedCall::Procedure { .. } => return Err(failure(1004, "unexpected lifecycle procedure")),
            }
            .map_err(|e| failure(1004, e.message))?;
            self.started = true;
        }
        Ok(None)
    }
}
impl Bridge {
    pub fn prepare_lifecycle_events(
        &mut self,
        session: &mut Session,
        ide: &mut Ide,
        doc: u64,
        event: LifecycleEvent,
    ) -> HostResult<LifecycleEventBatch> {
        self.check(session)?;
        let target = session.document(doc).ok_or_else(|| failure(91, "event workbook closed"))?;
        if !target.trust.allows_code() && !(doc == self.doc && self.consent) {
            return Err(failure(70, "target workbook requires consent before lifecycle callbacks"));
        }
        self.refresh_lifecycle_handlers(ide);
        let (name, save_ui, success, cancellable) = match event {
            LifecycleEvent::BeforeSave { save_as_ui } => ("BeforeSave", save_as_ui, true, true),
            LifecycleEvent::BeforeClose => ("BeforeClose", false, true, true),
            LifecycleEvent::AfterSave { success } => ("AfterSave", false, success, false),
        };
        let calls = if self.enable_events && self.lifecycle_handlers.load(std::sync::atomic::Ordering::Relaxed) {
            ide.prepare_host_events(&mut self.host(session)).map_err(|e| failure(1004, e.message))?;
            self.ensure_group(session, doc).map_err(|e| failure(7, e.to_string()))?;
            self.lifecycle_event_calls(session, ide, doc, name, save_ui, success)?
        } else {
            VecDeque::new()
        };
        Ok(LifecycleEventBatch { calls, started: false, cancel: false, cancellable })
    }
}
pub(crate) fn validate_document_signature(ide: &Ide, code: &str, handler: &str, event: &str) -> HostResult<()> {
    let module = ide
        .modules
        .iter()
        .find(|m| m.code_name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(code)))
        .ok_or_else(|| failure(91, "event module unavailable"))?;
    let source = module.source().map_err(|e| failure(1004, e.message))?;
    let parsed = vba_syntax::parse(&source, vba_syntax::ModuleKind::Document).map_err(|e| failure(1004, e.message))?;
    let procedure = parsed
        .module
        .procedures
        .iter()
        .find(|p| p.name.spelling.eq_ignore_ascii_case(handler))
        .ok_or_else(|| failure(1004, "event handler unavailable"))?;
    let expected = if event == "BeforeSave" { 2 } else { 1 };
    if procedure.parameters.len() != expected || procedure.kind != vba_syntax::ProcedureKind::Sub {
        return Err(failure(1004, "invalid workbook event signature"));
    }
    for (i, p) in procedure.parameters.iter().enumerate() {
        let by_val = event == "AfterSave" || event == "BeforeSave" && i == 0;
        if p.by_val != by_val
            || p.array
            || p.optional
            || p.param_array
            || p.ty.as_ref().is_some_and(|ty| !matches!(ty.spelling.to_ascii_lowercase().as_str(), "boolean" | "variant"))
        {
            return Err(failure(1004, "invalid workbook event parameter signature"));
        }
    }
    Ok(())
}
