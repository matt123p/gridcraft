//! Fork-owned Excel object profile over the separate tiny-vba interpreter.
//! Host calls are native typed Rust; JSON carries workbook metadata and command envelopes.
#![forbid(unsafe_code)]
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
mod address;
mod formatting;
mod objects;
mod ranges;
mod references;
mod search;
mod structure;
mod types;
mod values;
pub mod worksheet;

use gridcraft_engine::core::RangeRef;
use gridcraft_engine::{GroupToken, Session};
use objects::Object;
use std::collections::{BTreeMap, VecDeque};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, RwLock};
use vba_foreign::{Completion, ForeignError, HostCall, InvokeKind, Reply};
pub use vba_ide::{self, Action, Host, Ide, Mode, ObjectRef, Value};
use vba_ide::{ApiMember, Diagnostic, Import};

pub type Result<T> = std::result::Result<T, Diagnostic>;
type HostResult<T> = std::result::Result<T, ForeignError>;
static NEXT_HOST: AtomicU64 = AtomicU64::new(1);
pub const MAX_OBJECTS: usize = 8192;
pub const MAX_CELLS: usize = 65_536;
const MAX_QUEUED_EVENTS: usize = 256;
pub const DEMO: &str = "Option Explicit\nPublic Sub Main()\n    Dim i As Long\n    For i = 1 To 10\n        Cells(i, 1).Value2 = i\n    Next i\n    Range(\"B1\").Formula = \"=SUM(A1:A10)\"\nEnd Sub\n";

fn diagnostic(e: impl std::fmt::Display) -> Diagnostic {
    Diagnostic::new("GRIDCRAFT_VBA", e.to_string())
}
fn failure(number: i32, text: impl Into<String>) -> ForeignError {
    ForeignError { number, description: text.into() }
}

/// State persists across bounded IDE ticks, but never holds a document borrow.
pub struct Bridge {
    identity: u64,
    pub(crate) doc: u64,
    sheets: BTreeMap<u64, Vec<gridcraft_engine::model::sheet::SheetIdentity>>,
    context: u64,
    range_epochs: BTreeMap<u64, u64>,
    objects: Arc<types::Objects>,
    group: BTreeMap<u64, GroupToken>,
    consent: bool,
    enable_events: bool,
    pending_events: VecDeque<(String, String, Vec<Value>)>,
    event_work: usize,
    pub(crate) event_owner_sheet: Option<usize>,
    document_code_names: Arc<RwLock<std::collections::BTreeSet<String>>>,
    closed: bool,
    find: search::FindState,
    scan_work: usize,
    match_work: usize,
}
impl Bridge {
    pub fn new(session: &Session) -> Result<Self> {
        let doc = session.doc().map_err(diagnostic)?;
        #[allow(deprecated)] // fetch_update is available at the workspace MSRV.
        let identity =
            NEXT_HOST.fetch_update(Ordering::Relaxed, Ordering::Relaxed, |n| n.checked_add(1)).map_err(|_| diagnostic("host identity exhausted"))?;
        Ok(Self {
            identity,
            doc: doc.uid,
            sheets: BTreeMap::from([(doc.uid, doc.wb.sheets.iter().map(|s| s.identity.clone()).collect())]),
            context: doc.uid,
            range_epochs: BTreeMap::from([(doc.uid, doc.range_epoch)]),
            objects: Arc::new(types::Objects::new(identity, doc.uid, doc.sheet_index())),
            group: BTreeMap::new(),
            consent: false,
            enable_events: true,
            pending_events: VecDeque::new(),
            event_work: 0,
            event_owner_sheet: None,
            document_code_names: Arc::new(RwLock::new(std::collections::BTreeSet::new())),
            closed: false,
            find: search::FindState::default(),
            scan_work: 0,
            match_work: 0,
        })
    }
    pub fn document(&self) -> u64 {
        self.doc
    }
    pub fn permitted(&self, session: &Session) -> bool {
        self.check(session).is_ok()
    }
    pub fn current(&self, session: &Session) -> bool {
        !self.closed && session.doc().is_ok_and(|doc| doc.uid == self.doc)
    }
    pub(crate) fn check(&self, session: &Session) -> HostResult<()> {
        if self.closed {
            return Err(failure(91, "VBA host is closed"));
        }
        let doc = session.document(self.doc).ok_or_else(|| failure(91, "workbook closed"))?;
        if doc.uid != self.doc {
            return Err(failure(91, "workbook or worksheet topology changed; reset the editor"));
        }
        if !doc.trust.allows_code() && !self.consent {
            return Err(failure(70, "VBA requires explicit consent for this workbook"));
        }
        Ok(())
    }
    fn sync_sheets(&mut self, session: &Session) -> HostResult<()> {
        let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
        if self.range_epochs.get(&doc.uid).is_some_and(|epoch| *epoch != doc.range_epoch) {
            self.objects.invalidate_ranges(doc.uid)?;
        }
        self.range_epochs.insert(doc.uid, doc.range_epoch);
        if doc.wb.sheets.len() > MAX_OBJECTS {
            return Err(failure(7, "worksheet identity budget"));
        }
        let current: Vec<_> = doc.wb.sheets.iter().map(|s| s.identity.clone()).collect();
        if self.sheets.get(&doc.uid).is_some_and(|old| old.len() == current.len() && old.iter().zip(&current).all(|(a, b)| a.same(b))) {
            self.context = doc.uid;
            return Ok(());
        }
        let mapping: Vec<_> = self.sheets.get(&doc.uid).into_iter().flatten().map(|old| current.iter().position(|id| id.same(old))).collect();
        self.objects.remap_sheets(doc.uid, &mapping)?;
        if let Some(owner) = self.event_owner_sheet {
            self.event_owner_sheet = Some(mapping.get(owner).copied().flatten().unwrap_or(usize::MAX));
        }
        self.context = doc.uid;
        self.sheets.insert(doc.uid, current);
        Ok(())
    }
    fn ensure_group(&mut self, session: &mut Session, doc: u64) -> Result<()> {
        if !self.group.contains_key(&doc) {
            if self.group.len() >= 64 {
                return Err(diagnostic("workbook group budget"));
            }
            let token =
                session.with_target(gridcraft_engine::Target::Doc(doc), |s| s.group_begin("VBA Macro")).map_err(diagnostic)?.map_err(diagnostic)?;
            self.group.insert(doc, token);
        }
        Ok(())
    }
    fn object(&self, session: &Session, reference: ObjectRef) -> HostResult<Object> {
        match self.objects.get(reference)? {
            Object::Document(doc, object) if session.doc().is_ok_and(|d| d.uid == doc) => Ok(*object),
            Object::Document(..) => Err(failure(1004, "object belongs to another workbook")),
            object => Ok(object),
        }
    }
    fn reference(&self, object: usize) -> ObjectRef {
        ObjectRef { host: self.identity, object: object as u64, generation: self.doc }
    }
    pub(crate) fn intern(&mut self, object: Object) -> HostResult<Value> {
        if let Object::Range(_, range) = &object {
            range.validate()?;
        }
        let object = if matches!(object, Object::Application) { object } else { Object::Document(self.context, Box::new(object)) };
        let index = self.objects.intern(object)?;
        Ok(Value::Object(self.reference(index)))
    }
    pub fn configure(&self, ide: &mut Ide) -> Result<()> {
        if let Ok(mut names) = self.document_code_names.write() {
            names.clear();
            names.extend(ide.modules.iter().filter_map(|module| {
                (module.kind == vba_ide::ModuleKind::Document).then(|| module.code_name.clone()).flatten().map(|name| name.to_ascii_lowercase())
            }));
        }
        ide.configure_host(self.imports(), self.roots(), catalog())?;
        ide.configure_reference_bindings(types::references()?)?;
        ide.configure_type_resolver(Some(self.objects.clone()))
    }
    pub(crate) fn roots(&self) -> BTreeMap<String, ObjectRef> {
        BTreeMap::from([
            ("Excel".into(), self.reference(0)),
            ("Workbook".into(), self.reference(1)),
            ("Worksheet".into(), self.reference(2)),
            ("Worksheets".into(), self.reference(3)),
            ("Sheets".into(), self.reference(4)),
            ("Cells".into(), self.reference(5)),
            ("Rows".into(), self.reference(6)),
            ("Columns".into(), self.reference(7)),
        ])
    }
    pub(crate) fn imports(&self) -> Vec<Import> {
        let mut imports = vec![
            Import { name: "Application".into(), root: "Excel".into(), member: None, kind: InvokeKind::Get },
            Import { name: "ThisWorkbook".into(), root: "Workbook".into(), member: None, kind: InvokeKind::Get },
        ];
        for name in ["ActiveWorkbook", "ActiveSheet", "ActiveCell", "Selection", "Worksheets", "Sheets", "Range", "Cells", "Rows", "Columns"] {
            imports.push(Import {
                name: name.into(),
                root: "Excel".into(),
                member: Some(name.into()),
                kind: if name == "Range" { InvokeKind::Call } else { InvokeKind::Get },
            });
        }
        imports
    }
    /// Consent applies only to this editor/document instance. It never changes provenance.
    pub fn consent(&mut self, allow: bool) {
        self.consent = allow;
    }
    pub fn begin(&mut self, session: &mut Session) -> Result<()> {
        self.check(session).map_err(|e| diagnostic(e.description))?;
        if self.group.is_empty() {
            self.scan_work = 0;
            self.match_work = 0;
            if self.pending_events.is_empty() {
                self.event_work = 0;
            }
        }
        self.ensure_group(session, self.doc)?;
        Ok(())
    }
    fn cancel_pending_events(&mut self) {
        self.pending_events.clear();
        self.event_work = 0;
        self.event_owner_sheet = None;
    }
    fn reset_event_budget_if_idle(&mut self) {
        if self.group.is_empty() && self.pending_events.is_empty() {
            self.event_work = 0;
        }
    }
    fn change_owner(&self, session: &Session, sheet: usize) -> HostResult<Option<String>> {
        if !self.enable_events {
            return Ok(None);
        }
        let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
        let Some(code_name) = doc.wb.sheet(sheet).and_then(|sheet| sheet.code_name.as_deref()) else {
            return Ok(None);
        };
        let known = self.document_code_names.read().map_err(|_| failure(91, "document code-name registry unavailable"))?;
        Ok(known.contains(&code_name.to_ascii_lowercase()).then(|| code_name.to_string()))
    }
    fn document_owner_object(&mut self, session: &Session, code_name: &str) -> HostResult<Value> {
        let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
        if doc.uid != self.doc {
            return Err(failure(91, "document event owner is not in the bound workbook"));
        }
        self.context = doc.uid;
        let object = if doc.wb.code_name.as_deref().is_some_and(|name| name.eq_ignore_ascii_case(code_name)) {
            Object::Workbook
        } else {
            let sheet = doc
                .wb
                .sheets
                .iter()
                .position(|sheet| sheet.code_name.as_deref().is_some_and(|name| name.eq_ignore_ascii_case(code_name)))
                .ok_or_else(|| failure(91, "document code name no longer identifies a workbook object"))?;
            Object::Sheet(sheet)
        };
        self.intern(object)
    }
    fn preflight_change(&mut self, session: &Session, sheet: usize) -> HostResult<Option<String>> {
        let Some(code_name) = self.change_owner(session, sheet)? else {
            return Ok(None);
        };
        self.reset_event_budget_if_idle();
        if self.event_work >= MAX_QUEUED_EVENTS || self.pending_events.len() >= MAX_QUEUED_EVENTS {
            return Err(failure(28, "worksheet event budget exceeded"));
        }
        Ok(Some(code_name))
    }
    fn queue_change(&mut self, code_name: String, sheet: usize, range: ranges::Range) -> HostResult<()> {
        if self.group.is_empty() && self.pending_events.is_empty() {
            self.event_work = 0;
        }
        if self.pending_events.len() >= MAX_QUEUED_EVENTS || self.event_work >= MAX_QUEUED_EVENTS {
            return Err(failure(28, "worksheet event budget exceeded"));
        }
        let target = self.intern(Object::Range(sheet, range))?;
        self.pending_events.try_reserve(1).map_err(|_| failure(7, "worksheet event queue allocation"))?;
        self.pending_events.push_back((code_name, "Worksheet_Change".into(), vec![target]));
        self.event_work = self.event_work.saturating_add(1);
        Ok(())
    }
    pub fn notify_sheet_change(&mut self, session: &Session, sheet: usize, range: RangeRef) -> Result<()> {
        if !self.current(session) || !self.permitted(session) {
            return Ok(());
        }
        let Some(code_name) = self.preflight_change(session, sheet).map_err(|error| diagnostic(error.description))? else {
            return Ok(());
        };
        self.queue_change(code_name, sheet, range.into()).map_err(|error| diagnostic(error.description))
    }
    /// Queue Workbook_Open only after the embedding has explicitly authorized execution.
    /// Opening or decoding a package never calls this method automatically.
    pub fn queue_workbook_open(&mut self, session: &Session) -> Result<()> {
        if !self.current(session) {
            return Err(diagnostic("Workbook_Open can only target the bound workbook"));
        }
        self.check(session).map_err(|error| diagnostic(error.description))?;
        let doc = session.doc().map_err(diagnostic)?;
        let code_name = doc.wb.code_name.as_deref().ok_or_else(|| diagnostic("workbook has no VBA code name"))?;
        let registered = self
            .document_code_names
            .read()
            .map_err(|_| diagnostic("document code-name registry unavailable"))?
            .contains(&code_name.to_ascii_lowercase());
        if !registered {
            return Err(diagnostic("workbook code name has no imported document module"));
        }
        self.reset_event_budget_if_idle();
        if self.pending_events.len() >= MAX_QUEUED_EVENTS || self.event_work >= MAX_QUEUED_EVENTS {
            return Err(diagnostic("document event budget exceeded"));
        }
        self.pending_events.try_reserve(1).map_err(|_| diagnostic("document event queue allocation"))?;
        self.pending_events.push_back((code_name.into(), "Workbook_Open".into(), vec![]));
        self.event_work = self.event_work.saturating_add(1);
        Ok(())
    }
    /// Decode VBA source without executing it or modifying the package. Document code names
    /// must match the workbook/sheet identities preserved by XLSX.
    pub fn import_project(&self, session: &Session) -> Result<Ide> {
        let doc = session.doc().map_err(diagnostic)?;
        if doc.uid != self.doc {
            return Err(diagnostic("workbook changed"));
        }
        let bytes = doc.wb.vba_project.as_ref().ok_or_else(|| diagnostic("workbook has no VBA project"))?;
        let project = vba_project::decode(bytes, &vba_project::Limits::default())?;
        let mut ide = Ide::new();
        ide.title = format!("{} — VBA source", doc.title);
        for module in project.modules {
            match module.kind {
                vba_project::ModuleKind::Standard => {
                    ide.add_module(&format!("{}.bas", module.name), &module.source)?;
                }
                vba_project::ModuleKind::Class => {
                    ide.add_module_with_kind(&format!("{}.cls", module.name), &module.source, vba_ide::ModuleKind::Class)?;
                }
                vba_project::ModuleKind::Document => {
                    let code_name = module.code_name.as_deref().ok_or_else(|| diagnostic("document module has no code name"))?;
                    let owns_workbook = doc.wb.code_name.as_deref().is_some_and(|name| name.eq_ignore_ascii_case(code_name));
                    let owns_sheet =
                        doc.wb.sheets.iter().any(|sheet| sheet.code_name.as_deref().is_some_and(|name| name.eq_ignore_ascii_case(code_name)));
                    if !owns_workbook && !owns_sheet {
                        return Err(diagnostic(format!("{}: code name does not match workbook metadata", module.name)));
                    }
                    ide.add_document_module(&format!("{}.cls", module.name), code_name, &module.source)?;
                }
            }
        }
        self.configure(&mut ide)?;
        Ok(ide)
    }
    pub fn host<'a>(&'a mut self, session: &'a mut Session) -> WorkbookHost<'a> {
        WorkbookHost { bridge: self, session }
    }
    pub fn dispatch(&mut self, session: &mut Session, ide: &mut Ide, action: Action) -> Result<()> {
        let executes = matches!(action, Action::Run | Action::StepInto | Action::StepOver | Action::StepOut);
        if executes {
            self.begin(session)?;
        }
        if matches!(action, Action::Reset) {
            self.cancel_pending_events();
        }
        let result = ide.dispatch(action, &mut self.host(session));
        if result.is_err() || matches!(action, Action::Reset) {
            self.finish(session)?;
        }
        result
    }
    pub fn tick(&mut self, session: &mut Session, ide: &mut Ide) -> Result<()> {
        let active = matches!(ide.mode(), Mode::Running | Mode::Waiting | Mode::Break | Mode::Stale)
            || matches!(ide.mode(), Mode::Completed | Mode::Design) && !self.pending_events.is_empty();
        let validation = if active {
            if session.document(self.doc).is_none() { Err(failure(91, "workbook closed before VBA event delivery")) } else { self.check(session) }
        } else {
            Ok(())
        };
        if let Err(e) = validation {
            // Cancel while the original host identity is still available; no side effects replay.
            let _ = ide.dispatch(Action::Reset, &mut self.host(session));
            self.cancel_pending_events();
            self.finish(session)?;
            self.closed = true;
            self.objects.close();
            return Err(diagnostic(e.description));
        }
        let result = ide.tick(&mut self.host(session));
        if result.is_ok() && matches!(ide.mode(), Mode::Completed | Mode::Design) {
            if let Some((code_name, handler, arguments)) = self.pending_events.pop_front() {
                self.event_owner_sheet = session.doc().ok().and_then(|doc| {
                    doc.wb.sheets.iter().position(|sheet| sheet.code_name.as_deref().is_some_and(|name| name.eq_ignore_ascii_case(&code_name)))
                });
                if self.group.is_empty() {
                    self.ensure_group(session, self.doc)?;
                }
                let owner = self.document_owner_object(session, &code_name).map_err(|error| diagnostic(error.description))?;
                let Value::Object(owner) = owner else { return Err(diagnostic("document owner handle is not an object")) };
                if let Err(error) = ide.dispatch_document_event_for_owner(&code_name, &handler, arguments, Some(owner), &mut self.host(session)) {
                    self.cancel_pending_events();
                    self.finish(session)?;
                    return Err(error);
                }
            }
        }
        if result.is_err() || matches!(ide.mode(), Mode::Failed) {
            self.cancel_pending_events();
        }
        if result.is_err()
            || matches!(ide.mode(), Mode::Failed | Mode::Design)
            || matches!(ide.mode(), Mode::Completed) && self.pending_events.is_empty()
        {
            if self.pending_events.is_empty() {
                self.event_owner_sheet = None;
            }
            self.finish(session)?;
        }
        result
    }
    /// Retain successful writes on failure/stop and make them one ordinary undo step.
    pub fn finish(&mut self, session: &mut Session) -> Result<()> {
        for (doc, token) in std::mem::take(&mut self.group) {
            if session.document_index(doc).is_some() {
                session.group_end(token).map_err(diagnostic)?;
            }
        }
        Ok(())
    }
    pub fn close(&mut self, session: &mut Session, ide: &mut Ide, discard: bool) -> Result<()> {
        ide.close(&mut self.host(session), discard)?;
        self.cancel_pending_events();
        self.finish(session)?;
        self.closed = true;
        self.objects.close();
        Ok(())
    }
}
pub struct WorkbookHost<'a> {
    bridge: &'a mut Bridge,
    session: &'a mut Session,
}
impl Drop for Bridge {
    fn drop(&mut self) {
        self.objects.close();
    }
}
impl Host for WorkbookHost<'_> {
    fn identity(&self) -> u64 {
        self.bridge.identity
    }
    fn invoke(&mut self, call: HostCall) -> Reply {
        let result = (|| {
            self.bridge.check(self.session)?;
            if call.target.host != self.bridge.identity || call.target.generation != self.bridge.doc {
                return Err(failure(91, "stale or foreign object reference"));
            }
            self.bridge.sync_sheets(self.session)?;
            let object = self.bridge.objects.get(call.target)?;
            let (doc, object) = match object {
                Object::Document(doc, object) => (doc, *object),
                object => (self.session.doc().map_err(|_| failure(91, "workbook closed"))?.uid, object),
            };
            let mutation = match &object {
                Object::Range(sheet, range)
                    if matches!(call.kind, InvokeKind::Let | InvokeKind::Set)
                        && matches!(call.member.to_ascii_lowercase().as_str(), "value" | "value2" | "formula") =>
                {
                    self.bridge.preflight_change(self.session, *sheet)?.map(|code_name| (code_name, *sheet, range.clone()))
                }
                Object::Range(sheet, range) if call.kind == InvokeKind::Call && call.member.eq_ignore_ascii_case("ClearContents") => {
                    self.bridge.preflight_change(self.session, *sheet)?.map(|code_name| (code_name, *sheet, range.clone()))
                }
                _ => None,
            };
            let target = self.session.document(doc).ok_or_else(|| failure(91, "workbook closed"))?;
            if !target.trust.allows_code() && !(doc == self.bridge.doc && self.bridge.consent) {
                return Err(failure(70, "target workbook requires consent"));
            }
            self.bridge.ensure_group(self.session, doc).map_err(|e| failure(7, e.to_string()))?;
            let result = self
                .session
                .with_target(gridcraft_engine::Target::Doc(doc), |session| {
                    self.bridge.sync_sheets(session)?;
                    // Resolve again after sheet topology synchronization.
                    let object = match self.bridge.objects.get(call.target)? {
                        Object::Document(_, object) => *object,
                        _ => object,
                    };
                    self.bridge.invoke_object(session, object, call)
                })
                .map_err(|e| failure(91, e.to_string()))?;
            self.bridge.sync_sheets(self.session)?;
            if result.is_ok()
                && let Some((code_name, sheet, range)) = mutation
            {
                self.bridge.queue_change(code_name, sheet, range)?;
            }
            result
        })();
        match result {
            Ok(value) => Reply::Complete(Completion { value, writes: vec![] }),
            Err(e) => Reply::Error(e),
        }
    }
    fn resume(&mut self, _: u64) -> Reply {
        Reply::Error(failure(5, "host has no suspended operations"))
    }
    fn cancel(&mut self, _: u64) {}
}

pub fn catalog() -> Vec<ApiMember> {
    [
        ("Application", "ActiveCell", "Property Get ActiveCell() As Range"),
        ("Application", "Selection", "Property Get Selection() As Range"),
        ("Application", "Calculation", "Property Calculation As XlCalculation"),
        ("Application", "EnableEvents", "Property EnableEvents As Boolean"),
        ("Application", "Calculate", "Sub Calculate()"),
        ("Workbook", "Calculate", "Sub Calculate()"),
        ("Worksheet", "Name", "Property Name As String"),
        ("Range", "NumberFormat", "Property NumberFormat As Variant"),
        ("Range", "Font", "Property Get Font() As Font"),
        ("Range", "Interior", "Property Get Interior() As Interior"),
        ("Font", "Bold", "Property Bold As Variant"),
        ("Font", "Italic", "Property Italic As Variant"),
        ("Font", "Name", "Property Name As Variant"),
        ("Font", "Size", "Property Size As Variant"),
        ("Font", "Color", "Property Color As Variant"),
        ("Interior", "Color", "Property Color As Variant"),
        ("Application", "ThisWorkbook", "Property Get ThisWorkbook() As Workbook"),
        ("Application", "ActiveWorkbook", "Property Get ActiveWorkbook() As Workbook"),
        ("Application", "Range", "Property Get Range(Cell1 As Variant, Optional Cell2 As Variant) As Range"),
        ("Application", "Cells", "Property Get Cells(Optional RowIndex As Variant, Optional ColumnIndex As Variant) As Range"),
        ("Application", "Rows", "Property Get Rows(Optional RowIndex As Variant) As Range"),
        ("Application", "Columns", "Property Get Columns(Optional RowIndex As Variant) As Range"),
        ("Application", "ActiveSheet", "Property Get ActiveSheet() As Worksheet"),
        ("Workbook", "Worksheets", "Property Get Worksheets() As Worksheets"),
        ("Workbook", "Sheets", "Property Get Sheets() As Sheets"),
        ("Application", "Worksheets", "Property Get Worksheets() As Worksheets"),
        ("Application", "Sheets", "Property Get Sheets() As Sheets"),
        ("Worksheets", "Count", "Property Get Count() As Long"),
        ("Worksheets", "Item", "Property Get Item(Index As Variant) As Worksheet"),
        ("Worksheets", "Parent", "Property Get Parent() As Workbook"),
        ("Sheets", "Count", "Property Get Count() As Long"),
        ("Sheets", "Item", "Property Get Item(Index As Variant) As Worksheet"),
        ("Sheets", "Parent", "Property Get Parent() As Workbook"),
        ("Range", "Item", "Property Item(RowIndex As Variant, Optional ColumnIndex As Variant) As Range"),
        ("Workbook", "Parent", "Property Get Parent() As Application"),
        ("Worksheet", "Application", "Property Get Application() As Application"),
        ("Worksheet", "Parent", "Property Get Parent() As Workbook"),
        ("Worksheet", "Range", "Property Get Range(Cell1 As Variant, Optional Cell2 As Variant) As Range"),
        ("Worksheet", "Cells", "Property Get Cells(Optional RowIndex As Variant, Optional ColumnIndex As Variant) As Range"),
        ("Worksheet", "Rows", "Property Get Rows(Optional RowIndex As Variant) As Range"),
        ("Worksheet", "Columns", "Property Get Columns(Optional RowIndex As Variant) As Range"),
        ("Range", "Cells", "Property Get Cells(Optional RowIndex As Variant, Optional ColumnIndex As Variant) As Range"),
        ("Range", "Rows", "Property Get Rows(Optional RowIndex As Variant) As Range"),
        ("Range", "Columns", "Property Get Columns(Optional RowIndex As Variant) As Range"),
        ("Worksheet", "UsedRange", "Property Get UsedRange() As Range"),
        ("Range", "Range", "Property Get Range(Cell1 As Variant, Optional Cell2 As Variant) As Range"),
        ("Range", "Areas", "Property Get Areas() As Areas"),
        ("Areas", "Count", "Property Get Count() As Long"),
        ("Areas", "Item", "Property Get Item(Index As Variant) As Range"),
        ("Range", "Offset", "Property Get Offset(Optional RowOffset As Variant, Optional ColumnOffset As Variant) As Range"),
        ("Range", "Resize", "Property Get Resize(Optional RowSize As Variant, Optional ColumnSize As Variant) As Range"),
        ("Range", "EntireRow", "Property Get EntireRow() As Range"),
        ("Range", "EntireColumn", "Property Get EntireColumn() As Range"),
        ("Range", "End", "Property Get End(Direction As Long) As Range"),
        ("Range", "Find", "Function Find(What As Variant, Optional After As Range, Optional LookIn As Long, Optional LookAt As Long, Optional SearchOrder As Long, Optional SearchDirection As Long, Optional MatchCase As Boolean, Optional MatchByte As Boolean, Optional SearchFormat As Boolean) As Range"),
        ("Range", "FindNext", "Function FindNext(Optional After As Range) As Range"),
        ("Range", "FindPrevious", "Function FindPrevious(Optional After As Range) As Range"),
        ("Range", "Insert", "Sub Insert(Optional Shift As Long, Optional CopyOrigin As Long)"),
        ("Range", "Delete", "Sub Delete(Optional Shift As Long)"),
        ("Range", "CountLarge", "Property Get CountLarge() As Variant"),
        ("Range", "Value", "Property Value As Variant"),
        ("Range", "Application", "Property Get Application() As Application"),
        ("Range", "Parent", "Property Get Parent() As Worksheet"),
        ("Range", "Worksheet", "Property Get Worksheet() As Worksheet"),
        ("Range", "Row", "Property Get Row() As Long"),
        ("Range", "Column", "Property Get Column() As Long"),
        ("Range", "Count", "Property Get Count() As Long"),
        ("Range", "Value2", "Property Value2 As Variant"),
        ("Range", "Formula", "Property Formula As Variant"),
        ("Range", "ClearContents", "Sub ClearContents()"),
        ("Range", "Address", "Property Get Address(Optional RowAbsolute As Variant, Optional ColumnAbsolute As Variant, Optional ReferenceStyle As Long, Optional External As Variant, Optional RelativeTo As Variant) As String"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| ApiMember {
        library: "Excel".into(),
        owner: owner.into(),
        name: name.into(),
        signature: signature.into(),
        supported: true,
    })
    .chain([("ScreenUpdating", "Property ScreenUpdating As Boolean")].into_iter().map(|(name, signature)| ApiMember { library: "Excel".into(), owner: "Application".into(), name: name.into(), signature: signature.into(), supported: false }))
    .collect()
}
