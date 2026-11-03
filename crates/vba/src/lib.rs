//! Fork-owned Excel object profile over the separate tiny-vba interpreter.
//! Host calls are native typed Rust; JSON carries workbook metadata and command envelopes.
#![forbid(unsafe_code)]
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
mod address;
mod application;
mod host_ops;
pub use host_ops::{LifecyclePrompt, PromptRequest, PromptResponse, PromptServices};
mod chart_api;
mod chart_parts_api;
mod clipboard;
mod condition_details;
mod conditional_api;
mod data_api;
mod drawing_api;
mod features;
mod format_api;
mod formatting;
mod formula_api;
mod lifecycle;
mod lifecycle_events;
mod series_api;
pub use lifecycle_events::{LifecycleEvent, LifecycleEventBatch};
mod named_style_api;
mod objects;
mod outline_api;
mod protection_api;
mod ranges;
mod references;
mod scenario_api;
mod search;
mod structure;
mod table_api;
mod types;
mod values;
mod whatif_api;
pub mod worksheet;

use gridcraft_engine::core::RangeRef;
use gridcraft_engine::{GroupToken, Session};
use objects::Object;
use std::collections::{BTreeMap, VecDeque};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, RwLock};
use vba_foreign::{Completion, ForeignError, HostCall, InvokeKind, Reply};
pub use vba_ide::{self, Action, Host, Ide, Limits, Mode, ObjectRef, Value};
use vba_ide::{ApiMember, Diagnostic, Import};

pub type Result<T> = std::result::Result<T, Diagnostic>;
type HostResult<T> = std::result::Result<T, ForeignError>;
static NEXT_HOST: AtomicU64 = AtomicU64::new(1);
pub const MAX_OBJECTS: usize = 8192;
pub const MAX_CELLS: usize = 65_536;
const MAX_QUEUED_EVENTS: usize = 256;

/// Per-invocation macro limits for the workbook host. The workbook model and
/// clipboard already cap individual operations; these limits bound aggregate work.
pub const MACRO_EXECUTION_LIMITS: Limits = Limits {
    instructions: 25_000_000,
    slice: 4096,
    depth: 128,
    locations: 4096,
    stack_entries: 65_536,
    host_calls: 12_000,
    host_polls: 100_000,
    // The VM reserves 64 KiB of reply/snapshot space per host call. Keep this
    // aggregate cap within the IDE's 1 GiB maximum while allowing recorded loops.
    payload: 1024 * 1024 * 1024,
    callbacks: 256,
};

/// Availability declared once for completion, the object browser, dispatch
/// validation and the offline object-model audit.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MemberAvailability {
    Implemented,
    OutOfScope,
    ServiceRequired,
}

/// GridCraft's supported Excel member profile. The signature is retained in
/// VBA syntax because tiny-vba consumes it for completion and nominal types.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct MemberDescriptor {
    pub owner: &'static str,
    pub name: &'static str,
    pub signature: &'static str,
    pub availability: MemberAvailability,
    pub implementation_symbol: &'static str,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CompatibilityStatus {
    Supported,
    OutOfScope,
    ServiceRequired,
    Missing,
    Unknown,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct CompatibilityFinding<'a> {
    pub receiver: Option<&'a str>,
    pub member: &'a str,
    pub status: CompatibilityStatus,
    pub signature: Option<&'static str>,
}
pub const DEMO: &str = "Option Explicit\nPublic Sub Main()\n    Dim i As Long\n    For i = 1 To 10\n        Cells(i, 1).Value2 = i\n    Next i\n    Range(\"B1\").Formula = \"=SUM(A1:A10)\"\nEnd Sub\n";

fn diagnostic(e: impl std::fmt::Display) -> Diagnostic {
    Diagnostic::new("GRIDCRAFT_VBA", e.to_string())
}
fn failure(number: i32, text: impl Into<String>) -> ForeignError {
    ForeignError { number, description: text.into() }
}

type PendingEvent = (String, String, Vec<Value>);

/// Fully materialized callback data and reserved queue capacity. Preparation and
/// commit surround one synchronous operation; no callback runs between them.
#[derive(Default)]
struct PreparedChanges {
    events: Vec<PendingEvent>,
}

/// State persists across bounded IDE ticks, but never holds a document borrow.
pub struct Bridge {
    identity: u64,
    pub(crate) doc: u64,
    sheets: BTreeMap<u64, Vec<gridcraft_engine::model::sheet::SheetIdentity>>,
    context: u64,
    sorts: Vec<data_api::SortState>,
    formula_locale: gridcraft_engine::formula::FormulaLocale,
    range_epochs: BTreeMap<u64, u64>,
    objects: Arc<types::Objects>,
    group: BTreeMap<u64, GroupToken>,
    consent: bool,
    enable_events: bool,
    pending_events: VecDeque<PendingEvent>,
    event_work: usize,
    pub(crate) event_owner_sheet: Option<usize>,
    document_code_names: Arc<RwLock<std::collections::BTreeSet<String>>>,
    closed: bool,
    operations: host_ops::Operations,
    lifecycle_handlers: std::sync::atomic::AtomicBool,
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
            sorts: Vec::new(),
            formula_locale: gridcraft_engine::formula::FormulaLocale::En,
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
            operations: Default::default(),
            lifecycle_handlers: std::sync::atomic::AtomicBool::new(false),
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
        if doc.uid == self.doc
            && let Some(owner) = self.event_owner_sheet
        {
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
        let object = if matches!(object, Object::Application | Object::Workbooks | Object::Document(..)) {
            object
        } else {
            Object::Document(self.context, Box::new(object))
        };
        let index = self.objects.intern(object)?;
        Ok(Value::Object(self.reference(index)))
    }
    /// The embedding supplies its existing formula-entry locale; headless defaults to canonical English.
    pub fn set_formula_locale(&mut self, locale: gridcraft_engine::formula::FormulaLocale) {
        self.formula_locale = locale;
    }
    pub fn configure(&self, ide: &mut Ide) -> Result<()> {
        self.configure_with_workbook(ide, None)
    }
    /// Declare the open workbook's document context: implicit `Me` members per
    /// document module (derived from the supported catalogue), instance types for
    /// completion and result types for host globals. Metadata only; code aids stay
    /// read-only and never invoke host members.
    pub fn configure_document(&self, ide: &mut Ide, wb: &gridcraft_model::Workbook) -> Result<()> {
        self.configure_with_workbook(ide, Some(wb))
    }
    fn configure_with_workbook(&self, ide: &mut Ide, wb: Option<&gridcraft_model::Workbook>) -> Result<()> {
        self.refresh_lifecycle_handlers(ide);
        ide.configure_execution_limits(MACRO_EXECUTION_LIMITS.clone())?;
        let comparison_context =
            vba_vm::IntrinsicContextConfig::new("invariant".into(), "Gregorian".into(), "UTF-16".into(), usize::BITS as u8).map_err(diagnostic)?;
        ide.configure_string_comparison_service(Arc::new(vba_runtime::BinaryStringComparisonService), comparison_context).map_err(diagnostic)?;
        if let Ok(mut names) = self.document_code_names.write() {
            names.clear();
            names.extend(ide.modules.iter().filter_map(|module| {
                (module.kind == vba_ide::ModuleKind::Document).then(|| module.code_name.clone()).flatten().map(|name| name.to_ascii_lowercase())
            }));
        }
        ide.configure_host(self.imports(), self.roots(), catalog())?;
        ide.configure_reference_bindings(types::references()?)?;
        ide.configure_type_resolver(Some(self.objects.clone()))?;
        let Some(wb) = wb else {
            return Ok(());
        };
        let supported = |owner: &str| {
            let mut names: Vec<String> =
                catalog().iter().filter(|item| item.supported && item.owner == owner).map(|item| item.name.clone()).collect();
            names.sort_by_key(|name| name.to_ascii_lowercase());
            names.dedup_by(|a, b| a.eq_ignore_ascii_case(b));
            names
        };
        let mut document_members = BTreeMap::new();
        let mut document_types = BTreeMap::new();
        if let Some(code) = wb.code_name.as_deref() {
            let key = code.to_ascii_lowercase();
            document_members.insert(key.clone(), supported("Workbook"));
            document_types.insert(key, "Workbook".into());
        }
        for sheet in &wb.sheets {
            let Some(code) = sheet.code_name.as_deref() else {
                continue;
            };
            let key = code.to_ascii_lowercase();
            document_members.insert(key.clone(), supported("Worksheet"));
            document_types.insert(key, "Worksheet".into());
        }
        ide.configure_document_context(
            document_members,
            document_types,
            Self::HOST_GLOBAL_TYPES.iter().map(|(name, ty)| (name.to_string(), ty.to_string())).collect(),
        )?;
        Ok(())
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
            Import { name: "Run".into(), root: "Excel".into(), member: Some("Run".into()), kind: InvokeKind::Call },
        ];
        for name in [
            "ActiveWorkbook",
            "ActiveSheet",
            "ActiveCell",
            "Selection",
            "Evaluate",
            "WorksheetFunction",
            "Names",
            "Workbooks",
            "Worksheets",
            "Sheets",
            "Range",
            "Cells",
            "Rows",
            "Columns",
        ] {
            imports.push(Import {
                name: name.into(),
                root: "Excel".into(),
                member: Some(name.into()),
                kind: if name == "Range" { InvokeKind::Call } else { InvokeKind::Get },
            });
        }
        imports
    }
    /// Declared result type of each host global for read-only completion metadata.
    const HOST_GLOBAL_TYPES: &'static [(&'static str, &'static str)] = &[
        ("Application", "Application"),
        ("ThisWorkbook", "Workbook"),
        ("ActiveWorkbook", "Workbook"),
        ("ActiveSheet", "Worksheet"),
        ("ActiveCell", "Range"),
        ("Selection", "Range"),
        ("WorksheetFunction", "WorksheetFunction"),
        ("Names", "Names"),
        ("Workbooks", "Workbooks"),
        ("Worksheets", "Worksheets"),
        ("Sheets", "Sheets"),
        ("Range", "Range"),
        ("Cells", "Range"),
        ("Rows", "Range"),
        ("Columns", "Range"),
    ];
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
        self.operations.pending.clear();
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
        // Imported document modules belong only to this bridge's workbook.
        // Code names are not globally unique across open workbooks.
        if doc.uid != self.doc {
            return Ok(None);
        }
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
    fn prepare_changes(&mut self, session: &Session, changes: &[(usize, ranges::Range)]) -> HostResult<PreparedChanges> {
        let mut owners = Vec::new();
        owners.try_reserve(changes.len()).map_err(|_| failure(7, "worksheet event preparation allocation"))?;
        for (sheet, range) in changes {
            if let Some(owner) = self.change_owner(session, *sheet)? {
                range.validate()?;
                owners.push((owner, *sheet, range));
            }
        }
        if owners.is_empty() {
            return Ok(PreparedChanges::default());
        }
        self.reset_event_budget_if_idle();
        let count = owners.len();
        if self.event_work.checked_add(count).is_none_or(|n| n > MAX_QUEUED_EVENTS)
            || self.pending_events.len().checked_add(count).is_none_or(|n| n > MAX_QUEUED_EVENTS)
        {
            return Err(failure(28, "worksheet event budget exceeded"));
        }
        self.pending_events.try_reserve(count).map_err(|_| failure(7, "worksheet event queue allocation"))?;
        let mut objects = Vec::new();
        objects.try_reserve(count * 2).map_err(|_| failure(7, "worksheet event target allocation"))?;
        for (_, sheet, range) in &owners {
            // Reserve both the callback owner and its exact Target, qualified by
            // the actual target document rather than the last ambient context.
            objects.push(Object::Document(self.doc, Box::new(Object::Sheet(*sheet))));
            objects.push(Object::Document(self.doc, Box::new(Object::Range(*sheet, (*range).clone()))));
        }
        let handles = self.objects.intern_many(objects)?;
        let mut events = Vec::new();
        events.try_reserve(count).map_err(|_| failure(7, "worksheet event preparation allocation"))?;
        for ((owner, _, _), pair) in owners.into_iter().zip(handles.as_chunks::<2>().0) {
            events.push((owner, "Worksheet_Change".into(), vec![Value::Object(self.reference(pair[1]))]));
        }
        Ok(PreparedChanges { events })
    }
    fn prepare_change(&mut self, session: &Session, sheet: usize, range: ranges::Range) -> HostResult<PreparedChanges> {
        self.prepare_changes(session, &[(sheet, range)])
    }
    /// Infallible after a successful synchronous mutation: all handles, callback
    /// arguments and queue storage were prepared before the engine was called.
    fn queue_changes(&mut self, changes: PreparedChanges) {
        self.event_work += changes.events.len();
        self.pending_events.extend(changes.events);
    }
    fn write_range(
        &mut self,
        session: &mut Session,
        sheet: usize,
        range: &ranges::Range,
        value: &Value,
        formula: bool,
        typed: bool,
    ) -> HostResult<Value> {
        let rectangle = range.rectangle()?;
        let changes = self.prepare_change(session, sheet, range.clone())?;
        values::write(session, sheet, rectangle, value, formula, typed)?;
        self.queue_changes(changes);
        Ok(Value::Empty)
    }
    pub fn notify_sheet_change(&mut self, session: &Session, sheet: usize, range: RangeRef) -> Result<()> {
        if !self.current(session) || !self.permitted(session) {
            return Ok(());
        }
        let changes = self.prepare_change(session, sheet, range.into()).map_err(|error| diagnostic(error.description))?;
        self.queue_changes(changes);
        Ok(())
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
        self.configure_with_workbook(&mut ide, Some(doc.wb.as_ref()))?;
        Ok(ide)
    }
    pub fn host<'a>(&'a mut self, session: &'a mut Session) -> WorkbookHost<'a> {
        WorkbookHost { bridge: self, session }
    }
    pub fn dispatch(&mut self, session: &mut Session, ide: &mut Ide, action: Action) -> Result<()> {
        self.refresh_lifecycle_handlers(ide);
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
        self.tick_events(session, ide, true)
    }
    /// The shell owns the active lifecycle batch and must collect its ByRef outputs
    /// before ordinary queued mutations can start another callback.
    pub fn tick_lifecycle(&mut self, session: &mut Session, ide: &mut Ide) -> Result<()> {
        self.tick_events(session, ide, false)
    }
    fn tick_events(&mut self, session: &mut Session, ide: &mut Ide, deliver_events: bool) -> Result<()> {
        self.refresh_lifecycle_handlers(ide);
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
        if !self.prepare_operation(session, ide)? {
            return Ok(());
        }
        let result = ide.tick(&mut self.host(session));
        self.drain_host_results(session, ide)?;
        if result.is_ok()
            && deliver_events
            && matches!(ide.mode(), Mode::Completed | Mode::Design)
            && let Some((code_name, handler, arguments)) = self.pending_events.pop_front()
        {
            if self.group.is_empty() {
                self.ensure_group(session, self.doc)?;
            }
            let owner = session
                .with_target(gridcraft_engine::Target::Doc(self.doc), |session| {
                    self.sync_sheets(session)?;
                    self.event_owner_sheet = session.doc().ok().and_then(|doc| {
                        doc.wb.sheets.iter().position(|sheet| sheet.code_name.as_deref().is_some_and(|name| name.eq_ignore_ascii_case(&code_name)))
                    });
                    self.document_owner_object(session, &code_name)
                })
                .map_err(diagnostic)?
                .map_err(|error| diagnostic(error.description))?;
            let Value::Object(owner) = owner else { return Err(diagnostic("document owner handle is not an object")) };
            if let Err(error) = ide.dispatch_document_event_for_owner(&code_name, &handler, arguments, Some(owner), &mut self.host(session)) {
                self.cancel_pending_events();
                self.finish(session)?;
                return Err(error);
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
        let mut suspended = None;
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
            let target = self.session.document(doc).ok_or_else(|| failure(91, "workbook closed"))?;
            if !target.trust.allows_code() && !(doc == self.bridge.doc && self.bridge.consent) {
                return Err(failure(70, "target workbook requires consent"));
            }
            self.bridge.ensure_group(self.session, doc).map_err(|e| failure(7, e.to_string()))?;
            let activate = call.kind == InvokeKind::Call
                && call.member.eq_ignore_ascii_case("Activate")
                && matches!(object, Object::Workbook | Object::Sheet(_));
            let result = self
                .session
                .with_target(gridcraft_engine::Target::Doc(doc), |session| {
                    self.bridge.sync_sheets(session)?;
                    // Resolve again after sheet topology synchronization.
                    let object = match self.bridge.objects.get(call.target)? {
                        Object::Document(_, object) => *object,
                        _ => object,
                    };
                    self.bridge.sync_chart_owner(session, &object)?;
                    let object = match self.bridge.objects.get(call.target)? {
                        Object::Document(_, object) => *object,
                        object => object,
                    };
                    if let Some(token) = self.bridge.defer_call(session, &object, &call)? {
                        suspended = Some(token);
                        return Ok(Value::Empty);
                    }
                    self.bridge.invoke_object(session, object, call)
                })
                .map_err(|e| failure(91, e.to_string()))?;
            if result.is_ok()
                && activate
                && let Some(index) = self.session.document_index(doc)
            {
                self.session.set_active(index);
            }
            if self.session.active().is_some() {
                self.bridge.sync_sheets(self.session)?;
            }
            result
        })();
        match result {
            Ok(value) => {
                if let Some(token) = suspended {
                    Reply::Wait(token)
                } else {
                    Reply::Complete(Completion { value, writes: vec![] })
                }
            }
            Err(e) => Reply::Error(e),
        }
    }
    fn resume(&mut self, token: u64) -> Reply {
        if let Err(error) = self.bridge.check(self.session) {
            self.bridge.cancel_operation(token);
            return Reply::Error(error);
        }
        self.bridge.operation_reply(token)
    }
    fn cancel(&mut self, token: u64) {
        self.bridge.cancel_operation(token);
    }
}

pub fn member_descriptors() -> Vec<MemberDescriptor> {
    [
        ("Application", "Workbooks", "Property Get Workbooks() As Workbooks"),
        ("Application", "Name", "Property Get Name() As String"),
        ("Application", "Version", "Property Get Version() As String"),
        ("Application", "OperatingSystem", "Property Get OperatingSystem() As String"),
        ("Application", "CalculateFull", "Sub CalculateFull()"),
        ("Application", "CalculateFullRebuild", "Sub CalculateFullRebuild()"),
        ("Application", "Iteration", "Property Iteration As Boolean"),
        ("Application", "MaxIterations", "Property MaxIterations As Long"),
        ("Application", "MaxChange", "Property MaxChange As Double"),
        ("Workbooks", "Count", "Property Get Count() As Long"),
        ("Workbooks", "Item", "Property Get Item(Index As Variant) As Workbook"),
        ("Workbooks", "Add", "Function Add(Optional Template As Variant) As Workbook"),
        ("Workbooks", "Open", "Function Open(Filename As Variant, Optional UpdateLinks As Variant, Optional ReadOnly As Variant, Optional Format As Variant, Optional Password As Variant, Optional WriteResPassword As Variant, Optional IgnoreReadOnlyRecommended As Variant, Optional Origin As Variant, Optional Delimiter As Variant, Optional Editable As Variant, Optional Notify As Variant, Optional Converter As Variant, Optional AddToMru As Variant, Optional Local As Variant, Optional CorruptLoad As Variant) As Workbook"),
        ("Workbooks", "Parent", "Property Get Parent() As Application"),
        ("Workbooks", "Application", "Property Get Application() As Application"),
        ("Workbook", "Activate", "Sub Activate()"),
        ("Workbook", "Path", "Property Get Path() As String"),
        ("Workbook", "FullName", "Property Get FullName() As String"),
        ("Workbook", "CodeName", "Property Get CodeName() As String"),
        ("Workbook", "Saved", "Property Get Saved() As Boolean"),
        ("Workbook", "Protect", "Sub Protect(Optional Password As Variant, Optional Structure As Variant, Optional Windows As Variant)"),
        ("Workbook", "Unprotect", "Sub Unprotect(Optional Password As Variant)"),
        ("Workbook", "ProtectStructure", "Property Get ProtectStructure() As Boolean"),
        ("Workbook", "Save", "Sub Save()"),
        ("Workbook", "SaveAs", "Sub SaveAs(Filename As Variant, Optional FileFormat As Variant, Optional Password As Variant, Optional WriteResPassword As Variant, Optional ReadOnlyRecommended As Variant, Optional CreateBackup As Variant, Optional AccessMode As Variant, Optional ConflictResolution As Variant, Optional AddToMru As Variant, Optional TextCodepage As Variant, Optional TextVisualLayout As Variant, Optional Local As Variant)"),
        ("Workbook", "SaveCopyAs", "Sub SaveCopyAs(Filename As Variant)"),
        ("Workbook", "Close", "Sub Close(Optional SaveChanges As Variant, Optional Filename As Variant, Optional RouteWorkbook As Variant)"),
        ("Worksheets", "Add", "Function Add(Optional Before As Variant, Optional After As Variant, Optional Count As Variant, Optional Type As Variant) As Worksheet"),
        ("Sheets", "Add", "Function Add(Optional Before As Variant, Optional After As Variant, Optional Count As Variant, Optional Type As Variant) As Worksheet"),
        ("Worksheet", "Index", "Property Get Index() As Long"),
        ("Worksheet", "CodeName", "Property Get CodeName() As String"),
        ("Worksheet", "Visible", "Property Visible As XlSheetVisibility"),
        ("Worksheet", "Delete", "Function Delete() As Boolean"),
        ("Worksheet", "Copy", "Sub Copy(Optional Before As Variant, Optional After As Variant)"),
        ("Worksheet", "Move", "Sub Move(Optional Before As Variant, Optional After As Variant)"),
        ("Application", "Application", "Property Get Application() As Application"),
        ("Application", "Parent", "Property Get Parent() As Application"),
        ("Application", "ActiveCell", "Property Get ActiveCell() As Range"),
        ("Application", "Selection", "Property Get Selection() As Range"),
        ("Application", "Goto", "Sub Goto(Optional Reference As Variant, Optional Scroll As Variant)"),
        ("Application", "PreviousSelections", "Property Get PreviousSelections(Optional Index As Variant) As Variant"),
        ("Application", "Union", "Function Union(Arg1 As Range, Arg2 As Range, Optional Arg3 As Variant) As Range"),
        ("Application", "Intersect", "Function Intersect(Arg1 As Range, Arg2 As Range, Optional Arg3 As Variant) As Range"),
        ("Application", "Calculation", "Property Calculation As XlCalculation"),
        ("Application", "EnableEvents", "Property EnableEvents As Boolean"),
        ("Application", "CutCopyMode", "Property CutCopyMode As Variant"),
        ("Application", "Calculate", "Sub Calculate()"),
        ("Workbook", "Calculate", "Sub Calculate()"),
        ("Workbook", "Application", "Property Get Application() As Application"),
        ("Workbook", "ActiveSheet", "Property Get ActiveSheet() As Worksheet"),
        ("Workbook", "Name", "Property Get Name() As String"),
        ("Worksheet", "Name", "Property Name As String"),
        ("Worksheet", "Select", "Sub Select(Optional Replace As Variant)"),
        ("Worksheet", "Activate", "Sub Activate()"),
        ("Worksheet", "Paste", "Sub Paste(Optional Destination As Variant, Optional Link As Variant)"),
        ("Range", "Select", "Sub Select()"),
        ("Range", "Activate", "Sub Activate()"),
        ("Range", "Copy", "Sub Copy(Optional Destination As Variant)"),
        ("Range", "Cut", "Sub Cut(Optional Destination As Variant)"),
        ("Range", "PasteSpecial", "Function PasteSpecial(Optional Paste As XlPasteType, Optional Operation As XlPasteSpecialOperation, Optional SkipBlanks As Variant, Optional Transpose As Variant) As Variant"),
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
        ("Worksheets", "Application", "Property Get Application() As Application"),
        ("Sheets", "Count", "Property Get Count() As Long"),
        ("Sheets", "Item", "Property Get Item(Index As Variant) As Worksheet"),
        ("Sheets", "Parent", "Property Get Parent() As Workbook"),
        ("Sheets", "Application", "Property Get Application() As Application"),
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
        ("Areas", "Application", "Property Get Application() As Application"),
        ("Areas", "Parent", "Property Get Parent() As Range"),
        ("Range", "Offset", "Property Get Offset(Optional RowOffset As Variant, Optional ColumnOffset As Variant) As Range"),
        ("Range", "Resize", "Property Get Resize(Optional RowSize As Variant, Optional ColumnSize As Variant) As Range"),
        ("Range", "CurrentRegion", "Property Get CurrentRegion() As Range"),
        ("Range", "CurrentArray", "Property Get CurrentArray() As Range"),
        ("Range", "HasArray", "Property Get HasArray() As Variant"),
        ("Range", "HasSpill", "Property Get HasSpill() As Variant"),
        ("Range", "SpillParent", "Property Get SpillParent() As Range"),
        ("Range", "Next", "Property Get Next() As Range"),
        ("Range", "Previous", "Property Get Previous() As Range"),
        ("Range", "EntireRow", "Property Get EntireRow() As Range"),
        ("Range", "EntireColumn", "Property Get EntireColumn() As Range"),
        ("Range", "End", "Property Get End(Direction As Long) As Range"),
        ("Range", "Find", "Function Find(What As Variant, Optional After As Range, Optional LookIn As Long, Optional LookAt As Long, Optional SearchOrder As Long, Optional SearchDirection As Long, Optional MatchCase As Boolean, Optional MatchByte As Boolean, Optional SearchFormat As Boolean) As Range"),
        ("Range", "FindNext", "Function FindNext(Optional After As Range) As Range"),
        ("Range", "FindPrevious", "Function FindPrevious(Optional After As Range) As Range"),
        ("Range", "Replace", "Function Replace(What As Variant, Replacement As Variant, Optional LookAt As XlLookAt, Optional SearchOrder As XlSearchOrder, Optional MatchCase As Variant, Optional MatchByte As Variant, Optional SearchFormat As Variant, Optional ReplaceFormat As Variant, Optional FormulaVersion As XlReplaceFormula) As Boolean"),
        ("Range", "FillDown", "Function FillDown() As Variant"),
        ("Range", "FillUp", "Function FillUp() As Variant"),
        ("Range", "FillLeft", "Function FillLeft() As Variant"),
        ("Range", "FillRight", "Function FillRight() As Variant"),
        ("Range", "AutoFill", "Function AutoFill(Destination As Range, Optional Type As XlAutoFillType) As Variant"),
        ("Range", "SpecialCells", "Function SpecialCells(Type As XlCellType, Optional Value As Variant) As Range"),
        ("Range", "DataSeries", "Function DataSeries(Optional Rowcol As Variant, Optional Type As XlDataSeriesType, Optional Date As XlDataSeriesDate, Optional Step As Variant, Optional Stop As Variant, Optional Trend As Variant) As Variant"),
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
        ("Range", "Clear", "Sub Clear()"),
        ("Range", "ClearFormats", "Sub ClearFormats()"),
        ("Range", "ClearComments", "Sub ClearComments()"),
        ("Range", "ClearNotes", "Sub ClearNotes()"),
        ("Range", "ClearHyperlinks", "Sub ClearHyperlinks()"),
        ("Range", "Address", "Property Get Address(Optional RowAbsolute As Variant, Optional ColumnAbsolute As Variant, Optional ReferenceStyle As Long, Optional External As Variant, Optional RelativeTo As Variant) As String"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| MemberDescriptor {
        owner,
        name,
        signature,
        availability: MemberAvailability::Implemented,
        implementation_symbol: if owner == "Workbooks"
            || owner == "Application" && name == "Workbooks"
            || owner == "Workbook" && matches!(name, "Activate" | "Save" | "SaveAs" | "SaveCopyAs" | "Close" | "Path" | "FullName" | "Saved" | "CodeName" | "ProtectStructure" | "Protect" | "Unprotect")
            || matches!(owner, "Worksheets" | "Sheets") && name == "Add"
            || owner == "Worksheet" && matches!(name, "Index" | "CodeName" | "Visible" | "Delete" | "Copy" | "Move") {
            "crates/vba/src/lifecycle.rs::Bridge::lifecycle"
        } else if owner == "Application" && matches!(name, "Name" | "Version" | "OperatingSystem" | "CalculateFull" | "CalculateFullRebuild" | "Iteration" | "MaxIterations" | "MaxChange") {
            "crates/vba/src/application.rs::Bridge::application_state"
        } else if matches!(owner, "Font" | "Interior") {
            "crates/vba/src/formatting.rs::Bridge::format"
        } else {
            "crates/vba/src/objects.rs::Bridge::invoke_object"
        },
    })
    .chain([
        ("Application", "Run", "Function Run(Macro As Variant, Optional Arg1 As Variant, Optional Arg2 As Variant, Optional Arg3 As Variant, Optional Arg4 As Variant, Optional Arg5 As Variant, Optional Arg6 As Variant, Optional Arg7 As Variant, Optional Arg8 As Variant, Optional Arg9 As Variant, Optional Arg10 As Variant, Optional Arg11 As Variant, Optional Arg12 As Variant, Optional Arg13 As Variant, Optional Arg14 As Variant, Optional Arg15 As Variant, Optional Arg16 As Variant, Optional Arg17 As Variant, Optional Arg18 As Variant, Optional Arg19 As Variant, Optional Arg20 As Variant, Optional Arg21 As Variant, Optional Arg22 As Variant, Optional Arg23 As Variant, Optional Arg24 As Variant, Optional Arg25 As Variant, Optional Arg26 As Variant, Optional Arg27 As Variant, Optional Arg28 As Variant, Optional Arg29 As Variant, Optional Arg30 As Variant ) As Variant"),
        ("Application", "ScreenUpdating", "Property ScreenUpdating As Boolean"),
        ("Worksheet", "PasteSpecial", "Sub PasteSpecial(Optional Format As Variant, Optional Link As Variant, Optional DisplayAsIcon As Variant, Optional IconFileName As Variant, Optional IconIndex As Variant, Optional IconLabel As Variant, Optional NoHTMLFormatting As Variant)"),
        ("Range", "CopyPicture", "Sub CopyPicture(Optional Appearance As XlPictureAppearance, Optional Format As XlCopyPictureFormat)"),
    ].into_iter().map(|(owner, name, signature)| MemberDescriptor {
        owner,
        name,
        signature,
        availability: if owner == "Application" && name == "Run" { MemberAvailability::Implemented } else { MemberAvailability::OutOfScope },
        implementation_symbol: if owner == "Application" && name == "Run" { "crates/vba/src/host_ops.rs::Bridge::defer_call" } else { "crates/vba/src/objects.rs::Bridge::invoke_object" },
    }))
    .chain(chart_api::descriptors())
    .chain(chart_parts_api::descriptors())
    .chain(series_api::descriptors())
    .chain(drawing_api::descriptors())
    .chain(formula_api::descriptors())
    .chain(format_api::descriptors())
    .chain(named_style_api::descriptors())
    .chain(protection_api::descriptors())
    .chain(conditional_api::descriptors())
    .chain(condition_details::descriptors())
    .chain(data_api::descriptors())
    .chain(outline_api::descriptors())
    .chain(table_api::descriptors())
    .chain(whatif_api::descriptors())
    .chain(scenario_api::descriptors())
    .collect()
}

pub fn catalog() -> Vec<ApiMember> {
    member_descriptors()
        .into_iter()
        .map(|member| ApiMember {
            library: if types::OFFICE_TYPES.contains(&member.owner) { "Office" } else { "Excel" }.into(),
            owner: member.owner.into(),
            name: member.name.into(),
            signature: member.signature.into(),
            supported: member.availability == MemberAvailability::Implemented,
        })
        .collect()
}

/// Classify a statically resolved host call without executing it. `None` means
/// the receiver is dynamic and must remain UNKNOWN rather than being certified.
pub fn compatibility<'a>(receiver: Option<&'a str>, member: &'a str) -> CompatibilityFinding<'a> {
    let Some(receiver) = receiver else {
        return CompatibilityFinding { receiver: None, member, status: CompatibilityStatus::Unknown, signature: None };
    };
    let descriptor = member_descriptors()
        .into_iter()
        .find(|candidate| candidate.owner.eq_ignore_ascii_case(receiver) && candidate.name.eq_ignore_ascii_case(member));
    match descriptor {
        Some(descriptor) => CompatibilityFinding {
            receiver: Some(receiver),
            member,
            status: match descriptor.availability {
                MemberAvailability::Implemented => CompatibilityStatus::Supported,
                MemberAvailability::ServiceRequired => CompatibilityStatus::ServiceRequired,
                MemberAvailability::OutOfScope => CompatibilityStatus::OutOfScope,
            },
            signature: Some(descriptor.signature),
        },
        None => CompatibilityFinding { receiver: Some(receiver), member, status: CompatibilityStatus::Missing, signature: None },
    }
}

/// Produce the JSON ledger consumed by nonexecuting compatibility tooling.
/// Callers pass `None` when static analysis cannot prove a receiver type.
pub fn compatibility_report<'a>(calls: impl IntoIterator<Item = (Option<&'a str>, &'a str)>) -> serde_json::Value {
    let findings = calls
        .into_iter()
        .map(|(receiver, member)| {
            let finding = compatibility(receiver, member);
            let status = match finding.status {
                CompatibilityStatus::Supported => "SUPPORTED",
                CompatibilityStatus::ServiceRequired => "SERVICE_REQUIRED",
                CompatibilityStatus::OutOfScope => "OUT_OF_SCOPE_GRIDCRAFT",
                CompatibilityStatus::Missing => "MISSING",
                CompatibilityStatus::Unknown => "UNKNOWN",
            };
            serde_json::json!({
                "receiver": finding.receiver,
                "member": finding.member,
                "status": status,
                "signature": finding.signature,
            })
        })
        .collect::<Vec<_>>();
    serde_json::json!({"schema":1,"profile":"GridCraft.Excel.v1","findings":findings})
}

#[cfg(test)]
mod descriptor_tests {
    use super::{CompatibilityStatus, catalog, compatibility, compatibility_report, member_descriptors};
    use std::collections::BTreeSet;

    #[test]
    fn catalogue_is_derived_from_unique_member_descriptors() {
        let descriptors = member_descriptors();
        let keys = descriptors.iter().map(|member| (member.owner.to_ascii_lowercase(), member.name.to_ascii_lowercase())).collect::<BTreeSet<_>>();
        assert_eq!(keys.len(), descriptors.len());
        assert_eq!(catalog().len(), descriptors.len());
    }

    #[test]
    fn compatibility_never_certifies_dynamic_receivers() {
        assert_eq!(compatibility(Some("Range"), "FillDown").status, CompatibilityStatus::Supported);
        assert_eq!(compatibility(Some("Range"), "CopyPicture").status, CompatibilityStatus::OutOfScope);
        assert_eq!(compatibility(Some("Range"), "DefinitelyMissing").status, CompatibilityStatus::Missing);
        assert_eq!(compatibility(None, "FillDown").status, CompatibilityStatus::Unknown);
        let report = compatibility_report([(Some("Range"), "FillDown"), (None, "FillDown")]);
        assert_eq!(report["findings"][0]["status"], "SUPPORTED");
        assert_eq!(report["findings"][1]["status"], "UNKNOWN");
    }
}
