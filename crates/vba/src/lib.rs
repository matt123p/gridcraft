//! Fork-owned Excel object profile over the separate tiny-vba interpreter.
//! Host calls are native typed Rust; JSON carries workbook metadata and command envelopes.
#![forbid(unsafe_code)]
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
mod objects;
mod values;
pub mod worksheet;

use gridcraft_engine::{GroupToken, Session};
use objects::Object;
use std::collections::BTreeMap;
use std::sync::atomic::{AtomicU64, Ordering};
use vba_foreign::{Completion, ForeignError, HostCall, InvokeKind, Reply};
pub use vba_ide::{self, Action, Host, Ide, Mode, ObjectRef, Value};
use vba_ide::{ApiMember, Diagnostic, Import};

pub type Result<T> = std::result::Result<T, Diagnostic>;
type HostResult<T> = std::result::Result<T, ForeignError>;
static NEXT_HOST: AtomicU64 = AtomicU64::new(1);
pub const MAX_OBJECTS: usize = 8192;
pub const MAX_CELLS: usize = 65_536;
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
    sheets: Vec<String>,
    objects: Vec<Object>,
    group: Option<GroupToken>,
    consent: bool,
    closed: bool,
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
            sheets: doc.wb.sheets.iter().map(|s| s.name.clone()).collect(),
            objects: vec![Object::Application, Object::Workbook, Object::Sheet(doc.sheet_index())],
            group: None,
            consent: false,
            closed: false,
        })
    }
    pub fn document(&self) -> u64 {
        self.doc
    }
    pub fn permitted(&self, session: &Session) -> bool {
        self.check(session).is_ok()
    }
    pub fn current(&self, session: &Session) -> bool {
        !self.closed && session.doc().is_ok_and(|doc| doc.uid == self.doc && doc.wb.sheets.iter().map(|s| &s.name).eq(self.sheets.iter()))
    }
    pub(crate) fn check(&self, session: &Session) -> HostResult<()> {
        if self.closed {
            return Err(failure(91, "VBA host is closed"));
        }
        let doc = session.doc().map_err(|_| failure(91, "workbook closed"))?;
        if doc.uid != self.doc || !doc.wb.sheets.iter().map(|s| &s.name).eq(self.sheets.iter()) {
            return Err(failure(91, "workbook or worksheet topology changed; reset the editor"));
        }
        if !doc.trust.allows_code() && !self.consent {
            return Err(failure(70, "VBA requires explicit consent for this workbook"));
        }
        Ok(())
    }
    fn reference(&self, object: usize) -> ObjectRef {
        ObjectRef { host: self.identity, object: object as u64, generation: self.doc }
    }
    pub(crate) fn intern(&mut self, object: Object) -> HostResult<Value> {
        let index = if let Some(i) = self.objects.iter().position(|o| o == &object) {
            i
        } else {
            if self.objects.len() >= MAX_OBJECTS {
                return Err(failure(7, "object handle budget exceeded"));
            }
            self.objects.push(object);
            self.objects.len().saturating_sub(1)
        };
        Ok(Value::Object(self.reference(index)))
    }
    pub fn configure(&self, ide: &mut Ide) -> Result<()> {
        ide.configure_host(self.imports(), self.roots(), catalog())
    }
    pub(crate) fn roots(&self) -> BTreeMap<String, ObjectRef> {
        BTreeMap::from([("Excel".into(), self.reference(0)), ("Workbook".into(), self.reference(1)), ("Worksheet".into(), self.reference(2))])
    }
    pub(crate) fn imports(&self) -> Vec<Import> {
        [
            ("Application", None, InvokeKind::Get),
            ("Worksheets", Some("Worksheets"), InvokeKind::Call),
            ("Range", Some("Range"), InvokeKind::Call),
            ("Cells", Some("Cells"), InvokeKind::Call),
        ]
        .into_iter()
        .map(|(name, member, kind)| Import { name: name.into(), root: "Excel".into(), member: member.map(str::to_string), kind })
        .chain([("ThisWorkbook", "Workbook"), ("ActiveWorkbook", "Workbook"), ("ActiveSheet", "Worksheet")].into_iter().map(|(name, root)| Import {
            name: name.into(),
            root: root.into(),
            member: None,
            kind: InvokeKind::Get,
        }))
        .collect()
    }
    /// Consent applies only to this editor/document instance. It never changes provenance.
    pub fn consent(&mut self, allow: bool) {
        self.consent = allow;
    }
    pub fn begin(&mut self, session: &mut Session) -> Result<()> {
        self.check(session).map_err(|e| diagnostic(e.description))?;
        if self.group.is_none() {
            self.group = Some(session.group_begin("VBA Macro").map_err(diagnostic)?);
        }
        Ok(())
    }
    /// Decode source without running it or modifying the package. Only standard modules
    /// are executable in this profile; class/document modules are refused explicitly.
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
            if module.kind != vba_project::ModuleKind::Standard {
                // Preserve inert document declarations, but refuse executable class code.
                if module.source.lines().any(|line| {
                    let l = line.trim();
                    !l.is_empty() && !l.starts_with("Attribute ") && !l.starts_with("Option ") && !l.starts_with('\'')
                }) {
                    return Err(diagnostic(format!("{}: executable class/document modules are not supported", module.name)));
                }
                continue;
            }
            ide.add_module(&format!("{}.bas", module.name), &module.source)?;
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
        let result = ide.dispatch(action, &mut self.host(session));
        if result.is_err() || matches!(action, Action::Reset) {
            self.finish(session)?;
        }
        result
    }
    pub fn tick(&mut self, session: &mut Session, ide: &mut Ide) -> Result<()> {
        if matches!(ide.mode(), Mode::Running | Mode::Waiting | Mode::Break | Mode::Stale)
            && let Err(e) = self.check(session)
        {
            // Cancel while the original host identity is still available; no side effects replay.
            let _ = ide.dispatch(Action::Reset, &mut self.host(session));
            self.finish(session)?;
            self.closed = true;
            return Err(diagnostic(e.description));
        }
        let result = ide.tick(&mut self.host(session));
        if result.is_err() || matches!(ide.mode(), Mode::Completed | Mode::Failed | Mode::Design) {
            self.finish(session)?;
        }
        result
    }
    /// Retain successful writes on failure/stop and make them one ordinary undo step.
    pub fn finish(&mut self, session: &mut Session) -> Result<()> {
        if let Some(token) = self.group.take() {
            // Closing a document already aborts its engine groups.
            if session.document_index(self.doc).is_some() {
                session.group_end(token).map_err(diagnostic)?;
            }
        }
        Ok(())
    }
    pub fn close(&mut self, session: &mut Session, ide: &mut Ide, discard: bool) -> Result<()> {
        ide.close(&mut self.host(session), discard)?;
        self.finish(session)?;
        self.closed = true;
        Ok(())
    }
}
pub struct WorkbookHost<'a> {
    bridge: &'a mut Bridge,
    session: &'a mut Session,
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
            let index = usize::try_from(call.target.object).map_err(|_| failure(91, "invalid object"))?;
            let object = self.bridge.objects.get(index).cloned().ok_or_else(|| failure(91, "invalid object"))?;
            self.bridge.invoke_object(self.session, object, call)
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
        ("Application", "ThisWorkbook", "Property Get ThisWorkbook() As Workbook"),
        ("Application", "ActiveWorkbook", "Property Get ActiveWorkbook() As Workbook"),
        ("Application", "ActiveSheet", "Property Get ActiveSheet() As Worksheet"),
        ("Workbook", "Worksheets", "Function Worksheets(Index As Variant) As Worksheet"),
        ("Worksheet", "Range", "Function Range(Cell1 As String) As Range"),
        ("Worksheet", "Cells", "Function Cells(Row As Long, Column As Long) As Range"),
        ("Range", "Value", "Property Value As Variant"),
        ("Range", "Value2", "Property Value2 As Variant"),
        ("Range", "Formula", "Property Formula As Variant"),
        ("Range", "ClearContents", "Sub ClearContents()"),
        ("Range", "Address", "Property Get Address() As String"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| ApiMember {
        library: "GridCraft.Excel.v1".into(),
        owner: owner.into(),
        name: name.into(),
        signature: signature.into(),
        supported: true,
    })
    .collect()
}
