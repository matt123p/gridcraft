//! Explicit edited-package export snapshots; ordinary document saves stay separate.
use crate::{Bridge, Ide, Result, diagnostic};
use gridcraft_engine::Session;
use gridcraft_model::Workbook;
use std::sync::Arc;
#[derive(Clone)]
pub struct NativeWorkbookSave {
    document: u64,
    revision: u64,
    original: Arc<Workbook>,
    project: vba_ide::NativeProjectSave,
}
impl NativeWorkbookSave {
    pub fn workbook(&self) -> &Workbook {
        &self.original
    }
    pub fn original_project(&self) -> &[u8] {
        self.original.vba_project.as_deref().map_or(&[], Vec::as_slice)
    }
    pub fn encoded_project(&self) -> &[u8] {
        self.project.bytes()
    }
}
impl Bridge {
    pub fn prepare_native_project_export(&self, session: &Session, ide: &Ide) -> Result<NativeWorkbookSave> {
        if !self.current(session) {
            return Err(diagnostic("native save workbook identity changed"));
        }
        let doc = session.doc().map_err(diagnostic)?;
        let bytes = doc.wb.vba_project.as_deref().ok_or_else(|| diagnostic("workbook has no native project"))?;
        let project = ide.export_native_project(bytes)?;
        Ok(NativeWorkbookSave { document: doc.uid, revision: doc.revision, original: doc.wb.clone(), project })
    }
    pub fn validate_native_project_export(&self, session: &Session, ide: &Ide, save: &NativeWorkbookSave) -> Result<()> {
        let doc = session.doc().map_err(diagnostic)?;
        if !self.current(session) || doc.uid != save.document || doc.revision != save.revision || !Arc::ptr_eq(&doc.wb, &save.original) {
            return Err(diagnostic("workbook changed during native project export"));
        }
        let bytes = doc.wb.vba_project.as_deref().ok_or_else(|| diagnostic("native project removed"))?;
        ide.validate_native_project_save(&save.project, bytes)
    }
    /// Acknowledges an explicitly saved edited package copy, never an ordinary save
    /// of the original workbook. The caller must have removed signatures and committed IO.
    pub fn acknowledge_native_project_export(&self, session: &Session, ide: &mut Ide, save: &NativeWorkbookSave) -> Result<()> {
        self.validate_native_project_export(session, ide, save)?;
        ide.acknowledge_native_project_save(&save.project)
    }
}
