//! Excel type definitions stay in GridCraft. tiny-vba sees opaque nominal identities.
use super::*;
use std::sync::{RwLock, atomic::AtomicBool};
use vba_foreign::{ObjectTypeResolver, ReferenceLibrary, ReferenceSet, ReferenceType, ReferenceTypeKind};
use vba_values::NominalType;

// Adapter profile identity, not a claim about an installed Office typelib version.
const EXCEL_TYPES_VERSION: &str = "GridCraft.Excel.v1";

pub(crate) fn references() -> Result<ReferenceSet> {
    ReferenceSet::new(vec![ReferenceLibrary {
        name: "Excel".into(),
        version: EXCEL_TYPES_VERSION.into(),
        available: true,
        types: ["Application", "Workbook", "Worksheet", "Range"]
            .into_iter()
            .map(|name| ReferenceType { name: name.into(), kind: ReferenceTypeKind::Object })
            .collect(),
        constants: vec![],
    }])
}

/// This is the bridge's existing object store, shared read-only with the type
/// resolver. It is not a second registry or a document/engine borrow.
#[derive(Debug)]
pub(crate) struct Objects {
    identity: u64,
    document: u64,
    closed: AtomicBool,
    values: RwLock<Vec<Object>>,
}
impl Objects {
    pub(crate) fn new(identity: u64, document: u64, sheet: usize) -> Self {
        Self {
            identity,
            document,
            closed: AtomicBool::new(false),
            values: RwLock::new(vec![Object::Application, Object::Workbook, Object::Sheet(sheet)]),
        }
    }
    pub(crate) fn close(&self) {
        self.closed.store(true, Ordering::Relaxed);
    }
    pub(crate) fn get(&self, reference: ObjectRef) -> HostResult<Object> {
        if self.closed.load(Ordering::Relaxed) || reference.host != self.identity || reference.generation != self.document {
            return Err(failure(91, "stale or foreign object reference"));
        }
        let index = usize::try_from(reference.object).map_err(|_| failure(91, "invalid object index"))?;
        self.values
            .read()
            .map_err(|_| failure(91, "object store unavailable"))?
            .get(index)
            .cloned()
            .ok_or_else(|| failure(91, "invalid object reference"))
    }
    pub(crate) fn intern(&self, object: Object) -> HostResult<usize> {
        if self.closed.load(Ordering::Relaxed) {
            return Err(failure(91, "object store closed"));
        }
        let mut objects = self.values.write().map_err(|_| failure(91, "object store unavailable"))?;
        if let Some(index) = objects.iter().position(|o| o == &object) {
            return Ok(index);
        }
        if objects.len() >= MAX_OBJECTS {
            return Err(failure(7, "object handle budget exceeded"));
        }
        let index = objects.len();
        objects.push(object);
        Ok(index)
    }
}
impl ObjectTypeResolver for Objects {
    fn identity(&self) -> u64 {
        self.identity
    }
    fn supports_type(&self, reference: ObjectRef, expected: &NominalType) -> HostResult<bool> {
        let object = self.get(reference)?;
        let name = match object {
            Object::Application => "application",
            Object::Workbook => "workbook",
            Object::Sheet(_) => "worksheet",
            Object::Range(..) => "range",
        };
        Ok(expected.library() == "excel" && expected.version() == EXCEL_TYPES_VERSION && expected.name() == name)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
    #[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
    fn type_queries_use_current_handles_and_shared_store_lifetime() -> std::result::Result<(), Box<dyn std::error::Error>> {
        let objects = Objects::new(5, 7, 0);
        let sheet = ObjectRef { host: 5, object: 2, generation: 7 };
        let worksheet = NominalType::new("Excel", EXCEL_TYPES_VERSION, "Worksheet")?;
        let range = NominalType::new("Excel", EXCEL_TYPES_VERSION, "Range")?;
        assert!(objects.supports_type(sheet, &worksheet).map_err(|e| e.description)?);
        assert!(!objects.supports_type(sheet, &range).map_err(|e| e.description)?);
        assert!(objects.supports_type(ObjectRef { generation: 6, ..sheet }, &worksheet).is_err());
        assert!(objects.supports_type(ObjectRef { host: 6, ..sheet }, &worksheet).is_err());
        let index = objects.intern(Object::Range(0, gridcraft_engine::core::RangeRef::parse("A1").ok_or("range")?)).map_err(|e| e.description)?;
        assert!(objects.supports_type(ObjectRef { object: index as u64, ..sheet }, &range).map_err(|e| e.description)?);
        objects.close();
        assert!(objects.supports_type(sheet, &worksheet).is_err());
        Ok(())
    }
}
