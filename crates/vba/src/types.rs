//! Excel type definitions stay in GridCraft. tiny-vba sees opaque nominal identities.
use super::*;
use std::sync::{RwLock, atomic::AtomicBool};
use vba_foreign::{ObjectTypeResolver, ReferenceConstant, ReferenceLibrary, ReferenceSet, ReferenceType, ReferenceTypeKind};
use vba_values::NominalType;

// Adapter profile identity, not a claim about an installed Office typelib version.
const EXCEL_TYPES_VERSION: &str = "GridCraft.Excel.v1";

pub(crate) fn references() -> Result<ReferenceSet> {
    ReferenceSet::new(vec![ReferenceLibrary {
        name: "Excel".into(),
        version: EXCEL_TYPES_VERSION.into(),
        available: true,
        types: ["Application", "Workbook", "Worksheet", "Range", "Worksheets", "Sheets", "Areas", "Font", "Interior"]
            .into_iter()
            .map(|name| ReferenceType { name: name.into(), kind: ReferenceTypeKind::Object })
            .chain(
                [
                    "XlCalculation",
                    "XlReferenceStyle",
                    "XlDirection",
                    "XlFindLookIn",
                    "XlLookAt",
                    "XlSearchOrder",
                    "XlSearchDirection",
                    "XlInsertShiftDirection",
                    "XlDeleteShiftDirection",
                    "XlInsertFormatOrigin",
                ]
                .into_iter()
                .map(|name| ReferenceType { name: name.into(), kind: ReferenceTypeKind::Enum }),
            )
            .collect(),
        constants: [
            ("xlCalculationAutomatic", -4105),
            ("xlCalculationManual", -4135),
            ("xlCalculationSemiautomatic", 2),
            ("xlA1", 1),
            ("xlR1C1", -4150),
            ("xlUp", -4162),
            ("xlDown", -4121),
            ("xlToLeft", -4159),
            ("xlToRight", -4161),
            ("xlValues", -4163),
            ("xlFormulas", -4123),
            ("xlWhole", 1),
            ("xlPart", 2),
            ("xlByRows", 1),
            ("xlByColumns", 2),
            ("xlNext", 1),
            ("xlPrevious", 2),
            ("xlShiftDown", -4121),
            ("xlShiftToRight", -4161),
            ("xlShiftUp", -4162),
            ("xlShiftToLeft", -4159),
            ("xlFormatFromLeftOrAbove", 0),
            ("xlFormatFromRightOrBelow", 1),
        ]
        .into_iter()
        .map(|(name, n)| ReferenceConstant {
            name: name.into(),
            module: Some(
                match name {
                    "xlCalculationAutomatic" | "xlCalculationManual" | "xlCalculationSemiautomatic" => "XlCalculation",
                    "xlA1" | "xlR1C1" => "XlReferenceStyle",
                    "xlUp" | "xlDown" | "xlToLeft" | "xlToRight" => "XlDirection",
                    "xlValues" | "xlFormulas" => "XlFindLookIn",
                    "xlWhole" | "xlPart" => "XlLookAt",
                    "xlByRows" | "xlByColumns" => "XlSearchOrder",
                    "xlNext" | "xlPrevious" => "XlSearchDirection",
                    "xlShiftDown" | "xlShiftToRight" => "XlInsertShiftDirection",
                    "xlShiftUp" | "xlShiftToLeft" => "XlDeleteShiftDirection",
                    _ => "XlInsertFormatOrigin",
                }
                .into(),
            ),
            value: Value::Long(n),
        })
        .collect(),
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
            values: RwLock::new(
                vec![
                    Object::Application,
                    Object::Workbook,
                    Object::Sheet(sheet),
                    Object::Worksheets,
                    Object::Sheets,
                    Object::Range(sheet, gridcraft_engine::core::RangeRef::all().into()),
                    Object::Range(sheet, ranges::Range { areas: vec![gridcraft_engine::core::RangeRef::all()], unit: ranges::Unit::Rows }),
                    Object::Range(sheet, ranges::Range { areas: vec![gridcraft_engine::core::RangeRef::all()], unit: ranges::Unit::Columns }),
                ]
                .into_iter()
                .map(|object| if object == Object::Application { object } else { Object::Document(document, Box::new(object)) })
                .collect(),
            ),
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
        let object = self
            .values
            .read()
            .map_err(|_| failure(91, "object store unavailable"))?
            .get(index)
            .cloned()
            .ok_or_else(|| failure(91, "invalid object reference"))?;
        Ok(object)
    }
    pub(crate) fn invalidate_ranges(&self, document: u64) -> HostResult<()> {
        let mut objects = self.values.write().map_err(|_| failure(91, "object store unavailable"))?;
        for stored in objects.iter_mut() {
            if let Object::Document(doc, object) = stored
                && *doc == document
            {
                let areas = matches!(object.as_ref(), Object::Areas(..));
                if matches!(object.as_ref(), Object::Range(..) | Object::Areas(..) | Object::Format(..)) {
                    **object = Object::DeletedRange(areas);
                }
            }
        }
        Ok(())
    }
    pub(crate) fn remap_sheets(&self, document: u64, mapping: &[Option<usize>]) -> HostResult<()> {
        let mut objects = self.values.write().map_err(|_| failure(91, "object store unavailable"))?;
        for stored in objects.iter_mut() {
            let object = match stored {
                Object::Document(doc, object) if *doc == document => object.as_mut(),
                _ => continue,
            };
            let sheet = match object {
                Object::Sheet(s) | Object::Range(s, _) | Object::Areas(s, _) | Object::Format(s, _, _) => s,
                _ => continue,
            };
            if let Some(next) = mapping.get(*sheet).copied().flatten() {
                *sheet = next;
            } else {
                *sheet = usize::MAX;
            }
        }
        Ok(())
    }
    pub(crate) fn rebase(&self, document: u64, sheet: usize, edit: &gridcraft_engine::formula::adjust::Edit) -> HostResult<Vec<Object>> {
        use gridcraft_engine::formula::{self, Expr};
        let objects = self.values.read().map_err(|_| failure(91, "object store unavailable"))?;
        let mut result = Vec::new();
        result.try_reserve(objects.len()).map_err(|_| failure(7, "handle rebase allocation"))?;
        for stored in objects.iter() {
            let object = match stored {
                Object::Document(doc, object) if *doc == document => object.as_ref(),
                _ => {
                    result.push(stored.clone());
                    continue;
                }
            };
            let (owner, range, areas) = match object {
                Object::Range(s, r) => (*s, r, false),
                Object::Areas(s, r) => (*s, r, true),
                Object::Format(s, r, _) => (*s, r, false),
                _ => {
                    result.push(stored.clone());
                    continue;
                }
            };
            if owner != sheet {
                result.push(stored.clone());
                continue;
            }
            let mut r = range.clone();
            r.areas.clear();
            for rect in &range.areas {
                if *rect == gridcraft_engine::core::RangeRef::all() {
                    r.areas.push(*rect);
                    continue;
                }
                let expr = formula::parse(&rect.a1()).map_err(|_| failure(1004, "invalid saved range"))?;
                if let Expr::Ref(reference) = formula::adjust::adjust(expr, "sheet", "sheet", edit) {
                    r.areas.push(reference.range());
                }
            }
            if r.areas.is_empty() {
                result.push(Object::Document(document, Box::new(Object::DeletedRange(areas))));
            } else {
                r.validate()?;
                result.push(Object::Document(
                    document,
                    Box::new(if let Object::Format(_, _, font) = object {
                        Object::Format(owner, r, *font)
                    } else if areas {
                        Object::Areas(owner, r)
                    } else {
                        Object::Range(owner, r)
                    }),
                ));
            }
        }
        Ok(result)
    }
    pub(crate) fn replace(&self, values: Vec<Object>) -> HostResult<()> {
        *self.values.write().map_err(|_| failure(91, "object store unavailable"))? = values;
        Ok(())
    }
    pub(crate) fn intern_many(&self, values: Vec<Object>) -> HostResult<Vec<usize>> {
        if self.closed.load(Ordering::Relaxed) {
            return Err(failure(91, "object store closed"));
        }
        let mut objects = self.values.write().map_err(|_| failure(91, "object store unavailable"))?;
        // Preflight the whole snapshot before extending the existing registry.
        let missing = values.iter().filter(|v| !objects.contains(v)).count();
        if objects.len().checked_add(missing).is_none_or(|n| n > MAX_OBJECTS) {
            return Err(failure(7, "object handle budget exceeded"));
        }
        objects.try_reserve(missing).map_err(|_| failure(7, "object allocation"))?;
        let mut handles = Vec::new();
        handles.try_reserve(values.len()).map_err(|_| failure(7, "enumeration allocation"))?;
        for value in values {
            let i = if let Some(i) = objects.iter().position(|o| o == &value) {
                i
            } else {
                let i = objects.len();
                objects.push(value);
                i
            };
            handles.push(i);
        }
        Ok(handles)
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
        let object = match object {
            Object::Document(_, object) => *object,
            object => object,
        };
        let name = match object {
            Object::Application => "application",
            Object::Document(..) => return Err(failure(91, "invalid nested document owner")),
            Object::Workbook => "workbook",
            Object::Worksheets => "worksheets",
            Object::Sheets => "sheets",
            Object::Sheet(_) => "worksheet",
            Object::Range(..) => "range",
            Object::Areas(..) => "areas",
            Object::Format(_, _, font) => {
                if font {
                    "font"
                } else {
                    "interior"
                }
            }
            Object::DeletedRange(areas) => {
                if areas {
                    "areas"
                } else {
                    "range"
                }
            }
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
        let index =
            objects.intern(Object::Range(0, gridcraft_engine::core::RangeRef::parse("A1").ok_or("range")?.into())).map_err(|e| e.description)?;
        assert!(objects.supports_type(ObjectRef { object: index as u64, ..sheet }, &range).map_err(|e| e.description)?);
        objects.close();
        assert!(objects.supports_type(sheet, &worksheet).is_err());
        Ok(())
    }
}
