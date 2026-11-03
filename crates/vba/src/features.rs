//! Private family nodes share the bridge's single document-qualified registry.
use super::{Bridge, HostCall, HostResult, InvokeKind, MAX_OBJECTS, Object, Session, Value, failure, ranges};
use gridcraft_engine::model::sheet::SheetIdentity;
use vba_values::{ScalarKind, VbaArray};

#[derive(Clone, Debug)]
pub(crate) struct Feature {
    pub kind: &'static str,
    pub sheet: Option<usize>,
    pub key: u64,
    pub text: Option<String>,
    pub part: i32,
    pub range: Option<ranges::Range>,
    pub identity: Option<SheetIdentity>,
}
impl PartialEq for Feature {
    fn eq(&self, other: &Self) -> bool {
        self.kind == other.kind
            && self.sheet == other.sheet
            && self.text == other.text
            && self.key == other.key
            && self.part == other.part
            && self.range == other.range
            && match (&self.identity, &other.identity) {
                (None, None) => true,
                (Some(a), Some(b)) => a.same(b),
                _ => false,
            }
    }
}
impl Feature {
    pub fn new(kind: &'static str, sheet: Option<usize>) -> Self {
        Self { kind, sheet, text: None, key: 0, part: 0, range: None, identity: None }
    }
    pub fn object(self) -> Object {
        Object::Feature(self)
    }
}
pub(crate) fn reads(call: &HostCall) -> bool {
    matches!(call.kind, InvokeKind::Get | InvokeKind::Call | InvokeKind::DefaultGet)
}
pub(crate) fn absent(value: &Value) -> bool {
    matches!(value, Value::Missing)
}
impl Bridge {
    pub(crate) fn feature(&mut self, feature: Feature) -> HostResult<Value> {
        self.intern(feature.object())
    }
    pub(crate) fn object_array(&mut self, objects: Vec<Object>) -> HostResult<Value> {
        if objects.len() > MAX_OBJECTS {
            return Err(failure(7, "collection budget"));
        }
        let handles = self.objects.intern_many(
            objects
                .into_iter()
                .map(|o| if matches!(o, Object::Application | Object::Document(..)) { o } else { Object::Document(self.context, Box::new(o)) })
                .collect(),
        )?;
        let mut array = if handles.is_empty() {
            VbaArray::empty(ScalarKind::Object, 1)
        } else {
            VbaArray::new(ScalarKind::Object, vec![(1, handles.len() as i32)], false).map_err(|e| failure(7, e.to_string()))?
        };
        for (i, handle) in handles.into_iter().enumerate() {
            array.set(i, Value::Object(self.reference(handle))).map_err(|e| failure(7, e.to_string()))?;
        }
        Ok(Value::Array(array))
    }
    pub(crate) fn feature_dispatch(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if let Some(result) = self.protection_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.formula_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.named_style_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.format_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.conditional_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.outline_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.data_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.table_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.whatif_api(session, object, call) {
            return Some(result);
        }
        self.scenario_api(session, object, call)
    }
}
