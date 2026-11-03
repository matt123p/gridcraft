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
    pub label: Option<String>,
    pub part: i32,
    pub range: Option<ranges::Range>,
    pub identity: Option<SheetIdentity>,
    pub ancestors: Vec<SheetIdentity>,
    pub items: Vec<Feature>,
    pub span: Option<(u32, u32)>,
}
impl PartialEq for Feature {
    fn eq(&self, other: &Self) -> bool {
        self.kind == other.kind
            && self.sheet == other.sheet
            && self.text == other.text
            && self.label == other.label
            && self.key == other.key
            && self.part == other.part
            && self.range == other.range
            && self.span == other.span
            && self.items == other.items
            && self.ancestors.len() == other.ancestors.len()
            && self.ancestors.iter().zip(&other.ancestors).all(|(a, b)| a.same(b))
            && match (&self.identity, &other.identity) {
                (None, None) => true,
                (Some(a), Some(b)) => a.same(b),
                _ => false,
            }
    }
}
impl Feature {
    pub fn new(kind: &'static str, sheet: Option<usize>) -> Self {
        Self {
            kind,
            sheet,
            text: None,
            label: None,
            key: 0,
            part: 0,
            range: None,
            identity: None,
            ancestors: Vec::new(),
            items: Vec::new(),
            span: None,
        }
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
fn chart_lifetime(f: &Feature) -> Option<&SheetIdentity> {
    if matches!(
        f.kind,
        "Series" | "SeriesInterior" | "SeriesBorder" | "SeriesFormat" | "Trendlines" | "Trendline" | "TrendlineBorder" | "DataLabels" | "DataLabel"
    ) || matches!(f.kind, "FillFormat" | "LineFormat" | "ColorFormat") && matches!(f.text.as_deref(), Some("series" | "trendline"))
    {
        return f.ancestors.first();
    }
    if matches!(
        f.kind,
        "Chart"
            | "ChartObject"
            | "ChartTitle"
            | "Legend"
            | "SeriesCollection"
            | "Axes"
            | "Axis"
            | "AxisTitle"
            | "Gridlines"
            | "ChartArea"
            | "PlotArea"
    ) || f.kind == "Shape" && f.text.as_deref() == Some("chart")
    {
        return f.identity.as_ref();
    }
    None
}

impl Bridge {
    /// Native undo or an external chart.move can transfer the same chart lifetime.
    /// Repair only the requested chart graph; ordinary cell calls do not scan drawings.
    pub(crate) fn sync_chart_owner(&self, session: &Session, object: &Object) -> HostResult<()> {
        let Object::Feature(f) = object else {
            return Ok(());
        };
        let nodes = if f.kind == "ShapeRange" { f.items.iter().collect::<Vec<_>>() } else { vec![f] };
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        for node in nodes {
            let Some(identity) = chart_lifetime(node) else {
                continue;
            };
            let Some(source) = node.sheet else {
                continue;
            };
            if doc.wb.sheet(source).is_some_and(|sh| sh.charts.iter().any(|c| c.identity.same(identity))) {
                continue;
            }
            let target = doc
                .wb
                .sheets
                .iter()
                .enumerate()
                .flat_map(|(index, sh)| sh.charts.iter().map(move |c| (index, c)))
                .take(MAX_OBJECTS)
                .find(|(_, c)| c.identity.same(identity))
                .map(|(index, _)| index);
            if let Some(target) = target {
                let staged = self.objects.relocate_chart(doc.uid, source, target, identity)?;
                self.objects.replace(staged)?;
            }
        }
        Ok(())
    }
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
        if let Some(result) = self.pivot_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.slicer_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.drawing_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.chart_parts_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.series_api(session, object, call) {
            return Some(result);
        }
        if let Some(result) = self.chart_api(session, object, call) {
            return Some(result);
        }
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
