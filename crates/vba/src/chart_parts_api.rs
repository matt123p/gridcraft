//! Axes, titles, chart/plot areas and data labels are native computed/stored views.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::model::{Chart, ChartKind};
use serde_json::json;
fn child(f: &Feature, kind: &'static str, part: i32) -> Feature {
    let mut f = f.clone();
    f.kind = kind;
    f.part = part;
    if kind == "Chart" {
        f.span = None;
        f.text = None;
    }
    f
}
fn axes(kind: ChartKind) -> bool {
    !matches!(kind, ChartKind::Pie | ChartKind::Doughnut | ChartKind::Radar | ChartKind::Treemap | ChartKind::Sunburst | ChartKind::Funnel)
}
impl Bridge {
    pub(crate) fn chart_parts_api(&mut self, s: &mut Session, o: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        let Object::Feature(f) = o else {
            return None;
        };
        let m = call.member.to_ascii_lowercase();
        if f.kind == "Chart"
            && matches!(m.as_str(), "axes" | "hasaxis" | "chartarea" | "plotarea" | "applydatalabels" | "hastitle" | "name" | "location")
        {
            if m == "hastitle" {
                return None;
            }
            return Some(self.chart_parts(s, f, call));
        }
        if f.kind == "Series" && matches!(m.as_str(), "datalabels" | "applydatalabels" | "hasdatalabels") {
            return Some(self.chart_parts(s, f, call));
        }
        if !matches!(f.kind, "Axes" | "Axis" | "AxisTitle" | "Gridlines" | "ChartArea" | "PlotArea" | "DataLabels" | "DataLabel") {
            return None;
        }
        Some(self.chart_parts(s, f, call))
    }
    fn chart_parts(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet deleted"))?;
        let wb = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.clone();
        let id = f.ancestors.first().or(f.identity.as_ref()).ok_or_else(|| failure(91, "chart identity"))?;
        let c = wb
            .sheet(sheet)
            .and_then(|sh| sh.charts.iter().find(|c| u64::from(c.id) == f.key && id.same(&c.identity)))
            .ok_or_else(|| failure(91, "chart deleted"))?;
        if c.series.len() > 255 {
            return Err(failure(7, "chart series budget"));
        }
        let m = call.member.to_ascii_lowercase();
        let group = f.span.map_or(1, |s| s.0);
        let has_axis = axes(c.kind);
        let secondary = has_axis && c.series.iter().any(|se| se.secondary);
        let axis_count = if has_axis { 2 + i32::from(secondary) } else { 0 };
        if m == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if f.kind == "Chart" {
            return self.chart_elements(s, f, c, call);
        }
        if f.kind == "Axes" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                let mut nodes = if has_axis { (1..=2).map(|i| child(f, "Axis", i).object()).collect::<Vec<_>>() } else { vec![] };
                if secondary {
                    let mut n = child(f, "Axis", 2);
                    n.span = Some((2, 0));
                    nodes.push(n.object());
                }
                return self.object_array(nodes);
            }
            match m.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(axis_count));
                }
                "parent" => {
                    arguments(call, &[])?;
                    return self.feature(child(f, "Chart", 0));
                }
                "item" | "" | "axes" => {
                    let a = arguments(call, &["Type", "AxisGroup"])?;
                    let axis = numeric_index(required(&a, 0)?)?;
                    let g = if absent(&a[1]) { 1 } else { numeric_index(&a[1])? };
                    return self.axis_node(s, f, c, axis, g);
                }
                _ => return Err(failure(438, "unknown Axes member")),
            }
        }
        if matches!(f.kind, "ChartArea" | "PlotArea") {
            if !reads(call) {
                return Err(failure(1004, "native chart areas use automatic layout"));
            }
            arguments(call, &[])?;
            if m == "parent" {
                return self.feature(child(f, "Chart", 0));
            }
            let bounds = if f.kind == "ChartArea" {
                [0., 0., c.anchor.width, c.anchor.height]
            } else {
                native_layout(&wb, sheet, c)?.plot.ok_or_else(|| failure(1004, "no native plot area"))?
            };
            return match m.as_str() {
                "left" | "insideleft" => Ok(Value::Double(f64::from(bounds[0]))),
                "top" | "insidetop" => Ok(Value::Double(f64::from(bounds[1]))),
                "width" | "insidewidth" => Ok(Value::Double(f64::from(bounds[2]))),
                "height" | "insideheight" => Ok(Value::Double(f64::from(bounds[3]))),
                _ => Err(failure(1004, "no independently configurable chart area field")),
            };
        }
        if matches!(f.kind, "Series" | "DataLabels" | "DataLabel") {
            return self.data_labels(s, f, c, call);
        }
        if !has_axis {
            return Err(failure(91, "chart has no Cartesian axes"));
        }
        if f.kind == "AxisTitle" {
            let title = if f.part == 1 { c.x_title.as_deref() } else { c.y_title.as_deref() }.ok_or_else(|| failure(91, "axis title removed"))?;
            if m == "parent" && reads(call) {
                arguments(call, &[])?;
                return self.feature(child(f, "Axis", f.part));
            }
            if reads(call) && matches!(m.as_str(), "text" | "caption") {
                arguments(call, &[])?;
                return values::vba_text(title);
            }
            let text = if m == "delete" && reads(call) {
                arguments(call, &[])?;
                serde_json::Value::Null
            } else {
                let a = arguments(call, &["Value"])?;
                if call.kind != InvokeKind::Let || !matches!(m.as_str(), "text" | "caption") {
                    return Err(failure(1004, "only native axis title text is mutable"));
                }
                json!(string(required(&a, 0)?)?)
            };
            return self.chart_command(s, sheet, c.id, "chart.set", json!({if f.part==1{"xTitle"}else{"yTitle"}:text})).map(|()| Value::Empty);
        }
        if f.kind == "Gridlines" {
            if !c.gridlines {
                return Err(failure(91, "gridlines removed"));
            }
            arguments(call, &[])?;
            return if m == "parent" && reads(call) {
                self.feature(child(f, "Axis", f.part))
            } else {
                Err(failure(1004, "gridline styling is not independently represented"))
            };
        }
        let title = if f.part == 1 { c.x_title.as_ref() } else { c.y_title.as_ref() };
        if reads(call) {
            arguments(call, &[])?;
            return match m.as_str() {
                "parent" => self.feature(child(f, "Chart", 0)),
                "type" => Ok(Value::Long(f.part)),
                "axisgroup" => Ok(Value::Long(group as i32)),
                "hastitle" => Ok(Value::Boolean(group == 1 && title.is_some())),
                "axistitle" if group == 1 && title.is_some() => self.feature(child(f, "AxisTitle", f.part)),
                "hasmajorgridlines" => Ok(Value::Boolean(group == 1 && f.part == 2 && c.gridlines)),
                "majorgridlines" if group == 1 && f.part == 2 && c.gridlines => self.feature(child(f, "Gridlines", f.part)),
                "minimumscaleisauto" | "maximumscaleisauto" | "majorunitisauto" => Ok(Value::Boolean(true)),
                "minimumscale" | "maximumscale" | "majorunit" => {
                    let layout = native_layout(&wb, sheet, c)?;
                    let scale = if group == 2 {
                        layout.secondary
                    } else if f.part == 1 {
                        layout.category
                    } else {
                        layout.value
                    }
                    .ok_or_else(|| failure(1004, "categorical axis has no numeric scale"))?;
                    Ok(Value::Double(match m.as_str() {
                        "minimumscale" => scale.min,
                        "maximumscale" => scale.max,
                        _ => scale.step,
                    }))
                }
                _ => Err(failure(1004, "axis field unavailable in native automatic layout")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "axis property requires Let"));
        }
        if group != 1 {
            return Err(failure(1004, "secondary axis settings are automatic"));
        }
        let a = arguments(call, &["Value"])?;
        let on = super::address::boolean(required(&a, 0)?)?;
        let p = match m.as_str() {
            "hastitle" => {
                json!({if f.part==1{"xTitle"}else{"yTitle"}:if on{json!(title.map(String::as_str).unwrap_or("Axis Title"))}else{serde_json::Value::Null}})
            }
            "hasmajorgridlines" if f.part == 2 => json!({"gridlines":on}),
            _ => return Err(failure(1004, "native axes do not store manual scale/style options")),
        };
        self.chart_command(s, sheet, c.id, "chart.set", p)?;
        Ok(Value::Empty)
    }
    fn chart_elements(&mut self, s: &mut Session, f: &Feature, c: &Chart, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        if !reads(call) {
            return Err(failure(1004, "native chart element is read-only"));
        }
        let m = call.member.to_ascii_lowercase();
        match m.as_str() {
            "location" => {
                let a = arguments(call, &["Where", "Name"])?;
                if numeric_index(required(&a, 0)?)? != 2 {
                    return Err(failure(1004, "chart sheets have no native model"));
                }
                let name = string(required(&a, 1)?)?;
                let doc = s.doc().map_err(|e| failure(91, e.to_string()))?;
                let target = doc.wb.sheet_index(&name).ok_or_else(|| failure(9, "destination worksheet not found"))?;
                if target == sheet {
                    return self.feature(f.clone());
                }
                let dest = doc.wb.sheet(target).ok_or_else(|| failure(91, "sheet"))?;
                if doc.wb.sheet(sheet).is_none_or(|sh| sh.is_protected()) || dest.is_protected() {
                    return Err(failure(1004, "drawing objects are protected"));
                }
                if dest.charts.iter().any(|other| other.id == c.id) {
                    return Err(failure(1004, "native destination chart ID collision"));
                }
                let staged = self.objects.relocate_chart(self.context, sheet, target, &c.identity)?;
                s.execute("chart.move", json!({"sourceSheet":sheet,"sheet":name,"chart":c.id})).map_err(|e| failure(1004, e.to_string()))?;
                self.objects.replace(staged)?;
                let mut n = f.clone();
                n.sheet = Some(target);
                self.feature(n)
            }
            "name" => {
                arguments(call, &[])?;
                values::vba_text(&format!("Chart {}", c.id))
            }
            "chartarea" | "plotarea" => {
                arguments(call, &[])?;
                self.feature(child(f, if m == "chartarea" { "ChartArea" } else { "PlotArea" }, 0))
            }
            "axes" => {
                let a = arguments(call, &["Type", "AxisGroup"])?;
                if absent(&a[0]) && absent(&a[1]) {
                    self.feature(child(f, "Axes", 0))
                } else {
                    self.axis_node(s, f, c, numeric_index(required(&a, 0)?)?, if absent(&a[1]) { 1 } else { numeric_index(&a[1])? })
                }
            }
            "hasaxis" => {
                let a = arguments(call, &["Index1", "Index2"])?;
                let ty = if absent(&a[0]) { 2 } else { numeric_index(&a[0])? };
                let group = if absent(&a[1]) { 1 } else { numeric_index(&a[1])? };
                Ok(Value::Boolean(
                    axes(c.kind) && ((group == 1 && matches!(ty, 1 | 2)) || (group == 2 && ty == 2 && c.series.iter().any(|s| s.secondary))),
                ))
            }
            "applydatalabels" => {
                let a = arguments(
                    call,
                    &[
                        "Type",
                        "LegendKey",
                        "AutoText",
                        "HasLeaderLines",
                        "ShowSeriesName",
                        "ShowCategoryName",
                        "ShowValue",
                        "ShowPercentage",
                        "ShowBubbleSize",
                        "Separator",
                    ],
                )?;
                if (!absent(&a[0]) && numeric_index(&a[0])? != 2) || a.iter().skip(1).enumerate().any(|(i, v)| !absent(v) && i != 5) {
                    return Err(failure(1004, "only native value labels are supported"));
                }
                let on = absent(&a[6]) || super::address::boolean(&a[6])?;
                self.chart_command(s, sheet, c.id, "chart.set", json!({"dataLabels":on}))?;
                Ok(Value::Empty)
            }
            _ => Err(failure(438, "unsupported chart element")),
        }
    }
    fn axis_node(&mut self, s: &Session, f: &Feature, c: &Chart, axis: i32, group: i32) -> HostResult<Value> {
        if !axes(c.kind) || !matches!(axis, 1 | 2) || !matches!(group, 1 | 2) || group == 2 && (axis != 2 || !c.series.iter().any(|s| s.secondary)) {
            return Err(failure(1004, "axis is not present"));
        }
        let _ = s;
        let mut n = child(f, "Axis", axis);
        n.span = Some((group as u32, 0));
        self.feature(n)
    }
    fn data_labels(&mut self, s: &mut Session, f: &Feature, c: &Chart, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        let identity = f.ancestors.get(1).or(f.identity.as_ref()).ok_or_else(|| failure(91, "series"))?;
        let se = c.series.iter().find(|se| identity.same(&se.identity)).ok_or_else(|| failure(91, "series replaced"))?;
        let m = call.member.to_ascii_lowercase();
        if f.kind == "Series" {
            if m == "hasdatalabels" && reads(call) {
                arguments(call, &[])?;
                return Ok(Value::Boolean(c.data_labels));
            }
            if m == "datalabels" && reads(call) {
                if !c.data_labels {
                    return Err(failure(1004, "enable data labels first"));
                }
                let mut n = f.clone();
                n.kind = "DataLabels";
                return if call.arguments.is_empty() { self.feature(n) } else { self.data_labels(s, &n, c, call) };
            }
            if c.series.len() != 1 {
                return Err(failure(1004, "native labels are chart-wide; use Chart.ApplyDataLabels"));
            }
            let on = if m == "hasdatalabels" {
                let a = arguments(call, &["Value"])?;
                super::address::boolean(required(&a, 0)?)?
            } else {
                arguments(call, &[])?;
                true
            };
            return self.chart_command(s, sheet, c.id, "chart.set", json!({"dataLabels":on})).map(|()| Value::Empty);
        }
        if !c.data_labels {
            return Err(failure(91, "data labels removed"));
        }
        let data = super::series_api::binding_value(s, sheet, &se.values)?;
        let Value::Array(data) = data else {
            return Err(failure(13, "series data"));
        };
        let count = data.elements().len();
        if m == "parent" && reads(call) {
            arguments(call, &[])?;
            let mut n = f.clone();
            n.kind = "Series";
            n.part = 0;
            return self.feature(n);
        }
        if f.kind == "DataLabels" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                if count > super::MAX_OBJECTS {
                    return Err(failure(7, "data label collection budget"));
                }
                return self.object_array(
                    (0..count)
                        .map(|i| {
                            let mut n = f.clone();
                            n.kind = "DataLabel";
                            n.part = i as i32;
                            n.object()
                        })
                        .collect(),
                );
            }
            if m == "count" {
                arguments(call, &[])?;
                return Ok(Value::Long(count as i32));
            }
            if !matches!(m.as_str(), "item" | "" | "datalabels") {
                return Err(failure(438, "unknown DataLabels member"));
            }
            let a = arguments(call, &["Index"])?;
            let index = numeric_index(required(&a, 0)?)?.checked_sub(1).ok_or_else(|| failure(9, "label index"))?;
            if usize::try_from(index).ok().is_none_or(|i| i >= count) {
                return Err(failure(9, "label index"));
            }
            let mut n = f.clone();
            n.kind = "DataLabel";
            n.part = index;
            return self.feature(n);
        }
        let _value = data.elements().get(usize::try_from(f.part).map_err(|_| failure(91, "label"))?).ok_or_else(|| failure(91, "label removed"))?;
        arguments(call, &[])?;
        if !reads(call) {
            return Err(failure(1004, "native label text and placement are automatic"));
        }
        match m.as_str() {
            "text" | "caption" => {
                let data = gridcraft_chart::resolve(&s.doc().map_err(|e| failure(91, e.to_string()))?.wb, sheet, c);
                let index = c.series.iter().position(|entry| entry.identity.same(&se.identity)).ok_or_else(|| failure(91, "series"))?;
                let se = data.series.get(index).ok_or_else(|| failure(7, "native series render limit"))?;
                let text = se
                    .values
                    .get(f.part as usize)
                    .copied()
                    .flatten()
                    .map(|number| gridcraft_chart::format_number(number, &se.number_format, None))
                    .unwrap_or_default();
                values::vba_text(&text)
            }
            "showvalue" => Ok(Value::Boolean(true)),
            _ => Err(failure(1004, "no per-point label formatting state")),
        }
    }
}
fn native_layout(wb: &gridcraft_engine::model::Workbook, sheet: usize, c: &Chart) -> HostResult<gridcraft_chart::Layout> {
    if c.series.len() > 255 {
        return Err(failure(7, "chart layout budget"));
    }
    let data = gridcraft_chart::resolve(wb, sheet, c);
    Ok(gridcraft_chart::render_layout(c, &data, c.anchor.width, c.anchor.height, &gridcraft_chart::ApproxMeasure).1)
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Chart","Location","Function Location(Where As XlChartLocation, Optional Name As Variant) As Chart",super::MemberAvailability::Implemented),
        ("Chart","ChartArea","Property Get ChartArea() As ChartArea",super::MemberAvailability::Implemented),
        ("Chart","PlotArea","Property Get PlotArea() As PlotArea",super::MemberAvailability::Implemented),
        ("Chart","Name","Property Get Name() As String",super::MemberAvailability::Implemented),
        ("Chart","Axes","Function Axes(Optional Type As Variant, Optional AxisGroup As XlAxisGroup) As Object",super::MemberAvailability::Implemented),
        ("Chart","HasAxis","Property Get HasAxis(Optional Index1 As Variant, Optional Index2 As Variant) As Variant",super::MemberAvailability::Implemented),
        ("Chart","ApplyDataLabels","Sub ApplyDataLabels(Optional Type As XlDataLabelsType, Optional LegendKey As Variant, Optional AutoText As Variant, Optional HasLeaderLines As Variant, Optional ShowSeriesName As Variant, Optional ShowCategoryName As Variant, Optional ShowValue As Variant, Optional ShowPercentage As Variant, Optional ShowBubbleSize As Variant, Optional Separator As Variant)",super::MemberAvailability::Implemented),
        ("ChartArea","Parent","Property Get Parent() As Chart",super::MemberAvailability::Implemented),
        ("ChartArea","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("ChartArea","Left","Property Get Left() As Double",super::MemberAvailability::Implemented),
        ("ChartArea","Top","Property Get Top() As Double",super::MemberAvailability::Implemented),
        ("ChartArea","Width","Property Get Width() As Double",super::MemberAvailability::Implemented),
        ("ChartArea","Height","Property Get Height() As Double",super::MemberAvailability::Implemented),
        ("PlotArea","Parent","Property Get Parent() As Chart",super::MemberAvailability::Implemented),
        ("PlotArea","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("PlotArea","Left","Property Get Left() As Double",super::MemberAvailability::Implemented),
        ("PlotArea","Top","Property Get Top() As Double",super::MemberAvailability::Implemented),
        ("PlotArea","Width","Property Get Width() As Double",super::MemberAvailability::Implemented),
        ("PlotArea","Height","Property Get Height() As Double",super::MemberAvailability::Implemented),
        ("PlotArea","InsideLeft","Property Get InsideLeft() As Double",super::MemberAvailability::Implemented),
        ("PlotArea","InsideTop","Property Get InsideTop() As Double",super::MemberAvailability::Implemented),
        ("PlotArea","InsideWidth","Property Get InsideWidth() As Double",super::MemberAvailability::Implemented),
        ("PlotArea","InsideHeight","Property Get InsideHeight() As Double",super::MemberAvailability::Implemented),
        ("Axes","Parent","Property Get Parent() As Chart",super::MemberAvailability::Implemented),
        ("Axes","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Axes","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("Axes","Item","Property Get Item(Type As XlAxisType, Optional AxisGroup As XlAxisGroup) As Axis",super::MemberAvailability::Implemented),
        ("Axis","Parent","Property Get Parent() As Chart",super::MemberAvailability::Implemented),
        ("Axis","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Axis","Type","Property Get Type() As XlAxisType",super::MemberAvailability::Implemented),
        ("Axis","AxisGroup","Property Get AxisGroup() As XlAxisGroup",super::MemberAvailability::Implemented),
        ("Axis","HasTitle","Property HasTitle As Boolean",super::MemberAvailability::Implemented),
        ("Axis","AxisTitle","Property Get AxisTitle() As AxisTitle",super::MemberAvailability::Implemented),
        ("Axis","HasMajorGridlines","Property HasMajorGridlines As Boolean",super::MemberAvailability::Implemented),
        ("Axis","MajorGridlines","Property Get MajorGridlines() As Gridlines",super::MemberAvailability::Implemented),
        ("Axis","MinimumScale","Property Get MinimumScale() As Double",super::MemberAvailability::Implemented),
        ("Axis","MaximumScale","Property Get MaximumScale() As Double",super::MemberAvailability::Implemented),
        ("Axis","MajorUnit","Property Get MajorUnit() As Double",super::MemberAvailability::Implemented),
        ("Axis","MinimumScaleIsAuto","Property Get MinimumScaleIsAuto() As Boolean",super::MemberAvailability::Implemented),
        ("Axis","MaximumScaleIsAuto","Property Get MaximumScaleIsAuto() As Boolean",super::MemberAvailability::Implemented),
        ("Axis","MajorUnitIsAuto","Property Get MajorUnitIsAuto() As Boolean",super::MemberAvailability::Implemented),
        ("AxisTitle","Parent","Property Get Parent() As Axis",super::MemberAvailability::Implemented),
        ("AxisTitle","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("AxisTitle","Text","Property Text As String",super::MemberAvailability::Implemented),
        ("AxisTitle","Caption","Property Caption As String",super::MemberAvailability::Implemented),
        ("AxisTitle","Delete","Function Delete() As Variant",super::MemberAvailability::Implemented),
        ("Gridlines","Parent","Property Get Parent() As Axis",super::MemberAvailability::Implemented),
        ("Gridlines","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Series","HasDataLabels","Property HasDataLabels As Boolean",super::MemberAvailability::Implemented),
        ("Series","DataLabels","Function DataLabels(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("Series","ApplyDataLabels","Sub ApplyDataLabels()",super::MemberAvailability::Implemented),
        ("DataLabels","Parent","Property Get Parent() As Series",super::MemberAvailability::Implemented),
        ("DataLabels","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("DataLabels","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("DataLabels","Item","Function Item(Index As Variant) As DataLabel",super::MemberAvailability::Implemented),
        ("DataLabel","Parent","Property Get Parent() As Series",super::MemberAvailability::Implemented),
        ("DataLabel","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("DataLabel","Text","Property Get Text() As String",super::MemberAvailability::Implemented),
        ("DataLabel","Caption","Property Get Caption() As String",super::MemberAvailability::Implemented),
        ("DataLabel","ShowValue","Property Get ShowValue() As Boolean",super::MemberAvailability::Implemented),
    ].into_iter().map(|(owner,name,signature,availability)|super::MemberDescriptor{owner,name,signature,availability,implementation_symbol:"crates/vba/src/chart_parts_api.rs::Bridge::chart_parts_api"}).collect()
}
