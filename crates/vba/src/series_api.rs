//! Native series and trendline views retain both chart and element lifetime identity.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::{
    core::Value as CellValue,
    model::{Chart, Series, Trendline, TrendlineKind},
};
use serde_json::json;
use vba_values::{ScalarKind, VbaArray};
pub(crate) fn series_node(kind: &'static str, sheet: usize, c: &Chart, se: &Series) -> Feature {
    let mut f = Feature::new(kind, Some(sheet));
    f.key = u64::from(c.id);
    f.identity = Some(se.identity.clone());
    f.ancestors.push(c.identity.clone());
    f
}
pub(crate) fn array(values: Vec<Value>) -> HostResult<Value> {
    if values.len() > super::MAX_CELLS {
        return Err(failure(7, "series value budget"));
    }
    let mut a = if values.is_empty() {
        VbaArray::empty(ScalarKind::Variant, 1)
    } else {
        VbaArray::new(ScalarKind::Variant, vec![(1, i32::try_from(values.len()).map_err(|_| failure(7, "array budget"))?)], false)
            .map_err(|e| failure(7, e.to_string()))?
    };
    for (i, v) in values.into_iter().enumerate() {
        a.set(i, v).map_err(|e| failure(13, e.to_string()))?;
    }
    Ok(Value::Array(a))
}
pub(crate) fn binding_value(s: &Session, sheet: usize, text: &str) -> HostResult<Value> {
    if text.len() > 65536 {
        return Err(failure(7, "series binding budget"));
    }
    let value = gridcraft_engine::calc::evaluate(
        &s.doc().map_err(|e| failure(91, e.to_string()))?.wb,
        sheet,
        gridcraft_engine::core::CellRef::default(),
        text,
    );
    match value {
        CellValue::Array(a) if a.data.len() > super::MAX_CELLS => Err(failure(7, "series value budget")),
        CellValue::Array(a) => array(a.data.iter().cloned().map(values::from_cell).collect::<HostResult<Vec<_>>>()?),
        other => array(vec![values::from_cell(other)?]),
    }
}
pub(crate) fn binding(s: &Session, value: &Value) -> HostResult<String> {
    if let Value::Object(_) = value {
        return Err(failure(13, "range binding requires host object resolution"));
    }
    if let Value::Array(a) = value {
        if a.elements().len() > 65536 {
            return Err(failure(7, "series array budget"));
        }
        let mut out = String::from("{");
        for (i, v) in a.elements().iter().enumerate() {
            if i > 0 {
                out.push(',');
            }
            match v {
                Value::String(t) => {
                    out.push('"');
                    out.push_str(&String::from_utf16(t.units()).map_err(|_| failure(13, "invalid text"))?.replace('"', "\"\""));
                    out.push('"');
                }
                Value::Boolean(v) => out.push_str(if *v { "TRUE" } else { "FALSE" }),
                Value::Empty => out.push('0'),
                _ => {
                    let cell = values::content(v, false)?;
                    let gridcraft_engine::CellContent::Value(cell) = cell else {
                        return Err(failure(13, "literal series values required"));
                    };
                    out.push_str(&cell.display());
                }
            }
            if out.len() > 65536 {
                return Err(failure(7, "series literal budget"));
            }
        }
        out.push('}');
        return Ok(out);
    }
    let _ = s;
    let t = string(value)?;
    if t.len() > 65536 {
        return Err(failure(7, "series binding budget"));
    }
    Ok(t.trim_start_matches('=').to_owned())
}
impl Bridge {
    pub(crate) fn range_binding(&self, s: &Session, value: &Value) -> HostResult<String> {
        if let Value::Object(r) = value {
            let Object::Range(sheet, range) = self.object(s, *r)? else {
                return Err(failure(13, "Range required"));
            };
            let rect = range.rectangle()?;
            if rect.count() > super::MAX_CELLS as u64 {
                return Err(failure(7, "series range budget"));
            }
            let name = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet deleted"))?.name;
            return Ok(format!("{}!{}", gridcraft_engine::formula::quote_sheet(name), rect.a1()));
        }
        binding(s, value)
    }
    pub(crate) fn series_api(&mut self, s: &mut Session, o: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        let Object::Feature(f) = o else {
            return None;
        };
        if f.kind == "Chart" && call.member.eq_ignore_ascii_case("SeriesCollection") && reads(call) {
            return Some({
                let mut n = f.clone();
                n.kind = "SeriesCollection";
                if call.arguments.is_empty() { self.feature(n) } else { self.series_dispatch(s, &n, call) }
            });
        }
        if !matches!(f.kind, "SeriesCollection" | "Series" | "SeriesInterior" | "SeriesBorder" | "Trendlines" | "Trendline" | "TrendlineBorder") {
            return None;
        }
        Some(self.series_dispatch(s, f, call))
    }
    fn series_dispatch(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet deleted"))?;
        let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = d.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet deleted"))?;
        let chart_identity = f.ancestors.first().or(f.identity.as_ref()).ok_or_else(|| failure(91, "chart identity"))?;
        let c =
            sh.charts.iter().find(|c| u64::from(c.id) == f.key && chart_identity.same(&c.identity)).ok_or_else(|| failure(91, "chart deleted"))?;
        if c.series.len() > 255 {
            return Err(failure(7, "series collection budget"));
        }
        if f.kind == "SeriesCollection" {
            return self.series_collection(s, f, call);
        }
        let si = f.ancestors.get(1).or(f.identity.as_ref()).ok_or_else(|| failure(91, "series identity"))?;
        let index = c.series.iter().position(|se| si.same(&se.identity)).ok_or_else(|| failure(91, "series replaced"))?;
        let se = c.series.get(index).ok_or_else(|| failure(91, "series deleted"))?;
        if [se.name.as_deref(), Some(se.values.as_str()), se.categories.as_deref(), se.bubble_sizes.as_deref()]
            .into_iter()
            .flatten()
            .any(|t| t.len() > 65536)
        {
            return Err(failure(7, "series metadata budget"));
        }
        if matches!(f.kind, "Trendlines" | "Trendline" | "TrendlineBorder") {
            return self.trendline_dispatch(s, f, index, call);
        }
        let member = call.member.to_ascii_lowercase();
        if member == "parent" && reads(call) {
            arguments(call, &[])?;
            let mut parent = Feature::new(if f.kind == "Series" { "Chart" } else { "Series" }, Some(sheet));
            parent.key = f.key;
            if f.kind == "Series" {
                parent.identity = f.ancestors.first().cloned();
            } else {
                parent = f.clone();
                parent.kind = "Series";
                parent.part = 0;
                parent.text = None;
            }
            return self.feature(parent);
        }
        if member == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if matches!(f.kind, "SeriesInterior" | "SeriesBorder") {
            if member != "color" {
                return Err(failure(438, "only native series color is available"));
            }
            if reads(call) {
                arguments(call, &[])?;
                return super::format_api::ole(
                    se.color.unwrap_or_else(|| {
                        let [r, g, b, _] = gridcraft_chart::series_color(&d.wb.theme, index);
                        gridcraft_engine::model::Color::rgb(r, g, b)
                    }),
                    s,
                );
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(438, "color requires Let"));
            }
            let a = arguments(call, &["Value"])?;
            let color = super::format_api::rgb(required(&a, 0)?)?;
            let color = color.resolve(&d.wb.theme).ok_or_else(|| failure(13, "color"))?;
            return self
                .chart_command(
                    s,
                    sheet,
                    c.id,
                    "chart.formatSelection",
                    json!({"series":index,"color":format!("#{:02X}{:02X}{:02X}",color[0],color[1],color[2])}),
                )
                .map(|()| Value::Empty);
        }
        if member == "trendlines" && reads(call) {
            let mut n = f.clone();
            n.kind = "Trendlines";
            return if call.arguments.is_empty() { self.feature(n) } else { self.trendline_dispatch(s, &n, index, call) };
        }
        if matches!(member.as_str(), "interior" | "border") && reads(call) {
            arguments(call, &[])?;
            let mut n = f.clone();
            n.kind = if member == "interior" { "SeriesInterior" } else { "SeriesBorder" };
            return self.feature(n);
        }
        if member == "delete" && reads(call) {
            arguments(call, &[])?;
            if c.series.len() == 1 {
                return Err(failure(1004, "native charts require a series"));
            }
            return self.chart_command(s, sheet, c.id, "chart.selectData", json!({"removeSeries":[index]})).map(|()| Value::Empty);
        }
        if reads(call) {
            arguments(call, &[])?;
            return match member.as_str() {
                "name" => {
                    let name = se
                        .name
                        .as_deref()
                        .map(|name| gridcraft_chart::resolve_series_name(&d.wb, sheet, name))
                        .filter(|n| !n.is_empty())
                        .unwrap_or_else(|| format!("Series{}", index + 1));
                    values::vba_text(&name)
                }
                "values" => binding_value(s, sheet, &se.values),
                "xvalues" => binding_value(s, sheet, se.categories.as_deref().unwrap_or("{}")),
                "bubblesizes" => binding_value(s, sheet, se.bubble_sizes.as_deref().unwrap_or("{}")),
                "formula" => values::vba_text(&format!(
                    "=SERIES({},{},{},{})",
                    se.name.as_deref().unwrap_or(""),
                    se.categories.as_deref().unwrap_or(""),
                    se.values,
                    index + 1
                )),
                "plotorder" => Ok(Value::Long(index as i32 + 1)),
                "quartilecalculationinclusivemedian" => Ok(Value::Boolean(se.quartiles == gridcraft_engine::model::QuartileMethod::Inclusive)),
                "charttype" => Ok(Value::Long(super::chart_api::type_number(gridcraft_chart::series_plot_kind(c, index)))),
                "axisgroup" => Ok(Value::Long(if se.secondary { 2 } else { 1 })),
                "smooth" => Ok(Value::Boolean(se.smooth)),
                "hasdatalabels" => Ok(Value::Boolean(c.data_labels)),
                _ => Err(failure(438, "unsupported Series member")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "unsupported series invocation"));
        }
        let a = arguments(call, &["Value"])?;
        let v = required(&a, 0)?;
        let id = c.id;
        let mut patch = json!({});
        match member.as_str() {
            "name" => patch["name"] = json!(string(v)?),
            "values" | "xvalues" | "bubblesizes" => {
                patch[match member.as_str() {
                    "values" => "values",
                    "xvalues" => "categories",
                    _ => "bubbleSizes",
                }] = json!(self.range_binding(s, v)?)
            }
            "smooth" => patch["smooth"] = json!(super::address::boolean(v)?),
            "quartilecalculationinclusivemedian" => patch["inclusiveQuartiles"] = json!(super::address::boolean(v)?),
            "plotorder" => {
                let order = numeric_index(v)?;
                if order < 1 || order as usize > c.series.len() {
                    return Err(failure(1004, "plot order bounds"));
                }
                patch["plotOrder"] = json!(order - 1);
            }
            "axisgroup" => {
                patch["secondary"] = json!(match numeric_index(v)? {
                    1 => false,
                    2 => true,
                    _ => return Err(failure(1004, "invalid AxisGroup")),
                })
            }
            "charttype" => {
                let number = numeric_index(v)?;
                let (_, ty, sub) = super::chart_api::CHART_TYPES
                    .iter()
                    .find(|(k, _, _)| super::chart_api::type_number(*k) == number)
                    .ok_or_else(|| failure(1004, "unsupported series type"))?;
                patch["type"] = json!(ty);
                patch["subtype"] = json!(sub);
            }
            _ => return Err(failure(1004, "this series option has no native mutation")),
        }
        self.chart_command(s, sheet, id, "chart.selectData", json!({"series":index,"seriesPatch":patch}))?;
        Ok(Value::Empty)
    }
    fn series_collection(&mut self, s: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet deleted"))?;
        let doc = s.doc().map_err(|e| failure(91, e.to_string()))?;
        let c = doc
            .wb
            .sheet(sheet)
            .and_then(|sh| sh.charts.iter().find(|c| u64::from(c.id) == f.key && f.identity.as_ref().is_some_and(|id| id.same(&c.identity))))
            .ok_or_else(|| failure(91, "chart deleted"))?;
        if call.kind == InvokeKind::Enumerate {
            arguments(call, &[])?;
            return self.object_array(c.series.iter().map(|se| series_node("Series", sheet, c, se).object()).collect());
        }
        if !reads(call) {
            return Err(failure(438, "series collection is read-only"));
        }
        match call.member.to_ascii_lowercase().as_str() {
            "count" => {
                arguments(call, &[])?;
                Ok(Value::Long(c.series.len() as i32))
            }
            "parent" => {
                arguments(call, &[])?;
                let mut n = f.clone();
                n.kind = "Chart";
                self.feature(n)
            }
            "application" => {
                arguments(call, &[])?;
                self.intern(Object::Application)
            }
            "item" | "" | "seriescollection" => {
                let a = arguments(call, &["Index"])?;
                let v = required(&a, 0)?;
                let se = if matches!(v, Value::String(_)) {
                    let text = string(v)?;
                    c.series.iter().find(|se| {
                        se.name.as_deref().is_some_and(|n| {
                            n.eq_ignore_ascii_case(&text) || gridcraft_chart::resolve_series_name(&doc.wb, sheet, n).eq_ignore_ascii_case(&text)
                        })
                    })
                } else {
                    usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "series index"))?).ok().and_then(|i| c.series.get(i))
                }
                .ok_or_else(|| failure(9, "series not found"))?;
                self.feature(series_node("Series", sheet, c, se))
            }
            "add" => {
                let a = arguments(call, &["Source", "Rowcol", "SeriesLabels", "CategoryLabels", "Replace"])?;
                if a.iter().skip(2).any(|v| !absent(v)) {
                    return Err(failure(1004, "native series insertion infers headers and categories"));
                }
                let source = self.range_binding(s, required(&a, 0)?)?;
                if !matches!(required(&a, 0)?, Value::Object(_)) {
                    return Err(failure(13, "Source must be a Range"));
                }
                let mut p = json!({"range":source,"append":true});
                if !absent(&a[1]) {
                    p["byRows"] = json!(match numeric_index(&a[1])? {
                        1 => true,
                        2 => false,
                        _ => return Err(failure(1004, "invalid Rowcol")),
                    });
                }
                let id = c.id;
                let before = c.series.len();
                self.objects.reserve(1)?;
                self.chart_command(s, sheet, id, "chart.selectData", p)?;
                let c = s
                    .doc()
                    .map_err(|e| failure(91, e.to_string()))?
                    .wb
                    .sheet(sheet)
                    .and_then(|sh| sh.charts.iter().find(|c| c.id == id))
                    .ok_or_else(|| failure(91, "chart"))?;
                let se = c.series.get(before).ok_or_else(|| failure(1004, "source has no series"))?;
                self.feature(series_node("Series", sheet, c, se))
            }
            _ => Err(failure(1004, "use SetSourceData to create native data-backed series")),
        }
    }
    fn trendline_dispatch(&mut self, s: &mut Session, f: &Feature, series_index: usize, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet deleted"))?;
        let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
        let c = d.wb.sheet(sheet).and_then(|sh| sh.charts.iter().find(|c| u64::from(c.id) == f.key)).ok_or_else(|| failure(91, "chart deleted"))?;
        let se = c.series.get(series_index).ok_or_else(|| failure(91, "series deleted"))?;
        if se.trendlines.len() > 16 {
            return Err(failure(7, "trendline budget"));
        }
        let member = call.member.to_ascii_lowercase();
        if f.kind == "Trendlines" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(se.trendlines.iter().map(|t| trend_node(sheet, c, se, t).object()).collect());
            }
            if !reads(call) {
                return Err(failure(438, "trendline collection is read-only"));
            }
            match member.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(se.trendlines.len() as i32));
                }
                "parent" => {
                    arguments(call, &[])?;
                    return self.feature(series_node("Series", sheet, c, se));
                }
                "application" => {
                    arguments(call, &[])?;
                    return self.intern(Object::Application);
                }
                "item" | "" | "trendlines" => {
                    let a = arguments(call, &["Index"])?;
                    let i = usize::try_from(numeric_index(required(&a, 0)?)?.checked_sub(1).ok_or_else(|| failure(9, "trendline index"))?)
                        .map_err(|_| failure(9, "trendline index"))?;
                    let t = se.trendlines.get(i).ok_or_else(|| failure(9, "trendline not found"))?;
                    return self.feature(trend_node(sheet, c, se, t));
                }
                "add" => return self.add_trendline(s, f, series_index, call),
                _ => return Err(failure(438, "unsupported Trendlines member")),
            }
        }
        let index = se
            .trendlines
            .iter()
            .position(|t| f.identity.as_ref().is_some_and(|id| id.same(&t.identity)))
            .ok_or_else(|| failure(91, "trendline deleted"))?;
        let t = se.trendlines.get(index).ok_or_else(|| failure(91, "trendline deleted"))?;
        if member == "border" && reads(call) {
            arguments(call, &[])?;
            let mut n = f.clone();
            n.kind = "TrendlineBorder";
            return self.feature(n);
        }
        if f.kind == "TrendlineBorder" {
            if member == "parent" && reads(call) {
                arguments(call, &[])?;
                let mut n = f.clone();
                n.kind = "Trendline";
                n.part = 0;
                n.text = None;
                return self.feature(n);
            }
            if member == "application" && reads(call) {
                arguments(call, &[])?;
                return self.intern(Object::Application);
            }
            if reads(call) {
                arguments(call, &[])?;
                return match member.as_str() {
                    "color" => super::format_api::ole(
                        t.color.or(se.color).unwrap_or_else(|| {
                            let [r, g, b, _] = gridcraft_chart::series_color(&d.wb.theme, series_index);
                            gridcraft_engine::model::Color::rgb(r, g, b)
                        }),
                        s,
                    ),
                    "weight" => Ok(Value::Double(f64::from(t.width.unwrap_or(1.5)))),
                    "dashstyle" => Ok(Value::Long(dash_number(t.dash.unwrap_or_default()))),
                    _ => Err(failure(1004, "native trendline line option unavailable")),
                };
            }
            if call.kind != InvokeKind::Let {
                return Err(failure(438, "line property requires Let"));
            }
            let a = arguments(call, &["Value"])?;
            let v = required(&a, 0)?;
            let mut p = json!({"element":"trendline","series":series_index,"trendline":index});
            match member.as_str() {
                "color" => p["color"] = json!(super::format_api::rgb(v)?),
                "weight" => {
                    let n = super::format_api::number(v)?;
                    if !(0.25..=12.0).contains(&n) {
                        return Err(failure(1004, "trendline width bounds"));
                    }
                    p["width"] = json!(n);
                }
                "dashstyle" => p["dash"] = json!(dash_kind(numeric_index(v)?)?),
                _ => return Err(failure(1004, "line option unavailable")),
            }
            self.chart_command(s, sheet, c.id, "chart.addElement", p)?;
            return Ok(Value::Empty);
        }

        if member == "parent" && reads(call) {
            arguments(call, &[])?;
            return self.feature(series_node("Series", sheet, c, se));
        }
        if member == "application" && reads(call) {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if member == "delete" && reads(call) {
            arguments(call, &[])?;
            return self
                .chart_command(s, sheet, c.id, "chart.addElement", json!({"element":"trendline","series":series_index,"trendline":index,"on":false}))
                .map(|()| Value::Empty);
        }
        if reads(call) {
            arguments(call, &[])?;
            return match member.as_str() {
                "type" => Ok(Value::Long(trend_type(t.kind))),
                "order" => Ok(Value::Long(if let TrendlineKind::Polynomial(n) = t.kind { i32::from(n) } else { 2 })),
                "period" => Ok(Value::Long(if let TrendlineKind::MovingAverage(n) = t.kind {
                    i32::try_from(n).map_err(|_| failure(6, "trendline count overflow"))?
                } else {
                    2
                })),
                "forward" | "forward2" => Ok(Value::Double(t.forward)),
                "backward" | "backward2" => Ok(Value::Double(t.backward)),
                "intercept" => Ok(Value::Double(t.intercept.unwrap_or(0.0))),
                "interceptisauto" => Ok(Value::Boolean(t.intercept.is_none())),
                "displayequation" => Ok(Value::Boolean(t.display_equation)),
                "displayrsquared" => Ok(Value::Boolean(t.display_r_squared)),
                "name" => {
                    let series_name = se
                        .name
                        .as_deref()
                        .map(|n| gridcraft_chart::resolve_series_name(&d.wb, sheet, n))
                        .filter(|n| !n.is_empty())
                        .unwrap_or_else(|| format!("Series{}", series_index + 1));
                    values::vba_text(t.name.as_deref().unwrap_or(&gridcraft_chart::trendline_name(t.kind, &series_name)))
                }
                "nameisauto" => Ok(Value::Boolean(t.name.is_none())),
                "index" => Ok(Value::Long(index as i32 + 1)),
                _ => Err(failure(438, "unsupported trendline field")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "trendline property requires Let"));
        }
        let a = arguments(call, &["Value"])?;
        let v = required(&a, 0)?;
        let mut p = json!({"element":"trendline","series":series_index,"trendline":index});
        match member.as_str() {
            "name" => p["name"] = json!(string(v)?),
            "nameisauto" if super::address::boolean(v)? => p["name"] = serde_json::Value::Null,
            "type" => p["type"] = json!(trend_name(numeric_index(v)?)?),
            "order" => {
                let n = numeric_index(v)?;
                if !(2..=6).contains(&n) {
                    return Err(failure(1004, "polynomial order"));
                }
                p["type"] = json!("polynomial");
                p["order"] = json!(n);
            }
            "period" => {
                let n = numeric_index(v)?;
                if !(2..=255).contains(&n) {
                    return Err(failure(1004, "moving average period"));
                }
                p["type"] = json!("movingAverage");
                p["period"] = json!(n);
            }
            "forward" | "backward" | "forward2" | "backward2" => {
                let n = super::format_api::number(v)?;
                if !(0.0..=Trendline::MAX_FORECAST).contains(&n) {
                    return Err(failure(1004, "forecast bounds"));
                }
                p[member.trim_end_matches('2')] = json!(n);
            }
            "intercept" => p["intercept"] = json!(super::format_api::number(v)?),
            "interceptisauto" if super::address::boolean(v)? => p["intercept"] = serde_json::Value::Null,
            "displayequation" => p["displayEquation"] = json!(super::address::boolean(v)?),
            "displayrsquared" => p["displayRSquared"] = json!(super::address::boolean(v)?),
            _ => return Err(failure(1004, "unsupported trendline option")),
        }
        self.chart_command(s, sheet, c.id, "chart.addElement", p)?;
        Ok(Value::Empty)
    }
    fn add_trendline(&mut self, s: &mut Session, f: &Feature, series_index: usize, call: &HostCall) -> HostResult<Value> {
        let a = arguments(call, &["Type", "Order", "Period", "Forward", "Backward", "Intercept", "DisplayEquation", "DisplayRSquared", "Name"])?;
        let mut p = json!({"element":"trendline","series":series_index});
        if !absent(&a[0]) {
            p["type"] = json!(trend_name(numeric_index(&a[0])?)?);
        }
        for (i, key, lo, hi) in [(1, "order", 2, 6), (2, "period", 2, 255)] {
            if !absent(&a[i]) {
                let n = numeric_index(&a[i])?;
                if !(lo..=hi).contains(&n) {
                    return Err(failure(1004, "trendline bounds"));
                }
                p[key] = json!(n);
            }
        }
        for (i, key) in [(3, "forward"), (4, "backward"), (5, "intercept")] {
            if !absent(&a[i]) {
                let n = super::format_api::number(&a[i])?;
                if i != 5 && !(0.0..=Trendline::MAX_FORECAST).contains(&n) {
                    return Err(failure(1004, "forecast bounds"));
                }
                p[key] = json!(n);
            }
        }
        for (i, key) in [(6, "displayEquation"), (7, "displayRSquared")] {
            if !absent(&a[i]) {
                p[key] = json!(super::address::boolean(&a[i])?);
            }
        }
        if !absent(&a[8]) {
            p["name"] = json!(string(&a[8])?);
        }
        let sheet = f.sheet.ok_or_else(|| failure(91, "sheet"))?;
        self.objects.reserve(1)?;
        if s.doc()
            .map_err(|e| failure(91, e.to_string()))?
            .wb
            .sheet(sheet)
            .and_then(|sh| sh.charts.iter().find(|c| u64::from(c.id) == f.key))
            .and_then(|c| c.series.get(series_index))
            .is_none_or(|se| se.trendlines.len() >= 16)
        {
            return Err(failure(7, "trendline budget"));
        }
        self.chart_command(s, sheet, u32::try_from(f.key).map_err(|_| failure(91, "chart"))?, "chart.addElement", p)?;
        let c = s
            .doc()
            .map_err(|e| failure(91, e.to_string()))?
            .wb
            .sheet(sheet)
            .and_then(|sh| sh.charts.iter().find(|c| u64::from(c.id) == f.key))
            .ok_or_else(|| failure(91, "chart"))?;
        let se = c.series.get(series_index).ok_or_else(|| failure(91, "series"))?;
        let t = se.trendlines.last().ok_or_else(|| failure(1004, "trendline not added"))?;
        self.feature(trend_node(sheet, c, se, t))
    }
}
fn trend_node(sheet: usize, c: &Chart, se: &Series, t: &Trendline) -> Feature {
    let mut f = series_node("Trendline", sheet, c, se);
    f.identity = Some(t.identity.clone());
    f.ancestors.push(se.identity.clone());
    f
}
fn trend_type(k: TrendlineKind) -> i32 {
    match k {
        TrendlineKind::Linear => -4132,
        TrendlineKind::Exponential => 5,
        TrendlineKind::Logarithmic => -4133,
        TrendlineKind::Power => 4,
        TrendlineKind::Polynomial(_) => 3,
        TrendlineKind::MovingAverage(_) => 6,
    }
}
fn trend_name(n: i32) -> HostResult<&'static str> {
    match n {
        -4132 => Ok("linear"),
        5 => Ok("exponential"),
        -4133 => Ok("logarithmic"),
        4 => Ok("power"),
        3 => Ok("polynomial"),
        6 => Ok("movingAverage"),
        _ => Err(failure(1004, "unsupported trendline type")),
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Series","QuartileCalculationInclusiveMedian","Property QuartileCalculationInclusiveMedian As Boolean",super::MemberAvailability::Implemented),
        ("Chart","SeriesCollection","Property Get SeriesCollection() As Object",super::MemberAvailability::Implemented),
        ("SeriesCollection","Parent","Property Get Parent() As Chart",super::MemberAvailability::Implemented),
        ("SeriesCollection","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("SeriesCollection","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("SeriesCollection","Item","Property Get Item(Index As Variant) As Series",super::MemberAvailability::Implemented),
        ("SeriesCollection","Add","Function Add(Source As Range, Optional Rowcol As Variant, Optional SeriesLabels As Variant, Optional CategoryLabels As Variant, Optional Replace As Variant) As Series",super::MemberAvailability::Implemented),
        ("SeriesCollection","NewSeries","Function NewSeries() As Series",super::MemberAvailability::OutOfScope),
        ("Series","Parent","Property Get Parent() As Chart",super::MemberAvailability::Implemented),
        ("Series","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Series","Name","Property Name As String",super::MemberAvailability::Implemented),
        ("Series","Values","Property Values As Variant",super::MemberAvailability::Implemented),
        ("Series","XValues","Property XValues As Variant",super::MemberAvailability::Implemented),
        ("Series","BubbleSizes","Property BubbleSizes As Variant",super::MemberAvailability::Implemented),
        ("Series","Formula","Property Get Formula() As String",super::MemberAvailability::Implemented),
        ("Series","PlotOrder","Property PlotOrder As Long",super::MemberAvailability::Implemented),
        ("Series","ChartType","Property ChartType As XlChartType",super::MemberAvailability::Implemented),
        ("Series","AxisGroup","Property AxisGroup As XlAxisGroup",super::MemberAvailability::Implemented),
        ("Series","Smooth","Property Smooth As Boolean",super::MemberAvailability::Implemented),
        ("Series","Trendlines","Function Trendlines(Optional Index As Variant) As Object",super::MemberAvailability::Implemented),
        ("Series","Delete","Function Delete() As Variant",super::MemberAvailability::Implemented),
        ("Trendlines","Parent","Property Get Parent() As Series",super::MemberAvailability::Implemented),
        ("Trendlines","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Trendlines","Count","Property Get Count() As Long",super::MemberAvailability::Implemented),
        ("Trendlines","Item","Property Get Item(Index As Long) As Trendline",super::MemberAvailability::Implemented),
        ("Trendlines","Add","Function Add(Optional Type As Variant, Optional Order As Variant, Optional Period As Variant, Optional Forward As Variant, Optional Backward As Variant, Optional Intercept As Variant, Optional DisplayEquation As Variant, Optional DisplayRSquared As Variant, Optional Name As Variant) As Trendline",super::MemberAvailability::Implemented),
        ("Trendline","Parent","Property Get Parent() As Series",super::MemberAvailability::Implemented),
        ("Trendline","Application","Property Get Application() As Application",super::MemberAvailability::Implemented),
        ("Trendline","Type","Property Type As XlTrendlineType",super::MemberAvailability::Implemented),
        ("Trendline","Order","Property Order As Long",super::MemberAvailability::Implemented),
        ("Trendline","Period","Property Period As Long",super::MemberAvailability::Implemented),
        ("Trendline","Forward2","Property Forward2 As Double",super::MemberAvailability::Implemented),
        ("Trendline","Backward2","Property Backward2 As Double",super::MemberAvailability::Implemented),
        ("Trendline","Intercept","Property Intercept As Double",super::MemberAvailability::Implemented),
        ("Trendline","InterceptIsAuto","Property InterceptIsAuto As Boolean",super::MemberAvailability::Implemented),
        ("Trendline","DisplayEquation","Property DisplayEquation As Boolean",super::MemberAvailability::Implemented),
        ("Trendline","DisplayRSquared","Property DisplayRSquared As Boolean",super::MemberAvailability::Implemented),
        ("Trendline","Name","Property Name As String",super::MemberAvailability::Implemented),
        ("Trendline","NameIsAuto","Property NameIsAuto As Boolean",super::MemberAvailability::Implemented),
        ("Trendline","Index","Property Get Index() As Long",super::MemberAvailability::Implemented),
        ("Trendline","Delete","Function Delete() As Variant",super::MemberAvailability::Implemented),
    ].into_iter().map(|(owner,name,signature,availability)|super::MemberDescriptor{owner,name,signature,availability,implementation_symbol:"crates/vba/src/series_api.rs::Bridge::series_api"}).collect()
}
fn dash_kind(n: i32) -> HostResult<gridcraft_engine::model::LineDash> {
    use gridcraft_engine::model::LineDash as D;
    match n {
        1 => Ok(D::Solid),
        2 => Ok(D::SysDot),
        3 => Ok(D::Dot),
        4 => Ok(D::Dash),
        5 => Ok(D::DashDot),
        6 => Ok(D::SysDashDotDot),
        7 => Ok(D::LgDash),
        8 => Ok(D::LgDashDot),
        _ => Err(failure(1004, "dash style unavailable")),
    }
}
fn dash_number(n: gridcraft_engine::model::LineDash) -> i32 {
    use gridcraft_engine::model::LineDash as D;
    match n {
        D::Solid => 1,
        D::Dot => 3,
        D::SysDot => 2,
        D::Dash | D::SysDash => 4,
        D::DashDot | D::SysDashDot => 5,
        D::SysDashDotDot | D::LgDashDotDot => 6,
        D::LgDash => 7,
        D::LgDashDot => 8,
    }
}
