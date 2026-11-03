//! OM12 first slice: live embedded chart containers and native chart settings.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::model::{Anchor, AnchorMode, ChartKind, LegendPos, sheet::SheetIdentity};
use serde_json::{Value as Json, json};

/// Owned scalar metadata only: ordinary property calls never clone native series payloads.
struct ChartView {
    id: u32,
    identity: SheetIdentity,
    kind: ChartKind,
    anchor: Anchor,
    title: Option<String>,
    legend: LegendPos,
    style: u32,
    source: Option<String>,
    by_rows: bool,
}
fn node(kind: &'static str, sheet: usize, id: u32, identity: &SheetIdentity) -> Feature {
    let mut f = Feature::new(kind, Some(sheet));
    f.key = u64::from(id);
    f.identity = Some(identity.clone());
    f
}
fn chart_name(id: u32) -> String {
    format!("Chart {id}")
}
fn has_title(chart: &ChartView) -> bool {
    chart.title.as_ref().is_some_and(|t| !t.is_empty())
}
impl Bridge {
    pub(crate) fn chart_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if let Object::Sheet(sheet) = object
            && call.member.eq_ignore_ascii_case("ChartObjects")
            && reads(call)
        {
            let f = Feature::new("ChartObjects", Some(*sheet));
            return Some(if call.arguments.is_empty() { self.feature(f) } else { self.chart_dispatch(session, &f, call) });
        }
        let Object::Feature(f) = object else {
            return None;
        };
        if !matches!(f.kind, "ChartObjects" | "ChartObject" | "Chart" | "ChartTitle" | "Legend") {
            return None;
        }
        Some(self.chart_dispatch(session, f, call))
    }
    fn chart_dispatch(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "chart worksheet deleted"))?;
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "chart worksheet deleted"))?;
        if sh.charts.len() > super::MAX_OBJECTS {
            return Err(failure(7, "chart collection budget"));
        }
        if f.kind == "ChartObjects" {
            return self.chart_collection(session, sheet, call);
        }
        let chart = sh
            .charts
            .iter()
            .find(|c| u64::from(c.id) == f.key && f.identity.as_ref().is_some_and(|id| id.same(&c.identity)))
            .ok_or_else(|| failure(91, "chart was deleted or replaced"))?;
        if chart.title.as_ref().is_some_and(|s| s.len() > 65536) || chart.source.as_ref().is_some_and(|s| s.len() > 65536) {
            return Err(failure(7, "chart metadata budget"));
        }
        let chart = ChartView {
            id: chart.id,
            identity: chart.identity.clone(),
            kind: chart.kind,
            anchor: chart.anchor,
            title: chart.title.clone(),
            legend: chart.legend,
            style: chart.style,
            source: chart.source.clone(),
            by_rows: chart.by_rows,
        };
        if f.kind == "ChartTitle" && !has_title(&chart) || f.kind == "Legend" && chart.legend == LegendPos::None {
            return Err(failure(91, "chart child is unavailable"));
        }
        let member = call.member.to_ascii_lowercase();
        if reads(call) {
            return self.chart_read(session, f, &chart, &member, call);
        }
        if call.kind == InvokeKind::Let {
            return self.chart_write(session, f, &chart, &member, call);
        }
        Err(failure(438, "unsupported chart invocation"))
    }
    fn chart_collection(&mut self, session: &mut Session, sheet: usize, call: &HostCall) -> HostResult<Value> {
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if !reads(call) && call.kind != InvokeKind::Enumerate {
            return Err(failure(438, "ChartObjects is read-only"));
        }
        if call.kind == InvokeKind::Enumerate {
            arguments(call, &[])?;
            return self.object_array(sh.charts.iter().map(|c| node("ChartObject", sheet, c.id, &c.identity).object()).collect());
        }
        match call.member.to_ascii_lowercase().as_str() {
            "count" => {
                arguments(call, &[])?;
                Ok(Value::Long(i32::try_from(sh.charts.len()).map_err(|_| failure(6, "chart Count overflow"))?))
            }
            "parent" => {
                arguments(call, &[])?;
                self.intern(Object::Sheet(sheet))
            }
            "application" => {
                arguments(call, &[])?;
                self.intern(Object::Application)
            }
            "item" | "" | "chartobjects" => {
                let args = arguments(call, &["Index"])?;
                let index = required(&args, 0)?;
                let chart = if matches!(index, Value::String(_)) {
                    let name = string(index)?;
                    sh.charts.iter().find(|c| chart_name(c.id).eq_ignore_ascii_case(&name))
                } else {
                    usize::try_from(numeric_index(index)?.checked_sub(1).ok_or_else(|| failure(9, "chart index"))?)
                        .ok()
                        .and_then(|i| sh.charts.get(i))
                }
                .ok_or_else(|| failure(9, "chart not found"))?;
                self.feature(node("ChartObject", sheet, chart.id, &chart.identity))
            }
            "delete" => {
                arguments(call, &[])?;
                let ids = sh.charts.iter().map(|c| c.id).collect::<Vec<_>>();
                for id in ids {
                    self.chart_command(session, sheet, id, "chart.delete", json!({}))?;
                }
                Ok(Value::Empty)
            }
            "add" => Err(failure(1004, "blank-chart creation has no native command; create a data-backed chart through insert.chart")),
            _ => Err(failure(438, "unsupported ChartObjects member")),
        }
    }
    fn chart_read(&mut self, session: &mut Session, f: &Feature, chart: &ChartView, member: &str, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "chart worksheet deleted"))?;
        let kind = f.kind;
        if member == "setsourcedata" && kind == "Chart" {
            return self.chart_source(session, sheet, chart, call);
        }
        arguments(call, &[])?;
        if member == "application" {
            return self.intern(Object::Application);
        }
        if member == "parent" {
            return if kind == "ChartObject" {
                self.intern(Object::Sheet(sheet))
            } else {
                self.feature(node(if kind == "Chart" { "ChartObject" } else { "Chart" }, sheet, chart.id, &chart.identity))
            };
        }
        if member == "delete" && matches!(kind, "ChartObject" | "ChartTitle" | "Chart" | "Legend") {
            self.chart_command(
                session,
                sheet,
                chart.id,
                if matches!(kind, "ChartObject" | "Chart") { "chart.delete" } else { "chart.set" },
                if kind == "Legend" { json!({"legend":"none"}) } else { json!({"title":null}) },
            )?;
            return Ok(Value::Empty);
        }
        if kind == "Chart" && member == "refresh" {
            session.calculate_workbook().map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if kind == "ChartObject" && member == "index" {
            let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "sheet"))?;
            return Ok(Value::Long(sh.charts.iter().position(|c| c.identity.same(&chart.identity)).ok_or_else(|| failure(91, "chart"))? as i32 + 1));
        }
        match (kind, member) {
            ("ChartObject", "chart") => self.feature(node("Chart", sheet, chart.id, &chart.identity)),
            ("ChartObject", "name") => values::vba_text(&chart_name(chart.id)),
            ("ChartObject", "width") => Ok(Value::Double(f64::from(chart.anchor.width))),
            ("ChartObject", "height") => Ok(Value::Double(f64::from(chart.anchor.height))),
            ("ChartObject", "left" | "top") => {
                let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
                Ok(Value::Double(if member == "left" {
                    sh.col_left(chart.anchor.cell.col) + f64::from(chart.anchor.dx)
                } else {
                    sh.row_top(chart.anchor.cell.row) + f64::from(chart.anchor.dy)
                }))
            }
            ("ChartObject", "placement") => Ok(Value::Long(match chart.anchor.mode {
                AnchorMode::MoveAndSize => 1,
                AnchorMode::MoveOnly => 2,
                AnchorMode::Absolute => 3,
            })),
            ("ChartObject", "topleftcell") => self.intern(Object::Range(sheet, gridcraft_engine::core::RangeRef::cell(chart.anchor.cell).into())),
            ("Chart", "charttype") => Ok(Value::Long(type_number(chart.kind))),
            ("Chart", "chartstyle") => Ok(Value::Long(i32::try_from(chart.style).map_err(|_| failure(6, "ChartStyle overflow"))?)),
            ("Chart", "plotby") => Ok(Value::Long(if chart.by_rows { 1 } else { 2 })),
            ("Chart", "hastitle") => Ok(Value::Boolean(has_title(chart))),
            ("Chart", "haslegend") => Ok(Value::Boolean(chart.legend != LegendPos::None)),
            ("Chart", "charttitle") if has_title(chart) => self.feature(node("ChartTitle", sheet, chart.id, &chart.identity)),
            ("Chart", "legend") if chart.legend != LegendPos::None => self.feature(node("Legend", sheet, chart.id, &chart.identity)),
            ("Chart", "charttitle" | "legend") => Err(failure(1004, "enable the chart element first")),
            ("ChartTitle", "text" | "caption") => values::vba_text(chart.title.as_deref().unwrap_or("")),
            ("Legend", "position") => Ok(Value::Long(match chart.legend {
                LegendPos::Bottom => -4107,
                LegendPos::Top => -4160,
                LegendPos::Left => -4131,
                LegendPos::Right => -4152,
                LegendPos::None => -4142,
            })),
            _ => Err(failure(438, "unsupported chart member")),
        }
    }
    fn chart_write(&mut self, session: &mut Session, f: &Feature, chart: &ChartView, member: &str, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "chart worksheet deleted"))?;
        let kind = f.kind;
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        let mut params = json!({});
        match (kind, member) {
            ("Chart", "charttype") => {
                let requested = numeric_index(value)?;
                let (_, ty, sub) = CHART_TYPES
                    .iter()
                    .find(|(kind, _, _)| type_number(*kind) == requested)
                    .map(|(kind, ty, sub)| (kind, ty, sub))
                    .ok_or_else(|| failure(1004, "ChartType is outside the native chart profile"))?;
                params["type"] = json!(ty);
                params["subtype"] = json!(sub);
            }
            ("Chart", "chartstyle") => {
                let n = numeric_index(value)?;
                if !(1..=16).contains(&n) {
                    return Err(failure(1004, "native ChartStyle must be 1..16"));
                }
                params["style"] = json!(n);
            }
            ("Chart", "hastitle") => {
                params["title"] = if super::address::boolean(value)? {
                    json!(chart.title.as_deref().filter(|t| !t.is_empty()).unwrap_or("Chart Title"))
                } else {
                    Json::Null
                };
            }
            ("Chart", "haslegend") => {
                params["legend"] = json!(if super::address::boolean(value)? {
                    match chart.legend {
                        LegendPos::Top => "top",
                        LegendPos::Left => "left",
                        LegendPos::Right => "right",
                        _ => "bottom",
                    }
                } else {
                    "none"
                });
            }
            ("ChartTitle", "text" | "caption") => {
                let text = string(value)?;
                if text.len() > 65536 {
                    return Err(failure(7, "chart title text budget"));
                }
                params["title"] = json!(text);
            }
            ("Legend", "position") => {
                params["legend"] = json!(match numeric_index(value)? {
                    -4107 => "bottom",
                    -4160 => "top",
                    -4131 => "left",
                    -4152 => "right",
                    _ => return Err(failure(1004, "unsupported Legend.Position")),
                });
            }
            ("Chart", "plotby") => {
                let by_rows = plot_by(value)?;
                let source = chart.source.as_deref().ok_or_else(|| failure(1004, "set a native chart source before changing PlotBy"))?;
                self.chart_command(session, sheet, chart.id, "chart.selectData", json!({"range":source,"byRows":by_rows}))?;
                return Ok(Value::Empty);
            }
            ("ChartObject", "placement") => {
                let mode = match numeric_index(value)? {
                    1 => "moveAndSize",
                    2 => "moveOnly",
                    3 => "absolute",
                    _ => return Err(failure(1004, "invalid Placement")),
                };
                self.chart_command(session, sheet, chart.id, "object.setAnchorMode", json!({"kind":"chart","id":chart.id,"mode":mode}))?;
                return Ok(Value::Empty);
            }
            ("ChartObject", "width" | "height" | "left" | "top") => {
                let n = super::format_api::number(value)?;
                let size = matches!(member, "width" | "height");
                if !(if size { 20.0..=10000.0 } else { 0.0..=1.0e9 }).contains(&n) {
                    return Err(failure(1004, "geometry is outside native bounds"));
                }
                if size {
                    params[member] = json!(n);
                } else {
                    let sh =
                        session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
                    let mut cell = chart.anchor.cell;
                    let offset = if member == "left" {
                        cell.col = sh.col_at(n);
                        sh.col_left(cell.col)
                    } else {
                        cell.row = sh.row_at(n);
                        sh.row_top(cell.row)
                    };
                    params["at"] = json!(cell.a1());
                    params[if member == "left" { "dx" } else { "dy" }] = json!(n - offset);
                }
            }
            _ => return Err(failure(438, "unsupported chart property write")),
        }
        self.chart_command(session, sheet, chart.id, "chart.set", params)?;
        Ok(Value::Empty)
    }
    fn chart_source(&mut self, session: &mut Session, sheet: usize, chart: &ChartView, call: &HostCall) -> HostResult<Value> {
        if call.kind != InvokeKind::Call {
            return Err(failure(438, "SetSourceData requires Call"));
        }
        let args = arguments(call, &["Source", "PlotBy"])?;
        let Value::Object(reference) = required(&args, 0)? else {
            return Err(failure(13, "Source must be a Range"));
        };
        let Object::Range(source_sheet, range) = self.object(session, *reference)? else {
            return Err(failure(13, "Source must be a Range in this workbook"));
        };
        let rect = range.rectangle()?;
        if rect.count() > super::MAX_CELLS as u64 {
            return Err(failure(7, "chart source cell budget"));
        }
        let name = &session
            .doc()
            .map_err(|e| failure(91, e.to_string()))?
            .wb
            .sheet(source_sheet)
            .ok_or_else(|| failure(91, "source worksheet deleted"))?
            .name;
        let mut params = json!({"range":format!("{}!{}", gridcraft_engine::formula::quote_sheet(name),rect.a1())});
        if !absent(&args[1]) {
            params["byRows"] = json!(plot_by(&args[1])?);
        }
        self.chart_command(session, sheet, chart.id, "chart.selectData", params)?;
        Ok(Value::Empty)
    }
    pub(crate) fn chart_command(&mut self, session: &mut Session, sheet: usize, id: u32, command: &str, mut params: Json) -> HostResult<()> {
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if sh.is_protected() {
            return Err(failure(1004, "worksheet drawing objects are protected"));
        }
        if sh.charts.iter().filter(|c| c.id == id).count() != 1 {
            return Err(failure(1004, "native chart ID is absent or ambiguous"));
        }
        params["chart"] = json!(id);
        params["sheet"] = json!(sheet);
        session.execute(command, params).map_err(|e| failure(1004, e.to_string()))?;
        Ok(())
    }
}
fn plot_by(value: &Value) -> HostResult<bool> {
    match numeric_index(value)? {
        1 => Ok(true),
        2 => Ok(false),
        _ => Err(failure(1004, "PlotBy must be xlRows or xlColumns")),
    }
}

pub(crate) const CHART_TYPES: &[(ChartKind, &str, &str)] = &[
    (ChartKind::ColumnClustered, "column", ""),
    (ChartKind::ColumnStacked, "column", "stacked"),
    (ChartKind::ColumnStacked100, "column", "stacked100"),
    (ChartKind::BarClustered, "bar", ""),
    (ChartKind::BarStacked, "bar", "stacked"),
    (ChartKind::BarStacked100, "bar", "stacked100"),
    (ChartKind::Line, "line", ""),
    (ChartKind::LineMarkers, "line", "markers"),
    (ChartKind::LineStacked, "line", "stacked"),
    (ChartKind::Pie, "pie", ""),
    (ChartKind::Doughnut, "doughnut", ""),
    (ChartKind::Area, "area", ""),
    (ChartKind::AreaStacked, "area", "stacked"),
    (ChartKind::AreaStacked100, "area", "stacked100"),
    (ChartKind::Scatter, "scatter", ""),
    (ChartKind::ScatterLines, "scatter", "lines"),
    (ChartKind::Bubble, "bubble", ""),
    (ChartKind::Radar, "radar", ""),
    (ChartKind::Histogram, "histogram", ""),
    (ChartKind::Waterfall, "waterfall", ""),
    (ChartKind::Funnel, "funnel", ""),
    (ChartKind::Treemap, "treemap", ""),
    (ChartKind::Sunburst, "sunburst", ""),
    (ChartKind::BoxWhisker, "boxWhisker", ""),
    (ChartKind::Stock, "stock", ""),
    (ChartKind::Combo, "combo", ""),
];
pub(crate) fn type_number(kind: ChartKind) -> i32 {
    match kind {
        ChartKind::ColumnClustered => 51,
        ChartKind::ColumnStacked => 52,
        ChartKind::ColumnStacked100 => 53,
        ChartKind::BarClustered => 57,
        ChartKind::BarStacked => 58,
        ChartKind::BarStacked100 => 59,
        ChartKind::Line => 4,
        ChartKind::LineMarkers => 65,
        ChartKind::LineStacked => 63,
        ChartKind::Pie => 5,
        ChartKind::Doughnut => -4120,
        ChartKind::Area => 1,
        ChartKind::AreaStacked => 76,
        ChartKind::AreaStacked100 => 77,
        ChartKind::Scatter => -4169,
        ChartKind::ScatterLines => 74,
        ChartKind::Bubble => 15,
        ChartKind::Radar => -4151,
        ChartKind::Histogram => 118,
        ChartKind::Waterfall => 119,
        ChartKind::Funnel => 123,
        ChartKind::Treemap => 117,
        ChartKind::Sunburst => 120,
        ChartKind::BoxWhisker => 121,
        ChartKind::Stock => 88,
        ChartKind::Combo => -4152,
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("ChartObjects", "Delete", "Function Delete() As Variant"),
        ("Chart", "Delete", "Function Delete() As Variant"),
        ("Chart", "Refresh", "Sub Refresh()"),
        ("ChartObject", "Index", "Property Get Index() As Long"),
        ("Legend", "Delete", "Function Delete() As Variant"),
        ("Worksheet", "ChartObjects", "Function ChartObjects(Optional Index As Variant) As Object"),
        ("ChartObjects", "Count", "Property Get Count() As Long"),
        ("ChartObjects", "Item", "Property Get Item(Index As Variant) As ChartObject"),
        ("ChartObjects", "Parent", "Property Get Parent() As Worksheet"),
        ("ChartObjects", "Application", "Property Get Application() As Application"),
        ("ChartObject", "Chart", "Property Get Chart() As Chart"),
        ("ChartObject", "Name", "Property Get Name() As String"),
        ("ChartObject", "Parent", "Property Get Parent() As Worksheet"),
        ("ChartObject", "Application", "Property Get Application() As Application"),
        ("ChartObject", "Delete", "Function Delete() As Variant"),
        ("ChartObject", "TopLeftCell", "Property Get TopLeftCell() As Range"),
        ("ChartObject", "Left", "Property Left As Double"),
        ("ChartObject", "Top", "Property Top As Double"),
        ("ChartObject", "Width", "Property Width As Double"),
        ("ChartObject", "Height", "Property Height As Double"),
        ("ChartObject", "Placement", "Property Placement As XlPlacement"),
        ("Chart", "Parent", "Property Get Parent() As ChartObject"),
        ("Chart", "Application", "Property Get Application() As Application"),
        ("Chart", "ChartTitle", "Property Get ChartTitle() As ChartTitle"),
        ("Chart", "Legend", "Property Get Legend() As Legend"),
        ("Chart", "SetSourceData", "Sub SetSourceData(Source As Range, Optional PlotBy As XlRowCol)"),
        ("Chart", "ChartType", "Property ChartType As XlChartType"),
        ("Chart", "ChartStyle", "Property ChartStyle As Variant"),
        ("Chart", "PlotBy", "Property PlotBy As XlRowCol"),
        ("Chart", "HasTitle", "Property HasTitle As Boolean"),
        ("Chart", "HasLegend", "Property HasLegend As Boolean"),
        ("ChartTitle", "Text", "Property Text As String"),
        ("ChartTitle", "Caption", "Property Caption As String"),
        ("ChartTitle", "Parent", "Property Get Parent() As Chart"),
        ("ChartTitle", "Application", "Property Get Application() As Application"),
        ("ChartTitle", "Delete", "Function Delete() As Variant"),
        ("Legend", "Position", "Property Position As XlLegendPosition"),
        ("Legend", "Parent", "Property Get Parent() As Chart"),
        ("Legend", "Application", "Property Get Application() As Application"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/chart_api.rs::Bridge::chart_api",
    })
    .chain(std::iter::once(super::MemberDescriptor {
        owner: "ChartObjects",
        name: "Add",
        signature: "Function Add(Left As Double, Top As Double, Width As Double, Height As Double) As ChartObject",
        availability: super::MemberAvailability::OutOfScope,
        implementation_symbol: "crates/vba/src/chart_api.rs::Bridge::chart_collection",
    }))
    .collect()
}
