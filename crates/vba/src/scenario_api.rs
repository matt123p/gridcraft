//! Native Scenario Manager projections; records retain their native opaque identity.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::{
    CellContent,
    cmd::whatif::{Scenario, scenario_objects},
    core::Value as CellValue,
};
use serde_json::{Value as Json, json};
fn from_json(value: &Json) -> HostResult<Value> {
    match value {
        Json::Null => Ok(Value::Empty),
        Json::Bool(b) => Ok(Value::Boolean(*b)),
        Json::Number(n) => Ok(Value::Double(n.as_f64().ok_or_else(|| failure(6, "scenario number overflow"))?)),
        Json::String(s) => values::vba_text(s),
        _ => {
            let error = value
                .get("error")
                .and_then(Json::as_str)
                .and_then(gridcraft_engine::core::CellError::parse)
                .ok_or_else(|| failure(13, "unsupported scenario value"))?;
            values::from_cell(CellValue::Error(error))
        }
    }
}
fn to_json(value: &Value) -> HostResult<Json> {
    match values::content(value, false)? {
        CellContent::Value(v) => Ok(match v {
            CellValue::Empty => Json::Null,
            CellValue::Bool(b) => json!(b),
            CellValue::Number(n) => json!(n),
            CellValue::Text(s) => json!(s.as_ref()),
            CellValue::Error(error) => json!({"error":error.as_str()}),
            _ => return Err(failure(13, "scenario scalar required")),
        }),
        _ => Err(failure(13, "scenario scalar required")),
    }
}
impl Bridge {
    pub(crate) fn scenario_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if let Object::Sheet(sheet) = object
            && call.member.eq_ignore_ascii_case("Scenarios")
            && reads(call)
        {
            let f = Feature::new("Scenarios", Some(*sheet));
            return Some(if call.arguments.is_empty() {
                self.feature(f)
            } else {
                let mut call = call.clone();
                call.member = "Item".into();
                self.scenario_dispatch(session, &f, &call)
            });
        }
        if let Object::Feature(f) = object
            && matches!(f.kind, "Scenarios" | "Scenario")
        {
            return Some(self.scenario_dispatch(session, f, call));
        }
        None
    }
    fn scenario_node(sheet: usize, sc: &Scenario) -> Feature {
        let mut f = Feature::new("Scenario", Some(sheet));
        f.key = sc.identity;
        f
    }
    fn scenario_dispatch(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "scenario worksheet deleted"))?;
        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = doc.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let sheet_name = sh.name.clone();
        let list = scenario_objects(&doc.wb).into_iter().filter(|sc| sc.sheet.eq_ignore_ascii_case(&sheet_name)).collect::<Vec<_>>();
        let member = call.member.to_ascii_lowercase();
        if reads(call) && member == "application" {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if reads(call) && member == "parent" {
            arguments(call, &[])?;
            return self.intern(Object::Sheet(sheet));
        }
        if f.kind == "Scenarios" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(list.iter().map(|sc| Self::scenario_node(sheet, sc).object()).collect());
            }
            match member.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(list.len() as i32));
                }
                "item" | "" => {
                    let args = arguments(call, &["Index"])?;
                    let index = required(&args, 0)?;
                    let sc = if matches!(index, Value::String(_)) {
                        let name = string(index)?;
                        list.iter().find(|sc| sc.name.eq_ignore_ascii_case(&name))
                    } else {
                        usize::try_from(numeric_index(index)?.checked_sub(1).ok_or_else(|| failure(9, "scenario index"))?)
                            .ok()
                            .and_then(|i| list.get(i))
                    }
                    .ok_or_else(|| failure(9, "scenario not found"))?;
                    return self.feature(Self::scenario_node(sheet, sc));
                }
                "add" => {
                    let args = arguments(call, &["Name", "ChangingCells", "Values", "Comment", "Locked", "Hidden"])?;
                    if !absent(&args[4]) || !absent(&args[5]) {
                        return Err(failure(1004, "scenario locking/visibility are unavailable"));
                    }
                    if sh.is_protected() {
                        return Err(failure(1004, "scenario changes require unprotected worksheet"));
                    }
                    let name = string(required(&args, 0)?)?;
                    let Value::Object(reference) = required(&args, 1)? else {
                        return Err(failure(13, "ChangingCells requires a Range"));
                    };
                    let Object::Range(owner, range) = self.object(session, *reference)? else {
                        return Err(failure(13, "ChangingCells requires a Range"));
                    };
                    if owner != sheet || range.count()? > 32 {
                        return Err(failure(1004, "scenario needs at most 32 cells on its worksheet"));
                    }
                    let mut params = json!({"name":name,"changing":range.areas.iter().map(|r|Self::qualified_rectangle(session,sheet,*r)).collect::<HostResult<Vec<_>>>()?.join(",")});
                    if !absent(&args[2]) {
                        let items = match &args[2] {
                            Value::Array(a) => a.elements().to_vec(),
                            v => vec![v.clone()],
                        };
                        if items.len() != range.count()? as usize {
                            return Err(failure(1004, "scenario values do not match changing cells"));
                        }
                        params["values"] = json!(items.iter().map(to_json).collect::<HostResult<Vec<_>>>()?);
                    }
                    if !absent(&args[3]) {
                        let comment = string(&args[3])?;
                        if comment.len() > 4096 {
                            return Err(failure(7, "scenario comment budget"));
                        }
                        params["comment"] = json!(comment);
                    }
                    self.objects.reserve(1)?;
                    session.execute("data.scenarioAdd", params).map_err(|e| failure(1004, e.to_string()))?;
                    let sc = scenario_objects(&session.doc().map_err(|e| failure(91, e.to_string()))?.wb)
                        .into_iter()
                        .find(|sc| sc.sheet.eq_ignore_ascii_case(&sheet_name) && sc.name.eq_ignore_ascii_case(&name))
                        .ok_or_else(|| failure(1004, "scenario not created"))?;
                    return self.feature(Self::scenario_node(sheet, &sc));
                }
                "createsummary" => {
                    let args = arguments(call, &["ReportType", "ResultCells"])?;
                    if !absent(&args[0]) && numeric_index(&args[0])? != 1 {
                        return Err(failure(1004, "only native scenario summary reports are available"));
                    }
                    let mut params = json!({"sheet":sheet_name});
                    if !absent(&args[1]) {
                        let (owner, r) = self.owned_rectangle(session, sheet, &args[1])?;
                        if owner != sheet {
                            return Err(failure(1004, "summary result cells belong to another worksheet"));
                        }
                        params["resultCells"] = json!(Self::qualified_rectangle(session, sheet, r)?);
                    }
                    session.execute("data.scenarioSummary", params).map_err(|e| failure(1004, e.to_string()))?;
                    self.sync_sheets(session)?;
                    return Ok(Value::Empty);
                }
                _ => return Err(failure(438, "unsupported Scenarios member")),
            }
        }
        let sc = list.into_iter().find(|sc| sc.identity == f.key).ok_or_else(|| failure(91, "scenario was deleted"))?;
        if reads(call) {
            match member.as_str() {
                "name" => {
                    arguments(call, &[])?;
                    return values::vba_text(&sc.name);
                }
                "comment" => {
                    arguments(call, &[])?;
                    return values::vba_text(&sc.comment);
                }
                "index" => {
                    arguments(call, &[])?;
                    let index = scenario_objects(&doc.wb)
                        .into_iter()
                        .filter(|sc| sc.sheet.eq_ignore_ascii_case(&sheet_name))
                        .position(|entry| entry.identity == f.key)
                        .ok_or_else(|| failure(91, "scenario deleted"))?;
                    return Ok(Value::Long(index as i32 + 1));
                }
                "changingcells" => {
                    arguments(call, &[])?;
                    let (_, r) = Self::resolve_text(session, sheet, &sc.changing, None, 0, &mut 0)?;
                    return self.intern(Object::Range(sheet, r));
                }
                "values" => {
                    let args = arguments(call, &["Index"])?;
                    if !absent(&args[0]) {
                        let i = usize::try_from(numeric_index(&args[0])?.checked_sub(1).ok_or_else(|| failure(9, "scenario value index"))?)
                            .map_err(|_| failure(9, "scenario value index"))?;
                        return from_json(sc.values.get(i).ok_or_else(|| failure(9, "scenario value index"))?);
                    }
                    let mut array = vba_values::VbaArray::new(vba_values::ScalarKind::Variant, vec![(1, sc.values.len() as i32)], false)
                        .map_err(|e| failure(7, e.to_string()))?;
                    for (i, v) in sc.values.iter().enumerate() {
                        array.set(i, from_json(v)?).map_err(|e| failure(7, e.to_string()))?;
                    }
                    return Ok(Value::Array(array));
                }
                "show" | "delete" => {
                    if sh.is_protected() {
                        return Err(failure(1004, "scenario changes require unprotected worksheet"));
                    }
                    arguments(call, &[])?;
                    let changes = if member == "show" {
                        let (_, r) = Self::resolve_text(session, sheet, &sc.changing, None, 0, &mut 0)?;
                        self.prepare_change(session, sheet, r)?
                    } else {
                        super::PreparedChanges::default()
                    };
                    session
                        .execute(
                            if member == "show" { "data.scenarioShow" } else { "data.scenarioDelete" },
                            json!({"name":sc.name,"sheet":sheet_name}),
                        )
                        .map_err(|e| failure(1004, e.to_string()))?;
                    self.queue_changes(changes);
                    return Ok(Value::Empty);
                }
                "changescenario" => {
                    if sh.is_protected() {
                        return Err(failure(1004, "scenario changes require unprotected worksheet"));
                    }
                    let args = arguments(call, &["ChangingCells", "Values"])?;
                    let changing = if absent(&args[0]) {
                        sc.changing.clone()
                    } else {
                        let Value::Object(reference) = &args[0] else {
                            return Err(failure(13, "ChangingCells requires a Range"));
                        };
                        let Object::Range(owner, r) = self.object(session, *reference)? else {
                            return Err(failure(13, "ChangingCells requires a Range"));
                        };
                        if owner != sheet || r.count()? > 32 {
                            return Err(failure(1004, "scenario ChangingCells owner/budget"));
                        }
                        r.areas.iter().map(|r| r.a1()).collect::<Vec<_>>().join(",")
                    };
                    let mut params = json!({"name":sc.name,"changing":changing.split(',').map(|r|format!("{}!{r}",gridcraft_engine::core::quote_sheet(&sheet_name))).collect::<Vec<_>>().join(","),"replace":true,"comment":sc.comment,"values":sc.values});
                    if !absent(&args[1]) {
                        let values = match &args[1] {
                            Value::Array(a) => a.elements().to_vec(),
                            v => vec![v.clone()],
                        };
                        params["values"] = json!(values.iter().map(to_json).collect::<HostResult<Vec<_>>>()?);
                    }
                    session.execute("data.scenarioAdd", params).map_err(|e| failure(1004, e.to_string()))?;
                    return Ok(Value::Empty);
                }
                _ => {}
            }
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(438, "unsupported Scenario member"));
        }
        if sh.is_protected() {
            return Err(failure(1004, "scenario changes require unprotected worksheet"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        let mut params = json!({"name":sc.name,"oldName":sc.name,"changing":sc.changing.split(',').map(|r|format!("{}!{r}",gridcraft_engine::core::quote_sheet(&sheet_name))).collect::<Vec<_>>().join(","),"replace":true,"comment":sc.comment,"values":sc.values});
        match member.as_str() {
            "name" => params["name"] = json!(string(value)?),
            "comment" => {
                let comment = string(value)?;
                if comment.len() > 4096 {
                    return Err(failure(7, "scenario comment budget"));
                }
                params["comment"] = json!(comment);
            }
            _ => return Err(failure(438, "unsupported Scenario assignment")),
        };
        session.execute("data.scenarioAdd", params).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    let entries = [
        ("Worksheet", "Scenarios", "Property Get Scenarios() As Scenarios"),
        ("Scenarios", "Count", "Property Get Count() As Long"),
        ("Scenarios", "Item", "Property Get Item(Index As Variant) As Scenario"),
        (
            "Scenarios",
            "Add",
            "Function Add(Name As String, ChangingCells As Range, Optional Values As Variant, Optional Comment As Variant, Optional Locked As Variant, Optional Hidden As Variant) As Scenario",
        ),
        ("Scenarios", "CreateSummary", "Sub CreateSummary(Optional ReportType As Variant, Optional ResultCells As Variant)"),
        ("Scenario", "Name", "Property Name As String"),
        ("Scenario", "Comment", "Property Comment As String"),
        ("Scenario", "Index", "Property Get Index() As Long"),
        ("Scenario", "ChangingCells", "Property Get ChangingCells() As Range"),
        ("Scenario", "Values", "Property Get Values(Optional Index As Variant) As Variant"),
        ("Scenario", "Show", "Sub Show()"),
        ("Scenario", "Delete", "Sub Delete()"),
        ("Scenario", "ChangeScenario", "Sub ChangeScenario(Optional ChangingCells As Variant, Optional Values As Variant)"),
        ("Scenarios", "Parent", "Property Get Parent() As Worksheet"),
        ("Scenarios", "Application", "Property Get Application() As Application"),
        ("Scenario", "Parent", "Property Get Parent() As Worksheet"),
        ("Scenario", "Application", "Property Get Application() As Application"),
    ];
    entries
        .into_iter()
        .map(|(owner, name, signature)| super::MemberDescriptor {
            owner,
            name,
            signature,
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/scenario_api.rs::Bridge::scenario_api",
        })
        .collect()
}
