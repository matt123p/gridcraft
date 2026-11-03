//! Remaining native analysis/data tools; no calculation algorithms are implemented here.
use super::features::{absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, Object, Session, Value, failure, values};
use gridcraft_engine::core::RangeRef;
use serde_json::json;
fn aggregate(n: i32) -> HostResult<&'static str> {
    match n {
        -4157 => Ok("sum"),
        -4106 => Ok("average"),
        -4112 => Ok("count"),
        -4113 => Ok("countNums"),
        -4136 => Ok("max"),
        -4139 => Ok("min"),
        -4149 => Ok("product"),
        _ => Err(failure(1004, "unsupported consolidation function")),
    }
}
impl Bridge {
    pub(crate) fn owned_rectangle(&self, session: &Session, owner: usize, value: &Value) -> HostResult<(usize, RangeRef)> {
        let (sheet, range) = if let Value::Object(reference) = value {
            let Object::Range(sheet, range) = self.object(session, *reference)? else {
                return Err(failure(13, "Range required"));
            };
            (sheet, range)
        } else {
            Self::resolve_text(session, owner, &string(value)?, None, 0, &mut 0)?
        };
        let rect = range.rectangle()?;
        values::dimensions(rect)?;
        Ok((sheet, rect))
    }
    pub(crate) fn qualified_rectangle(session: &Session, sheet: usize, rect: RangeRef) -> HostResult<String> {
        let name =
            session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?.name.as_str();
        Ok(format!("{}!{}", gridcraft_engine::core::quote_sheet(name), rect.a1()))
    }
    pub(crate) fn whatif_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        let Object::Range(sheet, range) = object else {
            return None;
        };
        if !reads(call) {
            return None;
        }
        let member = call.member.to_ascii_lowercase();
        if !matches!(member.as_str(), "goalseek" | "advancedfilter" | "consolidate" | "table") {
            return None;
        }
        Some((|| {
            let rect = range.rectangle()?;
            values::dimensions(rect)?;
            let mut params = json!({"range":Self::qualified_rectangle(session,*sheet,rect)?});
            if member == "goalseek" {
                let args = arguments(call, &["Goal", "ChangingCell"])?;
                let (owner, changing) = self.owned_rectangle(session, *sheet, required(&args, 1)?)?;
                if !rect.is_single() || !changing.is_single() {
                    return Err(failure(1004, "GoalSeek requires single set/changing cells"));
                }
                params["set"] = json!(Self::qualified_rectangle(session, *sheet, rect)?);
                params["changing"] = json!(Self::qualified_rectangle(session, owner, changing)?);
                params["to"] = json!(super::format_api::number(required(&args, 0)?)?);
                let changes = self.prepare_change(session, owner, changing.into())?;
                let result = session.execute("data.goalSeek", params).map_err(|e| failure(1004, e.to_string()))?;
                self.queue_changes(changes);
                return Ok(Value::Boolean(result.get("converged").and_then(serde_json::Value::as_bool).unwrap_or(false)));
            }
            let command = if member == "advancedfilter" {
                let args = arguments(call, &["Action", "CriteriaRange", "CopyToRange", "Unique"])?;
                let action = numeric_index(required(&args, 0)?)?;
                if !matches!(action, 1 | 2) {
                    return Err(failure(1004, "invalid AdvancedFilter action"));
                }
                if !absent(&args[1]) {
                    let (owner, criteria) = self.owned_rectangle(session, *sheet, &args[1])?;
                    params["criteria"] = json!(Self::qualified_rectangle(session, owner, criteria)?);
                }
                if action == 2 {
                    let (owner, destination) = self.owned_rectangle(session, *sheet, required(&args, 2)?)?;
                    if !destination.is_single() {
                        return Err(failure(1004, "native AdvancedFilter copy requires a destination anchor cell"));
                    }
                    params["copyTo"] = json!(Self::qualified_rectangle(session, owner, destination)?);
                } else if !absent(&args[2]) {
                    return Err(failure(1004, "CopyToRange requires xlFilterCopy"));
                }
                params["unique"] = json!(!absent(&args[3]) && super::address::boolean(&args[3])?);
                "data.advancedFilter"
            } else if member == "consolidate" {
                let args = arguments(call, &["Sources", "Function", "TopRow", "LeftColumn", "CreateLinks"])?;
                if !absent(&args[4]) && super::address::boolean(&args[4])? {
                    return Err(failure(1004, "linked consolidation is unavailable"));
                }
                let source = required(&args, 0)?;
                let items = if let Value::Array(a) = source { a.elements().to_vec() } else { vec![source.clone()] };
                if items.is_empty() || items.len() > 256 {
                    return Err(failure(7, "consolidation source budget"));
                }
                let mut sources = Vec::new();
                let mut cells = 0u64;
                for source in items {
                    let (owner, area) = self.owned_rectangle(session, *sheet, &source)?;
                    cells = cells.saturating_add(area.count());
                    if cells > super::MAX_CELLS as u64 {
                        return Err(failure(7, "consolidation cell budget"));
                    }
                    sources.push(Self::qualified_rectangle(session, owner, area)?);
                }
                params["sources"] = json!(sources);
                params["destination"] = json!(Self::qualified_rectangle(session, *sheet, RangeRef::cell(rect.start))?);
                params["function"] = json!(aggregate(if absent(&args[1]) { -4157 } else { numeric_index(&args[1])? })?);
                params["topRow"] = json!(!absent(&args[2]) && super::address::boolean(&args[2])?);
                params["leftColumn"] = json!(!absent(&args[3]) && super::address::boolean(&args[3])?);
                "data.consolidate"
            } else {
                let args = arguments(call, &["RowInput", "ColumnInput"])?;
                if args.iter().all(absent) {
                    return Err(failure(449, "a data table needs an input cell"));
                }
                for (i, key) in [(0, "rowInput"), (1, "columnInput")] {
                    if !absent(&args[i]) {
                        let (owner, input) = self.owned_rectangle(session, *sheet, &args[i])?;
                        if !input.is_single() {
                            return Err(failure(1004, "data-table input must be one cell"));
                        }
                        params[key] = json!(Self::qualified_rectangle(session, owner, input)?);
                    }
                }
                "data.dataTable"
            };
            let changes = if command == "data.dataTable" {
                if rect.height() < 2 || rect.width() < 2 {
                    return Err(failure(1004, "data table needs at least two rows and columns"));
                }
                let body = RangeRef::new(gridcraft_engine::core::CellRef::new(rect.start.row + 1, rect.start.col + 1), rect.end);
                self.prepare_change(session, *sheet, body.into())?
            } else {
                match session.preview_data_mutation(command, &params).map_err(|e| failure(1004, e.to_string()))? {
                    Some((sheet, rect)) => {
                        values::dimensions(rect)?;
                        self.prepare_change(session, sheet, rect.into())?
                    }
                    None => super::PreparedChanges::default(),
                }
            };
            session.execute(command, params).map_err(|e| failure(1004, e.to_string()))?;
            self.queue_changes(changes);
            Ok(Value::Empty)
        })())
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
 ("GoalSeek","Function GoalSeek(Goal As Variant, ChangingCell As Range) As Boolean"),
 ("AdvancedFilter","Sub AdvancedFilter(Action As Long, Optional CriteriaRange As Variant, Optional CopyToRange As Variant, Optional Unique As Variant)"),
 ("Consolidate","Sub Consolidate(Sources As Variant, Optional Function As Variant, Optional TopRow As Variant, Optional LeftColumn As Variant, Optional CreateLinks As Variant)"),
 ("Table","Sub Table(Optional RowInput As Variant, Optional ColumnInput As Variant)"),
].into_iter().map(|(name,signature)|super::MemberDescriptor{owner:"Range",name,signature,availability:super::MemberAvailability::Implemented,implementation_symbol:"crates/vba/src/whatif_api.rs::Bridge::whatif_api"}).collect()
}
