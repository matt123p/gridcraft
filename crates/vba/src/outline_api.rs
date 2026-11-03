//! Existing worksheet outline state and native visibility commands.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure};
use serde_json::json;
impl Bridge {
    pub(crate) fn outline_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if let Object::Sheet(sheet) = object
            && call.member.eq_ignore_ascii_case("Outline")
            && reads(call)
        {
            return Some(arguments(call, &[]).and_then(|_| self.feature(Feature::new("Outline", Some(*sheet)))));
        }
        if let Object::Range(sheet, range) = object
            && call.member.eq_ignore_ascii_case("OutlineLevel")
            && reads(call)
        {
            return Some((|| {
                arguments(call, &[])?;
                let r = range.rectangle()?;
                let rows = r.is_full_rows();
                if !rows && !r.is_full_cols() {
                    return Err(failure(1004, "OutlineLevel requires whole rows or columns"));
                }
                let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(*sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
                let (a, b, map) = if rows { (r.start.row, r.end.row, &sh.rows) } else { (r.start.col, r.end.col, &sh.cols) };
                let first = map.get(&a).map_or(0, |v| v.outline);
                if (a..=b).any(|i| map.get(&i).map_or(0, |v| v.outline) != first) {
                    return Ok(Value::Null);
                }
                Ok(Value::Long(i32::from(first) + 1))
            })());
        }
        let Object::Feature(f) = object else {
            return None;
        };
        if f.kind != "Outline" {
            return None;
        }
        Some((|| {
            let sheet = f.sheet.ok_or_else(|| failure(91, "outline worksheet deleted"))?;
            let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
            let member = call.member.to_ascii_lowercase();
            if reads(call) {
                match member.as_str() {
                    "parent" => {
                        arguments(call, &[])?;
                        return self.intern(Object::Sheet(sheet));
                    }
                    "application" => {
                        arguments(call, &[])?;
                        return self.intern(Object::Application);
                    }
                    "summaryrow" | "summarycolumn" => {
                        arguments(call, &[])?;
                        return Ok(Value::Long(if member == "summaryrow" {
                            if sh.summary_below { 1 } else { 0 }
                        } else if sh.summary_right {
                            -4152
                        } else {
                            -4131
                        }));
                    }
                    _ => {}
                }
            }
            if sh.is_protected() {
                return Err(failure(1004, "outline changes require unprotected worksheet"));
            }
            if member == "showlevels" && reads(call) {
                let args = arguments(call, &["RowLevels", "ColumnLevels"])?;
                let mut params = json!({"sheet":sheet});
                let mut any = false;
                for (i, key) in [(0, "rows"), (1, "cols")] {
                    if !absent(&args[i]) {
                        let n = numeric_index(&args[i])?;
                        if !(0..=8).contains(&n) {
                            return Err(failure(1004, "outline level outside 0..8"));
                        }
                        if n > 0 {
                            params[key] = json!(n);
                            any = true;
                        }
                    }
                }
                if any {
                    session.execute("data.showLevels", params).map_err(|e| failure(1004, e.to_string()))?;
                }
                return Ok(Value::Empty);
            }
            if call.kind == InvokeKind::Let && matches!(member.as_str(), "summaryrow" | "summarycolumn") {
                let args = arguments(call, &["Value"])?;
                let n = numeric_index(required(&args, 0)?)?;
                let below = if member == "summaryrow" {
                    match n {
                        1 => true,
                        0 => false,
                        _ => return Err(failure(1004, "invalid summary row direction")),
                    }
                } else {
                    match n {
                        -4152 => true,
                        -4131 => false,
                        _ => return Err(failure(1004, "invalid summary column direction")),
                    }
                };
                let key = if member == "summaryrow" { "summaryBelow" } else { "summaryRight" };
                session.execute("data.outlineSettings", json!({"sheet":sheet,key:below})).map_err(|e| failure(1004, e.to_string()))?;
                return Ok(Value::Empty);
            }
            Err(failure(438, "unsupported Outline member"))
        })())
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Worksheet", "Outline", "Property Get Outline() As Outline"),
        ("Range", "OutlineLevel", "Property Get OutlineLevel() As Variant"),
        ("Outline", "ShowLevels", "Sub ShowLevels(Optional RowLevels As Variant, Optional ColumnLevels As Variant)"),
        ("Outline", "SummaryRow", "Property SummaryRow As Long"),
        ("Outline", "SummaryColumn", "Property SummaryColumn As Long"),
        ("Outline", "Parent", "Property Get Parent() As Worksheet"),
        ("Outline", "Application", "Property Get Application() As Application"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/outline_api.rs::Bridge::outline_api",
    })
    .collect()
}
