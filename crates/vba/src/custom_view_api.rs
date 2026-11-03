//! Existing native saved views; collection edits invalidate prior stash snapshots.
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::model::sheet::SheetIdentity;
use serde_json::{Value as Json, json};
fn views(s: &Session) -> HostResult<(Json, Option<SheetIdentity>)> {
    let wb = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb;
    let Some(n) = wb.names.iter().find(|n| n.name == "_sc_custom_views") else {
        return Ok((json!({}), None));
    };
    if n.formula.len() > 1_048_576 {
        return Err(failure(7, "saved view metadata budget"));
    }
    let text =
        n.formula.strip_prefix('"').and_then(|t| t.strip_suffix('"')).ok_or_else(|| failure(1004, "invalid saved views"))?.replace("\"\"", "\"");
    let v: Json = serde_json::from_str(&text).map_err(|_| failure(1004, "invalid saved views"))?;
    let map = v.as_object().ok_or_else(|| failure(1004, "invalid saved views"))?;
    if map.len() > super::MAX_OBJECTS || map.keys().any(|k| k.len() > 1024) {
        return Err(failure(7, "saved view collection budget"));
    }
    Ok((v, Some(n.identity.clone())))
}
fn node(name: &str, identity: Option<SheetIdentity>) -> Feature {
    let mut f = Feature::new("CustomView", None);
    f.text = Some(name.into());
    f.identity = identity;
    f
}
impl Bridge {
    pub(crate) fn custom_view_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        if *o == Object::Workbook && c.member.eq_ignore_ascii_case("CustomViews") && reads(c) {
            return Some(if c.arguments.is_empty() {
                self.feature(Feature::new("CustomViews", None))
            } else {
                self.custom_view_dispatch(s, &Feature::new("CustomViews", None), c)
            });
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if !matches!(f.kind, "CustomViews" | "CustomView") {
            return None;
        }
        Some(self.custom_view_dispatch(s, f, c))
    }
    fn custom_view_dispatch(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let (all, identity) = views(s)?;
        let m = c.member.to_ascii_lowercase();
        if f.kind == "CustomViews" {
            let map = all.as_object().ok_or_else(|| failure(1004, "invalid saved views"))?;
            if c.kind == InvokeKind::Enumerate {
                arguments(c, &[])?;
                return self.object_array(map.keys().map(|name| node(name, identity.clone()).object()).collect());
            }
            if m == "add" && reads(c) {
                let a = arguments(c, &["ViewName", "PrintSettings", "RowColSettings"])?;
                let name = string(required(&a, 0)?)?;
                if name.is_empty() || name.len() > 1024 || map.len() >= super::MAX_OBJECTS {
                    return Err(failure(1004, "saved view name or collection budget"));
                }
                // Native views do not capture either Excel option; omitted defaults are unsupported.
                for i in [1, 2] {
                    if a.get(i).is_none_or(|v| matches!(v, Value::Missing)) || super::address::boolean(required(&a, i)?)? {
                        return Err(failure(1004, "native views require PrintSettings=False and RowColSettings=False"));
                    }
                }
                if map.keys().any(|k| k.eq_ignore_ascii_case(&name)) {
                    return Err(failure(1004, "saved view already exists"));
                }
                let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
                let sh = d.wb.active().ok_or_else(|| failure(91, "worksheet missing"))?;
                if d.selection.ranges.len() > 64 {
                    return Err(failure(7, "saved selection budget"));
                }
                let mut prospective = all.clone();
                prospective[&name] = json!({"sheet":d.sheet_index(),"selection":d.selection.a1(),"zoom":sh.zoom,"freeze":sh.freeze});
                let text = prospective.to_string();
                if text.len().saturating_add(text.bytes().filter(|b| *b == b'"').count()).saturating_add(2) > 1_048_576 {
                    return Err(failure(7, "saved view metadata budget"));
                }
                self.objects.reserve(1)?;
                s.execute("view.customViews", json!({"save":name})).map_err(|e| failure(1004, e.to_string()))?;
                let (_, id) = views(s)?;
                return self.feature(node(&name, id));
            }
            if !reads(c) {
                return Err(failure(438, "read-only custom view collection"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(c, &[])?;
                    Ok(Value::Long(map.len() as i32))
                }
                "parent" => {
                    arguments(c, &[])?;
                    self.intern(Object::Workbook)
                }
                "application" => {
                    arguments(c, &[])?;
                    self.intern(Object::Application)
                }
                "item" | "" | "customviews" => {
                    let a = arguments(c, &["Index"])?;
                    let v = required(&a, 0)?;
                    let name = if matches!(v, Value::String(_)) {
                        let name = string(v)?;
                        map.keys().find(|k| k.eq_ignore_ascii_case(&name))
                    } else {
                        usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "saved view index"))?)
                            .ok()
                            .and_then(|i| map.keys().nth(i))
                    }
                    .ok_or_else(|| failure(9, "saved view not found"))?;
                    self.feature(node(name, identity))
                }
                _ => Err(failure(438, "unsupported saved view member")),
            };
        }
        if !f.identity.as_ref().zip(identity.as_ref()).is_some_and(|(a, b)| a.same(b)) {
            return Err(failure(91, "saved view collection was edited"));
        }
        let name = f.text.as_deref().ok_or_else(|| failure(91, "saved view missing"))?;
        let v = all.get(name).ok_or_else(|| failure(91, "saved view deleted"))?;
        arguments(c, &[])?;
        if c.kind == InvokeKind::Call && matches!(m.as_str(), "show" | "delete") {
            if m == "show" {
                // Imported malformed metadata must fail before any native command changes focus.
                let si = v["sheet"].as_u64().and_then(|n| usize::try_from(n).ok()).ok_or_else(|| failure(1004, "invalid view sheet"))?;
                if s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).is_none() {
                    return Err(failure(1004, "view sheet no longer exists"));
                }
                let selection = v["selection"].as_str().ok_or_else(|| failure(1004, "invalid view selection"))?;
                if selection.len() > 4096
                    || selection.split(',').count() > 64
                    || selection.split(',').any(|a| gridcraft_engine::core::RangeRef::parse(a).is_none())
                {
                    return Err(failure(1004, "invalid view selection"));
                }
                if !v["zoom"].as_u64().is_some_and(|z| (10..=400).contains(&z)) {
                    return Err(failure(1004, "invalid view zoom"));
                }
                if let Some(freeze) = v.get("freeze")
                    && !freeze.is_null()
                    && (!freeze.as_array().is_some_and(|a| a.len() == 2)
                        || !freeze[0].as_u64().is_some_and(|r| r < u64::from(gridcraft_engine::core::MAX_ROWS))
                        || !freeze[1].as_u64().is_some_and(|c| c < u64::from(gridcraft_engine::core::MAX_COLS)))
                {
                    return Err(failure(1004, "invalid view freeze bounds"));
                }
                let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
                let sh = d.wb.sheet(si).ok_or_else(|| failure(91, "worksheet missing"))?;
                let active = selection
                    .split(',')
                    .next_back()
                    .and_then(gridcraft_engine::core::RangeRef::parse)
                    .map(|r| r.start)
                    .ok_or_else(|| failure(1004, "invalid view selection"))?;
                if let Some(p) = &sh.protection {
                    let locked = d.wb.styles.get(sh.style_id(active)).protection.locked;
                    if (locked && !p.select_locked) || (!locked && !p.select_unlocked) {
                        return Err(failure(1004, "view selection is protected"));
                    }
                }
            }
            s.execute("view.customViews", if m == "show" { json!({"show":name}) } else { json!({"delete":name}) })
                .map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if !reads(c) {
            return Err(failure(438, "read-only saved view"));
        }
        match m.as_str() {
            "name" => values::vba_text(name),
            "parent" => self.intern(Object::Workbook),
            "application" => self.intern(Object::Application),
            "printsettings" | "rowcolsettings" => Ok(Value::Boolean(false)),
            _ => Err(failure(438, "unsupported saved view member")),
        }
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Workbook", "CustomViews", "Property Get CustomViews() As CustomViews"),
        ("CustomViews", "Count", "Property Get Count() As Long"),
        ("CustomViews", "Item", "Property Get Item(Index As Variant) As CustomView"),
        (
            "CustomViews",
            "Add",
            "Function Add(ViewName As String, Optional PrintSettings As Variant, Optional RowColSettings As Variant) As CustomView",
        ),
        ("CustomViews", "Parent", "Property Get Parent() As Workbook"),
        ("CustomViews", "Application", "Property Get Application() As Application"),
        ("CustomView", "Parent", "Property Get Parent() As Workbook"),
        ("CustomView", "Application", "Property Get Application() As Application"),
        ("CustomView", "Name", "Property Get Name() As String"),
        ("CustomView", "PrintSettings", "Property Get PrintSettings() As Boolean"),
        ("CustomView", "RowColSettings", "Property Get RowColSettings() As Boolean"),
        ("CustomView", "Show", "Sub Show()"),
        ("CustomView", "Delete", "Sub Delete()"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/custom_view_api.rs::Bridge::custom_view_api",
    })
    .collect()
}
