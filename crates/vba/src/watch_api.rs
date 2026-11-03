//! Session-wide view over native per-document watch lists.
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure};
use serde_json::json;
fn nodes(s: &Session) -> HostResult<Vec<Object>> {
    let mut all = Vec::new();
    for d in s.documents() {
        let Some(n) = d.wb.names.iter().find(|n| n.name == "_sc_watch") else {
            continue;
        };
        if n.formula.len() > 1_048_576 {
            return Err(failure(7, "watch metadata budget"));
        }
        let text = n
            .formula
            .strip_prefix('"')
            .and_then(|t| t.strip_suffix('"'))
            .ok_or_else(|| failure(1004, "invalid watch metadata"))?
            .replace("\"\"", "\"");
        let list: Vec<String> = serde_json::from_str(&text).map_err(|_| failure(1004, "invalid watch list"))?;
        if list.len() > super::MAX_OBJECTS || list.len().saturating_add(all.len()) > super::MAX_OBJECTS || list.iter().any(|t| t.len() > 4096) {
            return Err(failure(7, "watch collection budget"));
        }
        for text in list {
            let mut f = Feature::new("Watch", None);
            f.text = Some(text);
            f.identity = Some(n.identity.clone());
            all.push(Object::Document(d.uid, Box::new(f.object())));
        }
    }
    Ok(all)
}
impl Bridge {
    fn watch_permission(&mut self, s: &mut Session, doc: u64) -> HostResult<()> {
        let d = s.document(doc).ok_or_else(|| failure(91, "workbook closed"))?;
        if !d.trust.allows_code() && !(doc == self.doc && self.consent) {
            return Err(failure(70, "target workbook requires consent"));
        }
        self.ensure_group(s, doc).map_err(|e| failure(7, e.to_string()))
    }
    pub(crate) fn watch_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        if *o == Object::Application && c.member.eq_ignore_ascii_case("Watches") && reads(c) {
            return Some(if c.arguments.is_empty() {
                self.feature(Feature::new("Watches", None))
            } else {
                self.watch_dispatch(s, &Feature::new("Watches", None), c)
            });
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if !matches!(f.kind, "Watch" | "Watches") {
            return None;
        }
        Some(self.watch_dispatch(s, f, c))
    }
    fn watch_dispatch(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let all = nodes(s)?;
        let m = c.member.to_ascii_lowercase();
        if f.kind == "Watches" {
            if c.kind == InvokeKind::Enumerate {
                arguments(c, &[])?;
                return self.object_array(all);
            }
            if m == "add" && reads(c) {
                let a = arguments(c, &["Source"])?;
                let Value::Object(reference) = required(&a, 0)? else {
                    return Err(failure(13, "Watch.Source requires a single-cell Range"));
                };
                let Object::Document(doc, inner) = self.objects.get(*reference)? else {
                    return Err(failure(13, "Watch.Source requires a Range"));
                };
                let Object::Range(si, r) = *inner else {
                    return Err(failure(13, "Watch.Source requires a Range"));
                };
                let rect = r.rectangle()?;
                if rect.count() != 1 || all.len() >= super::MAX_OBJECTS {
                    return Err(failure(1004, "native watches require one cell; collection is bounded"));
                }
                let sh = s.document(doc).and_then(|d| d.wb.sheet(si)).ok_or_else(|| failure(91, "source worksheet deleted"))?;
                let text = format!("'{}'!{}", sh.name.replace('\'', "''"), rect.start.a1());
                if let Some(existing) = all.iter().find(|o| matches!(o,Object::Document(uid,inner) if *uid==doc && matches!(inner.as_ref(),Object::Feature(f) if f.text.as_deref()==Some(&text)))) {
                    self.watch_permission(s, doc)?;
                    return self.intern(existing.clone());
                }
                let mut prospective = all
                    .iter()
                    .filter_map(|o| match o {
                        Object::Document(uid, inner) if *uid == doc => match inner.as_ref() {
                            Object::Feature(f) => f.text.clone(),
                            _ => None,
                        },
                        _ => None,
                    })
                    .collect::<Vec<_>>();
                prospective.push(text.clone());
                let encoded = json!(prospective).to_string();
                if encoded.len().saturating_add(encoded.bytes().filter(|b| *b == b'"').count()).saturating_add(2) > 1_048_576 {
                    return Err(failure(7, "watch metadata budget"));
                }
                self.objects.reserve(1)?;
                self.watch_permission(s, doc)?;
                s.execute_in(gridcraft_engine::Target::Doc(doc), "formulas.watchWindow", json!({"add":text}))
                    .map_err(|e| failure(1004, e.to_string()))?;
                let found=nodes(s)?.into_iter().find(|o|matches!(o,Object::Document(uid,inner) if *uid==doc && matches!(inner.as_ref(),Object::Feature(f) if f.text.as_deref()==Some(&text)))).ok_or_else(||failure(91,"watch missing"))?;
                return self.intern(found);
            }
            if m == "delete" && c.kind == InvokeKind::Call {
                arguments(c, &[])?;
                // Validate every destination before removing anything.
                for o in &all {
                    if let Object::Document(doc, _) = o {
                        self.watch_permission(s, *doc)?;
                    }
                }
                for o in all {
                    if let Object::Document(doc, inner) = o
                        && let Object::Feature(f) = *inner
                    {
                        s.execute_in(gridcraft_engine::Target::Doc(doc), "formulas.watchWindow", json!({"remove":f.text}))
                            .map_err(|e| failure(1004, e.to_string()))?;
                    }
                }
                return Ok(Value::Empty);
            }
            if !reads(c) {
                return Err(failure(438, "read-only watches collection"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(c, &[])?;
                    Ok(Value::Long(all.len() as i32))
                }
                "application" | "parent" => {
                    arguments(c, &[])?;
                    self.intern(Object::Application)
                }
                "item" | "" | "watches" => {
                    let a = arguments(c, &["Index"])?;
                    let i = usize::try_from(numeric_index(required(&a, 0)?)?.checked_sub(1).ok_or_else(|| failure(9, "watch index"))?)
                        .map_err(|_| failure(9, "watch index"))?;
                    self.intern(all.get(i).cloned().ok_or_else(|| failure(9, "watch index"))?)
                }
                _ => Err(failure(438, "unsupported watches member")),
            };
        }
        let doc = s.doc().map_err(|e| failure(91, e.to_string()))?.uid;
        if !all.iter().any(|o| matches!(o,Object::Document(uid,inner) if *uid==doc && inner.as_ref()==&f.clone().object())) {
            return Err(failure(91, "watch list edited or watch deleted"));
        }
        arguments(c, &[])?;
        if m == "delete" && c.kind == InvokeKind::Call {
            s.execute("formulas.watchWindow", json!({"remove":f.text})).map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if !reads(c) {
            return Err(failure(438, "read-only watch"));
        }
        match m.as_str() {
            "application" | "parent" => self.intern(Object::Application),
            "source" => {
                let text = f.text.as_deref().ok_or_else(|| failure(91, "watch missing"))?;
                let (sheet, body) = gridcraft_engine::core::split_sheet(text);
                if sheet.is_none() || gridcraft_engine::core::CellRef::parse(body).is_none() {
                    return Err(failure(1004, "native watch lacks a stable single-cell source"));
                }
                // Resolve qualified single-cell text using native workbook rules.
                let si = s.doc().map_err(|e| failure(91, e.to_string()))?.sheet_index();
                let (si, r) = Self::resolve_text(s, si, text, None, 0, &mut 0)?;
                self.intern(Object::Range(si, r))
            }
            _ => Err(failure(438, "unsupported watch member")),
        }
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Application", "Watches", "Property Get Watches() As Watches"),
        ("Watches", "Application", "Property Get Application() As Application"),
        ("Watches", "Parent", "Property Get Parent() As Application"),
        ("Watches", "Count", "Property Get Count() As Long"),
        ("Watches", "Item", "Property Get Item(Index As Variant) As Watch"),
        ("Watches", "Add", "Function Add(Source As Variant) As Watch"),
        ("Watches", "Delete", "Sub Delete()"),
        ("Watch", "Application", "Property Get Application() As Application"),
        ("Watch", "Parent", "Property Get Parent() As Application"),
        ("Watch", "Source", "Property Get Source() As Variant"),
        ("Watch", "Delete", "Sub Delete()"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/watch_api.rs::Bridge::watch_api",
    })
    .collect()
}
