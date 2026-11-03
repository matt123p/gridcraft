//! Native workbook cell-style catalog and application. No parallel style store.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::model::Style;
use serde_json::json;
fn catalog(session: &Session) -> HostResult<Vec<(String, Style, bool)>> {
    let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
    if doc.wb.cell_styles.len() > super::MAX_OBJECTS {
        return Err(failure(7, "style collection budget"));
    }
    let mut styles = gridcraft_engine::cmd::format::CELL_STYLE_NAMES
        .iter()
        .filter_map(|name| gridcraft_engine::cmd::format::builtin_cell_style(name, &doc.wb.theme).map(|style| ((*name).to_string(), style, true)))
        .collect::<Vec<_>>();
    for (name, style) in &doc.wb.cell_styles {
        if let Some(item) = styles.iter_mut().find(|item| item.0.eq_ignore_ascii_case(name)) {
            *item = (name.clone(), style.clone(), false);
        } else {
            styles.push((name.clone(), style.clone(), false));
        }
    }
    Ok(styles)
}
fn node(name: String) -> Feature {
    let mut f = Feature::new("Style", None);
    f.text = Some(name);
    f
}
impl Bridge {
    pub(crate) fn named_style_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        let member = call.member.to_ascii_lowercase();
        if matches!(object, Object::Workbook) && member == "styles" && reads(call) {
            return Some(if call.arguments.is_empty() {
                self.feature(Feature::new("Styles", None))
            } else {
                let mut call = call.clone();
                call.member = "Item".into();
                self.named_style_dispatch(session, &Feature::new("Styles", None), &call)
            });
        }
        if let Object::Range(sheet, range) = object
            && member == "style"
        {
            return Some((|| {
                if reads(call) {
                    arguments(call, &[])?;
                    let all = catalog(session)?;
                    let styles = self.styles(session, *sheet, range)?;
                    let name = |style: &&Style| all.iter().find(|(_, s, _)| s == *style).map(|s| s.0.as_str());
                    let first = styles.first().ok_or_else(|| failure(1004, "empty range"))?;
                    if styles.iter().any(|s| name(s) != name(first)) {
                        return Ok(Value::Null);
                    }
                    return values::vba_text(name(first).ok_or_else(|| failure(1004, "cell has no exact native named style"))?);
                }
                if !matches!(call.kind, InvokeKind::Let | InvokeKind::Set) {
                    return Err(failure(438, "unsupported style operation"));
                }
                let args = arguments(call, &["Value"])?;
                let v = required(&args, 0)?;
                let name = if let Value::Object(reference) = v {
                    let Object::Feature(f) = self.object(session, *reference)? else {
                        return Err(failure(13, "Style required"));
                    };
                    if f.kind != "Style" {
                        return Err(failure(13, "Style required"));
                    }
                    f.text.ok_or_else(|| failure(91, "style missing"))?
                } else {
                    string(v)?
                };
                values::dimensions(range.rectangle()?)?;
                session
                    .execute("home.cellStyle", json!({"sheet":sheet,"range":range.rectangle()?.a1(),"name":name}))
                    .map_err(|e| failure(1004, e.to_string()))?;
                Ok(Value::Empty)
            })());
        }
        if let Object::Feature(f) = object
            && matches!(f.kind, "Style" | "Styles")
        {
            return Some(self.named_style_dispatch(session, f, call));
        }
        None
    }
    fn named_style_dispatch(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let member = call.member.to_ascii_lowercase();
        if reads(call) && matches!(member.as_str(), "parent" | "application") {
            arguments(call, &[])?;
            return self.intern(if member == "parent" { Object::Workbook } else { Object::Application });
        }
        let styles = catalog(session)?;
        if f.kind == "Styles" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(styles.into_iter().map(|(name, _, _)| node(name).object()).collect());
            }
            if !reads(call) {
                return Err(failure(438, "unsupported Styles assignment"));
            }
            match member.as_str() {
                "count" => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(styles.len() as i32));
                }
                "item" | "" => {
                    let args = arguments(call, &["Index"])?;
                    let value = required(&args, 0)?;
                    let style = if matches!(value, Value::String(_)) {
                        let name = string(value)?;
                        styles.iter().find(|s| s.0.eq_ignore_ascii_case(&name))
                    } else {
                        usize::try_from(numeric_index(value)?.checked_sub(1).ok_or_else(|| failure(9, "style index"))?)
                            .ok()
                            .and_then(|i| styles.get(i))
                    }
                    .ok_or_else(|| failure(9, "style not found"))?;
                    return self.feature(node(style.0.clone()));
                }
                "add" => {
                    let args = arguments(call, &["Name", "BasedOn"])?;
                    let name = string(required(&args, 0)?)?;
                    if name.is_empty() || name.len() > 255 {
                        return Err(failure(1004, "invalid style name"));
                    }
                    let (sheet, cell) = if absent(&args[1]) {
                        let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
                        (doc.sheet_index(), doc.selection.active)
                    } else {
                        let Value::Object(reference) = &args[1] else {
                            return Err(failure(13, "BasedOn requires Range"));
                        };
                        let Object::Range(sheet, range) = self.object(session, *reference)? else {
                            return Err(failure(13, "BasedOn requires Range"));
                        };
                        (sheet, range.first()?.start)
                    };
                    self.objects.reserve(1)?;
                    session
                        .execute("home.newCellStyle", json!({"name":name,"sheet":sheet,"fromCell":cell.a1()}))
                        .map_err(|e| failure(1004, e.to_string()))?;
                    return self.feature(node(name));
                }
                _ => return Err(failure(438, "unsupported Styles member")),
            }
        }
        if !reads(call) {
            return Err(failure(1004, "native named-style editing is unavailable"));
        }
        let name = f.text.as_deref().ok_or_else(|| failure(91, "style deleted"))?;
        let (name, style, builtin) = styles.iter().find(|s| s.0.eq_ignore_ascii_case(name)).ok_or_else(|| failure(91, "style deleted"))?;
        arguments(call, &[])?;
        match member.as_str() {
            "name" | "namelocal" | "value" | "" => values::vba_text(name),
            "builtin" => Ok(Value::Boolean(*builtin)),
            "numberformat" => values::vba_text(style.num_fmt.as_str()),
            "locked" => Ok(Value::Boolean(style.protection.locked)),
            "formulahidden" => Ok(Value::Boolean(style.protection.hidden)),
            _ => Err(failure(438, "unsupported Style member")),
        }
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Workbook", "Styles", "Property Get Styles() As Styles"),
        ("Range", "Style", "Property Style As Variant"),
        ("Styles", "Add", "Function Add(Name As String, Optional BasedOn As Variant) As Style"),
        ("Styles", "Count", "Property Get Count() As Long"),
        ("Styles", "Item", "Property Get Item(Index As Variant) As Style"),
        ("Styles", "Parent", "Property Get Parent() As Workbook"),
        ("Styles", "Application", "Property Get Application() As Application"),
        ("Style", "Name", "Property Get Name() As String"),
        ("Style", "NameLocal", "Property Get NameLocal() As String"),
        ("Style", "Value", "Property Get Value() As String"),
        ("Style", "BuiltIn", "Property Get BuiltIn() As Boolean"),
        ("Style", "NumberFormat", "Property Get NumberFormat() As String"),
        ("Style", "Locked", "Property Get Locked() As Boolean"),
        ("Style", "FormulaHidden", "Property Get FormulaHidden() As Boolean"),
        ("Style", "Parent", "Property Get Parent() As Workbook"),
        ("Style", "Application", "Property Get Application() As Application"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/named_style_api.rs::Bridge::named_style_api",
    })
    .collect()
}
