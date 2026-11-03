//! Native protection permissions and Allow Edit Range lifetimes; no ACL/IRM emulation.
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, ranges, values};
use gridcraft_engine::model::EditRange;
use serde_json::json;
fn node(si: usize, r: &EditRange) -> Feature {
    let mut f = Feature::new("AllowEditRange", Some(si));
    f.identity = Some(r.identity.clone());
    f
}
fn password(v: &Value) -> HostResult<String> {
    let text = string(v)?;
    if text.len() > 1024 {
        return Err(failure(7, "password payload budget"));
    }
    Ok(text)
}
fn title(v: &Value) -> HostResult<String> {
    let text = string(v)?;
    if text.trim() != text || text.is_empty() || text.len() > 255 {
        return Err(failure(1004, "invalid edit-range title"));
    }
    Ok(text)
}
fn lookup(s: &Session, f: &Feature) -> HostResult<(usize, EditRange)> {
    let si = f.sheet.ok_or_else(|| failure(91, "worksheet deleted"))?;
    let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
    if sh.edit_ranges.len() > super::MAX_OBJECTS {
        return Err(failure(7, "edit-range collection budget"));
    }
    let r = sh
        .edit_ranges
        .iter()
        .find(|r| f.identity.as_ref().is_some_and(|id| id.same(&r.identity)))
        .ok_or_else(|| failure(91, "edit range deleted or replaced"))?;
    if r.title.len() > 1024 || r.ranges.len() > 64 {
        return Err(failure(7, "edit-range metadata budget"));
    }
    let range = ranges::Range { areas: r.ranges.clone(), unit: ranges::Unit::Cells };
    range.validate()?;
    Ok((si, r.clone()))
}
impl Bridge {
    pub(crate) fn permission_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        let m = c.member.to_ascii_lowercase();
        if let Object::Sheet(si) = o {
            if m == "protection" && reads(c) {
                return Some((|| {
                    arguments(c, &[])?;
                    self.feature(Feature::new("Protection", Some(*si)))
                })());
            }
            if matches!(m.as_str(), "enableselection" | "protectionmode" | "protectdrawingobjects" | "protectscenarios") {
                return Some(self.sheet_permissions(s, *si, c));
            }
        }
        if let Object::Range(si, r) = o
            && m == "allowedit"
        {
            return Some((|| {
                if !reads(c) {
                    return Err(failure(438, "AllowEdit is read-only"));
                }
                arguments(c, &[])?;
                if r.count()? > super::MAX_CELLS as u64 {
                    return Err(failure(7, "AllowEdit cell budget"));
                }
                let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
                let sh = d.wb.sheet(*si).ok_or_else(|| failure(91, "worksheet deleted"))?;
                Ok(Value::Boolean(
                    !sh.is_protected()
                        || r.areas.iter().all(|r| {
                            r.iter().all(|cell| {
                                !d.wb.styles.get(sh.style_id(cell)).protection.locked
                                    || gridcraft_engine::cmd::locked_cell_error(sh, cell, &d.unlocked_ranges).is_none()
                            })
                        }),
                ))
            })());
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if !matches!(f.kind, "Protection" | "AllowEditRanges" | "AllowEditRange") {
            return None;
        }
        Some(self.permission_dispatch(s, f, c))
    }
    fn sheet_permissions(&mut self, s: &mut Session, si: usize, c: &HostCall) -> HostResult<Value> {
        let m = c.member.to_ascii_lowercase();
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if m == "enableselection" && c.kind == InvokeKind::Let {
            let a = arguments(c, &["Value"])?;
            let v = numeric_index(required(&a, 0)?)?;
            if !matches!(v, 0 | 1 | -4142) {
                return Err(failure(1004, "invalid selection mode"));
            }
            if sh.protection.is_none() {
                return Err(failure(1004, "native selection settings require a protected worksheet"));
            }
            s.execute("review.protectSheet", json!({"sheet":si,"selectionOnly":true,"selectLockedCells":v==0,"selectUnlockedCells":v!=-4142}))
                .map_err(|e| failure(1004, e.to_string()))?;
            return Ok(Value::Empty);
        }
        if !reads(c) {
            return Err(failure(438, "read-only protection status"));
        }
        arguments(c, &[])?;
        match m.as_str() {
            "enableselection" => Ok(Value::Long(match sh.protection.as_ref().map(|p| (p.select_locked, p.select_unlocked)) {
                None | Some((true, true)) => 0,
                Some((false, true)) => 1,
                Some((false, false)) => -4142,
                Some((true, false)) => return Err(failure(1004, "native locked-only selection has no Excel enum value")),
            })),
            "protectionmode" => Ok(Value::Boolean(false)), // No UI-only bypass in the native model.
            "protectdrawingobjects" | "protectscenarios" => Ok(Value::Boolean(sh.is_protected())),
            _ => Err(failure(438, "unsupported protection status")),
        }
    }
    fn permission_dispatch(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let si = f.sheet.ok_or_else(|| failure(91, "worksheet deleted"))?;
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let m = c.member.to_ascii_lowercase();
        if f.kind == "Protection" {
            if !reads(c) {
                return Err(failure(438, "protection options are read-only"));
            }
            if m == "alloweditranges" {
                let child = Feature::new("AllowEditRanges", Some(si));
                return if c.arguments.is_empty() { self.feature(child) } else { self.permission_dispatch(s, &child, c) };
            }
            arguments(c, &[])?;
            let p = sh.protection.as_ref();
            let allowed = match m.as_str() {
                "allowformattingcells" => p.is_some_and(|p| p.format_cells),
                "allowformattingcolumns" => p.is_some_and(|p| p.format_columns),
                "allowformattingrows" => p.is_some_and(|p| p.format_rows),
                "allowinsertingcolumns" => p.is_some_and(|p| p.insert_columns),
                "allowinsertingrows" => p.is_some_and(|p| p.insert_rows),
                "allowdeletingcolumns" => p.is_some_and(|p| p.delete_columns),
                "allowdeletingrows" => p.is_some_and(|p| p.delete_rows),
                "allowsorting" => p.is_some_and(|p| p.sort),
                "allowfiltering" => p.is_some_and(|p| p.autofilter),
                "allowinsertinghyperlinks" | "allowusingpivottables" => false,
                _ => return Err(failure(438, "unsupported protection member")),
            };
            return Ok(Value::Boolean(allowed));
        }
        if f.kind == "AllowEditRanges" {
            if sh.edit_ranges.len() > super::MAX_OBJECTS {
                return Err(failure(7, "edit-range collection budget"));
            }
            if c.kind == InvokeKind::Enumerate {
                arguments(c, &[])?;
                return self.object_array(sh.edit_ranges.iter().map(|r| node(si, r).object()).collect());
            }
            if m == "add" && reads(c) {
                let a = arguments(c, &["Title", "Range", "Password"])?;
                let name = title(required(&a, 0)?)?;
                let Value::Object(reference) = required(&a, 1)? else {
                    return Err(failure(13, "Range required"));
                };
                let Object::Range(owner, r) = self.object(s, *reference)? else {
                    return Err(failure(13, "Range required"));
                };
                r.validate()?;
                if owner != si {
                    return Err(failure(1004, "edit range must belong to the receiver worksheet"));
                }
                if sh.edit_ranges.len() >= super::MAX_OBJECTS {
                    return Err(failure(7, "edit-range budget"));
                }
                let mut p = json!({"sheet":si,"action":"add","title":name,"range":r.areas.iter().map(|r|r.a1()).collect::<Vec<_>>().join(",")});
                if !matches!(a[2], Value::Missing) {
                    p["password"] = json!(password(&a[2])?);
                }
                self.objects.reserve(1)?;
                s.execute("review.allowEditRanges", p).map_err(|e| failure(1004, e.to_string()))?;
                let r = s
                    .doc()
                    .map_err(|e| failure(91, e.to_string()))?
                    .wb
                    .sheet(si)
                    .and_then(|sh| sh.edit_ranges.last())
                    .ok_or_else(|| failure(91, "edit range missing"))?;
                return self.feature(node(si, r));
            }
            if !reads(c) {
                return Err(failure(438, "read-only edit-range collection"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(c, &[])?;
                    Ok(Value::Long(sh.edit_ranges.len() as i32))
                }
                "item" | "" | "alloweditranges" => {
                    let a = arguments(c, &["Index"])?;
                    let v = required(&a, 0)?;
                    let r = if matches!(v, Value::String(_)) {
                        let name = string(v)?;
                        sh.edit_ranges.iter().find(|r| r.title.eq_ignore_ascii_case(&name))
                    } else {
                        usize::try_from(numeric_index(v)?.checked_sub(1).ok_or_else(|| failure(9, "edit-range index"))?)
                            .ok()
                            .and_then(|i| sh.edit_ranges.get(i))
                    }
                    .ok_or_else(|| failure(9, "edit range not found"))?;
                    self.feature(node(si, r))
                }
                _ => Err(failure(438, "unsupported edit-range collection member")),
            };
        }
        let (si, r) = lookup(s, f)?;
        let mut p = json!({"sheet":si,"title":r.title});
        if m == "title" && c.kind == InvokeKind::Let {
            let a = arguments(c, &["Value"])?;
            p["action"] = json!("modify");
            p["newTitle"] = json!(title(required(&a, 0)?)?);
        } else if matches!(m.as_str(), "delete" | "unprotect" | "changepassword") && c.kind == InvokeKind::Call {
            let a = arguments(c, if m == "delete" { &[] } else { &["Password"] })?;
            p["action"] = json!(match m.as_str() {
                "delete" => "delete",
                "unprotect" => "unlock",
                _ => "modify",
            });
            if m == "changepassword" {
                p["password"] = json!(password(required(&a, 0)?)?);
            }
            if m == "unprotect" && !matches!(a[0], Value::Missing) {
                p["password"] = json!(password(&a[0])?);
            }
        } else {
            if !reads(c) {
                return Err(failure(438, "unsupported edit-range assignment"));
            }
            arguments(c, &[])?;
            return match m.as_str() {
                "title" => values::vba_text(&r.title),
                "range" => self.intern(Object::Range(si, ranges::Range { areas: r.ranges, unit: ranges::Unit::Cells })),
                _ => Err(failure(438, "unsupported edit-range member; native user ACL enforcement is unavailable")),
            };
        }
        s.execute("review.allowEditRanges", p).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    let mut out = vec![
        ("Worksheet", "Protection", "Property Get Protection() As Protection"),
        ("Worksheet", "EnableSelection", "Property EnableSelection As XlEnableSelection"),
        ("Worksheet", "ProtectionMode", "Property Get ProtectionMode() As Boolean"),
        ("Worksheet", "ProtectDrawingObjects", "Property Get ProtectDrawingObjects() As Boolean"),
        ("Worksheet", "ProtectScenarios", "Property Get ProtectScenarios() As Boolean"),
        ("Range", "AllowEdit", "Property Get AllowEdit() As Boolean"),
        ("Protection", "AllowEditRanges", "Property Get AllowEditRanges() As AllowEditRanges"),
        ("AllowEditRanges", "Count", "Property Get Count() As Long"),
        ("AllowEditRanges", "Item", "Property Get Item(Index As Variant) As AllowEditRange"),
        ("AllowEditRanges", "Add", "Function Add(Title As String, Range As Range, Optional Password As Variant) As AllowEditRange"),
        ("AllowEditRange", "Title", "Property Title As String"),
        ("AllowEditRange", "Range", "Property Get Range() As Range"),
        ("AllowEditRange", "ChangePassword", "Sub ChangePassword(Password As String)"),
        ("AllowEditRange", "Unprotect", "Sub Unprotect(Optional Password As Variant)"),
        ("AllowEditRange", "Delete", "Sub Delete()"),
    ];
    out.extend([
        ("Protection", "AllowFormattingCells", "Property Get AllowFormattingCells() As Boolean"),
        ("Protection", "AllowFormattingColumns", "Property Get AllowFormattingColumns() As Boolean"),
        ("Protection", "AllowFormattingRows", "Property Get AllowFormattingRows() As Boolean"),
        ("Protection", "AllowInsertingColumns", "Property Get AllowInsertingColumns() As Boolean"),
        ("Protection", "AllowInsertingRows", "Property Get AllowInsertingRows() As Boolean"),
        ("Protection", "AllowDeletingColumns", "Property Get AllowDeletingColumns() As Boolean"),
        ("Protection", "AllowDeletingRows", "Property Get AllowDeletingRows() As Boolean"),
        ("Protection", "AllowSorting", "Property Get AllowSorting() As Boolean"),
        ("Protection", "AllowFiltering", "Property Get AllowFiltering() As Boolean"),
        ("Protection", "AllowInsertingHyperlinks", "Property Get AllowInsertingHyperlinks() As Boolean"),
        ("Protection", "AllowUsingPivotTables", "Property Get AllowUsingPivotTables() As Boolean"),
    ]);
    out.into_iter()
        .map(|(owner, name, signature)| super::MemberDescriptor {
            owner,
            name,
            signature,
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/permission_api.rs::Bridge::permission_api",
        })
        .collect()
}
