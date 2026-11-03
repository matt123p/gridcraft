//! Ordinary High protection bindings shared by all feature families.
use super::features::{absent, reads};
use super::objects::{arguments, string};
use super::{Bridge, HostCall, HostResult, Object, Session, Value, failure};
use serde_json::json;
impl Bridge {
    pub(crate) fn protection_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if !matches!(object, Object::Sheet(_)) || !reads(call) {
            return None;
        }
        let member = call.member.to_ascii_lowercase();
        if !matches!(member.as_str(), "protect" | "unprotect" | "protectcontents") {
            return None;
        }
        Some((|| {
            let doc = session.doc().map_err(|e| failure(91, e.to_string()))?;
            if member == "protectcontents"
                && let Object::Sheet(sheet) = object
            {
                arguments(call, &[])?;
                return Ok(Value::Boolean(doc.wb.sheet(*sheet).ok_or_else(|| failure(91, "worksheet deleted"))?.is_protected()));
            }
            if !matches!(member.as_str(), "protect" | "unprotect") {
                return Err(failure(438, "unsupported protection status"));
            }
            if call.kind != super::InvokeKind::Call {
                return Err(failure(438, "protection mutations require a method call"));
            }
            let unprotect = member == "unprotect";
            let args = arguments(
                call,
                if unprotect {
                    &["Password"]
                } else {
                    &[
                        "Password",
                        "DrawingObjects",
                        "Contents",
                        "Scenarios",
                        "UserInterfaceOnly",
                        "AllowFormattingCells",
                        "AllowFormattingColumns",
                        "AllowFormattingRows",
                        "AllowInsertingColumns",
                        "AllowInsertingRows",
                        "AllowInsertingHyperlinks",
                        "AllowDeletingColumns",
                        "AllowDeletingRows",
                        "AllowSorting",
                        "AllowFiltering",
                        "AllowUsingPivotTables",
                    ]
                },
            )?;
            let mut params = json!({});
            if !absent(&args[0]) {
                let password = string(&args[0])?;
                if password.len() > 1024 {
                    return Err(failure(7, "password payload budget"));
                }
                params["password"] = json!(password);
            }
            let command = {
                let Object::Sheet(sheet) = object else {
                    return Err(failure(91, "worksheet missing"));
                };
                params["sheet"] = json!(sheet);
                if !unprotect {
                    for i in [1, 2, 3] {
                        if !absent(&args[i]) && !super::address::boolean(&args[i])? {
                            return Err(failure(1004, "independent object/content/scenario protection is unavailable"));
                        }
                    }
                    for i in [4, 10, 15] {
                        if !absent(&args[i]) && super::address::boolean(&args[i])? {
                            return Err(failure(1004, "this protection mode is unavailable"));
                        }
                    }
                    for (i, key) in [
                        (5, "formatCells"),
                        (6, "formatColumns"),
                        (7, "formatRows"),
                        (8, "insertColumns"),
                        (9, "insertRows"),
                        (11, "deleteColumns"),
                        (12, "deleteRows"),
                        (13, "sort"),
                        (14, "autofilter"),
                    ] {
                        if !absent(&args[i]) {
                            params[key] = json!(super::address::boolean(&args[i])?);
                        }
                    }
                }
                if unprotect { "review.unprotectSheet" } else { "review.protectSheet" }
            };
            session.execute(command, params).map_err(|e| failure(1004, e.to_string()))?;
            Ok(Value::Empty)
        })())
    }
}
pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [("Worksheet","Protect","Sub Protect(Optional Password As Variant, Optional DrawingObjects As Variant, Optional Contents As Variant, Optional Scenarios As Variant, Optional UserInterfaceOnly As Variant, Optional AllowFormattingCells As Variant, Optional AllowFormattingColumns As Variant, Optional AllowFormattingRows As Variant, Optional AllowInsertingColumns As Variant, Optional AllowInsertingRows As Variant, Optional AllowInsertingHyperlinks As Variant, Optional AllowDeletingColumns As Variant, Optional AllowDeletingRows As Variant, Optional AllowSorting As Variant, Optional AllowFiltering As Variant, Optional AllowUsingPivotTables As Variant)"),("Worksheet","Unprotect","Sub Unprotect(Optional Password As Variant)"),("Worksheet","ProtectContents","Property Get ProtectContents() As Boolean")].into_iter().map(|(owner,name,signature)|super::MemberDescriptor{owner,name,signature,availability:super::MemberAvailability::Implemented,implementation_symbol:"crates/vba/src/protection_api.rs::Bridge::protection_api"}).collect()
}
