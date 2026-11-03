//! OM09 native protection permission and edit-range acceptance.
use gridcraft_engine::{Session, core::CellRef};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Protection.bas", source)?;
    b.configure(&mut ide)?;
    Ok((s, b, ide))
}
fn finish(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    b.dispatch(s, ide, Action::Run)?;
    for _ in 0..10000 {
        b.tick(s, ide)?;
        if !matches!(ide.mode(), Mode::Running | Mode::Waiting) {
            break;
        }
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn typed_permissions_and_selection_preserve_native_password_and_rights() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim p As Protection
Set p=ActiveSheet.Protection
Range("B2").Locked=False
ActiveSheet.Protect Password:="secret", AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowInsertingColumns:=True, AllowDeletingRows:=True, AllowDeletingColumns:=True, AllowSorting:=True, AllowFiltering:=True
If Not (p Is ActiveSheet.Protection) Then Err.Raise 5
If Not p.AllowFormattingCells Or Not p.AllowFormattingColumns Or Not p.AllowFormattingRows Then Err.Raise 5
If Not p.AllowInsertingRows Or Not p.AllowInsertingColumns Then Err.Raise 5
If Not p.AllowDeletingRows Or Not p.AllowDeletingColumns Then Err.Raise 5
If Not p.AllowSorting Or Not p.AllowFiltering Then Err.Raise 5
If p.AllowInsertingHyperlinks Or p.AllowUsingPivotTables Then Err.Raise 5
If ActiveSheet.ProtectionMode Then Err.Raise 5
If Not ActiveSheet.ProtectDrawingObjects Or Not ActiveSheet.ProtectScenarios Then Err.Raise 5
ActiveSheet.EnableSelection=xlUnlockedCells
If ActiveSheet.EnableSelection <> xlUnlockedCells Then Err.Raise 5
On Error Resume Next
Range("A1").Select
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
Range("B2").Select
ActiveSheet.EnableSelection=xlNoSelection
If ActiveSheet.EnableSelection <> xlNoSelection Then Err.Raise 5
ActiveSheet.EnableSelection=xlNoRestrictions
On Error Resume Next
ActiveSheet.Unprotect "bad"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If Not ActiveSheet.ProtectContents Then Err.Raise 5
ActiveSheet.Unprotect "secret"
If ActiveSheet.ProtectContents Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn edit_ranges_use_native_authentication_and_locked_cell_policy() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim ranges As AllowEditRanges, secure As AllowEditRange, entry As AllowEditRange
Set ranges=ActiveSheet.Protection.AllowEditRanges
Set secure=ranges.Add("Secret",Range("A1:A2"),"range-pw")
ranges.Add "Open",Range("B2,C3")
If ranges.Count <> 2 Then Err.Raise 5
If Not (secure Is ranges("secret")) Then Err.Raise 5
If secure.Range.Address <> "$A$1:$A$2" Then Err.Raise 5
For Each entry In ranges
If entry.Title = "" Then Err.Raise 5
Next
Range("D4").Locked=False
ActiveSheet.Protect "sheet-pw"
If Range("A1").AllowEdit Then Err.Raise 5
If Not Range("B2,C3,D4").AllowEdit Then Err.Raise 5
If Range("A1:B2").AllowEdit Then Err.Raise 5
On Error Resume Next
secure.Unprotect "wrong"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
Range("A1").Value2=9
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
secure.Unprotect "range-pw"
If Not Range("A1:A2").AllowEdit Then Err.Raise 5
Range("A1").Value2=11
Range("B2").Value2=12
If Not ActiveSheet.ProtectContents Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::parse("A1").ok_or("cell")?).display(), "11");
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn edit_range_identity_and_grants_do_not_rebind_after_password_change_or_recreation() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim old As AllowEditRange, fresh As AllowEditRange, x As String
Set old=ActiveSheet.Protection.AllowEditRanges.Add("First",Range("A1"),"old")
old.Unprotect "old"
old.Title="Renamed"
If Not (old Is ActiveSheet.Protection.AllowEditRanges("Renamed")) Then Err.Raise 5
old.ChangePassword "new"
ActiveSheet.Protect
If Range("A1").AllowEdit Then Err.Raise 5
On Error Resume Next
old.Unprotect "old"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
old.Unprotect "new"
If Not Range("A1").AllowEdit Then Err.Raise 5
ActiveSheet.Unprotect
old.Delete
Set fresh=ActiveSheet.Protection.AllowEditRanges.Add("Renamed",Range("A1"),"other")
ActiveSheet.Protect
If Range("A1").AllowEdit Then Err.Raise 5
On Error Resume Next
x=old.Title
If Err.Number <> 91 Then Err.Raise 5
Err.Clear
On Error GoTo 0
fresh.Unprotect "other"
If Not Range("A1").AllowEdit Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn edit_ranges_target_inactive_sheets_and_follow_sheet_reordering() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim owner As Worksheet, entry As AllowEditRange
Set owner=Worksheets("Other")
Set entry=owner.Protection.AllowEditRanges.Add("Inputs",owner.Range("D4"))
owner.Name="Moved"
owner.Move Before:=Worksheets(1)
If entry.Range.Address <> "$D$4" Then Err.Raise 5
If Not (entry Is owner.Protection.AllowEditRanges(1)) Then Err.Raise 5
owner.Protect
If Not owner.Range("D4").AllowEdit Then Err.Raise 5
If owner.Range("E5").AllowEdit Then Err.Raise 5
On Error Resume Next
owner.Protection.AllowEditRanges.Add "Bad",Worksheets("Sheet1").Range("F6")
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If owner.Protection.AllowEditRanges.Count <> 1 Then Err.Raise 5
End Sub"#,
    )?;
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("sheet.activate", json!({"sheet":0}))?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unsupported_permissions_and_protected_edits_refuse_before_effects() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim entry As AllowEditRange, x As Variant
Set entry=ActiveSheet.Protection.AllowEditRanges.Add("One",Range("A1"))
On Error Resume Next
ActiveSheet.EnableSelection=xlNoSelection
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
x=ActiveSheet.Protect
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
ActiveSheet.Protect UserInterfaceOnly:=True
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
ActiveSheet.Protect AllowUsingPivotTables:=True
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
If ActiveSheet.ProtectContents Then Err.Raise 5
ActiveSheet.Protect
entry.Title="Bad"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
entry.ChangePassword "Bad"
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
entry.Delete
If Err.Number <> 1004 Then Err.Raise 5
Err.Clear
Set x=entry.Users
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
x=Range("A:A").AllowEdit
If Err.Number <> 7 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If entry.Title <> "One" Then Err.Raise 5
If ActiveSheet.Protection.AllowEditRanges.Count <> 1 Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn sheet_and_range_protection_hashes_roundtrip_but_session_unlocks_do_not() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim entry As AllowEditRange
Set entry=ActiveSheet.Protection.AllowEditRanges.Add("Secret",Range("A1:B2"),"range-pw")
ActiveSheet.Protect Password:="sheet-pw",AllowFiltering:=True
ActiveSheet.EnableSelection=xlUnlockedCells
entry.Unprotect "range-pw"
If Not Range("A1").AllowEdit Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let sh = wb.sheet(0).ok_or("sheet")?;
    let p = sh.protection.as_ref().ok_or("protection")?;
    assert!(p.password.as_ref().ok_or("password")?.verify("sheet-pw")?);
    assert!(p.autofilter);
    assert!(!p.select_locked);
    assert!(sh.edit_ranges.first().and_then(|r| r.password.as_ref()).ok_or("range password")?.verify("range-pw")?);
    let d = gridcraft_engine::DocState::new(wb, None, "Reopened".into(), gridcraft_engine::Trust::FromFile);
    assert!(d.unlocked_ranges.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_definition_undo_preserves_identity_and_revokes_prior_unlocks() -> R {
    let mut s = Session::new();
    s.new_workbook();
    s.execute("review.allowEditRanges", json!({"action":"add","title":"Secret","range":"A1","password":"old"}))?;
    let identity = s.doc()?.wb.sheet(0).and_then(|sh| sh.edit_ranges.first()).ok_or("range")?.identity.clone();
    s.execute("review.allowEditRanges", json!({"action":"unlock","title":"Secret","password":"old"}))?;
    assert_eq!(s.doc()?.unlocked_ranges.len(), 1);
    s.execute("review.allowEditRanges", json!({"action":"modify","title":"Secret","password":"new"}))?;
    assert!(s.doc()?.unlocked_ranges.is_empty());
    assert!(identity.same(&s.doc()?.wb.sheet(0).and_then(|sh| sh.edit_ranges.first()).ok_or("range")?.identity));
    s.execute("review.allowEditRanges", json!({"action":"delete","title":"Secret"}))?;
    s.execute("edit.undo", json!({}))?;
    assert!(identity.same(&s.doc()?.wb.sheet(0).and_then(|sh| sh.edit_ranges.first()).ok_or("range")?.identity));
    assert!(s.doc()?.unlocked_ranges.is_empty());
    s.execute("review.protectSheet", json!({"password":"sheet"}))?;
    assert!(s.execute("review.protectSheet", json!({"selectionOnly":true,"password":"replacement","selectLockedCells":false})).is_err());
    assert!(s.doc()?.wb.sheet(0).and_then(|sh| sh.protection.as_ref()).and_then(|p| p.password.as_ref()).ok_or("password")?.verify("sheet")?);
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn permissions_retain_target_consent_and_do_not_claim_irm_or_encryption_state() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim x As Variant
On Error Resume Next
Set x=Workbooks(2).Worksheets(1).Protection
If Err.Number <> 70 Then Err.Raise 5
Err.Clear
Set x=ThisWorkbook.Permission
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
Set x=Application.ProtectedViewWindows
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
ThisWorkbook.Password="does-not-encrypt"
If Err.Number <> 438 Then Err.Raise 5
Err.Clear
On Error GoTo 0
If ActiveSheet.ProtectContents Then Err.Raise 5
End Sub"#,
    )?;
    let first = s.active_window().ok_or("window")?;
    s.new_workbook();
    s.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
    s.activate_window(first);
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn deleting_an_edit_range_owner_revokes_its_session_grants() -> R {
    let mut s = Session::new();
    s.new_workbook();
    s.execute("review.allowEditRanges", json!({"action":"add","title":"Secret","range":"A1","password":"pw"}))?;
    s.execute("review.allowEditRanges", json!({"action":"unlock","title":"Secret","password":"pw"}))?;
    assert_eq!(s.doc()?.unlocked_ranges.len(), 1);
    s.execute("home.insertSheet", json!({"name":"Other"}))?;
    s.execute("home.deleteSheet", json!({"sheet":0}))?;
    assert!(s.doc()?.unlocked_ranges.is_empty());
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.unlocked_ranges.is_empty());
    Ok(())
}
