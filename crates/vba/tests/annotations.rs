//! OM19 local native profile, exercised through real VBA dispatch.
use gridcraft_engine::{Session, core::CellRef};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use serde_json::json;
type R = Result<(), Box<dyn std::error::Error>>;
fn setup(source: &str) -> Result<(Session, Bridge, Ide), Box<dyn std::error::Error>> {
    let mut s = Session::new();
    s.new_workbook();
    let b = Bridge::new(&s)?;
    let mut ide = Ide::new();
    ide.add_module("Annotations.bas", source)?;
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
fn notes_text_visibility_collections_and_structural_rebase() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim n As Comment, other As Comment, all As Comments
If Not Range("B2").Comment Is Nothing Then Err.Raise 5
Set n=Range("B2").AddComment("abcd")
Set other=Range("D4").AddComment
Set all=ActiveSheet.Comments
If all.Count<>2 Then Err.Raise 5
If Not (n Is all(1)) Then Err.Raise 5
If n.Parent.Address<>"$B$2" Then Err.Raise 5
If Not (n.Application Is Application) Then Err.Raise 5
If Not (all.Parent Is ActiveSheet) Then Err.Raise 5
If n.Author="" Then Err.Raise 5
n.Text "XY",2,False
If n.Text<>"aXYbcd" Then Err.Raise 5
n.Text "z",3,True
If n.Text<>"aXzbcd" Then Err.Raise 5
n.Visible=True
n.Visible=True
If Not n.Visible Then Err.Raise 5
n.Visible=False
If n.Visible Then Err.Raise 5
If Not (n.Next Is other) Then Err.Raise 5
If Not (other.Previous Is n) Then Err.Raise 5
If Not other.Next Is Nothing Then Err.Raise 5
For Each other In all
If other.Text="" Then other.Text "empty"
Next
Rows(1).Insert
If n.Parent.Address<>"$B$3" Then Err.Raise 5
If Not (n Is Range("B3").Comment) Then Err.Raise 5
On Error Resume Next
Range("B3").AddComment "duplicate"
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
n.Text "x",0
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>5 Then Err.Raise 5
On Error Resume Next
n.Delete
n.Text
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>91 Then Err.Raise 5
On Error Resume Next
On Error GoTo 0
If all.Count<>1 Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert_eq!(s.doc()?.wb.sheet(0).ok_or("sheet")?.comments.len(), 1);
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.comments.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn local_threads_and_replies_preserve_native_text_without_remote_identity() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim t As CommentThreaded, r As CommentThreaded, replies As CommentsThreaded
Set t=Range("A1").AddCommentThreaded("thread")
Set r=t.AddReply("reply")
Set replies=t.Replies
If replies.Count<>1 Then Err.Raise 5
If Not (r Is replies(1)) Then Err.Raise 5
If Not (r.Parent Is t) Then Err.Raise 5
If r.Text<>"reply" Then Err.Raise 5
If ActiveSheet.Comments.Count<>0 Then Err.Raise 5
If ActiveSheet.CommentsThreaded.Count<>1 Then Err.Raise 5
If Not Range("A1").Comment Is Nothing Then Err.Raise 5
For Each r In replies
If r.Text<>"reply" Then Err.Raise 5
Next
If Not (replies.Parent Is t) Then Err.Raise 5
t.Text "edited"
If t.Text<>"edited" Or replies(1).Text<>"reply" Then Err.Raise 5
On Error Resume Next
r.Text "unsupported"
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
Dim person As Object
Set person=t.Author
On Error GoTo 0
If person.Name="" Then Err.Raise 5
On Error Resume Next
x=person.UserID
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
On Error GoTo 0
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    let c = s.doc()?.wb.sheet(0).ok_or("sheet")?.comments.get(&CellRef::new(0, 0)).ok_or("comment")?;
    assert_eq!(c.text, "edited");
    assert_eq!(c.replies.len(), 1);
    assert!(c.threaded);
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    let c = wb.sheet(0).ok_or("sheet")?.comments.get(&CellRef::new(0, 0)).ok_or("thread")?;
    assert_eq!(c.text, "edited");
    assert_eq!(c.replies.len(), 1);
    assert_eq!(c.replies[0].1, "reply");
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn hyperlinks_and_builtin_metadata_round_trip_and_undo() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim h As Hyperlink, links As Hyperlinks, p As Office.DocumentProperties, prop As Office.DocumentProperty
Set links=ActiveSheet.Hyperlinks
Set h=links.Add(Range("B2"),"https://example.org", "section", "tip", "label")
If links.Count<>1 Then Err.Raise 5
If Not (h Is Range("B2").Hyperlinks(1)) Then Err.Raise 5
If h.Address<>"https://example.org" Or h.SubAddress<>"section" Then Err.Raise 5
If h.TextToDisplay<>"label" Or h.ScreenTip<>"tip" Then Err.Raise 5
If h.Range.Address<>"$B$2" Or h.Type<>0 Then Err.Raise 5
h.ScreenTip="changed"
h.TextToDisplay="new label"
h.Address="https://example.net"
h.SubAddress="new"
If h.Address<>"https://example.net" Or h.SubAddress<>"new" Then Err.Raise 5
If h.ScreenTip<>"changed" Or h.TextToDisplay<>"new label" Then Err.Raise 5
Set p=ThisWorkbook.BuiltinDocumentProperties
If p.Count<>6 Then Err.Raise 5
Set prop=p("Title")
If prop.Name<>"Title" Or prop.Type<>4 Or prop.LinkToContent Then Err.Raise 5
If Not (prop.Parent Is p) Then Err.Raise 5
prop.Value="Report"
p(21).Value="Craft"
p("Comments").Value="Description"
If p(1).Value<>"Report" Or p("Company").Value<>"Craft" Then Err.Raise 5
For Each h In links
If h.Type<>0 Then Err.Raise 5
Next
Rows(1).Insert
If h.Range.Address<>"$B$3" Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert_eq!(wb.props.title, "Report");
    assert_eq!(wb.props.company, "Craft");
    assert_eq!(wb.props.description, "Description");
    let h = wb.sheet(0).ok_or("sheet")?.hyperlinks.get(&CellRef::new(2, 1)).ok_or("link")?;
    assert_eq!(h.target, "https://example.net#new");
    assert_eq!(h.tooltip.as_deref(), Some("changed"));
    s.execute("edit.undo", json!({}))?;
    assert!(s.doc()?.wb.props.title.is_empty());
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.hyperlinks.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn annotations_fail_without_side_effects_on_protected_or_untrusted_targets() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
ActiveSheet.Protect
On Error Resume Next
Range("A1").AddComment "blocked"
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
ActiveSheet.Hyperlinks.Add Range("A1"), "https://example.org"
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
Set x=Workbooks(2).Worksheets(1).Comments
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>70 Then Err.Raise 5
On Error Resume Next
On Error GoTo 0
End Sub"#,
    )?;
    s.new_workbook();
    s.doc_mut()?.trust = gridcraft_engine::Trust::FromFile;
    s.set_active(0);
    finish(&mut s, &mut b, &mut ide)?;
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.comments.is_empty());
    assert!(s.doc()?.wb.sheet(0).ok_or("sheet")?.hyperlinks.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn note_text_unicode_and_copy_move_keep_native_payloads() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Range("A1").NoteText "abcdef"
Range("A1").NoteText "XY",2,2
If Range("A1").NoteText<>"aXYdef" Then Err.Raise 5
If Range("A1").NoteText(,2,2)<>"XY" Then Err.Raise 5
Range("A1").NoteText "!",100
If Range("A1").Comment.Text<>"aXYdef!" Then Err.Raise 5
Range("A1").Copy Range("B2")
Range("B2").Cut Range("D4")
If Not Range("B2").Comment Is Nothing Then Err.Raise 5
If Range("D4").Comment.Text<>"aXYdef!" Then Err.Raise 5
Range("F6").AddComment "A😀B"
Range("F6").Comment.Text "Z",4
If Range("F6").Comment.Text<>"A😀Z" Then Err.Raise 5
On Error Resume Next
Range("F6").Comment.Text "x",3,False
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>5 Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    let bytes = gridcraft_xlsx::write_xlsx(&s.doc()?.wb)?;
    let (wb, _) = gridcraft_xlsx::read_xlsx(&bytes)?;
    assert_eq!(wb.sheet(0).ok_or("sheet")?.comments.get(&CellRef::new(3, 3)).ok_or("note")?.text, "aXYdef!");
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn deleted_annotation_handles_do_not_alias_replacements_and_reply_navigation_is_local() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim old As Comment, fresh As Comment, h As Hyperlink, t As CommentThreaded, first As CommentThreaded, last As CommentThreaded
Set old=Range("A1").AddComment("old")
old.Delete
Set fresh=Range("A1").AddComment("fresh")
If old Is fresh Then Err.Raise 5
On Error Resume Next
old.Text
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>91 Then Err.Raise 5
Set h=ActiveSheet.Hyperlinks.Add(Range("B2"),"https://example.org")
h.Delete
ActiveSheet.Hyperlinks.Add Range("B2"),"https://example.net"
On Error Resume Next
x=h.Address
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>91 Then Err.Raise 5
Set t=Range("C3").AddCommentThreaded("thread")
Set first=t.AddReply("first")
Set last=t.AddReply("last")
If Not (first.Next Is last) Then Err.Raise 5
If Not (last.Previous Is first) Then Err.Raise 5
If Not last.Next Is Nothing Then Err.Raise 5
If Not first.Previous Is Nothing Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn internal_follow_and_external_service_refusal_are_explicit() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
Dim h As Hyperlink
Set h=ActiveSheet.Hyperlinks.Add(Range("A1"),"","Sheet1!D4")
If h.Address<>"" Or h.SubAddress<>"Sheet1!D4" Then Err.Raise 5
h.Follow
If ActiveCell.Address<>"$D$4" Then Err.Raise 5
Set h=ActiveSheet.Hyperlinks.Add(Range("B2"),"https://example.org")
On Error Resume Next
h.Follow
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
If ActiveCell.Address<>"$D$4" Then Err.Raise 5
Range("B2").Hyperlinks.Delete
If ActiveSheet.Hyperlinks.Count<>1 Then Err.Raise 5
On Error Resume Next
ThisWorkbook.BuiltinDocumentProperties(6).Value="unsupported"
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>9 Then Err.Raise 5
End Sub"#,
    )?;
    finish(&mut s, &mut b, &mut ide)?;
    assert!(s.ui_requests.is_empty());
    Ok(())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn oversized_native_annotation_text_and_invalid_arguments_refuse_before_allocation() -> R {
    let (mut s, mut b, mut ide) = setup(
        r#"Sub Main()
On Error Resume Next
x=Range("A1").Comment.Text
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>7 Then Err.Raise 5
On Error Resume Next
x=Range("A1").NoteText
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>7 Then Err.Raise 5
On Error Resume Next
ActiveSheet.Hyperlinks.Add Range("B2:C3"),"https://example.org"
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>1004 Then Err.Raise 5
On Error Resume Next
Range("C3").AddComment WrongName:="not ignored"
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>448 Then Err.Raise 5
On Error Resume Next
Range("D4").NoteText Space(256)
actual=Err.Number
Err.Clear
On Error GoTo 0
If actual<>5 Then Err.Raise 5
End Sub"#,
    )?;
    s.execute("review.newNote", json!({"cell":"A1","text":"seed"}))?;
    std::sync::Arc::make_mut(&mut s.doc_mut()?.wb).sheet_mut(0).ok_or("sheet")?.comments.get_mut(&CellRef::new(0, 0)).ok_or("note")?.text =
        "x".repeat(1_048_577);
    finish(&mut s, &mut b, &mut ide)?;
    let sh = s.doc()?.wb.sheet(0).ok_or("sheet")?;
    assert_eq!(sh.comments.len(), 1);
    assert!(sh.hyperlinks.is_empty());
    Ok(())
}
