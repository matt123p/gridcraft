//! Authored native form import/attachment evidence; no Office or downloaded macro execution.
#[path = "support/native_forms.rs"]
mod fixture;
mod support;
use gridcraft_engine::{DocState, Session, Trust, model::Workbook};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use gridcraft_xlsx::{XlsxKind, read_xlsx, write_xlsx_as};
type R = Result<(), Box<dyn std::error::Error>>;
fn settle(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    for _ in 0..5000 {
        if ide.mode() != Mode::Running {
            return Ok(());
        }
        b.tick(s, ide)?;
    }
    Err("native form timeout".into())
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_forms_attach_from_xlsm_inertly_run_after_trust_and_retain_original_resources() -> R {
    let mut wb = Workbook::new();
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", fixture::SOURCE),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    wb.vba_project = Some(original.clone().into());
    let package = write_xlsx_as(&wb, XlsxKind::MacroWorkbook)?;
    let (wb, _) = read_xlsx(&package)?;
    let mut s = Session::new();
    s.add_document(DocState::from_saved(wb, None, "Forms.xlsm".into(), Trust::FromFile));
    let mut b = Bridge::new(&s)?;
    let mut ide = b.import_project(&s)?;
    let id = ide.modules.iter().find(|m| m.name == "NativeForm.frm").ok_or("native source")?.id;
    assert!(ide.form(id).is_some());
    assert!(ide.output().is_empty());
    assert_eq!(ide.mode(), Mode::Design);
    assert!(ide.native_form(id).is_some());
    ide.entry = "Main".into();
    assert!(b.dispatch(&mut s, &mut ide, Action::Run).is_err());
    b.consent(true);
    b.dispatch(&mut s, &mut ide, Action::Run)?;
    settle(&mut s, &mut b, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Waiting, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["initialized".to_string()]);
    ide.form_event(id, Some(13), None, &mut b.host(&mut s))?;
    settle(&mut s, &mut b, &mut ide)?;
    b.tick(&mut s, &mut ide)?;
    settle(&mut s, &mut b, &mut ide)?;
    assert_eq!(ide.output(), vec!["initialized".to_string(), "initial".to_string()]);
    b.dispatch(&mut s, &mut ide, Action::Reset)?;
    assert_eq!(ide.modules.iter().find(|m| m.id == id).ok_or("source")?.text, fixture::SOURCE);
    let saved = write_xlsx_as(s.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

fn boolean_project(project: Vec<u8>) -> Vec<u8> {
    form_project(project, fixture::boolean_companion())
}
fn form_project(project: Vec<u8>, resource: Vec<u8>) -> Vec<u8> {
    use std::io::{Cursor, Read, Write};
    let mut file = cfb::CompoundFile::open(Cursor::new(fixture::add_designer(project))).unwrap();
    let mut resource = cfb::CompoundFile::open(Cursor::new(resource[4..].to_vec())).unwrap();
    let streams = resource.walk().filter(|entry| entry.is_stream()).map(|entry| entry.path().to_path_buf()).collect::<Vec<_>>();
    for name in streams {
        let mut bytes = Vec::new();
        resource.open_stream(&name).unwrap().read_to_end(&mut bytes).unwrap();
        let target = format!("/NativeForm{}", name.display());
        let parent = std::path::Path::new(&target).parent().unwrap();
        file.create_storage_all(parent).unwrap();
        file.create_stream(&target).unwrap().write_all(&bytes).unwrap();
    }
    file.into_inner().into_inner()
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_boolean_instances_run_through_the_workbook_bridge_and_keep_package_bytes() -> R {
    let original = boolean_project(support::project(&[
        ("NativeForm", "BaseClass", fixture::SOURCE),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,c As MSForms.Control,u As MSForms.UserForm,n As Long\nDebug.Print VBA.UserForms.Count\nf.Check1.Value=True\nf.ChoiceA.Value=True\nf.ChoiceB.Value=True\nDebug.Print f.ChoiceA.Value\nDebug.Print f.ChoiceB.Value\nf.Show 0\nDebug.Print f.ActiveControl Is f.Entry\nFor Each c In f.Controls\nn=n+1\nNext c\nDebug.Print n\nn=0\nFor Each u In VBA.UserForms\nIf u Is f Then n=n+1\nNext u\nDebug.Print n\nEnd Sub",
        ),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Forms.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    assert!(ide.output().is_empty());
    ide.entry = "Main".into();
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["0", "initialized", "False", "True", "True", "7", "1"]);
    let instance = *ide.form_instances()?.first().ok_or("instance")?;
    assert_eq!(ide.form_instance_field(instance, Some(14), "Value")?, vba_values::Value::Boolean(true));
    ide.form_instance_event(instance, Some(14), None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(instance, Some(14), "Value")?, vba_values::Value::Boolean(false));
    assert_eq!(ide.form_instance_field(instance, None, "VbaActiveControl")?, ide.form_instance_field(instance, None, "VbaCheck1")?);
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_text_selection_and_input_preserve_xlsm_project_bytes() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,t As MSForms.TextBox\nSet t=f.Entry\nt.Text=\"abcde\"\nt.SelStart=1\nt.SelLength=2\nt.SelText=\"X\"\nf.Show 0\nDebug.Print t.MultiLine\nDebug.Print t.PasswordChar\nDebug.Print t.Text\nDebug.Print t.SelStart\nEnd Sub",
            ),
        ]),
        fixture::text_companion(0xacf0481b, 42),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Text.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output(), vec!["initialized", "True", "*", "aXde", "2"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    ide.form_instance_text(f, 12, vba_values::VbaString::from_utf8("z")?, (1, 0), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaSelStart")?, vba_values::Value::Long(1));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_numeric_controls_keep_xlsm_identity_and_canonical_input() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,s As MSForms.SpinButton\nSet s=f.Spin1\ns.Value=7\nf.Show 0\nDebug.Print s.Value\nDebug.Print f.Scroll1.Min\nEnd Sub",
            ),
        ]),
        fixture::numeric_companion(),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Numbers.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output(), vec!["initialized", "7", "10"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    ide.form_instance_numeric(f, 14, vba_ide::FormNumericInput::Step { up: true, large: false }, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    ide.form_instance_numeric(f, 15, vba_ide::FormNumericInput::Step { up: true, large: true }, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(14), "Value")?, vba_values::Value::Long(9));
    assert_eq!(ide.form_instance_field(f, Some(15), "Value")?, vba_values::Value::Long(-5));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_listbox_rows_and_authorized_input_preserve_xlsm_project_bytes() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,l As MSForms.ListBox\nSet l=f.List1\nDim rows(-1 To 0,3 To 4) As String\nrows(-1,3)=\"alpha\"\nrows(-1,4)=\"A\"\nrows(0,3)=\"beta\"\nrows(0,4)=\"B\"\nl.List=rows\nrows(-1,4)=\"changed source\"\nDim snapshot As Variant\nsnapshot=l.List\nl.List=snapshot\nsnapshot(0,1)=\"detached snapshot\"\nl.ListIndex=0\nf.Show 0\nDebug.Print l.ListCount\nDebug.Print l.Column(1)\nEnd Sub",
            ),
        ]),
        fixture::list_companion(),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Lists.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output(), vec!["initialized", "2", "A"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    ide.form_instance_list(f, 14, 1, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(14), "ListIndex")?, vba_values::Value::Long(1));
    assert_eq!(ide.form_instance_field(f, Some(14), "Value")?, vba_values::Value::String(vba_values::VbaString::from_utf8("B")?));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_combobox_text_selection_and_input_preserve_xlsm_project_bytes() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,l As MSForms.ComboBox\nSet l=f.Combo1\nDim columns(3 To 4,-1 To 0) As String\ncolumns(3,-1)=\"alpha\"\ncolumns(4,-1)=\"A\"\ncolumns(3,0)=\"beta\"\ncolumns(4,0)=\"B\"\nl.Column=columns\ncolumns(4,-1)=\"changed source\"\nDim snapshot As Variant\nsnapshot=l.Column\nl.Column=snapshot\nsnapshot(1,0)=\"detached snapshot\"\nl.ListIndex=0\nf.Show 0\nDebug.Print l.ListCount\nDebug.Print l.Column(1)\nEnd Sub",
            ),
        ]),
        fixture::combo_companion(3, 2),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Combos.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output(), vec!["initialized", "2", "A"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    ide.form_instance_list(f, 14, 1, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(14), "ListIndex")?, vba_values::Value::Long(1));
    assert_eq!(ide.form_instance_field(f, Some(14), "Value")?, vba_values::Value::String(vba_values::VbaString::from_utf8("B")?));
    ide.form_instance_text(f, 14, vba_values::VbaString::from_utf8("free😀")?, (6, 0), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.form_instance_field(f, Some(14), "ListIndex")?, vba_values::Value::Long(-1));
    assert_eq!(ide.form_instance_field(f, Some(14), "Value")?, vba_values::Value::String(vba_values::VbaString::from_utf8("free😀")?));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_metadata_and_control_tips_keep_original_xlsm_resources() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,c As MSForms.Control
Set c=f.Entry
Debug.Print c.Tag,c.ControlTipText,c.HelpContextID
c.Tag="runtime tag"
c.ControlTipText="Runtime help"
c.HelpContextID=42
f.Tag="form tag"
f.HelpContextID=43
Debug.Print f.Tag,f.HelpContextID,c.Tag,c.ControlTipText,c.HelpContextID
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::metadata_companion(),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Metadata.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output(), vec!["initialized", "tag-12\ttip-12\t112", "form tag\t43\truntime tag\tRuntime help\t42"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    assert_eq!(ide.form_instance_field(f, None, "HelpContextID")?, vba_values::Value::Long(43));
    assert_eq!(ide.form_instance_field(f, Some(12), "ControlTipText")?, vba_values::Value::String(vba_values::VbaString::from_utf8("Runtime help")?));
    let native = ide.native_form(f.source).ok_or("native")?.clone();
    ide.form_instance_text(f, 12, vba_values::VbaString::from_utf8("edited")?, (6, 0), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.native_form(f.source), Some(&native));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_nested_frame_collections_and_input_preserve_xlsm_resources() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,g As MSForms.Frame
Set g=f.Frame1
Debug.Print f.Controls.Count,g.Controls.Count,f.Frame2.Controls.Count
Debug.Print f.Entry.Parent Is f.Frame2
f.Show 0
Debug.Print f.ActiveControl.Name,g.ActiveControl.Name
End Sub"#,
            ),
        ]),
        fixture::frame_companion(2),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Frames.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(ide.form_instances()?.is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["initialized", "3\t1\t1", "True", "Frame1\tFrame2"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    let native = ide.native_form(f.source).ok_or("native")?.clone();
    ide.form_instance_text(f, 12, vba_values::VbaString::from_utf8("nested😀")?, (8, 0), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaText")?, vba_values::Value::String(vba_values::VbaString::from_utf8("nested😀")?));
    assert_eq!(ide.native_form(f.source), Some(&native));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_tabstrip_objects_and_input_preserve_xlsm_resources() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,s As MSForms.TabStrip
Set s=f.Tabs1
Debug.Print f.Controls.Count,s.Tabs.Count,s.SelectedItem Is s.FirstTab
s.Value=1
Debug.Print s.SelectedItem.Caption
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::tab_companion(0),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Tabs.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(ide.form_instances()?.is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["initialized", "4\t3\tTrue", "Second😀"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    let native = ide.native_form(f.source).ok_or("native")?.clone();
    assert!(ide.form_instance_tab(f, 14, 2, &mut bridge.host(&mut session)).is_err());
    ide.form_instance_tab(f, 14, 0, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(14), "Value")?, vba_values::Value::Long(0));
    assert_eq!(ide.native_form(f.source), Some(&native));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_multipage_objects_and_input_preserve_xlsm_resources() -> R {
    let form = format!(
        "{}\nPrivate Sub Multi1_MouseDown(Index As Long,ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nDebug.Print Index,TypeName(X)\nEnd Sub",
        fixture::SOURCE
    );
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", &form),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,s As MSForms.MultiPage
Set s=f.Multi1
Debug.Print f.Controls.Count,s.Pages.Count,s.SelectedItem Is s.FirstTab
s.Value=1
Debug.Print s.SelectedItem.Caption
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::multipage_companion(0),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Tabs.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(ide.form_instances()?.is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["initialized", "4\t3\tTrue", "Second😀"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    let native = ide.native_form(f.source).ok_or("native")?.clone();
    assert!(ide.form_instance_tab(f, 20, 2, &mut bridge.host(&mut session)).is_err());
    ide.form_instance_tab(f, 20, 0, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(20), "Value")?, vba_values::Value::Long(0));
    assert!(ide.form_instance_text(f, 41, vba_values::VbaString::from_utf8("hidden")?, (6, 0), &mut bridge.host(&mut session)).is_err());
    ide.form_instance_text(f, 40, vba_values::VbaString::from_utf8("selected")?, (8, 0), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(40), "VbaText")?, vba_values::Value::String(vba_values::VbaString::from_utf8("selected")?));
    bridge.begin(&mut session)?;
    let epoch = ide.form_instance_mouse(
        f,
        Some(20),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 0, x: 5., y: 5., index: 0 },
        &mut bridge.host(&mut session),
    )?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_mouse_result(f, epoch)?, Some(true));
    assert_eq!(ide.output().last().map(String::as_str), Some("0\tSingle"));
    assert_eq!(ide.native_form(f.source), Some(&native));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_image_objects_and_input_preserve_xlsm_resources() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,s As MSForms.Image
Set s=f.Image1
Dim p As stdole.IPictureDisp
Set p=s.Picture
Set f.Image2.Picture=p
Debug.Print f.Controls.Count,p.Width,p.Height,p.Type
Debug.Print f.Image2.Picture Is p
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::image_companion(),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Tabs.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(ide.form_instances()?.is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["initialized", "5\t1200\t720\t1", "True"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    let native = ide.native_form(f.source).ok_or("native")?.clone();
    ide.form_instance_event(f, Some(14), None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, None, "VbaActiveControl")?, ide.form_instance_field(f, None, "VbaEntry")?);
    let first = ide.form_instance_picture(f, 14)?.ok_or("picture")?;
    let second = ide.form_instance_picture(f, 15)?.ok_or("picture")?;
    assert!(std::sync::Arc::ptr_eq(&first.rgba, &second.rgba));
    assert_eq!(ide.native_form(f.source), Some(&native));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_refedit_text_preserves_xlsm_resources() -> R {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,r As RefEdit.RefEdit
Set r=f.Entry
Debug.Print r.Value,r.Parent Is f
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::refedit_companion(),
    );
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "RefEdit.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.form_instances()?.is_empty());
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["initialized", "initial\tTrue"]);
    let f = *ide.form_instances()?.first().ok_or("instance")?;
    let native = ide.native_form(f.source).ok_or("native")?.clone();
    assert!(ide.form_instance_drop(f, 12, true, &mut bridge.host(&mut session)).is_err());
    ide.form_instance_text(f, 12, vba_values::VbaString::from_utf8("Sheet1!$B$2:$C$4")?, (0, 0), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaText")?, gridcraft_vba::Value::String(vba_values::VbaString::from_utf8("Sheet1!$B$2:$C$4")?));
    assert_eq!(ide.form_range_selection()?, None);
    assert_eq!(ide.native_form(f.source), Some(&native));
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    let (back, _) = read_xlsx(&saved)?;
    assert_eq!(back.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_dynamic_tabs_and_pages_keep_aliases_trust_and_xlsm_resources() -> R {
    for pages in [false, true] {
        let (control, collection, id) = if pages { ("Multi1", "Pages", 20) } else { ("Tabs1", "Tabs", 14) };
        let source = format!(
            "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,s As Object,t As Object\nSet s=f.{control}\nSet t=s.{collection}(0)\nt.Index=1\ns.{collection}.Add Name:=\"Runtime\",Caption:=\"Added\",Index:=0\ns.{collection}.Remove \"BlockedTab\"\nDebug.Print s.Value,s.{collection}(2) Is t,s.{collection}.Count\nf.Show 0\nEnd Sub"
        );
        let original = form_project(
            support::project(&[("NativeForm", "BaseClass", fixture::SOURCE), ("Launch", "Module", &source)]),
            if pages { fixture::multipage_companion(0) } else { fixture::tab_companion(0) },
        );
        let mut wb = Workbook::new();
        wb.vba_project = Some(original.clone().into());
        let mut session = Session::new();
        session.add_document(DocState::from_saved(wb, None, "DynamicForms.xlsm".into(), Trust::FromFile));
        let mut bridge = Bridge::new(&session)?;
        let mut ide = bridge.import_project(&session)?;
        ide.entry = "Main".into();
        assert!(ide.form_instances()?.is_empty());
        assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
        bridge.consent(true);
        bridge.dispatch(&mut session, &mut ide, Action::Run)?;
        settle(&mut session, &mut bridge, &mut ide)?;
        assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
        assert_eq!(ide.output(), vec!["initialized", "2\tTrue\t3"]);
        let f = *ide.form_instances()?.first().ok_or("form")?;
        let native = ide.native_form(f.source).ok_or("native")?.clone();
        ide.form_instance_tab(f, id, 0, &mut bridge.host(&mut session))?;
        settle(&mut session, &mut bridge, &mut ide)?;
        assert_eq!(ide.form_instance_field(f, Some(id), "Value")?, vba_values::Value::Long(0));
        if pages {
            assert!(ide.form_instance_text(f, 40, vba_values::VbaString::from_utf8("hidden")?, (0, 0), &mut bridge.host(&mut session)).is_err());
        }
        assert_eq!(ide.native_form(f.source), Some(&native));
        let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
        assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
        bridge.consent(false);
        assert!(bridge.begin(&mut session).is_err());
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_excel_cell_and_row_bindings_use_the_existing_range_bridge_and_keep_package_bytes() -> R {
    let source = r#"Attribute VB_Name = "Launch"
Sub Main()
Range("A1").Value="bound"
Range("B1").Value="alpha"
Range("C1").Value="A"
Range("B2").Value="beta"
Range("C2").Value="B"
Range("D1").Value="B"
Dim f As New NativeForm,l As MSForms.ListBox
Set l=f.List1
f.Entry.ControlSource="A1"
l.RowSource="B1:C2"
l.ControlSource="D1"
Debug.Print f.Entry.Value,l.ListCount,l.ListIndex,l.Value
f.Entry.Value="written"
Debug.Print Range("A1").Value
Range("A1").Value="outside"
Debug.Print f.Entry.Value
Range("C2").Value="B2"
Debug.Print l.Value,l.List(1,1),Range("D1").Value
Range("D1").Value="A"
Debug.Print l.ListIndex,l.Value,Range("D1").Value
On Error Resume Next
l.AddItem "blocked"
Debug.Print Err.Number,l.ListCount
Err.Clear
l.List(0,0)="blocked"
Debug.Print Err.Number,l.List(0,0)
Err.Clear
f.Entry.ControlSource="A1:A2"
Debug.Print Err.Number,f.Entry.ControlSource,f.Entry.Value
Err.Clear
l.RowSource="A1:K2"
Debug.Print Err.Number,l.RowSource,l.ListCount
Err.Clear
l.RowSource=""
Debug.Print l.ListCount,IsNull(l.Value)
f.Entry.ControlSource=""
f.Entry.Value="detached"
Debug.Print Range("A1").Value,f.Entry.Value
f.Show 0
End Sub"#;
    let original =
        form_project(support::project(&[("NativeForm", "BaseClass", fixture::SOURCE), ("Launch", "Module", source)]), fixture::list_companion());
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Bindings.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.form_instances()?.is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(
        ide.output(),
        vec![
            "initialized",
            "bound\t2\t1\tB",
            "written",
            "outside",
            "B2\tB2\tB2",
            "0\tA\tA",
            "70\t2",
            "70\talpha",
            "380\tA1\toutside",
            "7\tB1:C2\t2",
            "0\tTrue",
            "outside\tdetached"
        ]
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_bound_text_refreshes_from_engine_edits_and_keeps_stopped_input_exactly_once() -> R {
    let source = fixture::SOURCE.to_owned()
        + r#"
Private Sub Entry_Change()
If Not Me.Visible Then Exit Sub
Debug.Print "changed",Entry.Value,Range("A1").Value
If Entry.Value="physical" Then
Stop
Debug.Print "resumed"
End If
End Sub
"#;
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &source),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nSub Main()\nRange(\"A1\").Value=\"seed\"\nNativeForm.Entry.ControlSource=\"A1\"\nNativeForm.Show 0\nEnd Sub",
        ),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "BoundInput.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    let f = *ide.form_instances()?.first().ok_or("form")?;
    let native = ide.native_form(f.source).ok_or("native")?.clone();
    assert_eq!(ide.form_instance_field(f, Some(12), "ControlSource")?, vba_values::Value::String(vba_values::VbaString::from_utf8("A1")?));
    session.execute("cell.set", serde_json::json!({"cell":"A1","input":"external"}))?;
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaText")?, vba_values::Value::String(vba_values::VbaString::from_utf8("seed")?));
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output(), vec!["initialized", "changed\texternal\texternal"]);
    ide.form_instance_text(f, 12, vba_values::VbaString::from_utf8("physical")?, (8, 0), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break, "{:?}", ide.diagnostics);
    assert_eq!(ide.output().last().map(String::as_str), Some("changed\tphysical\tphysical"));
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output(), vec!["initialized", "changed\texternal\texternal", "changed\tphysical\tphysical", "resumed"]);
    let output = ide.output().to_vec();
    for _ in 0..20 {
        bridge.tick(&mut session, &mut ide)?;
    }
    assert_eq!(ide.output(), output);
    bridge.consent(false);
    session.execute("cell.set", serde_json::json!({"cell":"A1","input":"deferred"}))?;
    bridge.tick(&mut session, &mut ide)?;
    assert_eq!(ide.output(), output);
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaText")?, vba_values::Value::String(vba_values::VbaString::from_utf8("physical")?));
    bridge.consent(true);
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output().last().map(String::as_str), Some("changed\tdeferred\tdeferred"));
    assert_eq!(ide.native_form(f.source), Some(&native));
    assert_eq!(read_xlsx(&write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?)?.0.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_bound_row_refresh_stays_read_only_inside_change_and_large_lists_reuse_epochs() -> R {
    let form = fixture::SOURCE.to_owned()
        + r#"
Private Sub List1_Change()
On Error Resume Next
List1.AddItem "illegal"
Debug.Print "guard",Err.Number,List1.ListCount
Err.Clear
List1.Clear
Debug.Print "clear",Err.Number,List1.ListCount
End Sub
"#;
    let source = r#"Attribute VB_Name = "Launch"
Sub Main()
Dim data(0 To 127,0 To 1) As Variant,i As Long,n As Long
For i=0 To 127
data(i,0)="row " & i
data(i,1)="key " & i
Next i
Range("B1:C128").Value=data
NativeForm.List1.RowSource="B1:C128"
NativeForm.List1.ListIndex=127
For i=0 To 39
n=n+NativeForm.List1.ListCount
If NativeForm.List1.List(127,1)<>"key 127" Then Err.Raise 5
Next i
Debug.Print n
Range("C128").Value="updated"
Debug.Print NativeForm.List1.Value,NativeForm.List1.ListCount
NativeForm.Show 0
End Sub"#;
    let original = form_project(support::project(&[("NativeForm", "BaseClass", &form), ("Launch", "Module", source)]), fixture::list_companion());
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "BoundRows.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["initialized", "guard\t70\t128", "clear\t70\t128", "5120", "guard\t70\t128", "clear\t70\t128", "updated\t128"]);
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_bound_protected_cell_failures_preserve_control_state_and_sources() -> R {
    let source = r#"Attribute VB_Name = "Launch"
Sub Main()
Range("A1").Value="seed"
Range("A2").Value=10
Dim spin As MSForms.SpinButton
Set spin=NativeForm.Controls.Add("Forms.SpinButton.1","BoundSpin")
spin.ControlSource="A2"
Range("D1").Value="beta"
NativeForm.Entry.ControlSource="A1"
NativeForm.List1.BoundColumn=1
NativeForm.List1.AddItem "alpha"
NativeForm.List1.AddItem "beta"
NativeForm.List1.ControlSource="D1"
ActiveSheet.Protect
On Error Resume Next
spin.Min=20
Debug.Print Err.Number,spin.Min,spin.Value
Err.Clear
NativeForm.Entry.VbaSetTextSelection "blocked",2,1
Debug.Print Err.Number,NativeForm.Entry.Value,Range("A1").Value,NativeForm.Entry.VbaDirty
Err.Clear
NativeForm.List1.ListIndex=0
Debug.Print Err.Number,NativeForm.List1.ListIndex,NativeForm.List1.Value
Err.Clear
NativeForm.List1.RemoveItem 1
Debug.Print Err.Number,NativeForm.List1.ListCount,NativeForm.List1.ListIndex,NativeForm.List1.Value
Err.Clear
NativeForm.List1.Clear
Debug.Print Err.Number,NativeForm.List1.ListCount,NativeForm.List1.Value
Err.Clear
NativeForm.Entry.ControlSource="A1:A2"
Debug.Print Err.Number,NativeForm.Entry.ControlSource,NativeForm.Entry.Value
NativeForm.Show 0
End Sub"#;
    let original =
        form_project(support::project(&[("NativeForm", "BaseClass", fixture::SOURCE), ("Launch", "Module", source)]), fixture::list_companion());
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "ProtectedBindings.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(
        ide.output(),
        vec!["initialized", "1004\t0\t10", "1004\tseed\tseed\tFalse", "1004\t1\tbeta", "1004\t2\t1\tbeta", "1004\t2\tbeta", "380\tA1\tseed"]
    );
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_persisted_bindings_attach_before_initialize_and_keep_xlsm_resources_exact() -> R {
    let form = fixture::SOURCE
        .replace("Debug.Print \"initialized\"", "Debug.Print Entry.Value,List1.ListCount,List1.Value,Entry.ControlSource,List1.RowSource");
    let source = "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nEnd Sub";
    let original = form_project(support::project(&[("NativeForm", "BaseClass", &form), ("Launch", "Module", source)]), fixture::binding_companion());
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let package = write_xlsx_as(&wb, XlsxKind::MacroWorkbook)?;
    let (wb, _) = read_xlsx(&package)?;
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "NativeSources.xlsm".into(), Trust::FromFile));
    for (cell, input) in [("A1", "from cell"), ("B1", "first"), ("C1", "A"), ("B2", "second"), ("C2", "B"), ("D1", "B")] {
        session.execute("cell.set", serde_json::json!({"cell":cell,"input":input}))?;
    }
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.form_instances()?.is_empty());
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["from cell\t2\tB\t$A$1\t$B$1:$C$2"]);
    bridge.dispatch(&mut session, &mut ide, Action::Reset)?;
    assert!(ide.form_instances()?.is_empty());
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_spilled_cell_binding_refreshes_after_recalculation_without_content_revision() -> R {
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", fixture::SOURCE),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Entry.ControlSource=\"B1\"\nNativeForm.Show 0\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "CalculatedBinding.xlsm".into(), Trust::FromFile));
    session.doc_mut()?.calc.fixed_now = Some(45000.);
    session.execute("cell.set", serde_json::json!({"cell":"A1","input":"=SEQUENCE(1,2,NOW(),1)"}))?;
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    bridge.consent(true);
    ide.entry = "Main".into();
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let f = *ide.form_instances()?.first().ok_or("form")?;
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaText")?, vba_values::Value::String(vba_values::VbaString::from_utf8("45001")?));
    let revision = session.doc()?.revision;
    session.doc_mut()?.calc.fixed_now = Some(45010.);
    session.execute("formulas.calculateNow", serde_json::json!({}))?;
    assert_eq!(session.doc()?.revision, revision);
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaText")?, vba_values::Value::String(vba_values::VbaString::from_utf8("45011")?));
    assert_eq!(read_xlsx(&write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?)?.0.vba_project.as_deref(), Some(&original));
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_form_arrays_keep_callback_receivers_trust_and_xlsm_resources() -> R {
    let launch = "Attribute VB_Name = \"Launch\"\nSub Main()\nDim forms(0 To 1) As New NativeForm,a As NativeForm\nDebug.Print UserForms.Count\nforms(0).Show 0\nforms(0).Entry.Text=\"first\"\nforms(1).Show 0\nforms(1).Entry.Text=\"second\"\nSet a=forms(0)\nSet forms(0)=Nothing\nforms(0).Show 0\nDebug.Print a.Entry.Text,forms(1).Entry.Text,forms(0).Entry.Text\nEnd Sub";
    let form = fixture::SOURCE.replace("Debug.Print Entry.Text", "Range(\"A1\")=Entry.Text\r\nDebug.Print Entry.Text");
    let original = fixture::add_designer(support::project(&[("NativeForm", "BaseClass", &form), ("Launch", "Module", launch)]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let package = write_xlsx_as(&wb, XlsxKind::MacroWorkbook)?;
    let mut session = Session::new();
    session.add_document(DocState::from_saved(read_xlsx(&package)?.0, None, "ArrayForms.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(ide.form_instances()?.is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    assert_eq!(ide.output(), vec!["0", "initialized", "initialized", "initialized", "first\tsecond\tinitial"]);
    let instances = ide.form_instances()?;
    assert_eq!(instances.len(), 3);
    let expected = vba_values::Value::String(vba_values::VbaString::from_utf8("second")?);
    let target = instances
        .into_iter()
        .find(|instance| ide.form_instance_field(*instance, Some(12), "VbaText").is_ok_and(|value| value == expected))
        .ok_or("second instance")?;
    ide.form_instance_event(target, Some(13), None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.output().last().map(String::as_str), Some("second"));
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Text("second".into())
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_keyboard_waits_for_debugger_and_preserves_workbook_authority_and_resources() -> R {
    let form = format!(
        "{}\nPrivate Sub Entry_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)\nDebug.Print KeyANSI.Value\nStop\nRange(\"A1\").Value=KeyANSI.Value\nKeyANSI.Value=0\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nNativeForm.Entry.SetFocus\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Keys.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let f = *ide.form_instances()?.first().ok_or("form")?;
    bridge.begin(&mut session)?;
    let epoch =
        ide.form_instance_key(f, Some(12), vba_ide::FormKeyInput::Text(vba_values::VbaString::from_utf8("a")?), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.form_instance_key_result(f, epoch)?.is_none());
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()), gridcraft_engine::core::Value::Empty);
    assert_eq!(ide.output(), vec!["initialized", "97"]);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_key_result(f, epoch)?.ok_or("receipt")?.code, 0);
    assert_eq!(ide.output(), vec!["initialized", "97"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Number(97.0)
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_root_keys_distinguish_tabstop_from_disabled_controls_and_keep_xlsm_resources() -> R {
    for disabled in [false, true] {
        let form = format!(
            "{}\nPrivate Sub UserForm_KeyDown(ByVal k As MSForms.ReturnInteger,ByVal s As Integer)\nStop\nRange(\"A1\").Value=\"root\"\nDebug.Print \"root\"\nk.Value=0\nEnd Sub",
            fixture::SOURCE
        );
        let launch = format!(
            "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Entry.TabStop=False\nNativeForm.Apply.TabStop=False\n{}\nNativeForm.Show 0\nEnd Sub",
            if disabled { "Dim c As MSForms.Control\nFor Each c In NativeForm.Controls\nc.Enabled=False\nNext c" } else { "" }
        );
        let original = fixture::add_designer(support::project(&[("NativeForm", "BaseClass", &form), ("Launch", "Module", &launch)]));
        let mut wb = Workbook::new();
        wb.vba_project = Some(original.clone().into());
        let mut session = Session::new();
        session.add_document(DocState::from_saved(wb, None, "RootKeys.xlsm".into(), Trust::FromFile));
        let mut bridge = Bridge::new(&session)?;
        let mut ide = bridge.import_project(&session)?;
        ide.entry = "Main".into();
        bridge.consent(true);
        bridge.dispatch(&mut session, &mut ide, Action::Run)?;
        settle(&mut session, &mut bridge, &mut ide)?;
        let f = *ide.form_instances()?.first().ok_or("form")?;
        bridge.begin(&mut session)?;
        let epoch = ide.form_instance_key(
            f,
            None,
            vba_ide::FormKeyInput::Code { event: vba_ide::FormKeyEvent::Down, code: 65, shift: 0 },
            &mut bridge.host(&mut session),
        )?;
        settle(&mut session, &mut bridge, &mut ide)?;
        assert_eq!(
            session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
            gridcraft_engine::core::Value::Empty
        );
        if disabled {
            assert_eq!(ide.mode(), Mode::Break);
            assert!(ide.form_instance_key_result(f, epoch)?.is_none());
            bridge.dispatch(&mut session, &mut ide, Action::Run)?;
            settle(&mut session, &mut bridge, &mut ide)?;
        }
        assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
        assert_eq!(ide.form_instance_key_result(f, epoch)?.ok_or("receipt")?.code, 0);
        assert_eq!(ide.output(), if disabled { vec!["initialized", "root"] } else { vec!["initialized"] });
        if disabled {
            assert_eq!(
                session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
                gridcraft_engine::core::Value::Text("root".into())
            );
        }
        assert_eq!(read_xlsx(&write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?)?.0.vba_project.as_deref(), Some(&original));
        bridge.consent(false);
        assert!(bridge.begin(&mut session).is_err());
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_accelerator_click_waits_for_debugger_and_preserves_workbook_resources() -> R {
    let form = fixture::SOURCE.replace("Debug.Print Entry.Text", "Stop\r\nRange(\"A1\").Value=\"accelerated\"\r\nDebug.Print Entry.Text");
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nNativeForm.Apply.Accelerator=\"A\"\nNativeForm.Entry.SetFocus\nEnd Sub",
        ),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Keys.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let f = *ide.form_instances()?.first().ok_or("form")?;
    bridge.begin(&mut session)?;
    let epoch = ide.form_instance_key(
        f,
        Some(12),
        vba_ide::FormKeyInput::Code { event: vba_ide::FormKeyEvent::Down, code: 65, shift: 4 },
        &mut bridge.host(&mut session),
    )?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.form_instance_key_result(f, epoch)?.is_none());
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()), gridcraft_engine::core::Value::Empty);
    assert_eq!(ide.output(), vec!["initialized"]);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_key_result(f, epoch)?.ok_or("receipt")?.code, 0);
    assert_eq!(ide.output(), vec!["initialized", "initial"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Text("accelerated".into())
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_mouse_callback_waits_for_debugger_and_preserves_workbook_resources() -> R {
    let form = format!(
        "{}\nPrivate Sub Apply_MouseDown(ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nStop\nRange(\"A1\").Value=\"mouse\"\nDebug.Print \"down\"\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nNativeForm.Entry.SetFocus\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Keys.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let f = *ide.form_instances()?.first().ok_or("form")?;
    bridge.begin(&mut session)?;
    let epoch = ide.form_instance_mouse(
        f,
        Some(13),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut bridge.host(&mut session),
    )?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.form_instance_mouse_result(f, epoch)?.is_none());
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()), gridcraft_engine::core::Value::Empty);
    assert_eq!(ide.output(), vec!["initialized"]);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_mouse_result(f, epoch)?, Some(true));
    assert_eq!(ide.output(), vec!["initialized", "down"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Text("mouse".into())
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_text_pointer_selection_after_mouse_preserves_workbook_resources() -> R {
    let form = format!(
        "{}\nPrivate Sub Entry_MouseDown(ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nStop\nRange(\"A1\").Value=\"mouse\"\nDebug.Print \"down\"\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nNativeForm.Entry.SetFocus\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Keys.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let f = *ide.form_instances()?.first().ok_or("form")?;
    bridge.begin(&mut session)?;
    let epoch = ide.form_instance_mouse(
        f,
        Some(12),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut bridge.host(&mut session),
    )?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.form_instance_mouse_result(f, epoch)?.is_none());
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()), gridcraft_engine::core::Value::Empty);
    assert_eq!(ide.output(), vec!["initialized"]);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_mouse_result(f, epoch)?, Some(true));
    bridge.begin(&mut session)?;
    ide.form_instance_text(f, 12, vba_values::VbaString::from_utf8("initial")?, (1, 3), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaSelStart")?, vba_values::Value::Long(1));
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaSelLength")?, vba_values::Value::Long(3));
    assert_eq!(ide.form_instance_field(f, Some(12), "AutoWordSelect")?, vba_values::Value::Boolean(false));
    assert_eq!(ide.output(), vec!["initialized", "down"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Text("mouse".into())
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_modal_pointer_callback_preserves_workbook_resources() -> R {
    let form = format!(
        "{}\nPrivate Sub Entry_MouseDown(ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nStop\nRange(\"A1\").Value=\"mouse\"\nDebug.Print \"down\"\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nNativeForm.Entry.SetFocus\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Keys.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let f = *ide.form_instances()?.first().ok_or("form")?;
    bridge.begin(&mut session)?;
    let epoch = ide.form_instance_mouse(
        f,
        Some(12),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut bridge.host(&mut session),
    )?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.form_instance_has_modal_window(f));
    assert!(!ide.form_instance_can_deliver(f));
    assert!(ide.form_instance_mouse_result(f, epoch)?.is_none());
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()), gridcraft_engine::core::Value::Empty);
    assert_eq!(ide.output(), vec!["initialized"]);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_mouse_result(f, epoch)?, Some(true));
    assert!(ide.form_instance_has_modal_window(f));
    bridge.begin(&mut session)?;
    ide.form_instance_text(f, 12, vba_values::VbaString::from_utf8("initial")?, (1, 3), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaSelStart")?, vba_values::Value::Long(1));
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaSelLength")?, vba_values::Value::Long(3));
    assert_eq!(ide.form_instance_field(f, Some(12), "AutoWordSelect")?, vba_values::Value::Boolean(false));
    assert_eq!(ide.output(), vec!["initialized", "down"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Text("mouse".into())
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_nested_modal_pointer_callback_preserves_workbook_resources() -> R {
    let form = format!(
        "{}\nPrivate Sub Entry_MouseDown(ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nDim child As NativeForm\nIf Me.Caption=\"Parent\" Then\nSet child=New NativeForm\nchild.Caption=\"Child\"\nchild.Show\nElse\nStop\nRange(\"A1\").Value=\"mouse\"\nDebug.Print \"down\"\nEnd If\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Caption=\"Parent\"\nNativeForm.Show\nNativeForm.Entry.SetFocus\nEnd Sub",
        ),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Keys.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let f = *ide.form_instances()?.first().ok_or("form")?;
    bridge.begin(&mut session)?;
    let epoch = ide.form_instance_mouse(
        f,
        Some(12),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut bridge.host(&mut session),
    )?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Waiting);
    let parent_epoch = epoch;
    let child = *ide.form_instances()?.last().ok_or("child")?;
    assert_ne!(child, f);
    assert_eq!(ide.form_instance_modal_rank(f), Some(0));
    assert_eq!(ide.form_instance_modal_rank(child), Some(1));
    assert!(!ide.form_instance_can_deliver(f));
    bridge.begin(&mut session)?;
    let epoch = ide.form_instance_mouse(
        child,
        Some(12),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut bridge.host(&mut session),
    )?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.form_instance_has_modal_window(f));
    assert!(!ide.form_instance_can_deliver(f));
    assert!(ide.form_instance_mouse_result(child, epoch)?.is_none());
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()), gridcraft_engine::core::Value::Empty);
    assert_eq!(ide.output(), vec!["initialized", "initialized"]);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_mouse_result(child, epoch)?, Some(true));
    assert!(ide.form_instance_mouse_result(f, parent_epoch)?.is_none());
    bridge.begin(&mut session)?;
    ide.form_instance_event(child, None, None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_mouse_result(f, parent_epoch)?, Some(true));
    assert!(ide.form_instance_has_modal_window(f));
    bridge.begin(&mut session)?;
    ide.form_instance_text(f, 12, vba_values::VbaString::from_utf8("initial")?, (1, 3), &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaSelStart")?, vba_values::Value::Long(1));
    assert_eq!(ide.form_instance_field(f, Some(12), "VbaSelLength")?, vba_values::Value::Long(3));
    assert_eq!(ide.form_instance_field(f, Some(12), "AutoWordSelect")?, vba_values::Value::Boolean(false));
    assert_eq!(ide.output(), vec!["initialized", "initialized", "down"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Text("mouse".into())
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_window_deactivation_preserves_trust_stop_and_workbook_resources() -> R {
    let form = format!(
        "{}\nPrivate Sub UserForm_Activate()\nDebug.Print Me.Caption & \" active\"\nEnd Sub\nPrivate Sub UserForm_Deactivate()\nDebug.Print Me.Caption & \" inactive\"\nIf Me.Caption=\"B\" Then\nStop\nRange(\"A1\").Value=\"deactivated\"\nEnd If\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nPublic a As NativeForm,b As NativeForm\nSub Main()\nSet a=New NativeForm\nSet b=New NativeForm\na.Caption=\"A\"\nb.Caption=\"B\"\na.Show 0\nb.Show 0\nEnd Sub",
        ),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Windows.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let active = ide.form_active_window().ok_or("active window")?;
    assert_eq!(ide.output(), vec!["initialized", "initialized", "A active", "A inactive", "B active"]);
    bridge.begin(&mut session)?;
    ide.form_instance_window_focus(active, false, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert!(ide.form_active_window().is_none());
    assert!(!ide.form_instance_can_deliver(active));
    assert_eq!(session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()), gridcraft_engine::core::Value::Empty);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Text("deactivated".into())
    );
    assert_eq!(ide.output().iter().filter(|line| *line == "B inactive").count(), 1);
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_colours_keep_canonical_bits_trusted_continuations_and_original_workbook_resources() -> R {
    let form = "Attribute VB_Name = \"NativeForm\"\nAttribute VB_PredeclaredId = True\nPrivate Sub UserForm_Initialize()\nMe.BackColor=RGB(230,240,250)\nEntry.ForeColor=RGB(80,30,20)\nEnd Sub\nPrivate Sub Apply_Click()\nEntry.BackColor=RGB(250,230,200)\nRange(\"A1\").Value=Entry.BackColor\nStop\nEntry.ForeColor=RGB(1,2,3)\nRange(\"B1\").Value=Entry.ForeColor\nMe.Hide\nEnd Sub";
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Colours.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let owner = ide.form_active_window().ok_or("form")?;
    assert_eq!(ide.form_instance_field(owner, None, "BackColor")?, vba_values::Value::Long(16445670));
    bridge.begin(&mut session)?;
    ide.form_instance_event(owner, Some(13), None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(ide.form_instance_field(owner, Some(12), "BackColor")?, vba_values::Value::Long(13166330));
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Number(13166330.)
    );
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed);
    assert_eq!(ide.form_instance_field(owner, Some(12), "ForeColor")?, vba_values::Value::Long(197121));
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::new(0, 1)),
        gridcraft_engine::core::Value::Number(197121.)
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_fonts_keep_typed_aliases_trusted_continuations_and_original_workbook_resources() -> R {
    let form = "Attribute VB_Name = \"NativeForm\"\nAttribute VB_PredeclaredId = True\nPrivate Sub UserForm_Initialize()\nEntry.Font.Name=\"Authored face\"\nEntry.Font.Size=12.25\nEnd Sub\nPrivate Sub Apply_Click()\nDim font As MSForms.NewFont\nSet font=Entry.Font\nRange(\"A1\").Value=font.Size\nStop\nfont.Bold=True\nfont.Size=14.5\nRange(\"B1\").Value=font.Weight\nMe.Hide\nEnd Sub";
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Fonts.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let owner = ide.form_active_window().ok_or("form")?;
    assert_eq!(ide.form_instance_font(owner, Some(12))?.size, 122500);
    bridge.begin(&mut session)?;
    ide.form_instance_event(owner, Some(13), None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(ide.form_instance_font(owner, Some(12))?.name, "Authored face");
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Number(12.25)
    );
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed);
    assert_eq!((ide.form_instance_font(owner, Some(12))?.size, ide.form_instance_font(owner, Some(12))?.weight), (145000, 700));
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::new(0, 1)),
        gridcraft_engine::core::Value::Number(700.)
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_font_assignment_keeps_shared_identity_trusted_continuations_and_original_workbook_resources() -> R {
    let form = "Attribute VB_Name = \"NativeForm\"\nAttribute VB_PredeclaredId = True\nPrivate Sub UserForm_Initialize()\nEntry.Font.Name=\"Authored face\"\nEntry.Font.Size=12.25\nEnd Sub\nPrivate Sub Apply_Click()\nDim font As MSForms.NewFont\nSet font=Entry.Font\nSet Heading.Font=font\nSet Me.Font=font\nDebug.Print Heading.Font Is font,Me.Font Is font\nRange(\"A1\").Value=font.Size\nStop\nfont.Bold=True\nfont.Size=14.5\nRange(\"B1\").Value=font.Weight\nMe.Hide\nEnd Sub";
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Fonts.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let owner = ide.form_active_window().ok_or("form")?;
    assert_eq!(ide.form_instance_font(owner, Some(12))?.size, 122500);
    bridge.begin(&mut session)?;
    ide.form_instance_event(owner, Some(13), None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(ide.form_instance_font(owner, Some(12))?.name, "Authored face");
    assert_eq!(ide.output(), vec!["True\tTrue"]);
    assert_eq!(ide.form_instance_font(owner, Some(11))?.size, 122500);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Number(12.25)
    );
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed);
    assert_eq!((ide.form_instance_font(owner, Some(12))?.size, ide.form_instance_font(owner, Some(12))?.weight), (145000, 700));
    assert_eq!(ide.form_instance_font(owner, Some(11))?.size, 145000);
    assert_eq!(ide.form_instance_font(owner, None)?.weight, 700);
    assert_eq!(ide.output(), vec!["True\tTrue"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::new(0, 1)),
        gridcraft_engine::core::Value::Number(700.)
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_page_appearance_uses_multipage_without_non_office_members_trusted_continuations_and_original_workbook_resources() -> R {
    let form = "Attribute VB_Name = \"NativeForm\"\nAttribute VB_PredeclaredId = True\nPrivate Sub UserForm_Initialize()\nEntry.Font.Name=\"Authored face\"\nEntry.Font.Size=12.25\nEnd Sub\nPrivate Sub Apply_Click()\nDim font As MSForms.NewFont\nSet font=Entry.Font\nSet Heading.Font=font\nSet Me.Font=font\nDebug.Print Heading.Font Is font,Me.Font Is font\nDim m As Object,p As Object,c As Object\nSet m=Me.Controls.Add(\"Forms.MultiPage.1\")\nSet p=m.Pages.Add()\nm.FontSize=17.25\nSet c=p.Controls.Add(\"Forms.TextBox.1\")\nDebug.Print c.FontSize,c.Parent Is p,p.Parent Is m\nOn Error Resume Next\nDebug.Print p.Font\nDebug.Print Err.Number\nErr.Clear\np.BackColor=1\nDebug.Print Err.Number\nOn Error GoTo 0\nRange(\"A1\").Value=font.Size\nStop\nfont.Bold=True\nfont.Size=14.5\nRange(\"B1\").Value=font.Weight\nMe.Hide\nEnd Sub";
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Fonts.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let owner = ide.form_active_window().ok_or("form")?;
    assert_eq!(ide.form_instance_font(owner, Some(12))?.size, 122500);
    bridge.begin(&mut session)?;
    ide.form_instance_event(owner, Some(13), None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(ide.form_instance_font(owner, Some(12))?.name, "Authored face");
    assert_eq!(ide.output(), vec!["True\tTrue", "17.25\tTrue\tTrue", "438", "438"]);
    assert_eq!(ide.form_instance_font(owner, Some(11))?.size, 122500);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Number(12.25)
    );
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed);
    assert_eq!((ide.form_instance_font(owner, Some(12))?.size, ide.form_instance_font(owner, Some(12))?.weight), (145000, 700));
    assert_eq!(ide.form_instance_font(owner, Some(11))?.size, 145000);
    assert_eq!(ide.form_instance_font(owner, None)?.weight, 700);
    assert_eq!(ide.output(), vec!["True\tTrue", "17.25\tTrue\tTrue", "438", "438"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::new(0, 1)),
        gridcraft_engine::core::Value::Number(700.)
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_appearance_and_font_aliases_preserve_typed_continuations_trusted_continuations_and_original_workbook_resources() -> R {
    let form = "Attribute VB_Name = \"NativeForm\"\nAttribute VB_PredeclaredId = True\nPrivate Sub UserForm_Initialize()\nEntry.Font.Name=\"Authored face\"\nEntry.Font.Size=12.25\nEnd Sub\nPrivate Sub Apply_Click()\nDim font As MSForms.NewFont\nSet font=Entry.Font\nSet Heading.Font=font\nSet Me.Font=font\nDebug.Print Heading.Font Is font,Me.Font Is font\nEntry.BackStyle=0\nEntry.BorderColor=RGB(1,2,3)\nEntry.BorderStyle=1\nEntry.FontSize=12.25\nDebug.Print VarType(Entry.BorderStyle),Entry.SpecialEffect\nRange(\"A1\").Value=font.Size\nStop\nfont.Bold=True\nfont.Size=14.5\nEntry.SpecialEffect=3\nRange(\"B1\").Value=font.Weight\nMe.Hide\nEnd Sub";
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    let mut wb = Workbook::new();
    wb.vba_project = Some(original.clone().into());
    let mut session = Session::new();
    session.add_document(DocState::from_saved(wb, None, "Fonts.xlsm".into(), Trust::FromFile));
    let mut bridge = Bridge::new(&session)?;
    let mut ide = bridge.import_project(&session)?;
    ide.entry = "Main".into();
    assert!(ide.output().is_empty());
    assert!(bridge.dispatch(&mut session, &mut ide, Action::Run).is_err());
    bridge.consent(true);
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    let owner = ide.form_active_window().ok_or("form")?;
    assert_eq!(ide.form_instance_font(owner, Some(12))?.size, 122500);
    bridge.begin(&mut session)?;
    ide.form_instance_event(owner, Some(13), None, &mut bridge.host(&mut session))?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Break);
    assert_eq!(ide.form_instance_font(owner, Some(12))?.name, "Authored face");
    assert_eq!(ide.output(), vec!["True\tTrue", "3\t0"]);
    assert_eq!(ide.form_instance_font(owner, Some(11))?.size, 122500);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::default()),
        gridcraft_engine::core::Value::Number(12.25)
    );
    bridge.dispatch(&mut session, &mut ide, Action::Run)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    bridge.tick(&mut session, &mut ide)?;
    settle(&mut session, &mut bridge, &mut ide)?;
    assert_eq!(ide.mode(), Mode::Completed);
    assert_eq!((ide.form_instance_font(owner, Some(12))?.size, ide.form_instance_font(owner, Some(12))?.weight), (145000, 700));
    assert_eq!(ide.form_instance_font(owner, Some(11))?.size, 145000);
    assert_eq!(ide.form_instance_font(owner, None)?.weight, 700);
    assert_eq!(ide.form_instance_field(owner, Some(12), "BorderStyle")?, vba_values::Value::Long(0));
    assert_eq!(ide.form_instance_field(owner, Some(12), "SpecialEffect")?, vba_values::Value::Long(3));
    assert_eq!(ide.form_instance_field(owner, Some(12), "BackStyle")?, vba_values::Value::Long(0));
    assert_eq!(ide.output(), vec!["True\tTrue", "3\t0"]);
    assert_eq!(
        session.doc()?.wb.sheet(0).ok_or("sheet")?.value(gridcraft_engine::core::CellRef::new(0, 1)),
        gridcraft_engine::core::Value::Number(700.)
    );
    let saved = write_xlsx_as(session.doc()?.wb.as_ref(), XlsxKind::MacroWorkbook)?;
    assert_eq!(read_xlsx(&saved)?.0.vba_project.as_deref(), Some(&original));
    bridge.consent(false);
    assert!(bridge.begin(&mut session).is_err());
    Ok(())
}
