//! Authored native form import/attachment evidence; no Office or downloaded macro execution.
#[path = "support/native_forms.rs"]
mod fixture;
mod support;
use gridcraft_engine::{DocState, Session, Trust, model::Workbook};
use gridcraft_vba::{Action, Bridge, Ide, Mode};
use gridcraft_xlsx::{XlsxKind, read_xlsx, write_xlsx_as};
type R = Result<(), Box<dyn std::error::Error>>;
fn settle(s: &mut Session, b: &mut Bridge, ide: &mut Ide) -> R {
    for _ in 0..500 {
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
    for name in ["f", "o"] {
        let mut bytes = Vec::new();
        resource.open_stream(format!("/{name}")).unwrap().read_to_end(&mut bytes).unwrap();
        file.create_stream(format!("/NativeForm/{name}")).unwrap().write_all(&bytes).unwrap();
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
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,l As MSForms.ListBox\nSet l=f.List1\nl.AddItem \"alpha\"\nl.AddItem \"beta\"\nl.List(0,1)=\"A\"\nl.List(1,1)=\"B\"\nl.ListIndex=0\nf.Show 0\nDebug.Print l.ListCount\nDebug.Print l.Value\nEnd Sub",
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
