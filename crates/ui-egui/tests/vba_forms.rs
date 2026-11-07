#![cfg(feature = "vba")]
use gridcraft_engine::{
    Session,
    core::{CellRef, Value},
};
use gridcraft_ui_egui::{Services, SheetApp};
use gridcraft_vba::vba_ide::{
    form_archive,
    forms::{Bounds, Control, ControlKind, Form},
};
use serde_json::json;
use std::fs;
use vba_ui_egui::{Command, FormCommand};

type R = Result<(), Box<dyn std::error::Error>>;

#[test]
fn imported_tvform_runs_textbox_and_button_callbacks_on_the_workbook_host() -> R {
    let source = "Private Sub TextBox1_Change()\nRange(\"A1\").Value2 = Len(TextBox1.Text)\nEnd Sub\nPrivate Sub CommandButton1_Click()\nRange(\"A2\").Value2 = 42\nMe.Hide\nEnd Sub\n";
    let form = Form {
        name: "M5Form".into(),
        caption: "M5 Form".into(),
        width: 300.0,
        height: 160.0,
        controls: vec![
            Control {
                id: 1,
                name: "TextBox1".into(),
                kind: ControlKind::TextBox,
                bounds: Bounds { left: 12.0, top: 12.0, width: 120.0, height: 24.0 },
                caption: String::new(),
                text: String::new(),
                enabled: true,
                visible: true,
                tab_stop: true,
                group: None,
                parent: None,
            },
            Control {
                id: 2,
                name: "CommandButton1".into(),
                kind: ControlKind::CommandButton,
                bounds: Bounds { left: 12.0, top: 48.0, width: 120.0, height: 24.0 },
                caption: "Apply".into(),
                text: String::new(),
                enabled: true,
                visible: true,
                tab_stop: true,
                group: None,
                parent: None,
            },
        ],
        tab_order: vec![1, 2],
    };
    let archive = form_archive::encode(&form, source)?;
    let path = std::env::temp_dir().join(format!("gridcraft-m5-form-{}.tvform", std::process::id()));
    fs::write(&path, archive.as_bytes())?;
    let picked = path.clone();
    let services = Services { pick_open: Some(Box::new(move || Some(picked.to_string_lossy().into_owned()))), ..Default::default() };
    let mut session = Session::new();
    session.new_workbook();
    let mut app = SheetApp::new(session, services);
    app.run("vba.open", json!({})).map_err(std::io::Error::other)?;
    let import = app.run("vba.importSource", json!({})).map_err(std::io::Error::other)?;
    assert_eq!(import, json!({"imported":true}));

    let form_id = app
        .vba
        .as_ref()
        .and_then(|editor| editor.ide.modules.iter().find(|module| editor.ide.form(module.id).is_some()))
        .map(|module| module.id)
        .ok_or("imported form missing from embedded IDE")?;
    assert_eq!(app.vba.as_ref().and_then(|editor| editor.ide.form(form_id)).map(|designer| designer.form().name.as_str()), Some("M5Form"));

    dispatch_form(&mut app, FormCommand::Run(form_id, true))?;
    settle(&mut app)?;
    let editor = app.vba.as_ref().ok_or("editor closed")?;
    assert_eq!(editor.ide.mode(), gridcraft_vba::Mode::Waiting, "{:?}", editor.ide.diagnostics);
    assert!(editor.ide.form_visible(form_id), "{:?}", editor.ide.diagnostics);
    assert!(editor.ide.form_is_modal(form_id), "{:?}", editor.ide.diagnostics);
    dispatch_form(&mut app, FormCommand::Event(form_id, Some(1), Some("hello".into())))?;
    settle(&mut app)?;
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::new(0, 0)), Value::Number(5.0));
    dispatch_form(&mut app, FormCommand::Event(form_id, Some(2), None))?;
    settle(&mut app)?;
    for _ in 0..20 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        if editor.ide.mode() == gridcraft_vba::Mode::Completed {
            break;
        }
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
    }
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet missing")?.value(CellRef::new(1, 0)), Value::Number(42.0));
    let editor = app.vba.as_ref().ok_or("editor closed")?;
    assert!(!editor.ide.form_visible(form_id), "{:?}", editor.ide.diagnostics);
    assert!(!editor.ide.form_is_modal(form_id), "{:?}", editor.ide.diagnostics);
    assert_eq!(app.vba.as_ref().map(|editor| editor.ide.mode()), Some(gridcraft_vba::Mode::Completed));
    let editor = app.vba.as_mut().ok_or("editor closed")?;
    editor.bridge.finish(&mut app.session)?;
    assert!(!app.session.in_group(editor.bridge.document()));

    app.run("vba.reset", json!({})).map_err(std::io::Error::other)?;
    dispatch_form(&mut app, FormCommand::Run(form_id, true))?;
    settle(&mut app)?;
    assert!(app.vba.as_ref().is_some_and(|editor| editor.ide.form_is_modal(form_id)));
    dispatch_form(&mut app, FormCommand::Event(form_id, None, None))?;
    settle(&mut app)?;
    for _ in 0..20 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        if editor.ide.mode() == gridcraft_vba::Mode::Completed {
            break;
        }
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
    }
    let editor = app.vba.as_ref().ok_or("editor closed")?;
    assert_eq!(editor.ide.mode(), gridcraft_vba::Mode::Completed, "{:?}", editor.ide.diagnostics);
    assert!(!editor.ide.form_visible(form_id));
    assert!(!editor.ide.form_is_modal(form_id));

    fs::remove_file(path)?;
    Ok(())
}

fn dispatch_form(app: &mut SheetApp, command: FormCommand) -> R {
    let editor = app.vba.as_mut().ok_or("editor closed")?;
    let mut host = editor.bridge.host(&mut app.session);
    vba_ui_egui::apply(&mut editor.ide, Command::Form(command), &mut host)?;
    Ok(())
}

fn settle(app: &mut SheetApp) -> R {
    for _ in 0..1000 {
        let editor = app.vba.as_mut().ok_or("editor closed")?;
        if editor.ide.mode() != gridcraft_vba::Mode::Running {
            return Ok(());
        }
        editor.bridge.tick(&mut app.session, &mut editor.ide)?;
    }
    Err("form execution exceeded the acceptance tick budget".into())
}

#[test]
fn form_launch_and_instance_callbacks_require_workbook_execution_consent() -> R {
    use gridcraft_engine::{DocState, Trust, model::Workbook};
    let mut session = Session::new();
    session.add_document(DocState::new(Workbook::new(), None, "Untrusted".into(), Trust::FromFile));
    let mut app = SheetApp::new(session, Services::default());
    app.run("vba.open", json!({})).map_err(std::io::Error::other)?;
    let editor = app.vba.as_mut().ok_or("editor")?;
    let id = editor.ide.add_form("GuardForm")?;
    editor.ide.set_text("Private Sub UserForm_QueryClose(Cancel As Integer,CloseMode As Integer)\nDebug.Print \"executed\"\nEnd Sub".into())?;
    assert!(gridcraft_ui_egui::vba::dispatch_editor_command(editor, &mut app.session, Command::Form(FormCommand::Run(id, false))).is_err());
    assert!(editor.ide.output().is_empty());
    editor.bridge.consent(true);
    gridcraft_ui_egui::vba::dispatch_editor_command(editor, &mut app.session, Command::Form(FormCommand::Run(id, false)))?;
    settle(&mut app)?;
    let editor = app.vba.as_mut().ok_or("editor")?;
    let instance = *editor.ide.form_instances()?.first().ok_or("form instance")?;
    editor.bridge.consent(false);
    assert!(
        gridcraft_ui_egui::vba::dispatch_editor_command(editor, &mut app.session, Command::Form(FormCommand::InstanceEvent(instance, None, None)))
            .is_err()
    );
    assert!(editor.ide.output().is_empty());
    assert!(editor.ide.form_instance_visible(instance));
    Ok(())
}

#[path = "../../vba/tests/support/native_forms.rs"]
mod native_fixture;

#[test]
fn native_refedit_uses_authorized_sheet_selection_and_cancels_on_trust_loss() -> R {
    use gridcraft_vba::vba_ide::{NativeFormResource, Value as VbaValue, VbaString};
    let mut session = Session::new();
    session.add_document(gridcraft_engine::DocState::new(
        gridcraft_engine::model::Workbook::new(),
        None,
        "RefEdit.xlsm".into(),
        gridcraft_engine::Trust::FromFile,
    ));
    let mut app = SheetApp::new(session, Services::default());
    app.run("vba.open", json!({"source":"Sub Main()\nNativeForm.Show\nEnd Sub"})).map_err(std::io::Error::other)?;
    let source = native_fixture::SOURCE.replace(
        "Private Sub Apply_Click()",
        "Private Sub Entry_Change()\nRange(\"A1\").Value2=Len(Entry.Text)\nDebug.Print Entry.Text\nEnd Sub\nPrivate Sub Apply_Click()",
    );
    let export = String::from_utf8(native_fixture::export())?.replace(native_fixture::SOURCE, &source);
    let editor = app.vba.as_mut().ok_or("editor")?;
    let id = editor.ide.import_native_form_export(
        "NativeForm.frm",
        export.as_bytes(),
        1252,
        &[NativeFormResource { name: "NativeForm.frx".into(), bytes: native_fixture::refedit_companion() }],
    )?;
    let original = editor.ide.native_form(id).ok_or("native")?.clone();
    app.run("vba.trust", json!({"allow":true})).map_err(std::io::Error::other)?;
    app.run("vba.run", json!({"entry":"Main"})).map_err(std::io::Error::other)?;
    settle(&mut app)?;
    let f = *app.vba.as_ref().ok_or("editor")?.ide.form_instances()?.first().ok_or("instance")?;
    {
        let editor = app.vba.as_mut().ok_or("editor")?;
        gridcraft_ui_egui::vba::dispatch_editor_command(editor, &mut app.session, Command::Form(FormCommand::InstanceDrop(f, 12, true)))?;
    }
    settle(&mut app)?;
    let ctx = egui::Context::default();
    app.logic(&ctx);
    app.run("selection.set", json!({"range":"B2:C4"})).map_err(std::io::Error::other)?;
    app.logic(&ctx);
    settle(&mut app)?;
    let editor = app.vba.as_ref().ok_or("editor")?;
    assert_eq!(editor.ide.form_instance_field(f, Some(12), "VbaText")?, VbaValue::String(VbaString::from_utf8("Sheet1!$B$2:$C$4")?));
    assert_eq!(editor.ide.output(), vec!["initialized", "Sheet1!$B$2:$C$4"]);
    assert_eq!(editor.ide.form_range_selection()?, Some((f, 12)));
    assert_eq!(editor.ide.native_form(id), Some(&original));
    assert!(app.run("cell.set", json!({"cell":"B1","input":"99"})).is_err());
    assert_eq!(app.session.doc()?.wb.sheet(0).ok_or("sheet")?.value(CellRef::new(0, 0)), Value::Number(16.));
    app.run("selection.set", json!({"range":"D5,E6"})).map_err(std::io::Error::other)?;
    app.logic(&ctx);
    settle(&mut app)?;
    assert_eq!(app.vba.as_ref().ok_or("editor")?.ide.output().last().map(String::as_str), Some("Sheet1!$D$5,Sheet1!$E$6"));
    use egui_kittest::kittest::Queryable;
    app.run("vba.return", json!({})).map_err(std::io::Error::other)?;
    app.use_prepared_fonts();
    let snapshot = std::env::var("GRIDCRAFT_REFEDIT_SNAPSHOT").ok();
    let builder = egui_kittest::Harness::builder().with_size(egui::vec2(1300., 900.)).with_context_setup(|ctx| {
        ctx.set_fonts(gridcraft_ui_egui::theme::pinned_font_definitions());
        gridcraft_ui_egui::theme::apply(ctx, false);
    });
    let builder = if snapshot.is_some() { builder.wgpu() } else { builder };
    let mut h = builder.build_ui_state(
        |ui, a: &mut SheetApp| {
            a.logic(ui.ctx());
            a.ui(ui);
        },
        app,
    );
    h.run_steps(6);
    h.get_by_label("Restore dialog").click();
    h.run_steps(8);
    assert_eq!(h.state().vba.as_ref().ok_or("editor")?.ide.form_range_selection()?, None);
    h.get_by_label("Select range").click();
    h.run_steps(8);
    assert_eq!(h.state().vba.as_ref().ok_or("editor")?.ide.form_range_selection()?, Some((f, 12)));
    // Real worksheet pointer input must reach the same canonical RefEdit value.
    let a = h.state();
    let sh = a.session.doc()?.wb.active().ok_or("sheet")?;
    let position = gridcraft_ui_egui::grid::Geo::new(sh, a.grid.rect.ok_or("grid")?, a.view().scroll).cell_rect(sh, CellRef::new(7, 5)).center();
    for pressed in [true, false] {
        h.event(egui::Event::PointerMoved(position));
        h.event(egui::Event::PointerButton { pos: position, button: egui::PointerButton::Primary, pressed, modifiers: egui::Modifiers::NONE });
        h.run_steps(4);
    }
    assert_eq!(
        h.state().vba.as_ref().ok_or("editor")?.ide.form_instance_field(f, Some(12), "VbaText")?,
        VbaValue::String(VbaString::from_utf8("Sheet1!$F$8")?)
    );
    h.key_press(egui::Key::F2);
    h.run_steps(3);
    h.event(egui::Event::Text("99".into()));
    h.run_steps(3);
    assert!(h.state().editor.is_none(), "Range selection must not enter cell editing");
    assert_eq!(h.state().session.doc()?.wb.active().ok_or("sheet")?.value(CellRef::new(7, 5)), Value::Empty);
    if let Some(path) = snapshot {
        h.render()?.save(path)?;
    }
    h.state_mut().run("vba.trust", json!({"allow":false})).map_err(std::io::Error::other)?;
    h.run_steps(4);
    let editor = h.state().vba.as_ref().ok_or("editor")?;
    assert_eq!(editor.ide.form_range_selection()?, None);
    assert_eq!(editor.ide.native_form(id), Some(&original));
    Ok(())
}

#[test]
fn native_bound_widgets_write_cells_and_refresh_from_sheet_commands() -> R {
    use egui_kittest::kittest::{By, Queryable};
    use gridcraft_vba::vba_ide::{NativeFormResource, Value as VbaValue, VbaString};
    let mut session = Session::new();
    session.add_document(gridcraft_engine::DocState::new(
        gridcraft_engine::model::Workbook::new(),
        None,
        "Bindings.xlsm".into(),
        gridcraft_engine::Trust::FromFile,
    ));
    for (cell, input) in [("A1", "seed"), ("B1", "first"), ("C1", "A"), ("B2", "second"), ("C2", "B"), ("D1", "A")] {
        session.execute("cell.set", json!({"cell":cell,"input":input}))?;
    }
    let mut app = SheetApp::new(session, Services::default());
    app.run("vba.open", json!({"source":"Sub Main()\nNativeForm.Show 0\nEnd Sub"})).map_err(std::io::Error::other)?;
    let editor = app.vba.as_mut().ok_or("editor")?;
    let id = editor.ide.import_native_form_export(
        "NativeForm.frm",
        &native_fixture::export(),
        1252,
        &[NativeFormResource { name: "NativeForm.frx".into(), bytes: native_fixture::binding_companion() }],
    )?;
    let original = editor.ide.native_form(id).ok_or("native")?.clone();
    app.run("vba.trust", json!({"allow":true})).map_err(std::io::Error::other)?;
    app.run("vba.run", json!({"entry":"Main"})).map_err(std::io::Error::other)?;
    settle(&mut app)?;
    let f = *app.vba.as_ref().ok_or("editor")?.ide.form_instances()?.first().ok_or("instance")?;
    app.run("vba.return", json!({})).map_err(std::io::Error::other)?;
    app.use_prepared_fonts();
    let snapshot = std::env::var("GRIDCRAFT_BINDING_SNAPSHOT").ok();
    let builder = egui_kittest::Harness::builder().with_size(egui::vec2(1300., 900.)).with_context_setup(|ctx| {
        ctx.set_fonts(gridcraft_ui_egui::theme::pinned_font_definitions());
        gridcraft_ui_egui::theme::apply(ctx, false);
    });
    let builder = if snapshot.is_some() { builder.wgpu() } else { builder };
    let mut h = builder.build_ui_state(
        |ui, a: &mut SheetApp| {
            a.logic(ui.ctx());
            a.ui(ui);
        },
        app,
    );
    h.run_steps(12);
    h.get(By::new().label("Entry").role(egui::accesskit::Role::TextInput)).click();
    h.run_steps(8);
    for pressed in [true, false] {
        h.event(egui::Event::Key {
            key: egui::Key::A,
            physical_key: None,
            pressed,
            repeat: false,
            modifiers: egui::Modifiers { ctrl: true, command: true, ..Default::default() },
        });
    }
    h.run_steps(2);
    h.event(egui::Event::Text("typed".into()));
    h.run_steps(40);
    assert_eq!(h.state().session.doc()?.wb.active().ok_or("sheet")?.value(CellRef::new(0, 0)), Value::Text("typed".into()));
    h.get_by_label("List1 row 1: second | B").click();
    h.run_steps(40);
    assert_eq!(h.state().session.doc()?.wb.active().ok_or("sheet")?.value(CellRef::new(0, 3)), Value::Text("B".into()));
    h.state_mut().run("cell.set", json!({"cell":"A1","input":"external"})).map_err(std::io::Error::other)?;
    h.state_mut().run("cell.set", json!({"cell":"B2","input":"updated"})).map_err(std::io::Error::other)?;
    h.run_steps(60);
    let editor = h.state().vba.as_ref().ok_or("editor")?;
    assert_eq!(editor.ide.form_instance_field(f, Some(12), "VbaText")?, VbaValue::String(VbaString::from_utf8("external")?));
    assert_eq!(editor.ide.native_form(id), Some(&original));
    assert!(h.query_by_label("List1 row 1: updated | B").is_some());
    if let Some(path) = snapshot {
        h.render()?.save(path)?;
    }
    h.state_mut().run("vba.trust", json!({"allow":false})).map_err(std::io::Error::other)?;
    h.state_mut().run("cell.set", json!({"cell":"A1","input":"deferred"})).map_err(std::io::Error::other)?;
    h.run_steps(10);
    assert_eq!(
        h.state().vba.as_ref().ok_or("editor")?.ide.form_instance_field(f, Some(12), "VbaText")?,
        VbaValue::String(VbaString::from_utf8("external")?)
    );
    Ok(())
}

#[test]
fn native_document_click_waits_for_deactivate_continue_and_reset_discards_it() -> R {
    use gridcraft_vba::vba_ide::NativeFormResource;
    for (reset, prefix) in [(false, false), (true, false), (false, true), (true, true)] {
        let mut session = Session::new();
        session.new_workbook();
        let mut app = SheetApp::new(session, Services::default());
        app.run("vba.open", json!({"source":"Sub Main()\nNativeForm.Show 0\nEnd Sub"})).map_err(std::io::Error::other)?;
        let pause = if prefix { "Debug.Print Entry.Text" } else { "Stop" };
        let code = format!(
            "Attribute VB_Name = \"NativeForm\"\nAttribute VB_PredeclaredId = True\nPrivate Sub UserForm_Initialize()\nMe.BackColor=RGB(230,240,250)\nEntry.BackColor=RGB(250,230,200)\nEntry.ForeColor=RGB(80,30,20)\nEntry.FontSize=12.25\nEntry.FontItalic=True\nEntry.FontUnderline=True\nEntry.TextAlign=fmTextAlignRight\nHeading.TextAlign=fmTextAlignLeft\nHeading.WordWrap=False\nApply.WordWrap=True\nEntry.WordWrap=False\nEntry.BorderColor=RGB(80,30,20)\nEntry.BorderStyle=1\nSet Apply.Font=Entry.Font\nApply.BackStyle=0\nApply.BackColor=RGB(190,220,240)\nDim rows As MSForms.ListBox\nSet rows=Me.Controls.Add(\"Forms.ListBox.1\",\"Rows\")\nrows.Left=145\nrows.Top=70\nrows.Width=110\nrows.Height=90\nrows.BackColor=RGB(31,53,79)\nrows.ForeColor=RGB(220,230,240)\nrows.SpecialEffect=0\nrows.BorderStyle=0\nrows.TextAlign=fmTextAlignCenter\nrows.AddItem \"Flat list\"\nEnd Sub\nPrivate Sub UserForm_Deactivate()\nDebug.Print \"deactivated\"\n{pause}\nRange(\"A1\").Value=42\nEnd Sub{}",
            if prefix {
                "\nPrivate Sub Entry_KeyPress(ByVal k As MSForms.ReturnInteger)\nDebug.Print \"press\"\nk.Value=122\nEnd Sub\nPrivate Sub Entry_Change()\nDebug.Print \"change\"\nStop\nEnd Sub"
            } else {
                ""
            }
        );
        let export = String::from_utf8(native_fixture::export())?.replace(native_fixture::SOURCE, &code);
        let editor = app.vba.as_mut().ok_or("editor")?;
        editor.ide.import_native_form_export(
            "NativeForm.frm",
            export.as_bytes(),
            1252,
            &[NativeFormResource { name: "NativeForm.frx".into(), bytes: native_fixture::companion() }],
        )?;
        app.run("vba.trust", json!({"allow":true})).map_err(std::io::Error::other)?;
        app.run("vba.run", json!({"entry":"Main"})).map_err(std::io::Error::other)?;
        settle(&mut app)?;
        app.run("vba.return", json!({})).map_err(std::io::Error::other)?;
        app.use_prepared_fonts();
        let ctx = egui::Context::default();
        ctx.set_fonts(gridcraft_ui_egui::theme::pinned_font_definitions());
        gridcraft_ui_egui::theme::apply(&ctx, false);
        let step = |app: &mut SheetApp, events: Vec<egui::Event>| {
            let mut raw = egui::RawInput {
                screen_rect: Some(egui::Rect::from_min_size(egui::Pos2::ZERO, egui::vec2(1300., 900.))),
                events,
                ..Default::default()
            };
            app.raw_input_hook(&ctx, &mut raw);
            let mut out = ctx.run_ui(raw, |ui| {
                app.logic(ui.ctx());
                app.ui(ui);
            });
            out.textures_delta.clear();
        };
        for _ in 0..12 {
            step(&mut app, vec![]);
        }
        let rect = app.grid.cells_rect.ok_or("canvas")?;
        let pos = rect.right_bottom() - egui::vec2(30., 30.);
        let before = app.session.doc()?.selection.current();
        let mut events = if prefix {
            vec![
                egui::Event::Key {
                    key: egui::Key::A,
                    physical_key: Some(egui::Key::A),
                    pressed: true,
                    repeat: false,
                    modifiers: egui::Modifiers::NONE,
                },
                egui::Event::Text("a".into()),
                egui::Event::Key {
                    key: egui::Key::A,
                    physical_key: Some(egui::Key::A),
                    pressed: false,
                    repeat: false,
                    modifiers: egui::Modifiers::NONE,
                },
            ]
        } else {
            vec![]
        };
        events.extend([true, false].into_iter().map(|pressed| egui::Event::PointerButton {
            pos,
            button: egui::PointerButton::Primary,
            pressed,
            modifiers: egui::Modifiers::NONE,
        }));
        step(&mut app, events);
        for _ in 0..40 {
            step(&mut app, vec![]);
        }
        assert_eq!(app.vba.as_ref().ok_or("editor")?.ide.mode(), gridcraft_vba::Mode::Break);
        assert_eq!(app.session.doc()?.selection.current(), before);
        assert_eq!(app.session.doc()?.wb.active().ok_or("sheet")?.value(CellRef::default()), Value::Empty);
        assert_eq!(app.vba.as_ref().ok_or("editor")?.ide.output(), if prefix { vec!["press", "change"] } else { vec!["deactivated"] });
        app.run(if reset { "vba.reset" } else { "vba.run" }, json!({})).map_err(std::io::Error::other)?;
        for _ in 0..90 {
            step(&mut app, vec![]);
        }
        assert!(!app.vba.as_ref().ok_or("editor")?.view.document_input_pending());
        if reset {
            assert_eq!(app.session.doc()?.selection.current(), before);
        } else {
            assert_ne!(app.session.doc()?.selection.current(), before);
            assert_eq!(app.session.doc()?.wb.active().ok_or("sheet")?.value(CellRef::default()), Value::Number(42.));
            assert_eq!(
                app.vba.as_ref().ok_or("editor")?.ide.output(),
                if prefix { vec!["press", "change", "deactivated", "z"] } else { vec!["deactivated"] }
            );
            if let Ok(path) = std::env::var("GRIDCRAFT_DOCUMENT_SNAPSHOT") {
                let mut h = egui_kittest::Harness::builder()
                    .with_size(egui::vec2(1300., 900.))
                    .with_context_setup(|ctx| {
                        ctx.set_fonts(gridcraft_ui_egui::theme::pinned_font_definitions());
                        gridcraft_ui_egui::theme::apply(ctx, false);
                    })
                    .wgpu()
                    .build_ui_state(
                        |ui, a: &mut SheetApp| {
                            a.logic(ui.ctx());
                            a.ui(ui);
                        },
                        app,
                    );
                h.run_steps(5);
                h.render()?.save(path)?;
            }
        }
    }
    Ok(())
}
