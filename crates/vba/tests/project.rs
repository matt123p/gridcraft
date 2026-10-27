//! Small authored MS-OVBA project, built in memory; no Office-authored binaries.
use gridcraft_engine::Session;
use gridcraft_vba::{Action, Bridge, Mode};
use std::io::{Cursor, Write};

fn tlv(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend(id.to_le_bytes());
    out.extend((data.len() as u32).to_le_bytes());
    out.extend(data);
}
fn utf16(s: &str) -> Vec<u8> {
    s.encode_utf16().flat_map(u16::to_le_bytes).collect()
}
fn packed(data: &[u8]) -> Vec<u8> {
    assert!(!data.is_empty() && data.len() < 3000);
    let mut tokens = vec![];
    for group in data.chunks(8) {
        tokens.push(0);
        tokens.extend(group);
    }
    let mut out = vec![1];
    out.extend((0xb000u16 | (tokens.len() as u16 - 1)).to_le_bytes());
    out.extend(tokens);
    out
}
fn protection(data: &[u8]) -> String {
    let seed = 0u8;
    let key = 0x35u8;
    let mut bytes = vec![seed, seed ^ 2, seed ^ key];
    let mut plain = key;
    let mut previous = seed ^ 2;
    let mut last = seed ^ key;
    for b in (data.len() as u32).to_le_bytes().into_iter().chain(data.iter().copied()) {
        let encrypted = b ^ previous.wrapping_add(plain);
        bytes.push(encrypted);
        previous = last;
        last = encrypted;
        plain = b;
    }
    format!("\"{}\"", bytes.iter().map(|b| format!("{b:02X}")).collect::<String>())
}
fn project(document: bool) -> Vec<u8> {
    let mut dir = vec![];
    for (id, n) in [(1, 1u32), (2, 0x409), (0x14, 0x409)] {
        tlv(&mut dir, id, &n.to_le_bytes());
    }
    tlv(&mut dir, 3, &1252u16.to_le_bytes());
    tlv(&mut dir, 4, b"AuthoredProject");
    for id in [5, 0x40, 6, 0x3d] {
        tlv(&mut dir, id, b"");
    }
    for id in [7, 8] {
        tlv(&mut dir, id, &0u32.to_le_bytes());
    }
    tlv(&mut dir, 9, &1u32.to_le_bytes());
    dir.extend(0u16.to_le_bytes());
    tlv(&mut dir, 0xf, &1u16.to_le_bytes());
    tlv(&mut dir, 0x13, &0u16.to_le_bytes());
    tlv(&mut dir, 0x19, b"Module1");
    tlv(&mut dir, 0x47, &utf16("Module1"));
    tlv(&mut dir, 0x1a, b"Source1");
    tlv(&mut dir, 0x32, &utf16("Source1"));
    tlv(&mut dir, 0x1c, b"");
    tlv(&mut dir, 0x48, b"");
    tlv(&mut dir, 0x31, &0u32.to_le_bytes());
    tlv(&mut dir, 0x1e, &0u32.to_le_bytes());
    tlv(&mut dir, 0x2c, &0u16.to_le_bytes());
    for id in [if document { 0x22 } else { 0x21 }, 0x2b, 0x10] {
        tlv(&mut dir, id, b"");
    }
    let kind = if document { "Document=Module1/00000000" } else { "Module=Module1" };
    let metadata = format!(
        "ID=\"{{00000000-0000-0000-0000-000000000000}}\"\r\n{kind}\r\nName=\"AuthoredProject\"\r\nCMG={}\r\nDPB={}\r\nGC={}\r\n",
        protection(&[0; 4]),
        protection(&[0]),
        protection(&[255])
    );
    let source = "Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nRange(\"A1\").Value2 = 99\r\nEnd Sub\r\nSub Auto_Open()\r\nRange(\"A2\").Value2 = 111\r\nEnd Sub\r\n";
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    file.create_storage("/VBA").unwrap();
    for (path, bytes) in [
        ("/PROJECT", metadata.into_bytes()),
        ("/VBA/dir", packed(&dir)),
        ("/VBA/Source1", packed(source.as_bytes())),
        ("/VBA/_VBA_PROJECT", vec![0xcc, 0x61, 0xff, 0xff, 0, 0, 0]),
    ] {
        file.create_stream(path).unwrap().write_all(&bytes).unwrap();
    }
    file.into_inner().into_inner()
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn import_is_inert_preserves_source_and_runs_only_explicit_entry() {
    let mut s = Session::new();
    s.new_workbook();
    let bytes = project(false);
    std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb).vba_project = Some(bytes.clone().into());
    let mut b = Bridge::new(&s).unwrap();
    let mut ide = b.import_project(&s).unwrap();
    assert_eq!(ide.mode(), Mode::Design);
    assert_eq!(ide.modules.len(), 1);
    assert!(ide.modules[0].text.contains("\r\n"));
    assert_eq!(s.doc().unwrap().undo.len(), 0);
    b.dispatch(&mut s, &mut ide, Action::Run).unwrap();
    for _ in 0..30 {
        b.tick(&mut s, &mut ide).unwrap();
    }
    assert_eq!(ide.mode(), Mode::Completed, "{:?}", ide.diagnostics);
    let sheet = s.doc().unwrap().wb.sheet(0).unwrap();
    assert_eq!(sheet.value(gridcraft_engine::cell(0, 0)).display(), "99");
    assert!(sheet.value(gridcraft_engine::cell(1, 0)).is_empty());
    assert_eq!(s.doc().unwrap().wb.vba_project.as_deref().unwrap().as_slice(), bytes);
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn executable_document_module_is_refused_without_mutation() {
    let mut s = Session::new();
    s.new_workbook();
    let bytes = project(true);
    std::sync::Arc::make_mut(&mut s.doc_mut().unwrap().wb).vba_project = Some(bytes.into());
    let b = Bridge::new(&s).unwrap();
    let before = s.doc().unwrap().wb.clone();
    let Err(error) = b.import_project(&s) else { panic!("expected refusal") };
    assert!(error.message.contains("class/document"), "{error}");
    assert_eq!(*before, *s.doc().unwrap().wb);
}
