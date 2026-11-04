//! Original deterministic MS-OVBA fixture builder; test-only, never Office output.
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
pub fn project(modules: &[(&str, &str, &str)]) -> Vec<u8> {
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
    tlv(&mut dir, 0xf, &(modules.len() as u16).to_le_bytes());
    tlv(&mut dir, 0x13, &0u16.to_le_bytes());
    let mut metadata = String::from("ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n");
    for (name, kind, _) in modules {
        tlv(&mut dir, 0x19, name.as_bytes());
        tlv(&mut dir, 0x47, &utf16(name));
        tlv(&mut dir, 0x1a, name.as_bytes());
        tlv(&mut dir, 0x32, &utf16(name));
        tlv(&mut dir, 0x1c, b"");
        tlv(&mut dir, 0x48, b"");
        tlv(&mut dir, 0x31, &0u32.to_le_bytes());
        tlv(&mut dir, 0x1e, &0u32.to_le_bytes());
        tlv(&mut dir, 0x2c, &0u16.to_le_bytes());
        for id in [if *kind == "Module" { 0x21 } else { 0x22 }, 0x2b] {
            tlv(&mut dir, id, b"");
        }
        let declaration = if *kind == "Document" { format!("Document={name}/&H00000000") } else { format!("{kind}={name}") };
        metadata.push_str(&format!("{declaration}\r\n"));
    }
    tlv(&mut dir, 0x10, b"");
    metadata.push_str(&format!(
        "Name=\"AuthoredProject\"\r\nCMG={}\r\nDPB={}\r\nGC={}\r\n",
        protection(&[0; 4]),
        protection(&[0]),
        protection(&[255])
    ));
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    file.create_storage("/VBA").unwrap();
    for (path, bytes) in
        [("/PROJECT", metadata.into_bytes()), ("/VBA/dir", packed(&dir)), ("/VBA/_VBA_PROJECT", vec![0xcc, 0x61, 0xff, 0xff, 0, 0, 0])]
    {
        file.create_stream(path).unwrap().write_all(&bytes).unwrap();
    }
    for (name, _, source) in modules {
        file.create_stream(format!("/VBA/{name}")).unwrap().write_all(&packed(source.as_bytes())).unwrap();
    }
    // CFB storage timestamps otherwise depend on ambient time, even in authored fixtures.
    let paths: Vec<_> = file.walk().map(|entry| entry.path().to_owned()).collect();
    for path in paths {
        file.set_created_time(&path, web_time::SystemTime::UNIX_EPOCH).unwrap();
        file.set_modified_time(&path, web_time::SystemTime::UNIX_EPOCH).unwrap();
    }
    file.into_inner().into_inner()
}
