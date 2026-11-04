//! Independently authored MS-OFORMS fixtures. MIT OR Apache-2.0, no Office output.
#![allow(dead_code, clippy::unwrap_used, clippy::indexing_slicing)]
use std::io::{Cursor, Write};
pub const SOURCE: &str = "Attribute VB_Name = \"NativeForm\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\nOption Explicit\r\nPrivate Sub UserForm_Initialize()\r\nDebug.Print \"initialized\"\r\nEnd Sub\r\nPrivate Sub Apply_Click()\r\nDebug.Print Entry.Text\r\nMe.Hide\r\nEnd Sub\r\n";
pub fn frame(blob: bool) -> String {
    format!(
        "VERSION 5.00\r\nBegin {{C62A69F0-16DC-11CE-9E98-00AA00574A4F}} NativeForm\r\n Caption = \"Native form\"\r\n ClientHeight = 4000\r\n ClientWidth = 6000\r\n{} StartUpPosition = 1 'CenterOwner\r\nEnd\r\n",
        if blob { " OleObjectBlob = \"NativeForm.frx\":0000\r\n" } else { "" }
    )
}
pub fn export() -> Vec<u8> {
    format!("{}{}", frame(true), SOURCE).into_bytes()
}
pub fn u32s(values: &[u32]) -> Vec<u8> {
    values.iter().flat_map(|n| n.to_le_bytes()).collect()
}
pub fn record(version: u16, body: &[u8]) -> Vec<u8> {
    let mut out = version.to_le_bytes().to_vec();
    out.extend((body.len() as u16).to_le_bytes());
    out.extend(body);
    out
}
fn string(value: &str) -> Vec<u8> {
    let mut bytes = value.as_bytes().to_vec();
    while !bytes.len().is_multiple_of(4) {
        bytes.push(0);
    }
    bytes
}
pub fn controls() -> Vec<Vec<u8>> {
    controls_with_properties(0)
}
fn controls_with_properties(max_length: u32) -> Vec<Vec<u8>> {
    let label = leaf(false, "Authored label");
    let mut b = if max_length == 0 {
        u32s(&[0x80400100, 0, 0x80000000 | 7, 2540, 635])
    } else {
        u32s(&[0x80400108, 0, max_length, 0x80000000 | 7, 2540, 635])
    };
    b.extend(string("initial"));
    let mut textbox = record(0x200, &b);
    textbox.extend(record(0x200, &[0; 4]));
    let button = leaf(false, "Apply");
    vec![label, textbox, button]
}
fn leaf(_label: bool, caption: &str) -> Vec<u8> {
    let mut b = u32s(&[0x28, 0x80000000 | caption.len() as u32]);
    b.extend(string(caption));
    b.extend(u32s(&[2540, 635]));
    let mut out = record(0x200, &b);
    out.extend(record(0x200, &[0; 4]));
    out
}
fn site(id: u32, name: &str, class: u16, tab: u16, size: usize) -> Vec<u8> {
    let mut b = u32s(&[0x1f5, 0x80000000 | name.len() as u32, id, 0x13, size as u32]);
    b.extend(tab.to_le_bytes());
    b.extend(class.to_le_bytes());
    b.extend(string(name));
    b.extend(u32s(&[423, 635 + (id - 11) * 1000]));
    record(0, &b)
}
pub fn streams() -> Vec<(String, Vec<u8>)> {
    streams_with_properties(0, false, false)
}
fn streams_with_properties(max_length: u32, default: bool, cancel: bool) -> Vec<(String, Vec<u8>)> {
    let controls = controls_with_properties(max_length);
    let mut b = u32s(&[(1 << 10) | (1 << 11) | (1 << 19) | (1 << 27), 0x80000000 | 11, 32000, 10583, 7056, 10583, 7056]);
    b.extend(string("Native form"));
    let mut f = record(0x400, &b);
    f.extend(0u16.to_le_bytes()); // empty class table, all three types use native cache indices
    f.extend(3u32.to_le_bytes());
    let mut sites = vec![0, 0x83, 1, 0]; // three immediate children, padded to four bytes
    for (((id, name), class), (tab, c)) in
        [(11, "Heading"), (12, "Entry"), (13, "Apply")].into_iter().zip([21, 23, 17]).zip(controls.iter().enumerate())
    {
        let mut encoded = site(id, name, class, tab as u16, c.len());
        if class == 17 {
            let flags = 0x13u32 | if default { 4 } else { 0 } | if cancel { 8 } else { 0 };
            encoded[16..20].copy_from_slice(&flags.to_le_bytes());
        }
        sites.extend(encoded);
    }
    f.extend((sites.len() as u32).to_le_bytes());
    f.extend(sites);
    vec![("f".into(), f), ("o".into(), controls.concat())]
}
pub fn companion() -> Vec<u8> {
    companion_with_properties(0, false, false)
}
pub fn companion_with_properties(max_length: u32, default: bool, cancel: bool) -> Vec<u8> {
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    for (name, bytes) in streams_with_properties(max_length, default, cancel) {
        file.create_stream(format!("/{name}")).unwrap().write_all(&bytes).unwrap();
    }
    let bytes = file.into_inner().into_inner();
    let mut out = (bytes.len() as u32).to_le_bytes().to_vec();
    out.extend(bytes);
    out
}
pub fn add_designer(project: Vec<u8>) -> Vec<u8> {
    let mut file = cfb::CompoundFile::open(Cursor::new(project)).unwrap();
    file.create_storage("/NativeForm").unwrap();
    file.create_stream("/NativeForm/\u{3}VBFrame").unwrap().write_all(frame(false).as_bytes()).unwrap();
    for (name, bytes) in streams() {
        file.create_stream(format!("/NativeForm/{name}")).unwrap().write_all(&bytes).unwrap();
    }
    file.into_inner().into_inner()
}

/// Original six-family form; native cached MorphData classes, never Office output.
pub fn boolean_companion() -> Vec<u8> {
    let mut all = controls();
    for (display, caption, group) in [(4u8, "Check", ""), (5, "Choice A", "choices"), (5, "Choice B", "choices"), (6, "Toggle", "")] {
        let mut body = u32s(&[0x80c00140, u32::from(!group.is_empty())]);
        body.extend([display, 0, 0, 0]);
        body.extend(u32s(&[0x80000001, 0x80000000 | caption.len() as u32]));
        if !group.is_empty() {
            body.extend(u32s(&[0x80000000 | group.len() as u32]));
        }
        body.extend(u32s(&[2540, 635]));
        body.extend(string("0"));
        body.extend(string(caption));
        if !group.is_empty() {
            body.extend(string(group));
        }
        let mut control = record(0x200, &body);
        control.extend(record(0x200, &[0; 4]));
        all.push(control);
    }
    let mut root = u32s(&[(1 << 10) | (1 << 11) | (1 << 19) | (1 << 27), 0x8000000b, 32000, 10583, 7056, 10583, 7056]);
    root.extend(string("Native form"));
    let mut f = record(0x400, &root);
    f.extend(0u16.to_le_bytes());
    f.extend((all.len() as u32).to_le_bytes());
    let mut sites = vec![0, 0x87, 1, 0];
    for (i, ((name, class), control)) in
        [("Heading", 21), ("Entry", 23), ("Apply", 17), ("Check1", 26), ("ChoiceA", 27), ("ChoiceB", 27), ("Toggle1", 28)]
            .into_iter()
            .zip(&all)
            .enumerate()
    {
        let mut encoded = site(11 + i as u32, name, class, i as u16, control.len());
        if i >= 3 {
            let n = encoded.len();
            encoded[n - 8..n - 4].copy_from_slice(&4000u32.to_le_bytes());
            encoded[n - 4..].copy_from_slice(&(635 + (i as u32 - 3) * 1000).to_le_bytes());
        }
        sites.extend(encoded);
    }
    f.extend((sites.len() as u32).to_le_bytes());
    f.extend(sites);
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    for (name, bytes) in [("f", f), ("o", all.concat())] {
        file.create_stream(format!("/{name}")).unwrap().write_all(&bytes).unwrap();
    }
    let bytes = file.into_inner().into_inner();
    let mut companion = (bytes.len() as u32).to_le_bytes().to_vec();
    companion.extend(bytes);
    companion
}

/// Original MorphData TextBox packet; no Office-produced resource is adopted.
pub fn text_companion(flags: u32, password: u16) -> Vec<u8> {
    let mut controls = controls();
    let mut body = u32s(&[0x80400301, 0, flags]);
    body.extend(password.to_le_bytes());
    body.extend([0, 0]);
    body.extend(u32s(&[0x80000007, 2540, 635]));
    body.extend(string("initial"));
    let mut text = record(0x200, &body);
    text.extend(record(0x200, &[0; 4]));
    controls[1] = text;
    let mut streams = streams();
    let root = &mut streams[0].1;
    let block_end = 4 + u16::from_le_bytes([root[2], root[3]]) as usize;
    let first = block_end + 2 + 4 + 4 + 4;
    let entry = first + 4 + u16::from_le_bytes([root[first + 2], root[first + 3]]) as usize;
    root[entry + 20..entry + 24].copy_from_slice(&(controls[1].len() as u32).to_le_bytes());
    streams[1].1 = controls.concat();
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    for (name, bytes) in streams {
        file.create_stream(format!("/{name}")).unwrap().write_all(&bytes).unwrap();
    }
    let bytes = file.into_inner().into_inner();
    let mut out = (bytes.len() as u32).to_le_bytes().to_vec();
    out.extend(bytes);
    out
}

/// Independent native numeric records; these controls have no TextProps tail.
pub fn numeric_companion() -> Vec<u8> {
    let mut all = controls();
    all.push(record(0x200, &u32s(&[0x4e8, 0, 10, 4, 2, 600, 1200])));
    all.push(record(0x200, &u32s(&[0x18e8, 10, (-10i32) as u32, 0, 2, 5, 1800, 600])));
    let mut root = u32s(&[(1 << 10) | (1 << 11) | (1 << 19) | (1 << 27), 0x8000000b, 32000, 10583, 7056, 10583, 7056]);
    root.extend(string("Native form"));
    let mut f = record(0x400, &root);
    f.extend(0u16.to_le_bytes());
    f.extend(5u32.to_le_bytes());
    let mut sites = vec![0, 0x85, 1, 0];
    for (((id, name), class), (tab, control)) in [(11, "Heading"), (12, "Entry"), (13, "Apply"), (14, "Spin1"), (15, "Scroll1")]
        .into_iter()
        .zip([21, 23, 17, 16, 47])
        .zip(all.iter().enumerate())
    {
        sites.extend(site(id, name, class, tab as u16, control.len()));
    }
    f.extend((sites.len() as u32).to_le_bytes());
    f.extend(sites);
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    file.create_stream("/f").unwrap().write_all(&f).unwrap();
    file.create_stream("/o").unwrap().write_all(&all.concat()).unwrap();
    let bytes = file.into_inner().into_inner();
    let mut out = (bytes.len() as u32).to_le_bytes().to_vec();
    out.extend(bytes);
    out
}

/// Original unbound ListBox; selected native column fields, no external list/binding data.
pub fn list_companion() -> Vec<u8> {
    let mut all = controls();
    let mut body = u32s(&[0x80003940, 0]);
    body.push(2); // native fmDisplayStyleList
    body.push(0); // alignment before BoundColumn
    body.extend(2u16.to_le_bytes());
    body.extend(1i16.to_le_bytes());
    body.extend(2i16.to_le_bytes());
    body.extend(u32s(&[6000, 1800]));
    let mut list = record(0x200, &body);
    list.extend(record(0x200, &[0; 4]));
    all.push(list);
    let mut root = u32s(&[(1 << 10) | (1 << 11) | (1 << 19) | (1 << 27), 0x8000000b, 32000, 10583, 7056, 10583, 7056]);
    root.extend(string("Native form"));
    let mut f = record(0x400, &root);
    f.extend(0u16.to_le_bytes());
    f.extend(4u32.to_le_bytes());
    let mut sites = vec![0, 0x84, 1, 0];
    for (((id, name), class), (tab, control)) in
        [(11, "Heading"), (12, "Entry"), (13, "Apply"), (14, "List1")].into_iter().zip([21, 23, 17, 24]).zip(all.iter().enumerate())
    {
        sites.extend(site(id, name, class, tab as u16, control.len()));
    }
    f.extend((sites.len() as u32).to_le_bytes());
    f.extend(sites);
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    file.create_stream("/f").unwrap().write_all(&f).unwrap();
    file.create_stream("/o").unwrap().write_all(&all.concat()).unwrap();
    let bytes = file.into_inner().into_inner();
    let mut out = (bytes.len() as u32).to_le_bytes().to_vec();
    out.extend(bytes);
    out
}

/// Independent ComboBox records: fmDisplayStyleCombo/DropList, shared list columns.
pub fn combo_companion(style: u8, matching: u8) -> Vec<u8> {
    let mut all = controls();
    let flags = if style == 3 { 0x2c80481b } else { 0x2c80081b };
    let mut body = u32s(&[0x80017941, 0, flags]);
    body.extend([style, 0]);
    body.extend(2u16.to_le_bytes());
    body.extend(1i16.to_le_bytes());
    body.extend(2i16.to_le_bytes());
    body.extend(5u16.to_le_bytes());
    body.extend([matching, 0]);
    body.extend(u32s(&[6000, 635]));
    let mut combo = record(0x200, &body);
    combo.extend(record(0x200, &[0; 4]));
    all.push(combo);
    let mut root = u32s(&[(1 << 10) | (1 << 11) | (1 << 19) | (1 << 27), 0x8000000b, 32000, 10583, 7056, 10583, 7056]);
    root.extend(string("Native form"));
    let mut f = record(0x400, &root);
    f.extend(0u16.to_le_bytes());
    f.extend(4u32.to_le_bytes());
    let mut sites = vec![0, 0x84, 1, 0];
    for (((id, name), class), (tab, control)) in
        [(11, "Heading"), (12, "Entry"), (13, "Apply"), (14, "Combo1")].into_iter().zip([21, 23, 17, 25]).zip(all.iter().enumerate())
    {
        sites.extend(site(id, name, class, tab as u16, control.len()));
    }
    f.extend((sites.len() as u32).to_le_bytes());
    f.extend(sites);
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    file.create_stream("/f").unwrap().write_all(&f).unwrap();
    file.create_stream("/o").unwrap().write_all(&all.concat()).unwrap();
    let bytes = file.into_inner().into_inner();
    let mut out = (bytes.len() as u32).to_le_bytes().to_vec();
    out.extend(bytes);
    out
}

/// Original native site metadata; no Office-produced resources.
pub fn metadata_companion() -> Vec<u8> {
    let controls = controls();
    let (_, mut f) = streams().remove(0);
    let header = 4 + usize::from(u16::from_le_bytes([f[2], f[3]])) + 2;
    f.truncate(header);
    f.extend(3u32.to_le_bytes());
    let mut sites = vec![0, 0x83, 1, 0];
    for (index, (id, name, class)) in [(11, "Heading", 21), (12, "Entry", 23), (13, "Apply", 17)].into_iter().enumerate() {
        let tag = format!("tag-{id}");
        let tip = format!("tip-{id}");
        let mut b = u32s(&[0x9ff, 0x80000000 | name.len() as u32, 0x80000000 | tag.len() as u32, id, 100 + id, 0x13, controls[index].len() as u32]);
        b.extend((index as u16).to_le_bytes());
        b.extend((class as u16).to_le_bytes());
        b.extend((0x80000000 | tip.len() as u32).to_le_bytes());
        b.extend(string(name));
        b.extend(string(&tag));
        b.extend(u32s(&[423, 635 + (id - 11) * 1000]));
        b.extend(string(&tip));
        sites.extend(record(0, &b));
    }
    f.extend((sites.len() as u32).to_le_bytes());
    f.extend(sites);
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    for (name, bytes) in [("f", f), ("o", controls.concat())] {
        file.create_stream(format!("/{name}")).unwrap().write_all(&bytes).unwrap();
    }
    let bytes = file.into_inner().into_inner();
    let mut out = (bytes.len() as u32).to_le_bytes().to_vec();
    out.extend(bytes);
    out
}
