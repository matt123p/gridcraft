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
    let label = leaf(false, "Authored label");
    let mut b = u32s(&[0x80400100, 0, 0x80000000 | 7, 2540, 635]);
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
    let controls = controls();
    let mut b = u32s(&[(1 << 10) | (1 << 11) | (1 << 19) | (1 << 27), 0x80000000 | 11, 32000, 10583, 7056, 10583, 7056]);
    b.extend(string("Native form"));
    let mut f = record(0x400, &b);
    f.extend(0u16.to_le_bytes()); // empty class table, all three types use native cache indices
    f.extend(3u32.to_le_bytes());
    let mut sites = vec![0, 0x83, 1, 0]; // three immediate children, padded to four bytes
    for (((id, name), class), (tab, c)) in
        [(11, "Heading"), (12, "Entry"), (13, "Apply")].into_iter().zip([21, 23, 17]).zip(controls.iter().enumerate())
    {
        sites.extend(site(id, name, class, tab as u16, c.len()));
    }
    f.extend((sites.len() as u32).to_le_bytes());
    f.extend(sites);
    vec![("f".into(), f), ("o".into(), controls.concat())]
}
pub fn companion() -> Vec<u8> {
    let mut file = cfb::CompoundFile::create(Cursor::new(Vec::new())).unwrap();
    for (name, bytes) in streams() {
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
