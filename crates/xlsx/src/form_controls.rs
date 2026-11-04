//! Inert native worksheet Button metadata. No rendering, geometry conversion or VBA launch.
use crate::{IoError, package::normalize};
use gridcraft_model::{PreservedPart, PreservedRel, Sheet};
use std::collections::{BTreeMap, BTreeSet};
mod xml;
use xml::{Budget, Node};
const SHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CTRL_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
const CTRL_CT: &str = "application/vnd.ms-excel.controlproperties+xml";
const VML: &str = "urn:schemas-microsoft-com:vml";
const OFFICE: &str = "urn:schemas-microsoft-com:office:office";
const EXCEL: &str = "urn:schemas-microsoft-com:office:excel";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MAX_CONTROLS: usize = 1024;
const MAX_FIELD: usize = 16384;
fn bad(message: impl Into<String>) -> IoError {
    IoError::Format(format!("native worksheet controls: {}", message.into()))
}
fn limited(message: &str) -> IoError {
    IoError::TooLarge(message.into())
}
/// Original worksheet control identity and decoded Button metadata. This snapshot is not
/// writable persistence state and does not grant execution permission. Geometry/style are raw.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct WorksheetButton {
    pub shape_id: u32,
    pub name: String,
    pub caption_text: String,
    pub macro_assignment: Option<String>,
    pub vml_anchor: Option<String>,
    pub vml_style: Option<String>,
}
/// A native control retained by the package layer but outside this decoder's Button slice.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct UnsupportedWorksheetControl {
    pub shape_id: u32,
    pub name: String,
    pub object_type: String,
}
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct WorksheetControlMetadata {
    pub buttons: Vec<WorksheetButton>,
    pub unsupported: Vec<UnsupportedWorksheetControl>,
}
/// Decode the sheet's native control/VML/property graph without changing it or executing code.
/// Each call is bounded to 4 MiB XML input and stored strings, 8192 nodes, depth 32, 1024
/// controls/shapes and 16 KiB per field. Invalid/ambiguous graphs return no partial snapshot.
/// Use a fresh snapshot after document changes; native controls are not yet rendered/activated.
pub fn read_worksheet_controls(sheet: &Sheet) -> Result<WorksheetControlMetadata, IoError> {
    let Some(kept) = sheet.preserved.as_deref() else {
        return Ok(WorksheetControlMetadata::default());
    };
    if kept.parts.len() > 4096 || kept.rels.len() > 10000 || kept.elements.len() > MAX_CONTROLS {
        return Err(limited("native control package graph count"));
    }
    let fragments: Vec<_> = kept.elements.iter().filter(|e| e.slot == "controls").collect();
    if fragments.is_empty() {
        return Ok(WorksheetControlMetadata::default());
    }
    if fragments.len() != 1 {
        return Err(bad("ambiguous worksheet controls fragments"));
    }
    let mut budget = Budget::default();
    let mut metadata_bytes = 0usize;
    let mut relationships = BTreeMap::new();
    for rel in &kept.rels {
        metadata_bytes = metadata_bytes.saturating_add(rel.id.len()).saturating_add(rel.target.len()).saturating_add(rel.rel_type.len());
        if metadata_bytes > 4 * 1024 * 1024 {
            return Err(limited("native control graph metadata"));
        }
        if rel.id.len() > MAX_FIELD || rel.target.len() > MAX_FIELD || rel.rel_type.len() > MAX_FIELD {
            return Err(limited("native control relationship field"));
        }
        if relationships.insert(rel.id.as_str(), rel).is_some() {
            return Err(bad("duplicate preserved relationship ID"));
        }
    }
    let mut parts = BTreeMap::new();
    for part in &kept.parts {
        metadata_bytes = metadata_bytes.saturating_add(part.name.len()).saturating_add(part.content_type.len());
        if metadata_bytes > 4 * 1024 * 1024 {
            return Err(limited("native control graph metadata"));
        }
        if part.name.len() > MAX_FIELD {
            return Err(limited("native control part name"));
        }
        if parts.insert(normalize(&part.name), part).is_some() {
            return Err(bad("duplicate preserved part name"));
        }
    }
    let controls = xml::parse(fragments.first().ok_or_else(|| bad("missing controls"))?.xml.as_bytes(), &mut budget)?;
    if !((controls.ns == SHEET && controls.name == "controls") || (controls.ns == MC && controls.name == "AlternateContent")) {
        return Err(bad("wrong worksheet controls root/namespace"));
    }
    let vml = kept.vml.as_ref().ok_or_else(|| bad("missing preserved VML"))?;
    let drawing = xml::parse(&vml.data, &mut budget)?;
    if drawing.name != "xml" || !(drawing.ns.is_empty() || drawing.ns == VML) {
        return Err(bad("wrong legacy VML root"));
    }
    let shapes = shape_map(&drawing)?;
    let mut nodes = Vec::new();
    controls.collect(SHEET, "control", &mut nodes)?;
    if nodes.is_empty() {
        return Err(bad("no SpreadsheetML controls in retained fragment"));
    }
    let mut ids = BTreeSet::new();
    let mut names = BTreeSet::new();
    let mut out = WorksheetControlMetadata::default();
    for node in nodes {
        let id = positive(node.attr("shapeId").ok_or_else(|| bad("missing shapeId"))?)?;
        let name = node.attr("name").filter(|n| !n.is_empty()).ok_or_else(|| bad("missing control name"))?;
        if !ids.insert(id) || !names.insert(name.to_ascii_lowercase()) {
            return Err(bad("duplicate control identity"));
        }
        let shape = shapes.get(&id).ok_or_else(|| bad("control has no matching VML shape"))?;
        let client = shape.child(EXCEL, "ClientData")?.ok_or_else(|| bad("missing VML ClientData"))?;
        let vml_type = client.attr("ObjectType").ok_or_else(|| bad("missing VML ObjectType"))?;
        let props = control_properties(&relationships, &parts, node)?;
        let props = xml::parse(&props.data, &mut budget)?;
        if props.ns != X14 || props.name != "formControlPr" {
            return Err(bad("wrong control-properties root/namespace"));
        }
        let kind = props.attr("objectType").unwrap_or(vml_type);
        if !kind.eq_ignore_ascii_case(vml_type) {
            return Err(bad("conflicting native control types"));
        }
        if kind != "Button" || vml_type != "Button" {
            out.unsupported.push(UnsupportedWorksheetControl { shape_id: id, name: name.into(), object_type: kind.into() });
            continue;
        }
        let legacy = optional_text(client.child(EXCEL, "FmlaMacro")?)?;
        let properties = node.child(SHEET, "controlPr")?;
        let current = properties.and_then(|n| n.attr("macro")).filter(|s| !s.trim().is_empty());
        if let (Some(a), Some(b)) = (legacy.as_deref(), current)
            && a != b
        {
            return Err(bad("conflicting native macro assignments"));
        }
        out.buttons.push(WorksheetButton {
            shape_id: id,
            name: name.into(),
            caption_text: shape.child(VML, "textbox")?.map(Node::text).transpose()?.unwrap_or_default(),
            macro_assignment: current.map(str::to_string).or(legacy),
            vml_anchor: optional_text(client.child(EXCEL, "Anchor")?)?,
            vml_style: shape.attr("style").map(str::to_string),
        });
    }
    Ok(out)
}
fn positive(value: &str) -> Result<u32, IoError> {
    value.parse::<u32>().ok().filter(|n| *n > 0).ok_or_else(|| bad("invalid native shape ID"))
}
fn optional_text(node: Option<&Node>) -> Result<Option<String>, IoError> {
    Ok(node.map(Node::text).transpose()?.filter(|s| !s.trim().is_empty()))
}
fn shape_map(root: &Node) -> Result<BTreeMap<u32, &Node>, IoError> {
    let mut nodes = Vec::new();
    root.collect(VML, "shape", &mut nodes)?;
    let mut out = BTreeMap::new();
    for shape in nodes {
        let id = shape.attr("id").and_then(|v| v.strip_prefix("_x0000_s")).map(positive).transpose()?;
        let spid = shape.attr_ns(OFFICE, "spid").map(|v| positive(v.strip_prefix("_x0000_s").ok_or_else(|| bad("invalid VML spid"))?)).transpose()?;
        if let (Some(a), Some(b)) = (id, spid)
            && a != b
        {
            return Err(bad("conflicting VML shape identities"));
        }
        if let Some(id) = spid.or(id) {
            if out.insert(id, shape).is_some() {
                return Err(bad("duplicate VML shape ID"));
            }
        } else if shape.child(EXCEL, "ClientData")?.is_some() {
            return Err(bad("missing native VML shape ID"));
        }
    }
    Ok(out)
}
fn control_properties<'a>(
    relationships: &BTreeMap<&str, &PreservedRel>,
    parts: &BTreeMap<String, &'a PreservedPart>,
    node: &Node,
) -> Result<&'a PreservedPart, IoError> {
    let rid = node.attr_ns(REL, "id").ok_or_else(|| bad("missing control relationship ID"))?;
    let rel = relationships.get(rid).ok_or_else(|| bad("missing control relationship"))?;
    if rel.external || rel.rel_type != CTRL_REL {
        return Err(bad("invalid control-properties relationship"));
    }
    let part = parts.get(&normalize(&rel.target)).ok_or_else(|| bad("missing control-properties part"))?;
    if part.content_type != CTRL_CT {
        return Err(bad("wrong control-properties content type"));
    }
    Ok(part)
}
