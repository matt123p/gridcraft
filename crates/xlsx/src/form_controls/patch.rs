//! Bounded edits of known XML nodes. Bytes outside an edited node are retained verbatim.
use super::{BTreeSet, IoError, Node, bad, limited};
type Patches = Vec<(usize, usize, String)>;
use crate::xml::esc_attr;
use quick_xml::{Reader, events::Event};
pub(super) fn qname(source: &str, node: &Node) -> Result<String, IoError> {
    let open = source.get(node.start..node.open_end).ok_or_else(|| bad("XML node span"))?;
    let mut r = Reader::from_str(open);
    match r.read_event().map_err(|e| bad(e.to_string()))? {
        Event::Start(e) | Event::Empty(e) => String::from_utf8(e.name().as_ref().into()).map_err(|_| bad("XML QName")),
        _ => Err(bad("XML opening node")),
    }
}
pub(super) fn apply(source: &str, mut patches: Vec<(usize, usize, String)>) -> Result<String, IoError> {
    patches.sort_by_key(|a| std::cmp::Reverse(a.0));
    let mut out = source.to_owned();
    let mut previous = source.len();
    for (start, end, text) in patches {
        if end > previous || source.get(start..end).is_none() {
            return Err(bad("overlapping XML patches"));
        }
        if out.len().saturating_add(text.len()).saturating_sub(end - start) > 4 * 1024 * 1024 {
            return Err(limited("native control edited XML"));
        }
        out.replace_range(start..end, &text);
        previous = start;
    }
    Ok(out)
}
pub(super) fn element(
    source: &str,
    node: &Node,
    attributes: &[(&str, String)],
    children: Vec<(usize, usize, String)>,
    append: &str,
) -> Result<String, IoError> {
    element_removing(source, node, attributes, children, append, &[])
}
pub(super) fn element_removing(
    source: &str,
    node: &Node,
    attributes: &[(&str, String)],
    children: Patches,
    append: &str,
    remove: &[&str],
) -> Result<String, IoError> {
    let original = source.get(node.start..node.end).ok_or_else(|| bad("XML source span"))?;
    let open = source.get(node.start..node.open_end).ok_or_else(|| bad("XML source opening span"))?;
    let mut r = Reader::from_str(open);
    let event = r.read_event().map_err(|e| bad(e.to_string()))?;
    let (e, empty) = match event {
        Event::Start(e) => (e, false),
        Event::Empty(e) => (e, true),
        _ => return Err(bad("XML opening element")),
    };
    let name = std::str::from_utf8(e.name().as_ref()).map_err(|_| bad("XML QName"))?.to_owned();
    let mut start = format!("<{name}");
    let mut seen = BTreeSet::new();
    for attr in e.attributes() {
        let attr = attr.map_err(|e| bad(e.to_string()))?;
        let key = std::str::from_utf8(attr.key.as_ref()).map_err(|_| bad("attribute name"))?;
        if remove.contains(&key) {
            continue;
        }
        let old = attr.decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, r.decoder()).map_err(|e| bad(e.to_string()))?;
        let value = attributes.iter().find(|(k, _)| *k == key).map(|(_, v)| v.as_str()).unwrap_or(&old);
        seen.insert(key.to_owned());
        start.push_str(&format!(" {key}=\"{}\"", esc_attr(value)));
    }
    for (key, value) in attributes {
        if !seen.contains(*key) {
            start.push_str(&format!(" {key}=\"{}\"", esc_attr(value)));
        }
    }
    if empty && append.is_empty() {
        start.push_str("/>");
        return Ok(start);
    }
    start.push('>');
    if empty {
        return Ok(format!("{start}{append}</{name}>"));
    }
    let patches = children.into_iter().map(|(s, e, t)| (s - node.start, e - node.start, t)).collect();
    let mut out = apply(original, patches)?;
    let close_start = node.end_start - node.start;
    // Descendant replacements can change the closing offset: find the retained final closing tag.
    let close = out.rfind("</").unwrap_or(close_start);
    out.insert_str(close, append);
    out.replace_range(..node.open_end - node.start, &start);
    Ok(out)
}
pub(super) fn child_changes(source: &str, parent: &Node, ns: &str, values: &[(&str, String)]) -> Result<(Patches, String), IoError> {
    let name = qname(source, parent)?;
    let prefix = name.rsplit_once(':').map(|(p, _)| format!("{p}:")).unwrap_or_default();
    let mut patches = vec![];
    let mut append = String::new();
    for (key, value) in values {
        let text = crate::xml::esc(value);
        if let Some(child) = parent.child(ns, key)? {
            let q = qname(source, child)?;
            patches.push((child.start, child.end, format!("<{q}>{text}</{q}>")));
        } else {
            append.push_str(&format!("<{prefix}{key}>{text}</{prefix}{key}>"));
        }
    }
    Ok((patches, append))
}
