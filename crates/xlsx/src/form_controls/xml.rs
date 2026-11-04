//! Strict, namespace-aware bounded XML for the native control metadata reader.
use super::{MAX_FIELD, MC, X14, bad, limited};
use crate::IoError;
use quick_xml::{
    NsReader, XmlVersion,
    events::Event,
    name::{QName, ResolveResult},
};
use std::collections::BTreeMap;
const MAX_BYTES: usize = 4 * 1024 * 1024;
const MAX_NODES: usize = 8192;
const MAX_DEPTH: usize = 32;
#[derive(Default)]
pub(super) struct Budget {
    input: usize,
    nodes: usize,
    stored: usize,
}
impl Budget {
    fn string(&mut self, value: &str) -> Result<String, IoError> {
        if value.chars().any(|c| !matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
            return Err(bad("invalid XML 1.0 character"));
        }
        if value.len() > MAX_FIELD || self.stored.saturating_add(value.len()) > MAX_BYTES {
            return Err(limited("native control XML text/attribute budget"));
        }
        self.stored += value.len();
        Ok(value.into())
    }
}
#[derive(Debug)]
enum Content {
    Text(String),
    Child(Node),
}
#[derive(Debug)]
pub(super) struct Node {
    pub start: usize,
    pub open_end: usize,
    pub end_start: usize,
    pub end: usize,
    pub ns: String,
    pub name: String,
    attrs: BTreeMap<(String, String), String>,
    content: Vec<Content>,
    supported_choice: bool,
}
impl Node {
    pub fn attr(&self, name: &str) -> Option<&str> {
        self.attr_ns("", name)
    }
    pub fn attr_ns(&self, ns: &str, name: &str) -> Option<&str> {
        self.attrs.get(&(ns.into(), name.into())).map(String::as_str)
    }
    pub fn children(&self) -> impl Iterator<Item = &Node> {
        self.content.iter().filter_map(|c| match c {
            Content::Child(n) => Some(n),
            Content::Text(_) => None,
        })
    }
    pub fn child(&self, ns: &str, name: &str) -> Result<Option<&Node>, IoError> {
        let mut matches = self.children().filter(|n| n.ns == ns && n.name == name);
        let first = matches.next();
        if matches.next().is_some() {
            return Err(bad("duplicate native control element"));
        }
        Ok(first)
    }
    pub fn text(&self) -> Result<String, IoError> {
        fn append(node: &Node, out: &mut String) -> Result<(), IoError> {
            for c in &node.content {
                match c {
                    Content::Text(t) => {
                        if out.len().saturating_add(t.len()) > MAX_FIELD {
                            return Err(limited("native control field text"));
                        }
                        out.push_str(t);
                    }
                    Content::Child(n) => append(n, out)?,
                }
            }
            Ok(())
        }
        let mut result = String::new();
        append(self, &mut result)?;
        Ok(result)
    }
    pub fn collect<'a>(&'a self, ns: &str, name: &str, out: &mut Vec<&'a Node>) -> Result<(), IoError> {
        if self.ns == ns && self.name == name {
            if out.len() >= super::MAX_CONTROLS {
                return Err(limited("native control/shape count"));
            }
            out.push(self);
        }
        if self.ns == MC && self.name == "AlternateContent" {
            let chosen = self.children().find(|n| n.ns == MC && n.name == "Choice" && n.supported_choice).or(self.child(MC, "Fallback")?);
            return chosen.ok_or_else(|| bad("unsupported native control markup-compatibility branch"))?.collect(ns, name, out);
        }
        for c in self.children() {
            c.collect(ns, name, out)?;
        }
        Ok(())
    }
}
fn namespace(value: ResolveResult<'_>) -> Result<String, IoError> {
    match value {
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Bound(ns) => {
            let text = std::str::from_utf8(ns.as_ref()).map_err(|_| bad("invalid XML namespace"))?;
            if text.len() > MAX_FIELD {
                return Err(limited("native control namespace"));
            }
            Ok(text.into())
        }
        ResolveResult::Unknown(_) => Err(bad("undeclared XML namespace prefix")),
    }
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, node: Node) -> Result<(), IoError> {
    if let Some(parent) = stack.last_mut() {
        parent.content.push(Content::Child(node));
    } else if root.replace(node).is_some() {
        return Err(bad("multiple XML roots"));
    }
    Ok(())
}
pub(super) fn parse(bytes: &[u8], budget: &mut Budget) -> Result<Node, IoError> {
    if budget.input.saturating_add(bytes.len()) > MAX_BYTES {
        return Err(limited("native control XML input"));
    }
    budget.input += bytes.len();
    let input = std::str::from_utf8(bytes).map_err(|_| bad("native control XML must be UTF-8 in this profile"))?;
    let mut reader = NsReader::from_str(input);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| bad(e.to_string()))?;
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                budget.nodes = budget.nodes.saturating_add(1);
                if budget.nodes > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(limited("native control XML nodes/depth"));
                }
                let (ns, local) = reader.resolver().resolve_element(e.name());
                let ns = budget.string(&namespace(ns)?)?;
                let name = budget.string(std::str::from_utf8(local.as_ref()).map_err(|_| bad("invalid XML name"))?)?;
                let end = reader.buffer_position() as usize;
                let mut node =
                    Node { start, open_end: end, end_start: end, end, ns, name, attrs: BTreeMap::new(), content: vec![], supported_choice: false };
                for (count, attribute) in e.attributes().enumerate() {
                    if count >= 64 {
                        return Err(limited("native control XML attribute count"));
                    }
                    let a = attribute.map_err(|e| bad(e.to_string()))?;
                    if a.key.as_ref() == b"xmlns" || a.key.as_ref().starts_with(b"xmlns:") {
                        continue;
                    }
                    let (ns, local) = reader.resolver().resolve_attribute(a.key);
                    let key = (
                        budget.string(&namespace(ns)?)?,
                        budget.string(std::str::from_utf8(local.as_ref()).map_err(|_| bad("invalid XML attribute"))?)?,
                    );
                    let value = a.decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder()).map_err(|e| bad(e.to_string()))?;
                    if node.attrs.insert(key, budget.string(&value)?).is_some() {
                        return Err(bad("duplicate expanded XML attribute"));
                    }
                }
                if node.ns == MC
                    && node.name == "Choice"
                    && let Some(requires) = node.attr("Requires").filter(|s| !s.trim().is_empty())
                {
                    let mut supported = true;
                    for prefix in requires.split_whitespace() {
                        let qualified = format!("{prefix}:feature");
                        let (ns, _) = reader.resolver().resolve_element(QName(qualified.as_bytes()));
                        if namespace(ns)? != X14 {
                            supported = false;
                        }
                    }
                    node.supported_choice = supported;
                }
                if matches!(event, Event::Empty(_)) {
                    attach(&mut stack, &mut root, node)?;
                } else {
                    stack.push(node);
                }
            }
            Event::End(_) => {
                let mut node = stack.pop().ok_or_else(|| bad("unmatched XML end"))?;
                node.end_start = start;
                node.end = reader.buffer_position() as usize;
                attach(&mut stack, &mut root, node)?;
            }
            Event::Text(t) => {
                let text = t.xml_content(XmlVersion::Explicit1_0).map_err(|e| bad(e.to_string()))?;
                push_text(&mut stack, &text, budget)?;
            }
            Event::CData(t) => {
                let text = std::str::from_utf8(t.as_ref()).map_err(|_| bad("invalid XML CDATA"))?;
                push_text(&mut stack, text, budget)?;
            }
            Event::GeneralRef(r) => {
                let text = if let Some(c) = r.resolve_char_ref().map_err(|e| bad(e.to_string()))? {
                    c.to_string()
                } else {
                    let name = std::str::from_utf8(r.as_ref()).map_err(|_| bad("invalid XML reference"))?;
                    quick_xml::escape::resolve_predefined_entity(name).ok_or_else(|| bad("unknown XML entity"))?.into()
                };
                push_text(&mut stack, &text, budget)?;
            }
            Event::DocType(_) => return Err(bad("DTD is not supported for native controls")),
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err(bad("truncated native control XML"));
    }
    root.ok_or_else(|| bad("missing native control XML root"))
}
fn push_text(stack: &mut [Node], text: &str, budget: &mut Budget) -> Result<(), IoError> {
    if let Some(parent) = stack.last_mut() {
        let value = budget.string(text)?;
        if let Some(Content::Text(previous)) = parent.content.last_mut() {
            if previous.len().saturating_add(value.len()) > MAX_FIELD {
                return Err(limited("native control text field"));
            }
            previous.push_str(&value);
        } else {
            parent.content.push(Content::Text(value));
        }
    } else if !text.trim().is_empty() {
        return Err(bad("text outside XML root"));
    }
    Ok(())
}
