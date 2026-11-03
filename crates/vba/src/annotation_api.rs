//! Local annotations project native command/model behavior; no remote identities are invented.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, MAX_OBJECTS, Object, Session, Value, failure, values};
use gridcraft_engine::core::CellRef;
use serde_json::json;

const PROPS: [(&str, &str, i32); 6] = [
    ("Title", "title", 1),
    ("Subject", "subject", 2),
    ("Author", "author", 3),
    ("Keywords", "keywords", 4),
    ("Comments", "description", 5),
    ("Company", "company", 21),
];
fn node(kind: &'static str, sheet: usize, cell: CellRef) -> Feature {
    let mut f = Feature::new(kind, Some(sheet));
    f.range = Some(gridcraft_engine::core::RangeRef::new(cell, cell).into());
    f
}
fn execute(s: &mut Session, id: &str, p: serde_json::Value) -> HostResult<()> {
    // Native annotation maps are separate from sparse cell content, so the cell-diff
    // protection check cannot guard them. The native protection model has no
    // independent annotation-edit or insert-hyperlink permission.
    if (id.starts_with("review.") || id == "insert.link")
        && let Some(si) = p.get("sheet").and_then(serde_json::Value::as_u64)
    {
        let si = usize::try_from(si).map_err(|_| failure(91, "worksheet missing"))?;
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if sh.is_protected() {
            return Err(failure(1004, "sheet protection forbids annotation edits"));
        }
    }
    s.execute(id, p).map_err(|e| failure(1004, e.to_string()))?;
    Ok(())
}
fn text_edit(old: &str, c: &HostCall) -> HostResult<Option<String>> {
    let a = arguments(c, &["Text", "Start", "Overwrite"])?;
    if absent(&a[0]) {
        if !absent(&a[1]) || !absent(&a[2]) {
            return Err(failure(449, "Text required for text editing"));
        }
        return Ok(None);
    }
    let text = string(&a[0])?;
    if text.len().saturating_add(old.len()) > 1_048_576 {
        return Err(failure(7, "annotation text budget"));
    }
    if absent(&a[1]) {
        return Ok(Some(text));
    }
    let start = numeric_index(&a[1])?;
    let mut units: Vec<u16> = old.encode_utf16().collect();
    let start = usize::try_from(start.checked_sub(1).ok_or_else(|| failure(5, "text start"))?).map_err(|_| failure(5, "text start"))?;
    if start > units.len() {
        return Err(failure(5, "text start"));
    }
    let added: Vec<u16> = text.encode_utf16().collect();
    let overwrite = absent(&a[2]) || super::address::boolean(&a[2])?;
    let end = if overwrite { start.saturating_add(added.len()).min(units.len()) } else { start };
    units.splice(start..end, added);
    Ok(Some(String::from_utf16(&units).map_err(|_| failure(5, "text edit splits a Unicode character"))?))
}
impl Bridge {
    pub(crate) fn annotation_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        let m = c.member.to_ascii_lowercase();
        let root = match o {
            Object::Sheet(si) if reads(c) && matches!(m.as_str(), "comments" | "commentsthreaded" | "hyperlinks") => Some(Feature::new(
                match m.as_str() {
                    "comments" => "Comments",
                    "commentsthreaded" => "CommentsThreaded",
                    _ => "Hyperlinks",
                },
                Some(*si),
            )),
            Object::Range(si, r)
                if reads(c)
                    && matches!(m.as_str(), "comment" | "commentthreaded" | "addcomment" | "addcommentthreaded" | "hyperlinks" | "notetext") =>
            {
                return Some((|| {
                    let cell = r.rectangle()?.start;
                    if m == "notetext" {
                        return self.note_text(s, *si, cell, c);
                    }
                    if m == "hyperlinks" {
                        let mut f = Feature::new("Hyperlinks", Some(*si));
                        f.range = Some(r.clone());
                        return if c.arguments.is_empty() { self.feature(f) } else { self.annotation_dispatch(s, &f, c) };
                    }
                    let threaded = m.ends_with("threaded");
                    let kind = if threaded { "CommentThreaded" } else { "Comment" };
                    let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(*si).ok_or_else(|| failure(91, "worksheet deleted"))?;
                    if m.starts_with("add") {
                        let a = arguments(c, &["Text"])?;
                        if sh.comments.contains_key(&cell) {
                            return Err(failure(1004, "cell already has an annotation"));
                        }
                        let text = if threaded {
                            string(required(&a, 0)?)?
                        } else if absent(&a[0]) {
                            String::new()
                        } else {
                            string(&a[0])?
                        };
                        if text.len() > 1_048_576 {
                            return Err(failure(7, "annotation text budget"));
                        }
                        self.objects.reserve(1)?;
                        execute(s, if threaded { "review.newComment" } else { "review.newNote" }, json!({"sheet":si,"cell":cell.a1(),"text":text}))?;
                        self.feature(node(kind, *si, cell))
                    } else {
                        arguments(c, &[])?;
                        if sh.comments.get(&cell).is_some_and(|v| v.threaded == threaded) {
                            self.feature(node(kind, *si, cell))
                        } else {
                            Ok(Value::Nothing)
                        }
                    }
                })());
            }
            Object::Workbook if m == "builtindocumentproperties" && reads(c) => Some(Feature::new("DocumentProperties", None)),
            _ => None,
        };
        if let Some(f) = root {
            return Some(if c.arguments.is_empty() { self.feature(f) } else { self.annotation_dispatch(s, &f, c) });
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if !matches!(
            f.kind,
            "Comments"
                | "Comment"
                | "CommentsThreaded"
                | "CommentThreaded"
                | "Author"
                | "Hyperlinks"
                | "Hyperlink"
                | "DocumentProperties"
                | "DocumentProperty"
        ) {
            return None;
        }
        Some(self.annotation_dispatch(s, f, c))
    }
    fn annotation_nodes(&self, s: &Session, f: &Feature) -> HostResult<Vec<Feature>> {
        if f.kind == "DocumentProperties" {
            return Ok(PROPS
                .iter()
                .map(|(name, _, i)| {
                    let mut n = Feature::new("DocumentProperty", None);
                    n.text = Some((*name).into());
                    n.part = *i;
                    n
                })
                .collect());
        }
        let si = f.sheet.ok_or_else(|| failure(91, "worksheet missing"))?;
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let mut out = Vec::new();
        if f.label.as_deref() == Some("replies") {
            let cell = f.range.as_ref().ok_or_else(|| failure(91, "comment deleted"))?.rectangle()?.start;
            let entry = sh.comments.get(&cell).filter(|v| v.threaded).ok_or_else(|| failure(91, "comment deleted"))?;
            if entry.replies.len() > MAX_OBJECTS {
                return Err(failure(7, "reply collection budget"));
            }
            for i in 0..entry.replies.len() {
                let mut n = node("CommentThreaded", si, cell);
                n.part = i32::try_from(i + 1).map_err(|_| failure(7, "reply budget"))?;
                out.push(n);
            }
            return Ok(out);
        }
        let within = |cell: CellRef| f.range.as_ref().is_none_or(|r| r.areas.iter().any(|r| r.contains(cell)));
        if f.kind == "Hyperlinks" {
            for cell in sh.hyperlinks.keys().filter(|c| within(**c)) {
                if out.len() >= MAX_OBJECTS {
                    return Err(failure(7, "link collection budget"));
                }
                out.push(node("Hyperlink", si, *cell));
            }
        } else {
            for (cell, _) in sh.comments.iter().filter(|(cell, c)| within(**cell) && c.threaded == (f.kind == "CommentsThreaded")) {
                if out.len() >= MAX_OBJECTS {
                    return Err(failure(7, "comment collection budget"));
                }
                out.push(node(if f.kind == "Comments" { "Comment" } else { "CommentThreaded" }, si, *cell));
            }
        }
        Ok(out)
    }
    fn annotation_dispatch(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let m = c.member.to_ascii_lowercase();
        if f.kind == "DocumentProperty" {
            return self.document_property(s, f, c);
        }
        let collection =
            matches!(f.kind, "Comments" | "CommentsThreaded" | "Hyperlinks" | "DocumentProperties") || f.label.as_deref() == Some("replies");
        if collection {
            if f.kind == "Hyperlinks" && m == "add" && reads(c) {
                return self.add_hyperlink(s, f, c);
            }
            let all = self.annotation_nodes(s, f)?;
            if c.kind == InvokeKind::Enumerate {
                arguments(c, &[])?;
                return self.object_array(all.into_iter().map(Feature::object).collect());
            }
            if m == "delete" && c.kind == InvokeKind::Call && f.kind == "Hyperlinks" {
                arguments(c, &[])?;
                for link in all {
                    let cell = link.range.ok_or_else(|| failure(91, "link missing"))?.rectangle()?.start;
                    execute(s, "insert.link", json!({"sheet":f.sheet,"cell":cell.a1(),"remove":true}))?;
                    self.objects.invalidate_annotation(self.context, f.sheet.ok_or_else(|| failure(91, "worksheet missing"))?, cell, true)?;
                }
                return Ok(Value::Empty);
            }
            if !reads(c) {
                return Err(failure(438, "read-only annotation collection"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(c, &[])?;
                    Ok(Value::Long(all.len() as i32))
                }
                "application" => {
                    arguments(c, &[])?;
                    self.intern(Object::Application)
                }
                "parent" => {
                    arguments(c, &[])?;
                    if f.kind == "DocumentProperties" {
                        self.intern(Object::Workbook)
                    } else if f.label.as_deref() == Some("replies") {
                        let mut n = f.clone();
                        n.label = None;
                        n.kind = "CommentThreaded";
                        self.feature(n)
                    } else {
                        self.intern(Object::Sheet(f.sheet.ok_or_else(|| failure(91, "worksheet missing"))?))
                    }
                }
                "item" | "" | "comments" | "commentsthreaded" | "hyperlinks" | "builtindocumentproperties" | "replies" => {
                    let a = arguments(c, &["Index"])?;
                    let index = required(&a, 0)?;
                    let n = if f.kind == "DocumentProperties" {
                        all.into_iter().find(|f| match index {
                            Value::String(_) => string(index).is_ok_and(|n| f.text.as_ref().is_some_and(|t| t.eq_ignore_ascii_case(&n))),
                            _ => numeric_index(index).is_ok_and(|i| i == f.part),
                        })
                    } else {
                        let i = usize::try_from(numeric_index(index)?.checked_sub(1).ok_or_else(|| failure(9, "annotation index"))?)
                            .map_err(|_| failure(9, "annotation index"))?;
                        all.get(i).cloned()
                    };
                    self.feature(n.ok_or_else(|| failure(9, "annotation index unavailable"))?)
                }
                _ => Err(failure(438, "unsupported annotation collection member")),
            };
        }
        let si = f.sheet.ok_or_else(|| failure(91, "worksheet missing"))?;
        let cell = f.range.as_ref().ok_or_else(|| failure(91, "annotation deleted"))?.rectangle()?.start;
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        if f.kind == "Hyperlink" {
            return self.hyperlink(s, f, c);
        }
        let entry = sh
            .comments
            .get(&cell)
            .filter(|v| v.threaded == matches!(f.kind, "CommentThreaded" | "Author"))
            .ok_or_else(|| failure(91, "annotation deleted"))?;
        let reply = usize::try_from(f.part).map_err(|_| failure(91, "reply deleted"))?;
        let (author, text) = if reply == 0 {
            (&entry.author, &entry.text)
        } else {
            let (a, t) = entry.replies.get(reply - 1).ok_or_else(|| failure(91, "reply deleted"))?;
            (a, t)
        };
        if author.len() > 1_048_576 || text.len() > 1_048_576 || entry.replies.len() > MAX_OBJECTS {
            return Err(failure(7, "native annotation payload budget"));
        }
        if f.kind == "Author" {
            arguments(c, &[])?;
            if !reads(c) {
                return Err(failure(438, "read-only local author"));
            }
            return match m.as_str() {
                "name" => values::vba_text(author),
                "application" => self.intern(Object::Application),
                "providerid" | "userid" => Err(failure(1004, "remote author identity service unavailable")),
                _ => Err(failure(438, "unsupported local author member")),
            };
        }
        if m == "text" && reads(c) {
            if let Some(text) = text_edit(text, c)? {
                if reply != 0 {
                    return Err(failure(1004, "native service cannot edit replies"));
                }
                execute(s, if entry.threaded { "review.newComment" } else { "review.newNote" }, json!({"sheet":si,"cell":cell.a1(),"text":text}))?;
                return values::vba_text(&text);
            }
            return values::vba_text(text);
        }
        if m == "addreply" && reads(c) && entry.threaded && reply == 0 {
            let a = arguments(c, &["Text"])?;
            let text = string(required(&a, 0)?)?;
            if text.len() > 1_048_576 || entry.replies.len() >= MAX_OBJECTS {
                return Err(failure(7, "reply budget"));
            }
            self.objects.reserve(1)?;
            let part = entry.replies.len() as i32 + 1;
            execute(s, "review.replyComment", json!({"sheet":si,"cell":cell.a1(),"text":text}))?;
            let mut n = f.clone();
            n.part = part;
            return self.feature(n);
        }
        if m == "delete" && c.kind == InvokeKind::Call {
            arguments(c, &[])?;
            if reply != 0 {
                return Err(failure(1004, "native service cannot delete individual replies"));
            }
            execute(s, "review.deleteComment", json!({"sheet":si,"range":cell.a1()}))?;
            self.objects.invalidate_annotation(self.context, si, cell, false)?;
            return Ok(Value::Empty);
        }
        if m == "visible" && c.kind == InvokeKind::Let && !entry.threaded {
            let a = arguments(c, &["Value"])?;
            let v = super::address::boolean(required(&a, 0)?)?;
            if v != entry.visible {
                execute(s, "review.showNote", json!({"sheet":si,"cell":cell.a1()}))?;
            }
            return Ok(Value::Empty);
        }
        if !reads(c) {
            return Err(failure(438, "unsupported annotation setter"));
        }
        if m == "replies" && entry.threaded && reply == 0 {
            let mut n = f.clone();
            n.label = Some("replies".into());
            n.kind = "CommentsThreaded";
            return if c.arguments.is_empty() { self.feature(n) } else { self.annotation_dispatch(s, &n, c) };
        }
        arguments(c, &[])?;
        match m.as_str() {
            "application" => self.intern(Object::Application),
            "parent" => {
                if reply == 0 {
                    self.intern(Object::Range(si, gridcraft_engine::core::RangeRef::new(cell, cell).into()))
                } else {
                    let mut n = f.clone();
                    n.part = 0;
                    self.feature(n)
                }
            }
            "author" if !entry.threaded => values::vba_text(author),
            "author" => {
                let mut n = f.clone();
                n.kind = "Author";
                self.feature(n)
            }
            "date" if entry.threaded => Err(failure(1004, "native local comments have no timestamp service")),
            "visible" if !entry.threaded => Ok(Value::Boolean(entry.visible)),
            "next" | "previous" if reply != 0 => {
                let next = if m == "next" { reply.checked_add(1) } else { reply.checked_sub(1) };
                if let Some(i) = next.filter(|i| *i > 0 && *i <= entry.replies.len()) {
                    let mut n = f.clone();
                    n.part = i as i32;
                    self.feature(n)
                } else {
                    Ok(Value::Nothing)
                }
            }
            "next" | "previous" => {
                let all = self.annotation_nodes(s, &Feature::new(if entry.threaded { "CommentsThreaded" } else { "Comments" }, Some(si)))?;
                let at = all.iter().position(|n| n.range == f.range).ok_or_else(|| failure(91, "annotation deleted"))?;
                let next = if m == "next" { at.checked_add(1) } else { at.checked_sub(1) };
                match next.and_then(|i| all.get(i)).cloned() {
                    Some(n) => self.feature(n),
                    None => Ok(Value::Nothing),
                }
            }
            _ => Err(failure(438, "unsupported annotation member")),
        }
    }
    fn note_text(&mut self, s: &mut Session, si: usize, cell: CellRef, c: &HostCall) -> HostResult<Value> {
        let a = arguments(c, &["Text", "Start", "Length"])?;
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let old = sh.comments.get(&cell);
        if old.is_some_and(|v| v.threaded) {
            return Err(failure(1004, "NoteText requires a classic note"));
        }
        if old.is_some_and(|v| v.text.len() > 1_048_576) {
            return Err(failure(7, "note text budget"));
        }
        let mut units: Vec<u16> = old.map(|v| v.text.encode_utf16().collect()).unwrap_or_default();
        let start = if absent(&a[1]) { 1 } else { numeric_index(&a[1])? };
        let start = usize::try_from(start.checked_sub(1).ok_or_else(|| failure(5, "note text start"))?)
            .map_err(|_| failure(5, "note text start"))?
            .min(units.len());
        let length = if absent(&a[2]) { 255 } else { usize::try_from(numeric_index(&a[2])?).map_err(|_| failure(5, "note text length"))? };
        if length > 255 {
            return Err(failure(5, "NoteText length is bounded to 255 UTF-16 units"));
        }
        let end = start.saturating_add(length).min(units.len());
        if absent(&a[0]) {
            let text = String::from_utf16(units.get(start..end).ok_or_else(|| failure(5, "note text bounds"))?)
                .map_err(|_| failure(5, "note text splits Unicode character"))?;
            return values::vba_text(&text);
        }
        let text = string(&a[0])?;
        if text.encode_utf16().count() > 255 {
            return Err(failure(5, "NoteText accepts at most 255 UTF-16 units per call"));
        }
        if units.len().saturating_add(text.len()) > 1_048_576 {
            return Err(failure(7, "note text budget"));
        }
        units.splice(start..end, text.encode_utf16());
        let text = String::from_utf16(&units).map_err(|_| failure(5, "note text splits Unicode character"))?;
        execute(s, "review.newNote", json!({"sheet":si,"cell":cell.a1(),"text":text}))?;
        values::vba_text(&text)
    }
    fn add_hyperlink(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let a = arguments(c, &["Anchor", "Address", "SubAddress", "ScreenTip", "TextToDisplay"])?;
        let Value::Object(reference) = required(&a, 0)? else {
            return Err(failure(13, "Anchor requires a Range"));
        };
        let Object::Document(doc, inner) = self.objects.get(*reference)? else {
            return Err(failure(13, "Anchor requires a Range"));
        };
        if doc != self.context {
            return Err(failure(1004, "Anchor belongs to another workbook"));
        }
        let Object::Range(si, r) = *inner else {
            return Err(failure(1004, "native hyperlinks require a cell Range"));
        };
        let rect = r.rectangle()?;
        if rect.count() != 1 || f.sheet != Some(si) || f.range.as_ref().is_some_and(|r| !r.areas.iter().any(|r| r.contains(rect.start))) {
            return Err(failure(1004, "Anchor must be one cell within the receiver"));
        }
        let address = string(required(&a, 1)?)?;
        let sub = if absent(&a[2]) { String::new() } else { string(&a[2])? };
        let target = if sub.is_empty() { address } else { format!("{address}#{sub}") };
        if target.is_empty() || target.len() > 1_048_576 {
            return Err(failure(1004, "hyperlink target empty or too long"));
        }
        let mut p = json!({"sheet":si,"cell":rect.start.a1(),"target":target});
        for (i, k) in [(3, "tooltip"), (4, "text")] {
            if !absent(&a[i]) {
                let t = string(&a[i])?;
                if t.len() > 1_048_576 {
                    return Err(failure(7, "hyperlink text budget"));
                }
                p[k] = json!(t);
            }
        }
        self.objects.reserve(1)?;
        execute(s, "insert.link", p)?;
        self.feature(node("Hyperlink", si, rect.start))
    }
    fn hyperlink(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let si = f.sheet.ok_or_else(|| failure(91, "worksheet missing"))?;
        let cell = f.range.as_ref().ok_or_else(|| failure(91, "link deleted"))?.rectangle()?.start;
        let sh = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let link = sh.hyperlinks.get(&cell).ok_or_else(|| failure(91, "hyperlink deleted"))?;
        if link.target.len() > 1_048_576 || link.tooltip.as_ref().is_some_and(|t| t.len() > 1_048_576) {
            return Err(failure(7, "native hyperlink payload budget"));
        }
        let link = link.clone();
        let text = sh.value(cell).display();
        let (address, sub) =
            if link.external { link.target.split_once('#').unwrap_or((&link.target, "")) } else { ("", link.target.trim_start_matches('#')) };
        let m = c.member.to_ascii_lowercase();
        if m == "delete" && c.kind == InvokeKind::Call {
            arguments(c, &[])?;
            execute(s, "insert.link", json!({"sheet":si,"cell":cell.a1(),"remove":true}))?;
            self.objects.invalidate_annotation(self.context, si, cell, true)?;
            return Ok(Value::Empty);
        }
        if c.kind == InvokeKind::Let && matches!(m.as_str(), "address" | "subaddress" | "screentip" | "texttodisplay") {
            let a = arguments(c, &["Value"])?;
            let value = string(required(&a, 0)?)?;
            if value.len() > 1_048_576 {
                return Err(failure(7, "hyperlink text budget"));
            }
            let target = match m.as_str() {
                "address" => {
                    if sub.is_empty() {
                        value.clone()
                    } else {
                        format!("{value}#{sub}")
                    }
                }
                "subaddress" => {
                    if value.is_empty() {
                        address.to_string()
                    } else {
                        format!("{address}#{value}")
                    }
                }
                _ => link.target.clone(),
            };
            if target.is_empty() {
                return Err(failure(1004, "empty hyperlink target"));
            }
            let mut p = json!({"sheet":si,"cell":cell.a1(),"target":target,"tooltip":link.tooltip});
            if m == "screentip" {
                p["tooltip"] = json!(value);
            } else if m == "texttodisplay" {
                p["text"] = json!(value);
            }
            execute(s, "insert.link", p)?;
            return Ok(Value::Empty);
        }
        if !reads(c) {
            return Err(failure(438, "unsupported hyperlink setter"));
        }
        if m == "follow" {
            let a = arguments(c, &["NewWindow", "AddHistory", "ExtraInfo", "Method", "HeaderInfo"])?;
            if !absent(&a[0]) && super::address::boolean(&a[0])? || a[1..].iter().any(|v| !absent(v)) {
                return Err(failure(1004, "native link navigation does not support browser/request options"));
            }
            if link.external {
                return Err(failure(1004, "external Follow requires a connected navigation continuation"));
            }
            execute(s, "link.open", json!({"sheet":si,"cell":cell.a1()}))?;
            return Ok(Value::Empty);
        }
        arguments(c, &[])?;
        match m.as_str() {
            "application" => self.intern(Object::Application),
            "parent" | "range" => self.intern(Object::Range(si, gridcraft_engine::core::RangeRef::new(cell, cell).into())),
            "address" => values::vba_text(address),
            "subaddress" => values::vba_text(sub),
            "screentip" => values::vba_text(link.tooltip.as_deref().unwrap_or("")),
            "texttodisplay" => values::vba_text(&text),
            "type" => Ok(Value::Long(0)),
            _ => Err(failure(438, "unsupported hyperlink member")),
        }
    }
    fn document_property(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let (name, key, _) = PROPS.iter().find(|(_, _, i)| *i == f.part).ok_or_else(|| failure(91, "document property unavailable"))?;
        let m = c.member.to_ascii_lowercase();
        if matches!(m.as_str(), "value" | "") && matches!(c.kind, InvokeKind::Let | InvokeKind::DefaultLet) {
            let a = arguments(c, &["Value"])?;
            let text = string(required(&a, 0)?)?;
            if text.len() > 1_048_576 {
                return Err(failure(7, "document property budget"));
            }
            let mut p = json!({});
            p[*key] = json!(text);
            execute(s, "file.properties", p)?;
            return Ok(Value::Empty);
        }
        if !reads(c) {
            return Err(failure(438, "unsupported document property setter"));
        }
        arguments(c, &[])?;
        match m.as_str() {
            "name" => values::vba_text(name),
            "type" => Ok(Value::Long(4)),
            "linktocontent" => Ok(Value::Boolean(false)),
            "application" => self.intern(Object::Application),
            "parent" => self.feature(Feature::new("DocumentProperties", None)),
            "value" | "" => {
                let p = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb.props;
                values::vba_text(match *key {
                    "title" => &p.title,
                    "subject" => &p.subject,
                    "author" => &p.author,
                    "keywords" => &p.keywords,
                    "description" => &p.description,
                    _ => &p.company,
                })
            }
            _ => Err(failure(438, "unsupported document property member")),
        }
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("CommentThreaded", "Author", "Property Get Author() As Author"),
        ("Author", "Name", "Property Get Name() As String"),
        ("Author", "Application", "Property Get Application() As Application"),
        ("Worksheet", "Hyperlinks", "Property Get Hyperlinks() As Hyperlinks"),
        ("Range", "Hyperlinks", "Property Get Hyperlinks() As Hyperlinks"),
        ("Range", "NoteText", "Function NoteText(Optional Text As Variant, Optional Start As Variant, Optional Length As Variant) As String"),
        ("Hyperlink", "Follow", "Sub Follow(Optional NewWindow As Variant, Optional AddHistory As Variant, Optional ExtraInfo As Variant, Optional Method As Variant, Optional HeaderInfo As Variant)"),
        ("Range", "Comment", "Property Get Comment() As Comment"),
        ("Range", "AddComment", "Function AddComment(Optional Text As Variant) As Comment"),
        ("Range", "CommentThreaded", "Property Get CommentThreaded() As CommentThreaded"),
        ("Range", "AddCommentThreaded", "Function AddCommentThreaded(Text As Variant) As CommentThreaded"),
        ("Worksheet", "Comments", "Property Get Comments() As Comments"),
        ("Comments", "Count", "Property Get Count() As Long"),
        ("Comments", "Item", "Property Get Item(Index As Variant) As Comment"),
        ("Comments", "Application", "Property Get Application() As Application"),
        ("Comments", "Parent", "Property Get Parent() As Object"),
        ("Worksheet", "CommentsThreaded", "Property Get CommentsThreaded() As CommentsThreaded"),
        ("CommentsThreaded", "Count", "Property Get Count() As Long"),
        ("CommentsThreaded", "Item", "Property Get Item(Index As Variant) As CommentThreaded"),
        ("CommentsThreaded", "Application", "Property Get Application() As Application"),
        ("CommentsThreaded", "Parent", "Property Get Parent() As Object"),
        ("Hyperlinks", "Count", "Property Get Count() As Long"),
        ("Hyperlinks", "Item", "Property Get Item(Index As Variant) As Hyperlink"),
        ("Hyperlinks", "Application", "Property Get Application() As Application"),
        ("Hyperlinks", "Parent", "Property Get Parent() As Object"),
        ("DocumentProperties", "Count", "Property Get Count() As Long"),
        ("DocumentProperties", "Item", "Property Get Item(Index As Variant) As DocumentProperty"),
        ("DocumentProperties", "Application", "Property Get Application() As Application"),
        ("DocumentProperties", "Parent", "Property Get Parent() As Object"),
        ("Workbook", "BuiltinDocumentProperties", "Property Get BuiltinDocumentProperties() As Object"),
        ("Comment", "Application", "Property Get Application() As Object"),
        ("Comment", "Parent", "Property Get Parent() As Object"),
        ("CommentThreaded", "Application", "Property Get Application() As Object"),
        ("CommentThreaded", "Parent", "Property Get Parent() As Object"),
        ("Hyperlink", "Application", "Property Get Application() As Object"),
        ("Hyperlink", "Parent", "Property Get Parent() As Object"),
        ("DocumentProperty", "Application", "Property Get Application() As Object"),
        ("DocumentProperty", "Parent", "Property Get Parent() As Object"),
        ("Comment", "Text", "Function Text(Optional Text As Variant, Optional Start As Variant, Optional Overwrite As Variant) As String"),
        ("Comment", "Delete", "Sub Delete()"),
        ("Comment", "Next", "Function Next() As Comment"),
        ("Comment", "Previous", "Function Previous() As Comment"),
        ("CommentThreaded", "Text", "Function Text(Optional Text As Variant, Optional Start As Variant, Optional Overwrite As Variant) As String"),
        ("CommentThreaded", "Delete", "Sub Delete()"),
        ("CommentThreaded", "Next", "Function Next() As CommentThreaded"),
        ("CommentThreaded", "Previous", "Function Previous() As CommentThreaded"),
        ("Comment", "Author", "Property Get Author() As String"),
        ("Comment", "Visible", "Property Visible As Boolean"),
        ("CommentThreaded", "AddReply", "Function AddReply(Text As String) As CommentThreaded"),
        ("CommentThreaded", "Replies", "Property Get Replies() As CommentsThreaded"),
        ("Hyperlinks", "Add", "Function Add(Anchor As Object, Address As String, Optional SubAddress As Variant, Optional ScreenTip As Variant, Optional TextToDisplay As Variant) As Hyperlink"),
        ("Hyperlinks", "Delete", "Sub Delete()"),
        ("Hyperlink", "Delete", "Sub Delete()"),
        ("Hyperlink", "Address", "Property Address As String"),
        ("Hyperlink", "SubAddress", "Property SubAddress As String"),
        ("Hyperlink", "ScreenTip", "Property ScreenTip As String"),
        ("Hyperlink", "TextToDisplay", "Property TextToDisplay As String"),
        ("Hyperlink", "Range", "Property Get Range() As Range"),
        ("Hyperlink", "Type", "Property Get Type() As Long"),
        ("DocumentProperty", "Name", "Property Get Name() As String"),
        ("DocumentProperty", "Type", "Property Get Type() As Long"),
        ("DocumentProperty", "LinkToContent", "Property Get LinkToContent() As Boolean"),
        ("DocumentProperty", "Value", "Property Value As Variant"),
    ].into_iter().map(|(owner,name,signature)| super::MemberDescriptor {owner,name,signature,availability:super::MemberAvailability::Implemented,implementation_symbol:"crates/vba/src/annotation_api.rs::Bridge::annotation_api"}).collect()
}
