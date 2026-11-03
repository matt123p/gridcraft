//! Sheet-owned native print settings and header/footer definition views.
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::cmd::print::{PrintOptions, join_header_footer, layout_sheet, split_header_footer};
use gridcraft_engine::model::{HfPage, Orientation};
use serde_json::json;
const PAPERS: &[(i32, &str)] = &[
    (1, "Letter"),
    (3, "Tabloid"),
    (4, "Ledger"),
    (5, "Legal"),
    (7, "Executive"),
    (8, "A3"),
    (9, "A4"),
    (11, "A5"),
    (12, "B4"),
    (13, "B5"),
    (17, "11x17"),
];
fn page_node(si: usize, part: i32, key: u64, s: &Session) -> HostResult<Feature> {
    let mut f = Feature::new("Page", Some(si));
    f.part = part;
    f.key = key;
    if key > 0 {
        f.text = Some(s.doc().map_err(|e| failure(91, e.to_string()))?.revision.to_string());
    }
    Ok(f)
}
pub(crate) fn fresh(s: &Session, f: &Feature) -> HostResult<()> {
    if let Some(revision) = &f.text
        && *revision != s.doc().map_err(|e| failure(91, e.to_string()))?.revision.to_string()
    {
        return Err(failure(91, "pagination view changed; reacquire it"));
    }
    Ok(())
}
fn section(m: &str) -> Option<(bool, usize)> {
    Some(match m {
        "leftheader" => (false, 0),
        "centerheader" => (false, 1),
        "rightheader" => (false, 2),
        "leftfooter" => (true, 0),
        "centerfooter" => (true, 1),
        "rightfooter" => (true, 2),
        _ => return None,
    })
}
fn page_kind(part: i32) -> HfPage {
    match part {
        1 => HfPage::Even,
        2 => HfPage::First,
        _ => HfPage::Odd,
    }
}
fn absolute(r: gridcraft_engine::core::RangeRef) -> String {
    let a = format!("${}${}", gridcraft_engine::core::col_to_letters(r.start.col), r.start.row + 1);
    if r.start == r.end { a } else { format!("{a}:${}${}", gridcraft_engine::core::col_to_letters(r.end.col), r.end.row + 1) }
}
pub(crate) fn validate_settings(pr: &gridcraft_engine::model::PrintSettings) -> HostResult<()> {
    for (span, limit) in [(pr.title_rows, gridcraft_engine::core::MAX_ROWS), (pr.title_cols, gridcraft_engine::core::MAX_COLS)] {
        if span.is_some_and(|(a, b)| a > b || b >= limit) {
            return Err(failure(1004, "invalid native print-title bounds"));
        }
    }
    if let Some(area) = pr.print_area {
        super::ranges::Range::from(area).validate()?;
    }
    if pr.paper.len() > 1024
        || [&pr.header, &pr.footer, &pr.first_header, &pr.first_footer, &pr.even_header, &pr.even_footer].iter().any(|s| s.len() > 4096)
    {
        return Err(failure(7, "native print metadata budget"));
    }

    Ok(())
}
impl Bridge {
    pub(crate) fn page_setup_api(&mut self, s: &mut Session, o: &Object, c: &HostCall) -> Option<HostResult<Value>> {
        if let Object::Sheet(si) = o
            && c.member.eq_ignore_ascii_case("PageSetup")
            && reads(c)
        {
            return Some((|| {
                arguments(c, &[])?;
                self.feature(Feature::new("PageSetup", Some(*si)))
            })());
        }
        let Object::Feature(f) = o else {
            return None;
        };
        if !matches!(f.kind, "PageSetup" | "Pages" | "Page" | "HeaderFooter") {
            return None;
        }
        Some(self.page_setup_dispatch(s, f, c))
    }
    fn page_setup_dispatch(&mut self, s: &mut Session, f: &Feature, c: &HostCall) -> HostResult<Value> {
        let si = f.sheet.ok_or_else(|| failure(91, "worksheet deleted"))?;
        let d = s.doc().map_err(|e| failure(91, e.to_string()))?;
        let sh = d.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let pr = &sh.print;
        validate_settings(pr)?;

        let m = c.member.to_ascii_lowercase();
        fresh(s, f)?;
        if f.kind == "Pages" {
            let lay = layout_sheet(&d.wb, si, &PrintOptions::default(), None).ok_or_else(|| failure(1004, "no pagination"))?;
            if c.kind == InvokeKind::Enumerate {
                arguments(c, &[])?;
                let nodes = (0..lay.pages.len())
                    .map(|i| {
                        page_node(
                            si,
                            match pr.hf_page(i == 0, i + 1) {
                                HfPage::Odd => 0,
                                HfPage::Even => 1,
                                HfPage::First => 2,
                            },
                            i as u64 + 1,
                            s,
                        )
                        .map(Feature::object)
                    })
                    .collect::<HostResult<Vec<_>>>()?;
                return self.object_array(nodes);
            }
            if !reads(c) {
                return Err(failure(438, "pages are read-only"));
            }
            return match m.as_str() {
                "count" => {
                    arguments(c, &[])?;
                    Ok(Value::Long(lay.pages.len() as i32))
                }
                "item" | "" | "pages" => {
                    let a = arguments(c, &["Index"])?;
                    let n = numeric_index(required(&a, 0)?)?;
                    if n < 1 || n as usize > lay.pages.len() {
                        return Err(failure(9, "page index"));
                    }
                    let part = match pr.hf_page(n == 1, n as usize) {
                        HfPage::Odd => 0,
                        HfPage::Even => 1,
                        HfPage::First => 2,
                    };
                    self.feature(page_node(si, part, n as u64, s)?)
                }
                _ => Err(failure(438, "unsupported Pages member")),
            };
        }
        if f.kind == "Page" {
            if !reads(c) {
                return Err(failure(438, "Page definition is not assignable"));
            }
            arguments(c, &[])?;
            let (footer, index) = section(&m).ok_or_else(|| failure(438, "unsupported Page member"))?;
            let mut child = Feature::new("HeaderFooter", Some(si));
            child.part = f.part * 10 + i32::from(footer) * 3 + index as i32;
            child.key = f.key;
            child.text = f.text.clone();
            return self.feature(child);
        }
        if f.kind == "HeaderFooter" {
            if m != "text" {
                return Err(failure(438, "header/footer pictures have no native backend"));
            }
            let kind = page_kind(f.part / 10);
            let footer = f.part % 10 >= 3;
            let index = (f.part % 10 % 3) as usize;
            if reads(c) {
                arguments(c, &[])?;
                return values::vba_text(split_header_footer(pr.hf(kind, footer)).get(index).map(String::as_str).unwrap_or_default());
            }
            if c.kind != InvokeKind::Let || f.key > 0 {
                return Err(failure(438, "numbered pagination views are read-only"));
            }
            let a = arguments(c, &["Value"])?;
            let text = string(required(&a, 0)?)?;
            return self.set_print_section(s, si, kind, footer, index, &text);
        }
        if reads(c) {
            if m == "pages" {
                let child = Feature::new("Pages", Some(si));
                return if c.arguments.is_empty() { self.feature(child) } else { self.page_setup_dispatch(s, &child, c) };
            }
            arguments(c, &[])?;
            if let Some((footer, index)) = section(&m) {
                return values::vba_text(split_header_footer(pr.hf(HfPage::Odd, footer)).get(index).map(String::as_str).unwrap_or_default());
            }
            if let Some(index) =
                ["leftmargin", "rightmargin", "topmargin", "bottommargin", "headermargin", "footermargin"].iter().position(|n| *n == m)
            {
                return Ok(Value::Double(f64::from(pr.margins.get(index).copied().unwrap_or_default()) * 72.0));
            }
            return match m.as_str() {
                "parent" => self.intern(Object::Sheet(si)),
                "application" => self.intern(Object::Application),
                "order" => Ok(Value::Long(1)),
                "orientation" => Ok(Value::Long(if pr.orientation == Orientation::Portrait { 1 } else { 2 })),
                "papersize" => Ok(Value::Long(
                    PAPERS
                        .iter()
                        .find(|(_, name)| name.eq_ignore_ascii_case(&pr.paper))
                        .map(|p| p.0)
                        .ok_or_else(|| failure(1004, "native paper has no supported Excel enum mapping"))?,
                )),
                "zoom" => Ok(if pr.fit_to.is_some() { Value::Boolean(false) } else { Value::Long(i32::from(pr.scale)) }),
                "fittopageswide" | "fittopagestall" => {
                    let (w, h) = pr.fit_to.ok_or_else(|| failure(1004, "native has no dormant fit counts; set Zoom=False first"))?;
                    let n = if m == "fittopageswide" { w } else { h };
                    Ok(if n == 0 { Value::Boolean(false) } else { Value::Long(i32::from(n)) })
                }
                "printarea" => values::vba_text(&pr.print_area.map(absolute).unwrap_or_default()),
                "printtitlerows" => values::vba_text(&pr.title_rows.map(|(a, b)| format!("${}:${}", a + 1, b + 1)).unwrap_or_default()),
                "printtitlecolumns" => values::vba_text(
                    &pr.title_cols
                        .map(|(a, b)| format!("${}:${}", gridcraft_engine::core::col_to_letters(a), gridcraft_engine::core::col_to_letters(b)))
                        .unwrap_or_default(),
                ),
                "firstpage" | "evenpage" => self.feature(page_node(si, if m == "firstpage" { 2 } else { 1 }, 0, s)?),
                _ => Ok(Value::Boolean(match m.as_str() {
                    "centerhorizontally" => pr.center_h,
                    "centervertically" => pr.center_v,
                    "printgridlines" => pr.gridlines,
                    "printheadings" => pr.headings,
                    "differentfirstpageheaderfooter" => pr.different_first,
                    "oddandevenpagesheaderfooter" => pr.different_odd_even,
                    "scalewithdocheaderfooter" => pr.hf_scale_with_doc,
                    "alignmarginsheaderfooter" => pr.hf_align_with_margins,
                    _ => return Err(failure(438, "unsupported page setting")),
                })),
            };
        }
        if c.kind != InvokeKind::Let {
            return Err(failure(438, "page setting requires Let"));
        }
        let a = arguments(c, &["Value"])?;
        let v = required(&a, 0)?;
        if let Some((footer, index)) = section(&m) {
            return self.set_print_section(s, si, HfPage::Odd, footer, index, &string(v)?);
        }
        let mut p = json!({"sheet":si});
        let mut command = "pageLayout.pageSetup";
        if let Some(index) = ["leftmargin", "rightmargin", "topmargin", "bottommargin", "headermargin", "footermargin"].iter().position(|n| *n == m) {
            let n = super::format_api::number(v)?;
            if !n.is_finite() || !(0.0..=1440.0).contains(&n) {
                return Err(failure(1004, "margin must be 0..1440 points"));
            }
            p["margins"] = json!({});
            p["margins"][["left", "right", "top", "bottom", "header", "footer"][index]] = json!(n / 72.0);
        } else {
            match m.as_str() {
                "orientation" => {
                    p["orientation"] = json!(match numeric_index(v)? {
                        1 => "portrait",
                        2 => "landscape",
                        _ => return Err(failure(1004, "unsupported orientation")),
                    })
                }
                "papersize" => {
                    p["paper"] = json!(
                        PAPERS
                            .iter()
                            .find(|(n, _)| numeric_index(v).is_ok_and(|value| value == *n))
                            .map(|p| p.1)
                            .ok_or_else(|| failure(1004, "unsupported native paper size"))?
                    )
                }
                "zoom" => {
                    if matches!(v, Value::Boolean(false)) {
                        p["fitToPage"] = json!(true);
                    } else {
                        let n = numeric_index(v)?;
                        if !(10..=400).contains(&n) {
                            return Err(failure(1004, "Zoom must be 10..400 or False"));
                        }
                        p["scale"] = json!(n);
                    }
                }
                "fittopageswide" | "fittopagestall" => {
                    if pr.fit_to.is_none() {
                        return Err(failure(1004, "set Zoom=False before fit counts"));
                    }
                    let n = if matches!(v, Value::Boolean(false)) { 0 } else { numeric_index(v)? };
                    if !(0..=32767).contains(&n) {
                        return Err(failure(1004, "fit count must be 0..32767"));
                    }
                    p[if m == "fittopageswide" { "fitWidth" } else { "fitHeight" }] = json!(n);
                }
                "printarea" | "printtitlerows" | "printtitlecolumns" => {
                    let text = if m == "printarea" && matches!(v, Value::Boolean(false)) { String::new() } else { string(v)? };
                    let key = match m.as_str() {
                        "printarea" => "printArea",
                        "printtitlerows" => "titleRows",
                        _ => "titleCols",
                    };
                    if text.is_empty() {
                        p[key] = json!("");
                    } else {
                        let (owner, r) = Self::resolve_text(s, si, &text, None, 0, &mut 0)?;
                        let rect = r.rectangle()?;
                        if owner != si || m == "printtitlerows" && !rect.is_full_rows() || m == "printtitlecolumns" && !rect.is_full_cols() {
                            return Err(failure(1004, "print reference must have the receiver sheet and correct geometry"));
                        }
                        p[key] = json!(rect.a1());
                    }
                }
                _ => {
                    let (key, hf) = match m.as_str() {
                        "centerhorizontally" => ("centerH", false),
                        "centervertically" => ("centerV", false),
                        "printgridlines" => ("gridlines", false),
                        "printheadings" => ("headings", false),
                        "differentfirstpageheaderfooter" => ("differentFirst", true),
                        "oddandevenpagesheaderfooter" => ("differentOddEven", true),
                        "scalewithdocheaderfooter" => ("scaleWithDoc", true),
                        "alignmarginsheaderfooter" => ("alignWithMargins", true),
                        _ => return Err(failure(438, "unsupported page setting")),
                    };
                    p[key] = json!(super::address::boolean(v)?);
                    if hf {
                        command = "pageLayout.headerFooter";
                    }
                }
            }
        }
        s.execute(command, p).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    fn set_print_section(&mut self, s: &mut Session, si: usize, kind: HfPage, footer: bool, index: usize, text: &str) -> HostResult<Value> {
        if text.len() > 4096 {
            return Err(failure(7, "header/footer input budget"));
        }
        let pr = &s.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(si).ok_or_else(|| failure(91, "worksheet deleted"))?.print;
        let mut parts = split_header_footer(pr.hf(kind, footer));
        let converted = gridcraft_engine::cmd::print::fields_to_codes(text);
        if let Some(slot) = parts.get_mut(index) {
            *slot = converted;
        }
        if join_header_footer(&parts).chars().count() > 255 {
            return Err(failure(1004, "combined header/footer exceeds native 255-character limit"));
        }
        s.execute("pageLayout.headerFooter",json!({"sheet":si,"page":kind.id(),"part":if footer{"footer"}else{"header"},"section":(["left","center","right"].get(index).copied().ok_or_else(||failure(1004,"header section"))?),"text":text})).map_err(|e|failure(1004,e.to_string()))?;
        Ok(Value::Empty)
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    [
        ("Worksheet", "PageSetup", "Property Get PageSetup() As PageSetup"),
        ("PageSetup", "Parent", "Property Get Parent() As Worksheet"),
        ("PageSetup", "Application", "Property Get Application() As Application"),
        ("PageSetup", "Orientation", "Property Orientation As XlPageOrientation"),
        ("PageSetup", "PaperSize", "Property PaperSize As XlPaperSize"),
        ("PageSetup", "Zoom", "Property Zoom As Variant"),
        ("PageSetup", "FitToPagesWide", "Property FitToPagesWide As Variant"),
        ("PageSetup", "FitToPagesTall", "Property FitToPagesTall As Variant"),
        ("PageSetup", "Order", "Property Get Order() As XlOrder"),
        ("PageSetup", "LeftMargin", "Property LeftMargin As Double"),
        ("PageSetup", "RightMargin", "Property RightMargin As Double"),
        ("PageSetup", "TopMargin", "Property TopMargin As Double"),
        ("PageSetup", "BottomMargin", "Property BottomMargin As Double"),
        ("PageSetup", "HeaderMargin", "Property HeaderMargin As Double"),
        ("PageSetup", "FooterMargin", "Property FooterMargin As Double"),
        ("PageSetup", "PrintArea", "Property PrintArea As String"),
        ("PageSetup", "PrintTitleRows", "Property PrintTitleRows As String"),
        ("PageSetup", "PrintTitleColumns", "Property PrintTitleColumns As String"),
        ("PageSetup", "LeftHeader", "Property LeftHeader As String"),
        ("PageSetup", "CenterHeader", "Property CenterHeader As String"),
        ("PageSetup", "RightHeader", "Property RightHeader As String"),
        ("PageSetup", "LeftFooter", "Property LeftFooter As String"),
        ("PageSetup", "CenterFooter", "Property CenterFooter As String"),
        ("PageSetup", "RightFooter", "Property RightFooter As String"),
        ("PageSetup", "CenterHorizontally", "Property CenterHorizontally As Boolean"),
        ("PageSetup", "CenterVertically", "Property CenterVertically As Boolean"),
        ("PageSetup", "PrintGridlines", "Property PrintGridlines As Boolean"),
        ("PageSetup", "PrintHeadings", "Property PrintHeadings As Boolean"),
        ("PageSetup", "DifferentFirstPageHeaderFooter", "Property DifferentFirstPageHeaderFooter As Boolean"),
        ("PageSetup", "OddAndEvenPagesHeaderFooter", "Property OddAndEvenPagesHeaderFooter As Boolean"),
        ("PageSetup", "ScaleWithDocHeaderFooter", "Property ScaleWithDocHeaderFooter As Boolean"),
        ("PageSetup", "AlignMarginsHeaderFooter", "Property AlignMarginsHeaderFooter As Boolean"),
        ("PageSetup", "FirstPage", "Property Get FirstPage() As Page"),
        ("PageSetup", "EvenPage", "Property Get EvenPage() As Page"),
        ("PageSetup", "Pages", "Property Get Pages() As Pages"),
        ("Pages", "Count", "Property Get Count() As Long"),
        ("Pages", "Item", "Property Get Item(Index As Variant) As Page"),
        ("Page", "LeftHeader", "Property Get LeftHeader() As HeaderFooter"),
        ("Page", "CenterHeader", "Property Get CenterHeader() As HeaderFooter"),
        ("Page", "RightHeader", "Property Get RightHeader() As HeaderFooter"),
        ("Page", "LeftFooter", "Property Get LeftFooter() As HeaderFooter"),
        ("Page", "CenterFooter", "Property Get CenterFooter() As HeaderFooter"),
        ("Page", "RightFooter", "Property Get RightFooter() As HeaderFooter"),
        ("HeaderFooter", "Text", "Property Text As String"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| super::MemberDescriptor {
        owner,
        name,
        signature,
        availability: super::MemberAvailability::Implemented,
        implementation_symbol: "crates/vba/src/page_setup_api.rs::Bridge::page_setup_api",
    })
    .collect()
}
