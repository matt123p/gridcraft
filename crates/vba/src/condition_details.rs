//! Typed conditional-format thresholds, colors and condition-specific properties.
use super::conditional_api::{rule_style, set_rule_style};
use super::features::{Feature, reads};
use super::objects::{arguments, numeric_index, required, string};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::model::{CfRule, CfValueKind, Color, CondFormat};
use serde_json::json;
const ICONS: &[(i32, &str)] = &[
    (1, "3Arrows"),
    (2, "3ArrowsGray"),
    (3, "3Flags"),
    (4, "3TrafficLights1"),
    (5, "3TrafficLights2"),
    (6, "3Signs"),
    (7, "3Symbols"),
    (8, "3Symbols2"),
    (9, "4Arrows"),
    (10, "4ArrowsGray"),
    (11, "4RedToBlack"),
    (12, "4Rating"),
    (13, "4TrafficLights"),
    (14, "5Arrows"),
    (15, "5ArrowsGray"),
    (16, "5Rating"),
    (17, "5Quarters"),
    (18, "3Stars"),
    (19, "3Triangles"),
    (20, "5Boxes"),
];
pub(crate) fn is_detail(kind: &str) -> bool {
    matches!(kind, "ConditionValue" | "FormatColor" | "ColorScaleCriteria" | "ColorScaleCriterion" | "IconCriteria" | "IconCriterion")
}
fn condition_type(value: &CfValueKind) -> i32 {
    match value {
        CfValueKind::Number(_) => 0,
        CfValueKind::Min => 1,
        CfValueKind::Max => 2,
        CfValueKind::Percent(_) => 3,
        CfValueKind::Formula(_) => 4,
        CfValueKind::Percentile(_) => 5,
    }
}
fn condition_value(value: &CfValueKind) -> HostResult<Value> {
    match value {
        CfValueKind::Number(n) | CfValueKind::Percent(n) | CfValueKind::Percentile(n) => Ok(Value::Double(*n)),
        CfValueKind::Formula(s) => values::vba_text(&format!("={}", s.trim_start_matches('='))),
        _ => Ok(Value::Empty),
    }
}
fn threshold(kind: i32, value: &Value) -> HostResult<CfValueKind> {
    Ok(match kind {
        0 => CfValueKind::Number(super::format_api::number(value)?),
        1 | 6 => CfValueKind::Min,
        2 | 7 => CfValueKind::Max,
        3 | 5 => {
            let n = super::format_api::number(value)?;
            if !(0.0..=100.0).contains(&n) {
                return Err(failure(1004, "threshold percentage must be between zero and 100"));
            }
            if kind == 3 { CfValueKind::Percent(n) } else { CfValueKind::Percentile(n) }
        }
        4 => CfValueKind::Formula(string(value)?.trim_start_matches('=').into()),
        _ => return Err(failure(1004, "unsupported condition value type")),
    })
}
pub(crate) fn literal(value: &Value) -> HostResult<String> {
    let value = super::data_api::text(value)?;
    Ok(if value.starts_with('=') || gridcraft_engine::core::parse::parse_number_text(&value).is_some() {
        value.trim_start_matches('=').into()
    } else {
        format!("\"{}\"", value.replace('"', "\"\""))
    })
}
impl Bridge {
    pub(crate) fn iconset_api(&mut self, session: &Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if *object == Object::Workbook && call.member.eq_ignore_ascii_case("IconSets") && reads(call) {
            let f = Feature::new("IconSets", None);
            return Some(if call.arguments.is_empty() { self.feature(f) } else { self.iconset_dispatch(session, &f, call) });
        }
        if let Object::Feature(f) = object
            && matches!(f.kind, "IconSet" | "IconSets")
        {
            return Some(self.iconset_dispatch(session, f, call));
        }
        None
    }
    fn iconset_dispatch(&mut self, _session: &Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let member = call.member.to_ascii_lowercase();
        if !reads(call) && call.kind != InvokeKind::Enumerate {
            return Err(failure(438, "icon-set catalog is read-only"));
        }
        if member == "parent" {
            arguments(call, &[])?;
            return self.intern(Object::Workbook);
        }
        if member == "application" {
            arguments(call, &[])?;
            return self.intern(Object::Application);
        }
        if f.kind == "IconSets" {
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(
                    ICONS
                        .iter()
                        .map(|(id, _)| {
                            let mut f = Feature::new("IconSet", None);
                            f.key = *id as u64;
                            f.object()
                        })
                        .collect(),
                );
            }
            if member == "count" {
                arguments(call, &[])?;
                return Ok(Value::Long(ICONS.len() as i32));
            }
            if matches!(member.as_str(), "item" | "" | "iconsets") {
                let args = arguments(call, &["Index"])?;
                let id = numeric_index(required(&args, 0)?)?;
                if !ICONS.iter().any(|(n, _)| *n == id) {
                    return Err(failure(9, "unknown icon set"));
                }
                let mut f = Feature::new("IconSet", None);
                f.key = id as u64;
                return self.feature(f);
            }
        }
        let (_, name) = ICONS.iter().find(|(id, _)| *id as u64 == f.key).ok_or_else(|| failure(9, "unknown icon set"))?;
        arguments(call, &[])?;
        match member.as_str() {
            "id" => Ok(Value::Long(f.key as i32)),
            "count" => Ok(Value::Long(gridcraft_engine::model::icon_count(name) as i32)),
            _ => Err(failure(438, "unsupported IconSet member")),
        }
    }
    pub(crate) fn condition_details(
        &mut self,
        session: &mut Session,
        f: &Feature,
        call: &HostCall,
        index: usize,
        cf: &CondFormat,
    ) -> Option<HostResult<Value>> {
        let member = call.member.to_ascii_lowercase();
        let known = is_detail(f.kind)
            || matches!(
                member.as_str(),
                "colorscalecriteria"
                    | "minpoint"
                    | "maxpoint"
                    | "barcolor"
                    | "iconcriteria"
                    | "iconset"
                    | "topbottom"
                    | "dupeunique"
                    | "abovebelow"
                    | "numstddev"
                    | "numberformat"
                    | "modify"
            );
        if !known {
            return None;
        }
        Some((|| {
            let sheet = f.sheet.ok_or_else(|| failure(91, "condition owner deleted"))?;
            if reads(call) {
                let child = match (&cf.rule, member.as_str()) {
                    (CfRule::ColorScale { .. }, "colorscalecriteria") => Some(("ColorScaleCriteria", 0)),
                    (CfRule::DataBar { .. }, "minpoint") => Some(("ConditionValue", 0)),
                    (CfRule::DataBar { .. }, "maxpoint") => Some(("ConditionValue", 1)),
                    (CfRule::DataBar { .. }, "barcolor") => Some(("FormatColor", -1)),
                    (CfRule::IconSet { .. }, "iconcriteria") => Some(("IconCriteria", 0)),
                    _ => None,
                };
                if let Some((kind, part)) = child {
                    let mut node = f.clone();
                    node.kind = kind;
                    node.part = part;
                    if call.arguments.is_empty() {
                        return self.feature(node);
                    }
                    let mut call = call.clone();
                    call.member = "Item".into();
                    return self.condition_details(session, &node, &call, index, cf).ok_or_else(|| failure(438, "condition child unavailable"))?;
                }
                if let CfRule::IconSet { set, .. } = &cf.rule
                    && member == "iconset"
                {
                    arguments(call, &[])?;
                    let (id, _) =
                        ICONS.iter().find(|(_, name)| name == set).ok_or_else(|| failure(1004, "native icon set has no portable enum identity"))?;
                    let mut node = Feature::new("IconSet", None);
                    node.key = *id as u64;
                    return self.feature(node);
                }
            }
            if matches!(f.kind, "ColorScaleCriteria" | "IconCriteria") {
                let count = match &cf.rule {
                    CfRule::ColorScale { stops } => stops.len(),
                    CfRule::IconSet { thresholds, .. } => thresholds.len(),
                    _ => return Err(failure(438, "condition no longer has criteria")),
                };
                if call.kind == InvokeKind::Enumerate {
                    arguments(call, &[])?;
                    return self.object_array(
                        (0..count)
                            .map(|i| {
                                let mut node = f.clone();
                                node.kind = if f.kind == "ColorScaleCriteria" { "ColorScaleCriterion" } else { "IconCriterion" };
                                node.part = i as i32;
                                node.object()
                            })
                            .collect(),
                    );
                }
                if member == "count" && reads(call) {
                    arguments(call, &[])?;
                    return Ok(Value::Long(count as i32));
                }
                if matches!(member.as_str(), "item" | "") && reads(call) {
                    let args = arguments(call, &["Index"])?;
                    let i = numeric_index(required(&args, 0)?)?;
                    if i < 1 || i as usize > count {
                        return Err(failure(9, "condition criterion index"));
                    }
                    let mut node = f.clone();
                    node.kind = if f.kind == "ColorScaleCriteria" { "ColorScaleCriterion" } else { "IconCriterion" };
                    node.part = i - 1;
                    return self.feature(node);
                }
                return Err(failure(438, "unsupported criteria collection member"));
            }
            let mut rule = cf.rule.clone();
            if f.kind == "FormatColor" {
                let color = match &mut rule {
                    CfRule::DataBar { color, .. } if f.part == -1 => color,
                    CfRule::ColorScale { stops } => &mut stops.get_mut(f.part as usize).ok_or_else(|| failure(91, "color criterion removed"))?.1,
                    _ => return Err(failure(438, "color is unavailable")),
                };
                if reads(call) {
                    arguments(call, &[])?;
                    return match member.as_str() {
                        "color" => super::format_api::ole(*color, session),
                        "themecolor" => match color {
                            Color::Theme(i, _) => Ok(Value::Long(super::format_api::theme_number(*i))),
                            _ => Err(failure(1004, "color has no theme identity")),
                        },
                        "tintandshade" => Ok(Value::Double(match color {
                            Color::Theme(_, t) => f64::from(*t) / 1000.0,
                            _ => 0.0,
                        })),
                        "parent" => {
                            let mut node = f.clone();
                            node.kind = if f.part == -1 { "Databar" } else { "ColorScaleCriterion" };
                            self.feature(node)
                        }
                        "application" => self.intern(Object::Application),
                        _ => Err(failure(438, "unsupported FormatColor member")),
                    };
                }
                if call.kind != InvokeKind::Let {
                    return Err(failure(13, "color assignment requires Let"));
                }
                let args = arguments(call, &["Value"])?;
                let value = required(&args, 0)?;
                *color = match member.as_str() {
                    "color" => super::format_api::rgb(value)?,
                    "themecolor" => Color::Theme(super::format_api::theme_index(numeric_index(value)?)?, 0),
                    "tintandshade" => {
                        let n = super::format_api::number(value)?;
                        if !(-1.0..=1.0).contains(&n) {
                            return Err(failure(1004, "invalid tint"));
                        }
                        let Color::Theme(i, _) = color else {
                            return Err(failure(1004, "RGB tint identity is unavailable"));
                        };
                        Color::Theme(*i, (n * 1000.0).round() as i16)
                    }
                    _ => return Err(failure(438, "unsupported FormatColor assignment")),
                };
                return self.condition_update(session, sheet, index, json!({"rule":rule}));
            }
            if matches!(f.kind, "ConditionValue" | "ColorScaleCriterion" | "IconCriterion") {
                if member == "formatcolor" && f.kind == "ColorScaleCriterion" && reads(call) {
                    arguments(call, &[])?;
                    let mut node = f.clone();
                    node.kind = "FormatColor";
                    return self.feature(node);
                }
                let value = match &mut rule {
                    CfRule::DataBar { min, max, .. } => {
                        if f.part == 0 {
                            min
                        } else {
                            max
                        }
                    }
                    CfRule::ColorScale { stops } => &mut stops.get_mut(f.part as usize).ok_or_else(|| failure(91, "color criterion removed"))?.0,
                    CfRule::IconSet { thresholds, .. } => thresholds.get_mut(f.part as usize).ok_or_else(|| failure(91, "icon criterion removed"))?,
                    _ => return Err(failure(438, "threshold is unavailable")),
                };
                if reads(call) && member != "modify" {
                    arguments(call, &[])?;
                    return match member.as_str() {
                        "type" => Ok(Value::Long(condition_type(value))),
                        "value" => condition_value(value),
                        "index" => Ok(Value::Long(f.part + 1)),
                        "operator" if f.kind == "IconCriterion" => Ok(Value::Long(7)),
                        "parent" => {
                            let mut node = f.clone();
                            node.kind = match f.kind {
                                "ConditionValue" => "Databar",
                                "ColorScaleCriterion" => "ColorScaleCriteria",
                                _ => "IconCriteria",
                            };
                            node.part = 0;
                            self.feature(node)
                        }
                        "application" => self.intern(Object::Application),
                        _ => Err(failure(438, "unsupported threshold member")),
                    };
                }
                let next = if member == "modify" {
                    let args = arguments(call, &["newType", "newValue"])?;
                    threshold(numeric_index(required(&args, 0)?)?, &args[1])?
                } else {
                    if call.kind != InvokeKind::Let {
                        return Err(failure(13, "criterion assignment requires Let"));
                    }
                    let args = arguments(call, &["Value"])?;
                    let v = required(&args, 0)?;
                    match member.as_str() {
                        "type" => threshold(numeric_index(v)?, &condition_value(value)?)?,
                        "value" => threshold(condition_type(value), v)?,
                        "operator" if f.kind == "IconCriterion" => {
                            if numeric_index(v)? != 7 {
                                return Err(failure(1004, "native icon thresholds use greater-or-equal comparison"));
                            }
                            value.clone()
                        }
                        _ => return Err(failure(438, "unsupported criterion assignment")),
                    }
                };
                *value = next;
                return self.condition_update(session, sheet, index, json!({"rule":rule}));
            }
            if member == "numberformat" {
                let mut style = rule_style(&rule)?;
                if reads(call) {
                    arguments(call, &[])?;
                    return values::vba_text(style.num_fmt.as_str());
                }
                let args = arguments(call, &["Value"])?;
                style.num_fmt = gridcraft_engine::model::NumFmt::new(&string(required(&args, 0)?)?);
                let rule = set_rule_style(&rule, &style)?;
                return self.condition_update(session, sheet, index, json!({"rule":rule}));
            }
            if member == "modify" {
                let args = arguments(call, &["Type", "Operator", "Formula1", "Formula2"])?;
                let style = Box::new(rule_style(&rule)?);
                rule = match numeric_index(required(&args, 0)?)? {
                    1 => {
                        let op = match numeric_index(required(&args, 1)?)? {
                            1 => gridcraft_engine::model::CfOperator::Between,
                            2 => gridcraft_engine::model::CfOperator::NotBetween,
                            3 => gridcraft_engine::model::CfOperator::Equal,
                            4 => gridcraft_engine::model::CfOperator::NotEqual,
                            5 => gridcraft_engine::model::CfOperator::Greater,
                            6 => gridcraft_engine::model::CfOperator::Less,
                            7 => gridcraft_engine::model::CfOperator::GreaterOrEqual,
                            8 => gridcraft_engine::model::CfOperator::LessOrEqual,
                            _ => return Err(failure(1004, "invalid condition operator")),
                        };
                        let a = literal(required(&args, 2)?)?;
                        let b = if super::features::absent(&args[3]) { None } else { Some(literal(&args[3])?) };
                        if matches!(op, gridcraft_engine::model::CfOperator::Between | gridcraft_engine::model::CfOperator::NotBetween) && b.is_none()
                        {
                            return Err(failure(449, "Formula2 is required for a between condition"));
                        }
                        CfRule::CellIs { op, a, b, style }
                    }
                    2 => CfRule::Expression { formula: string(required(&args, 2)?)?.trim_start_matches('=').into(), style },
                    _ => return Err(failure(1004, "Modify supports existing cell/expression condition types")),
                };
                return self.condition_update(session, sheet, index, json!({"rule":rule}));
            }
            if reads(call) {
                arguments(call, &[])?;
                return match (&rule, member.as_str()) {
                    (CfRule::Top10 { bottom, .. }, "topbottom") => Ok(Value::Long(if *bottom { 0 } else { 1 })),
                    (CfRule::Duplicate { .. }, "dupeunique") => Ok(Value::Long(1)),
                    (CfRule::Unique { .. }, "dupeunique") => Ok(Value::Long(0)),
                    (CfRule::AboveAverage { below, equal, std_dev, .. }, "abovebelow") => Ok(Value::Long(if *std_dev > 0 {
                        if *below { 5 } else { 4 }
                    } else {
                        match (below, equal) {
                            (false, false) => 0,
                            (true, false) => 1,
                            (false, true) => 2,
                            (true, true) => 3,
                        }
                    })),
                    (CfRule::AboveAverage { std_dev, .. }, "numstddev") => Ok(Value::Long(i32::from(*std_dev))),
                    _ => Err(failure(438, "property does not apply to this rule")),
                };
            }
            if !matches!(call.kind, InvokeKind::Let | InvokeKind::Set) {
                return Err(failure(13, "condition assignment required"));
            }
            let args = arguments(call, &["Value"])?;
            let value = required(&args, 0)?;
            match (&mut rule, member.as_str()) {
                (CfRule::Top10 { bottom, .. }, "topbottom") => {
                    *bottom = match numeric_index(value)? {
                        0 => true,
                        1 => false,
                        _ => return Err(failure(1004, "invalid TopBottom")),
                    }
                }
                (CfRule::Duplicate { style }, "dupeunique") | (CfRule::Unique { style }, "dupeunique") => {
                    let style = style.clone();
                    rule = match numeric_index(value)? {
                        0 => CfRule::Unique { style },
                        1 => CfRule::Duplicate { style },
                        _ => return Err(failure(1004, "invalid DupeUnique")),
                    };
                }
                (CfRule::AboveAverage { below, equal, std_dev, .. }, "abovebelow") => {
                    let n = numeric_index(value)?;
                    if !(0..=5).contains(&n) {
                        return Err(failure(1004, "invalid AboveBelow"));
                    }
                    *below = matches!(n, 1 | 3 | 5);
                    *equal = matches!(n, 2 | 3);
                    if n < 4 {
                        *std_dev = 0;
                    } else if *std_dev == 0 {
                        *std_dev = 1;
                    }
                }
                (CfRule::AboveAverage { std_dev, .. }, "numstddev") => {
                    let n = numeric_index(value)?;
                    if !(0..=3).contains(&n) {
                        return Err(failure(1004, "NumStdDev must be between zero and three"));
                    }
                    *std_dev = n as u8;
                }
                (CfRule::IconSet { set, thresholds, .. }, "iconset") => {
                    let Value::Object(reference) = value else {
                        return Err(failure(13, "IconSet object required"));
                    };
                    let Object::Feature(node) = self.object(session, *reference)? else {
                        return Err(failure(13, "IconSet required"));
                    };
                    if node.kind != "IconSet" {
                        return Err(failure(13, "IconSet required"));
                    }
                    let (_, name) = ICONS.iter().find(|(id, _)| *id as u64 == node.key).ok_or_else(|| failure(1004, "unknown icon set"))?;
                    *set = (*name).into();
                    *thresholds = gridcraft_engine::model::default_icon_thresholds(name);
                }
                _ => return Err(failure(438, "property does not apply to this condition")),
            }
            self.condition_update(session, sheet, index, json!({"rule":rule}))
        })())
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    let mut result = Vec::new();
    for (owner, members) in [
        ("Workbook", &["IconSets"][..]),
        ("IconSets", &["Item", "Count", "Parent", "Application"]),
        ("IconSet", &["ID", "Count", "Parent", "Application"]),
        ("ColorScale", &["ColorScaleCriteria"]),
        ("ColorScaleCriteria", &["Item", "Count"]),
        ("ColorScaleCriterion", &["Index", "Type", "Value", "FormatColor"]),
        ("Databar", &["MinPoint", "MaxPoint", "BarColor"]),
        ("ConditionValue", &["Type", "Value", "Modify", "Parent", "Application"]),
        ("FormatColor", &["Color", "ThemeColor", "TintAndShade", "Parent", "Application"]),
        ("IconSetCondition", &["IconSet", "IconCriteria"]),
        ("IconCriteria", &["Item", "Count"]),
        ("IconCriterion", &["Index", "Type", "Value", "Operator"]),
        ("AboveAverage", &["AboveBelow", "NumStdDev", "NumberFormat", "Font", "Interior"]),
        ("UniqueValues", &["DupeUnique", "NumberFormat", "Font", "Interior"]),
        ("Top10", &["NumberFormat", "Font", "Interior"]),
        ("FormatCondition", &["NumberFormat", "Modify"]),
    ] {
        for name in members {
            result.push(super::MemberDescriptor {
                owner,
                name,
                signature: match (owner, *name) {
                    (_, "Font") => "Property Get Font() As Font",
                    (_, "Interior") => "Property Get Interior() As Interior",
                    ("IconSets", "Item") => "Property Get Item(Index As Variant) As IconSet",
                    ("ColorScaleCriteria", "Item") => "Property Get Item(Index As Variant) As ColorScaleCriterion",
                    ("IconCriteria", "Item") => "Property Get Item(Index As Variant) As IconCriterion",
                    ("ConditionValue", "Modify") => "Sub Modify(newType As Long, Optional newValue As Variant)",
                    (_, "Modify") => {
                        "Sub Modify(Type As Long, Optional Operator As Variant, Optional Formula1 As Variant, Optional Formula2 As Variant)"
                    }
                    (_, "IconSets") => "Property Get IconSets() As IconSets",
                    (_, "ID") => "Property Get ID() As Long",
                    (_, "Count") => "Property Get Count() As Long",
                    (_, "Parent") => "Property Get Parent() As Object",
                    (_, "Application") => "Property Get Application() As Application",
                    (_, "ColorScaleCriteria") => "Property Get ColorScaleCriteria() As ColorScaleCriteria",
                    (_, "Index") => "Property Get Index() As Long",
                    (_, "Type") => "Property Type As Long",
                    (_, "Value") => "Property Value As Variant",
                    (_, "FormatColor") => "Property Get FormatColor() As FormatColor",
                    (_, "MinPoint") => "Property Get MinPoint() As ConditionValue",
                    (_, "MaxPoint") => "Property Get MaxPoint() As ConditionValue",
                    (_, "BarColor") => "Property Get BarColor() As FormatColor",
                    (_, "Color") => "Property Color As Variant",
                    (_, "ThemeColor") => "Property ThemeColor As Long",
                    (_, "TintAndShade") => "Property TintAndShade As Double",
                    (_, "IconSet") => "Property IconSet As IconSet",
                    (_, "IconCriteria") => "Property Get IconCriteria() As IconCriteria",
                    (_, "Operator") => "Property Operator As Long",
                    (_, "AboveBelow") => "Property AboveBelow As Long",
                    (_, "NumStdDev") => "Property NumStdDev As Long",
                    (_, "NumberFormat") => "Property NumberFormat As String",
                    _ => "",
                },
                availability: super::MemberAvailability::Implemented,
                implementation_symbol: "crates/vba/src/condition_details.rs::Bridge::condition_details",
            });
        }
    }
    result
}
