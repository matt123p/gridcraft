//! Conditional-format objects are views of native rules, never a parallel rule store.
use super::features::{Feature, absent, reads};
use super::objects::{arguments, numeric_index, required};
use super::{Bridge, HostCall, HostResult, InvokeKind, Object, Session, Value, failure, values};
use gridcraft_engine::model::{CfOperator, CfRule, CondFormat, PatternType, Style};
use serde_json::{Value as Json, json};
fn kind(rule: &CfRule) -> &'static str {
    match rule {
        CfRule::ColorScale { .. } => "ColorScale",
        CfRule::DataBar { .. } => "Databar",
        CfRule::IconSet { .. } => "IconSetCondition",
        CfRule::Top10 { .. } => "Top10",
        CfRule::Duplicate { .. } | CfRule::Unique { .. } => "UniqueValues",
        CfRule::AboveAverage { .. } => "AboveAverage",
        _ => "FormatCondition",
    }
}
fn type_number(rule: &CfRule) -> i32 {
    match rule {
        CfRule::CellIs { .. } => 1,
        CfRule::Expression { .. } => 2,
        CfRule::ColorScale { .. } => 3,
        CfRule::DataBar { .. } => 4,
        CfRule::Top10 { .. } => 5,
        CfRule::IconSet { .. } => 6,
        CfRule::Duplicate { .. } | CfRule::Unique { .. } => 8,
        CfRule::ContainsText { .. } | CfRule::NotContainsText { .. } | CfRule::BeginsWith { .. } | CfRule::EndsWith { .. } => 9,
        CfRule::Blanks { .. } => 10,
        CfRule::TimePeriod { .. } => 11,
        CfRule::NoBlanks { .. } => 13,
        CfRule::Errors { .. } => 16,
        CfRule::NoErrors { .. } => 17,
        CfRule::AboveAverage { .. } => 12,
    }
}
fn op_number(op: CfOperator) -> i32 {
    match op {
        CfOperator::Between => 1,
        CfOperator::NotBetween => 2,
        CfOperator::Equal => 3,
        CfOperator::NotEqual => 4,
        CfOperator::Greater => 5,
        CfOperator::Less => 6,
        CfOperator::GreaterOrEqual => 7,
        CfOperator::LessOrEqual => 8,
    }
}
fn operator(n: i32) -> HostResult<&'static str> {
    match n {
        1 => Ok("between"),
        2 => Ok("notBetween"),
        3 => Ok("equal"),
        4 => Ok("notEqual"),
        5 => Ok("greater"),
        6 => Ok("less"),
        7 => Ok("greaterOrEqual"),
        8 => Ok("lessOrEqual"),
        _ => Err(failure(1004, "invalid condition operator")),
    }
}
pub(crate) fn rule_style(rule: &CfRule) -> HostResult<Style> {
    let value = serde_json::to_value(rule).map_err(|e| failure(1004, e.to_string()))?;
    let style = value
        .as_object()
        .and_then(|o| o.values().next())
        .and_then(|v| v.get("style"))
        .ok_or_else(|| failure(438, "this rule has no differential style"))?;
    serde_json::from_value(style.clone()).map_err(|e| failure(1004, e.to_string()))
}
pub(crate) fn set_rule_style(rule: &CfRule, style: &Style) -> HostResult<CfRule> {
    let mut value = serde_json::to_value(rule).map_err(|e| failure(1004, e.to_string()))?;
    let body = value.as_object_mut().and_then(|o| o.values_mut().next()).ok_or_else(|| failure(438, "invalid rule"))?;
    let body = body.as_object_mut().filter(|o| o.contains_key("style")).ok_or_else(|| failure(438, "this rule has no style"))?;
    body.insert("style".into(), serde_json::to_value(style).map_err(|e| failure(1004, e.to_string()))?);
    serde_json::from_value(value).map_err(|e| failure(1004, e.to_string()))
}
impl Bridge {
    pub(crate) fn conditional_api(&mut self, session: &mut Session, object: &Object, call: &HostCall) -> Option<HostResult<Value>> {
        if let Some(result) = self.iconset_api(session, object, call) {
            return Some(result);
        }
        if let Object::Range(sheet, range) = object
            && call.member.eq_ignore_ascii_case("FormatConditions")
            && reads(call)
        {
            return Some({
                let mut f = Feature::new("FormatConditions", Some(*sheet));
                f.range = Some(range.clone());
                if call.arguments.is_empty() {
                    self.feature(f)
                } else {
                    let mut call = call.clone();
                    call.member = "Item".into();
                    self.condition_dispatch(session, &f, &call)
                }
            });
        }
        let Object::Feature(f) = object else {
            return None;
        };
        if !super::condition_details::is_detail(f.kind)
            && !matches!(
                f.kind,
                "FormatConditions"
                    | "FormatCondition"
                    | "ColorScale"
                    | "Databar"
                    | "IconSetCondition"
                    | "Top10"
                    | "UniqueValues"
                    | "AboveAverage"
                    | "ConditionFont"
                    | "ConditionInterior"
            )
        {
            return None;
        }
        Some(self.condition_dispatch(session, f, call))
    }
    fn condition_node(sheet: usize, rule: &CondFormat) -> Feature {
        let mut node = Feature::new(kind(&rule.rule), Some(sheet));
        node.identity = Some(rule.identity.clone());
        node
    }
    fn condition_dispatch(&mut self, session: &mut Session, f: &Feature, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "rule worksheet gone"))?;
        let member = call.member.to_ascii_lowercase();
        let sh = session.doc().map_err(|e| failure(91, e.to_string()))?.wb.sheet(sheet).ok_or_else(|| failure(91, "worksheet deleted"))?;
        let protected = sh.is_protected();
        let rule_count = sh.cond_formats.len();
        if rule_count > super::MAX_OBJECTS {
            return Err(failure(7, "conditional format collection budget"));
        }
        if f.kind == "FormatConditions" {
            let range = f.range.as_ref().ok_or_else(|| failure(91, "condition range deleted"))?;
            range.validate()?;
            let rules = sh
                .cond_formats
                .iter()
                .filter(|cf| cf.ranges.iter().any(|r| range.areas.iter().any(|a| a.intersects(r))))
                .cloned()
                .collect::<Vec<_>>();
            if call.kind == InvokeKind::Enumerate {
                arguments(call, &[])?;
                return self.object_array(rules.iter().map(|r| Self::condition_node(sheet, r).object()).collect());
            }
            match member.as_str() {
                "count" if reads(call) => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(rules.len() as i32));
                }
                "item" | "" if reads(call) => {
                    let args = arguments(call, &["Index"])?;
                    let index = usize::try_from(numeric_index(required(&args, 0)?)?.checked_sub(1).ok_or_else(|| failure(9, "condition index"))?)
                        .map_err(|_| failure(9, "condition index"))?;
                    return self.feature(Self::condition_node(sheet, rules.get(index).ok_or_else(|| failure(9, "condition index"))?));
                }
                "delete" => {
                    arguments(call, &[])?;
                    session
                        .execute("home.clearRules", json!({"sheet":sheet,"range":range.rectangle()?.a1()}))
                        .map_err(|e| failure(1004, e.to_string()))?;
                    return Ok(Value::Empty);
                }
                "add" | "addcolorscale" | "adddatabar" | "addiconsetcondition" | "addtop10" | "adduniquevalues" | "addaboveaverage" => {
                    let rule = match member.as_str() {
                        "add" => {
                            let args = arguments(
                                call,
                                &["Type", "Operator", "Formula1", "Formula2", "String", "TextOperator", "DateOperator", "ScopeType"],
                            )?;
                            if args.iter().skip(4).any(|v| !absent(v)) {
                                return Err(failure(1004, "this conditional-format option is unavailable"));
                            }
                            match numeric_index(required(&args, 0)?)? {
                                1 => {
                                    json!({"type":"cellIs","operator":operator(if absent(&args[1]){1}else{numeric_index(&args[1])?})?,"value":super::condition_details::literal(required(&args,2)?)?,"value2":if absent(&args[3]){None}else{Some(super::condition_details::literal(&args[3])?)}})
                                }
                                2 => json!({"type":"expression","formula":super::condition_details::literal(required(&args,2)?)?}),
                                _ => return Err(failure(1004, "use the typed Add method for this condition kind")),
                            }
                        }
                        "addcolorscale" => {
                            let args = arguments(call, &["ColorScaleType"])?;
                            let n = numeric_index(required(&args, 0)?)?;
                            if !matches!(n, 2 | 3) {
                                return Err(failure(1004, "ColorScaleType must be 2 or 3"));
                            }
                            json!({"type":"colorScale","colors":if n==2{vec!["#ED7080","#65BFA0"]}else{vec!["#ED7080","#EBCF65","#65BFA0"]}})
                        }
                        "adddatabar" => {
                            arguments(call, &[])?;
                            json!({"type":"dataBar"})
                        }
                        "addiconsetcondition" => {
                            arguments(call, &[])?;
                            json!({"type":"iconSet"})
                        }
                        "addtop10" => {
                            arguments(call, &[])?;
                            json!({"type":"top10"})
                        }
                        "adduniquevalues" => {
                            arguments(call, &[])?;
                            json!({"type":"duplicate"})
                        }
                        _ => {
                            arguments(call, &[])?;
                            json!({"type":"aboveAverage"})
                        }
                    };
                    values::dimensions(range.rectangle()?)?;
                    if protected {
                        return Err(failure(1004, "conditional-format changes require an unprotected sheet"));
                    }
                    self.objects.reserve(1)?;
                    session
                        .execute("home.conditionalFormat", json!({"sheet":sheet,"range":range.rectangle()?.a1(),"rule":rule}))
                        .map_err(|e| failure(1004, e.to_string()))?;
                    let rule = session
                        .doc()
                        .map_err(|e| failure(91, e.to_string()))?
                        .wb
                        .sheet(sheet)
                        .and_then(|sh| sh.cond_formats.first())
                        .ok_or_else(|| failure(1004, "condition not created"))?;
                    return self.feature(Self::condition_node(sheet, rule));
                }
                "parent" if reads(call) => {
                    arguments(call, &[])?;
                    return self.intern(Object::Range(sheet, range.clone()));
                }
                "application" if reads(call) => {
                    arguments(call, &[])?;
                    return self.intern(Object::Application);
                }
                _ => return Err(failure(438, "unsupported FormatConditions member")),
            }
        }
        let (index, cf) = sh
            .cond_formats
            .iter()
            .enumerate()
            .find(|(_, cf)| f.identity.as_ref().is_some_and(|id| id.same(&cf.identity)))
            .ok_or_else(|| failure(91, "conditional format was deleted"))?;
        let mut cf = cf.clone();
        if let CfRule::IconSet { set, thresholds, .. } = &mut cf.rule
            && thresholds.is_empty()
        {
            *thresholds = gridcraft_engine::model::default_icon_thresholds(set);
        }
        if let Some(result) = self.condition_details(session, f, call, index, &cf) {
            return result;
        }

        if matches!(f.kind, "ConditionFont" | "ConditionInterior") {
            return self.condition_style(session, sheet, index, &cf, f.kind == "ConditionFont", call);
        }
        if reads(call) {
            match member.as_str() {
                "font" | "interior" => {
                    arguments(call, &[])?;
                    rule_style(&cf.rule)?;
                    let mut node = f.clone();
                    node.kind = if member == "font" { "ConditionFont" } else { "ConditionInterior" };
                    return self.feature(node);
                }
                "appliesto" | "parent" => {
                    arguments(call, &[])?;
                    return self.intern(Object::Range(sheet, super::ranges::Range { areas: cf.ranges.clone(), unit: super::ranges::Unit::Cells }));
                }
                "application" => {
                    arguments(call, &[])?;
                    return self.intern(Object::Application);
                }
                "type" => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(type_number(&cf.rule)));
                }
                "priority" => {
                    arguments(call, &[])?;
                    return Ok(Value::Long(cf.priority as i32));
                }
                "stopiftrue" => {
                    arguments(call, &[])?;
                    return Ok(Value::Boolean(cf.stop_if_true));
                }
                "formula1" | "formula2" | "operator" => {
                    arguments(call, &[])?;
                    return match &cf.rule {
                        CfRule::CellIs { op, a, b, .. } => {
                            if member == "operator" {
                                Ok(Value::Long(op_number(*op)))
                            } else {
                                values::vba_text(if member == "formula1" { a } else { b.as_deref().unwrap_or("") })
                            }
                        }
                        CfRule::Expression { formula, .. } if member == "formula1" => values::vba_text(formula),
                        _ => Err(failure(438, "this condition has no comparison formula")),
                    };
                }
                "delete" => {
                    arguments(call, &[])?;
                    if protected {
                        return Err(failure(1004, "sheet protected"));
                    }
                    session.execute("home.manageRules", json!({"sheet":sheet,"delete":index})).map_err(|e| failure(1004, e.to_string()))?;
                    return Ok(Value::Empty);
                }
                "setfirstpriority" | "setlastpriority" => {
                    arguments(call, &[])?;
                    let n = if member == "setfirstpriority" { 1 } else { rule_count };
                    return self.condition_update(session, sheet, index, json!({"priority":n}));
                }
                "modifyappliestorange" => {
                    let args = arguments(call, &["Range"])?;
                    let Value::Object(reference) = required(&args, 0)? else {
                        return Err(failure(13, "Range required"));
                    };
                    let Object::Range(owner, range) = self.object(session, *reference)? else {
                        return Err(failure(13, "Range required"));
                    };
                    if owner != sheet {
                        return Err(failure(1004, "AppliesTo belongs to another worksheet"));
                    }
                    return self.condition_update(session, sheet, index, json!({"ranges":range.areas.iter().map(|r|r.a1()).collect::<Vec<_>>()}));
                }
                _ => {}
            }
        }
        if call.kind == InvokeKind::Let && matches!(member.as_str(), "priority" | "stopiftrue") {
            let args = arguments(call, &["Value"])?;
            let v = required(&args, 0)?;
            return self.condition_update(
                session,
                sheet,
                index,
                if member == "priority" { json!({"priority":numeric_index(v)?}) } else { json!({"stopIfTrue":super::address::boolean(v)?}) },
            );
        }
        // Native typed rule fields share one serde-checked mutation path.
        let field = match member.as_str() {
            "rank" => "rank",
            "percent" => "percent",
            "bottom" => "bottom",
            "showvalue" | "showicononly" => "show_value",
            "barcolor" => "color",
            "reverseorder" => "reverse",
            _ => return Err(failure(438, "unsupported conditional-format member")),
        };
        let mut rule = serde_json::to_value(&cf.rule).map_err(|e| failure(1004, e.to_string()))?;
        let body = rule
            .as_object_mut()
            .and_then(|o| o.values_mut().next())
            .and_then(Json::as_object_mut)
            .filter(|o| o.contains_key(field))
            .ok_or_else(|| failure(438, "field does not apply to this condition"))?;
        if reads(call) {
            arguments(call, &[])?;
            let v = body.get(field).ok_or_else(|| failure(438, "missing field"))?;
            return if let Some(b) = v.as_bool() {
                Ok(Value::Boolean(if member == "showicononly" { !b } else { b }))
            } else if let Some(n) = v.as_i64() {
                Ok(Value::Long(n as i32))
            } else {
                Err(failure(438, "field requires a specialized child object"))
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(13, "condition assignment requires Let"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        let value = if field == "rank" {
            let n = numeric_index(value)?;
            if !(1..=1000).contains(&n) {
                return Err(failure(1004, "invalid rank"));
            }
            json!(n)
        } else {
            json!(if member == "showicononly" { !super::address::boolean(value)? } else { super::address::boolean(value)? })
        };
        body.insert(field.into(), value);
        self.condition_update(session, sheet, index, json!({"rule":rule}))
    }
    pub(crate) fn condition_update(&self, session: &mut Session, sheet: usize, index: usize, mut update: Json) -> HostResult<Value> {
        update["index"] = json!(index);
        session.execute("home.manageRules", json!({"sheet":sheet,"update":update})).map_err(|e| failure(1004, e.to_string()))?;
        Ok(Value::Empty)
    }
    fn condition_style(&self, session: &mut Session, sheet: usize, index: usize, cf: &CondFormat, font: bool, call: &HostCall) -> HostResult<Value> {
        let mut style = rule_style(&cf.rule)?;
        let member = call.member.to_ascii_lowercase();
        if reads(call) {
            arguments(call, &[])?;
            return match (font, member.as_str()) {
                (true, "bold") => Ok(Value::Boolean(style.font.bold)),
                (true, "italic") => Ok(Value::Boolean(style.font.italic)),
                (true, "name") => values::vba_text(&style.font.name),
                (true, "size") => Ok(Value::Double(f64::from(style.font.size))),
                (true, "color") => super::format_api::ole(style.font.color, session),
                (false, "color") => super::format_api::ole(style.fill.fg, session),
                _ => Err(failure(438, "unsupported differential style member")),
            };
        }
        if call.kind != InvokeKind::Let {
            return Err(failure(13, "style assignment requires Let"));
        }
        let args = arguments(call, &["Value"])?;
        let value = required(&args, 0)?;
        match (font, member.as_str()) {
            (true, "bold") => style.font.bold = super::address::boolean(value)?,
            (true, "italic") => style.font.italic = super::address::boolean(value)?,
            (true, "color") => style.font.color = super::format_api::rgb(value)?,
            (false, "color") => {
                style.fill.fg = super::format_api::rgb(value)?;
                style.fill.pattern = PatternType::Solid;
            }
            _ => return Err(failure(438, "unsupported differential style assignment")),
        };
        let rule = set_rule_style(&cf.rule, &style)?;
        self.condition_update(session, sheet, index, json!({"rule":rule}))
    }
}

pub(crate) fn descriptors() -> Vec<super::MemberDescriptor> {
    vec![
        super::MemberDescriptor {
            owner: "Range",
            name: "FormatConditions",
            signature: "Property Get FormatConditions() As FormatConditions",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "Count",
            signature: "Property Get Count() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "Item",
            signature: "Property Get Item(Index As Variant) As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "Add",
            signature: "Function Add(Type As Long, Optional Operator As Variant, Optional Formula1 As Variant, Optional Formula2 As Variant, Optional String As Variant, Optional TextOperator As Variant, Optional DateOperator As Variant, Optional ScopeType As Variant) As FormatCondition",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "AddColorScale",
            signature: "Function AddColorScale(ColorScaleType As Long) As ColorScale",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "AddDatabar",
            signature: "Function AddDatabar() As Databar",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "AddIconSetCondition",
            signature: "Function AddIconSetCondition() As IconSetCondition",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "AddTop10",
            signature: "Function AddTop10() As Top10",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "AddUniqueValues",
            signature: "Function AddUniqueValues() As UniqueValues",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "AddAboveAverage",
            signature: "Function AddAboveAverage() As AboveAverage",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatConditions",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Type",
            signature: "Property Get Type() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "AppliesTo",
            signature: "Property Get AppliesTo() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Priority",
            signature: "Property Priority As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "StopIfTrue",
            signature: "Property StopIfTrue As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "SetFirstPriority",
            signature: "Sub SetFirstPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "SetLastPriority",
            signature: "Sub SetLastPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "ModifyAppliesToRange",
            signature: "Sub ModifyAppliesToRange(Range As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "Type",
            signature: "Property Get Type() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "AppliesTo",
            signature: "Property Get AppliesTo() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "Priority",
            signature: "Property Priority As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "StopIfTrue",
            signature: "Property StopIfTrue As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "SetFirstPriority",
            signature: "Sub SetFirstPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "SetLastPriority",
            signature: "Sub SetLastPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "ColorScale",
            name: "ModifyAppliesToRange",
            signature: "Sub ModifyAppliesToRange(Range As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "Type",
            signature: "Property Get Type() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "AppliesTo",
            signature: "Property Get AppliesTo() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "Priority",
            signature: "Property Priority As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "StopIfTrue",
            signature: "Property StopIfTrue As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "SetFirstPriority",
            signature: "Sub SetFirstPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "SetLastPriority",
            signature: "Sub SetLastPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "ModifyAppliesToRange",
            signature: "Sub ModifyAppliesToRange(Range As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "Type",
            signature: "Property Get Type() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "AppliesTo",
            signature: "Property Get AppliesTo() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "Priority",
            signature: "Property Priority As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "StopIfTrue",
            signature: "Property StopIfTrue As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "SetFirstPriority",
            signature: "Sub SetFirstPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "SetLastPriority",
            signature: "Sub SetLastPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "ModifyAppliesToRange",
            signature: "Sub ModifyAppliesToRange(Range As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "Type",
            signature: "Property Get Type() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "AppliesTo",
            signature: "Property Get AppliesTo() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "Priority",
            signature: "Property Priority As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "StopIfTrue",
            signature: "Property StopIfTrue As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "SetFirstPriority",
            signature: "Sub SetFirstPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "SetLastPriority",
            signature: "Sub SetLastPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "ModifyAppliesToRange",
            signature: "Sub ModifyAppliesToRange(Range As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "Type",
            signature: "Property Get Type() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "AppliesTo",
            signature: "Property Get AppliesTo() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "Priority",
            signature: "Property Priority As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "StopIfTrue",
            signature: "Property StopIfTrue As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "SetFirstPriority",
            signature: "Sub SetFirstPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "SetLastPriority",
            signature: "Sub SetLastPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "UniqueValues",
            name: "ModifyAppliesToRange",
            signature: "Sub ModifyAppliesToRange(Range As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "Parent",
            signature: "Property Get Parent() As Object",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "Application",
            signature: "Property Get Application() As Application",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "Type",
            signature: "Property Get Type() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "AppliesTo",
            signature: "Property Get AppliesTo() As Range",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "Priority",
            signature: "Property Priority As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "StopIfTrue",
            signature: "Property StopIfTrue As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "Delete",
            signature: "Sub Delete()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "SetFirstPriority",
            signature: "Sub SetFirstPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "SetLastPriority",
            signature: "Sub SetLastPriority()",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "AboveAverage",
            name: "ModifyAppliesToRange",
            signature: "Sub ModifyAppliesToRange(Range As Range)",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Font",
            signature: "Property Get Font() As Font",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Interior",
            signature: "Property Get Interior() As Interior",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Formula1",
            signature: "Property Get Formula1() As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Formula2",
            signature: "Property Get Formula2() As String",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "FormatCondition",
            name: "Operator",
            signature: "Property Get Operator() As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "Rank",
            signature: "Property Rank As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "Percent",
            signature: "Property Percent As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Top10",
            name: "TopBottom",
            signature: "Property TopBottom As Long",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "Databar",
            name: "ShowValue",
            signature: "Property ShowValue As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "ReverseOrder",
            signature: "Property ReverseOrder As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
        super::MemberDescriptor {
            owner: "IconSetCondition",
            name: "ShowIconOnly",
            signature: "Property ShowIconOnly As Boolean",
            availability: super::MemberAvailability::Implemented,
            implementation_symbol: "crates/vba/src/conditional_api.rs::Bridge::conditional_api",
        },
    ]
}
