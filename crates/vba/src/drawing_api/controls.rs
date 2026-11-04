use super::{
    Bridge, Feature, HostCall, HostResult, InvokeKind, Json, Session, Value, View, absent, arguments, child, failure, json, numeric_index, reads,
    required, root, string, values,
};
use gridcraft_engine::model::form_controls::effective_control;
impl Bridge {
    pub(super) fn add_form_control(&mut self, s: &mut Session, sheet: usize, call: &HostCall) -> HostResult<Value> {
        let a = arguments(call, &["Type", "Left", "Top", "Width", "Height"])?;
        let kind = numeric_index(required(&a, 0)?)?;
        let mut p = json!({"type":kind});
        for (i, key) in [(1, "left"), (2, "top"), (3, "width"), (4, "height")] {
            p[key] = json!(super::super::format_api::number(required(&a, i)?)? * 4. / 3.);
        }
        let result = self.drawing_command(s, sheet, "control.add", p)?;
        let id =
            result.get("shape").and_then(Json::as_u64).and_then(|n| u32::try_from(n).ok()).ok_or_else(|| failure(1004, "control was not created"))?;
        let identity = s
            .doc()
            .map_err(|e| failure(91, e.to_string()))?
            .wb
            .sheet(sheet)
            .and_then(|sh| sh.shapes.iter().find(|s| s.id == id))
            .map(|s| s.identity.clone())
            .ok_or_else(|| failure(91, "control"))?;
        self.feature(root(sheet, id, &identity, "shape"))
    }
    pub(super) fn native_control_format(&mut self, s: &mut Session, f: &Feature, v: &View, call: &HostCall) -> HostResult<Value> {
        let sheet = f.sheet.ok_or_else(|| failure(91, "control owner"))?;
        let wb = s.doc().map_err(|e| failure(91, e.to_string()))?.wb.clone();
        let shape = wb.sheet(sheet).and_then(|s| s.shapes.iter().find(|s| s.id == v.id)).ok_or_else(|| failure(91, "control deleted"))?;
        let c = effective_control(&wb, sheet, shape).map_err(|e| failure(1004, e))?;
        let m = call.member.to_ascii_lowercase();
        let list = c.kind.is_list();
        let numeric = c.kind.is_numeric();
        if matches!(
            m.as_str(),
            "list" | "listcount" | "listindex" | "listfillrange" | "additem" | "removeitem" | "removeallitems" | "multiselect" | "dropdownlines"
        ) && !list
        {
            return Err(failure(1004, "property requires a list control"));
        }
        if matches!(m.as_str(), "min" | "max" | "smallchange" | "largechange") && !numeric {
            return Err(failure(1004, "property requires a scroll bar or spinner"));
        }
        if (m == "multiselect" && c.kind != gridcraft_engine::model::FormControlKind::ListBox)
            || (m == "dropdownlines" && c.kind != gridcraft_engine::model::FormControlKind::DropDown)
        {
            return Err(failure(1004, "property does not apply to this control type"));
        }
        if matches!(m.as_str(), "value" | "linkedcell")
            && matches!(c.kind, gridcraft_engine::model::FormControlKind::Button | gridcraft_engine::model::FormControlKind::Label)
            || m == "value" && c.kind == gridcraft_engine::model::FormControlKind::GroupBox
        {
            return Err(failure(1004, "control does not have a linked value"));
        }
        if matches!(m.as_str(), "value" | "listindex") && c.multi_select {
            return Err(failure(1004, "single selected value is unavailable for a multiselect list"));
        }
        if m == "parent" && reads(call) {
            arguments(call, &[])?;
            return self.feature(child(f, "Shape", 0));
        }
        let key = match m.as_str() {
            "value" | "listindex" => "value",
            "enabled" => "enabled",
            "printobject" => "printObject",
            "min" => "min",
            "max" => "max",
            "smallchange" => "smallChange",
            "largechange" => "largeChange",
            "linkedcell" => "linkedCell",
            "listfillrange" => "listFillRange",
            "dropdownlines" => "dropLines",
            "multiselect" => "multiSelect",
            _ => "",
        };
        if m == "list" || matches!(m.as_str(), "additem" | "removeitem" | "removeallitems") {
            let mut items = c.items.clone();
            let mut value = c.value;
            let mut selected = c.selected.clone();
            if m == "list" && call.kind != InvokeKind::Let {
                let a = arguments(call, &["Index"])?;
                if let Some(index) = a.first().filter(|v| !absent(v)) {
                    let n = numeric_index(index)?;
                    let n = n.checked_sub(1).and_then(|n| usize::try_from(n).ok()).ok_or_else(|| failure(1004, "list index"))?;
                    return values::vba_text(items.get(n).ok_or_else(|| failure(1004, "list index"))?);
                }
                if items.is_empty() {
                    return Err(failure(1004, "the list has no entries"));
                }
                let mut array = vba_values::VbaArray::new(vba_values::ScalarKind::Variant, vec![(1, items.len() as i32)], false)
                    .map_err(|e| failure(7, e.to_string()))?;
                for (i, item) in items.iter().enumerate() {
                    array.set(i, values::vba_text(item)?).map_err(|e| failure(7, e.to_string()))?;
                }
                return Ok(Value::Array(array));
            }
            match m.as_str() {
                "list" => {
                    if call.arguments.len() == 1 {
                        let a = arguments(call, &["Value"])?;
                        let Value::Array(array) = required(&a, 0)? else { return Err(failure(13, "List requires an array")) };
                        if array.elements().len() > 1024 {
                            return Err(failure(7, "list item budget"));
                        }
                        items = array.elements().iter().map(string).collect::<HostResult<Vec<_>>>()?;
                        value = 0;
                        selected.clear();
                    } else {
                        let a = arguments(call, &["Index", "Value"])?;
                        let index = numeric_index(required(&a, 0)?)?;
                        let index = index.checked_sub(1).and_then(|n| usize::try_from(n).ok()).ok_or_else(|| failure(1004, "list index"))?;
                        *items.get_mut(index).ok_or_else(|| failure(1004, "list index"))? = string(required(&a, 1)?)?;
                    }
                }
                "additem" => {
                    let a = arguments(call, &["Text", "Index"])?;
                    let text = string(required(&a, 0)?)?;
                    let index = a.get(1).filter(|v| !absent(v)).map(numeric_index).transpose()?.unwrap_or(items.len() as i32 + 1);
                    let index = index
                        .checked_sub(1)
                        .and_then(|n| usize::try_from(n).ok())
                        .filter(|n| *n <= items.len())
                        .ok_or_else(|| failure(1004, "list insertion index"))?;
                    if items.len() >= 1024 {
                        return Err(failure(7, "list item budget"));
                    }
                    for selected in &mut selected {
                        if *selected > index as u32 {
                            *selected += 1;
                        }
                    }
                    items.insert(index, text);
                    if value > index as i32 {
                        value += 1;
                    }
                }
                "removeitem" => {
                    let a = arguments(call, &["Index", "Count"])?;
                    let index = numeric_index(required(&a, 0)?)?;
                    let count = a.get(1).filter(|v| !absent(v)).map(numeric_index).transpose()?.unwrap_or(1);
                    let start = index.checked_sub(1).and_then(|n| usize::try_from(n).ok()).ok_or_else(|| failure(1004, "list removal index"))?;
                    let count = usize::try_from(count).map_err(|_| failure(1004, "list removal count"))?;
                    let end =
                        start.checked_add(count).filter(|end| count > 0 && *end <= items.len()).ok_or_else(|| failure(1004, "list removal range"))?;
                    items.drain(start..end);
                    if value >= index && value < index + count as i32 {
                        value = 0;
                    } else if value >= index + count as i32 {
                        value -= count as i32;
                    }
                    selected.retain(|n| *n < index as u32 || *n >= index as u32 + count as u32);
                    for n in &mut selected {
                        if *n >= index as u32 + count as u32 {
                            *n -= count as u32;
                        }
                    }
                }
                "removeallitems" => {
                    arguments(call, &[])?;
                    items.clear();
                    value = 0;
                    selected.clear();
                }
                _ => {}
            }
            self.drawing_command(s, sheet, "control.set", json!({"id":v.id,"items":items,"value":value,"selected":selected}))?;
            return Ok(Value::Empty);
        }
        if call.kind == InvokeKind::Let {
            if key.is_empty() {
                return Err(failure(438, "ControlFormat member is read-only or unsupported"));
            }
            let a = arguments(call, &["Value"])?;
            let value = required(&a, 0)?;
            let val = match m.as_str() {
                "linkedcell" | "listfillrange" => json!(string(value)?),
                "enabled" | "printobject" => {
                    let Value::Boolean(v) = vba_runtime::bind(vba_values::ScalarKind::Boolean, value).map_err(values::runtime_error)? else {
                        return Err(failure(13, "Boolean required"));
                    };
                    json!(v)
                }
                "multiselect" => match numeric_index(value)? {
                    -4142 => json!(false),
                    -4154 | 3 => json!(true),
                    _ => return Err(failure(1004, "invalid XlEnableSelection value")),
                },
                _ => json!(numeric_index(value)?),
            };
            let mut p = json!({"id":v.id});
            p[key] = val;
            if m == "multiselect" {
                p["extendedSelect"] = json!(numeric_index(value)? == 3);
            }
            self.drawing_command(s, sheet, "control.set", p)?;
            return Ok(Value::Empty);
        }
        arguments(call, &[])?;
        match m.as_str() {
            "value" | "listindex" => Ok(Value::Long(c.value)),
            "enabled" => Ok(Value::Boolean(c.enabled)),
            "printobject" => Ok(Value::Boolean(c.print_object)),
            "min" => Ok(Value::Long(c.min)),
            "max" => Ok(Value::Long(c.max)),
            "smallchange" => Ok(Value::Long(c.small_change)),
            "largechange" => Ok(Value::Long(c.large_change)),
            "linkedcell" => {
                values::vba_text(if c.kind == gridcraft_engine::model::FormControlKind::GroupBox { &c.group_link } else { &c.linked_cell })
            }
            "listfillrange" => values::vba_text(&c.list_fill_range),
            "listcount" => Ok(Value::Long(c.items.len() as i32)),
            "dropdownlines" => Ok(Value::Long(c.drop_lines as i32)),
            "multiselect" => Ok(Value::Long(if c.extended_select {
                3
            } else if c.multi_select {
                -4154
            } else {
                -4142
            })),
            _ => Err(failure(1004, "ControlFormat member outside native control profile")),
        }
    }
}
