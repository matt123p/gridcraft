//! Native worksheet Form Controls share the existing Shape store and lifetime identity.
use crate::{Shape, Sheet, Workbook};
use gridcraft_core::{CellError, CellRef, RangeRef, Value, split_sheet};
use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ControlAttachment {
    pub from: CellRef,
    pub to: CellRef,
    pub fractions: [f32; 4],
}
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum FormControlKind {
    Button,
    CheckBox,
    DropDown,
    GroupBox,
    Label,
    ListBox,
    OptionButton,
    ScrollBar,
    Spinner,
}
impl FormControlKind {
    pub fn from_xml(s: &str) -> Option<Self> {
        Some(match s {
            "Button" => Self::Button,
            "CheckBox" => Self::CheckBox,
            "Drop" => Self::DropDown,
            "GBox" => Self::GroupBox,
            "Label" => Self::Label,
            "List" => Self::ListBox,
            "Radio" => Self::OptionButton,
            "Scroll" => Self::ScrollBar,
            "Spin" => Self::Spinner,
            _ => return None,
        })
    }
    pub fn xml(self) -> &'static str {
        match self {
            Self::Button => "Button",
            Self::CheckBox => "CheckBox",
            Self::DropDown => "Drop",
            Self::GroupBox => "GBox",
            Self::Label => "Label",
            Self::ListBox => "List",
            Self::OptionButton => "Radio",
            Self::ScrollBar => "Scroll",
            Self::Spinner => "Spin",
        }
    }
    pub fn excel_type(self) -> i32 {
        match self {
            Self::Button => 0,
            Self::CheckBox => 1,
            Self::DropDown => 2,
            Self::GroupBox => 4,
            Self::Label => 5,
            Self::ListBox => 6,
            Self::OptionButton => 7,
            Self::ScrollBar => 8,
            Self::Spinner => 9,
        }
    }
    pub fn from_excel(n: i32) -> Option<Self> {
        [Self::Button, Self::CheckBox, Self::DropDown, Self::GroupBox, Self::Label, Self::ListBox, Self::OptionButton, Self::ScrollBar, Self::Spinner]
            .into_iter()
            .find(|k| k.excel_type() == n)
    }
    pub fn is_list(self) -> bool {
        matches!(self, Self::ListBox | Self::DropDown)
    }
    pub fn is_numeric(self) -> bool {
        matches!(self, Self::ScrollBar | Self::Spinner)
    }
}

pub fn resolve_binding(wb: &Workbook, owner: usize, address: &str) -> Result<(usize, RangeRef), String> {
    let mut current = address.trim().trim_start_matches('=').to_owned();
    let mut scope = owner;
    for _ in 0..8 {
        if current.len() > 1024 || current.contains(['[', ']']) {
            return Err("external or oversized control reference".into());
        }
        let (sheet, body) = split_sheet(&current);
        let index =
            sheet.as_deref().map(|n| wb.sheet_index(n).ok_or_else(|| "control reference sheet missing".to_owned())).transpose()?.unwrap_or(scope);
        if let Some(range) = RangeRef::parse(body) {
            if !range.start.is_valid() || !range.end.is_valid() {
                return Err("control range bounds".into());
            }
            return Ok((index, range));
        }
        let name = wb.name(body, index).ok_or_else(|| "invalid control cell/range/name reference".to_owned())?;
        current = name.formula.clone();
        scope = name.scope.unwrap_or(index);
    }
    Err("control name-reference depth limit".into())
}
fn position(sh: &Sheet, shape: &Shape) -> (f64, f64, f64, f64) {
    let a = shape.anchor;
    (sh.col_left(a.cell.col) + f64::from(a.dx), sh.row_top(a.cell.row) + f64::from(a.dy), f64::from(a.width), f64::from(a.height))
}
/// Native option groups: smallest enclosing GroupBox, then the VML firstButton run.
fn radio_keys(sh: &Sheet) -> Vec<(u32, u32)> {
    let boxes: Vec<_> = sh
        .shapes
        .iter()
        .filter(|s| s.format.form_control.as_ref().is_some_and(|c| c.kind == FormControlKind::GroupBox))
        .map(|s| (s.id, position(sh, s)))
        .collect();
    let mut first = 0;
    sh.shapes
        .iter()
        .filter_map(|s| {
            let c = s.format.form_control.as_ref()?;
            if c.kind != FormControlKind::OptionButton {
                return None;
            }
            if c.first_button {
                first = s.id;
            }
            let (x, y, w, h) = position(sh, s);
            let (cx, cy) = (x + w / 2., y + h / 2.);
            let containing = boxes
                .iter()
                .filter(|(_, (a, b, w, h))| cx >= *a && cx <= a + w && cy >= *b && cy <= b + h)
                .min_by(|a, b| (a.1.2 * a.1.3).total_cmp(&(b.1.2 * b.1.3)));
            Some((s.id, containing.map_or(first, |b| b.0)))
        })
        .collect()
}
pub fn radio_group(sh: &Sheet, id: u32) -> Vec<u32> {
    let keys = radio_keys(sh);
    let Some((_, key)) = keys.iter().find(|(i, _)| *i == id) else { return vec![] };
    keys.iter().filter(|(_, k)| k == key).map(|(i, _)| *i).collect()
}
fn radio_link(sh: &Sheet, id: u32) -> Option<&str> {
    let key = radio_keys(sh).into_iter().find(|(i, _)| *i == id)?.1;
    sh.shapes
        .iter()
        .find(|s| s.id == key)
        .and_then(|s| s.format.form_control.as_ref())
        .filter(|c| c.kind == FormControlKind::GroupBox && !c.group_link.is_empty())
        .map(|c| c.group_link.as_str())
}
pub fn effective_control(wb: &Workbook, owner: usize, shape: &Shape) -> Result<FormControl, String> {
    let mut c = shape.format.form_control.clone().ok_or("not a worksheet Form Control")?;
    if !c.list_fill_range.is_empty() {
        if !c.kind.is_list() {
            return Err("ListFillRange requires a list control".into());
        }
        let (si, r) = resolve_binding(wb, owner, &c.list_fill_range)?;
        if r.count() > 1024 || r.width() != 1 {
            return Err("control list supports at most 1024 rows in one column".into());
        }
        let sh = wb.sheet(si).ok_or("list source sheet")?;
        c.items = r.iter().map(|cell| sh.value(cell).display()).collect();
    }
    if c.kind == FormControlKind::OptionButton
        && let Some(link) = wb.sheet(owner).and_then(|sh| radio_link(sh, shape.id))
    {
        c.linked_cell = link.into();
    }
    if !c.linked_cell.is_empty() && !c.multi_select {
        let (si, r) = resolve_binding(wb, owner, &c.linked_cell)?;
        if r.count() != 1 {
            return Err("LinkedCell requires a single cell".into());
        }
        let value = wb.sheet(si).ok_or("linked sheet")?.value(r.start);
        c.value = match (&c.kind, value) {
            (FormControlKind::CheckBox, Value::Bool(v)) => {
                if v {
                    1
                } else {
                    -4146
                }
            }
            (FormControlKind::CheckBox, Value::Empty) => -4146,
            (FormControlKind::CheckBox, Value::Error(CellError::NA)) => 2,
            (FormControlKind::CheckBox, Value::Number(v)) => {
                if v == 0. {
                    -4146
                } else {
                    1
                }
            }
            (FormControlKind::CheckBox, _) => return Err("invalid linked checkbox value".into()),
            (FormControlKind::OptionButton, Value::Number(v)) => {
                let sh = wb.sheet(owner).ok_or("control owner")?;
                let group = radio_group(sh, shape.id);
                if group.iter().position(|n| *n == shape.id).is_some_and(|i| v == (i + 1) as f64) { 1 } else { -4146 }
            }
            (FormControlKind::OptionButton, Value::Empty) => -4146,
            (k, Value::Empty) if k.is_numeric() => c.min,
            (_, Value::Empty) => 0,
            (_, Value::Number(n)) if n.is_finite() && n.fract() == 0. && (0. ..=30000.).contains(&n) => n as i32,
            _ => return Err("invalid linked control value".into()),
        };
    }
    c.validate()?;
    if c.kind.is_list() && (c.value as usize > c.items.len() || c.selected.iter().any(|n| *n == 0 || *n as usize > c.items.len())) {
        return Err("list selection outside item bounds".into());
    }
    Ok(c)
}
pub fn linked_write(wb: &Workbook, owner: usize, shape: &Shape, c: &FormControl) -> Result<Option<(usize, CellRef, Value)>, String> {
    if c.linked_cell.is_empty() || c.multi_select {
        return Ok(None);
    }
    let (si, r) = resolve_binding(wb, owner, &c.linked_cell)?;
    if r.count() != 1 {
        return Err("LinkedCell requires one cell".into());
    }
    let value = match c.kind {
        FormControlKind::CheckBox => match c.value {
            1 => Value::Bool(true),
            -4146 => Value::Bool(false),
            2 => Value::Error(CellError::NA),
            _ => return Err("invalid checkbox state".into()),
        },
        FormControlKind::OptionButton => {
            if c.value != 1 {
                return Ok(None);
            }
            let sh = wb.sheet(owner).ok_or("owner sheet")?;
            let index = radio_group(sh, shape.id).iter().position(|id| *id == shape.id).ok_or("radio group")?;
            Value::Number((index + 1) as f64)
        }
        _ => Value::Number(f64::from(c.value)),
    };
    Ok(Some((si, r.start, value)))
}
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct FormControl {
    pub kind: FormControlKind,
    pub name: String,
    pub enabled: bool,
    pub visible: bool,
    pub print_object: bool,
    /// Excel ControlFormat.Value: checkbox 1/-4146/2; radio 1/-4146; list one-based.
    pub value: i32,
    pub min: i32,
    pub max: i32,
    pub small_change: i32,
    pub large_change: i32,
    pub horizontal: bool,
    pub drop_lines: u32,
    pub linked_cell: String,
    pub list_fill_range: String,
    pub items: Vec<String>,
    pub multi_select: bool,
    #[serde(default)]
    pub extended_select: bool,
    pub selected: Vec<u32>,
    pub first_button: bool,
    pub group_link: String,
}
impl FormControl {
    pub fn new(kind: FormControlKind, name: String) -> Self {
        Self {
            kind,
            name,
            enabled: true,
            visible: true,
            print_object: true,
            value: if matches!(kind, FormControlKind::CheckBox | FormControlKind::OptionButton) { -4146 } else { 0 },
            min: 0,
            max: 100,
            small_change: 1,
            large_change: 10,
            horizontal: false,
            drop_lines: 8,
            linked_cell: String::new(),
            list_fill_range: String::new(),
            items: vec![],
            multi_select: false,
            extended_select: false,
            selected: vec![],
            first_button: false,
            group_link: String::new(),
        }
    }
    pub fn validate(&self) -> Result<(), String> {
        if ![&self.name, &self.linked_cell, &self.list_fill_range, &self.group_link].into_iter().all(|s| valid_control_text(s))
            || self.items.iter().any(|s| !valid_control_text(s))
        {
            return Err("native control text contains an invalid XML character".into());
        }
        if self.name.is_empty()
            || self.name.len() > 256
            || self.name.contains('\0')
            || self.linked_cell.len() > 1024
            || self.list_fill_range.len() > 1024
            || self.group_link.len() > 1024
            || self.items.len() > 1024
            || self.selected.len() > 1024
            || self.items.iter().any(|s| s.len() > 16384)
            || self.items.iter().fold(0usize, |n, s| n.saturating_add(s.len())) > 1024 * 1024
        {
            return Err("form control metadata budget".into());
        }
        if self.drop_lines > 30000
            || self.min < 0
            || self.max > 30000
            || self.min > self.max
            || !(0..=30000).contains(&self.small_change)
            || !(0..=30000).contains(&self.large_change)
        {
            return Err("invalid native control range or increment".into());
        }
        if self.kind.is_numeric() && !(self.min..=self.max).contains(&self.value) {
            return Err("control value outside range".into());
        }
        if matches!(self.kind, FormControlKind::CheckBox | FormControlKind::OptionButton) && ![-4146, 1, 2].contains(&self.value) {
            return Err("invalid check state".into());
        }
        if self.kind == FormControlKind::OptionButton && self.value == 2 {
            return Err("radio controls do not support mixed state".into());
        }
        if self.kind.is_list() && !(0..=30000).contains(&self.value) {
            return Err("invalid list selection".into());
        }
        Ok(())
    }
}

/// Office control XML strings cannot encode forbidden XML 1.0 scalar values.
pub fn valid_control_text(s: &str) -> bool {
    s.chars().all(|c| matches!(c,'\t'|'\r'|'\n'|'\u{20}'..='\u{d7ff}'|'\u{e000}'..='\u{fffd}'|'\u{10000}'..='\u{10ffff}'))
}
