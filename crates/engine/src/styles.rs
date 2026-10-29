//! Bounded native formatting shares the command's protection and undo path.
use crate::{EngineError, Result, Session, cmd};
use gridcraft_core::RangeRef;
use gridcraft_model::{Color, NumFmt, Style};
use serde_json::{Value as Json, json};
#[derive(Clone, Debug)]
pub enum StyleChange {
    NumberFormat(String),
    Bold(bool),
    Italic(bool),
    FontName(String),
    FontSize(f32),
    FontColor(u32),
    InteriorColor(u32),
}
impl StyleChange {
    fn patch(&self) -> Json {
        match self {
            Self::NumberFormat(v) => json!({"numFmt": v}),
            Self::Bold(v) => json!({"font":{"bold":v}}),
            Self::Italic(v) => json!({"font":{"italic":v}}),
            Self::FontName(v) => json!({"font":{"name":v,"minor":false}}),
            Self::FontSize(v) => json!({"font":{"size":v}}),
            Self::FontColor(v) => json!({"font":{"color": Self::color(*v)}}),
            Self::InteriorColor(v) => json!({"fill": gridcraft_model::Fill::solid(Self::color(*v))}),
        }
    }
    fn color(v: u32) -> Color {
        Color::rgb((v & 255) as u8, ((v >> 8) & 255) as u8, ((v >> 16) & 255) as u8)
    }
    fn apply(&self, s: &mut Style) {
        match self {
            Self::NumberFormat(v) => s.num_fmt = NumFmt::new(v),
            Self::Bold(v) => s.font.bold = *v,
            Self::Italic(v) => s.font.italic = *v,
            Self::FontName(v) => {
                s.font.name = v.clone();
                s.font.minor = false;
            }
            Self::FontSize(v) => s.font.size = *v,
            Self::FontColor(v) => s.font.color = Self::color(*v),
            Self::InteriorColor(v) => s.fill = gridcraft_model::Fill::solid(Self::color(*v)),
        }
    }
}
impl Session {
    pub fn set_range_style(&mut self, sheet: usize, range: RangeRef, change: StyleChange) -> Result<Json> {
        let bad = || EngineError::Other("Invalid or oversized native style change".into());
        if !range.start.is_valid()
            || !range.end.is_valid()
            || range.start.row > range.end.row
            || range.start.col > range.end.col
            || range.count() > 65536
            || self.doc()?.wb.sheet(sheet).is_none()
        {
            return Err(bad());
        }
        match &change {
            StyleChange::NumberFormat(s) if s.len() > 4096 => return Err(bad()),
            StyleChange::FontName(s) if s.is_empty() || s.len() > 255 => return Err(bad()),
            StyleChange::FontSize(n) if !n.is_finite() || !(1.0..=409.0).contains(n) => return Err(bad()),
            StyleChange::FontColor(n) | StyleChange::InteriorColor(n) if *n > 0xffffff => return Err(bad()),
            _ => (),
        }
        self.execute_with(
            "home.formatCells",
            || std::borrow::Cow::Owned(json!({"sheet":sheet,"range":range.a1(),"style":change.patch()})),
            |s, _| cmd::format::apply_style_to(s, sheet, &[range], |style| change.apply(style)),
        )
    }
}

impl Session {
    pub fn rename_worksheet(&mut self, sheet: usize, name: &str) -> Result<Json> {
        self.execute_with(
            "sheet.rename",
            || std::borrow::Cow::Owned(json!({"sheet":sheet,"name":name})),
            |s, _| cmd::sheet::rename_sheet_native(s, sheet, name),
        )
    }
    pub fn calculate_workbook(&mut self) -> Result<Json> {
        self.execute_with("formulas.calculateNow", || std::borrow::Cow::Owned(json!({})), |s, _| cmd::formulas::calc_now(s, &Json::Null))
    }
    pub fn set_calculation_mode(&mut self, mode: gridcraft_model::CalcMode) -> Result<Json> {
        self.execute_with("formulas.calculationOptions", || std::borrow::Cow::Owned(json!({"mode":match mode { gridcraft_model::CalcMode::Manual=>"manual", gridcraft_model::CalcMode::AutomaticExceptTables=>"automaticExceptTables", _=>"automatic" }})), |s, _| cmd::edit(s, |cx| { cx.wb.calc.mode = mode; if mode != gridcraft_model::CalcMode::Manual { cx.structural=true; } Ok(Json::Null) }))
    }
}

impl Session {
    pub fn application_calculation_mode(&self) -> Result<gridcraft_model::CalcMode> {
        Ok(self.calculation_mode.unwrap_or(self.doc()?.wb.calc.mode))
    }
    pub fn set_application_calculation_mode(&mut self, mode: gridcraft_model::CalcMode) -> Result<()> {
        if self.documents().len() > 64 {
            return Err(EngineError::Other("calculation document budget".into()));
        }
        let documents: Vec<_> = self.documents().iter().map(|d| d.uid).collect();
        for doc in documents {
            self.with_target(crate::Target::Doc(doc), |s| s.set_calculation_mode(mode))??;
        }
        self.calculation_mode = Some(mode);
        Ok(())
    }
    pub fn calculate_application(&mut self) -> Result<()> {
        if self.documents().len() > 64 {
            return Err(EngineError::Other("calculation document budget".into()));
        }
        let documents: Vec<_> = self.documents().iter().map(|d| d.uid).collect();
        for doc in documents {
            self.with_target(crate::Target::Doc(doc), |s| s.calculate_workbook())??;
        }
        Ok(())
    }
}
