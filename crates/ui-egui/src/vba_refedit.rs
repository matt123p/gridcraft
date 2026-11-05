//! Host-owned worksheet selection input for the shared native RefEdit control.
use super::{Editor, Result};
use gridcraft_engine::{
    Session,
    core::{CellRef, col_to_letters, quote_sheet},
};
use gridcraft_vba::vba_ide::{FormInstance, VbaString};
pub(super) struct Picker {
    owner: FormInstance,
    control: u32,
    selection: String,
}
fn selected(session: &Session) -> Result<String> {
    let doc = session.doc().map_err(super::error)?;
    let sheet = doc.wb.sheets.get(doc.wb.active_sheet()).ok_or("No active worksheet")?;
    if doc.selection.ranges.is_empty() || doc.selection.ranges.len() > 128 {
        return Err("RefEdit selection area limit".into());
    }
    let cell = |c: CellRef| format!("${}${}", col_to_letters(c.col), u64::from(c.row) + 1);
    let mut parts = Vec::new();
    for r in &doc.selection.ranges {
        if !r.start.is_valid() || !r.end.is_valid() {
            return Err("Invalid worksheet selection".into());
        }
        let address = if r.is_full_cols() {
            format!("${}:${}", col_to_letters(r.start.col), col_to_letters(r.end.col))
        } else if r.is_full_rows() {
            format!("${}:${}", u64::from(r.start.row) + 1, u64::from(r.end.row) + 1)
        } else if r.start == r.end {
            cell(r.start)
        } else {
            format!("{}:{}", cell(r.start), cell(r.end))
        };
        parts.push(format!("{}!{address}", quote_sheet(&sheet.name)));
    }
    let text = parts.join(",");
    if text.len() > 16384 {
        return Err("RefEdit reference length limit".into());
    }
    Ok(text)
}
pub(super) fn before_command(editor: &Editor, id: &str) -> Result<()> {
    if id.starts_with("vba.") {
        return Ok(());
    }
    if editor.ide.form_range_selection().map_err(super::error)?.is_none() {
        return Ok(());
    }
    if id.starts_with("selection.")
        || id.starts_with("vba.")
        || matches!(id, "sheet.activate" | "sheet.select" | "sheet.next" | "sheet.previous" | "ui.scroll" | "view.zoom" | "view.zoomToSelection")
    {
        return Ok(());
    }
    Err("A RefEdit range selection is active; restore the dialog or reset VBA before running this command.".into())
}
pub(super) fn poll(editor: &mut Editor, session: &mut Session) -> Result<()> {
    editor.view.range_selection_enabled = editor.bridge.permitted(session) && editor.bridge.current(session);
    let Some((owner, control)) = editor.ide.form_range_selection().map_err(super::error)? else {
        editor.refedit_picker = None;
        return Ok(());
    };
    if !editor.view.range_selection_enabled {
        return Err("RefEdit workbook changed or execution consent was revoked.".into());
    }
    let text = selected(session)?;
    let current = editor.refedit_picker.as_ref().filter(|p| p.owner == owner && p.control == control);
    let Some(previous) = current else {
        editor.refedit_picker = Some(Picker { owner, control, selection: text });
        return Ok(());
    };
    if previous.selection == text || !editor.ide.form_instance_can_deliver(owner) {
        return Ok(());
    }
    let value = VbaString::from_utf8(&text).map_err(super::error)?;
    let length = u32::try_from(value.units().len()).map_err(super::error)?;
    editor.bridge.begin(session).map_err(super::error)?;
    editor.ide.form_instance_text(owner, control, value, (0, length), &mut editor.bridge.host(session)).map_err(super::error)?;
    editor.refedit_picker = Some(Picker { owner, control, selection: text });
    Ok(())
}
