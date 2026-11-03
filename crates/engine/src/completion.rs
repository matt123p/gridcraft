//! The native column autocomplete algorithm shared by the editor and VBA.
use gridcraft_core::CellRef;
pub fn column_completion(sheet: &gridcraft_model::Sheet, at: CellRef, typed: &str) -> Option<String> {
    if !at.is_valid() || typed.is_empty() || typed.starts_with('=') || typed.chars().all(|c| c.is_ascii_digit() || c == '.' || c == '-') {
        return None;
    }
    let lower = typed.to_lowercase();
    let mut found: Option<String> = None;
    let mut ambiguous = false;
    let mut scan = |row: u32| -> bool {
        match sheet.value(CellRef::new(row, at.col)) {
            gridcraft_core::Value::Empty => false,
            gridcraft_core::Value::Text(t) => {
                let folded = t.to_lowercase();
                if folded.starts_with(&lower) && folded.len() > lower.len() {
                    match &found {
                        Some(f) if f.to_lowercase() != folded => ambiguous = true,
                        None => found = Some(t.to_string()),
                        _ => {}
                    }
                }
                true
            }
            _ => true,
        }
    };
    let mut r = at.row;
    let mut n = 0;
    while r > 0 && n < 5000 {
        r -= 1;
        n += 1;
        if !scan(r) {
            break;
        }
    }
    let mut r = at.row.saturating_add(1);
    n = 0;
    while n < 5000 && r < gridcraft_core::MAX_ROWS && scan(r) {
        r += 1;
        n += 1;
    }
    if ambiguous { None } else { found }
}
