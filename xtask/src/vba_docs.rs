//! Keep the reader-facing name indexes in sync with the live adapter descriptors.
use gridcraft_vba::{MemberAvailability, MemberDescriptor};
use serde_json::Value;
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::Path,
};

fn names(names: &BTreeSet<String>) -> String {
    let names = names.iter().collect::<Vec<_>>();
    let mut out = String::new();
    for chunk in names.chunks(8) {
        out.push_str("- ");
        out.push_str(&chunk.iter().map(|name| format!("`{name}`")).collect::<Vec<_>>().join(", "));
        out.push('\n');
    }
    out
}
fn replace_block(source: &str, block: &str, body: &str) -> Result<String, String> {
    let start = format!("<!-- BEGIN {block} -->");
    let end = format!("<!-- END {block} -->");
    if source.matches(&start).count() != 1 || source.matches(&end).count() != 1 {
        return Err(format!("missing/duplicate VBA documentation marker: {block}"));
    }
    let (before, rest) = source.split_once(&start).ok_or("missing VBA block start")?;
    let (_, after) = rest.split_once(&end).ok_or("missing VBA block end")?;
    Ok(format!("{before}{start}\n\n{body}\n{end}{after}"))
}
pub(super) fn emit(root: &Path, rows: &[Value], descriptors: &BTreeMap<(String, String), MemberDescriptor>, write: bool) -> Result<(), String> {
    let mut functions = BTreeSet::new();
    let mut members = BTreeMap::<String, BTreeSet<String>>::new();
    for descriptor in descriptors.values().filter(|d| d.availability == MemberAvailability::Implemented) {
        if descriptor.owner == "WorksheetFunction" && descriptor.signature.starts_with("Function ") {
            functions.insert(descriptor.name.to_owned());
        } else {
            members.entry(descriptor.owner.into()).or_default().insert(descriptor.name.into());
        }
    }
    let generic = rows
        .iter()
        .filter(|r| r["extension_capability_review"]["status"] == "GENERIC_DISPATCH")
        .map(|r| r["member"].as_str().map(str::to_owned).ok_or("generic function member missing"))
        .collect::<Result<BTreeSet<_>, _>>()?;
    let mut index = String::from("<details>\n<summary>Full implemented object member names</summary>\n\n| Object | Members |\n|---|---|\n");
    for (owner, names) in members {
        index.push_str(&format!("| `{owner}` | {} |\n", names.iter().map(|name| format!("`{name}`")).collect::<Vec<_>>().join(", ")));
    }
    index.push_str("\n</details>\n");
    let path = root.join("docs/vba/supported.md");
    let current = fs::read_to_string(&path).map_err(|e| e.to_string())?;
    let out = replace_block(&current, "VBA WORKSHEET FUNCTIONS", &names(&functions))?;
    let out = replace_block(&out, "VBA GENERIC FUNCTIONS", &names(&generic))?;
    let out = replace_block(&out, "VBA OBJECT MEMBERS", &index)?;
    if write {
        fs::write(path, out).map_err(|e| e.to_string())?;
    } else if current != out {
        return Err("VBA supported-name index stale; run vba-audit --write".into());
    }
    Ok(())
}
