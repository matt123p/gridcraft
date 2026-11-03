//! Source-backed capability review of the High extensions outside accepted OM20 obligations.
use serde_json::{Value, json};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::Path,
};

const LEDGER: &str = "docs/vba-object-model-high-extension-review.json";
const REPORT: &str = "docs/vba-object-model-high-extension-review.md";
const STATUSES: &[&str] = &["ELIGIBLE_NATIVE", "NATIVE_SUBSET", "NO_NATIVE_CAPABILITY", "HOST_SERVICE_REVIEW", "GENERIC_DISPATCH"];
fn array<'a>(v: &'a Value, field: &str) -> Result<&'a Vec<Value>, String> {
    v[field].as_array().ok_or_else(|| format!("High extension review missing {field}"))
}
fn text<'a>(v: &'a Value, field: &str) -> Result<&'a str, String> {
    v[field].as_str().filter(|s| !s.is_empty()).ok_or_else(|| format!("High extension review missing {field}"))
}
fn validate(root: &Path, rows: &[Value], ledger: &Value) -> Result<BTreeMap<String, Value>, String> {
    if ledger["documentation_commit"] != super::vba_audit::COMMIT || ledger["packet"] != "OM20-HIGH-REVIEW" || ledger["schema"] != 1 {
        return Err("High extension review provenance mismatch".into());
    }
    let expected = rows
        .iter()
        .filter(|r| r["v1_disposition"] == "HIGH_REFERENCE_EXTENSION")
        .map(|r| text(r, "source_path").map(str::to_owned))
        .collect::<Result<BTreeSet<_>, _>>()?;
    let mut groups = BTreeSet::new();
    let mut reviews = BTreeMap::new();
    for g in array(ledger, "groups")? {
        let id = text(g, "id")?;
        if !groups.insert(id) || !STATUSES.contains(&text(g, "status")?) {
            return Err(format!("invalid or duplicate High extension group {id}"));
        }
        text(g, "native_backing")?;
        text(g, "boundary")?;
        if array(g, "evidence")?.is_empty() || array(g, "source_paths")?.is_empty() {
            return Err(format!("empty High extension group {id}"));
        }
        for e in array(g, "evidence")? {
            let path = text(e, "path")?;
            if !path.starts_with("crates/") || path.contains("..") || !path.ends_with(".rs") {
                return Err(format!("invalid High extension evidence path {path}"));
            }
            let source = fs::read_to_string(root.join(path)).map_err(|e| e.to_string())?;
            if !source.contains(text(e, "anchor")?) {
                return Err(format!("stale High extension evidence: {id}: {path}"));
            }
        }
        for p in array(g, "source_paths")? {
            let path = p.as_str().ok_or("High extension source path must be text")?;
            if !expected.contains(path) || reviews.insert(path.into(), g.clone()).is_some() {
                return Err(format!("unknown, resolved or duplicate High extension source path: {path}"));
            }
        }
    }
    if reviews.keys().cloned().collect::<BTreeSet<_>>() != expected {
        return Err("High extension capability review does not cover every remaining High source page".into());
    }
    let mut queue_ids = BTreeSet::new();
    let mut queue_orders = BTreeSet::new();
    for q in array(ledger, "next_queue")? {
        if !queue_ids.insert(text(q, "id")?) || !queue_orders.insert(q["order"].as_u64().ok_or("High extension queue order missing")?) {
            return Err("duplicate High extension queue ID/order".into());
        }
        text(q, "goal")?;
        if array(q, "groups")?.is_empty() || array(q, "groups")?.iter().any(|g| g.as_str().is_none_or(|id| !groups.contains(id))) {
            return Err("High extension queue references unknown review group".into());
        }
    }
    Ok(reviews)
}

pub(super) fn reconcile(root: &Path, rows: &mut [Value], profile: &mut Value, write: bool) -> Result<(), String> {
    let ledger: Value = serde_json::from_slice(&fs::read(root.join(LEDGER)).map_err(|e| e.to_string())?).map_err(|e| e.to_string())?;
    let reviews = validate(root, rows, &ledger)?;
    let mut counts = BTreeMap::<String, usize>::new();
    for row in rows.iter_mut() {
        if let Some(group) = reviews.get(text(row, "source_path")?) {
            let status = text(group, "status")?;
            *counts.entry(status.into()).or_default() += 1;
            row["extension_capability_review"] = json!({"group":group["id"],"status":status,"ledger":LEDGER,
                "native_backing":group["native_backing"],"boundary":group["boundary"],"evidence":group["evidence"]});
            row["eligibility"] = json!(match status {
                "ELIGIBLE_NATIVE" => "ELIGIBLE_EXISTING",
                "NO_NATIVE_CAPABILITY" => "OUT_OF_SCOPE_GRIDCRAFT",
                "HOST_SERVICE_REVIEW" => "HOST_SERVICE_REVIEW_REQUIRED",
                "GENERIC_DISPATCH" => "GENERIC_DISPATCH_METADATA_REVIEW",
                _ => "PARTIAL_NATIVE_BACKING",
            });
            row["profile_review"] = json!(format!(
                "High extension capability review {}: {status}; not a new accepted runtime obligation. {}",
                text(group, "id")?,
                text(group, "boundary")?
            ));
        }
    }
    profile["high_extension_capability_review"] = json!({"ledger":LEDGER,"report":REPORT,"reviewed":ledger["reviewed"],
        "source_pages":reviews.len(),"groups":array(&ledger,"groups")?.len(),"status_counts":counts,
        "next_queue":ledger["next_queue"],"evidence_boundary":ledger["evidence_boundary"]});
    let mut out = format!(
        "# High VBA extension capability review\n\nReviewed {} against the pinned public reference and native source. **{} source pages** in **{} review groups**, outside the accepted OM20 runtime obligations. Priorities, source identities and literal catalogue result states remain unchanged. These are capability decisions and projection candidates, not additional accepted runtime contracts.\n\n| Review result | Source pages |\n|---|---:|\n",
        text(&ledger, "reviewed")?,
        reviews.len(),
        array(&ledger, "groups")?.len()
    );
    for (status, count) in &counts {
        out.push_str(&format!("| `{status}` | {count} |\n"));
    }
    out.push_str("\nNative-subset rows have related state or effects but still lack a named part of the Excel contract; they are not all eligible members. Host-service rows require an existing usable host seam. Exclusions are source-inspection judgments; checked anchors prove evidence freshness, not mathematical absence of every possible implementation.\n\n## Recommended next queue\n\n");
    for q in array(&ledger, "next_queue")? {
        out.push_str(&format!("{}. **{}**: {}\n", q["order"], text(q, "id")?, text(q, "goal")?));
    }
    out.push_str("\n## Candidates with direct native backing\n\n");
    for row in rows.iter().filter(|r| r["extension_capability_review"]["status"] == "ELIGIBLE_NATIVE") {
        out.push_str(&format!("- [{}]({}) — `{}`\n", text(row, "id")?, text(row, "url")?, text(&row["extension_capability_review"], "group")?));
    }
    out.push_str("\n## Review groups and limits\n\nThe [checked ledger](vba-object-model-high-extension-review.json) lists every exact source path and native evidence anchor. The [generated contracts](vba-object-model-contracts.json) attach each review to its original source row. Grouping does not imply every member has the same implementation cost or complete native support.\n\n| Group | Result | Pages | Native backing and boundary |\n|---|---|---:|---|\n");
    for g in array(&ledger, "groups")? {
        let backing = text(g, "native_backing")?.replace('|', "\\|");
        let boundary = text(g, "boundary")?.replace('|', "\\|");
        out.push_str(&format!(
            "| `{}` | `{}` | {} | {} {} |\n",
            text(g, "id")?,
            text(g, "status")?,
            array(g, "source_paths")?.len(),
            backing,
            boundary
        ));
    }
    out.push_str("\nNo runtime implementation or fresh Excel observation was performed for this review. Existing OM20 runtime acceptance and native/WASM evidence remain in the [profile sign-off](vba-object-model-profile-report.md). Proposed slices still need exact signatures, options, child/default handles, stale/foreign-object errors and appropriate native/WASM, protection, undo, persistence and UI evidence before delivery.\n");
    if write {
        fs::write(root.join(REPORT), out).map_err(|e| e.to_string())?;
    } else if fs::read_to_string(root.join(REPORT)).map_err(|e| e.to_string())? != out {
        return Err("High extension capability report stale; run vba-audit --write".into());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn capability_labels_do_not_promote_runtime_or_lower_priority() {
        let root = crate::root();
        let doc: Value = serde_json::from_slice(&fs::read(root.join("docs/vba-object-model-contracts.json")).unwrap()).unwrap();
        let mut rows = doc["pages"].as_array().unwrap().clone();
        let before = rows.clone();
        let mut profile = json!({});
        reconcile(&root, &mut rows, &mut profile, false).unwrap();
        for (before, after) in before.iter().zip(&rows) {
            for field in ["id", "source_path", "result_state", "v1_priority", "v1_disposition", "required_for_profile", "implementation_symbol"] {
                assert_eq!(before[field], after[field], "review changed {field}: {}", before["id"]);
            }
        }
        assert_eq!(profile["high_extension_capability_review"]["source_pages"], 885);
        let unsupported = rows.iter().find(|r| r["id"] == "worksheetfunction.webservice").unwrap();
        assert_eq!(unsupported["eligibility"], "OUT_OF_SCOPE_GRIDCRAFT");
        assert_eq!(unsupported["result_state"], "MISSING");
    }
    #[cfg(feature = "engine")]
    #[test]
    fn generic_function_review_matches_the_live_calculator() {
        let root = crate::root();
        let doc: Value = serde_json::from_slice(&fs::read(root.join("docs/vba-object-model-contracts.json")).unwrap()).unwrap();
        let mut count = 0;
        for row in doc["pages"].as_array().unwrap().iter().filter(|r| r["extension_capability_review"]["status"] == "GENERIC_DISPATCH") {
            let function = row["member"].as_str().unwrap().replace('_', ".").to_ascii_uppercase();
            assert!(gridcraft_engine::calc::is_known_function(&function), "no calculator registration: {function}");
            count += 1;
        }
        assert_eq!(count, 50);
    }
    #[test]
    fn capability_review_rejects_uncovered_duplicate_and_stale_evidence() {
        let root = crate::root();
        let doc: Value = serde_json::from_slice(&fs::read(root.join("docs/vba-object-model-contracts.json")).unwrap()).unwrap();
        let rows = doc["pages"].as_array().unwrap();
        let ledger: Value = serde_json::from_slice(&fs::read(root.join(LEDGER)).unwrap()).unwrap();
        assert_eq!(validate(&root, rows, &ledger).unwrap().len(), 885);
        let mut bad = ledger.clone();
        bad["groups"].as_array_mut().unwrap().pop();
        assert!(validate(&root, rows, &bad).unwrap_err().contains("does not cover"));
        let mut bad = ledger.clone();
        let duplicate = bad["groups"][0]["source_paths"][0].clone();
        bad["groups"][0]["source_paths"].as_array_mut().unwrap().push(duplicate);
        assert!(validate(&root, rows, &bad).unwrap_err().contains("duplicate"));
        let mut bad = ledger;
        bad["groups"][0]["evidence"][0]["anchor"] = json!("missing native evidence marker");
        assert!(validate(&root, rows, &bad).unwrap_err().contains("stale High extension evidence"));
    }
}
