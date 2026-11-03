//! Final bounded-profile acceptance; independent catalogue states are never promoted to parity.
use gridcraft_vba::{MemberAvailability, MemberDescriptor};
use serde_json::{Value, json};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs,
    path::Path,
};
fn fold(s: &str) -> String {
    s.to_ascii_lowercase()
}
fn strings(v: &Value) -> Result<Vec<String>, String> {
    v.as_array().ok_or("expected string array")?.iter().map(|v| v.as_str().map(str::to_owned).ok_or_else(|| "expected text".into())).collect()
}
fn member_id(d: &MemberDescriptor) -> String {
    format!("{}.{}", d.owner, d.name)
}
fn types_in(signature: &str) -> Vec<String> {
    signature
        .split(" As ")
        .skip(1)
        .filter_map(|s| {
            let ty = s.split(|c: char| !c.is_ascii_alphanumeric() && c != '_' && c != '.').next().unwrap_or("");
            (!ty.is_empty()).then(|| ty.to_owned())
        })
        .collect()
}
fn result_type(signature: &str) -> Option<String> {
    let (_, tail) = signature.rsplit_once(" As ")?;
    (!tail.contains([')', ',', '('])).then(|| tail.trim().to_owned())
}
fn scalar(ty: &str) -> bool {
    matches!(ty, "variant" | "object" | "string" | "boolean" | "byte" | "integer" | "long" | "single" | "double" | "currency" | "decimal" | "date")
}

pub(super) fn reconcile(root: &Path, rows: &mut [Value], descriptors: &BTreeMap<(String, String), MemberDescriptor>) -> Result<Value, String> {
    let spec: Value = serde_json::from_slice(&fs::read(root.join("docs/vba-object-model-om20-acceptance.json")).map_err(|e| e.to_string())?)
        .map_err(|e| e.to_string())?;
    if spec["profile"] != "GridCraft.Excel.v1"
        || spec["documentation_commit"] != super::vba_audit::COMMIT
        || spec["status"] != "ACCEPTED_BOUNDED_NATIVE_PROFILE"
    {
        return Err("OM20 profile provenance mismatch".into());
    }
    let pin = spec["tiny_vba_revision"].as_str().ok_or("dependency pin missing")?;
    let manifest = fs::read_to_string(root.join("crates/vba/Cargo.toml")).map_err(|e| e.to_string())?;
    let dependencies = manifest.lines().filter(|line| line.contains("git = \"https://github.com/matt123p/tiny-vba.git\"")).collect::<Vec<_>>();
    if dependencies.is_empty() || dependencies.iter().any(|line| !line.contains(&format!("rev = \"{pin}\""))) {
        return Err("tiny-vba revision does not match profile sign-off".into());
    }
    let expected_ledgers = std::iter::once("docs/vba-object-model-high-acceptance.json".to_owned())
        .chain(["om12", "om13", "om08", "om09", "om10", "om19", "om15"].map(|p| format!("docs/vba-object-model-{p}-acceptance.json")))
        .collect::<BTreeSet<_>>();
    let declared_ledgers = strings(&spec["accepted_ledgers"])?;
    if declared_ledgers.iter().cloned().collect::<BTreeSet<_>>() != expected_ledgers || declared_ledgers.len() != 8 {
        return Err("OM20 must retain all eight accepted native ledgers".into());
    }
    if strings(&spec["roots"])? != ["Application"] {
        return Err("OM20 graph must use the actual Application root".into());
    }
    let routing = spec["routing_corrections"]
        .as_array()
        .ok_or("routing review missing")?
        .iter()
        .map(|r| Ok((fold(r["id"].as_str().ok_or("routing ID missing")?), r["resolved_packet"].as_str().ok_or("routing packet missing")?.to_owned())))
        .collect::<Result<BTreeMap<_, _>, String>>()?;
    let mut required = BTreeMap::<String, BTreeSet<String>>::new();
    for ledger in strings(&spec["accepted_ledgers"])? {
        let doc: Value = serde_json::from_slice(&fs::read(root.join(&ledger)).map_err(|e| e.to_string())?).map_err(|e| e.to_string())?;
        let groups =
            doc.get("required_obligations").or_else(|| doc.get("requirements")).and_then(Value::as_array).ok_or("profile obligations missing")?;
        for g in groups {
            if g.get("required") == Some(&Value::Bool(false)) {
                continue;
            }
            for id in strings(&g["source_ids"])? {
                required.entry(fold(&id)).or_default().insert(ledger.clone());
            }
        }
    }
    for test in strings(&spec["test_ids"])? {
        super::vba_audit::validate_test_reference(root, &test)?;
    }
    let graph = reachability(root, descriptors, &spec)?;
    let mut duplicates = BTreeMap::<String, Vec<String>>::new();
    for row in rows.iter() {
        duplicates
            .entry(fold(row["id"].as_str().ok_or("page ID missing")?))
            .or_default()
            .push(row["source_path"].as_str().ok_or("source path missing")?.into());
    }
    let mut resolved_required = BTreeSet::new();
    let mut counts = BTreeMap::<String, usize>::new();
    let mut priorities = BTreeMap::<String, usize>::new();
    let mut legacy = 0;
    let mut seen = BTreeSet::new();
    let mut exceptions = BTreeSet::new();
    let allowed_empty = strings_array_ids(&spec["source_less_enum_values"])?;
    for row in rows.iter_mut() {
        let id = row["id"].as_str().ok_or("ID missing")?.to_owned();
        let path = row["source_path"].as_str().ok_or("source path missing")?.to_owned();
        let kind = row["kind"].as_str().unwrap_or("").to_owned();
        if !seen.insert(path.clone()) {
            return Err(format!("duplicate source page {path}"));
        }
        let packet = row["packet"].as_str().ok_or("source family missing")?;
        let family = packet.strip_prefix("OM").and_then(|n| n.parse::<u32>().ok());
        if packet != "OM00" && !family.is_some_and(|n| (2..=19).contains(&n)) {
            return Err(format!("unreviewed source family {packet}: {id}"));
        }
        let key = fold(&id);
        if row["member"].as_str().is_some_and(|m| m.eq_ignore_ascii_case("Creator")) && !row["implementation_symbol"].is_string() {
            row["result_state"] = json!("OUT_OF_SCOPE_GRIDCRAFT");
            row["native_capability_gate"] =
                json!("GridCraft does not expose an Office application creator signature; no Excel vendor identity is fabricated.");
        }
        if kind == "object"
            && row["source_namespace"] == "Excel"
            && !row["implementation_symbol"].is_string()
            && let Some(witness) = graph["root_witnesses"]
                .as_array()
                .and_then(|ws| ws.iter().find(|w| w["type"].as_str().is_some_and(|t| t.eq_ignore_ascii_case(row["owner"].as_str().unwrap_or("")))))
        {
            row["result_state"] = json!("PARTIAL");
            row["profile"] = json!("GridCraft.Excel.v1");
            row["implementation_symbol"] = json!("crates/vba/src/types.rs::references");
            row["dispatch_evidence"] = json!("reachable nominal receiver; source member/option support remains independent");
            row["receiver_witness"] = witness["path"].clone();
        }
        if kind == "event" && !row["implementation_symbol"].is_string() {
            row["result_state"] = json!("OUT_OF_SCOPE_GRIDCRAFT");
            row["native_capability_gate"] = json!(
                "No subscribed native signal/cancellation seam in the accepted event profile; additional Excel events are not synthesized from polling. See High ledger additional-event-signals."
            );
        }
        let required_ledgers = required.get(&key).filter(|_| kind != "event" || row["implementation_symbol"].is_string());
        if let Some(ledgers) = required_ledgers {
            if row["result_state"] != "PARTIAL" || !row["implementation_symbol"].is_string() || row["test_ids"].as_array().is_none_or(Vec::is_empty) {
                return Err(format!("unresolved required source obligation: {path}"));
            }
            resolved_required.insert(key.clone());
            row["accepted_ledgers"] = json!(ledgers);
        } else {
            row["accepted_ledgers"] = json!([]);
        }
        let namespace = row["source_namespace"].as_str().unwrap_or("Excel").to_owned();
        if row["baseline_status"] == "reference-triage" {
            if namespace != "ExcelGraph" || row["packet"] != "OM17" {
                return Err(format!("unreconciled legacy namespace: {path}"));
            }
            row["result_state"] = json!("REFERENCE_ALIAS");
            row["legacy_reconciliation"] = json!("EXCEL_GRAPH_SEPARATE_NAMESPACE; retained source/API names do not alias the modern Excel receiver");
            legacy += 1;
        }
        if kind == "enumeration" && row["enum_values"].as_array().is_none_or(Vec::is_empty) {
            if !allowed_empty.contains(&key) || row["source_deprecated"] != true {
                return Err(format!("enum without reviewed source-values exception: {id}"));
            }
            exceptions.insert(key.clone());
            row["enum_source_exception"] = json!("PINNED_DEPRECATED_PAGE_HAS_NO_MEMBER_VALUES");
        }
        if let Some(packet) = routing.get(&key) {
            row["resolved_packet"] = json!(packet);
        } else {
            row["resolved_packet"] = row["packet"].clone();
        }
        let priority = priority(row);
        if required_ledgers.is_some() && priority == "LOW" {
            return Err(format!("required native member conflicts with Low deferral; review routing: {id}"));
        }
        let (eligibility, disposition, reason) = disposition(row, &priority, required_ledgers.is_some());
        *counts.entry(disposition.into()).or_default() += 1;
        *priorities.entry(priority.clone()).or_default() += 1;
        row["v1_priority"] = json!(priority);
        row["v1_disposition"] = json!(disposition);
        row["eligibility"] = json!(eligibility);
        row["profile_review"] = json!(reason);
        row["required_for_profile"] = json!(required_ledgers.is_some());
        row["implementation_family"] = json!(format!(
            "{}::{}",
            row["resolved_packet"].as_str().unwrap_or(""),
            if namespace == "ExcelGraph" { "ExcelGraph" } else { row["owner"].as_str().unwrap_or("") }
        ));
        row["canonical_contract_key"] = json!(format!("{namespace}::{id}#{kind}"));
        row["source_record_key"] = json!(path);
        row["alias_review"] = json!(if namespace == "ExcelGraph" {
            "LEGACY_NAMESPACE_RETAINED"
        } else if duplicates.get(&key).is_some_and(|v| v.len() > 1) {
            "DISTINCT_SOURCE_KINDS; method/property/event provenance retained"
        } else {
            "CANONICAL_SOURCE_PROVENANCE_RETAINED"
        });
    }
    if exceptions != allowed_empty {
        return Err("stale enum source exception list".into());
    }
    let missing = required.keys().filter(|id| !resolved_required.contains(*id)).cloned().collect::<Vec<_>>();
    if !missing.is_empty() {
        return Err(format!("unresolved required profile IDs: {missing:?}"));
    }
    let collisions=duplicates.into_iter().filter(|(_,paths)|paths.len()>1).map(|(id,paths)|json!({"source_id":id,"source_paths":paths,"resolution":"SOURCE_PATH_AND_KIND_ARE_DISTINCT; no callable/event evidence conflation"})).collect::<Vec<_>>();
    Ok(
        json!({"schema":1,"profile":spec["profile"],"status":"ACCEPTED_BOUNDED_NATIVE_PROFILE","scope":spec["scope"],"tiny_vba_revision":spec["tiny_vba_revision"],"evidence_boundary":spec["evidence_boundary"],"documentation_commit":spec["documentation_commit"],"source_pages":rows.len(),"required_source_ids":resolved_required.len(),"unresolved_required_source_ids":missing,"disposition_counts":counts,"priority_counts":priorities,"legacy_graph_pages":legacy,"source_identity_collisions":collisions,"routing_corrections":spec["routing_corrections"],"enum_source_exceptions":spec["source_less_enum_values"],"accepted_ledgers":spec["accepted_ledgers"],"catalogue_policy":spec["catalogue_policy"],"reachability":graph}),
    )
}
fn strings_array_ids(v: &Value) -> Result<BTreeSet<String>, String> {
    v.as_array().ok_or("expected exception list")?.iter().map(|r| r["id"].as_str().map(fold).ok_or_else(|| "exception ID missing".into())).collect()
}
fn priority(row: &Value) -> String {
    let packet = row.get("resolved_packet").unwrap_or(&row["packet"]).as_str().unwrap_or("");
    let owner = fold(row["owner"].as_str().unwrap_or(""));
    let member = fold(row["member"].as_str().unwrap_or(""));
    if packet == "OM00" {
        return "METADATA".into();
    }
    if matches!(packet, "OM14" | "OM16" | "OM17")
        || owner.contains("connection")
        || owner.contains("querytable")
        || owner.contains("cubefield")
        || owner.starts_with("xml")
        || matches!(
            member.as_str(),
            "providerid"
                | "userid"
                | "vbproject"
                | "vbasigned"
                | "refreshall"
                | "link sources"
                | "linksources"
                | "updatelink"
                | "openlinks"
                | "querytables"
                | "customxmlparts"
        )
    {
        return "LOW".into();
    }
    if packet == "OM11" {
        return if matches!(owner.as_str(), "chart" | "window") || member.contains("pivot") || member.contains("hyperlink") {
            "MEDIUM"
        } else if member.contains("xml") || member.contains("sync") || member.contains("remote") || owner.contains("query") {
            "LOW"
        } else {
            "HIGH"
        }
        .into();
    }
    if packet == "OM09"
        && matches!(
            member.as_str(),
            "protect"
                | "unprotect"
                | "protectcontents"
                | "protectstructure"
                | "protectwindows"
                | "protection"
                | "protectionmode"
                | "protectdrawingobjects"
                | "protectscenarios"
                | "locked"
                | "allowedit"
        )
    {
        return "HIGH".into();
    }
    if matches!(packet, "OM08" | "OM09" | "OM10" | "OM12" | "OM13" | "OM15" | "OM19") {
        return "MEDIUM".into();
    }
    "HIGH".into()
}
fn disposition(row: &Value, priority: &str, required: bool) -> (&'static str, &'static str, String) {
    let packet = row["packet"].as_str().unwrap_or("");
    let id = row["id"].as_str().unwrap_or("");
    let state = row["result_state"].as_str().unwrap_or("");
    if priority == "LOW" {
        return (
            if row["result_state"] == "OUT_OF_SCOPE_GRIDCRAFT" {
                "OUT_OF_SCOPE_GRIDCRAFT"
            } else if row["result_state"] == "PLATFORM_SERVICE_NEEDED" {
                "SERVICE_UNAVAILABLE"
            } else if row["implementation_symbol"].is_string() {
                "ELIGIBLE_EXISTING"
            } else {
                "NOT_REQUIRED_FOR_DEFERRED_LOW"
            },
            "DEFERRED_LOW",
            format!("{packet}: accepted Low deferral; native result_state remains independent for {id}"),
        );
    }
    if priority == "METADATA" {
        return (
            "SOURCE_METADATA",
            "METADATA_ONLY",
            format!("OM00: source enum declaration/values retained for {id}; consumer priority and runtime option support remain separate"),
        );
    }
    if required {
        return (
            "ELIGIBLE_EXISTING",
            "ACCEPTED_REQUIRED_NATIVE",
            format!("{packet}: exact required contract linked to checked native acceptance ledgers"),
        );
    }
    if state == "OUT_OF_SCOPE_GRIDCRAFT" {
        return (
            "OUT_OF_SCOPE_GRIDCRAFT",
            "OUT_OF_SCOPE_GRIDCRAFT",
            row["native_capability_gate"].as_str().unwrap_or("Owning native acceptance ledger excludes the missing product graph/option").into(),
        );
    }
    if state == "PLATFORM_SERVICE_NEEDED" {
        return ("SERVICE_UNAVAILABLE", "SERVICE_REQUIRED", format!("{packet}: service/continuation unavailable for {id}; no success is implied"));
    }
    if row["implementation_symbol"].is_string() {
        return (
            "ELIGIBLE_EXISTING",
            "IMPLEMENTED_NATIVE_PARTIAL",
            format!("{packet}: existing adapter metadata/runtime subset; exact Microsoft options are not certified"),
        );
    }
    // A missing descriptor is evidence of a missing projection, not proof of a missing product feature.
    (
        "OUTSIDE_DECLARED_CAPABILITY_REVIEW",
        if priority == "HIGH" { "HIGH_REFERENCE_EXTENSION" } else { "UNSELECTED_MEDIUM" },
        format!(
            "{packet}: {id} is outside the declared native obligations; retained {} catalogue state and {} priority require explicit capability review before any new projection",
            state, priority
        ),
    )
}
fn reachability(root: &Path, descriptors: &BTreeMap<(String, String), MemberDescriptor>, spec: &Value) -> Result<Value, String> {
    let metadata = gridcraft_vba::reference_metadata().map_err(|e| e.to_string())?;
    let mut all = BTreeSet::new();
    let mut objects = BTreeMap::new();
    let mut enums = BTreeMap::new();
    let mut enum_dependencies = BTreeMap::<String, BTreeSet<String>>::new();
    for library in metadata.as_array().ok_or("reference metadata missing")? {
        let name = library["library"].as_str().ok_or("library name missing")?;
        for t in library["types"].as_array().ok_or("types missing")? {
            let n = t["name"].as_str().ok_or("type missing")?;
            all.insert(fold(n));
            all.insert(fold(&format!("{name}.{n}")));
            if t["kind"] == "object" {
                objects.insert(fold(n), n.to_owned());
            } else if t["kind"] == "enum" {
                enums.insert(fold(n), t["constant_count"].as_u64().ok_or("enum constant count missing")?);
            }
        }
    }
    let implemented = descriptors.values().filter(|d| d.availability == MemberAvailability::Implemented).collect::<Vec<_>>();
    let mut edges = Vec::<(String, String, String, String)>::new();
    for d in &implemented {
        if !objects.contains_key(&fold(d.owner)) {
            return Err(format!("unregistered implemented receiver: {}", d.owner));
        }
        for ty in types_in(d.signature) {
            let base = fold(ty.rsplit('.').next().unwrap_or(&ty));
            if let Some(count) = enums.get(&base) {
                if *count == 0 {
                    return Err(format!("implemented member requires enum without source values: {} -> {ty}", member_id(d)));
                }
                enum_dependencies.entry(base).or_default().insert(member_id(d));
            }
            if !scalar(&fold(&ty)) && !all.contains(&fold(&ty)) {
                return Err(format!("unregistered signature type: {} -> {ty}", member_id(d)));
            }
        }
        if let Some(ty) = result_type(d.signature) {
            let name = ty.rsplit('.').next().unwrap_or(&ty);
            if objects.contains_key(&fold(name)) {
                edges.push((fold(d.owner), fold(name), member_id(d), "TYPED_SIGNATURE".into()));
            }
        }
    }
    for edge in spec["dynamic_returns"].as_array().ok_or("dynamic graph review missing")? {
        let id = edge["member"].as_str().ok_or("dynamic member missing")?;
        let (owner, member) = id.split_once('.').ok_or("invalid dynamic member")?;
        let d = descriptors
            .get(&(fold(owner), fold(member)))
            .filter(|d| d.availability == MemberAvailability::Implemented)
            .ok_or_else(|| format!("unimplemented dynamic edge {id}"))?;
        let result = result_type(d.signature).ok_or("dynamic edge has no result type")?;
        if !matches!(fold(&result).as_str(), "object" | "variant") {
            return Err(format!("stale dynamic edge for typed result: {id}"));
        }
        let tests = strings(&edge["test_ids"])?;
        if tests.is_empty() {
            return Err("dynamic graph edge lacks tests".into());
        }
        for t in tests {
            super::vba_audit::validate_test_reference(root, &t)?;
        }
        for to in strings(&edge["returns"])? {
            if !objects.contains_key(&fold(&to)) {
                return Err(format!("unregistered dynamic child {to}"));
            }
            edges.push((fold(owner), fold(&to), id.into(), "REVIEWED_NATIVE_DYNAMIC_RETURN".into()));
        }
    }
    let mut paths = BTreeMap::<String, Vec<String>>::new();
    for r in strings(&spec["roots"])? {
        if !objects.contains_key(&fold(&r)) {
            return Err("unregistered graph root".into());
        }
        paths.insert(fold(&r), vec![r]);
    }
    loop {
        let mut added = false;
        for (from, to, member, _) in &edges {
            if !paths.contains_key(to)
                && let Some(path) = paths.get(from)
            {
                let mut next = path.clone();
                next.push(member.clone());
                paths.insert(to.clone(), next);
                added = true;
            }
        }
        if !added {
            break;
        }
    }
    let unreachable = objects.keys().filter(|ty| !paths.contains_key(*ty)).cloned().collect::<Vec<_>>();
    if !unreachable.is_empty() {
        return Err(format!("unreachable nominal receivers: {unreachable:?}"));
    }
    let expected_collections = implemented
        .iter()
        .filter(|d| d.name.eq_ignore_ascii_case("Item"))
        .filter(|d| descriptors.get(&(fold(d.owner), "count".into())).is_some_and(|count| count.availability == MemberAvailability::Implemented))
        .map(|d| fold(d.owner))
        .collect::<BTreeSet<_>>();
    let mut reviewed_collections = BTreeSet::new();
    let mut collections = Vec::new();
    for c in spec["collections"].as_array().ok_or("collection review missing")? {
        let owner = c["owner"].as_str().ok_or("collection owner missing")?;
        if c["default_member"] != "Item" || c["enumeration_evidence"].as_str().is_none_or(str::is_empty) || !reviewed_collections.insert(fold(owner))
        {
            return Err(format!("invalid collection/default review: {owner}"));
        }
        for member in ["Item", "Count"] {
            if descriptors.get(&(fold(owner), fold(member))).is_none_or(|d| d.availability != MemberAvailability::Implemented) {
                return Err(format!("incomplete collection/default contract: {owner}.{member}"));
            }
        }
        if !paths.contains_key(&fold(owner)) {
            return Err(format!("unreachable collection {owner}"));
        }
        collections.push(c.clone());
    }
    if reviewed_collections != expected_collections {
        return Err("collection/default review does not match implemented Count/Item receivers".into());
    }
    let witnesses = objects.iter().map(|(key, name)| json!({"type":name,"path":paths.get(key)})).collect::<Vec<_>>();
    Ok(
        json!({"registered_object_types":objects.len(),"implemented_member_descriptors":implemented.len(),"enum_dependencies":enum_dependencies,"unregistered_signature_types":[],"unreachable_receivers":unreachable,"root_witnesses":witnesses,"dynamic_edges":spec["dynamic_returns"],"collections":collections,"evidence_limit":"Static signature/registration reachability plus reviewed native factory tests; this is not exhaustive runtime option coverage. Variant scalars/arrays and language-owned Run results do not invent nominal Excel types."}),
    )
}
pub(super) fn report(profile: &Value) -> Result<String, String> {
    let mut out = format!(
        "# Excel VBA bounded native profile sign-off\n\nProfile `{}` · reviewed 2026-11-03 · status **{}**.\n\n{}\n\n{} pinned source pages retained; **{} required source IDs** resolved, **{} nominal object types** reachable from Application, with checked signatures and reviewed dynamic edges. Source-page counts are not an Excel compatibility percentage.\n\n## Planning dispositions\n\n| Disposition | Source pages |\n|---|---:|\n",
        profile["profile"].as_str().unwrap_or(""),
        profile["status"].as_str().unwrap_or(""),
        profile["scope"].as_str().unwrap_or(""),
        profile["source_pages"],
        profile["required_source_ids"],
        profile["reachability"]["registered_object_types"]
    );
    for (k, v) in profile["disposition_counts"].as_object().ok_or("dispositions missing")? {
        out.push_str(&format!("| `{k}` | {v} |\n"));
    }
    if let Some(review) = profile.get("high_extension_capability_review") {
        out.push_str(&format!("\n[High extension capability review](vba-object-model-high-extension-review.md): {} source pages checked against native backing; {} direct native candidates, {} existing generic-dispatch metadata candidates, {} partial-native rows, {} host-service reviews and {} excluded source contracts. This adds no runtime acceptance.\n", review["source_pages"],review["status_counts"]["ELIGIBLE_NATIVE"],review["status_counts"]["GENERIC_DISPATCH"],review["status_counts"]["NATIVE_SUBSET"],review["status_counts"]["HOST_SERVICE_REVIEW"],review["status_counts"]["NO_NATIVE_CAPABILITY"]));
    }
    out.push_str("\nHigh reference extensions retain High priority and their literal missing/partial catalogue state; they are outside the already accepted native obligations and are not certified or downgraded. A missing adapter does not establish a missing GridCraft feature. Unselected Medium and every Low runtime projection remain outside this sign-off.\n\n## Provenance and acceptance evidence\n\n");
    for path in strings(&profile["accepted_ledgers"])? {
        let name = path.strip_prefix("docs/").unwrap_or(&path);
        out.push_str(&format!("- [{name}]({name})\n"));
    }
    out.push_str(&format!("\nAll **{}** Graph topics are explicitly a separate legacy namespace and deferred Low. **{}** same-name source groups retain distinct method/property/event identities. Four deprecated enum pages contain no numeric table; no values are invented.\n\nThe [generated contracts](vba-object-model-contracts.json) contain the exact remaining source rows and per-row priority, disposition, eligibility-evidence boundary, namespace and stable source-path key. The [machine-checked profile](vba-object-model-profile.json) contains graph witnesses and collision/exception records. The [OM20 acceptance specification](vba-object-model-om20-acceptance.json) defines scope and dynamic factories. Existing native/wasm, UI, service, persistence and protection evidence remains owned by each ledger; current metadata does not imply new Excel event signals.\n",profile["legacy_graph_pages"],profile["source_identity_collisions"].as_array().map_or(0,Vec::len)));
    Ok(out)
}
pub(super) fn emit(root: &Path, profile: &Value, write: bool) -> Result<(), String> {
    let path = root.join("docs/vba-object-model-profile.json");
    let report_path = root.join("docs/vba-object-model-profile-report.md");
    let text = report(profile)?;
    if write {
        fs::write(&path, serde_json::to_vec_pretty(profile).map_err(|e| e.to_string())?).map_err(|e| e.to_string())?;
        fs::write(report_path, text).map_err(|e| e.to_string())?;
    } else {
        let current: Value = serde_json::from_slice(&fs::read(&path).map_err(|e| e.to_string())?).map_err(|e| e.to_string())?;
        if &current != profile || fs::read_to_string(report_path).map_err(|e| e.to_string())? != text {
            return Err("OM20 generated profile/report stale; run vba-audit --write".into());
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{disposition, fold, priority, reachability, reconcile};
    use serde_json::{Value, json};
    use std::{collections::BTreeMap, fs};
    fn descriptors() -> BTreeMap<(String, String), gridcraft_vba::MemberDescriptor> {
        gridcraft_vba::member_descriptors().into_iter().map(|d| ((fold(d.owner), fold(d.name)), d)).collect()
    }
    fn spec() -> Value {
        serde_json::from_slice(&fs::read(crate::root().join("docs/vba-object-model-om20-acceptance.json")).unwrap()).unwrap()
    }
    #[test]
    fn dangling_result_types_and_unreachable_dynamic_receivers_block_signoff() {
        let mut ds = descriptors();
        let d = ds.get_mut(&("range".into(), "replace".into())).unwrap();
        d.signature = "Function Replace() As MissingExcelType";
        assert!(reachability(&crate::root(), &ds, &spec()).unwrap_err().contains("unregistered signature type"));
        let mut ds = descriptors();
        ds.get_mut(&("range".into(), "replace".into())).unwrap().signature = "Function Replace() As XlRoutingSlipDelivery";
        assert!(reachability(&crate::root(), &ds, &spec()).unwrap_err().contains("enum without source values"));
        let mut ds = descriptors();
        ds.remove(&("formatconditions".into(), "count".into()));
        assert!(reachability(&crate::root(), &ds, &spec()).unwrap_err().contains("collection/default contract"));
        let mut review = spec();
        review["collections"][0]["default_member"] = json!("MissingDefault");
        assert!(reachability(&crate::root(), &descriptors(), &review).unwrap_err().contains("collection/default review"));
        let ds = descriptors();
        let mut review = spec();
        review["dynamic_returns"].as_array_mut().unwrap().retain(|e| e["member"] != "Chart.Axes");
        assert!(reachability(&crate::root(), &ds, &review).unwrap_err().contains("unreachable nominal receivers"));
    }
    #[test]
    fn missing_required_evidence_blocks_acceptance_without_reclassifying_it() {
        let doc: Value = serde_json::from_slice(&fs::read(crate::root().join("docs/vba-object-model-contracts.json")).unwrap()).unwrap();
        let mut rows = doc["pages"].as_array().unwrap().clone();
        let row = rows.iter_mut().find(|r| r["id"] == "Workbooks.Item").unwrap();
        row["implementation_symbol"] = Value::Null;
        assert!(reconcile(&crate::root(), &mut rows, &descriptors()).unwrap_err().contains("unresolved required source obligation"));
    }
    #[test]
    fn low_deferral_keeps_runtime_evidence_independent_and_high_gaps_stay_high() {
        let low = json!({"packet":"OM17","id":"Application.VBE","owner":"Application","member":"VBE","result_state":"MISSING"});
        assert_eq!(priority(&low), "LOW");
        assert_eq!(disposition(&low, "LOW", false).1, "DEFERRED_LOW");
        assert_eq!(low["result_state"], "MISSING");
        let high = json!({"packet":"OM04","id":"Worksheet.Next","owner":"Worksheet","member":"Next","result_state":"MISSING"});
        assert_eq!(priority(&high), "HIGH");
        assert_eq!(disposition(&high, "HIGH", false).1, "HIGH_REFERENCE_EXTENSION");
    }
}
