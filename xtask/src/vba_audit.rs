//! Offline integrity check for the pinned Excel VBA object-model inventory.
//!
//! The documentation archive is cached under `target/` and is never fetched by
//! this command. The committed JSON remains the reviewable inventory; this check
//! proves that its source page set still matches the pinned archive.

use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, BTreeSet};
use std::fs::{self, File};
use std::io::Read;
use std::path::Path;
use zip::ZipArchive;

use gridcraft_vba::{MemberAvailability, MemberDescriptor};

const ARCHIVE_SHA256: &str = "285218108a034fc33acc8296d714395fd182397ada27c848b93c3d5423efd75e";
pub(super) const COMMIT: &str = "b202c96d9d53896593b73d6fae57f5230601a5e3";
const PAGE_COUNT: usize = 5_996;
const DOCUMENTATION_EXCEPTIONS: &[(&str, &str, &str)] = &[
    ("workbook", "calculate", "Excel exposes workbook-scoped Calculate although the pinned VBA-Docs set has no Workbook.Calculate page"),
    (
        "range",
        "goalseek",
        "Native Goal Seek projection; pinned VBA-Docs omits this member, independently documented at https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range.goalseek",
    ),
    (
        "range",
        "evaluate",
        "GridCraft range-relative Evaluate extension explicitly requested by OM07; pinned Excel VBA reference only documents Application/Worksheet receivers, so this is not an Excel parity claim",
    ),
];

pub fn run(root: &Path, args: &[&str]) -> Result<(), String> {
    let write = args.contains(&"--write");
    if args.iter().any(|arg| *arg != "--write") {
        return Err("usage: cargo xtask vba-audit [--write]".into());
    }
    let cache = root.join("target/vba-docs-audit");
    let archive_path = cache.join("reference.zip");
    let inventory_path = root.join("docs/vba/reference/vba-object-model-inventory.json");
    let archive_bytes = std::fs::read(&archive_path)
        .map_err(|e| format!("{}: {e}; cache the pinned MicrosoftDocs/VBA-Docs archive first", archive_path.display()))?;
    let digest = Sha256::digest(&archive_bytes).iter().map(|byte| format!("{byte:02x}")).collect::<String>();
    if digest != ARCHIVE_SHA256 {
        return Err(format!("pinned documentation archive SHA256 mismatch: {digest}"));
    }
    let inventory: Value = serde_json::from_slice(&std::fs::read(&inventory_path).map_err(|e| format!("{}: {e}", inventory_path.display()))?)
        .map_err(|e| format!("invalid inventory JSON: {e}"))?;
    if inventory.get("documentation_commit").and_then(Value::as_str) != Some(COMMIT) {
        return Err("inventory documentation_commit does not match the pinned source".into());
    }
    if inventory.get("documentation_archive_sha256").and_then(Value::as_str) != Some(ARCHIVE_SHA256) {
        return Err("inventory archive digest does not match the pinned source".into());
    }
    let rows = inventory.get("pages").and_then(Value::as_array).ok_or("inventory has no pages array")?;
    let mut inventory_paths = BTreeSet::new();
    for row in rows {
        let source_path = row.get("source_path").and_then(Value::as_str).ok_or("inventory row has no source_path")?;
        if !inventory_paths.insert(source_path.to_owned()) {
            return Err(format!("duplicate inventory source_path: {source_path}"));
        }
    }
    if rows.len() != PAGE_COUNT || inventory_paths.len() != PAGE_COUNT {
        return Err(format!("inventory contains {} pages; expected {PAGE_COUNT}", rows.len()));
    }

    let file = File::open(&archive_path).map_err(|e| format!("{}: {e}", archive_path.display()))?;
    let mut archive = ZipArchive::new(file).map_err(|e| format!("invalid pinned archive: {e}"))?;
    let mut archive_paths = BTreeSet::new();
    for index in 0..archive.len() {
        let entry = archive.by_index(index).map_err(|e| format!("archive entry {index}: {e}"))?;
        let name = entry.name();
        if let Some((_, path)) = name.split_once("/api/")
            && path.starts_with("Excel.")
            && path.ends_with(".md")
        {
            archive_paths.insert(format!("api/{path}"));
        }
    }
    if archive_paths.len() != PAGE_COUNT {
        return Err(format!("pinned archive contains {} Excel API pages; expected {PAGE_COUNT}", archive_paths.len()));
    }
    let missing: Vec<_> = archive_paths.difference(&inventory_paths).take(10).cloned().collect();
    let extra: Vec<_> = inventory_paths.difference(&archive_paths).take(10).cloned().collect();
    if !missing.is_empty() || !extra.is_empty() {
        return Err(format!("inventory/source page set differs; missing: {missing:?}; extra: {extra:?}"));
    }
    let mut descriptors = BTreeMap::<(String, String), MemberDescriptor>::new();
    for descriptor in gridcraft_vba::member_descriptors() {
        let key = (fold_identifier(descriptor.owner), fold_identifier(descriptor.name));
        if descriptors.insert(key, descriptor).is_some() {
            return Err(format!("duplicate VBA member descriptor: {}.{}", descriptor.owner, descriptor.name));
        }
    }
    let acceptance = validate_high_acceptance(root, rows)?;
    let om12_acceptance = validate_native_acceptance(root, rows, &descriptors, "OM12")?;
    let om13_acceptance = validate_native_acceptance(root, rows, &descriptors, "OM13")?;
    let om08_acceptance = validate_native_acceptance(root, rows, &descriptors, "OM08")?;
    let om09_acceptance = validate_native_acceptance(root, rows, &descriptors, "OM09")?;
    let om10_acceptance = validate_native_acceptance(root, rows, &descriptors, "OM10")?;
    let om19_acceptance = validate_native_acceptance(root, rows, &descriptors, "OM19")?;
    let om15_acceptance = validate_native_acceptance(root, rows, &descriptors, "OM15")?;
    for test in OM12_TESTS.iter().chain(OM12_DRAWING_TESTS) {
        validate_test_reference(root, test)?;
    }
    let mut contracts = Vec::with_capacity(rows.len());
    let mut enums = Vec::new();
    let mut enum_types = BTreeSet::new();
    for row in rows {
        let source_path = row.get("source_path").and_then(Value::as_str).ok_or("inventory row has no source_path")?;
        let archive_name = format!("VBA-Docs-{COMMIT}/{source_path}");
        let mut entry = archive.by_name(&archive_name).map_err(|e| format!("missing {source_path}: {e}"))?;
        let mut markdown = String::new();
        entry.read_to_string(&mut markdown).map_err(|e| format!("invalid UTF-8 in {source_path}: {e}"))?;
        let mut contract = contract_row(row, &markdown, &descriptors)?;
        if let Some(tests) = row.get("id").and_then(Value::as_str).and_then(|id| acceptance.get(&fold_identifier(id))) {
            contract["test_ids"] = json!(tests);
            contract["high_acceptance"] =
                json!("IMPLEMENTED_NATIVE_PROFILE; see vba-object-model-high-acceptance.json for exact signatures and exclusions");
        }
        if let Some(tests) = row.get("id").and_then(Value::as_str).and_then(|id| om12_acceptance.get(&fold_identifier(id))) {
            contract["test_ids"] = json!(tests);
            contract["om12_acceptance"] = json!(
                "IMPLEMENTED_NATIVE_PROFILE; see vba-object-model-om12-acceptance.json for signatures, bounds and native capability exclusions"
            );
        }
        if let Some(tests) = row.get("id").and_then(Value::as_str).and_then(|id| om13_acceptance.get(&fold_identifier(id))) {
            merge_acceptance_tests(&mut contract, tests);
            contract["om13_acceptance"] = json!(
                "IMPLEMENTED_NATIVE_PROFILE; see vba-object-model-om13-acceptance.json for native capability limits, cache ownership and slicer persistence gaps"
            );
        }
        if let Some(tests) = row.get("id").and_then(Value::as_str).and_then(|id| om08_acceptance.get(&fold_identifier(id))) {
            merge_acceptance_tests(&mut contract, tests);
            contract["om08_acceptance"] = json!(
                "IMPLEMENTED_NATIVE_PROFILE; see vba-object-model-om08-acceptance.json for window ownership, stash lifetimes and unavailable view services"
            );
        }
        if let Some(tests) = row.get("id").and_then(Value::as_str).and_then(|id| om09_acceptance.get(&fold_identifier(id))) {
            merge_acceptance_tests(&mut contract, tests);
            contract["om09_acceptance"] = json!(
                "IMPLEMENTED_NATIVE_PROFILE; see vba-object-model-om09-acceptance.json for authentication, native selection and permission/encryption exclusions"
            );
        }
        if let Some(tests) = row.get("id").and_then(Value::as_str).and_then(|id| om10_acceptance.get(&fold_identifier(id))) {
            merge_acceptance_tests(&mut contract, tests);
            contract["om10_acceptance"] = json!(
                "IMPLEMENTED_NATIVE_PROFILE; see vba-object-model-om10-acceptance.json for pagination epochs, PDF option limits and native/browser service gaps"
            );
        }
        if let Some(tests) = row.get("id").and_then(Value::as_str).and_then(|id| om19_acceptance.get(&fold_identifier(id))) {
            merge_acceptance_tests(&mut contract, tests);
            contract["om19_acceptance"] = json!(
                "IMPLEMENTED_NATIVE_PROFILE; see vba-object-model-om19-acceptance.json for local annotation, cell-handle, metadata and external-navigation limits"
            );
        }
        if let Some(tests) = row["id"].as_str().and_then(|id| om15_acceptance.get(&fold_identifier(id))) {
            merge_acceptance_tests(&mut contract, tests);
            contract["om15_acceptance"] = json!(
                "IMPLEMENTED_NATIVE_PROFILE; see vba-object-model-om15-acceptance.json for modal ownership, native dialog/picker options, headless refusals and cell-control limits"
            );
        }
        reconcile_om15_native_scope(&mut contract);
        reconcile_om19_native_scope(&mut contract);
        reconcile_om10_native_scope(&mut contract);
        reconcile_om09_native_scope(&mut contract);
        reconcile_om08_native_scope(&mut contract);
        reconcile_om13_native_scope(&mut contract);
        if contract.get("enum_type").is_some_and(|ty| !ty.is_null()) {
            if let Some(ty) = contract.get("enum_type").and_then(Value::as_str) {
                enum_types.insert(ty.to_owned());
            }
            if let Some(values) = contract.get("enum_values").and_then(Value::as_array) {
                for value in values {
                    if let (Some(ty), Some(name), Some(number)) = (
                        contract.get("enum_type").and_then(Value::as_str),
                        value.get("name").and_then(Value::as_str),
                        value.get("number").and_then(Value::as_i64),
                    ) && let Ok(number) = i32::try_from(number)
                    {
                        enums.push((ty.to_owned(), name.to_owned(), number));
                    }
                }
            }
        }
        contracts.push(contract);
    }
    enums.sort();
    let contract_ids = contracts
        .iter()
        .filter_map(|row| Some((fold_identifier(row.get("owner")?.as_str()?), fold_identifier(row.get("member")?.as_str()?))))
        .collect::<BTreeSet<_>>();
    // Shared Office child types are documented in the same pinned archive, outside
    // the deliberately unchanged 5,996-page Excel inventory.
    let mut office_contracts = Vec::new();
    let mut office_keys = BTreeSet::new();
    for (key, d) in &descriptors {
        if contract_ids.contains(key) {
            continue;
        }
        let source = format!("api/Office.{}.{}.md", d.owner, d.name);
        if let Ok(mut entry) = archive.by_name(&format!("VBA-Docs-{COMMIT}/{source}")) {
            let mut md = String::new();
            entry.read_to_string(&mut md).map_err(|e| e.to_string())?;
            if md.is_empty() {
                return Err(format!("empty Office source: {source}"));
            }
            office_keys.insert(key.clone());
            office_contracts.push(json!({"id":format!("Office.{}.{}",d.owner,d.name),"source_path":source,"signature":d.signature,"implementation_symbol":d.implementation_symbol,"test_ids":if matches!(d.owner,"DocumentProperties"|"DocumentProperty") { &OM19_METADATA_TESTS[..] } else { OM12_DRAWING_TESTS }}));
        }
    }
    let mut cell_source = String::new();
    archive
        .by_name(&format!("VBA-Docs-{COMMIT}/api/Excel.Range.CellControl.md"))
        .map_err(|e| e.to_string())?
        .read_to_string(&mut cell_source)
        .map_err(|e| e.to_string())?;
    let mut inline_contracts = Vec::new();
    let mut inline_keys = BTreeSet::new();
    for (member, source_text) in [("Type", "Property: Type"), ("SetCheckbox", ".SetCheckbox")] {
        if !cell_source.contains(source_text) {
            return Err(format!("CellControl.{member} lacks pinned inline source"));
        }
        let key = (fold_identifier("CellControl"), fold_identifier(member));
        let d = descriptors.get(&key).ok_or("missing CellControl descriptor")?;
        inline_keys.insert(key);
        inline_contracts.push(json!({"id":format!("CellControl.{member}"),"source_path":"api/Excel.Range.CellControl.md","signature":d.signature,"implementation_symbol":d.implementation_symbol,"test_ids":["crates/vba/tests/interactions.rs::checkbox_controls_preserve_values_and_formulas_rebase_and_roundtrip","crates/vba/tests/interactions.rs::checkbox_value_changes_deliver_exact_native_event_targets_once"]}));
    }
    for (name, n) in gridcraft_vba::CELL_CONTROL_CONSTANTS {
        if !cell_source.contains(&format!("{name} = {n}")) {
            return Err(format!("missing inline constant {name}"));
        }
        inline_contracts.push(json!({"id":name,"value":n,"source_path":"api/Excel.Range.CellControl.md"}));
    }
    let office_types = [
        "MsoAutoShapeType",
        "MsoShapeType",
        "MsoZOrderCmd",
        "MsoAlignCmd",
        "MsoTextOrientation",
        "MsoTriState",
        "MsoConnectorType",
        "MsoArrowheadStyle",
        "MsoThemeColorIndex",
        "MsoVerticalAnchor",
        "MsoLineDashStyle",
        "MsoFillType",
    ];
    let mut office_values = Vec::new();
    for ty in office_types {
        let mut entry = archive.by_name(&format!("VBA-Docs-{COMMIT}/api/Office.{ty}.md")).map_err(|e| e.to_string())?;
        let mut md = String::new();
        entry.read_to_string(&mut md).map_err(|e| e.to_string())?;
        let vals = enum_values(&md);
        if vals.is_empty() {
            return Err(format!("no values for Office.{ty}"));
        }
        for v in vals {
            let n = v["number"].as_i64().and_then(|n| i32::try_from(n).ok()).ok_or("Office enum overflow")?;
            office_values.push((ty.to_owned(), v["name"].as_str().ok_or("Office enum name")?.to_owned(), n));
        }
    }
    office_values.sort();
    let office_source = render_office_enums(&office_types, &office_values);
    let unassigned_descriptors = descriptors
        .keys()
        .filter(|key| {
            !contract_ids.contains(*key)
                && !office_keys.contains(*key)
                && !inline_keys.contains(*key)
                && !DOCUMENTATION_EXCEPTIONS.iter().any(|(owner, member, _)| key.0 == *owner && key.1 == *member)
        })
        .map(|(owner, member)| format!("{owner}.{member}"))
        .collect::<Vec<_>>();
    if !unassigned_descriptors.is_empty() {
        return Err(format!("VBA member descriptors have no canonical documentation row: {unassigned_descriptors:?}"));
    }
    let aliases = duplicate_enum_constants(&enums);
    let constants = deduplicate_enum_constants(enums);
    let unsupported_claims = contracts
        .iter()
        .filter_map(|row| {
            let baseline = row.get("baseline_status").and_then(Value::as_str)?;
            let symbol = row.get("implementation_symbol").and_then(Value::as_str);
            (baseline == "implemented-subset" && symbol.is_none()).then(|| row.get("id").and_then(Value::as_str).unwrap_or("UNKNOWN").to_owned())
        })
        .collect::<Vec<_>>();
    if !unsupported_claims.is_empty() {
        return Err(format!("implemented-subset inventory claims lack an adapter dispatch/event symbol: {unsupported_claims:?}"));
    }
    let mut profile = super::vba_profile::reconcile(root, &mut contracts, &descriptors)?;
    super::vba_extensions::reconcile(root, &mut contracts, &mut profile, write)?;
    super::vba_profile::emit(root, &profile, write)?;
    super::vba_docs::emit(root, &contracts, &descriptors, write)?;
    let contract_file = root.join("docs/vba/reference/vba-object-model-contracts.json");
    let enum_file = root.join("crates/vba/src/vba_enum_data.rs");
    let report_file = root.join("docs/vba/reference/vba-object-model-audit-report.md");
    let contract_value = json!({
        "schema": 3,
        "documentation_commit": COMMIT,
        "documentation_archive_sha256": ARCHIVE_SHA256,
        "page_count": PAGE_COUNT,
        "runtime_extensions_without_source_page": DOCUMENTATION_EXCEPTIONS.iter().map(|(owner, member, reason)| json!({
            "id": format!("{owner}.{member}"),
            "reason": reason,
        })).collect::<Vec<_>>(),
        "office_runtime_extensions":office_contracts,
        "inline_excel_runtime_extensions":inline_contracts,
        "pages": contracts,
    });
    let enum_source = render_enum_source(&enum_types, &constants, &aliases);
    let report = render_report(&contract_value)?;
    if write {
        fs::write(&contract_file, serde_json::to_vec_pretty(&contract_value).map_err(|e| e.to_string())?)
            .map_err(|e| format!("{}: {e}", contract_file.display()))?;
        fs::write(root.join("crates/vba/src/vba_office_enums.rs"), &office_source).map_err(|e| e.to_string())?;
        fs::write(&enum_file, enum_source).map_err(|e| format!("{}: {e}", enum_file.display()))?;
        fs::write(&report_file, report).map_err(|e| format!("{}: {e}", report_file.display()))?;
    } else {
        compare_generated(&contract_file, &contract_value)?;
        compare_generated_text(&root.join("crates/vba/src/vba_office_enums.rs"), &office_source)?;
        compare_generated_text(&enum_file, &enum_source)?;
        compare_generated_text(&report_file, &report)?;
    }
    println!(
        "VBA object-model audit: {PAGE_COUNT} pinned source pages; {} semantic rows; {} enum types; {} typed constants; unresolved rows are listed in {}",
        contracts.len(),
        count_enum_types(&contracts),
        constants.len(),
        report_file.display()
    );
    Ok(())
}

const OM19_METADATA_TESTS: [&str; 1] = ["crates/vba/tests/annotations.rs::hyperlinks_and_builtin_metadata_round_trip_and_undo"];
const OM12_DRAWING_TESTS: &[&str] = &[
    "crates/vba/tests/drawings.rs::typed_series_arrays_trendlines_and_layout_share_native_state",
    "crates/vba/tests/drawings.rs::shapes_text_range_format_zorder_and_inactive_sheet_undo",
    "crates/vba/tests/drawings.rs::series_add_delete_and_rebind_never_retarget_parked_children",
    "crates/vba/tests/drawings.rs::drawing_ranges_follow_sheet_moves_and_refuse_stale_or_unsupported_bulk_writes",
    "crates/vba/tests/drawings.rs::drawing_protection_and_reused_ids_refuse_before_effects",
    "crates/vba/tests/drawings.rs::picture_reads_and_failed_services_preserve_native_drawing_state",
    "crates/vba/tests/drawings.rs::add_picture_embeds_native_file_with_original_dimensions_on_inactive_owner",
    "crates/vba/tests/drawings.rs::names_native_anchor_transforms_and_source_deletion_stay_live",
    "crates/vba/tests/drawings.rs::embedded_chart_location_retargets_parked_descendants_without_switching_view",
    "crates/vba/tests/drawings.rs::external_chart_moves_preserve_existing_graph_lifetimes",
];
fn merge_acceptance_tests(contract: &mut Value, tests: &[String]) {
    let mut all = contract["test_ids"].as_array().cloned().unwrap_or_default();
    for test in tests {
        let test = json!(test);
        if !all.contains(&test) {
            all.push(test);
        }
    }
    contract["test_ids"] = json!(all);
}

fn validate_native_acceptance(
    root: &Path,
    rows: &[Value],
    descriptors: &BTreeMap<(String, String), MemberDescriptor>,
    packet: &str,
) -> Result<BTreeMap<String, Vec<String>>, String> {
    let path = root.join(format!("docs/vba/reference/vba-object-model-{}-acceptance.json", packet.to_ascii_lowercase()));
    let value: Value = serde_json::from_slice(&fs::read(&path).map_err(|e| e.to_string())?).map_err(|e| e.to_string())?;
    if value["packet"] != packet || value["documentation_commit"] != COMMIT {
        return Err("native ledger provenance mismatch".into());
    }
    let source_ids = rows.iter().filter_map(|r| r["id"].as_str().map(fold_identifier)).collect::<BTreeSet<_>>();
    let mut result = BTreeMap::new();
    for group in value["required_obligations"].as_array().ok_or("native ledger obligations missing")? {
        let tests = group["test_ids"]
            .as_array()
            .ok_or("native tests missing")?
            .iter()
            .map(|t| t.as_str().map(str::to_owned).ok_or("native test ID must be text"))
            .collect::<Result<Vec<_>, _>>()?;
        if tests.is_empty() {
            return Err("native obligation has no tests".into());
        }
        for test in &tests {
            validate_test_reference(root, test)?;
        }
        for id in group["source_ids"].as_array().ok_or("native source IDs missing")? {
            let id = id.as_str().ok_or("native source ID must be text")?;
            let key = fold_identifier(id);
            if !source_ids.contains(&key) {
                return Err(format!("unknown native source ID: {id}"));
            }
            let (owner, member) = id.split_once('.').ok_or("native obligation must name a member")?;
            if descriptors.get(&(fold_identifier(owner), fold_identifier(member))).is_none_or(|d| d.availability != MemberAvailability::Implemented) {
                return Err(format!("native required member is not implemented: {id}"));
            }
            result.entry(key).or_insert_with(Vec::new).extend(tests.clone());
        }
    }
    if result.is_empty() {
        return Err("native ledger has no required members".into());
    }
    println!("{packet} native acceptance: {} required source contracts; native product gaps remain explicit", result.len());
    Ok(result)
}
fn reconcile_om15_native_scope(row: &mut Value) {
    if row["packet"] != "OM15" || row["implementation_symbol"].is_string() {
        return;
    }
    let owner = row["owner"].as_str().unwrap_or("").to_ascii_lowercase();
    let member = row["member"].as_str().unwrap_or("").to_ascii_lowercase();
    if member.is_empty() && matches!(owner.as_str(), "dialogs" | "dialog") {
        row["result_state"] = json!("PARTIAL");
        row["implementation_symbol"] = json!("crates/vba/src/types.rs::references");
        row["dispatch_evidence"] = json!("native selected dialog graph; Show uses owned UI continuation");
        return;
    }
    let service = owner == "range" && matches!(member.as_str(), "checkspelling" | "functionwizard");
    let reason = if service {
        "Existing worksheet spelling/function tools have no selected range-owned VBA completion contract; these adapters remain outside the selected OM15 profile, not new product features"
    } else if owner == "application" {
        "No corresponding native application-wide cursor, redraw/alert suppression, arbitrary key override, custom status text, insert/paste options, selection floaties or configurable Office.FileDialog graph. Native filename pickers and status visibility are projected separately"
    } else if owner == "autocorrect" {
        "No independent native AutoCorrect replacement list/capitalization/table-expansion settings graph; structured-table behavior is not application AutoCorrect state"
    } else if owner == "errorcheckingoptions" {
        "Native error indicators/checking do not expose independent per-rule enable switches, indicator palette index or background-check configuration"
    } else if matches!(owner.as_str(), "action" | "actions" | "research" | "speech") || member == "speak" {
        "No native smart-tag action/provider, research or speech runtime; provider-backed and Low interactions remain deferred"
    } else if owner == "spellingoptions" {
        "Native spell checking has no corresponding configurable Excel locale/dictionary/script-specific options object"
    } else if owner == "quickanalysis" {
        "No native Quick Analysis overlay/service object"
    } else if owner == "range" && member == "dialogbox" {
        "No native Excel 5 dialog-sheet runtime; generic tiny-vba user forms are a separate surface, not a dialog-sheet/control compatibility claim"
    } else if member == "creator" {
        "No Office application creator signature is invented for GridCraft"
    } else {
        return;
    };
    row["result_state"] = json!(if service { "PLATFORM_SERVICE_NEEDED" } else { "OUT_OF_SCOPE_GRIDCRAFT" });
    row["native_capability_gate"] = json!(reason);
}
fn reconcile_om19_native_scope(row: &mut Value) {
    if row["packet"] != "OM19" || row["implementation_symbol"].is_string() {
        return;
    }
    let owner = row["owner"].as_str().unwrap_or("").to_ascii_lowercase();
    let member = row["member"].as_str().unwrap_or("").to_ascii_lowercase();
    if member.is_empty()
        && matches!(owner.as_str(), "comments" | "comment" | "commentsthreaded" | "commentthreaded" | "author" | "hyperlinks" | "hyperlink")
    {
        row["result_state"] = json!("PARTIAL");
        row["implementation_symbol"] = json!("crates/vba/src/types.rs::references");
        row["dispatch_evidence"] = json!("native nominal annotation type; reachable through OM19 descriptors; local cell-map profile only");
        return;
    }
    let reason = if owner == "author" || owner == "commentthreaded" && matches!(member.as_str(), "author" | "date") {
        "Native comments retain local author labels/imported dates but have no remote Author provider/user identity or timestamp creation service; remote identities remain deferred Low"
    } else if matches!(owner.as_str(), "customproperties" | "customproperty") {
        "No native worksheet custom-property store, command or XLSX writer; native workbook DocProps has only six built-in text fields"
    } else if owner == "multithreadedcalculation" {
        "No Excel-style mutable per-application calculation thread-count/mode configuration graph; native scheduling is not this API"
    } else if owner == "soundnote" || member == "soundnote" {
        "No native sound-note recording/import/playback service or model; legacy audio members remain deferred Low"
    } else if member == "creator" {
        "No Office creator identity; do not invent an Excel application signature for GridCraft"
    } else if member == "shape" {
        "Native note boxes and cell hyperlinks do not own native drawing Shape objects; formatting/shape anchors require a missing product graph"
    } else if matches!(member.as_str(), "addtofavorites" | "createnewdocument") {
        "No native browser favorites or linked-document creation service; deferred Low"
    } else if owner == "hyperlink" && member == "name" {
        "Native cell hyperlinks have no independent persisted name; TextToDisplay is projected separately"
    } else {
        return;
    };
    row["result_state"] = json!("OUT_OF_SCOPE_GRIDCRAFT");
    row["native_capability_gate"] = json!(reason);
}
fn reconcile_om10_native_scope(row: &mut Value) {
    if row["packet"] != "OM10" || row["implementation_symbol"].is_string() {
        return;
    }
    let owner = row["owner"].as_str().unwrap_or("").to_ascii_lowercase();
    let member = row["member"].as_str().unwrap_or("").to_ascii_lowercase();
    if member.is_empty()
        && matches!(owner.as_str(), "pagesetup" | "pages" | "page" | "headerfooter" | "hpagebreaks" | "vpagebreaks" | "hpagebreak" | "vpagebreak")
    {
        row["result_state"] = json!("PARTIAL");
        row["implementation_symbol"] = json!("crates/vba/src/types.rs::references");
        row["dispatch_evidence"] = json!("native nominal print type; reachable through OM10 member descriptors");
        return;
    }
    let service = owner == "pagesetup" && member == "printquality";
    let reason = if service {
        "No connected printer-resolution service"
    } else if owner == "graphic" || owner == "headerfooter" || owner == "pagesetup" && member.ends_with("picture") {
        "No native header/footer picture object graph; worksheet images are a different owner"
    } else if owner.starts_with("chart") {
        "No independent chart print/page-layout or per-object print-inclusion state; native worksheet PDF still renders embedded charts"
    } else if matches!(owner.as_str(), "hpagebreak" | "vpagebreak") {
        "No independent partial-break region/drag-off or viewport interaction model"
    } else if owner == "pagesetup" {
        "No corresponding native configurable monochrome/draft/first-number/comment/note/error print setting; native renderer defaults are not independent settings"
    } else if owner == "worksheet" {
        "Native renderer has no comment-only printed-page model"
    } else if owner == "workbook" {
        "No per-window/personal-view print settings model"
    } else if owner == "controlformat" {
        "No native control-specific print inclusion flag"
    } else {
        return;
    };
    row["result_state"] = json!(if service { "PLATFORM_SERVICE_NEEDED" } else { "OUT_OF_SCOPE_GRIDCRAFT" });
    row["native_capability_gate"] = json!(reason);
}
fn reconcile_om09_native_scope(row: &mut Value) {
    if row["packet"] != "OM09" || row["implementation_symbol"].is_string() {
        return;
    }
    let owner = row["owner"].as_str().unwrap_or("").to_ascii_lowercase();
    let member = row["member"].as_str().unwrap_or("").to_ascii_lowercase();
    if member.is_empty() && matches!(owner.as_str(), "protection" | "alloweditranges" | "alloweditrange") {
        row["result_state"] = json!("PARTIAL");
        row["implementation_symbol"] = json!("crates/vba/src/types.rs::references");
        row["dispatch_evidence"] = json!("native nominal type; reachable through OM09 member descriptors");
        return;
    }
    let reason = if owner.starts_with("protectedview") || owner == "application" {
        "No native protected-view window/session graph; active-content consent is not a document protected-view or encryption session"
    } else if owner.starts_with("useraccess") || owner == "alloweditrange" && member == "users" {
        "Native securityDescriptor is opaque round-trip metadata; no principal resolver, user ACL enforcement or Windows permission service"
    } else if owner == "workbook" {
        "Native supports password-authenticated encrypted-file reading through Workbooks.Open only; no retained encryption/password-to-write/sharing/IRM/signature metadata or corresponding writer service. Workbook window protection is also absent"
    } else if owner.starts_with("chart") {
        "No independent chart/content/data/drawing/format/selection protection model; existing worksheet protection guards remain in force"
    } else if owner == "style" && member == "includeprotection" {
        "Native named styles apply whole styles; no independent protection-inclusion switch"
    } else {
        return;
    };
    row["result_state"] = json!("OUT_OF_SCOPE_GRIDCRAFT");
    row["native_capability_gate"] = json!(reason);
}
fn reconcile_om08_native_scope(row: &mut Value) {
    if row["packet"] != "OM08" || row["implementation_symbol"].is_string() {
        return;
    }
    let owner = row["owner"].as_str().unwrap_or("").to_ascii_lowercase();
    let member = row["member"].as_str().unwrap_or("").to_ascii_lowercase();
    if member.is_empty() && matches!(owner.as_str(), "windows" | "window" | "customviews" | "customview" | "watches" | "watch") {
        row["result_state"] = json!("PARTIAL");
        row["implementation_symbol"] = json!("crates/vba/src/types.rs::references");
        row["dispatch_evidence"] = json!("native nominal type; reachable through OM08 member descriptors");
        return;
    }
    let service = (owner == "window"
        && matches!(
            member.as_str(),
            "pointstoscreenpixelsx" | "pointstoscreenpixelsy" | "rangefrompoint" | "scrollintoview" | "usableheight" | "usablewidth" | "hwnd"
        ))
        || (owner == "range" && member == "show")
        || (owner == "application" && member == "windowstate");
    let reason = if service {
        "No native/browser window metrics or viewport service is connected to the VBA host"
    } else if matches!(owner.as_str(), "pane" | "panes") {
        "No independent native split-pane object graph; native Split currently aliases sheet-owned freeze state"
    } else if matches!(owner.as_str(), "sheetviews" | "worksheetview" | "dialogsheetview" | "moduleview" | "treeviewcontrol") {
        "No native per-window sheet-view, dialog/module-sheet view or tree-control object graph"
    } else if owner == "window" {
        "No corresponding independent native window-owned setting, chart focus, macro callback or Office creator identity; worksheet/global settings cannot be substituted"
    } else if owner == "windows" {
        "No native side-by-side synchronization/position model or Office creator identity"
    } else if owner == "application" {
        "No corresponding native application-wide window/pen/scrollbar state"
    } else if owner == "worksheet" && member == "scrollarea" {
        "No native worksheet scroll-area restriction model or command"
    } else if matches!(owner.as_str(), "customview" | "customviews" | "watch" | "watches") && member == "creator" {
        "No native Office creator identity"
    } else {
        return;
    };
    row["result_state"] = json!(if service { "PLATFORM_SERVICE_NEEDED" } else { "OUT_OF_SCOPE_GRIDCRAFT" });
    row["native_capability_gate"] = json!(reason);
}
fn reconcile_om13_native_scope(row: &mut Value) {
    if row["packet"] != "OM13" || row["implementation_symbol"].is_string() {
        return;
    }
    let owner = row["owner"].as_str().unwrap_or("").to_ascii_lowercase();
    let member = row["member"].as_str().unwrap_or("").to_ascii_lowercase();
    let native_types = [
        "pivottables",
        "pivottable",
        "pivotcaches",
        "pivotcache",
        "pivotfields",
        "pivotfield",
        "pivotitems",
        "pivotitem",
        "calculatedfields",
        "pivotcell",
        "pivotlayout",
        "slicercaches",
        "slicercache",
        "slicers",
        "slicer",
        "sliceritems",
        "sliceritem",
    ];
    if member.is_empty() && native_types.contains(&owner.as_str()) {
        row["result_state"] = json!("PARTIAL");
        row["implementation_symbol"] = json!("crates/vba/src/types.rs::references");
        row["dispatch_evidence"] = json!("native nominal type; reachable through OM13 member descriptors");
        return;
    }
    let reason = if owner.starts_with("model3d") {
        Some("No native 3-D drawing object graph")
    } else if owner.starts_with("cube")
        || owner.starts_with("model")
        || owner.starts_with("timeline")
        || owner.starts_with("calculatedmember")
        || matches!(owner.as_str(), "valuechange" | "pivottablechangelist" | "pivotvaluecell")
    {
        Some("No native OLAP, data-model, writeback or timeline service/object graph")
    } else if matches!(owner.as_str(), "slicerpivottables" | "slicercachelevel" | "slicercachelevels") {
        Some("Native slicers filter one table column; no pivot connections or OLAP levels")
    } else if matches!(owner.as_str(), "calculateditems" | "pivotfilters" | "pivotfilter") {
        Some("Native pivots support calculated fields and manual item selections; no calculated-item or advanced filter object model")
    } else if owner == "pivotcache" {
        Some("Native caches are local report-owned source views; no retained cache/provider lifecycle state or standalone/shared cache service")
    } else if owner == "pivottable"
        && matches!(
            member.as_str(),
            "allocatechanges"
                | "allocation"
                | "allocationmethod"
                | "allocationvalue"
                | "allocationweightexpression"
                | "changeconnection"
                | "changelist"
                | "commitchanges"
                | "createcubefile"
                | "cubefields"
                | "discardchanges"
                | "enabledatavalueediting"
                | "enablewriteback"
                | "mdx"
                | "visualtotals"
                | "visualtotalsforsets"
                | "viewcalculatedmembers"
                | "calculatedmembers"
        )
    {
        Some("No native OLAP/provider, data-model or writeback service")
    } else if owner == "slicercache"
        && matches!(member.as_str(), "timelinestate" | "slicercachelevels" | "pivottables" | "workbookconnection" | "cleardatefilter")
    {
        Some("No native timeline, shared pivot-slicer connection or provider service")
    } else {
        None
    };
    if let Some(reason) = reason {
        row["result_state"] = json!("OUT_OF_SCOPE_GRIDCRAFT");
        row["native_capability_gate"] = json!(reason);
    }
}
fn render_office_enums(types: &[&str], values: &[(String, String, i32)]) -> String {
    let mut out = String::from("// Generated from the pinned VBA-Docs Office enum pages; regenerate with cargo xtask vba-audit --write.\n");
    out.push_str("pub(crate) const OFFICE_ENUM_TYPES: &[&str] = &[\n");
    for ty in types {
        out.push_str(&format!("    {ty:?},\n"));
    }
    out.push_str("];\npub(crate) const OFFICE_ENUM_CONSTANTS: &[(&str, &str, i32)] = &[\n");
    for (ty, name, n) in values {
        out.push_str(&format!("    ({ty:?}, {name:?}, {n}),\n"));
    }
    out.push_str("];\n");
    out
}
const OM12_TESTS: &[&str] = &[
    "crates/vba/tests/charts.rs::typed_chart_graph_edits_inactive_sheet_and_retains_single_undo_group",
    "crates/vba/tests/charts.rs::chart_source_and_collection_defaults_use_live_native_state",
    "crates/vba/tests/charts.rs::native_chart_id_reuse_does_not_retarget_parked_handles",
    "crates/vba/tests/charts.rs::invalid_chart_options_and_protection_do_not_partially_write",
    "crates/vba/tests/charts.rs::chart_edits_persist_and_renderer_reads_the_same_native_model",
    "crates/vba/tests/charts.rs::deleting_title_and_chart_updates_live_counts_and_invalidates_children",
];
pub(super) fn validate_test_reference(root: &Path, test: &str) -> Result<(), String> {
    let (file, function) = test.split_once("::").ok_or("invalid acceptance test reference")?;
    if !file.starts_with("crates/") || !file.ends_with(".rs") || file.contains("..") {
        return Err(format!("invalid acceptance test path: {file}"));
    }
    let source = fs::read_to_string(root.join(file)).map_err(|e| format!("{file}: {e}"))?;
    if !source.contains(&format!("fn {function}(")) {
        return Err(format!("missing acceptance test function: {test}"));
    }
    Ok(())
}

/// Evidence ledger for the selected High profiles, independent of historical source counts.
fn validate_high_acceptance(root: &Path, source_rows: &[Value]) -> Result<BTreeMap<String, BTreeSet<String>>, String> {
    let path = root.join("docs/vba/reference/vba-object-model-high-acceptance.json");
    let ledger: Value = serde_json::from_slice(&fs::read(&path).map_err(|e| format!("{}: {e}", path.display()))?).map_err(|e| e.to_string())?;
    if ledger.get("schema").and_then(Value::as_u64) != Some(1) {
        return Err("unsupported High acceptance schema".into());
    }
    let rows = ledger.get("requirements").and_then(Value::as_array).ok_or("missing High requirements")?;
    let ids = source_rows.iter().filter_map(|row| row.get("id").and_then(Value::as_str)).map(fold_identifier).collect::<BTreeSet<_>>();
    let mut tests_by_source = BTreeMap::<String, BTreeSet<String>>::new();
    let mut obligations = BTreeSet::new();
    for row in rows {
        let id = string(row, "id")?;
        if !obligations.insert(id.to_owned()) {
            return Err(format!("duplicate High obligation: {id}"));
        }
        if row.get("required").and_then(Value::as_bool) != Some(true) || string(row, "status")? != "IMPLEMENTED_NATIVE_PROFILE" {
            return Err(format!("unresolved required High obligation: {id}"));
        }
        let tests = row.get("test_ids").and_then(Value::as_array).ok_or("High obligation without tests")?;
        if tests.is_empty() {
            return Err(format!("High obligation without test evidence: {id}"));
        }
        let mut test_ids = BTreeSet::new();
        for test in tests {
            let test = test.as_str().ok_or("invalid High test id")?;
            validate_test_reference(root, test)?;
            test_ids.insert(test.to_owned());
        }
        for file in row.get("native_evidence").and_then(Value::as_array).ok_or("High obligation without native evidence")? {
            let file = file.as_str().ok_or("invalid native evidence")?;
            if !file.starts_with("crates/") || file.contains("..") || !root.join(file).is_file() {
                return Err(format!("missing native High evidence: {file}"));
            }
        }
        for source in row.get("source_ids").and_then(Value::as_array).ok_or("High obligation without source ids")? {
            let source = fold_identifier(source.as_str().ok_or("invalid High source id")?);
            if !ids.contains(&source) {
                return Err(format!("unknown High source contract: {source}"));
            }
            tests_by_source.entry(source).or_default().extend(test_ids.iter().cloned());
        }
    }
    println!(
        "High native acceptance: {} required obligations with linked Rust tests; platform/scope exclusions retained separately",
        obligations.len()
    );
    Ok(tests_by_source)
}

fn contract_row(row: &Value, markdown: &str, descriptors: &BTreeMap<(String, String), MemberDescriptor>) -> Result<Value, String> {
    let id = string(row, "id")?;
    let kind = string(row, "kind")?;
    let owner = string(row, "owner")?;
    let member = string(row, "member")?;
    let status = string(row, "status")?;
    let source_path = string(row, "source_path")?;
    let packet = string(row, "packet")?;
    let url = string(row, "url")?;
    let audit_evidence = string(row, "evidence")?;
    let api_names = frontmatter_api_names(markdown);
    let signatures = signatures(markdown);
    let parameters = parameters(markdown);
    let return_type = return_type(markdown);
    let enum_values = enum_values(markdown);
    let enum_type = (kind == "enumeration").then_some(owner);
    // Events and callable properties/methods are distinct source contracts even
    // when Microsoft gives them identical owner/member names.
    let descriptor = if kind == "event" { None } else { descriptors.get(&(fold_identifier(owner), fold_identifier(member))) };
    let implementation_symbol = if kind == "enumeration" {
        Some("crates/vba/src/vba_enum_data.rs::EXCEL_ENUM_CONSTANTS")
    } else if kind == "event" && matches!(id, "Workbook.Open" | "Worksheet.Change") {
        Some("crates/vba/src/lib.rs::Bridge::tick")
    } else if kind == "event"
        && matches!(
            id,
            "Workbook.BeforeSave"
                | "Workbook.BeforeClose"
                | "Workbook.AfterSave"
                | "Application.WorkbookBeforeSave"
                | "Application.WorkbookBeforeClose"
                | "Application.WorkbookAfterSave"
        )
    {
        Some("crates/vba/src/host_ops.rs::Bridge::lifecycle_event_calls")
    } else {
        descriptor.map(|item| item.implementation_symbol)
    };
    let (mut result_state, mut profile) = match status {
        "implemented-subset" | "object-subset" | "context-only" | "partial-definition" => ("PARTIAL", Some("GridCraft.Excel.v1")),
        "platform-service-needed" => ("PLATFORM_SERVICE_NEEDED", None),
        "reference-triage" => ("LEGACY_REVIEW", None),
        "object-not-exposed" | "not-exposed" | "formula-bridge-not-exposed" | "definition-gap" => ("MISSING", None),
        _ => ("LEGACY_REVIEW", None),
    };
    // Baseline rows are historical. Current shared dispatch evidence supersedes
    // their old missing state, without certifying every Excel option or signature.
    if let Some(descriptor) = descriptor {
        (result_state, profile) = match descriptor.availability {
            MemberAvailability::Implemented => ("PARTIAL", Some("GridCraft.Excel.v1")),
            MemberAvailability::ServiceRequired => ("PLATFORM_SERVICE_NEEDED", None),
            MemberAvailability::OutOfScope => ("OUT_OF_SCOPE_GRIDCRAFT", None),
        };
    }
    let enum_values_complete = kind == "enumeration" && enum_table_member_count(markdown) == enum_values.len() && !enum_values.is_empty();
    if kind == "enumeration" {
        if enum_values_complete {
            result_state = "PARTIAL";
            profile = Some("GridCraft.Excel.v1");
        } else {
            result_state = "MISSING";
            profile = None;
        }
    }
    if kind == "event" && implementation_symbol.is_some() {
        result_state = "PARTIAL";
        profile = Some("GridCraft.Excel.v1");
    }
    let mut arguments = Vec::new();
    for (signature_index, signature) in signatures.iter().enumerate() {
        for (position, name) in signature.get("arguments").and_then(Value::as_array).into_iter().flatten().filter_map(Value::as_str).enumerate() {
            let record = parameters.get(&fold_identifier(name));
            arguments.push(json!({
                "signature_index": signature_index,
                "member": signature.get("member"),
                "position": position,
                "name": name,
                "required": record.and_then(|v| v.get("required")).cloned().unwrap_or(Value::Null),
                "required_source": record.and_then(|v| v.get("required_source")).cloned().unwrap_or(json!("NOT_SPECIFIED_IN_SOURCE")),
                "type": record.and_then(|v| v.get("type")).cloned().unwrap_or(Value::Null),
                "type_source": record.and_then(|v| v.get("type_source")).cloned().unwrap_or(json!("NOT_SPECIFIED_IN_SOURCE")),
                "default": record.and_then(|v| v.get("default")).cloned().unwrap_or(Value::Null),
                "default_source": record.and_then(|v| v.get("default_source")).cloned().unwrap_or(json!("NO_EXPLICIT_DEFAULT_IN_SOURCE")),
                "mode": "NOT_SPECIFIED_IN_SOURCE",
            }));
        }
    }
    let cancellation_parameters = if kind == "event" {
        arguments
            .iter()
            .filter(|argument| argument.get("name").and_then(Value::as_str).is_some_and(|name| name.to_ascii_lowercase().contains("cancel")))
            .map(|argument| argument.get("name").cloned().unwrap_or(Value::Null))
            .collect::<Vec<_>>()
    } else {
        Vec::new()
    };
    let api_aliases: Vec<String> = api_names.into_iter().filter(|name| name != id).collect();
    let signature_aliases = signatures
        .iter()
        .filter_map(|signature| signature.get("member").and_then(Value::as_str))
        .filter(|name| !format!("{owner}.{name}").eq_ignore_ascii_case(id))
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let related_member_ids = related_members(markdown).into_iter().filter(|candidate| !candidate.eq_ignore_ascii_case(id)).collect::<Vec<_>>();
    let signature_review = if kind == "enumeration" {
        if enum_values_complete { "ENUM_VALUES_EXTRACTED_RUNTIME_TESTS_PENDING" } else { "ENUM_VALUES_INCOMPLETE" }
    } else if matches!(kind, "object" | "reference") {
        "NOT_APPLICABLE"
    } else if !signatures.is_empty() {
        "SOURCE_EXTRACTED_WITH_PROVENANCE"
    } else if !related_member_ids.is_empty() {
        "REFERENCE_PAGE_POINTS_TO_CANONICAL_MEMBER"
    } else if status == "reference-triage" {
        "RECEIVER_OR_ALIAS_TRIAGE"
    } else {
        "SYNTAX_NOT_FOUND"
    };
    Ok(json!({
        "id": id,
        "source_page_ids": [id],
        "source_deprecated": markdown.to_ascii_lowercase().contains("deprecated"),
        "source_namespace":if markdown.lines().any(|line|line.starts_with("# ") && line.contains("(Excel Graph)")){"ExcelGraph"}else{"Excel"},
        "source_heading":markdown.lines().find_map(|line|line.strip_prefix("# ")).unwrap_or(""),
        "source_path": source_path,
        "url": url,
        "packet": packet,
        "baseline_status": status,
        "audit_evidence": audit_evidence,
        "aliases": api_aliases,
        "syntax_member_aliases": signature_aliases,
        "related_member_ids": related_member_ids,
        "source_api_names": frontmatter_api_names(markdown),
        "kind": kind,
        "owner": owner,
        "member": member,
        "signatures": signatures,
        "arguments": arguments,
        "return_type": return_type,
        "read_write": property_access(kind, markdown),
        "cancellation_parameters": cancellation_parameters,
        "enum_type": enum_type,
        "enum_values": enum_values,
        "enum_aliases": enum_value_aliases(markdown),
        "enum_alias_review": if kind == "enumeration" { "PRESERVED_AS_TYPED_SOURCE_ALIASES" } else { "NOT_APPLICABLE" },
        "enum_values_complete": enum_values_complete,
        "platform_notes": platform_notes(markdown),
        "platform_review": "SOURCE_REVIEWED",
        "dependencies": [],
        "profile": profile,
        "implementation_symbol": implementation_symbol,
        "descriptor_availability": descriptor.map(|item| match item.availability {
            MemberAvailability::Implemented => "IMPLEMENTED",
            MemberAvailability::ServiceRequired => "SERVICE_REQUIRED",
            MemberAvailability::OutOfScope=>"OUT_OF_SCOPE_GRIDCRAFT",
        }),
        "dispatch_evidence": if descriptor.is_some() { "shared-member-descriptor" } else if implementation_symbol.is_some() { "explicit-event-or-enum" } else { "none" },
        "test_ids": if descriptor.is_some_and(|d| d.implementation_symbol.starts_with("crates/vba/src/chart_api.rs")) {json!(OM12_TESTS)} else if descriptor.is_some_and(|d| ["series_api.rs","chart_parts_api.rs","drawing_api.rs"].iter().any(|file|d.implementation_symbol.contains(file))) {json!(OM12_DRAWING_TESTS)} else {json!([])},
        "result_state": result_state,
        "signature_review": signature_review,
    }))
}

fn property_access(kind: &str, markdown: &str) -> &'static str {
    if kind != "property" {
        return "NOT_APPLICABLE";
    }
    let lower = markdown.to_ascii_lowercase();
    if ["returns or sets", "gets or sets", "can be set to", "read/write"].iter().any(|phrase| lower.contains(phrase)) {
        "READ_WRITE"
    } else if ["returns ", "gets ", "read-only", "returns the"].iter().any(|phrase| lower.contains(phrase)) {
        "READ_ONLY"
    } else {
        "NOT_SPECIFIED_IN_SOURCE"
    }
}

fn platform_notes(markdown: &str) -> Vec<String> {
    let mut notes = BTreeSet::new();
    for paragraph in markdown.split("\n\n") {
        let lower = paragraph.to_ascii_lowercase();
        if lower.contains("windows") || lower.contains("macintosh") || lower.contains("mac os") || lower.contains("office for mac") {
            let text = paragraph.split_whitespace().collect::<Vec<_>>().join(" ");
            if !text.is_empty() && text.len() <= 1024 {
                notes.insert(text);
            }
        }
    }
    notes.into_iter().collect()
}

fn string<'a>(value: &'a Value, key: &str) -> Result<&'a str, String> {
    value.get(key).and_then(Value::as_str).ok_or_else(|| format!("inventory row has no {key}"))
}

fn frontmatter_api_names(markdown: &str) -> Vec<String> {
    let mut names = Vec::new();
    let mut in_frontmatter = false;
    for line in markdown.lines() {
        if line.trim() == "---" {
            if in_frontmatter {
                break;
            }
            in_frontmatter = true;
            continue;
        }
        if in_frontmatter && let Some(value) = line.trim().strip_prefix("- Excel.") {
            names.push(value.trim().to_owned());
        }
    }
    names
}

fn signatures(markdown: &str) -> Vec<Value> {
    let mut in_syntax = false;
    let mut out = Vec::<Value>::new();
    for line in markdown.lines() {
        if line.trim().eq_ignore_ascii_case("## Syntax") {
            in_syntax = true;
            continue;
        }
        if in_syntax && line.starts_with("## ") {
            break;
        }
        if !in_syntax || !line.contains("_expression_.") {
            continue;
        }
        let Some(start) = line.find("**") else {
            continue;
        };
        let Some(end_rel) = line[start + 2..].find("**") else {
            continue;
        };
        let end = start + 2 + end_rel;
        let member = clean_identifier(&line[start + 2..end]);
        if member.is_empty() {
            continue;
        }
        let tail = &line[end + 2..];
        let arguments = if let Some(open) = tail.find('(') {
            if let Some(close) = tail[open + 1..].find(')') {
                tail[open + 1..open + 1 + close].split(',').map(clean_identifier).filter(|name| !name.is_empty()).collect::<Vec<_>>()
            } else {
                Vec::new()
            }
        } else {
            Vec::new()
        };
        out.push(json!({"member":member,"arguments":arguments}));
    }
    let mut seen = BTreeSet::new();
    out.into_iter().filter(|signature| seen.insert(signature.to_string())).collect()
}

fn related_members(markdown: &str) -> Vec<String> {
    let mut related = BTreeSet::new();
    for line in markdown.lines() {
        let lower = line.to_ascii_lowercase();
        let Some(index) = lower.find("range.") else {
            continue;
        };
        let tail = &line[index..];
        let end = tail.find(|c: char| !c.is_ascii_alphanumeric() && c != '.' && c != '_').unwrap_or(tail.len());
        let candidate = tail[..end].trim_end_matches('.');
        if candidate.matches('.').count() == 1 && candidate.len() <= 128 {
            related.insert(candidate.to_owned());
        }
    }
    related.into_iter().collect()
}

fn parameters(markdown: &str) -> BTreeMap<String, Value> {
    let mut out = BTreeMap::new();
    let mut in_parameters = false;
    let mut in_table = false;
    for line in markdown.lines() {
        if line.trim().eq_ignore_ascii_case("## Parameters") {
            in_parameters = true;
            continue;
        }
        if in_parameters && line.starts_with("## ") {
            break;
        }
        if !in_parameters {
            continue;
        }
        if line.starts_with('|') {
            in_table = true;
            let cells: Vec<_> = line.trim_matches('|').split('|').map(str::trim).collect();
            if cells.len() < 3 || cells[0].contains("---") || cells[0].eq_ignore_ascii_case("name") {
                continue;
            }
            let name = clean_identifier(cells[0]);
            if name.is_empty() {
                continue;
            }
            let required = match cells[1].to_ascii_lowercase().as_str() {
                "required" => Some(true),
                "optional" => Some(false),
                _ => None,
            };
            let data_type = clean_type(cells[2]);
            let description = cells.get(3).copied().unwrap_or_default();
            let default = extract_default(description);
            let required_source = if required.is_some() { "PARAMETERS_TABLE" } else { "NOT_SPECIFIED_IN_SOURCE" };
            let type_source = if data_type.is_empty() { "NOT_SPECIFIED_IN_SOURCE" } else { "PARAMETERS_TABLE" };
            let default_source = if default.is_some() { "PARAMETER_DESCRIPTION" } else { "NO_EXPLICIT_DEFAULT_IN_SOURCE" };
            out.insert(
                fold_identifier(&name),
                json!({
                    "required": required,
                    "type": data_type,
                    "required_source": required_source,
                    "type_source": type_source,
                    "default": default.unwrap_or(Value::Null),
                    "default_source": default_source,
                }),
            );
        } else if in_table && !line.trim().is_empty() {
            in_table = false;
        }
    }
    out
}

fn return_type(markdown: &str) -> Value {
    let mut in_return = false;
    for line in markdown.lines() {
        if line.trim().eq_ignore_ascii_case("## Return value") {
            in_return = true;
            continue;
        }
        if in_return && line.starts_with("## ") {
            break;
        }
        if !in_return || line.trim().is_empty() || line.trim().starts_with('|') {
            continue;
        }
        let ty = clean_type(line.trim());
        if !ty.is_empty() && ty.len() <= 128 && !ty.contains(' ') {
            return Value::String(ty);
        }
    }
    Value::Null
}

fn enum_values(markdown: &str) -> Vec<Value> {
    let mut values = Vec::new();
    let mut in_table = false;
    for line in markdown.lines() {
        if line.starts_with('|') && line.to_ascii_lowercase().contains("name") && line.to_ascii_lowercase().contains("value") {
            in_table = true;
            continue;
        }
        let inline_row = line.starts_with('|') && line.trim_start_matches('|').trim_start().starts_with("**") && line.contains("**|");
        if !in_table && !inline_row {
            continue;
        }
        if !line.starts_with('|') {
            in_table = false;
            continue;
        }
        let cells: Vec<_> = line.trim_matches('|').split('|').map(str::trim).collect();
        if cells.len() < 2 || cells[0].contains("---") {
            continue;
        }
        let name = clean_identifier(cells[0]);
        if name.is_empty() {
            continue;
        }
        let raw_value = cells[1].replace("**", "");
        let Some(number) = integer_prefix(raw_value.trim()) else {
            continue;
        };
        values.push(json!({"name":name,"number":number,"source_value":raw_value.trim()}));
    }
    values.sort_by(|a, b| a.get("name").and_then(Value::as_str).cmp(&b.get("name").and_then(Value::as_str)));
    values
}

fn enum_table_member_count(markdown: &str) -> usize {
    let mut count = 0;
    let mut in_table = false;
    for line in markdown.lines() {
        if line.starts_with('|') && line.to_ascii_lowercase().contains("name") && line.to_ascii_lowercase().contains("value") {
            in_table = true;
            continue;
        }
        let inline_row = line.starts_with('|') && line.trim_start_matches('|').trim_start().starts_with("**") && line.contains("**|");
        if !in_table && !inline_row {
            continue;
        }
        if !line.starts_with('|') {
            in_table = false;
            continue;
        }
        let cells: Vec<_> = line.trim_matches('|').split('|').map(str::trim).collect();
        if cells.len() >= 2 && !cells[0].contains("---") && !clean_identifier(cells[0]).is_empty() {
            count += 1;
        }
    }
    count
}

fn enum_value_aliases(markdown: &str) -> Vec<Value> {
    let mut groups = BTreeMap::<i64, Vec<String>>::new();
    for value in enum_values(markdown) {
        if let (Some(number), Some(name)) = (value.get("number").and_then(Value::as_i64), value.get("name").and_then(Value::as_str)) {
            groups.entry(number).or_default().push(name.to_owned());
        }
    }
    groups.into_iter().filter(|(_, names)| names.len() > 1).map(|(number, names)| json!({"number":number,"names":names})).collect()
}

fn integer_prefix(value: &str) -> Option<i64> {
    let first = value.split_whitespace().next()?.trim_end_matches(',');
    if first.starts_with("&H") || first.starts_with("&h") {
        let raw = i64::from_str_radix(first.get(2..)?, 16).ok()?;
        return Some(if first.get(2..)?.len() == 8 && raw > i32::MAX as i64 { raw - (1i64 << 32) } else { raw });
    }
    first.parse().ok()
}

fn clean_identifier(value: &str) -> String {
    let value = value.trim().trim_start_matches('_').trim_end_matches('_');
    let value = value.trim_matches('*').trim();
    let value = value.split_whitespace().next().unwrap_or_default();
    value.trim_matches(|c: char| !c.is_ascii_alphanumeric() && c != '_').to_owned()
}

fn clean_type(value: &str) -> String {
    let value = value.trim().trim_matches('*');
    if let Some((label, _)) = value.split_once("](") {
        return label.trim_start_matches('[').to_owned();
    }
    value.trim_matches(|c: char| c == '`' || c == '[' || c == ']' || c == '*').trim().to_owned()
}

fn extract_default(description: &str) -> Option<Value> {
    let lower = description.to_ascii_lowercase();
    let phrases = ["default value is ", "default is "];
    let (index, phrase) = phrases.iter().filter_map(|phrase| lower.find(phrase).map(|index| (index, *phrase))).min_by_key(|(index, _)| *index)?;
    if description[..index].trim_end().to_ascii_lowercase().ends_with("no") {
        return None;
    }
    let value = description[index + phrase.len()..].split(['.', ';', ',']).next()?.trim();
    let value = value.trim_matches(|c: char| matches!(c, '*' | '_' | '`' | ' '));
    if value.is_empty() {
        None
    } else if value.eq_ignore_ascii_case("true") {
        Some(Value::Bool(true))
    } else if value.eq_ignore_ascii_case("false") {
        Some(Value::Bool(false))
    } else if let Ok(number) = value.parse::<i64>() {
        Some(json!(number))
    } else {
        Some(Value::String(value.to_owned()))
    }
}

fn fold_identifier(value: &str) -> String {
    value.to_ascii_lowercase()
}

fn deduplicate_enum_constants(values: Vec<(String, String, i32)>) -> Vec<(String, String, i32)> {
    let mut seen = BTreeSet::new();
    values.into_iter().filter(|(_, name, _)| seen.insert(name.to_ascii_lowercase())).collect()
}

fn duplicate_enum_constants(values: &[(String, String, i32)]) -> Vec<(String, String, i32)> {
    let mut seen = BTreeSet::new();
    values.iter().filter(|(_, name, _)| !seen.insert(name.to_ascii_lowercase())).cloned().collect()
}

fn render_enum_source(types: &BTreeSet<String>, values: &[(String, String, i32)], aliases: &[(String, String, i32)]) -> String {
    let mut out = String::from("// Generated from the pinned VBA-Docs enum pages; regenerate with cargo xtask vba-audit --write.\n");
    out.push_str("pub(super) const EXCEL_ENUM_TYPES: &[&str] = &[\n");
    for ty in types {
        out.push_str(&format!("    {ty:?},\n"));
    }
    out.push_str("];\npub(super) const EXCEL_ENUM_CONSTANTS: &[(&str, &str, i32)] = &[\n");
    for (ty, name, value) in values {
        out.push_str(&format!("    ({ty:?}, {name:?}, {value}),\n"));
    }
    out.push_str("];\npub(super) const EXCEL_ENUM_ALIASES: &[(&str, &str, i32)] = &[\n");
    for (ty, name, value) in aliases {
        out.push_str(&format!("    ({ty:?}, {name:?}, {value}),\n"));
    }
    out.push_str("];\n");
    out
}

fn render_report(contracts: &Value) -> Result<String, String> {
    let pages = contracts.get("pages").and_then(Value::as_array).ok_or("contracts has no pages")?;
    let mut counts = BTreeMap::<String, usize>::new();
    for row in pages {
        let state = row.get("result_state").and_then(Value::as_str).unwrap_or("LEGACY_REVIEW");
        *counts.entry(state.to_owned()).or_default() += 1;
    }
    let mut out = format!(
        "# VBA source catalogue\n\nThis generated index maps the pinned public Excel reference to GridCraft's adapter. For practical feature guidance, read [supported features](../supported.md) and [unsupported features](../unsupported.md). Catalogue PARTIAL/MISSING states describe source coverage, not promises about every Excel option.\n\nPinned source: MicrosoftDocs/VBA-Docs `{COMMIT}` (SHA256 `{ARCHIVE_SHA256}`). {} source pages retained; source-page counts are not compatibility percentages.\n\n## Source coverage\n\n",
        pages.len()
    );
    for (state, count) in counts {
        out.push_str(&format!("- `{state}`: {count}\n"));
    }
    out.push_str("\n## Native profile evidence\n\n[Native profile and scope](vba-object-model-profile-report.md) keeps planning disposition independent of raw catalogue evidence. Exact source paths and method/event kinds remain distinct.\n\n## Unresolved audit rows\n\nRows in `MISSING`, `PARTIAL`, `PLATFORM_SERVICE_NEEDED`, `LEGACY_REVIEW` or `REFERENCE_ALIAS` remain unresolved against the Microsoft catalogue. They are not automatically in-scope implementation requirements: the existing-GridCraft-capability gate is applied first. This report deliberately does not calculate an overall compatibility percentage.\n\n| ID | Source path | State | Priority | Profile disposition | Signature review |\n|---|---|---|---|---|---|\n");
    for row in pages {
        let state = row.get("result_state").and_then(Value::as_str).unwrap_or("LEGACY_REVIEW");
        if state == "IMPLEMENTED_TESTED" {
            continue;
        }
        let id = row.get("id").and_then(Value::as_str).unwrap_or("UNKNOWN");
        let review = row.get("signature_review").and_then(Value::as_str).unwrap_or("UNREVIEWED");
        let source = row["source_path"].as_str().unwrap_or("");
        let priority = row["v1_priority"].as_str().unwrap_or("");
        let disposition = row["v1_disposition"].as_str().unwrap_or("");
        out.push_str(&format!("| `{id}` | `{source}` | `{state}` | `{priority}` | `{disposition}` | `{review}` |\n"));
    }
    Ok(out)
}

fn count_enum_types(contracts: &[Value]) -> usize {
    contracts.iter().filter_map(|row| row.get("enum_type").and_then(Value::as_str)).collect::<BTreeSet<_>>().len()
}

fn compare_generated(path: &Path, expected: &Value) -> Result<(), String> {
    let actual: Value = serde_json::from_slice(&fs::read(path).map_err(|e| format!("{}: {e}; run vba-audit --write", path.display()))?)
        .map_err(|e| format!("{}: invalid JSON: {e}", path.display()))?;
    if &actual != expected {
        return Err(format!("{} is stale; run cargo xtask vba-audit --write", path.display()));
    }
    Ok(())
}

fn compare_generated_text(path: &Path, expected: &str) -> Result<(), String> {
    let actual = fs::read_to_string(path).map_err(|e| format!("{}: {e}; run vba-audit --write", path.display()))?;
    if actual != expected {
        return Err(format!("{} is stale; run cargo xtask vba-audit --write", path.display()));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{contract_row, deduplicate_enum_constants, duplicate_enum_constants, extract_default};
    use serde_json::Value;

    #[test]
    fn om10_keeps_pdf_layout_separate_from_printers_and_header_graphics() {
        for (owner, member, state) in [
            ("PageSetup", "PrintQuality", "PLATFORM_SERVICE_NEEDED"),
            ("Graphic", "Filename", "OUT_OF_SCOPE_GRIDCRAFT"),
            ("Chart", "PageSetup", "OUT_OF_SCOPE_GRIDCRAFT"),
            ("PageSetup", "", "PARTIAL"),
        ] {
            let mut row = serde_json::json!({"packet":"OM10","owner":owner,"member":member,"result_state":"MISSING"});
            super::reconcile_om10_native_scope(&mut row);
            assert_eq!(row["result_state"], state);
        }
    }

    #[test]
    fn om09_never_substitutes_consent_or_opaque_metadata_for_document_permissions() {
        for (owner, member) in [
            ("Workbook", "Permission"),
            ("Workbook", "HasPassword"),
            ("AllowEditRange", "Users"),
            ("UserAccess", "AllowEdit"),
            ("ProtectedViewWindow", "Edit"),
            ("Chart", "Protect"),
        ] {
            let mut row = serde_json::json!({"packet":"OM09","owner":owner,"member":member,"result_state":"MISSING"});
            super::reconcile_om09_native_scope(&mut row);
            assert_eq!(row["result_state"], "OUT_OF_SCOPE_GRIDCRAFT");
            assert!(row["native_capability_gate"].is_string());
        }
        let mut row = serde_json::json!({"packet":"OM09","owner":"Protection","member":"","result_state":"MISSING"});
        super::reconcile_om09_native_scope(&mut row);
        assert_eq!(row["result_state"], "PARTIAL");
    }

    #[test]
    fn om08_distinguishes_missing_window_state_from_unconnected_ui_services() {
        for (owner, member, state) in [
            ("Window", "Zoom", "OUT_OF_SCOPE_GRIDCRAFT"),
            ("Pane", "VisibleRange", "OUT_OF_SCOPE_GRIDCRAFT"),
            ("Window", "PointsToScreenPixelsX", "PLATFORM_SERVICE_NEEDED"),
            ("Application", "WindowState", "PLATFORM_SERVICE_NEEDED"),
            ("Window", "", "PARTIAL"),
        ] {
            let mut row = serde_json::json!({"packet":"OM08","owner":owner,"member":member,"result_state":"MISSING"});
            super::reconcile_om08_native_scope(&mut row);
            assert_eq!(row["result_state"], state);
        }
        let mut row =
            serde_json::json!({"packet":"OM08","owner":"Window","member":"ActiveCell","implementation_symbol":"native","result_state":"PARTIAL"});
        super::reconcile_om08_native_scope(&mut row);
        assert_eq!(row["result_state"], "PARTIAL");
    }

    #[test]
    fn extracts_typed_parameter_defaults_from_documentation() {
        assert_eq!(extract_default("The default is **False**."), Some(Value::Bool(false)));
        assert_eq!(extract_default("The default value is **True**."), Some(Value::Bool(true)));
        assert_eq!(extract_default("The default is **-4142**."), Some(serde_json::json!(-4142)));
        assert_eq!(extract_default("No default is documented."), None);
    }

    #[test]
    fn runtime_constant_table_deduplicates_names_preserved_in_source_contracts() {
        let constants = deduplicate_enum_constants(vec![
            ("Constants".into(), "xlManual".into(), -4135),
            ("XlSortOrder".into(), "xlManual".into(), -4135),
            ("Constants".into(), "xlManual".into(), -4135),
        ]);
        assert_eq!(constants.len(), 1);
        assert_eq!(constants[0].1, "xlManual");
        let aliases = duplicate_enum_constants(&[("Constants".into(), "xlManual".into(), -4135), ("XlSortOrder".into(), "xlManual".into(), -4135)]);
        assert_eq!(aliases, vec![("XlSortOrder".into(), "xlManual".into(), -4135)]);
    }
    #[test]
    fn current_dispatch_updates_result_without_rewriting_historical_inventory() -> Result<(), String> {
        let descriptors =
            gridcraft_vba::member_descriptors().into_iter().map(|d| ((d.owner.to_ascii_lowercase(), d.name.to_ascii_lowercase()), d)).collect();
        for (member, expected) in [("Workbooks", "PARTIAL"), ("Run", "PARTIAL"), ("ScreenUpdating", "OUT_OF_SCOPE_GRIDCRAFT")] {
            let row = serde_json::json!({
                "id":format!("Application.{member}"), "kind":"property", "owner":"Application", "member":member,
                "status":"not-exposed", "source_path":format!("api/Excel.Application.{member}.md"), "packet":"OM03",
                "url":"https://example.invalid/reference", "evidence":"historical audit"
            });
            let contract = contract_row(&row, "", &descriptors)?;
            assert_eq!(contract["baseline_status"], "not-exposed");
            assert_eq!(contract["result_state"], expected);
            assert_eq!(contract["dispatch_evidence"], "shared-member-descriptor");
        }
        Ok(())
    }
    #[test]
    fn method_activation_does_not_claim_same_named_events() -> Result<(), String> {
        let descriptors =
            gridcraft_vba::member_descriptors().into_iter().map(|d| ((d.owner.to_ascii_lowercase(), d.name.to_ascii_lowercase()), d)).collect();
        for owner in ["Workbook", "Worksheet"] {
            for kind in ["event", "method"] {
                let row = serde_json::json!({"id":format!("{owner}.Activate"),"kind":kind,"owner":owner,"member":"Activate","status":"not-exposed","source_path":format!("api/Excel.{owner}.Activate({kind}).md"),"packet":if kind=="event"{"OM11"}else{"OM04"},"url":"https://example.invalid/activation","evidence":"source kind remains distinct"});
                let contract = super::contract_row(&row, "", &descriptors)?;
                assert_eq!(contract["implementation_symbol"].is_string(), kind == "method");
                assert_eq!(contract["descriptor_availability"].is_string(), kind == "method");
            }
        }
        Ok(())
    }
    #[test]
    fn headerless_enum_tables_ignore_the_concatenated_pseudo_header() {
        let source = "## Members\n|xlFirstxlSecond|12|combined description|\n| **xlFirst**| **1**||\n| **xlSecond**| **2**||\n";
        let values = super::enum_values(source);
        assert_eq!(values.len(), 2);
        assert_eq!(super::enum_table_member_count(source), 2);
        assert_eq!(values[0]["name"], "xlFirst");
    }
}
