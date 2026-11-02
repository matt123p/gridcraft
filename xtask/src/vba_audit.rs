//! Offline integrity check for the pinned Excel VBA object-model inventory.
//!
//! The documentation archive is cached under `target/` and is never fetched by
//! this command. The committed JSON remains the reviewable inventory; this check
//! proves that its source page set still matches the pinned archive.

use serde_json::Value;
use sha2::{Digest, Sha256};
use std::collections::BTreeSet;
use std::fs::File;
use std::path::Path;
use zip::ZipArchive;

const ARCHIVE_SHA256: &str = "285218108a034fc33acc8296d714395fd182397ada27c848b93c3d5423efd75e";
const COMMIT: &str = "b202c96d9d53896593b73d6fae57f5230601a5e3";
const PAGE_COUNT: usize = 5_996;

pub fn run(root: &Path) -> Result<(), String> {
    let cache = root.join("target/vba-docs-audit");
    let archive_path = cache.join("reference.zip");
    let inventory_path = root.join("docs/vba-object-model-inventory.json");
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
    println!("VBA object-model inventory: {PAGE_COUNT} source pages match pinned commit {COMMIT} (sha256 {digest})");
    Ok(())
}
