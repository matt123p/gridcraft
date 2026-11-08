/// Publish an immutable companion first, then atomically replace the .frm commit point.
/// Existing companion bytes are never overwritten; collisions refuse. No acknowledgment here.
#[cfg(not(target_arch = "wasm32"))]
pub fn write_native_form(path: &std::path::Path, snapshot: &gridcraft_vba::vba_ide::NativeFormSave) -> Result<(), Box<dyn std::error::Error>> {
    use std::io::Read;
    let frm = snapshot.files().first().ok_or("native form file missing")?;
    let frx = snapshot.files().get(1).ok_or("native companion missing")?;
    if !path.extension().is_some_and(|extension| extension.eq_ignore_ascii_case("frm")) {
        return Err("Choose a native .frm file".into());
    }
    let parent = path.parent().filter(|p| !p.as_os_str().is_empty()).unwrap_or_else(|| std::path::Path::new("."));
    let permissions = std::fs::metadata(path).ok().filter(|metadata| metadata.is_file()).map(|metadata| metadata.permissions());
    let companion = parent.join(&frx.name);
    let staged = stage_native_file(parent, &frx.name, &frx.bytes, permissions.as_ref())?;
    let published = std::fs::hard_link(&staged, &companion);
    let _ = std::fs::remove_file(&staged);
    match published {
        Ok(()) => {}
        Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => {
            let metadata = std::fs::symlink_metadata(&companion)?;
            if !metadata.file_type().is_file() || metadata.len() != frx.bytes.len() as u64 {
                return Err("Native companion collision".into());
            }
            let mut bytes = Vec::new();
            std::fs::File::open(&companion)?.take(4 * 1024 * 1024 + 1).read_to_end(&mut bytes)?;
            if bytes != frx.bytes {
                return Err("Native companion collision".into());
            }
        }
        Err(error) => return Err(error.into()),
    }
    let temp = stage_native_file(parent, &frm.name, &frm.bytes, permissions.as_ref())?;
    let result = std::fs::rename(&temp, path);
    if result.is_err() {
        let _ = std::fs::remove_file(&temp);
    }
    result.map_err(Into::into)
}
#[cfg(not(target_arch = "wasm32"))]
fn stage_native_file(
    parent: &std::path::Path,
    name: &str,
    bytes: &[u8],
    permissions: Option<&std::fs::Permissions>,
) -> Result<std::path::PathBuf, Box<dyn std::error::Error>> {
    use std::io::Write;
    for attempt in 0..128 {
        let temp = parent.join(format!(".{name}.save-{}-{attempt}.tmp", std::process::id()));
        let mut options = std::fs::OpenOptions::new();
        options.write(true).create_new(true);
        #[cfg(unix)]
        if let Some(permissions) = permissions {
            use std::os::unix::fs::{OpenOptionsExt, PermissionsExt};
            options.mode(permissions.mode());
        }
        let mut file = match options.open(&temp) {
            Ok(file) => file,
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        };
        if let Err(error) = file.write_all(bytes).and_then(|()| file.sync_all()) {
            let _ = std::fs::remove_file(&temp);
            return Err(error.into());
        }
        #[cfg(not(unix))]
        if let Some(permissions) = permissions
            && let Err(error) = file.set_permissions(permissions.clone())
        {
            let _ = std::fs::remove_file(&temp);
            return Err(error.into());
        }
        return Ok(temp);
    }
    Err("Native save staging limit".into())
}

/// The owning host passes an encoded macro-enabled package after revision checks.
pub fn write_package(path: &std::path::Path, bytes: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let parent = path.parent().filter(|p| !p.as_os_str().is_empty()).unwrap_or_else(|| std::path::Path::new("."));
    let name = path.file_name().and_then(|n| n.to_str()).ok_or("package filename missing")?;
    let permissions = std::fs::metadata(path).ok().filter(|metadata| metadata.is_file()).map(|metadata| metadata.permissions());
    let temp = stage_native_file(parent, name, bytes, permissions.as_ref())?;
    let result = std::fs::rename(&temp, path);
    if result.is_err() {
        let _ = std::fs::remove_file(&temp);
    }
    result.map_err(Into::into)
}
