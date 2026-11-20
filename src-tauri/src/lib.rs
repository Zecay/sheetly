use std::path::PathBuf;
use tauri::{ipc::Response, Manager};

fn data_file(app: &tauri::AppHandle, name: &str) -> Result<PathBuf, String> {
    // Only plain file names are allowed in the app data folder.
    if name.is_empty() || name.contains('/') || name.contains('\\') || name.contains("..") {
        return Err("bad name".into());
    }
    let dir = app.path().app_data_dir().map_err(|e| e.to_string())?;
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir.join(name))
}

const MAX_READ_BYTES: u64 = 256 * 1024 * 1024;

fn read_limited(path: &std::path::Path) -> Result<Vec<u8>, String> {
    let len = std::fs::metadata(path).map_err(|e| e.to_string())?.len();
    if len > MAX_READ_BYTES {
        return Err("File is too large to open (over 256 MB)".into());
    }
    std::fs::read(path).map_err(|e| e.to_string())
}

/// Write to a temp file in the same folder, then rename over the target (atomic save).
fn write_atomic(path: &std::path::Path, bytes: &[u8]) -> Result<(), String> {
    let mut tmp_name = path
        .file_name()
        .ok_or("bad path")?
        .to_os_string();
    tmp_name.push(".sheetly-tmp");
    let tmp = path.with_file_name(tmp_name);
    if let Err(e) = std::fs::write(&tmp, bytes) {
        let _ = std::fs::remove_file(&tmp);
        return Err(e.to_string());
    }
    if let Err(e) = std::fs::rename(&tmp, path) {
        let _ = std::fs::remove_file(&tmp);
        return Err(e.to_string());
    }
    Ok(())
}

#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    read_limited(std::path::Path::new(&path)).map(Response::new)
}

#[tauri::command]
fn write_file(request: tauri::ipc::Request<'_>) -> Result<(), String> {
    let path = request
        .headers()
        .get("path")
        .and_then(|v| v.to_str().ok())
        .ok_or("missing path")?
        .to_string();
    let path = urldecode(&path);
    match request.body() {
        tauri::ipc::InvokeBody::Raw(bytes) => write_atomic(std::path::Path::new(&path), bytes),
        _ => Err("expected raw body".into()),
    }
}

#[tauri::command]
fn data_read(app: tauri::AppHandle, name: String) -> Result<Response, String> {
    let p = data_file(&app, &name)?;
    read_limited(&p).map(Response::new)
}

#[tauri::command]
fn data_write(app: tauri::AppHandle, request: tauri::ipc::Request<'_>) -> Result<(), String> {
    let name = request
        .headers()
        .get("name")
        .and_then(|v| v.to_str().ok())
        .ok_or("missing name")?
        .to_string();
    let p = data_file(&app, &name)?;
    match request.body() {
        tauri::ipc::InvokeBody::Raw(bytes) => write_atomic(&p, bytes),
        _ => Err("expected raw body".into()),
    }
}

#[tauri::command]
fn data_delete(app: tauri::AppHandle, name: String) -> Result<(), String> {
    let p = data_file(&app, &name)?;
    if p.exists() {
        std::fs::remove_file(p).map_err(|e| e.to_string())?;
    }
    Ok(())
}

#[tauri::command]
fn data_exists(app: tauri::AppHandle, name: String) -> Result<bool, String> {
    Ok(data_file(&app, &name)?.exists())
}

#[tauri::command]
fn path_exists(path: String) -> bool {
    std::path::Path::new(&path).is_file()
}

/// A spreadsheet path passed on the command line ("Open with Sheetly").
#[tauri::command]
fn cli_file() -> Option<String> {
    std::env::args().skip(1).find(|a| {
        let l = a.to_lowercase();
        l.ends_with(".xlsx") || l.ends_with(".csv") || l.ends_with(".sheetly")
    })
}

fn urldecode(s: &str) -> String {
    let b = s.as_bytes();
    let mut out = Vec::with_capacity(b.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'%' && i + 2 < b.len() {
            if let Ok(v) = u8::from_str_radix(&s[i + 1..i + 3], 16) {
                out.push(v);
                i += 3;
                continue;
            }
        }
        out.push(b[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            read_file, write_file, data_read, data_write, data_delete, data_exists, path_exists, cli_file
        ])
        .run(tauri::generate_context!())
        .expect("error while running Sheetly");
}
