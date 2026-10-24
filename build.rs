use std::fs;
use std::path::PathBuf;

fn emit_version(env_key: &str, version: &str) {
    println!("cargo:rustc-env={env_key}={version}");
}

fn parse_lock_version(lock: &str, name: &str) -> Option<String> {
    let needle = format!("name = \"{}\"", name);
    let mut in_pkg = false;
    let mut saw_name = false;
    for line in lock.lines() {
        let line = line.trim();
        if line == "[[package]]" {
            in_pkg = true;
            saw_name = false;
            continue;
        }
        if !in_pkg {
            continue;
        }
        if line.starts_with("name = ") {
            saw_name = line == needle;
        }
        if saw_name && line.starts_with("version = ") {
            if let Some(v) = line.split('"').nth(1) {
                return Some(v.to_string());
            }
        }
    }
    None
}

fn main() {
    pyo3_build_config::use_pyo3_cfgs();
    pyo3_build_config::add_extension_module_link_args();

    println!("cargo:rerun-if-changed=Cargo.toml");
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=Cargo.lock");

    let manifest_dir = std::env::var("CARGO_MANIFEST_DIR").ok();
    let Some(manifest_dir) = manifest_dir else {
        return;
    };
    let lock_path = PathBuf::from(manifest_dir).join("Cargo.lock");
    let Ok(lock) = fs::read_to_string(&lock_path) else {
        return;
    };

    if let Some(v) = parse_lock_version(&lock, "rust_xlsxwriter") {
        emit_version("WOLFXL_DEP_RUST_XLSXWRITER_VERSION", &v);
    }
}
