//! Spooled input is capped at a named per-format limit and fails with exit 2.

use std::fs::{self, File};
use std::path::{Path, PathBuf};

use assert_cmd::Command;

/// A directory under the system temp dir, removed on drop.
struct Scratch(PathBuf);

impl Scratch {
    fn new(name: &str) -> Self {
        let path =
            std::env::temp_dir().join(format!("wolfxl-input-limits-{}-{name}", std::process::id()));
        fs::create_dir(&path).expect("create scratch directory");
        Self(path)
    }
}

impl Drop for Scratch {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

/// Exit code and standard error of `wolfxl <args...> PATH --from format`.
fn run(args: &[&str], path: &Path, format: &str) -> (i32, String) {
    let out = Command::cargo_bin("wolfxl")
        .expect("wolfxl bin built")
        .args(args)
        .arg(path)
        .args(["--from", format])
        .output()
        .expect("run wolfxl");
    (
        out.status.code().expect("exit code"),
        String::from_utf8(out.stderr).expect("utf-8 stderr"),
    )
}

/// `wolfxl <args...>` over a sparse file one byte larger than `limit`, read
/// as `--from format` through a mismatched extension so it is spooled. The
/// file occupies no disk blocks and is refused before any byte is copied.
fn over_limit(args: &[&str], format: &str, limit: u64) -> (i32, String) {
    let scratch = Scratch::new(&format!("{}-{format}", args[0]));
    let path = scratch.0.join("oversized.bin");
    File::create(&path)
        .and_then(|file| file.set_len(limit + 1))
        .expect("create sparse file");
    run(args, &path, format)
}

const XLS_LIMIT: u64 = 512 * 1024 * 1024;
const CSV_LIMIT: u64 = 4 * 1024 * 1024 * 1024;

fn assert_limit_failure((code, stderr): (i32, String), message: &str) {
    assert_eq!(code, 2, "{stderr}");
    assert!(
        stderr.contains(&format!("oversized.bin exceeds the {message}")),
        "{stderr}"
    );
}

#[test]
fn oversized_xls_override_exits_two_naming_the_limit() {
    assert_limit_failure(
        over_limit(&["calc", "=1"], "xls", XLS_LIMIT),
        "512 MiB spool limit for --from xls input",
    );
    assert_limit_failure(
        over_limit(&["peek"], "xls", XLS_LIMIT),
        "512 MiB spool limit for --from xls input",
    );
}

#[test]
fn oversized_csv_override_exits_two_naming_the_limit() {
    assert_limit_failure(
        over_limit(&["calc", "=1"], "csv", CSV_LIMIT),
        "4 GiB spool limit for --from csv input",
    );
    assert_limit_failure(
        over_limit(&["peek"], "csv", CSV_LIMIT),
        "4 GiB spool limit for --from csv input",
    );
}

#[test]
fn peek_keeps_exit_one_for_other_input_failures() {
    let scratch = Scratch::new("peek-missing");
    let (code, stderr) = run(&["peek"], &scratch.0.join("missing.bin"), "csv");
    assert_eq!(code, 1, "{stderr}");
    assert!(stderr.contains("missing.bin"), "{stderr}");
}
