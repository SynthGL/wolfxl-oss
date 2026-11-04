//! `wolfxl peek --interactive` refusals. The test harness captures standard
//! output, so it is never a terminal here: the grid must refuse with exit 2
//! and draw nothing. Grid behavior itself is unit-tested in `src/tui`.

use std::path::{Path, PathBuf};

use assert_cmd::Command;

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

struct Run {
    code: i32,
    stdout: String,
    stderr: String,
}

fn peek(args: &[&str], stdin: Option<&str>) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.arg("peek").args(args);
    if let Some(stdin) = stdin {
        command.write_stdin(stdin);
    }
    let out = command.output().expect("run wolfxl peek");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

#[test]
fn interactive_without_a_terminal_exits_2() {
    let file = fixture("sample-minimal.csv");
    let file = file.to_str().expect("utf-8 fixture path");
    for flag in ["--interactive", "-i"] {
        let run = peek(&[file, flag], None);
        assert_eq!(run.code, 2, "{flag}: stderr={}", run.stderr);
        assert_eq!(run.stdout, "", "{flag} draws nothing");
        assert!(
            run.stderr
                .contains("error: --interactive needs a terminal on standard output"),
            "{flag}: stderr={}",
            run.stderr
        );
    }
}

#[test]
fn interactive_with_standard_input_data_exits_2() {
    for args in [
        &["-", "--from", "csv", "--interactive"][..],
        &["-", "--interactive"][..],
    ] {
        let run = peek(args, Some("A,B\n1,2\n"));
        assert_eq!(run.code, 2, "{args:?}: stderr={}", run.stderr);
        assert_eq!(run.stdout, "");
        assert!(
            run.stderr
                .contains("--interactive reads keys from standard input"),
            "{args:?}: stderr={}",
            run.stderr
        );
    }
}

#[test]
fn interactive_rejects_an_export_format() {
    let file = fixture("sample-minimal.csv");
    let file = file.to_str().expect("utf-8 fixture path");
    let run = peek(&[file, "--interactive", "-e", "csv"], None);
    assert_eq!(run.code, 2, "stderr={}", run.stderr);
    assert_eq!(run.stdout, "");
    assert!(run.stderr.contains("cannot be used with"), "{}", run.stderr);
}

#[test]
fn interactive_refusal_precedes_filter_validation() {
    let file = fixture("sample-minimal.csv");
    let file = file.to_str().expect("utf-8 fixture path");
    let run = peek(&[file, "-i", "--where", "NoSuchColumn=1"], None);
    assert_eq!(run.code, 2, "stderr={}", run.stderr);
    assert!(
        run.stderr.contains("needs a terminal"),
        "stderr={}",
        run.stderr
    );
}
