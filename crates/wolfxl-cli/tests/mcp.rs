//! Integration tests for `wolfxl mcp`: full newline-delimited JSON-RPC
//! sessions against the built binary covering initialize, tools/list, every
//! read and edit tool, dry runs, revision guards, `--read-only`, `--root`
//! confinement, undo/redo history, and protocol errors.

use std::collections::BTreeMap;
use std::fs;
use std::io::{BufRead, BufReader, Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, Command as Process, ExitStatus, Stdio};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::mpsc::{self, Receiver};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use assert_cmd::Command;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

/// How long one response may take before the test fails instead of hanging.
const RESPONSE_TIMEOUT: Duration = Duration::from_secs(60);

const READ_TOOLS: [&str; 7] = [
    "calc",
    "diff",
    "fingerprint",
    "get_workbook",
    "pivot",
    "read_range",
    "stats",
];

const MUTATING_TOOLS: [&str; 10] = [
    "add_sheet",
    "delete_columns",
    "delete_rows",
    "insert_columns",
    "insert_rows",
    "redo",
    "rename_sheet",
    "set_number_format",
    "undo",
    "write_cells",
];

static NEXT_TEMP: AtomicUsize = AtomicUsize::new(0);

/// A private directory, unique per test, removed on drop. The path is
/// canonical so it compares equal to the server's resolved paths.
struct TempDir(PathBuf);

impl TempDir {
    fn new(label: &str) -> Self {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|elapsed| elapsed.as_nanos())
            .unwrap_or_default();
        let path = std::env::temp_dir().join(format!(
            "wolfxl-mcp-{label}-{}-{}-{stamp}",
            std::process::id(),
            NEXT_TEMP.fetch_add(1, Ordering::Relaxed)
        ));
        fs::create_dir(&path).expect("create temp dir");
        Self(path.canonicalize().expect("canonical temp dir"))
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }

    fn root(&self) -> &str {
        self.0.to_str().expect("utf-8 temp dir")
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn text(path: &Path) -> &str {
    path.to_str().expect("utf-8 path")
}

/// The revision the server reports: lowercase hex SHA-256 of the file bytes.
fn revision_of(bytes: &[u8]) -> String {
    base16ct::lower::encode_string(&Sha256::digest(bytes))
}

fn revision(path: &Path) -> String {
    revision_of(&fs::read(path).expect("read workbook"))
}

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// A two-sheet XLSX package. `Data` holds a shared-string header row
/// (Item, Region, Amount, Flag) over three rows whose middle amount is
/// `gadget`; `Summary` holds an inline-string label and a cross-sheet SUM
/// with its cached total. The workbook defines `Amounts` over the amount
/// column and carries an extended-properties part no edit touches.
fn workbook_bytes(gadget: f64) -> Vec<u8> {
    let total = 175.0 + gadget;
    let types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>"#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>"#;
    let app = r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>wolfxl test fixture</Application></Properties>"#;
    let workbook = format!(
        r#"<workbook xmlns="{MAIN_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Summary" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="Amounts">Data!$C$2:$C$4</definedName></definedNames></workbook>"#
    );
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"#;
    let styles = format!(
        r#"<styleSheet xmlns="{MAIN_NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#
    );
    let strings = [
        "Item", "Region", "Amount", "Flag", "Widget", "West", "Gadget", "East", "Gizmo",
    ]
    .map(|text| format!("<si><t>{text}</t></si>"))
    .concat();
    let shared = format!(r#"<sst xmlns="{MAIN_NS}" count="10" uniqueCount="9">{strings}</sst>"#);
    let data = format!(
        r#"<worksheet xmlns="{MAIN_NS}"><dimension ref="A1:D4"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row><row r="2"><c r="A2" t="s"><v>4</v></c><c r="B2" t="s"><v>5</v></c><c r="C2"><v>100</v></c><c r="D2" t="b"><v>1</v></c></row><row r="3"><c r="A3" t="s"><v>6</v></c><c r="B3" t="s"><v>7</v></c><c r="C3"><v>{gadget}</v></c><c r="D3" t="b"><v>0</v></c></row><row r="4"><c r="A4" t="s"><v>8</v></c><c r="B4" t="s"><v>5</v></c><c r="C4"><v>75</v></c><c r="D4" t="b"><v>1</v></c></row></sheetData></worksheet>"#
    );
    let summary = format!(
        r#"<worksheet xmlns="{MAIN_NS}"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Total</t></is></c><c r="B1"><f>SUM(Data!C2:C4)</f><v>{total}</v></c></row></sheetData></worksheet>"#
    );
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, body) in [
        ("[Content_Types].xml", types),
        ("_rels/.rels", root),
        ("docProps/app.xml", app),
        ("xl/workbook.xml", workbook.as_str()),
        ("xl/_rels/workbook.xml.rels", rels),
        ("xl/styles.xml", styles.as_str()),
        ("xl/sharedStrings.xml", shared.as_str()),
        ("xl/worksheets/sheet1.xml", data.as_str()),
        ("xl/worksheets/sheet2.xml", summary.as_str()),
    ] {
        writer.start_file(name, options).expect("start member");
        writer.write_all(body.as_bytes()).expect("write member");
    }
    writer.finish().expect("finish package").into_inner()
}

/// The fixture workbook (Gadget amount 250.5) written as `name` in `directory`.
fn workbook(directory: &TempDir, name: &str) -> PathBuf {
    let path = directory.path(name);
    fs::write(&path, workbook_bytes(250.5)).expect("write workbook");
    path
}

/// Every member of a ZIP package by name, decompressed.
fn members(path: &Path) -> BTreeMap<String, Vec<u8>> {
    let mut archive = ZipArchive::new(fs::File::open(path).expect("open package")).expect("zip");
    (0..archive.len())
        .map(|index| {
            let mut member = archive.by_index(index).expect("package member");
            let mut bytes = Vec::new();
            member.read_to_end(&mut bytes).expect("read member");
            (member.name().expect("member name").to_string(), bytes)
        })
        .collect()
}

/// `value` with every number as an `f64`, so `100` and `100.0` compare equal.
fn floats(value: &Value) -> Value {
    match value {
        Value::Number(number) => json!(number.as_f64().expect("finite number")),
        Value::Array(items) => Value::Array(items.iter().map(floats).collect()),
        Value::Object(fields) => Value::Object(
            fields
                .iter()
                .map(|(key, field)| (key.clone(), floats(field)))
                .collect(),
        ),
        other => other.clone(),
    }
}

/// Assert two JSON values are equal, comparing numbers by value.
fn assert_cells(actual: &Value, expected: Value) {
    assert_eq!(floats(actual), floats(&expected), "actual: {actual}");
}

/// A `rows` by `cols` matrix of nulls, the shape of an empty range.
fn blanks(rows: usize, cols: usize) -> Value {
    json!(vec![vec![Value::Null; cols]; rows])
}

/// `arguments` with extra fields merged in.
fn with(arguments: &Value, extra: Value) -> Value {
    let mut merged = arguments.clone();
    for (key, field) in extra.as_object().expect("extra fields object") {
        merged[key] = field.clone();
    }
    merged
}

/// Assert `response` is a JSON-RPC error with `code`; return its message.
fn rpc_error(response: &Value, code: i64) -> String {
    assert!(
        response.get("result").is_none(),
        "expected an error: {response}"
    );
    assert_eq!(response["error"]["code"], code, "error code: {response}");
    response["error"]["message"]
        .as_str()
        .expect("error message")
        .to_string()
}

/// JSON a CLI command prints, accepting the 1 exit code `diff` uses for
/// differences.
fn cli_json(args: &[&str]) -> Value {
    let out = Command::cargo_bin("wolfxl")
        .expect("wolfxl bin built")
        .args(args)
        .output()
        .expect("run wolfxl");
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(
        matches!(out.status.code(), Some(0 | 1)),
        "wolfxl {args:?}: stderr={stderr}"
    );
    serde_json::from_slice(&out.stdout)
        .unwrap_or_else(|error| panic!("wolfxl {args:?} JSON ({error}): stderr={stderr}"))
}

fn cli_text(args: &[&str]) -> String {
    let out = Command::cargo_bin("wolfxl")
        .expect("wolfxl bin built")
        .args(args)
        .output()
        .expect("run wolfxl");
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert_eq!(
        out.status.code(),
        Some(0),
        "wolfxl {args:?}: stderr={stderr}"
    );
    String::from_utf8(out.stdout).expect("utf-8 stdout")
}

/// One `wolfxl mcp` process driven over its standard input and output.
struct Session {
    child: Child,
    stdin: Option<ChildStdin>,
    lines: Receiver<String>,
    next_id: u64,
}

impl Session {
    /// Start the server without initializing it.
    fn spawn(args: &[&str]) -> Self {
        let mut child = Process::new(env!("CARGO_BIN_EXE_wolfxl"))
            .arg("mcp")
            .args(args)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::inherit())
            .spawn()
            .expect("start wolfxl mcp");
        let stdin = child.stdin.take();
        let stdout = child.stdout.take().expect("server stdout");
        let (sender, lines) = mpsc::channel();
        thread::spawn(move || {
            for line in BufReader::new(stdout).lines() {
                let Ok(line) = line else { break };
                if sender.send(line).is_err() {
                    break;
                }
            }
        });
        Self {
            child,
            stdin,
            lines,
            next_id: 0,
        }
    }

    /// Start the server and complete the initialize handshake.
    fn start(args: &[&str]) -> Self {
        let mut session = Self::spawn(args);
        let response = session.initialize("2025-06-18");
        assert_eq!(response["result"]["protocolVersion"], "2025-06-18");
        session.notify("notifications/initialized", json!({}));
        session
    }

    fn initialize(&mut self, version: &str) -> Value {
        self.request(
            "initialize",
            json!({
                "protocolVersion": version,
                "capabilities": {},
                "clientInfo": {"name": "wolfxl-tests", "version": "1.0.0"}
            }),
        )
    }

    fn send_line(&mut self, line: &str) {
        let stdin = self.stdin.as_mut().expect("server stdin open");
        stdin.write_all(line.as_bytes()).expect("write request");
        stdin.write_all(b"\n").expect("write newline");
        stdin.flush().expect("flush request");
    }

    fn receive(&mut self) -> Value {
        let line = self
            .lines
            .recv_timeout(RESPONSE_TIMEOUT)
            .expect("server answers within the timeout");
        serde_json::from_str(&line)
            .unwrap_or_else(|error| panic!("response is one JSON line ({error}): {line}"))
    }

    /// Send a request and return the response, which must answer its id.
    fn request(&mut self, method: &str, params: Value) -> Value {
        self.next_id += 1;
        let id = self.next_id;
        let request = json!({"jsonrpc": "2.0", "id": id, "method": method, "params": params});
        self.send_line(&request.to_string());
        let response = self.receive();
        assert_eq!(response["jsonrpc"], "2.0", "{response}");
        assert_eq!(
            response["id"], id,
            "response answers request {id}: {response}"
        );
        response
    }

    fn notify(&mut self, method: &str, params: Value) {
        let notification = json!({"jsonrpc": "2.0", "method": method, "params": params});
        self.send_line(&notification.to_string());
    }

    fn call(&mut self, name: &str, arguments: Value) -> Value {
        self.request("tools/call", json!({"name": name, "arguments": arguments}))
    }

    /// Call a tool that must succeed; return its structured content after
    /// checking the text block carries the same JSON.
    fn ok(&mut self, name: &str, arguments: Value) -> Value {
        let response = self.call(name, arguments.clone());
        let result = &response["result"];
        assert_eq!(
            result["isError"], false,
            "{name} {arguments} failed: {response}"
        );
        assert_text_matches(result);
        result["structuredContent"].clone()
    }

    /// Call a tool that must fail as a tool result; return its error text.
    fn fail(&mut self, name: &str, arguments: Value) -> String {
        let response = self.call(name, arguments.clone());
        let result = &response["result"];
        assert_eq!(
            result["isError"], true,
            "{name} {arguments} should fail: {response}"
        );
        assert_text_matches(result);
        result["structuredContent"]["error"]
            .as_str()
            .unwrap_or_else(|| panic!("error text: {response}"))
            .to_string()
    }

    /// Call a tool whose name or arguments the server must reject as
    /// invalid params; return the JSON-RPC error message.
    fn invalid(&mut self, name: &str, arguments: Value) -> String {
        let response = self.call(name, arguments);
        rpc_error(&response, -32602)
    }

    fn tool_names(&mut self) -> Vec<String> {
        let response = self.request("tools/list", json!({}));
        let mut names: Vec<String> = response["result"]["tools"]
            .as_array()
            .unwrap_or_else(|| panic!("tools array: {response}"))
            .iter()
            .map(|tool| tool["name"].as_str().expect("tool name").to_string())
            .collect();
        names.sort();
        names
    }

    fn sheet_names(&mut self, path: &Path) -> Vec<String> {
        self.ok("get_workbook", json!({"path": text(path)}))["sheets"]
            .as_array()
            .expect("sheets array")
            .iter()
            .map(|sheet| sheet["name"].as_str().expect("sheet name").to_string())
            .collect()
    }

    fn values(&mut self, path: &Path, sheet: &str, range: &str) -> Value {
        self.ok(
            "read_range",
            json!({"path": text(path), "sheet": sheet, "range": range}),
        )["values"]
            .clone()
    }

    /// Close standard input and wait for the server to exit.
    fn finish(&mut self) -> ExitStatus {
        drop(self.stdin.take());
        let deadline = Instant::now() + RESPONSE_TIMEOUT;
        loop {
            if let Some(status) = self.child.try_wait().expect("poll server") {
                return status;
            }
            assert!(Instant::now() < deadline, "server exits after end of input");
            thread::sleep(Duration::from_millis(20));
        }
    }
}

impl Drop for Session {
    fn drop(&mut self) {
        drop(self.stdin.take());
        if matches!(self.child.try_wait(), Ok(None)) {
            let _ = self.child.kill();
        }
        let _ = self.child.wait();
    }
}

/// A successful tools/call result carries its structured content again as
/// one text block of identical JSON.
fn assert_text_matches(result: &Value) {
    let content = result["content"].as_array().expect("content array");
    assert_eq!(content.len(), 1, "{result}");
    assert_eq!(content[0]["type"], "text");
    let echoed: Value =
        serde_json::from_str(content[0]["text"].as_str().expect("text block")).expect("text JSON");
    assert_eq!(echoed, result["structuredContent"]);
}

/// Assert a successful mutation result wrote the file it reports.
fn assert_written(result: &Value, path: &Path) -> String {
    assert_eq!(result["dry_run"], false, "{result}");
    assert!(result["change"].is_object(), "change: {result}");
    let revision = result["revision"].as_str().expect("revision").to_string();
    assert_eq!(
        revision,
        self::revision(path),
        "revision matches the file bytes"
    );
    revision
}

#[test]
fn initialize_negotiates_protocol_versions() {
    for (requested, negotiated) in [
        ("2025-06-18", "2025-06-18"),
        ("2025-03-26", "2025-03-26"),
        ("2024-11-05", "2024-11-05"),
        ("1999-01-01", "2025-06-18"),
    ] {
        let mut session = Session::spawn(&[]);
        let response = session.initialize(requested);
        let result = &response["result"];
        assert_eq!(
            result["protocolVersion"], negotiated,
            "{requested}: {response}"
        );
        assert!(result["capabilities"]["tools"].is_object(), "{response}");
        assert_eq!(result["serverInfo"]["name"], "wolfxl");
        assert!(result["serverInfo"]["version"].is_string(), "{response}");
    }
}

#[test]
fn tools_list_describes_every_tool_with_an_object_schema() {
    let mut session = Session::start(&[]);
    let response = session.request("tools/list", json!({}));
    let tools = response["result"]["tools"].as_array().expect("tools array");
    let mut names: Vec<&str> = tools
        .iter()
        .map(|tool| tool["name"].as_str().expect("tool name"))
        .collect();
    names.sort_unstable();
    let mut expected: Vec<&str> = READ_TOOLS
        .iter()
        .chain(MUTATING_TOOLS.iter())
        .copied()
        .collect();
    expected.sort_unstable();
    assert_eq!(names, expected);
    for tool in tools {
        let name = tool["name"].as_str().unwrap();
        assert!(
            !tool["description"].as_str().unwrap_or_default().is_empty(),
            "{name} has a description"
        );
        let schema = &tool["inputSchema"];
        assert_eq!(schema["type"], "object", "{name}");
        assert_eq!(schema["properties"]["path"]["type"], "string", "{name}");
        let required = schema["required"].as_array().expect("required list");
        assert!(required.contains(&json!("path")), "{name} requires path");
        let mutating = MUTATING_TOOLS.contains(&name);
        for option in ["dry_run", "expected_revision"] {
            assert_eq!(
                schema["properties"].get(option).is_some(),
                mutating,
                "{name} {option}"
            );
        }
    }
    assert!(session.finish().success(), "clean exit at end of input");
}

#[test]
fn read_only_hides_and_refuses_mutating_tools() {
    let directory = TempDir::new("read-only");
    let path = workbook(&directory, "book.xlsx");
    let before = fs::read(&path).unwrap();
    let mut session = Session::start(&["--read-only"]);
    assert_eq!(session.tool_names(), READ_TOOLS);
    for (name, arguments) in [
        (
            "write_cells",
            json!({"path": text(&path), "sheet": "Data", "cells": [{"cell": "A1", "value": 1}]}),
        ),
        ("add_sheet", json!({"path": text(&path), "name": "New"})),
        ("undo", json!({"path": text(&path)})),
    ] {
        let message = session.invalid(name, arguments);
        assert!(message.contains(name), "{message}");
    }
    assert_eq!(fs::read(&path).unwrap(), before, "read-only never writes");
    assert_cells(
        &session.values(&path, "Data", "A1:B1"),
        json!([["Item", "Region"]]),
    );
}

#[test]
fn get_workbook_matches_map_and_reports_the_byte_revision() {
    let directory = TempDir::new("get-workbook");
    let path = workbook(&directory, "book.xlsx");
    let financials = fixture("sample-financials.xlsx");
    let mut session = Session::start(&[]);
    for path in [&path, &financials] {
        let document = session.ok("get_workbook", json!({"path": text(path)}));
        let map = cli_json(&["map", text(path)]);
        assert_eq!(document["format"], "xlsx");
        assert_eq!(document["revision"], revision(path));
        assert_eq!(document["sheets"], map["sheets"], "{}", path.display());
        if let Some(named) = document.get("named_ranges") {
            assert_eq!(named, &map["named_ranges"]);
        }
    }
    let document = session.ok("get_workbook", json!({"path": text(&path)}));
    assert_eq!(document["path"], text(&path));
    assert_eq!(
        document["sheets"][0]["headers"],
        json!(["Item", "Region", "Amount", "Flag"])
    );
    assert_eq!(document["sheets"][1]["name"], "Summary");
}

#[test]
fn read_range_returns_values_and_formula_text_for_the_requested_rectangle() {
    let directory = TempDir::new("read-range");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);

    let document = session.ok("read_range", json!({"path": text(&path), "range": "A1:D4"}));
    assert_eq!(document["path"], text(&path));
    assert_eq!(document["sheet"], "Data", "the first sheet is the default");
    assert_eq!(document["range"], "A1:D4");
    assert_eq!(document["revision"], revision(&path));
    assert_eq!(document["cell_cap"], 10000);
    assert_cells(
        &document["values"],
        json!([
            ["Item", "Region", "Amount", "Flag"],
            ["Widget", "West", 100, true],
            ["Gadget", "East", 250.5, false],
            ["Gizmo", "West", 75, true]
        ]),
    );
    assert_eq!(document["formulas"], blanks(4, 4));

    let summary = session.ok(
        "read_range",
        json!({"path": text(&path), "sheet": "Summary", "range": "A1:B1"}),
    );
    assert_cells(&summary["values"], json!([["Total", 425.5]]));
    assert_eq!(summary["formulas"], json!([[null, "SUM(Data!C2:C4)"]]));

    // Cells beyond the stored data read as blanks; one cell is a 1x1 range.
    assert_cells(
        &session.values(&path, "Data", "C4:E5"),
        json!([[75, true, null], [null, null, null]]),
    );
    assert_cells(&session.values(&path, "Data", "B3"), json!([["East"]]));
}

#[test]
fn read_range_enforces_the_cell_cap_and_names_bad_requests() {
    let directory = TempDir::new("read-cap");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);

    let full = session.values(&path, "Data", "A1:CV100");
    let rows = full.as_array().expect("rows");
    assert_eq!(rows.len(), 100);
    assert!(rows
        .iter()
        .all(|row| row.as_array().map(Vec::len) == Some(100)));

    let over = session.fail(
        "read_range",
        json!({"path": text(&path), "sheet": "Data", "range": "A1:CV101"}),
    );
    assert!(over.contains("10100") && over.contains("10000"), "{over}");

    let sheet = session.fail(
        "read_range",
        json!({"path": text(&path), "sheet": "Nope", "range": "A1"}),
    );
    assert!(sheet.contains("Nope"), "{sheet}");
    session.fail(
        "read_range",
        json!({"path": text(&path), "sheet": "Data", "range": "B2:"}),
    );
    let missing = directory.path("missing.xlsx");
    session.fail("get_workbook", json!({"path": text(&missing)}));
}

#[test]
fn delimited_files_are_readable_but_not_editable() {
    let directory = TempDir::new("csv");
    let path = directory.path("ledger.csv");
    fs::copy(fixture("sample-minimal.csv"), &path).expect("copy csv");
    let before = fs::read(&path).unwrap();
    let mut session = Session::start(&[]);

    let document = session.ok("get_workbook", json!({"path": text(&path)}));
    assert_eq!(document["format"], "csv");
    assert_eq!(document["sheets"][0]["name"], "ledger");
    let read = session.ok("read_range", json!({"path": text(&path), "range": "A1:E2"}));
    assert_eq!(
        read["values"][0],
        json!(["Account", "Jan", "Feb", "Mar", "Q1"])
    );
    assert_eq!(read["values"][1][0], "Revenue");
    assert_eq!(read["formulas"], blanks(2, 5));

    let message = session.fail(
        "write_cells",
        json!({"path": text(&path), "sheet": "ledger", "cells": [{"cell": "A1", "value": "x"}]}),
    );
    assert!(message.contains(".xlsx"), "{message}");
    session.fail("undo", json!({"path": text(&path)}));
    assert_eq!(fs::read(&path).unwrap(), before);
}

#[test]
fn analysis_tools_return_the_cli_json_documents() {
    let directory = TempDir::new("analysis");
    let left = workbook(&directory, "left.xlsx");
    let right = directory.path("right.xlsx");
    fs::write(&right, workbook_bytes(251.0)).unwrap();
    let (left_text, right_text) = (text(&left), text(&right));
    let mut session = Session::start(&[]);

    let calc = session.ok("calc", json!({"path": left_text, "formula": "=SUM(C2:C4)"}));
    assert_eq!(
        calc,
        cli_json(&["calc", "=SUM(C2:C4)", left_text, "--json"])
    );
    assert_cells(&calc["value"], json!(425.5));
    let calc = session.ok(
        "calc",
        json!({"path": left_text, "sheet": "Summary", "formula": "B1*2"}),
    );
    assert_eq!(
        calc,
        cli_json(&["calc", "B1*2", left_text, "--sheet", "Summary", "--json"])
    );
    assert_cells(&calc["value"], json!(851));

    let stats = session.ok("stats", json!({"path": left_text}));
    assert_eq!(stats, cli_json(&["stats", left_text, "-e", "json"]));
    let stats = session.ok(
        "stats",
        json!({"path": left_text, "filters": ["Region=West"], "select": ["Amount", "Item"], "top": 1}),
    );
    assert_eq!(
        stats,
        cli_json(&[
            "stats",
            left_text,
            "--where",
            "Region=West",
            "--select",
            "Amount,Item",
            "--top",
            "1",
            "-e",
            "json",
        ])
    );
    assert_eq!(stats["rows"], 2);

    let pivot = session.ok(
        "pivot",
        json!({"path": left_text, "rows": ["Region"], "values": ["sum:Amount", "count:Amount"]}),
    );
    assert_eq!(
        pivot,
        cli_json(&[
            "pivot",
            left_text,
            "--rows",
            "Region",
            "--values",
            "sum:Amount,count:Amount",
            "-e",
            "json",
        ])
    );
    assert_cells(&pivot["totals"], json!([425.5, 3]));
    let pivot = session.ok(
        "pivot",
        json!({"path": left_text, "rows": ["Region"], "column": "Flag", "values": ["sum:Amount"]}),
    );
    assert_eq!(
        pivot,
        cli_json(&[
            "pivot",
            left_text,
            "--rows",
            "Region",
            "--column",
            "Flag",
            "--values",
            "sum:Amount",
            "-e",
            "json",
        ])
    );

    let diff = session.ok(
        "diff",
        json!({
            "path": left_text, "right_path": right_text, "key": "Item", "compare": ["Amount"],
            "tolerance": 0.25, "key_transform": "lower"
        }),
    );
    assert_eq!(
        diff,
        cli_json(&[
            "diff",
            left_text,
            right_text,
            "--key",
            "Item",
            "--compare",
            "Amount",
            "--tolerance",
            "0.25",
            "--key-transform",
            "lower",
        ])
    );
    assert_eq!(diff["key_transform"], "lower");
    assert_eq!(
        diff["summary"],
        json!({"matched": 3, "only_left": 0, "only_right": 0, "diff": 1, "diff_outside_tolerance": 1})
    );

    let fingerprint = session.ok("fingerprint", json!({"path": left_text}));
    assert_eq!(fingerprint, cli_json(&["fingerprint", left_text, "--json"]));
    let financials = fixture("sample-financials.xlsx");
    assert_eq!(
        session.ok("fingerprint", json!({"path": text(&financials)})),
        cli_json(&["fingerprint", text(&financials), "--json"])
    );

    let unknown = session.fail(
        "pivot",
        json!({"path": left_text, "rows": ["Nope"], "values": ["sum:Amount"]}),
    );
    assert!(unknown.contains("Nope"), "{unknown}");
}

#[test]
fn write_cells_writes_every_value_kind_and_reads_back() {
    let directory = TempDir::new("write");
    let path = workbook(&directory, "book.xlsx");
    let original = members(&path);
    let mut session = Session::start(&[]);
    let start = session.ok("get_workbook", json!({"path": text(&path)}))["revision"].clone();

    let result = session.ok(
        "write_cells",
        json!({
            "path": text(&path),
            "sheet": "Data",
            "expected_revision": start,
            "cells": [
                {"cell": "B2", "value": "North"},
                {"cell": "C2", "value": 120.25},
                {"cell": "D2", "value": false},
                {"cell": "C3", "value": null},
                {"cell": "E1", "value": "Double"},
                {"cell": "E2", "formula": "C2*2"},
                {"cell": "E3", "formula": "=C4*2"},
                {"cell": "F1", "value": "=not a formula"}
            ]
        }),
    );
    let written = assert_written(&result, &path);
    assert_ne!(json!(written), start);

    let document = session.ok(
        "read_range",
        json!({"path": text(&path), "sheet": "Data", "range": "A1:F4"}),
    );
    assert_eq!(document["revision"], written);
    assert_cells(
        &document["values"],
        json!([
            [
                "Item",
                "Region",
                "Amount",
                "Flag",
                "Double",
                "=not a formula"
            ],
            ["Widget", "North", 120.25, false, null, null],
            ["Gadget", "East", null, false, null, null],
            ["Gizmo", "West", 75, true, null, null]
        ]),
    );
    assert_eq!(
        document["formulas"],
        json!([
            [null, null, null, null, null, null],
            [null, null, null, null, "C2*2", null],
            [null, null, null, null, "C4*2", null],
            [null, null, null, null, null, null]
        ])
    );

    // Stored values feed calc at once; new formulas carry no cached result,
    // so the workbook asks Excel to recalculate on open.
    assert_cells(
        &session.ok(
            "calc",
            json!({"path": text(&path), "formula": "SUM(C2:C4)"}),
        )["value"],
        json!(195.25),
    );
    let after = members(&path);
    let workbook_xml = String::from_utf8(after["xl/workbook.xml"].clone()).unwrap();
    assert!(
        workbook_xml.contains(r#"fullCalcOnLoad="1""#),
        "{workbook_xml}"
    );
    for part in ["docProps/app.xml", "xl/worksheets/sheet2.xml"] {
        assert_eq!(after[part], original[part], "{part} is untouched");
    }

    // The cross-sheet formula still reads its own cached value.
    assert_cells(&session.values(&path, "Summary", "B1"), json!([[425.5]]));
}

#[test]
fn number_only_edit_preserves_other_parts_and_requests_recalculation() {
    let directory = TempDir::new("preserve");
    let path = workbook(&directory, "book.xlsx");
    let original = members(&path);
    let mut session = Session::start(&[]);
    let result = session.ok(
        "write_cells",
        json!({"path": text(&path), "sheet": "Data", "cells": [{"cell": "C4", "value": 80}]}),
    );
    assert_written(&result, &path);
    let after = members(&path);
    assert_eq!(
        after.keys().collect::<Vec<_>>(),
        original.keys().collect::<Vec<_>>(),
        "same package members"
    );
    for (name, bytes) in &original {
        if name != "xl/worksheets/sheet1.xml" && name != "xl/workbook.xml" {
            assert_eq!(&after[name], bytes, "{name} is untouched");
        }
    }
    // Summary!B1 caches SUM(Data!C2:C4), which this edit makes stale.
    let workbook_xml = String::from_utf8(after["xl/workbook.xml"].clone()).unwrap();
    assert!(
        workbook_xml.contains(r#"fullCalcOnLoad="1""#),
        "{workbook_xml}"
    );
    assert_cells(&session.values(&path, "Data", "C4"), json!([[80]]));
}

#[test]
fn set_number_format_changes_display_but_not_stored_values() {
    let directory = TempDir::new("format");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);
    let result = session.ok(
        "set_number_format",
        json!({"path": text(&path), "sheet": "Data", "range": "C2:C4", "format_code": "$#,##0.00"}),
    );
    assert_written(&result, &path);
    let csv = cli_text(&["peek", text(&path), "-e", "csv"]);
    for shown in ["$100.00", "$250.50", "$75.00"] {
        assert!(csv.contains(shown), "{shown} in {csv}");
    }
    assert_cells(
        &session.values(&path, "Data", "C2:C4"),
        json!([[100], [250.5], [75]]),
    );
    session.fail(
        "set_number_format",
        json!({"path": text(&path), "sheet": "Nope", "range": "A1", "format_code": "0.0%"}),
    );
}

#[test]
fn structural_edits_shift_rows_and_columns() {
    let directory = TempDir::new("structure");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);
    let span = |at: u32, count: u32| json!({"path": text(&path), "sheet": "Data", "at": at, "count": count});

    assert_written(&session.ok("insert_rows", span(2, 1)), &path);
    assert_cells(
        &session.values(&path, "Data", "A1:A5"),
        json!([["Item"], [null], ["Widget"], ["Gadget"], ["Gizmo"]]),
    );
    assert_written(&session.ok("delete_rows", span(2, 1)), &path);
    assert_cells(
        &session.values(&path, "Data", "A1:A4"),
        json!([["Item"], ["Widget"], ["Gadget"], ["Gizmo"]]),
    );

    assert_written(&session.ok("insert_columns", span(2, 2)), &path);
    assert_cells(
        &session.values(&path, "Data", "A1:F2"),
        json!([
            ["Item", null, null, "Region", "Amount", "Flag"],
            ["Widget", null, null, "West", 100, true]
        ]),
    );
    assert_written(&session.ok("delete_columns", span(2, 2)), &path);
    assert_cells(
        &session.values(&path, "Data", "A1:D2"),
        json!([
            ["Item", "Region", "Amount", "Flag"],
            ["Widget", "West", 100, true]
        ]),
    );

    // Indices are 1-based and counts positive; sheet is required for edits.
    let before = fs::read(&path).unwrap();
    session.invalid("insert_rows", span(0, 1));
    session.invalid("delete_columns", span(1, 0));
    session.invalid(
        "insert_rows",
        json!({"path": text(&path), "at": 1, "count": 1}),
    );
    session.fail(
        "delete_rows",
        json!({"path": text(&path), "sheet": "Nope", "at": 1, "count": 1}),
    );
    assert_eq!(fs::read(&path).unwrap(), before);
}

#[test]
fn add_and_rename_sheets_update_names_and_references() {
    let directory = TempDir::new("sheets");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);

    assert_written(
        &session.ok("add_sheet", json!({"path": text(&path), "name": "Notes"})),
        &path,
    );
    assert_eq!(session.sheet_names(&path), ["Data", "Summary", "Notes"]);
    assert_written(
        &session.ok(
            "add_sheet",
            json!({"path": text(&path), "name": "Cover", "position": 0}),
        ),
        &path,
    );
    assert_eq!(
        session.sheet_names(&path),
        ["Cover", "Data", "Summary", "Notes"]
    );
    assert_cells(
        &session.values(&path, "Cover", "A1:B2"),
        json!([[null, null], [null, null]]),
    );

    assert_written(
        &session.ok(
            "rename_sheet",
            json!({"path": text(&path), "old": "Data", "new": "Ledger"}),
        ),
        &path,
    );
    assert_eq!(
        session.sheet_names(&path),
        ["Cover", "Ledger", "Summary", "Notes"]
    );
    let summary = session.ok(
        "read_range",
        json!({"path": text(&path), "sheet": "Summary", "range": "B1"}),
    );
    assert_eq!(summary["formulas"], json!([["SUM(Ledger!C2:C4)"]]));
    let names = session.ok("get_workbook", json!({"path": text(&path)}))["named_ranges"].clone();
    let amounts = names
        .as_array()
        .expect("named ranges")
        .iter()
        .find(|entry| entry["name"] == "Amounts")
        .expect("Amounts survives the rename")["formula"]
        .as_str()
        .expect("formula text")
        .to_string();
    assert!(
        amounts.contains("Ledger!") && !amounts.contains("Data!"),
        "{amounts}"
    );
    assert_cells(
        &session.values(&path, "Ledger", "A2:C2"),
        json!([["Widget", "West", 100]]),
    );

    let before = fs::read(&path).unwrap();
    session.fail("add_sheet", json!({"path": text(&path), "name": "Notes"}));
    session.fail(
        "rename_sheet",
        json!({"path": text(&path), "old": "Data", "new": "Other"}),
    );
    session.fail(
        "rename_sheet",
        json!({"path": text(&path), "old": "Ledger", "new": "Summary"}),
    );
    assert_eq!(fs::read(&path).unwrap(), before);
}

#[test]
fn dry_run_leaves_bytes_and_history_and_predicts_the_revision() {
    let directory = TempDir::new("dry-run");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);
    let file = text(&path);
    let original = fs::read(&path).unwrap();

    let preview = session.ok(
        "write_cells",
        json!({"path": file, "sheet": "Data", "dry_run": true, "cells": [{"cell": "A1", "value": 1}]}),
    );
    assert_eq!(preview["dry_run"], true);
    assert_eq!(fs::read(&path).unwrap(), original);
    let message = session.fail("undo", json!({"path": file}));
    assert!(!message.is_empty(), "a dry run records no history");

    for (name, arguments) in [
        (
            "write_cells",
            json!({"path": file, "sheet": "Data", "cells": [{"cell": "C2", "formula": "C3+C4"}]}),
        ),
        (
            "set_number_format",
            json!({"path": file, "sheet": "Data", "range": "C2:C4", "format_code": "0.0%"}),
        ),
        (
            "insert_rows",
            json!({"path": file, "sheet": "Data", "at": 2, "count": 2}),
        ),
        (
            "insert_columns",
            json!({"path": file, "sheet": "Data", "at": 1, "count": 1}),
        ),
        (
            "delete_rows",
            json!({"path": file, "sheet": "Data", "at": 2, "count": 1}),
        ),
        (
            "delete_columns",
            json!({"path": file, "sheet": "Data", "at": 1, "count": 1}),
        ),
        (
            "add_sheet",
            json!({"path": file, "name": "Draft", "position": 1}),
        ),
        (
            "rename_sheet",
            json!({"path": file, "old": "Summary", "new": "Totals"}),
        ),
        ("undo", json!({"path": file})),
        ("redo", json!({"path": file})),
    ] {
        let before = fs::read(&path).unwrap();
        let preview = session.ok(name, with(&arguments, json!({"dry_run": true})));
        assert_eq!(preview["dry_run"], true, "{name}: {preview}");
        assert_eq!(
            fs::read(&path).unwrap(),
            before,
            "{name} dry run writes nothing"
        );
        assert_ne!(
            preview["revision"],
            json!(revision_of(&before)),
            "{name} plans a change"
        );
        let applied = session.ok(name, arguments);
        assert_written(&applied, &path);
        assert_eq!(
            applied["revision"], preview["revision"],
            "{name} dry run predicts the revision"
        );
    }
    assert_eq!(session.sheet_names(&path), ["Data", "Draft", "Totals"]);
}

#[test]
fn expected_revision_refuses_stale_edits() {
    let directory = TempDir::new("revision");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);
    let file = text(&path);
    let edit = |revision: &str, value: f64| {
        json!({
            "path": file, "sheet": "Data", "expected_revision": revision,
            "cells": [{"cell": "C2", "value": value}]
        })
    };
    let start = revision(&path);

    let message = session.fail("write_cells", edit(&"0".repeat(64), 1.0));
    assert!(message.contains("revision"), "{message}");
    assert_eq!(revision(&path), start, "a conflict writes nothing");
    session.invalid("write_cells", edit("not-a-sha256", 1.0));

    let first = assert_written(&session.ok("write_cells", edit(&start, 2.0)), &path);
    let stale = session.fail("write_cells", edit(&start, 3.0));
    assert!(stale.contains("revision"), "{stale}");
    assert_eq!(revision(&path), first);

    // A change made outside the server invalidates the revision it last saw.
    let outside = workbook_bytes(999.0);
    fs::write(&path, &outside).unwrap();
    session.fail("write_cells", edit(&first, 4.0));
    session.fail(
        "add_sheet",
        json!({"path": file, "name": "X", "expected_revision": first}),
    );
    assert_eq!(fs::read(&path).unwrap(), outside);
    let current = revision(&path);
    assert_written(&session.ok("write_cells", edit(&current, 5.0)), &path);
    assert_cells(&session.values(&path, "Data", "C2:C3"), json!([[5], [999]]));
}

#[test]
fn undo_and_redo_restore_exact_bytes() {
    let directory = TempDir::new("undo");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);
    let file = text(&path);
    let original = fs::read(&path).unwrap();

    session.fail("undo", json!({"path": file}));
    session.ok(
        "write_cells",
        json!({"path": file, "sheet": "Data", "cells": [{"cell": "A2", "value": "Sprocket"}]}),
    );
    let first = fs::read(&path).unwrap();
    session.ok(
        "insert_rows",
        json!({"path": file, "sheet": "Data", "at": 2, "count": 1}),
    );
    let second = fs::read(&path).unwrap();

    let undone = session.ok(
        "undo",
        json!({"path": file, "expected_revision": revision_of(&second)}),
    );
    assert_eq!(fs::read(&path).unwrap(), first);
    assert_eq!(undone["revision"], revision_of(&first));
    assert_eq!(undone["dry_run"], false);
    session.ok("undo", json!({"path": file}));
    assert_eq!(fs::read(&path).unwrap(), original);
    session.fail("undo", json!({"path": file}));
    assert_eq!(fs::read(&path).unwrap(), original);

    let message = session.fail(
        "redo",
        json!({"path": file, "expected_revision": revision_of(&first)}),
    );
    assert!(message.contains("revision"), "{message}");
    session.ok("redo", json!({"path": file}));
    assert_eq!(fs::read(&path).unwrap(), first);
    session.ok("redo", json!({"path": file}));
    assert_eq!(fs::read(&path).unwrap(), second);
    session.fail("redo", json!({"path": file}));
    assert_cells(
        &session.values(&path, "Data", "A2:A3"),
        json!([[null], ["Sprocket"]]),
    );

    // A new edit after undo discards the redo history.
    session.ok("undo", json!({"path": file}));
    session.ok("add_sheet", json!({"path": file, "name": "Fresh"}));
    session.fail("redo", json!({"path": file}));
    session.ok("undo", json!({"path": file}));
    assert_eq!(fs::read(&path).unwrap(), first);
}

#[test]
fn undo_history_keeps_the_newest_32_edits() {
    let directory = TempDir::new("undo-limit");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);
    let file = text(&path);
    let mut snapshots = Vec::new();
    for value in 0..34 {
        snapshots.push(fs::read(&path).unwrap());
        session.ok(
            "write_cells",
            json!({"path": file, "sheet": "Data", "cells": [{"cell": "C2", "value": value}]}),
        );
    }
    for _ in 0..32 {
        session.ok("undo", json!({"path": file}));
    }
    // The two oldest snapshots were dropped, so history stops after edit two.
    assert_eq!(fs::read(&path).unwrap(), snapshots[2]);
    let message = session.fail("undo", json!({"path": file}));
    assert!(message.contains("nothing to undo"), "{message}");
}

#[test]
fn undo_and_redo_refuse_after_an_outside_change() {
    let directory = TempDir::new("undo-outside");
    let path = workbook(&directory, "book.xlsx");
    let mut session = Session::start(&[]);
    let file = text(&path);

    session.ok("add_sheet", json!({"path": file, "name": "Later"}));
    let outside = workbook_bytes(1.0);
    fs::write(&path, &outside).unwrap();
    let message = session.fail("undo", json!({"path": file}));
    assert!(message.contains("changed"), "{message}");
    assert_eq!(
        fs::read(&path).unwrap(),
        outside,
        "undo leaves the outside change"
    );

    let mut other = Session::start(&[]);
    let path_b = workbook(&directory, "other.xlsx");
    let file_b = text(&path_b);
    other.ok("add_sheet", json!({"path": file_b, "name": "Later"}));
    other.ok("undo", json!({"path": file_b}));
    fs::write(&path_b, &outside).unwrap();
    other.fail("redo", json!({"path": file_b}));
    assert_eq!(
        fs::read(&path_b).unwrap(),
        outside,
        "redo leaves the outside change"
    );

    // History belongs to one server process.
    let mut fresh = Session::start(&[]);
    fresh.fail("undo", json!({"path": file}));
}

#[test]
fn root_confines_every_path_including_symlinks() {
    let root = TempDir::new("root");
    let outside = TempDir::new("outside");
    let inside = workbook(&root, "book.xlsx");
    fs::create_dir(root.path("sub")).unwrap();
    let nested = workbook(&root, "sub/inner.xlsx");
    let nested_bytes = fs::read(&nested).unwrap();
    let other = workbook(&outside, "other.xlsx");
    let other_bytes = fs::read(&other).unwrap();
    let mut session = Session::start(&["--root", root.root()]);

    let relative = session.ok("get_workbook", json!({"path": "book.xlsx"}));
    assert_eq!(
        relative["path"],
        text(&inside),
        "relative paths resolve under --root"
    );
    assert_eq!(relative["revision"], revision(&inside));
    session.ok("get_workbook", json!({"path": text(&inside)}));
    session.ok(
        "read_range",
        json!({"path": "sub/inner.xlsx", "range": "A1"}),
    );
    assert_written(
        &session.ok(
            "write_cells",
            json!({"path": "sub/../book.xlsx", "sheet": "Data", "cells": [{"cell": "C2", "value": 1}]}),
        ),
        &inside,
    );
    assert_eq!(fs::read(&nested).unwrap(), nested_bytes);

    let escape = format!(
        "../{}/other.xlsx",
        outside.0.file_name().unwrap().to_str().unwrap()
    );
    for path in [text(&other), escape.as_str()] {
        let message = session.fail("get_workbook", json!({"path": path}));
        assert!(message.contains("root"), "{path}: {message}");
        session.fail(
            "write_cells",
            json!({"path": path, "sheet": "Data", "cells": [{"cell": "A1", "value": 1}]}),
        );
        session.fail("undo", json!({"path": path}));
    }
    session.fail(
        "diff",
        json!({"path": "book.xlsx", "right_path": text(&other), "key": "Item"}),
    );
    session.fail("get_workbook", json!({"path": "missing.xlsx"}));

    #[cfg(unix)]
    {
        std::os::unix::fs::symlink(&other, root.path("link.xlsx")).unwrap();
        std::os::unix::fs::symlink(&outside.0, root.path("escape")).unwrap();
        for path in ["link.xlsx", "escape/other.xlsx"] {
            let message = session.fail("get_workbook", json!({"path": path}));
            assert!(message.contains("root"), "{path}: {message}");
            session.fail("read_range", json!({"path": path, "range": "A1"}));
            session.fail(
                "write_cells",
                json!({"path": path, "sheet": "Data", "cells": [{"cell": "A1", "value": 1}]}),
            );
            session.fail("add_sheet", json!({"path": path, "name": "Escaped"}));
        }
    }
    assert_eq!(
        fs::read(&other).unwrap(),
        other_bytes,
        "outside files are never written"
    );
}

#[test]
fn protocol_errors_use_json_rpc_codes_and_keep_the_session_alive() {
    let directory = TempDir::new("protocol");
    let path = workbook(&directory, "book.xlsx");
    let original = fs::read(&path).unwrap();
    let mut session = Session::spawn(&[]);

    // Tools need initialize first; ping does not.
    let early = session.request("tools/list", json!({}));
    assert!(early.get("error").is_some(), "{early}");
    assert_eq!(session.request("ping", json!({}))["result"], json!({}));
    let response = session.request("initialize", json!({"capabilities": {}}));
    rpc_error(&response, -32602);
    assert_eq!(
        session.initialize("2025-06-18")["result"]["protocolVersion"],
        "2025-06-18"
    );
    session.notify("notifications/initialized", json!({}));

    // Notifications, known or not, get no response: the next line answers
    // the following request.
    session.notify("notifications/cancelled", json!({"requestId": 99}));
    session.notify("wolfxl/unknown", json!({}));
    rpc_error(&session.request("wolfxl/unknown", json!({})), -32601);

    session.send_line(r#"{"jsonrpc": "2.0", "id": 1, "method": "#);
    let parse = session.receive();
    assert_eq!(parse["id"], Value::Null);
    rpc_error(&parse, -32700);

    session.send_line(r#"{"jsonrpc": "1.0", "id": 7, "method": "ping"}"#);
    rpc_error(&session.receive(), -32600);

    session.send_line(r#"{"jsonrpc": "2.0", "id": "text-id", "method": "ping"}"#);
    let pong = session.receive();
    assert_eq!(pong["id"], "text-id");
    assert_eq!(pong["result"], json!({}));

    session.send_line(r#"{"jsonrpc": "2.0", "id": 8, "method": "tools/call", "params": [1]}"#);
    rpc_error(&session.receive(), -32602);
    rpc_error(
        &session.request("tools/call", json!({"arguments": {}})),
        -32602,
    );

    let file = text(&path);
    let unknown = session.invalid("delete_sheet", json!({"path": file, "name": "Data"}));
    assert!(unknown.contains("delete_sheet"), "{unknown}");
    session.invalid("read_range", json!({"range": "A1"}));
    session.invalid("read_range", json!({"path": file, "range": 5}));
    session.invalid(
        "read_range",
        json!({"path": file, "range": "A1", "extra": true}),
    );
    session.invalid("stats", json!({"path": file, "top": -1}));
    session.invalid(
        "pivot",
        json!({"path": file, "rows": [], "values": ["sum:Amount"]}),
    );
    session.invalid(
        "diff",
        json!({"path": file, "right_path": file, "key": "Item", "key_transform": "upper"}),
    );
    session.invalid(
        "write_cells",
        json!({"path": file, "sheet": "Data", "cells": []}),
    );
    session.invalid(
        "write_cells",
        json!({"path": file, "sheet": "Data", "cells": [{"cell": "A1", "value": 1, "formula": "1"}]}),
    );
    session.invalid(
        "write_cells",
        json!({"path": file, "sheet": "Data", "cells": [{"cell": "A1", "value": [1]}]}),
    );
    session.invalid(
        "write_cells",
        json!({"path": file, "sheet": "Data", "cells": [{"cell": "A1", "value": 1}], "dry_run": "yes"}),
    );
    assert_eq!(
        fs::read(&path).unwrap(),
        original,
        "rejected calls write nothing"
    );

    // The session still serves tools after every error above.
    assert_cells(&session.values(&path, "Data", "A1"), json!([["Item"]]));
    assert!(session.finish().success(), "clean exit at end of input");
}
