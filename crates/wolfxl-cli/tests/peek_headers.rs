//! Absolute headers, column letters, address provenance, and percent imports.

use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};
use std::process::Output;

use assert_cmd::Command;
use serde_json::Value;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

fn cities() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/data-table-cities.xlsx")
}

fn run(command: &str, args: &[&str], stdin: Option<&[u8]>) -> Output {
    let mut cmd = Command::cargo_bin("wolfxl").expect("wolfxl built");
    cmd.arg(command).args(args);
    if let Some(stdin) = stdin {
        cmd.write_stdin(stdin);
    }
    cmd.output().expect("run wolfxl")
}

fn success(output: Output) -> String {
    assert_eq!(
        output.status.code(),
        Some(0),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    String::from_utf8(output.stdout).expect("utf-8")
}

fn refused(output: Output) -> String {
    assert_eq!(
        output.status.code(),
        Some(2),
        "{}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(output.stdout.is_empty());
    String::from_utf8(output.stderr).expect("utf-8")
}

fn city_run(command: &str, args: &[&str]) -> Output {
    let file = cities();
    let mut full = vec![file.to_str().expect("fixture path"), "--sheet", "Table"];
    full.extend_from_slice(args);
    run(command, &full, None)
}

#[test]
fn absolute_headers_and_sheet_letters_filter_both_tables_without_mixing_them() {
    let out = success(city_run(
        "peek",
        &[
            "--header-row",
            "2",
            "--where",
            "F>15000000",
            "--select",
            "C,F",
            "-e",
            "json",
        ],
    ));
    let json: Value = serde_json::from_str(&out).expect("JSON");
    assert_eq!(json["headers"], serde_json::json!(["City", "Population"]));
    let names: Vec<&str> = json["data"]
        .as_array()
        .unwrap()
        .iter()
        .map(|row| row[0].as_str().unwrap())
        .collect();
    assert_eq!(
        names,
        [
            "Shanghai",
            "Karachi",
            "Beijing",
            "Sao Paulo",
            "Delhi",
            "Lagos",
            "Total"
        ]
    );
    let out = success(city_run(
        "peek",
        &[
            "--header-row",
            "2",
            "--where",
            "F>15000000",
            "--where",
            "C!=Total",
            "--select",
            "C,F",
            "-e",
            "json",
        ],
    ));
    let json: Value = serde_json::from_str(&out).unwrap();
    assert_eq!(json["data"].as_array().unwrap().len(), 6);
    assert_eq!(json["data"][0][1], 24256800);
}

#[test]
fn duplicate_headers_identify_both_cells_in_each_command() {
    for (command, fields) in [
        ("peek", vec!["--select", "City"]),
        ("peek", vec!["--where", "City=Shanghai"]),
        ("stats", vec!["--select", "City"]),
        ("pivot", vec!["--rows", "City", "--values", "sum:F"]),
        (
            "pivot",
            vec!["--rows", "C", "--column", "Population", "--values", "sum:F"],
        ),
        ("pivot", vec!["--rows", "C", "--values", "sum:Population"]),
    ] {
        let mut args = vec!["--header-row", "2"];
        args.extend(fields);
        let error = refused(city_run(command, &args));
        assert!(error.contains("ambiguous"), "{command}: {error}");
        let (left, right) = if error.contains("Population") {
            ("F2", "L2")
        } else {
            ("C2", "I2")
        };
        assert!(
            error.contains(left) && error.contains(right),
            "{command}: {error}"
        );
        assert!(error.contains("column letter"), "{command}: {error}");
    }
}

#[test]
fn out_of_range_header_rows_report_the_sheet_range() {
    for command in ["peek", "stats", "pivot"] {
        for row in ["0", "40"] {
            let mut args = vec!["--header-row", row];
            if command == "pivot" {
                args.extend(["--rows", "C", "--values", "sum:F"]);
            }
            let error = refused(city_run(command, &args));
            assert!(error.contains("rows 1-15"), "{command}: {error}");
        }
    }
}

#[test]
fn profiles_and_pivots_include_the_stored_total_as_data() {
    let out = success(city_run(
        "stats",
        &["--header-row", "2", "--select", "F", "-e", "json"],
    ));
    let json: Value = serde_json::from_str(&out).unwrap();
    assert_eq!(json["rows"], 13);
    assert_eq!(json["columns"][0]["header"], "Population");
    assert_eq!(json["columns"][0]["numeric"]["sum"], 397811400);
    let out = success(city_run(
        "pivot",
        &[
            "--header-row",
            "2",
            "--rows",
            "C",
            "--values",
            "sum:F",
            "-e",
            "csv",
        ],
    ));
    assert!(out.starts_with("City,sum:Population\n"), "{out}");
    assert!(out.contains("Shanghai,\"24,256,800\""), "{out}");
    assert!(out.ends_with("Total,\"397,811,400\"\n"), "{out}");
}

#[test]
fn address_previews_keep_original_rows_reordered_columns_and_computed_markers() {
    for format in ["box", "text", "markdown"] {
        let out = success(city_run(
            "peek",
            &[
                "--header-row",
                "2",
                "--addresses",
                "--add",
                "Double=F3*2",
                "--where",
                "C=Karachi",
                "--select",
                "Double,F,C",
                "-e",
                format,
            ],
        ));
        assert!(
            out.contains("Double") && out.contains("Karachi"),
            "{format}: {out}"
        );
        assert!(
            out.contains("47,000,000") || out.contains("47000000"),
            "{format}: {out}"
        );
        let lines: Vec<&str> = out.lines().collect();
        let address_line = lines
            .iter()
            .find(|line| line.contains('+') && line.contains('F') && line.contains('C'))
            .expect("column letters");
        assert!(address_line.find('+') < address_line.find('F'), "{out}");
        assert!(address_line.find('F') < address_line.find('C'), "{out}");
        let row = lines.iter().find(|line| line.contains("Karachi")).unwrap();
        assert!(row.contains('4'), "source row 4: {out}");
    }
    let out = success(city_run(
        "peek",
        &["--header-row", "2", "--addresses", "-n", "3"],
    ));
    let line = out
        .lines()
        .find(|line| {
            line.contains('C') && line.contains('D') && line.contains('L') && !line.contains("City")
        })
        .expect("C through L labels");
    for letter in ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L"] {
        assert!(line.contains(letter), "{out}");
    }
    for (row, city) in [(3, "Shanghai"), (4, "Karachi"), (5, "Beijing")] {
        let line = out.lines().find(|line| line.contains(city)).unwrap();
        assert!(line.contains(&row.to_string()), "{out}");
    }
    assert!(!out.contains("Tianjin"), "three data rows only: {out}");
}

#[test]
fn machine_exports_refuse_address_labels() {
    for format in ["csv", "json"] {
        let error = refused(city_run("peek", &["--addresses", "-e", format]));
        assert!(
            error.contains("--addresses") && error.contains("CSV or JSON"),
            "{error}"
        );
    }
}

/// Sparse workbook whose used range starts at AA5, not row 1 or column A.
fn offset_workbook() -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, text) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Offset" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="5"><c r="AA5" t="inlineStr"><is><t>Item</t></is></c><c r="AB5" t="inlineStr"><is><t>Amount</t></is></c></row><row r="6"><c r="AA6" t="inlineStr"><is><t>x</t></is></c><c r="AB6"><v>10</v></c></row><row r="7"><c r="AA7" t="inlineStr"><is><t>y</t></is></c><c r="AB7"><v>20</v></c></row></sheetData></worksheet>"#,
        ),
    ] {
        writer
            .start_file(name, SimpleFileOptions::default())
            .unwrap();
        writer.write_all(text.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

#[test]
fn sheet_row_numbers_and_multi_letter_columns_are_absolute() {
    let bytes = offset_workbook();
    let out = success(run(
        "peek",
        &[
            "-",
            "--from",
            "xlsx",
            "--header-row",
            "5",
            "--select",
            "AB,AA",
            "--where",
            "AB>10",
            "--addresses",
            "-e",
            "text",
        ],
        Some(&bytes),
    ));
    assert!(out.contains("AB\tAA") && out.contains("7\t20\ty"), "{out}");
    for row in ["1", "4", "8"] {
        let error = refused(run(
            "peek",
            &["-", "--from", "xlsx", "--header-row", row],
            Some(&bytes),
        ));
        assert!(error.contains("rows 5-7"), "{error}");
    }
    let out = success(run(
        "peek",
        &["-", "--from", "csv", "--select", "A", "-e", "csv"],
        Some(b"Name,A\nx,7\n"),
    ));
    assert_eq!(out, "A\n7\n", "unique exact header wins over column letter");
}

#[test]
fn chosen_header_names_apply_to_every_export() {
    for format in ["csv", "json", "text", "markdown", "box"] {
        let out = success(run(
            "peek",
            &[
                "-",
                "--from",
                "csv",
                "--header-row",
                "2",
                "--where",
                "amount>1",
                "--select",
                "Name,Amount",
                "-e",
                format,
            ],
            Some(b"Title,Ignore\nName,Amount\nx,2\ny,1\n"),
        ));
        assert!(
            out.contains("Name") && out.contains("Amount") && out.contains('x'),
            "{format}: {out}"
        );
        assert!(
            !out.contains("Title") && !out.contains("Ignore"),
            "{format}: {out}"
        );
    }
}

#[test]
fn chosen_headers_drive_stats_filters_and_pivot_column_fields() {
    let csv =
        b"Report,Ignore,Ignore\nRegion,Year,Amount\nWest,2025,10\nEast,2026,20\nWest,2026,5\n";
    let out = success(run(
        "stats",
        &[
            "-",
            "--from",
            "csv",
            "--header-row",
            "2",
            "--where",
            "C>5",
            "--select",
            "C",
            "-e",
            "json",
        ],
        Some(csv),
    ));
    let json: Value = serde_json::from_str(&out).unwrap();
    assert_eq!(json["rows"], 2);
    assert_eq!(json["columns"][0]["header"], "Amount");
    assert_eq!(json["columns"][0]["numeric"]["sum"], 30);
    let out = success(run(
        "pivot",
        &[
            "-",
            "--from",
            "csv",
            "--header-row",
            "2",
            "--rows",
            "A",
            "--column",
            "B",
            "--values",
            "sum:C",
            "-e",
            "csv",
        ],
        Some(csv),
    ));
    assert_eq!(
        out,
        "Region,2025,2026,Total\nEast,,20,20\nWest,10,5,15\nTotal,10,25,35\n"
    );
    let out = success(city_run(
        "peek",
        &["--header-row", "15", "--select", "F", "-e", "json"],
    ));
    let json: Value = serde_json::from_str(&out).unwrap();
    assert_eq!(json["rows"], 0);
    assert_eq!(json["data"], serde_json::json!([]));
}

#[test]
fn csv_export_still_imports_grouped_populations_for_calc() {
    let csv = success(city_run("peek", &["-e", "csv"]));
    let out = success(run(
        "calc",
        &["SUM(D3:D14)", "-", "--from", "csv"],
        Some(csv.as_bytes()),
    ));
    assert_eq!(out.trim(), "198905700");
}

const PERCENT: &str = "Name,Rate\npositive,12%\nnegative,-3.5%\naccounting,(12%)\ngrouped,\"1,200%\"\nmalformed,\"1,23,4%\"\nmixed,12% growth\nidentifier,01%\n";

#[test]
fn percent_filters_use_the_same_numeric_rule_on_cells_and_conditions() {
    for (condition, expected) in [
        ("Rate=0.12", "positive,12%"),
        ("Rate=-3.5%", "negative,-3.5%"),
        ("Rate=-0.12", "accounting,(12%)"),
        ("Rate=12", "grouped,\"1,200%\""),
    ] {
        let out = success(run(
            "peek",
            &["-", "--from", "csv", "--where", condition, "-e", "csv"],
            Some(PERCENT.as_bytes()),
        ));
        assert_eq!(out, format!("Name,Rate\n{expected}\n"));
    }
}

#[test]
fn calc_percent_imports_numbers_but_preserves_malformed_and_identifier_text() {
    for (formula, expected) in [
        ("B2", "0.12"),
        ("B3", "-0.035"),
        ("B4", "-0.12"),
        ("B5", "12"),
        ("B6", "1,23,4%"),
        ("B7", "12% growth"),
        ("B8", "01%"),
        ("SUM(B2:B5)", "11.965"),
    ] {
        let out = success(run(
            "calc",
            &[formula, "-", "--from", "csv"],
            Some(PERCENT.as_bytes()),
        ));
        assert_eq!(out.trim(), expected, "{formula}");
    }
}
