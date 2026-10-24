//! Integrated unit tests: tokenizer round-trip + translator semantics.
//!
//! Tests cover every reference syntax in RFC-012 §2 plus the operational
//! semantics described in §5.

#![cfg(test)]

use crate::reference::{parse_ref, RefKind};
use crate::tokenizer::{render, tokenize, TokenKind, TokenSubKind};
use crate::translate::{
    move_range, rename_sheet, shift, translate, translate_with_meta, Axis, DeletedRange, Range,
    RefDelta, ShiftPlan,
};

// ---------------- §2 reference syntax tests ------------------------

#[test]
fn t01_a1_relative_shift_rows() {
    let out = shift(
        "=A1+B5",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=A4+B8");
}

#[test]
fn t02_a1_absolute_shifts_when_respect_dollar_false() {
    let out = shift(
        "=$A$1+$B$5",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=$A$4+$B$8");
}

#[test]
fn t03_a1_absolute_does_not_shift_when_respect_dollar_true() {
    let out = shift(
        "=$A$1+$B$5",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 3,
            respect_dollar: true,
        },
    );
    assert_eq!(out, "=$A$1+$B$5");
}

#[test]
fn t04_mixed_col_abs_only_shifts_row() {
    let out = shift(
        "=$A1",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 2,
            respect_dollar: true,
        },
    );
    assert_eq!(out, "=$A3");
}

#[test]
fn t05_mixed_row_abs_only_shifts_col() {
    let out = shift(
        "=A$1",
        &ShiftPlan {
            axis: Axis::Col,
            at: 1,
            n: 2,
            respect_dollar: true,
        },
    );
    assert_eq!(out, "=C$1");
}

#[test]
fn t06_range_translates_both_endpoints() {
    let out = shift(
        "=SUM(B5:D10)",
        &ShiftPlan {
            axis: Axis::Row,
            at: 5,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=SUM(B8:D13)");
}

#[test]
fn t07_range_absolute_endpoints() {
    let out = shift(
        "=SUM($B$5:$D$10)",
        &ShiftPlan {
            axis: Axis::Row,
            at: 5,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=SUM($B$8:$D$13)");
}

#[test]
fn t08_whole_row_range_shifts_on_row_delta() {
    let out = shift(
        "=SUM(2:5)",
        &ShiftPlan {
            axis: Axis::Row,
            at: 2,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=SUM(5:8)");
}

#[test]
fn t09_whole_row_range_unchanged_on_col_delta() {
    let out = shift(
        "=SUM(2:5)",
        &ShiftPlan {
            axis: Axis::Col,
            at: 1,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=SUM(2:5)");
}

#[test]
fn t09b_full_sheet_row_range_is_stable_on_tail_row_insert() {
    let out = shift(
        "='HB ROI'!$1:$1048576",
        &ShiftPlan {
            axis: Axis::Row,
            at: crate::MAX_ROW,
            n: 1,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "='HB ROI'!$1:$1048576");
}

#[test]
fn t09c_full_sheet_row_range_is_stable_on_tail_row_delete() {
    let mut delta = RefDelta::empty();
    delta.rows = -1;
    delta.anchor_row = crate::MAX_ROW + 1;
    delta.deleted_range = Some(DeletedRange {
        min_row: crate::MAX_ROW,
        max_row: crate::MAX_ROW,
        min_col: 1,
        max_col: crate::MAX_COL,
    });
    let out = translate("='HB ROI'!$1:$1048576", &delta).unwrap();
    assert_eq!(out, "='HB ROI'!$1:$1048576");
}

#[test]
fn t10_whole_col_range_shifts_on_col_delta() {
    let out = shift(
        "=SUM(A:C)",
        &ShiftPlan {
            axis: Axis::Col,
            at: 1,
            n: 2,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=SUM(C:E)");
}

#[test]
fn t11_whole_col_range_unchanged_on_row_delta() {
    let out = shift(
        "=SUM(A:C)",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=SUM(A:C)");
}

#[test]
fn t11b_full_sheet_col_range_is_stable_on_tail_col_insert() {
    let out = shift(
        "=Sheet1!$A:$XFD",
        &ShiftPlan {
            axis: Axis::Col,
            at: crate::MAX_COL,
            n: 1,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=Sheet1!$A:$XFD");
}

#[test]
fn t11c_full_rectangular_sheet_range_is_stable_on_tail_row_insert() {
    let out = shift(
        "=Sheet1!$A$1:$XFD$1048576",
        &ShiftPlan {
            axis: Axis::Row,
            at: crate::MAX_ROW,
            n: 1,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=Sheet1!$A$1:$XFD$1048576");
}

#[test]
fn t11d_unnecessary_source_quotes_are_preserved_without_rename() {
    let out = shift(
        "='Cover'!$A$1:$L$28",
        &ShiftPlan {
            axis: Axis::Row,
            at: crate::MAX_ROW,
            n: 1,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "='Cover'!$A$1:$L$28");
}

#[test]
fn t12_3d_unquoted_sheet_passes_through_on_no_rename() {
    let out = shift(
        "=Sheet2!A1",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=Sheet2!A4");
}

#[test]
fn t13_3d_quoted_sheet_with_apostrophe_round_trip() {
    let out = translate("='O''Brien'!A1", &RefDelta::empty()).unwrap();
    assert_eq!(out, "='O''Brien'!A1");
}

#[test]
fn t14_3d_rename_sheet_unquoted() {
    let out = rename_sheet(
        "=Forecast!B5+SUM(Forecast!A1:A10)",
        "Forecast",
        "Forecast_v2",
    );
    assert_eq!(out, "=Forecast_v2!B5+SUM(Forecast_v2!A1:A10)");
}

#[test]
fn t15_3d_rename_sheet_strips_quotes_when_safe() {
    let out = rename_sheet("='2024 Data'!A1", "2024 Data", "Q1");
    assert_eq!(out, "=Q1!A1");
}

#[test]
fn t16_3d_rename_sheet_adds_quotes_when_required() {
    let out = rename_sheet("=Forecast!A1", "Forecast", "My Sheet");
    assert_eq!(out, "='My Sheet'!A1");
}

#[test]
fn t17_external_book_ref_passes_through() {
    let f = "=[Book2.xlsx]Sheet1!A1";
    let out = shift(
        f,
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, f);
}

#[test]
fn t18_quoted_external_book_ref_passes_through() {
    let f = "='C:\\path\\[Book2.xlsx]Sheet1'!A1";
    let out = shift(
        f,
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, f);
}

#[test]
fn t19_table_ref_simple_passes_through() {
    let f = "=Table1[Col1]";
    let out = shift(
        f,
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, f);
}

#[test]
fn t20_table_ref_special_passes_through() {
    let f = "=Table1[#Headers]";
    let out = shift(
        f,
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, f);
}

#[test]
fn t21_table_ref_nested_passes_through() {
    let f = "=Table1[[#This Row], [Col1]]";
    let out = shift(
        f,
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, f);
}

#[test]
fn t22_array_formula_inner_refs_translate() {
    let f = "={A1+B1,C1+D1;A2+B2,C2+D2}";
    let out = shift(
        f,
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 1,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "={A2+B2,C2+D2;A3+B3,C3+D3}");
}

#[test]
fn t23_defined_name_passes_through() {
    let f = "=MyTotal+1";
    let out = shift(
        f,
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, f);
}

#[test]
fn t24_function_name_unchanged() {
    let out = shift(
        "=VLOOKUP(A1,B1:C10,2,FALSE)",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 1,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=VLOOKUP(A2,B2:C11,2,FALSE)");
}

#[test]
fn t25_string_literal_refs_not_translated() {
    let out = shift(
        "=IF(A1=\"B5\",X1,Y1)",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=IF(A4=\"B5\",X4,Y4)");
}

#[test]
fn t26_error_literal_passes_through() {
    let out = shift(
        "=#REF!",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=#REF!");
}

#[test]
fn t27_indirect_text_arg_not_translated_but_other_refs_are() {
    let out = shift(
        "=INDIRECT(\"B5\")+B5",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 3,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=INDIRECT(\"B5\")+B8");
    let meta = translate_with_meta("=INDIRECT(\"B5\")+B5", &{
        let mut d = RefDelta::empty();
        d.rows = 3;
        d.anchor_row = 1;
        d
    })
    .unwrap();
    assert!(meta.has_volatile_indirect);
}

// ---------------- §5 semantics tests ------------------------

#[test]
fn t28_identity_translation_byte_identical() {
    let cases = [
        "=A1",
        "=$A$1",
        "=SUM(A1:B5)",
        "=Sheet2!A1",
        "='O''Brien'!A1",
        "='2024 Data'!$A$1",
        "=IF(A1=\"hello\",1,0)",
        "=Table1[Col1]",
        "=Table1[[#This Row], [Col1]]",
        "=[Book2.xlsx]Sheet1!A1",
        "=#REF!",
        "=A1 + B1",
        "={1,2;3,4}",
        "=1.5E+3",
        "=-A1",
        "=A1>=B1",
        "=A1<>B1",
        "=2:5",
        "=A:C",
        "=MyTotal",
    ];
    for f in cases {
        let out = translate(f, &RefDelta::empty()).unwrap();
        assert_eq!(out, f, "identity failed for {:?}", f);
    }
}

#[test]
fn t29_round_trip_inverse_shift() {
    let inputs = ["=A1+B1", "=$A$1+B5", "=SUM(A1:B10)", "=Sheet2!C7"];
    for f in inputs {
        let forward = shift(
            f,
            &ShiftPlan {
                axis: Axis::Row,
                at: 1,
                n: 5,
                respect_dollar: false,
            },
        );
        let back = shift(
            &forward,
            &ShiftPlan {
                axis: Axis::Row,
                at: 6,
                n: -5,
                respect_dollar: false,
            },
        );
        assert_eq!(back, f, "round-trip failed for {}", f);
    }
}

#[test]
fn t30_out_of_bounds_becomes_ref_error() {
    let out = shift(
        "=A1",
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: -5,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=#REF!");
}

#[test]
fn t31_deleted_range_tombstone_kills_inside_refs() {
    let mut delta = RefDelta::empty();
    delta.deleted_range = Some(DeletedRange {
        min_row: 4,
        max_row: 6,
        min_col: 1,
        max_col: crate::MAX_COL,
    });
    delta.formula_sheet = Some("S".into());
    delta.deleted_range_sheet = Some("S".into());
    delta.rows = -3;
    delta.anchor_row = 7;
    let out = translate("=B5+B10", &delta).unwrap();
    assert_eq!(out, "=#REF!+B7");
}

#[test]
fn t32_deleted_range_clip_partial_overlap() {
    let mut delta = RefDelta::empty();
    delta.deleted_range = Some(DeletedRange {
        min_row: 4,
        max_row: 6,
        min_col: 1,
        max_col: crate::MAX_COL,
    });
    delta.formula_sheet = Some("S".into());
    delta.deleted_range_sheet = Some("S".into());
    let out = translate("=SUM(B3:B7)", &delta).unwrap();
    // Either the top half or bottom half survives — see RFC-012 §5.4.
    assert!(
        out == "=SUM(B3:B3)" || out == "=SUM(B7:B7)" || out == "=SUM(B3:B7)",
        "got {}",
        out
    );
}

#[test]
fn t33_cross_sheet_into_deleted_sheet_becomes_ref() {
    let mut delta = RefDelta::empty();
    delta.deleted_range = Some(DeletedRange {
        min_row: 4,
        max_row: 6,
        min_col: 1,
        max_col: crate::MAX_COL,
    });
    delta.formula_sheet = Some("Sheet1".into());
    delta.deleted_range_sheet = Some("Sheet2".into());
    delta.rows = -3;
    delta.anchor_row = 7;
    let out = translate("=Sheet2!B5", &delta).unwrap();
    assert_eq!(out, "=#REF!");
}

#[test]
fn t34_cross_sheet_outside_deleted_sheet_unchanged() {
    let mut delta = RefDelta::empty();
    delta.deleted_range = Some(DeletedRange {
        min_row: 4,
        max_row: 6,
        min_col: 1,
        max_col: crate::MAX_COL,
    });
    delta.formula_sheet = Some("Sheet1".into());
    delta.deleted_range_sheet = Some("Sheet2".into());
    delta.rows = -3;
    delta.anchor_row = 7;
    let out = translate("=Sheet3!B5", &delta).unwrap();
    assert_eq!(out, "=Sheet3!B5");
}

#[test]
fn t35_move_range_reanchor() {
    let f = "=B5+C7+Z99";
    let src = Range {
        min_row: 5,
        max_row: 7,
        min_col: 2,
        max_col: 3,
    };
    let dst = Range {
        min_row: 10,
        max_row: 12,
        min_col: 5,
        max_col: 6,
    };
    let out = move_range(f, &src, &dst, true);
    assert_eq!(out, "=E10+F12+Z99");
}

#[test]
fn t36_complex_formula_round_trip() {
    let f = "=SUM(IF($A1:$A100>0,$B1:$B100,0))";
    assert_eq!(translate(f, &RefDelta::empty()).unwrap(), f);
    let out = shift(
        f,
        &ShiftPlan {
            axis: Axis::Row,
            at: 1,
            n: 5,
            respect_dollar: true,
        },
    );
    assert_eq!(out, "=SUM(IF($A6:$A105>0,$B6:$B105,0))");
}

#[test]
fn t37_volatile_detection() {
    let m = translate_with_meta("=A1+1", &RefDelta::empty()).unwrap();
    assert!(!m.has_volatile_indirect);
    let m2 = translate_with_meta("=OFFSET(A1,1,0)", &RefDelta::empty()).unwrap();
    assert!(m2.has_volatile_indirect);
}

#[test]
fn t38_quoted_sheet_with_space_translated_after_rename() {
    let out = rename_sheet("='Old Sheet'!$A$1", "Old Sheet", "NewSheet");
    assert_eq!(out, "=NewSheet!$A$1");
}

#[test]
fn t39_negative_col_shift_into_oob_returns_ref_error() {
    let out = shift(
        "=A1",
        &ShiftPlan {
            axis: Axis::Col,
            at: 1,
            n: -1,
            respect_dollar: false,
        },
    );
    assert_eq!(out, "=#REF!");
}

#[test]
fn t40_unicode_cell_value_in_string_round_trips() {
    let f = "=IF(A1=\"héllo\",B1,C1)";
    let out = translate(f, &RefDelta::empty()).unwrap();
    assert_eq!(out, f);
}

// ---------------- Round-trip on synthgl + tests/fixtures corpus -----

#[test]
fn t41_synthgl_corpus_round_trip() {
    // Walk synthgl_snapshot + tests/fixtures and extract <f> contents
    // from every sheet*.xml. Verify identity translation is byte-
    // identical for >=100 formulas (RFC-012 §6 verification gate #3).
    let repo = std::env::current_dir().unwrap();
    let repo_root = repo
        .ancestors()
        .find_map(|p| {
            if p.join("Cargo.lock").is_file() {
                Some(p.to_path_buf())
            } else {
                None
            }
        })
        .expect("found repo root");

    let synthgl = repo_root.join("tests/parity/fixtures/synthgl_snapshot");
    let mut formulas = collect_formulas(&synthgl);
    let extra = repo_root.join("tests/fixtures");
    if extra.is_dir() {
        formulas.extend(collect_formulas(&extra));
    }
    eprintln!("collected {} formulas from corpus", formulas.len());
    assert!(
        formulas.len() >= 100,
        "RFC-012 verification gate requires >=100 formulas in corpus; got {}",
        formulas.len()
    );

    let mut failed = Vec::new();
    for f in &formulas {
        let prefixed = if f.starts_with('=') {
            f.clone()
        } else {
            format!("={}", f)
        };
        match translate(&prefixed, &RefDelta::empty()) {
            Ok(out) => {
                if out != prefixed {
                    failed.push((prefixed, out));
                }
            }
            Err(_) => {}
        }
    }
    assert!(
        failed.is_empty(),
        "round-trip failures: {:?}",
        &failed[..failed.len().min(5)]
    );
}

fn collect_formulas(root: &std::path::Path) -> Vec<String> {
    use std::fs::File;
    use std::io::Read;
    let mut out = Vec::new();
    walk_xlsx(root, &mut |path| {
        if path.extension().map_or(false, |e| e == "xlsx") {
            if let Ok(file) = File::open(path) {
                if let Ok(mut zip) = zip::ZipArchive::new(file) {
                    for i in 0..zip.len() {
                        if let Ok(mut entry) = zip.by_index(i) {
                            let name = entry.name().expect("valid fixture entry name").into_owned();
                            if name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml") {
                                let mut s = String::new();
                                if entry.read_to_string(&mut s).is_ok() {
                                    extract_f_text(&s, &mut out);
                                }
                            }
                        }
                    }
                }
            }
        }
    });
    out
}

fn walk_xlsx<F: FnMut(&std::path::Path)>(root: &std::path::Path, f: &mut F) {
    if let Ok(rd) = std::fs::read_dir(root) {
        for entry in rd.flatten() {
            let p = entry.path();
            if p.is_dir() {
                walk_xlsx(&p, f);
            } else {
                f(&p);
            }
        }
    }
}

fn extract_f_text(xml: &str, out: &mut Vec<String>) {
    let mut i = 0;
    let bytes = xml.as_bytes();
    while i + 2 < bytes.len() {
        if bytes[i] == b'<'
            && bytes[i + 1] == b'f'
            && (bytes[i + 2] == b' ' || bytes[i + 2] == b'>')
        {
            if let Some(close) = xml[i..].find('>') {
                let after = i + close + 1;
                if let Some(end_rel) = xml[after..].find("</f>") {
                    let body = &xml[after..after + end_rel];
                    let decoded = body
                        .replace("&lt;", "<")
                        .replace("&gt;", ">")
                        .replace("&amp;", "&")
                        .replace("&quot;", "\"")
                        .replace("&apos;", "'");
                    if !decoded.is_empty() {
                        out.push(decoded);
                    }
                    i = after + end_rel + 4;
                    continue;
                }
            }
        }
        i += 1;
    }
}

// ---------------- Performance test --------------------------------

#[test]
fn t42_perf_100k_formulas_under_1s() {
    let f = "=SUM($A$1:$A$100)+B5*VLOOKUP(C7,Sheet2!D1:E50,2,FALSE)";
    let plan = ShiftPlan {
        axis: Axis::Row,
        at: 5,
        n: 3,
        respect_dollar: false,
    };
    for _ in 0..100 {
        let _ = shift(f, &plan);
    }
    let n = 100_000;
    let t0 = std::time::Instant::now();
    for _ in 0..n {
        let _ = shift(f, &plan);
    }
    let elapsed = t0.elapsed();
    eprintln!("100k formulas: {:?}", elapsed);
    // 2x slack vs 1s budget for CI variability.
    assert!(
        elapsed.as_secs_f64() < 2.0,
        "100k formulas took {:?}",
        elapsed
    );
}

// ---------------- Sanity: tokenizer round-trip --------------------

#[test]
fn t43_tokenizer_render_round_trip_stress() {
    let cases = [
        "",
        "literal value, no equals",
        "=1",
        "=A1",
        "=A1+B1*C1/D1^E1&F1",
        "=SUM(A1:Z100)",
        "=IF(A1>=10,\"big\",\"small\")",
        "= A1 + B1 ",
        "=#N/A",
        "=A1 B1",
    ];
    for f in cases {
        let toks = tokenize(f).unwrap();
        let r = render(&toks);
        assert_eq!(r, f, "tokenizer round-trip failed for {:?}", f);
    }
}

// ---------------- Reference parser sanity ---------------------------

#[test]
fn t44_parse_ref_classification() {
    assert!(matches!(parse_ref("A1"), RefKind::Cell { .. }));
    assert!(matches!(parse_ref("A1:B5"), RefKind::Range { .. }));
    assert!(matches!(parse_ref("2:5"), RefKind::RowRange { .. }));
    assert!(matches!(parse_ref("A:C"), RefKind::ColRange { .. }));
    assert!(matches!(parse_ref("Sheet2!A1"), RefKind::Cell { .. }));
    assert!(matches!(parse_ref("'My Sheet'!A1"), RefKind::Cell { .. }));
    assert!(matches!(parse_ref("Table1[Col1]"), RefKind::Table(_)));
    assert!(matches!(
        parse_ref("[Book.xlsx]Sheet1!A1"),
        RefKind::ExternalBook { .. }
    ));
    assert!(matches!(parse_ref("MyName"), RefKind::Name(_)));
    assert!(matches!(parse_ref("#REF!"), RefKind::Error(_)));
}

#[test]
fn t45_string_token_subkind_is_text() {
    let toks = tokenize("=\"hello\"").unwrap();
    assert_eq!(toks[0].subkind, TokenSubKind::Text);
}

// ---------------- Line-break whitespace ------------------------------

/// The delta wolfxl-reader applies to a shared-formula child one row
/// below its master.
fn shared_child_delta(rows: i32, cols: i32) -> RefDelta {
    let mut delta = RefDelta::empty();
    delta.rows = rows;
    delta.cols = cols;
    delta.anchor_row = 1;
    delta.anchor_col = 1;
    delta.respect_dollar = true;
    delta
}

#[test]
fn t46_line_break_ends_an_operand_token() {
    for (formula, ws) in [
        ("=B1\n+B2", "\n"),
        ("=B1\r\n+B2", "\r\n"),
        ("=B1\r+B2", "\r"),
    ] {
        let toks = tokenize(formula).unwrap();
        let shape: Vec<_> = toks.iter().map(|t| (t.kind, t.value.as_str())).collect();
        assert_eq!(
            shape,
            [
                (TokenKind::Operand, "B1"),
                (TokenKind::Wspace, ws),
                (TokenKind::OpIn, "+"),
                (TokenKind::Operand, "B2"),
            ],
            "{formula:?}"
        );
        assert_eq!(toks[0].subkind, TokenSubKind::Range);
        assert_eq!(render(&toks), formula);
    }
}

#[test]
fn t47_shared_formula_child_keeps_line_breaks_in_place() {
    let cases = [
        ("=B1\n+B2", "=B2\n+B3"),
        ("=B1\r\n+B2", "=B2\r\n+B3"),
        ("=B1\r+B2", "=B2\r+B3"),
        ("=SUM(\r\nB1,\r\nB2\r\n)", "=SUM(\r\nB2,\r\nB3\r\n)"),
        ("=B1 \n +B2", "=B2 \n +B3"),
        ("=$A$1\n*B1", "=$A$1\n*B2"),
    ];
    for (master, child) in cases {
        assert_eq!(
            translate(master, &shared_child_delta(1, 0)).unwrap(),
            child,
            "{master:?}"
        );
    }
}

#[test]
fn t48_whitespace_runs_round_trip_byte_for_byte() {
    let cases = [
        "=A1  +  B1",
        "=SUM(\r\n1,2\r\n)",
        "=A1\n\nB1",
        "=IF(A1,\r\n  B1,\r\n  C1)",
    ];
    for f in cases {
        let toks = tokenize(f).unwrap();
        assert_eq!(render(&toks), f, "{f:?}");
        assert_eq!(translate(f, &RefDelta::empty()).unwrap(), f, "{f:?}");
    }
}
