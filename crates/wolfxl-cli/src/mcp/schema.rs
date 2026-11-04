use serde_json::{json, Map, Value};

pub(super) const MUTATORS: &[&str] = &[
    "write_cells",
    "set_number_format",
    "insert_rows",
    "delete_rows",
    "insert_columns",
    "delete_columns",
    "add_sheet",
    "rename_sheet",
    "undo",
    "redo",
];

fn string(description: &str) -> Value {
    json!({"type":"string", "minLength":1, "description":description})
}

fn strings(description: &str) -> Value {
    json!({"type":"array", "items":string("Header or expression"), "description":description})
}

fn integer(description: &str, minimum: u64, maximum: u64) -> Value {
    json!({"type":"integer", "minimum":minimum, "maximum":maximum, "description":description})
}

fn tool(name: &str, description: &str, fields: &[(&str, Value)], required: &[&str]) -> Value {
    let mut properties = Map::new();
    properties.insert("path".into(), string("Existing workbook path; relative paths resolve under --root when configured, otherwise the working directory."));
    for (name, schema) in fields {
        properties.insert((*name).into(), schema.clone());
    }
    if MUTATORS.contains(&name) {
        properties.insert("dry_run".into(), json!({"type":"boolean", "default":false, "description":"Validate and return the planned change and would-be revision without writing or changing history."}));
        properties.insert("expected_revision".into(), json!({"type":"string", "pattern":"^[0-9a-f]{64}$", "description":"SHA-256 file revision from a prior read; refuse if the file has changed."}));
    }
    let required: Vec<&str> = std::iter::once("path")
        .chain(required.iter().copied())
        .collect();
    json!({"name":name, "description":description, "inputSchema":{"type":"object", "properties":properties, "required":required, "additionalProperties":false}, "annotations":{"readOnlyHint":!MUTATORS.contains(&name), "destructiveHint":MUTATORS.contains(&name), "openWorldHint":false}})
}

pub(super) fn tools(read_only: bool) -> Vec<Value> {
    let sheet = || {
        (
            "sheet",
            string("Exact worksheet name; omitted reads use the first sheet."),
        )
    };
    let range = || {
        ("range", string("Inclusive A1 rectangle, for example B2:D9, or one cell B2. Rows and columns are 1-based."))
    };
    let cells = json!({"type":"array", "minItems":1, "items":{"type":"object", "properties":{"cell":string("A1 cell reference without a sheet prefix; rows and columns are 1-based."), "value":{"type":["number","string","boolean","null"], "description":"Stored scalar; null clears value and formula while retaining style. Text beginning with = stays text."}, "formula":string("Formula with or without leading =; no cached value is written.")}, "required":["cell"], "additionalProperties":false, "oneOf":[{"required":["value"]},{"required":["formula"]}]}});
    let mut result = vec![
        tool("get_workbook", "Orient before reading: list sheets, dimensions in rows/columns, headers, source format, and SHA-256 file revision. Reads all supported spreadsheet and delimited formats.", &[], &[]),
        tool("read_range", "Read stored values and formula text (without leading =) as row-major matrices in an inclusive A1 rectangle. Coordinates are 1-based, blanks are null, and formulas are not recalculated. Maximum 10,000 cells per call; page by requesting adjacent smaller row or column rectangles.", &[sheet(), range()], &["range"]),
        tool("write_cells", "Batch write typed values, formulas, or clears in an existing XLSX/XLSM sheet, preserving untouched package parts. Use expected_revision after reading; use dry_run to preview. Formulas recalculate when Excel opens the file.", &[sheet(), ("cells",cells)], &["sheet","cells"]),
        tool("set_number_format", "Apply an Excel number-format code to an inclusive A1 range while keeping other style properties. Does not change stored values or their units; XLSX/XLSM only.", &[sheet(),range(),("format_code",string("Excel format code, for example $#,##0.00 or 0.0%."))], &["sheet","range","format_code"]),
        tool("add_sheet", "Add a blank worksheet to an existing XLSX/XLSM workbook. Omit position to append; position is a 0-based sheet-list insertion index, unlike cell coordinates.", &[("name",string("New unique worksheet name.")),("position",integer("0-based insertion position; omit to append.",0,u32::MAX.into()))], &["name"]),
        tool("rename_sheet", "Rename an XLSX/XLSM worksheet and rewrite references in formulas and defined names. Read get_workbook first to discover exact names.", &[("old",string("Existing worksheet name.")),("new",string("New unique worksheet name."))], &["old","new"]),
        tool("undo", "Restore the exact prior file bytes from this process's per-path history. Refuses if the file changed outside this server since its last write. History is in memory and disappears on exit.", &[], &[]),
        tool("redo", "Restore the exact bytes reverted by undo for this path. A new edit clears redo history; outside changes are refused. History is local to this server process.", &[], &[]),
        tool("calc", "Evaluate one formula against a sheet's stored values without modifying the file or recalculating existing formula caches. Numeric values use underlying Excel units; dates are Excel serial days. Returns formula, sheet, type, and value.", &[sheet(),("formula",string("Formula with or without leading =; A1 references are 1-based."))], &["formula"]),
        tool("stats", "Profile columns using the first occupied row as headers: types, blanks, distinct values, numeric sum/mean/min/max/population standard deviation, and frequent values. Numeric summaries use stored units and recognize money-style text.", &[sheet(),("filters",strings("Filter expressions COLUMN<op>VALUE, combined with AND; operators = != < > <= >= ~ (contains).")),("select",strings("Headers to profile, in requested order; omit for all.")),("top",integer("Most frequent values per column; 0 returns all distinct values.",0,u32::MAX.into()))], &[]),
        tool("pivot", "Group data rows by header-named fields and aggregate with grand totals. The first occupied row supplies headers. sum/avg/min/max use numeric stored values in their original units; count counts non-empty values.", &[sheet(),("rows",json!({"type":"array","minItems":1,"items":string("Grouping header")})),("column",string("Optional header whose distinct values become columns.")),("values",json!({"type":"array","minItems":1,"items":string("Aggregation AGG:FIELD; AGG is sum, count, avg, min, or max.")}))], &["rows","values"]),
        tool("diff", "Reconcile this path against right_path by a unique header-named key. Returns the existing keyed diff JSON report, including unmatched rows and numeric deltas (right minus left) in source units. Both paths are confined by --root.", &[sheet(),("right_path",string("Existing right-hand spreadsheet path.")),("key",string("Unique non-empty key column header on both sides.")),("compare",strings("Headers to compare; omit for shared non-key columns.")),("tolerance",json!({"type":"number","minimum":0,"default":0,"description":"Largest absolute numeric delta accepted, in stored value units."})),("key_transform",json!({"type":"string","enum":["none","trim","lower","digits"],"default":"none","description":"Key normalization: exact, trim, trim and lowercase, or ASCII digits only."}))], &["right_path","key"]),
        tool("fingerprint", "Compute the versioned style-independent wxf1 content fingerprint and sheet/cell/formula counts. Use to compare workbook contents; unlike revision, formatting and ZIP bytes do not affect it.", &[], &[]),
    ];
    for (name, description, limit) in [
        ("insert_rows", "Insert rows at a 1-based row index, shifting rows at and below it down. Count is in rows; XLSX/XLSM only.",1_048_576),
        ("delete_rows", "Delete count rows starting at a 1-based row index, shifting later rows up. Count is in rows; XLSX/XLSM only.",1_048_576),
        ("insert_columns", "Insert columns at a 1-based column index (A=1), shifting columns at and after it right. Count is in columns; XLSX/XLSM only.",16_384),
        ("delete_columns", "Delete count columns starting at a 1-based column index (A=1), shifting later columns left. Count is in columns; XLSX/XLSM only.",16_384),
    ] {
        result.push(tool(name,description,&[sheet(),("at",integer("1-based starting index.",1,limit)),("count",integer("Positive number of rows or columns.",1,limit))],&["sheet","at","count"]));
    }
    result.retain(|entry| {
        !read_only || !MUTATORS.contains(&entry["name"].as_str().unwrap_or_default())
    });
    result
}

/// Validate the small JSON Schema vocabulary used by this server's tool contracts.
pub(super) fn validate(schema: &Value, value: &Value, location: &str) -> Result<(), String> {
    if let Some(kind) = schema.get("type") {
        let accepts = |kind: &str| match kind {
            "object" => value.is_object(),
            "array" => value.is_array(),
            "string" => value.is_string(),
            "number" => value.is_number(),
            "integer" => value.as_u64().is_some() || value.as_i64().is_some(),
            "boolean" => value.is_boolean(),
            "null" => value.is_null(),
            _ => false,
        };
        let valid = kind.as_str().map_or_else(
            || {
                kind.as_array().is_some_and(|kinds| {
                    kinds.iter().any(|kind| kind.as_str().is_some_and(accepts))
                })
            },
            accepts,
        );
        if !valid {
            return Err(format!("{location}: expected {kind}"));
        }
    }
    if let Some(allowed) = schema["enum"].as_array() {
        if !allowed.contains(value) {
            return Err(format!("{location}: expected one of {}", schema["enum"]));
        }
    }
    if let Some(text) = value.as_str() {
        if schema["minLength"]
            .as_u64()
            .is_some_and(|min| text.chars().count() < min as usize)
        {
            return Err(format!("{location}: must not be empty"));
        }
        if schema.get("pattern").is_some()
            && (text.len() != 64
                || !text
                    .bytes()
                    .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte)))
        {
            return Err(format!(
                "{location}: expected 64 lowercase SHA-256 hex digits"
            ));
        }
    }
    if let Some(number) = value.as_f64() {
        if schema["minimum"].as_f64().is_some_and(|min| number < min)
            || schema["maximum"].as_f64().is_some_and(|max| number > max)
        {
            return Err(format!("{location}: outside allowed numeric bounds"));
        }
    }
    if let Some(object) = value.as_object() {
        if let Some(required) = schema["required"].as_array() {
            for name in required.iter().filter_map(Value::as_str) {
                if !object.contains_key(name) {
                    return Err(format!("{location}.{name}: required"));
                }
            }
        }
        for (name, value) in object {
            if let Some(child) = schema["properties"].get(name) {
                validate(child, value, &format!("{location}.{name}"))?;
            } else if schema["additionalProperties"] == false {
                return Err(format!("{location}.{name}: unknown property"));
            }
        }
    }
    if let Some(items) = value.as_array() {
        if schema["minItems"]
            .as_u64()
            .is_some_and(|min| items.len() < min as usize)
        {
            return Err(format!("{location}: too few items"));
        }
        if let Some(child) = schema.get("items") {
            for (index, value) in items.iter().enumerate() {
                validate(child, value, &format!("{location}[{index}]"))?;
            }
        }
    }
    if let Some(alternatives) = schema["oneOf"].as_array() {
        if alternatives
            .iter()
            .filter(|child| validate(child, value, location).is_ok())
            .count()
            != 1
        {
            return Err(format!(
                "{location}: supply exactly one of value or formula"
            ));
        }
    }
    Ok(())
}
