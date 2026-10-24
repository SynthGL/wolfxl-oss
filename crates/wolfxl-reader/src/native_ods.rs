// SPDX-License-Identifier: MIT
//! Independently authored ODS value reader from ODF 1.3 Parts 2, 3 and 4.
//! Dates/date-times and time/duration caches are lossless `DateIso` + `String`:
//! `P`/`-P` prefixes distinguish durations from calendar dates. No display-string
//! parsing or lossy epoch conversion is performed. `date1904` honors null-date.
mod package;
mod xml;

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Cursor, Read, Seek};
use std::path::Path;

use crate::{
    Cell, CellDataType, CellValue, NamedRange, ReaderError, Result, SheetState, WorksheetData,
};
use xml::{malformed, Element, Limits, Name, Token, OFFICE, STYLE, TABLE, TEXT};

#[derive(Debug)]
pub struct NativeOdsBook {
    sheets: Vec<Sheet>,
    names: Vec<NamedRange>,
    date1904: bool,
    name_bases: Vec<Option<String>>,
}

#[derive(Debug)]
struct Sheet {
    name: String,
    state: SheetState,
    style: Option<String>,
    data: WorksheetData,
}

impl NativeOdsBook {
    pub fn open_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::open(File::open(path)?, Limits::default())
    }

    pub fn open_bytes(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        Self::open(Cursor::new(bytes.into()), Limits::default())
    }

    fn open<R: Read + Seek>(input: R, limits: Limits) -> Result<Self> {
        let mut parts = package::read(input, limits)?;
        let content = parts
            .remove("content.xml")
            .ok_or_else(|| ReaderError::MissingPart("content.xml".into()))?;
        let mut parser = Parser::new(limits);
        if let Some(styles) = parts.remove("styles.xml") {
            xml::walk(&styles, limits, |t, ancestors| parser.styles(t, ancestors))?;
        }
        xml::walk(&content, limits, |t, ancestors| {
            parser.content(t, ancestors)
        })?;
        parser.finish()
    }

    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    pub fn sheet_state(&self, name: &str) -> Result<SheetState> {
        Ok(self.sheet(name)?.state)
    }

    pub fn named_ranges(&self) -> &[NamedRange] {
        &self.names
    }

    /// ODF relative defined names are interpreted as offsets from this base.
    /// Indices correspond exactly to `named_ranges()` document order.
    pub fn named_range_base_cell_address(&self, index: usize) -> Option<&str> {
        self.name_bases.get(index).and_then(Option::as_deref)
    }

    pub fn date1904(&self) -> bool {
        self.date1904
    }

    // ODF number styles are not Excel style ids. Typed dates/durations do not
    // need a synthetic number format, and numeric cells keep their raw value.
    pub fn number_format_for_style_id(&self, _style_id: u32) -> Option<&str> {
        None
    }

    pub fn worksheet(&self, name: &str) -> Result<WorksheetData> {
        Ok(self.sheet(name)?.data.clone())
    }

    fn sheet(&self, name: &str) -> Result<&Sheet> {
        self.sheets
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| ReaderError::SheetNotFound(name.into()))
    }
}

#[derive(Default)]
struct TableStyle {
    parent: Option<String>,
    display: Option<bool>,
}

struct Parser {
    book: NativeOdsBook,
    limits: Limits,
    styles: HashMap<String, TableStyle>,
    style: Option<(usize, String)>,
    table_depth: Option<usize>,
    row: Option<Row>,
    cell: Option<CellBuilder>,
    next_row: u64,
    next_column: u64,
    cells: usize,
    expanded_text: usize,
    spreadsheet_seen: bool,
    null_date_seen: bool,
    ignored_depth: Option<usize>,
    sheet_names: HashSet<String>,
    defined_names: HashSet<(Option<usize>, String)>,
}

struct Row {
    depth: usize,
    repeat: u32,
    next_col: u64,
    cells: Vec<CellTemplate>,
}

struct CellTemplate {
    col: u32,
    repeat: u32,
    rows_spanned: u32,
    cols_spanned: u32,
    cell: Option<Cell>,
}

struct CellBuilder {
    depth: usize,
    col: u32,
    repeat: u32,
    rows_spanned: u32,
    cols_spanned: u32,
    kind: Option<String>,
    scalar: Option<String>,
    formula: Option<String>,
    error: bool,
    text: String,
    paragraphs: usize,
    paragraph_depth: Option<usize>,
    pending_space: bool,
}

impl Parser {
    fn new(limits: Limits) -> Self {
        Self {
            book: NativeOdsBook {
                sheets: Vec::new(),
                names: Vec::new(),
                name_bases: Vec::new(),
                date1904: false,
            },
            limits,
            styles: HashMap::new(),
            style: None,
            table_depth: None,
            row: None,
            cell: None,
            next_row: 1,
            next_column: 1,
            cells: 0,
            expanded_text: 0,
            spreadsheet_seen: false,
            null_date_seen: false,
            ignored_depth: None,
            sheet_names: HashSet::new(),
            defined_names: HashSet::new(),
        }
    }

    fn styles(&mut self, token: Token<'_>, ancestors: &[Name]) -> Result<()> {
        match token {
            Token::Start(e)
                if e.name.is(STYLE, "style") && e.get(STYLE, "family") == Some("table") =>
            {
                let name = e.required(STYLE, "name")?.to_owned();
                self.styles.insert(
                    name.clone(),
                    TableStyle {
                        parent: e.get(STYLE, "parent-style-name").map(str::to_owned),
                        display: None,
                    },
                );
                self.style = Some((ancestors.len(), name));
            }
            Token::Start(e) if e.name.is(STYLE, "table-properties") => {
                if let Some((depth, name)) = &self.style {
                    if ancestors.len() == depth + 1 {
                        self.styles.get_mut(name).unwrap().display =
                            optional_bool(e.get(TABLE, "display"))?;
                    }
                }
            }
            Token::End(_)
                if self
                    .style
                    .as_ref()
                    .is_some_and(|(depth, _)| *depth == ancestors.len()) =>
            {
                self.style = None
            }
            _ => {}
        }
        Ok(())
    }

    fn content(&mut self, token: Token<'_>, ancestors: &[Name]) -> Result<()> {
        if let Some(depth) = self.ignored_depth {
            if matches!(&token, Token::End(_)) && ancestors.len() == depth {
                self.ignored_depth = None;
            }
            return Ok(());
        }
        if let Token::Start(e) = &token {
            if self
                .cell
                .as_ref()
                .is_some_and(|cell| ancestors.len() > cell.depth)
                && (e.name.uri != TEXT
                    || matches!(
                        e.name.local.as_str(),
                        "note" | "ruby-text" | "tracked-changes"
                    ))
            {
                self.ignored_depth = Some(ancestors.len());
                return Ok(());
            }
        }
        // ODF automatic styles may override styles.xml definitions.
        match &token {
            Token::Start(e) => self.styles(Token::Start(e), ancestors)?,
            Token::End(n) => self.styles(Token::End(n), ancestors)?,
            Token::Text(_) => {}
        }
        match token {
            Token::Start(e) => {
                if ancestors.is_empty() && !e.name.is(OFFICE, "document-content") {
                    return Err(malformed("content.xml is not office:document-content"));
                }
                if e.name.is(OFFICE, "spreadsheet") {
                    if self.spreadsheet_seen
                        || !ancestors.last().is_some_and(|n| n.is(OFFICE, "body"))
                    {
                        return Err(malformed("invalid spreadsheet body"));
                    }
                    self.spreadsheet_seen = true;
                }
                if !ancestors.iter().any(|n| n.is(OFFICE, "spreadsheet")) {
                    return Ok(());
                }
                if e.name.is(TABLE, "table") {
                    if self.table_depth.is_some() {
                        return Err(malformed("nested spreadsheet table"));
                    }
                    let name = e.required(TABLE, "name")?.to_owned();
                    if name.is_empty() || !self.sheet_names.insert(name.clone()) {
                        return Err(malformed("empty or duplicate sheet name"));
                    }
                    self.book.sheets.push(Sheet {
                        name,
                        state: if optional_bool(e.get(TABLE, "display"))? == Some(false) {
                            SheetState::Hidden
                        } else {
                            SheetState::Visible
                        },
                        style: e.get(TABLE, "style-name").map(str::to_owned),
                        data: WorksheetData::default(),
                    });
                    self.table_depth = Some(ancestors.len());
                    self.next_row = 1;
                    self.next_column = 1;
                } else if e.name.is(TABLE, "table-column") {
                    if self.table_depth.is_none() || self.row.is_some() {
                        return Err(malformed("invalid column declaration nesting"));
                    }
                    let repeat = e.positive(TABLE, "number-columns-repeated")?;
                    checked_end(self.next_column, repeat)?;
                    self.next_column += u64::from(repeat);
                } else if e.name.is(TABLE, "table-row") {
                    if self.table_depth.is_none() || self.row.is_some() {
                        return Err(malformed("invalid row nesting"));
                    }
                    let repeat = e.positive(TABLE, "number-rows-repeated")?;
                    checked_end(self.next_row, repeat)?;
                    self.row = Some(Row {
                        depth: ancestors.len(),
                        repeat,
                        next_col: 1,
                        cells: Vec::new(),
                    });
                } else if e.name.is(TABLE, "table-cell") || e.name.is(TABLE, "covered-table-cell") {
                    let row = self
                        .row
                        .as_mut()
                        .ok_or_else(|| malformed("cell outside row"))?;
                    if self.cell.is_some() || ancestors.len() != row.depth + 1 {
                        return Err(malformed("invalid cell nesting"));
                    }
                    let repeat = e.positive(TABLE, "number-columns-repeated")?;
                    checked_end(row.next_col, repeat)?;
                    self.cell = Some(CellBuilder::new(
                        e,
                        ancestors.len(),
                        row.next_col as u32,
                        repeat,
                    )?);
                    row.next_col += u64::from(repeat);
                } else if e.name.is(TABLE, "null-date") {
                    if self.null_date_seen {
                        return Err(malformed("duplicate null-date"));
                    }
                    self.null_date_seen = true;
                    let date = e.required(TABLE, "date-value")?;
                    if !valid_date(date) {
                        return Err(malformed("invalid null-date"));
                    }
                    self.book.date1904 = date == "1904-01-01";
                } else if e.name.is(TABLE, "named-range") || e.name.is(TABLE, "named-expression") {
                    if !ancestors
                        .last()
                        .is_some_and(|n| n.is(TABLE, "named-expressions"))
                    {
                        return Err(malformed("name outside named-expressions"));
                    }
                    let local_sheet_id = self.table_depth.map(|_| self.book.sheets.len() - 1);
                    let name = e.required(TABLE, "name")?.to_owned();
                    if name.is_empty() || !self.defined_names.insert((local_sheet_id, name.clone()))
                    {
                        return Err(malformed("empty or duplicate defined name"));
                    }
                    let refers_to = if e.name.is(TABLE, "named-range") {
                        normalize_address_flags(e.required(TABLE, "cell-range-address")?, true)?
                    } else {
                        normalize_formula_flags(e.required(TABLE, "expression")?, true)?
                    };
                    let base = e
                        .get(TABLE, "base-cell-address")
                        .map(normalize_address)
                        .transpose()?;
                    self.book.name_bases.push(base);
                    self.book.names.push(NamedRange {
                        name,
                        scope: if local_sheet_id.is_some() {
                            "sheet"
                        } else {
                            "workbook"
                        }
                        .into(),
                        local_sheet_id,
                        refers_to,
                        ..NamedRange::default()
                    });
                }
                if let Some(cell) = &mut self.cell {
                    cell.start_text(e, ancestors.len(), self.limits)?;
                }
            }
            Token::Text(text) => {
                if let Some(cell) = &mut self.cell {
                    if cell.paragraph_depth.is_some() {
                        cell.append_text(text, self.limits)?;
                    }
                }
            }
            Token::End(name) => {
                if self
                    .cell
                    .as_ref()
                    .is_some_and(|c| c.depth == ancestors.len())
                {
                    let cell = self.cell.take().unwrap().finish()?;
                    if cell.cell.is_some() || cell.rows_spanned > 1 || cell.cols_spanned > 1 {
                        self.row.as_mut().unwrap().cells.push(cell);
                    }
                } else if let Some(cell) = &mut self.cell {
                    if cell.paragraph_depth == Some(ancestors.len()) {
                        cell.paragraph_depth = None;
                        cell.pending_space = false;
                    }
                }
                if name.is(TABLE, "table-row")
                    && self
                        .row
                        .as_ref()
                        .is_some_and(|r| r.depth == ancestors.len())
                {
                    self.finish_row()?;
                }
                if name.is(TABLE, "table") && self.table_depth == Some(ancestors.len()) {
                    self.table_depth = None;
                    let data = &mut self.book.sheets.last_mut().unwrap().data;
                    data.dimension = dimension(&data.cells);
                }
            }
        }
        Ok(())
    }

    fn finish_row(&mut self) -> Result<()> {
        let row = self.row.take().unwrap();
        let mut emitted = 0usize;
        let mut text = 0usize;
        let mut merges = 0usize;
        for template in &row.cells {
            if row.repeat > 1 && template.rows_spanned > 1 {
                return Err(malformed("repeated row contains vertical merge"));
            }
            if template.repeat > 1 && (template.rows_spanned > 1 || template.cols_spanned > 1) {
                return Err(malformed("repeated cell is merged"));
            }
            let count = (template.repeat as usize)
                .checked_mul(row.repeat as usize)
                .ok_or_else(|| malformed("repeat overflow"))?;
            if let Some(cell) = &template.cell {
                emitted = emitted
                    .checked_add(count)
                    .ok_or_else(|| malformed("cell count overflow"))?;
                let size = match &cell.value {
                    CellValue::String(s) | CellValue::Error(s) => s.len(),
                    _ => 0,
                };
                let size = size
                    .checked_add(cell.formula.as_ref().map_or(0, String::len))
                    .ok_or_else(|| malformed("text count overflow"))?;
                text = text
                    .checked_add(
                        size.checked_mul(count)
                            .ok_or_else(|| malformed("repeat text overflow"))?,
                    )
                    .ok_or_else(|| malformed("repeat text overflow"))?;
            }
            if template.rows_spanned > 1 || template.cols_spanned > 1 {
                merges = merges
                    .checked_add(count)
                    .ok_or_else(|| malformed("merge count overflow"))?;
            }
            checked_end(
                u64::from(template.col) + u64::from(template.repeat) - 1,
                template.cols_spanned,
            )?;
            checked_end(
                self.next_row + u64::from(row.repeat) - 1,
                template.rows_spanned,
            )?;
        }
        self.cells = self
            .cells
            .checked_add(emitted)
            .and_then(|n| n.checked_add(merges))
            .ok_or_else(|| malformed("expanded cell overflow"))?;
        self.expanded_text = self
            .expanded_text
            .checked_add(text)
            .ok_or_else(|| malformed("expanded text overflow"))?;
        if self.cells > self.limits.cells || self.expanded_text > self.limits.bytes {
            return Err(malformed("expanded cell/text budget exceeded"));
        }
        // Empty rows and empty/covered repeats only advance checked coordinates.
        if !row.cells.is_empty() {
            let data = &mut self.book.sheets.last_mut().unwrap().data;
            for dy in 0..row.repeat {
                let r = (self.next_row + u64::from(dy)) as u32;
                for template in &row.cells {
                    for dx in 0..template.repeat {
                        let c = template.col + dx;
                        if let Some(source) = &template.cell {
                            let mut cell = source.clone();
                            cell.row = r;
                            cell.col = c;
                            cell.coordinate = a1(r, c);
                            data.cells.push(cell);
                        }
                        if template.rows_spanned > 1 || template.cols_spanned > 1 {
                            data.merged_ranges.push(format!(
                                "{}:{}",
                                a1(r, c),
                                a1(
                                    r + (template.rows_spanned - 1),
                                    c + (template.cols_spanned - 1)
                                )
                            ));
                        }
                    }
                }
            }
        }
        self.next_row += u64::from(row.repeat);
        Ok(())
    }

    fn finish(mut self) -> Result<NativeOdsBook> {
        if !self.spreadsheet_seen
            || self.table_depth.is_some()
            || self.row.is_some()
            || self.cell.is_some()
        {
            return Err(malformed("incomplete spreadsheet"));
        }
        // Resolve style ancestry once per style, not once per sheet. A long
        // style chain reused by many sheets must remain linear in input size.
        let mut resolved = HashMap::<&str, bool>::new();
        for sheet in &mut self.book.sheets {
            let mut current = sheet.style.as_deref();
            let mut path = Vec::new();
            let mut seen = HashSet::new();
            let display = loop {
                let Some(name) = current else {
                    break true;
                };
                if let Some(display) = resolved.get(name) {
                    break *display;
                }
                if !seen.insert(name) {
                    return Err(malformed("cyclic table style inheritance"));
                }
                path.push(name);
                let Some(style) = self.styles.get(name) else {
                    break true;
                };
                if let Some(display) = style.display {
                    break display;
                }
                current = style.parent.as_deref();
            };
            for name in path {
                resolved.insert(name, display);
            }
            if !display {
                sheet.state = SheetState::Hidden;
            }
        }
        Ok(self.book)
    }
}

impl CellBuilder {
    fn new(e: &Element, depth: usize, col: u32, repeat: u32) -> Result<Self> {
        let kind = e.get(OFFICE, "value-type").map(str::to_owned);
        let value_attribute = match kind.as_deref() {
            Some("float" | "percentage" | "currency") => "value",
            Some("boolean") => "boolean-value",
            Some("date") => "date-value",
            Some("time") => "time-value",
            _ => "string-value",
        };
        Ok(Self {
            depth,
            col,
            repeat,
            rows_spanned: e.positive(TABLE, "number-rows-spanned")?,
            cols_spanned: e.positive(TABLE, "number-columns-spanned")?,
            kind,
            scalar: e.get(OFFICE, value_attribute).map(str::to_owned),
            formula: e.get(TABLE, "formula").map(normalize_formula).transpose()?,
            error: e.get(
                "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
                "value-type",
            ) == Some("error"),
            text: String::new(),
            paragraphs: 0,
            paragraph_depth: None,
            pending_space: false,
        })
    }

    fn start_text(&mut self, e: &Element, depth: usize, limits: Limits) -> Result<()> {
        if e.name.is(TEXT, "p") || e.name.is(TEXT, "h") {
            if self.paragraph_depth.is_some() {
                return Err(malformed("nested text paragraph"));
            }
            if self.paragraphs > 0 {
                self.text.push('\n');
            }
            self.paragraphs += 1;
            self.paragraph_depth = Some(depth);
            self.pending_space = false;
        } else if self.paragraph_depth.is_some() {
            if e.name.is(TEXT, "s") {
                self.flush_space();
                let count = e.positive(TEXT, "c")? as usize;
                if self
                    .text
                    .len()
                    .checked_add(count)
                    .is_none_or(|n| n > limits.text)
                {
                    return Err(malformed("text space repeat budget exceeded"));
                }
                self.text.extend(std::iter::repeat_n(' ', count));
            } else if e.name.is(TEXT, "tab") {
                self.flush_space();
                self.text.push('\t');
            } else if e.name.is(TEXT, "line-break") {
                self.pending_space = false;
                self.text.push('\n');
            }
        }
        if self.text.len() > limits.text {
            return Err(malformed("cell text budget exceeded"));
        }
        Ok(())
    }

    fn flush_space(&mut self) {
        if self.pending_space {
            self.text.push(' ');
            self.pending_space = false;
        }
    }

    fn append_text(&mut self, text: &str, limits: Limits) -> Result<()> {
        for ch in text.chars() {
            if matches!(ch, ' ' | '\n' | '\r' | '\t') {
                if !self.text.is_empty() && !self.text.ends_with('\n') {
                    self.pending_space = true;
                }
            } else {
                self.flush_space();
                self.text.push(ch);
            }
            if self.text.len() > limits.text {
                return Err(malformed("cell text budget exceeded"));
            }
        }
        Ok(())
    }

    fn finish(self) -> Result<CellTemplate> {
        let kind = self.kind.as_deref();
        let string_type = if self.formula.is_some() {
            CellDataType::FormulaString
        } else {
            CellDataType::InlineString
        };
        let (data_type, value) = if self.error || kind == Some("error") {
            (
                CellDataType::Error,
                CellValue::Error(if self.text.is_empty() {
                    self.scalar.unwrap_or_default()
                } else {
                    self.text
                }),
            )
        } else {
            match kind {
                Some("float" | "percentage" | "currency") => {
                    let raw = self
                        .scalar
                        .as_deref()
                        .ok_or_else(|| malformed("missing numeric value"))?
                        .trim();
                    let number = match raw {
                        "INF" => f64::INFINITY,
                        "-INF" => f64::NEG_INFINITY,
                        "NaN" => f64::NAN,
                        _ if valid_double(raw) => raw
                            .parse::<f64>()
                            .map_err(|_| malformed("invalid numeric value"))?,
                        _ => return Err(malformed("invalid numeric value")),
                    };
                    (CellDataType::Number, CellValue::Number(number))
                }
                Some("boolean") => (
                    CellDataType::Bool,
                    CellValue::Bool(
                        optional_bool(self.scalar.as_deref())?
                            .ok_or_else(|| malformed("missing boolean value"))?,
                    ),
                ),
                Some("date") => {
                    let mut value = self.scalar.ok_or_else(|| malformed("missing date value"))?;
                    if value.trim().len() != value.len() {
                        value = value.trim().to_owned();
                    }
                    if !valid_date(&value) {
                        return Err(malformed("invalid date value"));
                    }
                    (CellDataType::DateIso, CellValue::String(value))
                }
                Some("time") => {
                    let mut value = self.scalar.ok_or_else(|| malformed("missing time value"))?;
                    if value.trim().len() != value.len() {
                        value = value.trim().to_owned();
                    }
                    if !valid_duration(&value) {
                        return Err(malformed("invalid duration value"));
                    }
                    (CellDataType::DateIso, CellValue::String(value))
                }
                Some("string") => (
                    string_type,
                    CellValue::String(self.scalar.unwrap_or(self.text)),
                ),
                None if self.paragraphs > 0 || self.scalar.is_some() => (
                    string_type,
                    CellValue::String(self.scalar.unwrap_or(self.text)),
                ),
                None | Some("void") => (CellDataType::Number, CellValue::Empty),
                Some(other) => return Err(malformed(format!("unknown value type {other}"))),
            }
        };
        let cell = if value == CellValue::Empty && self.formula.is_none() {
            None
        } else {
            Some(Cell {
                row: 0,
                col: self.col,
                coordinate: String::new(),
                style_id: None,
                data_type,
                value,
                formula: self.formula,
                formula_kind: None,
                formula_shared_index: None,
                array_formula: None,
                rich_text: None,
            })
        };
        Ok(CellTemplate {
            col: self.col,
            repeat: self.repeat,
            rows_spanned: self.rows_spanned,
            cols_spanned: self.cols_spanned,
            cell,
        })
    }
}

fn optional_bool(value: Option<&str>) -> Result<Option<bool>> {
    match value.map(str::trim) {
        None => Ok(None),
        Some("true" | "1") => Ok(Some(true)),
        Some("false" | "0") => Ok(Some(false)),
        _ => Err(malformed("invalid boolean")),
    }
}

fn valid_double(raw: &str) -> bool {
    let b = raw.as_bytes();
    let mut i = usize::from(b.first().is_some_and(|b| matches!(b, b'+' | b'-')));
    let mut digits = 0;
    while b.get(i).is_some_and(u8::is_ascii_digit) {
        i += 1;
        digits += 1;
    }
    if b.get(i) == Some(&b'.') {
        i += 1;
        while b.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
            digits += 1;
        }
    }
    if digits == 0 {
        return false;
    }
    if b.get(i).is_some_and(|b| matches!(b, b'e' | b'E')) {
        i += 1;
        if b.get(i).is_some_and(|b| matches!(b, b'+' | b'-')) {
            i += 1;
        }
        let start = i;
        while b.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        if i == start {
            return false;
        }
    }
    i == b.len()
}

fn checked_end(start: u64, count: u32) -> Result<u32> {
    start
        .checked_add(u64::from(count) - 1)
        .filter(|n| *n <= u64::from(u32::MAX))
        .map(|n| n as u32)
        .ok_or_else(|| malformed("row/column coordinate overflow"))
}

fn a1(row: u32, mut col: u32) -> String {
    let mut bytes = [0u8; 7];
    let mut index = bytes.len();
    while col != 0 {
        index -= 1;
        bytes[index] = b'A' + ((col - 1) % 26) as u8;
        col = (col - 1) / 26;
    }
    format!("{}{row}", std::str::from_utf8(&bytes[index..]).unwrap())
}

fn dimension(cells: &[Cell]) -> Option<String> {
    let first = cells.first()?;
    let (mut min_r, mut min_c, mut max_r, mut max_c) = (first.row, first.col, first.row, first.col);
    for cell in cells.iter().skip(1) {
        min_r = min_r.min(cell.row);
        min_c = min_c.min(cell.col);
        max_r = max_r.max(cell.row);
        max_c = max_c.max(cell.col);
    }
    Some(format!("{}:{}", a1(min_r, min_c), a1(max_r, max_c)))
}

// ODF 1.3 Part 3 9.2.1 and Part 4 5/6. Names and string literals retain
// their spelling; only the formula envelope, reference grammar and argument
// separators are normalized. This is not a formula evaluator.
fn normalize_formula(raw: &str) -> Result<String> {
    normalize_formula_flags(raw, false)
}

fn normalize_formula_flags(raw: &str, preserve_sheet_absolute: bool) -> Result<String> {
    let raw = raw.trim();
    let raw = if let Some(formula) = raw.strip_prefix('=') {
        formula
    } else if let Some((prefix, formula)) = raw.split_once(":=") {
        if prefix.is_empty()
            || prefix
                .chars()
                .any(|c| !(c.is_alphanumeric() || c == '_' || c == '-' || c == '.'))
        {
            return Err(malformed("invalid formula prefix"));
        }
        formula
    } else {
        raw
    };
    let mut result = String::with_capacity(raw.len());
    let mut chars = raw.chars().peekable();
    let mut quoted = false;
    let mut array_depth = 0usize;
    while let Some(ch) = chars.next() {
        match ch {
            '"' => {
                result.push(ch);
                if quoted && chars.peek() == Some(&'"') {
                    result.push(chars.next().unwrap());
                } else {
                    quoted = !quoted;
                }
            }
            '[' if !quoted => {
                let mut reference = String::new();
                let mut sheet_quote = false;
                let mut closed = false;
                while let Some(c) = chars.next() {
                    if c == '\'' {
                        sheet_quote = !sheet_quote;
                    }
                    if c == ']' && !sheet_quote {
                        closed = true;
                        break;
                    }
                    reference.push(c);
                }
                if !closed {
                    return Err(malformed("unterminated formula reference"));
                }
                result.push_str(&normalize_address_flags(
                    &reference,
                    preserve_sheet_absolute,
                )?);
            }
            ';' if !quoted => result.push(','),
            '{' if !quoted => {
                array_depth += 1;
                result.push(ch);
            }
            '}' if !quoted => {
                array_depth = array_depth.saturating_sub(1);
                result.push(ch);
            }
            '|' if !quoted && array_depth > 0 => result.push(';'),
            _ => result.push(ch),
        }
    }
    if quoted {
        return Err(malformed("unterminated formula string"));
    }
    Ok(result)
}

fn normalize_address(raw: &str) -> Result<String> {
    normalize_address_flags(raw, false)
}

fn normalize_address_flags(raw: &str, preserve_sheet_absolute: bool) -> Result<String> {
    let mut result = String::with_capacity(raw.len());
    let mut part = String::new();
    let mut quoted = false;
    for ch in raw.chars().chain(std::iter::once(':')) {
        if ch == '\'' {
            quoted = !quoted;
        }
        if ch == ':' && !quoted {
            if part.is_empty() {
                return Err(malformed("empty range endpoint"));
            }
            let mut dot = None;
            let mut q = false;
            for (i, c) in part.char_indices() {
                if c == '\'' {
                    q = !q;
                } else if c == '.' && !q {
                    dot = Some(i);
                }
            }
            if !result.is_empty() {
                result.push(':');
            }
            if let Some(dot) = dot {
                let sheet = if preserve_sheet_absolute {
                    &part[..dot]
                } else {
                    part[..dot].strip_prefix('$').unwrap_or(&part[..dot])
                };
                if !sheet.is_empty() {
                    result.push_str(sheet);
                    result.push('!');
                }
                result.push_str(&part[dot + 1..]);
            } else {
                result.push_str(&part);
            }
            part.clear();
        } else {
            part.push(ch);
        }
    }
    if quoted {
        return Err(malformed("unterminated sheet name"));
    }
    Ok(result)
}

fn valid_date(raw: &str) -> bool {
    let raw = raw.strip_prefix('-').unwrap_or(raw);
    let Some((year, rest)) = raw.split_once('-') else {
        return false;
    };
    if year.len() < 4 || !year.bytes().all(|b| b.is_ascii_digit()) {
        return false;
    }
    if year.bytes().all(|b| b == b'0') || (year.len() > 4 && year.starts_with('0')) {
        return false;
    }
    // Arbitrarily long XML Schema years remain lossless strings. Only the
    // Gregorian leap-year residue is needed, never an overflowing year integer.
    let year = year.bytes().fold(0u16, |residue, b| {
        (residue * 10 + u16::from(b - b'0')) % 400
    });
    let bytes = rest.as_bytes();
    if bytes.len() < 5
        || bytes[2] != b'-'
        || !bytes[..2]
            .iter()
            .chain(&bytes[3..5])
            .all(u8::is_ascii_digit)
    {
        return false;
    }
    let month = (bytes[0] - b'0') * 10 + bytes[1] - b'0';
    let day = (bytes[3] - b'0') * 10 + bytes[4] - b'0';
    let days = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        _ => return false,
    };
    if day == 0 || day > days {
        return false;
    }
    let rest = &rest[5..];
    if let Some(time) = rest.strip_prefix('T') {
        valid_time(time)
    } else {
        valid_zone(rest)
    }
}

fn valid_zone(zone: &str) -> bool {
    if zone.is_empty() || zone == "Z" {
        return true;
    }
    let b = zone.as_bytes();
    if b.len() != 6
        || !matches!(b[0], b'+' | b'-')
        || b[3] != b':'
        || ![b[1], b[2], b[4], b[5]].iter().all(u8::is_ascii_digit)
    {
        return false;
    }
    let h = (b[1] - b'0') * 10 + b[2] - b'0';
    let m = (b[4] - b'0') * 10 + b[5] - b'0';
    h <= 14 && m < 60 && (h != 14 || m == 0)
}

fn valid_time(raw: &str) -> bool {
    let b = raw.as_bytes();
    if b.len() < 8
        || b[2] != b':'
        || b[5] != b':'
        || ![b[0], b[1], b[3], b[4], b[6], b[7]]
            .iter()
            .all(u8::is_ascii_digit)
    {
        return false;
    }
    let h = (b[0] - b'0') * 10 + b[1] - b'0';
    let m = (b[3] - b'0') * 10 + b[4] - b'0';
    let s = (b[6] - b'0') * 10 + b[7] - b'0';
    if h > 24 || m > 59 || s > 59 || (h == 24 && (m != 0 || s != 0)) {
        return false;
    }
    let mut end = 8;
    if b.get(end) == Some(&b'.') {
        end += 1;
        let start = end;
        while b.get(end).is_some_and(u8::is_ascii_digit) {
            if h == 24 && b[end] != b'0' {
                return false;
            }
            end += 1;
        }
        if end == start {
            return false;
        }
    }
    valid_zone(&raw[end..])
}

fn valid_duration(raw: &str) -> bool {
    let Some(raw) = raw.strip_prefix('-').unwrap_or(raw).strip_prefix('P') else {
        return false;
    };
    let mut time = false;
    let mut last = 0;
    let mut count = 0;
    let mut time_count = 0;
    let mut number = String::new();
    for ch in raw.chars() {
        if ch == 'T' && !time && number.is_empty() {
            time = true;
            last = 0;
            continue;
        }
        if ch.is_ascii_digit() || ch == '.' {
            number.push(ch);
            continue;
        }
        let rank = match (time, ch) {
            (false, 'Y') | (true, 'H') => 1,
            (false, 'M') | (true, 'M') => 2,
            (false, 'D') | (true, 'S') => 3,
            _ => return false,
        };
        if rank <= last || number.is_empty() {
            return false;
        }
        if let Some((whole, fraction)) = number.split_once('.') {
            if ch != 'S'
                || !time
                || whole.is_empty()
                || fraction.is_empty()
                || !whole
                    .bytes()
                    .chain(fraction.bytes())
                    .all(|b| b.is_ascii_digit())
            {
                return false;
            }
        } else if !number.bytes().all(|b| b.is_ascii_digit()) {
            return false;
        }
        last = rank;
        count += 1;
        if time {
            time_count += 1;
        }
        number.clear();
    }
    count > 0 && number.is_empty() && (!time || time_count > 0)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fixed_formula_and_iso_expectations() {
        assert_eq!(
            normalize_formula("of:=SUM([.$A$1:.B2];\"a;b\";['Two words'.C3])").unwrap(),
            "SUM($A$1:B2,\"a;b\",'Two words'!C3)"
        );
        assert_eq!(
            normalize_address("$'Two words'.$A$1:.$B$9").unwrap(),
            "'Two words'!$A$1:$B$9"
        );
        for s in [
            "2024-02-29",
            "2024-02-29T23:59:59.125+05:30",
            "1904-01-01Z",
            "2001-01-01T24:00:00",
        ] {
            assert!(valid_date(s), "{s}");
        }
        for s in [
            "2023-02-29",
            "2024-13-01",
            "2024-01-00",
            "2024-01-01T25:00:00",
            "2024-01-01+14:01",
        ] {
            assert!(!valid_date(s), "{s}");
        }
        for s in ["PT0S", "-P2DT3H4M5.25S", "P1Y2M3DT4H5M6S"] {
            assert!(valid_duration(s), "{s}");
        }
        for s in ["P", "PT", "P1DT", "PT1.5H", "PT1S2M", "PT1..2S"] {
            assert!(!valid_duration(s), "{s}");
        }
    }
}
