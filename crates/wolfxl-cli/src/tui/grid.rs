//! A sheet prepared for the interactive grid: each cell's displayed text,
//! spelled the way the box preview spells it, plus its numeric value for
//! sorting.

use unicode_width::UnicodeWidthStr;
use wolfxl_core::Sheet;

use crate::filter::cell_number;
use crate::input::Provenance;
use crate::render::display_sheet_cell;

/// Data rows measured when sizing columns, so a long sheet opens without a
/// full width pass.
const WIDTH_SAMPLE: usize = 1000;

/// One cell as the grid shows and sorts it.
#[derive(Clone, Debug, PartialEq)]
pub(crate) struct GridCell {
    /// Displayed text, with control characters (line breaks, tabs) shown as
    /// spaces so one cell stays on one screen line.
    pub(crate) text: String,
    /// Numbers as stored, and text that reads as a number, money spellings
    /// included. `None` for everything else.
    pub(crate) number: Option<f64>,
}

impl GridCell {
    fn blank() -> Self {
        Self {
            text: String::new(),
            number: None,
        }
    }

    /// No visible text.
    pub(crate) fn is_blank(&self) -> bool {
        self.text.trim().is_empty()
    }
}

/// A sheet's header row and data rows, every row as wide as the header.
#[derive(Debug)]
pub(crate) struct Grid {
    pub(crate) name: String,
    pub(crate) headers: Vec<String>,
    pub(crate) rows: Vec<Vec<GridCell>>,
    /// Column widths in terminal cells: the widest header or sampled value,
    /// between 1 and the width cap.
    pub(crate) widths: Vec<usize>,
    /// Shown on the status line, for example why filters do not apply.
    pub(crate) note: Option<String>,
    /// Original sheet coordinates, including the header at row index zero.
    pub(crate) provenance: Provenance,
}

impl Grid {
    /// A source-order sheet, used by grid tests.
    #[cfg(test)]
    pub(crate) fn new(sheet: &Sheet, max_width: usize, note: Option<String>) -> Self {
        Self::with_provenance(sheet, max_width, note, &Provenance::of(sheet))
    }

    /// Row 0 is the chosen header; provenance survives shaping and sorting.
    pub(crate) fn with_provenance(
        sheet: &Sheet,
        max_width: usize,
        note: Option<String>,
        provenance: &Provenance,
    ) -> Self {
        let columns = sheet.rows().iter().map(Vec::len).max().unwrap_or(0);
        let mut headers: Vec<String> = sheet
            .headers()
            .iter()
            .map(String::as_str)
            .map(one_line)
            .collect();
        headers.resize(columns, String::new());
        let rows: Vec<Vec<GridCell>> = sheet
            .rows()
            .iter()
            .skip(1)
            .map(|row| {
                (0..columns)
                    .map(|column| match row.get(column) {
                        Some(cell) => GridCell {
                            text: one_line(&display_sheet_cell(sheet, cell, true)),
                            number: cell_number(&cell.value),
                        },
                        None => GridCell::blank(),
                    })
                    .collect()
            })
            .collect();
        let widths = (0..columns)
            .map(|column| {
                let header = headers.get(column).map_or(0, |header| header.width());
                rows.iter()
                    .take(WIDTH_SAMPLE)
                    .filter_map(|row| row.get(column))
                    .map(|cell| cell.text.width())
                    .fold(header, usize::max)
                    .clamp(1, max_width.max(1))
            })
            .collect();
        Self {
            name: sheet.name.clone(),
            headers,
            rows,
            widths,
            note,
            provenance: provenance.clone(),
        }
    }

    /// Stand-in for a sheet that failed to load, so Tab can move past it.
    pub(crate) fn unavailable(name: &str, note: String) -> Self {
        Self {
            name: name.to_owned(),
            headers: Vec::new(),
            rows: Vec::new(),
            widths: Vec::new(),
            note: Some(note),
            provenance: Provenance {
                rows: Vec::new(),
                columns: Vec::new(),
            },
        }
    }
}

/// `text` with every control character replaced by a space.
pub(super) fn one_line(text: &str) -> String {
    if text.contains(char::is_control) {
        text.replace(char::is_control, " ")
    } else {
        text.to_owned()
    }
}

#[cfg(test)]
mod tests {
    use wolfxl_core::{Cell, CellValue};

    use super::*;

    fn cell(value: CellValue) -> Cell {
        Cell {
            value,
            number_format: None,
        }
    }

    #[test]
    fn cells_keep_display_text_and_numeric_value() {
        let sheet = Sheet::from_rows(
            "Data".into(),
            vec![
                vec![
                    cell(CellValue::String("Name".into())),
                    cell(CellValue::String("Amount".into())),
                ],
                vec![
                    cell(CellValue::String("line\nbreak".into())),
                    cell(CellValue::Int(1200)),
                ],
                vec![cell(CellValue::String("$1,200.00".into()))],
            ],
        );
        let grid = Grid::new(&sheet, 5, None);
        assert_eq!(grid.headers, ["Name", "Amount"]);
        assert_eq!(grid.rows[0][0].text, "line break");
        assert_eq!(grid.rows[0][1].text, "1,200");
        assert_eq!(grid.rows[0][1].number, Some(1200.0));
        assert_eq!(grid.rows[1][0].number, Some(1200.0));
        assert!(grid.rows[1][1].is_blank(), "short rows are padded");
        assert_eq!(grid.widths, [5, 5], "widths are capped");
    }
}
