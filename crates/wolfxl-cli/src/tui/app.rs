//! Interactive grid state: cursor, scrolling, search, sort, value
//! frequencies, and sheet cycling. Keys go in through [`App::handle_key`];
//! the event loop acts on the returned [`Action`] and drawing reads the
//! state, so every behavior here runs without a terminal.

use std::borrow::Cow;
use std::cmp::Ordering;
use std::collections::HashMap;
use std::mem;
use std::path::Path;

use ratatui::crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
use unicode_width::UnicodeWidthStr;

use super::draw::{elide, fit};
use super::grid::{one_line, Grid};

/// What the event loop does after a key.
#[derive(Debug, PartialEq, Eq)]
pub(crate) enum Action {
    /// Redraw and wait for the next key.
    Continue,
    /// Restore the terminal and exit.
    Quit,
    /// Load sheet `index`, then pass it to [`App::insert`].
    Load(usize),
}

/// Rows ordered by one column.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) struct Sort {
    pub(super) column: usize,
    pub(super) descending: bool,
    /// Every non-blank value in the column is a number.
    pub(super) numeric: bool,
}

/// One sheet's grid and where the user is in it.
pub(super) struct View {
    pub(super) grid: Grid,
    /// Display position to data row: the source order, or the sorted order.
    pub(super) order: Vec<usize>,
    pub(super) sort: Option<Sort>,
    /// Cursor, as a display position and a column.
    pub(super) row: usize,
    pub(super) col: usize,
    /// First display position and first column on screen.
    pub(super) top: usize,
    pub(super) left: usize,
}

impl View {
    fn new(grid: Grid) -> Self {
        let order = (0..grid.rows.len()).collect();
        Self {
            grid,
            order,
            sort: None,
            row: 0,
            col: 0,
            top: 0,
            left: 0,
        }
    }
}

/// Distinct values of one column with their row counts, most frequent
/// first, ties in text order.
#[derive(Debug, PartialEq, Eq)]
pub(super) struct Frequency {
    pub(super) header: String,
    pub(super) counts: Vec<(String, usize)>,
    pub(super) rows: usize,
}

/// What keys currently do.
pub(super) enum Mode {
    /// Move, sort, search, and switch sheets.
    Browse,
    /// Typing a search query after `/`.
    Search(String),
    /// Showing a frequency popup; any key closes it.
    Frequency(Frequency),
}

pub(crate) struct App {
    pub(super) file: String,
    pub(super) names: Vec<String>,
    /// Index of the sheet in `view`.
    pub(super) current: usize,
    pub(super) view: View,
    /// Sheets visited before and not shown now, by sheet index.
    parked: Vec<Option<View>>,
    /// The committed search, as typed.
    pub(super) query: Option<String>,
    /// Cells containing the query in the current view, as display
    /// `(position, column)` in row-major order.
    pub(super) matches: Vec<(usize, usize)>,
    pub(super) mode: Mode,
    /// One-key notice on the status line.
    pub(super) message: Option<String>,
    /// Screen columns available to the grid and the status line.
    width: usize,
    /// Data rows on one screen.
    page: usize,
}

impl App {
    /// `grid` is sheet `current` of `names`, loaded from `file`.
    pub(crate) fn new(file: String, names: Vec<String>, current: usize, grid: Grid) -> Self {
        let parked = names.iter().map(|_| None).collect();
        Self {
            file,
            names,
            current,
            view: View::new(grid),
            parked,
            query: None,
            matches: Vec::new(),
            mode: Mode::Browse,
            message: None,
            width: usize::from(u16::MAX),
            page: 1,
        }
    }

    /// Name of sheet `index`.
    pub(crate) fn sheet_name(&self, index: usize) -> &str {
        self.names.get(index).map_or("", String::as_str)
    }

    /// Fit scrolling to a `width` x `height` screen: one header line, data
    /// lines, one status line.
    pub(crate) fn set_viewport(&mut self, width: u16, height: u16) {
        self.width = usize::from(width);
        self.page = usize::from(height.saturating_sub(2)).max(1);
        self.scroll_into_view();
    }

    /// Show sheet `index`, freshly loaded as `grid`, after [`Action::Load`].
    pub(crate) fn insert(&mut self, index: usize, grid: Grid) {
        self.show(index, View::new(grid));
    }

    pub(crate) fn handle_key(&mut self, key: KeyEvent) -> Action {
        if !key.is_press() {
            return Action::Continue;
        }
        self.message = None;
        if key.code == KeyCode::Char('c') && key.modifiers.contains(KeyModifiers::CONTROL) {
            return Action::Quit;
        }
        let action = match mem::replace(&mut self.mode, Mode::Browse) {
            Mode::Frequency(_) => Action::Continue,
            Mode::Search(query) => {
                self.search_key(key.code, query);
                Action::Continue
            }
            Mode::Browse => self.browse_key(key.code),
        };
        self.scroll_into_view();
        action
    }

    fn browse_key(&mut self, code: KeyCode) -> Action {
        match code {
            KeyCode::Char('q') | KeyCode::Esc => return Action::Quit,
            KeyCode::Up => self.up(1),
            KeyCode::Down => self.down(1),
            KeyCode::PageUp => self.up(self.page),
            KeyCode::PageDown => self.down(self.page),
            KeyCode::Left => self.view.col = self.view.col.saturating_sub(1),
            KeyCode::Right => self.view.col = (self.view.col + 1).min(self.last_col()),
            KeyCode::Home => self.view.col = 0,
            KeyCode::End => self.view.col = self.last_col(),
            KeyCode::Char('g') => self.view.row = 0,
            KeyCode::Char('G') => self.view.row = self.last_row(),
            KeyCode::Char('/') => self.mode = Mode::Search(String::new()),
            KeyCode::Char('n') => self.next_match(),
            KeyCode::Char('N') => self.previous_match(),
            KeyCode::Char('s') => self.sort(),
            KeyCode::Char('S') => self.restore_order(),
            KeyCode::Char('f') => self.frequency(),
            KeyCode::Tab => return self.cycle(true),
            KeyCode::BackTab => return self.cycle(false),
            _ => {}
        }
        Action::Continue
    }

    fn search_key(&mut self, code: KeyCode, mut query: String) {
        match code {
            KeyCode::Esc => {}
            KeyCode::Enter => self.commit_search(query),
            KeyCode::Backspace => {
                query.pop();
                self.mode = Mode::Search(query);
            }
            KeyCode::Char(typed) => {
                query.push(typed);
                self.mode = Mode::Search(query);
            }
            _ => self.mode = Mode::Search(query),
        }
    }

    fn last_row(&self) -> usize {
        self.view.order.len().saturating_sub(1)
    }

    fn last_col(&self) -> usize {
        self.view.grid.headers.len().saturating_sub(1)
    }

    fn up(&mut self, rows: usize) {
        self.view.row = self.view.row.saturating_sub(rows);
    }

    fn down(&mut self, rows: usize) {
        self.view.row = self.view.row.saturating_add(rows).min(self.last_row());
    }

    /// Scroll so the cursor is on screen, moving as little as possible.
    fn scroll_into_view(&mut self) {
        let view = &mut self.view;
        if view.row < view.top {
            view.top = view.row;
        } else if view.row >= view.top + self.page {
            view.top = view.row + 1 - self.page;
        }
        if view.col < view.left {
            view.left = view.col;
        }
        while view.left < view.col && span(&view.grid.widths, view.left, view.col) > self.width {
            view.left += 1;
        }
    }

    /// Switch to `index`, or ask the event loop to load it first.
    fn cycle(&mut self, forward: bool) -> Action {
        let count = self.names.len();
        if count < 2 {
            self.message = Some("only one sheet".into());
            return Action::Continue;
        }
        let index = if forward {
            (self.current + 1) % count
        } else {
            (self.current + count - 1) % count
        };
        match self.parked.get_mut(index).and_then(Option::take) {
            Some(view) => {
                self.show(index, view);
                Action::Continue
            }
            None => Action::Load(index),
        }
    }

    fn show(&mut self, index: usize, view: View) {
        let previous = mem::replace(&mut self.view, view);
        if let Some(slot) = self.parked.get_mut(self.current) {
            *slot = Some(previous);
        }
        self.current = index;
        self.refresh_matches();
        self.scroll_into_view();
    }

    fn commit_search(&mut self, query: String) {
        if query.is_empty() {
            self.query = None;
            self.matches.clear();
            return;
        }
        self.query = Some(query);
        self.refresh_matches();
        self.next_match();
    }

    /// Recompute `matches` for the current view and order.
    fn refresh_matches(&mut self) {
        self.matches.clear();
        let Some(query) = &self.query else {
            return;
        };
        let needle = query.to_lowercase();
        for (position, &source) in self.view.order.iter().enumerate() {
            let Some(row) = self.view.grid.rows.get(source) else {
                continue;
            };
            for (column, cell) in row.iter().enumerate() {
                if cell.text.to_lowercase().contains(&needle) {
                    self.matches.push((position, column));
                }
            }
        }
    }

    /// Move to the first match after the cursor, wrapping to the top.
    fn next_match(&mut self) {
        if !self.has_matches() {
            return;
        }
        let here = (self.view.row, self.view.col);
        let mut index = self.matches.partition_point(|&found| found <= here);
        if index == self.matches.len() {
            index = 0;
            self.message = Some("search wrapped to the top".into());
        }
        self.jump(index);
    }

    /// Move to the last match before the cursor, wrapping to the bottom.
    fn previous_match(&mut self) {
        if !self.has_matches() {
            return;
        }
        let here = (self.view.row, self.view.col);
        let index = match self.matches.partition_point(|&found| found < here) {
            0 => {
                self.message = Some("search wrapped to the bottom".into());
                self.matches.len() - 1
            }
            after => after - 1,
        };
        self.jump(index);
    }

    /// Whether `n`/`N` have somewhere to go; sets the reason when not.
    fn has_matches(&mut self) -> bool {
        match &self.query {
            None => {
                self.message = Some("no search yet; press / to search".into());
                false
            }
            Some(query) if self.matches.is_empty() => {
                self.message = Some(format!("no match for /{query}"));
                false
            }
            Some(_) => true,
        }
    }

    fn jump(&mut self, index: usize) {
        if let Some(&(row, col)) = self.matches.get(index) {
            self.view.row = row;
            self.view.col = col;
        }
    }

    /// Sort by the cursor's column: ascending, or the reverse when the view
    /// is already ascending by it. Blanks stay last either way, and equal
    /// values keep their source order.
    fn sort(&mut self) {
        let column = self.view.col;
        let view = &mut self.view;
        let Some(header) = view.grid.headers.get(column) else {
            self.message = Some("no column to sort".into());
            return;
        };
        let mut non_blank = view
            .grid
            .rows
            .iter()
            .filter_map(|row| row.get(column))
            .filter(|cell| !cell.is_blank())
            .peekable();
        let numeric = non_blank.peek().is_some() && non_blank.all(|cell| cell.number.is_some());
        let keys: Vec<SortKey> = view
            .grid
            .rows
            .iter()
            .map(|row| match row.get(column) {
                Some(cell) if !cell.is_blank() => match cell.number {
                    Some(number) if numeric => SortKey::Number(number),
                    _ => SortKey::Text(cell.text.to_lowercase()),
                },
                _ => SortKey::Blank,
            })
            .collect();
        let descending =
            matches!(view.sort, Some(sort) if sort.column == column && !sort.descending);
        let mut order: Vec<usize> = (0..keys.len()).collect();
        order.sort_by(|&a, &b| compare(&keys[a], &keys[b], descending));
        view.order = order;
        view.sort = Some(Sort {
            column,
            descending,
            numeric,
        });
        self.message = Some(format!(
            "sorted by {header}, {}",
            if descending {
                "descending"
            } else {
                "ascending"
            }
        ));
        self.refresh_matches();
    }

    fn restore_order(&mut self) {
        let view = &mut self.view;
        view.order = (0..view.grid.rows.len()).collect();
        view.sort = None;
        self.message = Some("source order".into());
        self.refresh_matches();
    }

    fn frequency(&mut self) {
        let column = self.view.col;
        let grid = &self.view.grid;
        let Some(header) = grid.headers.get(column) else {
            self.message = Some("no column to count".into());
            return;
        };
        let mut tally: HashMap<&str, usize> = HashMap::new();
        for row in &grid.rows {
            let text = row.get(column).map_or("", |cell| cell.text.as_str());
            *tally.entry(text).or_default() += 1;
        }
        let mut counts: Vec<(String, usize)> = tally
            .into_iter()
            .map(|(text, count)| (text.to_owned(), count))
            .collect();
        counts.sort_by(|a, b| b.1.cmp(&a.1).then_with(|| a.0.cmp(&b.0)));
        self.mode = Mode::Frequency(Frequency {
            header: header.clone(),
            counts,
            rows: grid.rows.len(),
        });
    }

    /// The cursor: display row and column, then the cell's sheet address
    /// (none for a computed column). Sorting moves rows, not addresses, so the
    /// address comes from the data row now at the cursor.
    fn cursor(&self) -> (String, Option<String>) {
        let view = &self.view;
        let rows = view.order.len();
        let columns = view.grid.headers.len();
        let position = format!(
            "R{}/{rows} C{}/{columns}",
            if rows == 0 { 0 } else { view.row + 1 },
            if columns == 0 { 0 } else { view.col + 1 },
        );
        let address = view
            .order
            .get(view.row)
            .and_then(|&source| view.grid.provenance.address(source + 1, view.col));
        (position, address)
    }

    /// The status line, on one screen line in this order: file, sheet,
    /// cursor with the column header, sort, search, note, and the notice
    /// (the last key's message, or the frequency popup's title). Every name
    /// and note is cut to a bound.
    ///
    /// A narrow screen cannot hold every part, so parts claim the width in
    /// priority order: the cursor and its address, whole; the notice, cut
    /// with a trailing `…` only when the cursor leaves too little; the file
    /// name; the search position, whole or not at all; the sheet; the note;
    /// the file's directories, dropped from the front; the header; the sort.
    /// The address gives way only when the cursor alone overflows.
    pub(super) fn status(&self) -> String {
        let view = &self.view;
        let grid = &view.grid;
        let (position, address) = self.cursor();
        let full = match &address {
            Some(address) => format!("{position} {address}"),
            None => position.clone(),
        };
        if let Mode::Search(query) = &self.mode {
            // Typing happens at the end, so a long query shows its end.
            let cursor = if full.width() + 1 + SEPARATOR.len() <= self.width {
                full
            } else {
                position
            };
            let room = self.width.saturating_sub(cursor.width() + SEPARATOR.len());
            let query = one_line(query);
            let query = elide(&query, room.saturating_sub(1), 0);
            return fit(&format!("/{query}{SEPARATOR}{cursor}"), self.width).into_owned();
        }
        let cursor = if full.width() <= self.width {
            full
        } else {
            position
        };
        if cursor.width() >= self.width {
            return fit(&cursor, self.width).into_owned();
        }
        let mut room = Room(self.width - cursor.width());

        let notice = match &self.mode {
            Mode::Frequency(frequency) => Some(format!(
                "{}: {} distinct in {} rows",
                field(&frequency.header, FIELD_WIDTH),
                frequency.counts.len(),
                frequency.rows,
            )),
            _ => self
                .message
                .as_deref()
                .map(|message| field(message, NOTE_WIDTH)),
        };
        let notice = notice.and_then(|notice| room.take(SEPARATOR, |cells| fit(&notice, cells)));

        let file = one_line(&self.file);
        let name = Path::new(&self.file)
            .file_name()
            .map_or(Cow::Borrowed(self.file.as_str()), |name| {
                name.to_string_lossy()
            });
        let name = one_line(&name);
        let shown_name = room.take(SEPARATOR, |cells| {
            let cells = cells.min(FIELD_WIDTH);
            elide(&name, cells, cells / 2)
        });

        let search = self.query.as_ref().and_then(|query| {
            let query = field(query, FIELD_WIDTH);
            let search = match self.matches.binary_search(&(view.row, view.col)) {
                Ok(index) => format!("/{query} {}/{}", index + 1, self.matches.len()),
                Err(_) => format!("/{query} {} hits", self.matches.len()),
            };
            room.take_whole(SEPARATOR, search)
        });

        let sheet = format!(
            "{} [{}/{}]",
            field(&grid.name, FIELD_WIDTH),
            self.current + 1,
            self.names.len().max(1),
        );
        let sheet = room.take(SEPARATOR, |cells| fit(&sheet, cells));
        let note = grid.note.as_deref().and_then(|note| {
            let note = field(note, NOTE_WIDTH);
            room.take(SEPARATOR, |cells| fit(&note, cells))
        });

        // Directories only grow a file name shown whole.
        let path = shown_name.map(|shown| {
            if shown != name {
                return shown;
            }
            let limit = (shown.width() + room.0).min(FIELD_WIDTH.max(shown.width()));
            let path = shrink_path(&file, &name, limit);
            room.0 -= path.width() - shown.width();
            path
        });

        let header = grid.headers.get(view.col).and_then(|header| {
            let header = field(header, FIELD_WIDTH);
            room.take(" ", |cells| fit(&header, cells))
        });
        let sort = view.sort.and_then(|sort| {
            let header = grid.headers.get(sort.column).map_or("", String::as_str);
            let sort = format!(
                "sort {} {} {}",
                field(header, FIELD_WIDTH),
                if sort.descending { "desc" } else { "asc" },
                if sort.numeric { "numeric" } else { "text" }
            );
            room.take(SEPARATOR, |cells| fit(&sort, cells))
        });

        let mut status = String::new();
        for lead in [path, sheet].into_iter().flatten() {
            status.push_str(&lead);
            status.push_str(SEPARATOR);
        }
        status.push_str(&cursor);
        if let Some(header) = header {
            status.push(' ');
            status.push_str(&header);
        }
        for part in [sort, search, note, notice].into_iter().flatten() {
            status.push_str(SEPARATOR);
            status.push_str(&part);
        }
        status
    }
}

/// Between status line parts.
const SEPARATOR: &str = " | ";

/// Widest file, sheet, header, or query on the status line, in cells.
const FIELD_WIDTH: usize = 40;

/// Widest note or message on the status line, in cells.
const NOTE_WIDTH: usize = 80;

/// Status line cells not yet claimed by a part.
struct Room(usize);

impl Room {
    /// Claim `separator` and the text `cut` makes for the cells left after
    /// it; `None`, claiming nothing, when no character of the text fits.
    fn take<'a>(
        &mut self,
        separator: &str,
        cut: impl FnOnce(usize) -> Cow<'a, str>,
    ) -> Option<String> {
        let cells = self.0.checked_sub(separator.width())?;
        let text = cut(cells);
        if text.is_empty() || text == "…" {
            return None;
        }
        self.0 -= separator.width() + text.width();
        Some(text.into_owned())
    }

    /// Claim `separator` and all of `text`, or nothing when it does not fit.
    fn take_whole(&mut self, separator: &str, text: String) -> Option<String> {
        let needed = separator.width() + text.width();
        if needed > self.0 {
            return None;
        }
        self.0 -= needed;
        Some(text)
    }
}

/// `path` in at most `width` cells, which must hold `name`, its last part:
/// whole when it fits, else `…` and the trailing directories that fit,
/// else `name` alone.
fn shrink_path(path: &str, name: &str, width: usize) -> String {
    if path.width() <= width {
        return path.to_owned();
    }
    let mut shown = Cow::Borrowed(name);
    for (index, ch) in path.char_indices().rev() {
        if !std::path::is_separator(ch) {
            continue;
        }
        let tail = &path[index..];
        if tail.width() + 1 > width {
            break;
        }
        shown = Cow::Owned(format!("…{tail}"));
    }
    shown.into_owned()
}

/// `text` with control characters shown as spaces, cut to `width` cells.
fn field(text: &str, width: usize) -> String {
    fit(&one_line(text), width).into_owned()
}

/// Screen columns needed to show columns `from..=to` with one separator
/// between neighbours.
pub(super) fn span(widths: &[usize], from: usize, to: usize) -> usize {
    widths.get(from..=to).map_or(0, |shown| {
        shown.iter().sum::<usize>() + shown.len().saturating_sub(1)
    })
}

enum SortKey {
    Blank,
    Number(f64),
    Text(String),
}

fn compare(a: &SortKey, b: &SortKey, descending: bool) -> Ordering {
    let ordering = match (a, b) {
        (SortKey::Blank, SortKey::Blank) => return Ordering::Equal,
        (SortKey::Blank, _) => return Ordering::Greater,
        (_, SortKey::Blank) => return Ordering::Less,
        (SortKey::Number(a), SortKey::Number(b)) => a.total_cmp(b),
        (SortKey::Text(a), SortKey::Text(b)) => a.cmp(b),
        (SortKey::Number(_), SortKey::Text(_)) => Ordering::Less,
        (SortKey::Text(_), SortKey::Number(_)) => Ordering::Greater,
    };
    if descending {
        ordering.reverse()
    } else {
        ordering
    }
}

#[cfg(test)]
pub(super) mod tests {
    use wolfxl_core::{Cell, CellValue, Sheet};

    use super::*;

    /// A grid from delimited text: first line headers, cells split on `|`,
    /// every field kept as text the way a delimited source arrives.
    pub(in crate::tui) fn grid(name: &str, table: &str) -> Grid {
        let rows = table
            .lines()
            .map(|line| {
                line.split('|')
                    .map(|field| Cell {
                        value: if field.is_empty() {
                            CellValue::Empty
                        } else {
                            CellValue::String(field.to_owned())
                        },
                        number_format: None,
                    })
                    .collect()
            })
            .collect();
        Grid::new(&Sheet::from_rows(name.to_owned(), rows), 12, None)
    }

    pub(in crate::tui) fn app(csv: &str) -> App {
        App::new(
            "data.csv".into(),
            vec!["Sheet1".into()],
            0,
            grid("Sheet1", csv),
        )
    }

    /// The `Table` sheet of the cities fixture as `--header-row 2` shows it:
    /// the header on sheet row 2 below a title row, the first data row on
    /// row 3, opened as `file`.
    pub(in crate::tui) fn cities(file: &str) -> App {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures")
            .join("data-table-cities.xlsx");
        let mut input =
            crate::input::open(Some(path.as_path()), None).expect("open data-table-cities.xlsx");
        let names = input.sheet_names().to_vec();
        let current = names
            .iter()
            .position(|name| name == "Table")
            .expect("a Table sheet");
        let sheet = input.raw_sheet("Table").expect("load the Table sheet");
        let (sheet, provenance) = crate::input::header_sheet(sheet, Some(2)).expect("header row 2");
        let grid = Grid::with_provenance(&sheet, 12, None, &provenance);
        App::new(file.to_owned(), names, current, grid)
    }

    /// Move the cursor to the first `Population` column; returns its index.
    pub(in crate::tui) fn to_population(app: &mut App) -> usize {
        let column = app
            .view
            .grid
            .headers
            .iter()
            .position(|header| header == "Population")
            .expect("a Population column");
        for _ in 0..column {
            press(app, KeyCode::Right);
        }
        column
    }

    pub(in crate::tui) fn press(app: &mut App, code: KeyCode) -> Action {
        app.handle_key(KeyEvent::new(code, KeyModifiers::NONE))
    }

    fn keys(app: &mut App, typed: &str) {
        for typed in typed.chars() {
            press(app, KeyCode::Char(typed));
        }
    }

    /// City at the cursor row of the view.
    fn city(app: &App) -> &str {
        &app.view.grid.rows[app.view.order[app.view.row]][0].text
    }

    fn address(app: &App) -> Option<String> {
        app.cursor().1
    }

    #[test]
    fn cursor_address_follows_sorted_rows_below_a_title_row() {
        let mut app = cities("data-table-cities.xlsx");
        assert_eq!(app.view.grid.headers[0], "City");
        assert_eq!(app.view.grid.rows.len(), 13, "twelve cities and a total");
        assert_eq!(city(&app), "Shanghai, China");
        assert_eq!(
            address(&app).as_deref(),
            Some("C3"),
            "below the header row 2"
        );
        let population = to_population(&mut app);
        assert_eq!(address(&app).as_deref(), Some("F3"));

        press(&mut app, KeyCode::Char('s'));
        assert_eq!(
            app.view.sort,
            Some(Sort {
                column: population,
                descending: false,
                numeric: true
            })
        );
        assert_eq!(
            app.message.as_deref(),
            Some("sorted by Population, ascending")
        );
        assert_eq!(city(&app), "Shenzhen, China", "the smallest city first");
        assert_eq!(address(&app).as_deref(), Some("F14"), "its own sheet row");
        press(&mut app, KeyCode::Char('G'));
        assert_eq!(
            (city(&app), address(&app).as_deref()),
            ("Total", Some("F15"))
        );

        press(&mut app, KeyCode::Char('s'));
        assert!(app.view.sort.is_some_and(|sort| sort.descending));
        assert_eq!(city(&app), "Shenzhen, China", "now the last row");
        assert_eq!(address(&app).as_deref(), Some("F14"));
        press(&mut app, KeyCode::Up);
        assert_eq!(city(&app), "Guangzhou, China");
        assert_eq!(address(&app).as_deref(), Some("F13"));
        let position = format!("R12/13 C{}/", population + 1);
        assert!(app.status().contains(&position), "{}", app.status());

        press(&mut app, KeyCode::Char('S'));
        assert_eq!(city(&app), "Shenzhen, China", "source order again");
        assert_eq!(address(&app).as_deref(), Some("F14"));
        press(&mut app, KeyCode::Char('g'));
        assert_eq!(
            (city(&app), address(&app).as_deref()),
            ("Shanghai, China", Some("F3"))
        );
    }

    /// Column `column` of the view, top to bottom.
    fn column(app: &App, column: usize) -> Vec<String> {
        app.view
            .order
            .iter()
            .map(|&source| app.view.grid.rows[source][column].text.clone())
            .collect()
    }

    fn cursor(app: &App) -> (usize, usize) {
        (app.view.row, app.view.col)
    }

    const PEOPLE: &str = "\
Name|Amount|City
ann|30|Oslo
Bob|$1,200.00|Lima
cy||Oslo
Dee|(5)|Rome
eve|7|oslo";

    fn tall(rows: usize) -> String {
        let mut table = String::from("A|B|C");
        for row in 0..rows {
            table.push_str(&format!("\n{row}|x{row}|y{row}"));
        }
        table
    }

    #[test]
    fn cursor_stays_inside_the_sheet() {
        let mut app = app(&tall(100));
        app.set_viewport(40, 12);
        press(&mut app, KeyCode::Up);
        press(&mut app, KeyCode::Left);
        assert_eq!(cursor(&app), (0, 0));
        press(&mut app, KeyCode::End);
        press(&mut app, KeyCode::Right);
        assert_eq!(cursor(&app), (0, 2));
        press(&mut app, KeyCode::Home);
        assert_eq!(cursor(&app), (0, 0));
        press(&mut app, KeyCode::PageDown);
        assert_eq!(
            cursor(&app),
            (10, 0),
            "a page is the screen minus header and status"
        );
        assert_eq!(app.view.top, 1);
        press(&mut app, KeyCode::Char('G'));
        press(&mut app, KeyCode::Down);
        press(&mut app, KeyCode::PageDown);
        assert_eq!(cursor(&app), (99, 0));
        assert_eq!(app.view.top, 90, "the last row sits on the last line");
        press(&mut app, KeyCode::PageUp);
        assert_eq!(cursor(&app), (89, 0));
        assert_eq!(app.view.top, 89);
        press(&mut app, KeyCode::Char('g'));
        press(&mut app, KeyCode::PageUp);
        assert_eq!((cursor(&app), app.view.top), ((0, 0), 0));
    }

    #[test]
    fn empty_sheet_navigation_is_a_no_op() {
        let mut app = app("");
        app.set_viewport(20, 5);
        for code in [
            KeyCode::Down,
            KeyCode::Right,
            KeyCode::End,
            KeyCode::Char('G'),
            KeyCode::PageDown,
            KeyCode::Char('s'),
            KeyCode::Char('f'),
        ] {
            assert_eq!(press(&mut app, code), Action::Continue);
        }
        assert_eq!(cursor(&app), (0, 0));
        assert!(matches!(app.mode, Mode::Browse));
    }

    #[test]
    fn wide_columns_scroll_horizontally() {
        let mut app = app("Alpha|Beta|Gamma|Delta\naaaaaaaaaa|bbbbbbbbbb|cccccccccc|dddddddddd");
        app.set_viewport(21, 5);
        press(&mut app, KeyCode::Right);
        assert_eq!(
            app.view.left, 0,
            "two 10-wide columns and a separator fit in 21"
        );
        press(&mut app, KeyCode::Right);
        assert_eq!(app.view.left, 1);
        press(&mut app, KeyCode::Home);
        assert_eq!(app.view.left, 0);
    }

    #[test]
    fn search_is_case_insensitive_and_wraps() {
        let mut app = app(PEOPLE);
        press(&mut app, KeyCode::Char('/'));
        keys(&mut app, "OSLx");
        press(&mut app, KeyCode::Backspace);
        assert_eq!(app.status(), "/OSL | R1/5 C1/3 A2");
        press(&mut app, KeyCode::Enter);
        assert_eq!(app.matches, [(0, 2), (2, 2), (4, 2)]);
        assert_eq!(cursor(&app), (0, 2));
        assert!(app.status().ends_with("| /OSL 1/3"), "{}", app.status());
        press(&mut app, KeyCode::Char('n'));
        assert_eq!(cursor(&app), (2, 2));
        press(&mut app, KeyCode::Char('n'));
        assert_eq!(cursor(&app), (4, 2));
        press(&mut app, KeyCode::Char('n'));
        assert_eq!(cursor(&app), (0, 2));
        assert_eq!(app.message.as_deref(), Some("search wrapped to the top"));
        press(&mut app, KeyCode::Char('N'));
        assert_eq!(cursor(&app), (4, 2));
        assert_eq!(app.message.as_deref(), Some("search wrapped to the bottom"));
        press(&mut app, KeyCode::Char('N'));
        assert_eq!(cursor(&app), (2, 2));
        assert_eq!(app.message, None);
    }

    #[test]
    fn search_reports_misses_and_escape_cancels() {
        let mut app = app(PEOPLE);
        press(&mut app, KeyCode::Char('n'));
        assert_eq!(
            app.message.as_deref(),
            Some("no search yet; press / to search")
        );
        press(&mut app, KeyCode::Char('/'));
        keys(&mut app, "zzz");
        press(&mut app, KeyCode::Enter);
        assert_eq!(app.message.as_deref(), Some("no match for /zzz"));
        assert_eq!(cursor(&app), (0, 0));
        press(&mut app, KeyCode::Char('/'));
        keys(&mut app, "q");
        press(&mut app, KeyCode::Esc);
        assert!(matches!(app.mode, Mode::Browse));
        assert_eq!(
            app.query.as_deref(),
            Some("zzz"),
            "escape keeps the old query"
        );
    }

    #[test]
    fn numeric_sort_reverses_and_restores() {
        let mut app = app(PEOPLE);
        press(&mut app, KeyCode::Right);
        press(&mut app, KeyCode::Char('s'));
        assert_eq!(column(&app, 1), ["(5)", "7", "30", "$1,200.00", ""]);
        assert_eq!(
            app.view.sort,
            Some(Sort {
                column: 1,
                descending: false,
                numeric: true
            })
        );
        press(&mut app, KeyCode::Char('s'));
        assert_eq!(
            column(&app, 1),
            ["$1,200.00", "30", "7", "(5)", ""],
            "blanks stay last"
        );
        assert!(app.status().contains("| sort Amount desc numeric"));
        press(&mut app, KeyCode::Char('s'));
        assert_eq!(column(&app, 1), ["(5)", "7", "30", "$1,200.00", ""]);
        press(&mut app, KeyCode::Char('S'));
        assert_eq!(column(&app, 0), ["ann", "Bob", "cy", "Dee", "eve"]);
        assert_eq!(app.view.sort, None);
    }

    #[test]
    fn text_sort_ignores_case_and_keeps_ties_in_source_order() {
        let mut app = app(PEOPLE);
        press(&mut app, KeyCode::Char('s'));
        assert_eq!(column(&app, 0), ["ann", "Bob", "cy", "Dee", "eve"]);
        press(&mut app, KeyCode::Char('s'));
        assert_eq!(column(&app, 0), ["eve", "Dee", "cy", "Bob", "ann"]);
        press(&mut app, KeyCode::End);
        press(&mut app, KeyCode::Char('s'));
        assert_eq!(column(&app, 2), ["Lima", "Oslo", "Oslo", "oslo", "Rome"]);
        assert_eq!(column(&app, 0), ["Bob", "ann", "cy", "eve", "Dee"]);
        assert!(app.view.sort.is_some_and(|sort| !sort.numeric));
    }

    #[test]
    fn one_text_value_makes_the_sort_textual() {
        let mut app = app("N\n10\n9\nn/a\n100");
        press(&mut app, KeyCode::Char('s'));
        assert_eq!(column(&app, 0), ["10", "100", "9", "n/a"]);
        assert!(app.view.sort.is_some_and(|sort| !sort.numeric));
    }

    #[test]
    fn search_matches_follow_the_sort() {
        let mut app = app(PEOPLE);
        press(&mut app, KeyCode::Char('/'));
        keys(&mut app, "dee");
        press(&mut app, KeyCode::Enter);
        assert_eq!(app.matches, [(3, 0)]);
        press(&mut app, KeyCode::Char('s'));
        press(&mut app, KeyCode::Char('s'));
        assert_eq!(app.matches, [(1, 0)]);
    }

    #[test]
    fn frequency_counts_each_displayed_value() {
        let mut app = app(PEOPLE);
        press(&mut app, KeyCode::End);
        press(&mut app, KeyCode::Char('f'));
        let Mode::Frequency(frequency) = &app.mode else {
            panic!("f opens the frequency popup");
        };
        assert_eq!(
            *frequency,
            Frequency {
                header: "City".into(),
                counts: vec![
                    ("Oslo".into(), 2),
                    ("Lima".into(), 1),
                    ("Rome".into(), 1),
                    ("oslo".into(), 1),
                ],
                rows: 5,
            }
        );
        assert_eq!(press(&mut app, KeyCode::Char('q')), Action::Continue);
        assert!(matches!(app.mode, Mode::Browse), "any key closes the popup");
        press(&mut app, KeyCode::Home);
        press(&mut app, KeyCode::Right);
        press(&mut app, KeyCode::Char('f'));
        let Mode::Frequency(frequency) = &app.mode else {
            panic!("f opens the frequency popup");
        };
        assert!(
            frequency.counts.contains(&(String::new(), 1)),
            "blanks count"
        );
    }

    #[test]
    fn quit_keys() {
        for key in [
            KeyEvent::new(KeyCode::Char('q'), KeyModifiers::NONE),
            KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE),
            KeyEvent::new(KeyCode::Char('c'), KeyModifiers::CONTROL),
        ] {
            assert_eq!(app(PEOPLE).handle_key(key), Action::Quit);
        }
    }

    #[test]
    fn tab_loads_each_sheet_once_and_keeps_its_state() {
        let mut app = App::new(
            "book.xlsx".into(),
            vec!["One".into(), "Two".into(), "Three".into()],
            1,
            grid("Two", "B\n1\n2"),
        );
        press(&mut app, KeyCode::Down);
        assert_eq!(press(&mut app, KeyCode::Tab), Action::Load(2));
        app.insert(2, grid("Three", "C\nz"));
        assert_eq!((app.current, app.view.grid.name.as_str()), (2, "Three"));
        assert_eq!(press(&mut app, KeyCode::Tab), Action::Load(0));
        app.insert(0, grid("One", "A\nx"));
        assert_eq!(press(&mut app, KeyCode::Tab), Action::Continue);
        assert_eq!(app.view.grid.name, "Two");
        assert_eq!(cursor(&app), (1, 0), "the cursor is where it was left");
        assert_eq!(press(&mut app, KeyCode::BackTab), Action::Continue);
        assert_eq!(app.view.grid.name, "One");
        assert!(app
            .status()
            .starts_with("book.xlsx | One [1/3] | R1/1 C1/1 A2 A"));
    }

    #[test]
    fn tab_with_one_sheet_says_so() {
        let mut app = app(PEOPLE);
        assert_eq!(press(&mut app, KeyCode::Tab), Action::Continue);
        assert_eq!(app.message.as_deref(), Some("only one sheet"));
    }

    #[test]
    fn status_shows_names_and_load_errors_on_one_bounded_line() {
        let long = "日本".repeat(40);
        let file = format!("/tmp/private\ndir/{long}\tbook\u{1b}[31m.xlsx");
        let sheet = format!("Q1\r\n{long}");
        let note = format!("not loaded: bad\nzip\u{1b}]0;title\u{7} {long}");
        let mut app = App::new(
            file.clone(),
            vec!["One".into(), sheet.clone()],
            0,
            grid("One", "A\nx"),
        );
        assert_eq!(press(&mut app, KeyCode::Tab), Action::Load(1));
        app.insert(1, Grid::unavailable(&sheet, note.clone()));
        let status = app.status();
        assert!(!status.contains(char::is_control), "{status:?}");
        assert!(!status.contains("private"), "only the file name: {status}");
        assert!(status.contains("book [31m.xlsx | Q1  日本"), "{status}");
        assert!(
            status.contains("[2/2] | R0/0 C0/0 | not loaded: bad zip ]0;title"),
            "{status}"
        );
        for part in status.split(SEPARATOR) {
            assert!(part.width() <= NOTE_WIDTH, "{part}");
        }
        assert!(
            status.width() <= 2 * FIELD_WIDTH + NOTE_WIDTH + 30,
            "{status}"
        );
        assert_eq!(app.file, file, "the source path is kept");
        assert_eq!(app.view.grid.name, sheet);
        assert_eq!(app.sheet_name(1), sheet);
        assert_eq!(app.view.grid.note.as_deref(), Some(note.as_str()));
    }

    #[test]
    fn narrow_status_keeps_the_cursor_whole() {
        let mut app = App::new(
            format!("/data/{}.csv", "quarterly-revenue-".repeat(5)),
            vec!["Sheet1".into()],
            0,
            grid(&"Very long sheet name ".repeat(4), PEOPLE),
        );
        press(&mut app, KeyCode::Down);
        press(&mut app, KeyCode::End);
        for width in [9, 12, 16, 20, 24, 30, 48, 200] {
            app.set_viewport(width, 6);
            let status = app.status();
            assert!(status.contains("R2/5 C3/3"), "{width}: {status}");
            assert!(status.width() <= usize::from(width), "{width}: {status}");
        }
        assert!(
            app.status()
                .starts_with("quarterly-revenue-qu…rterly-revenue-.csv | Very long"),
            "{}",
            app.status()
        );
        press(&mut app, KeyCode::Char('/'));
        keys(&mut app, "needle-in-a-very-long-haystack-query");
        for width in [14, 16, 20, 30, 200] {
            app.set_viewport(width, 6);
            let status = app.status();
            assert!(status.starts_with('/'), "{width}: {status}");
            assert!(status.contains(" | R2/5 C3/3"), "{width}: {status}");
            assert!(status.width() <= usize::from(width), "{width}: {status}");
        }
        app.set_viewport(20, 6);
        assert_eq!(app.status(), "/…ery | R2/5 C3/3 C3", "typing shows the end");
        app.set_viewport(14, 6);
        assert_eq!(
            app.status(),
            "/… | R2/5 C3/3",
            "the address gives way before the position"
        );
        press(&mut app, KeyCode::Enter);
        assert_eq!(
            app.query.as_deref(),
            Some("needle-in-a-very-long-haystack-query")
        );
    }

    #[test]
    fn typed_control_characters_are_shown_as_spaces_and_searched_as_typed() {
        let mut app = app("Note\na\tb\nab");
        press(&mut app, KeyCode::Char('/'));
        keys(&mut app, "a\tb");
        assert_eq!(app.status(), "/a b | R1/2 C1/1 A2");
        press(&mut app, KeyCode::Enter);
        assert_eq!(app.query.as_deref(), Some("a\tb"));
        assert_eq!(app.message.as_deref(), Some("no match for /a\tb"));
        let status = app.status();
        assert!(!status.contains(char::is_control), "{status:?}");
        assert!(
            status.ends_with("| R1/2 C1/1 A2 Note | /a b 0 hits | no match for /a b"),
            "{status}"
        );
    }
}
