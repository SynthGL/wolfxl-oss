//! Drawing the interactive grid: a frozen header line, the data rows on
//! screen, a status line, and the frequency popup. Only the visible window
//! of rows and columns is formatted.

use std::borrow::Cow;

use ratatui::buffer::Buffer;
use ratatui::layout::Rect;
use ratatui::style::{Color, Modifier, Style};
use ratatui::text::Line;
use ratatui::widgets::{Block, Clear, Paragraph};
use ratatui::Frame;
use unicode_width::{UnicodeWidthChar, UnicodeWidthStr};

use super::app::{App, Frequency, Mode};

/// Widest frequency popup, in terminal cells.
const POPUP_WIDTH: u16 = 60;

pub(crate) fn draw(frame: &mut Frame, app: &mut App) {
    let area = frame.area();
    app.set_viewport(area.width, area.height);
    if area.width == 0 || area.height == 0 {
        return;
    }
    let buffer = frame.buffer_mut();
    draw_grid(buffer, area, app);
    draw_status(buffer, area, app);
    if let Mode::Frequency(frequency) = &app.mode {
        draw_frequency(frame, area, frequency);
    }
}

fn draw_grid(buffer: &mut Buffer, area: Rect, app: &App) {
    let view = &app.view;
    let grid = &view.grid;
    if grid.headers.is_empty() {
        let style = Style::default().add_modifier(Modifier::DIM);
        buffer.set_stringn(
            area.x,
            area.y,
            "(empty sheet)",
            usize::from(area.width),
            style,
        );
        return;
    }
    let header_style = Style::default()
        .fg(Color::Cyan)
        .add_modifier(Modifier::BOLD);
    let separator_style = Style::default().fg(Color::DarkGray);
    let body_lines = area.height.saturating_sub(2);
    let right = area.right();
    let mut x = area.x;
    let mut column = view.left;
    while column < grid.headers.len() && x < right {
        let width = grid
            .widths
            .get(column)
            .map_or(1, |&width| u16::try_from(width).unwrap_or(u16::MAX))
            .min(right - x);
        let selected = column == view.col;
        let style = if selected {
            header_style.add_modifier(Modifier::REVERSED)
        } else {
            header_style
        };
        put(
            buffer,
            x,
            area.y,
            &grid.headers[column],
            width,
            false,
            style,
        );
        for line in 0..body_lines {
            let position = view.top + usize::from(line);
            let Some(cell) = view
                .order
                .get(position)
                .and_then(|&source| grid.rows.get(source))
                .and_then(|row| row.get(column))
            else {
                break;
            };
            let style = if selected && position == view.row {
                Style::default().add_modifier(Modifier::REVERSED)
            } else {
                Style::default()
            };
            let numeric = cell.number.is_some();
            put(
                buffer,
                x,
                area.y + 1 + line,
                &cell.text,
                width,
                numeric,
                style,
            );
        }
        x += width;
        if x < right {
            for line in 0..=body_lines {
                buffer.set_string(x, area.y + line, "│", separator_style);
            }
            x += 1;
        }
        column += 1;
    }
}

fn draw_status(buffer: &mut Buffer, area: Rect, app: &App) {
    let y = area.bottom() - 1;
    let style = Style::default().add_modifier(Modifier::REVERSED);
    buffer.set_style(Rect::new(area.x, y, area.width, 1), style);
    buffer.set_stringn(area.x, y, app.status(), usize::from(area.width), style);
}

fn draw_frequency(frame: &mut Frame, area: Rect, frequency: &Frequency) {
    let width = area.width.saturating_sub(4).min(POPUP_WIDTH);
    let height = area.height.saturating_sub(2);
    if width < 8 || height < 3 {
        return;
    }
    let lines_room = usize::from(height - 2);
    let shown = if frequency.counts.len() > lines_room {
        lines_room.saturating_sub(1)
    } else {
        frequency.counts.len()
    };
    let count_width = frequency
        .counts
        .first()
        .map_or(1, |(_, count)| count.to_string().len());
    let value_width = usize::from(width - 2).saturating_sub(count_width + 2);
    let mut lines: Vec<Line> = frequency.counts[..shown]
        .iter()
        .map(|(value, count)| {
            let value = if value.is_empty() {
                "(blank)"
            } else {
                value.as_str()
            };
            Line::from(format!(
                "{count:>count_width$}  {}",
                fit(value, value_width)
            ))
        })
        .collect();
    if shown < frequency.counts.len() {
        lines.push(Line::from(format!(
            "… {} more values",
            frequency.counts.len() - shown
        )));
    }
    let used = u16::try_from(lines.len()).unwrap_or(height - 2);
    let popup = Rect::new(
        area.x + (area.width - width) / 2,
        area.y + 1,
        width,
        used + 2,
    );
    let block = Block::bordered()
        .title(format!(
            " {}: {} distinct in {} rows ",
            frequency.header,
            frequency.counts.len(),
            frequency.rows
        ))
        .title_bottom(" any key closes ");
    frame.render_widget(Clear, popup);
    frame.render_widget(Paragraph::new(lines).block(block), popup);
}

/// Write `text` into `width` cells at `(x, y)`, cut with an ellipsis when
/// too wide and padded to fill the cells.
fn put(buffer: &mut Buffer, x: u16, y: u16, text: &str, width: u16, right: bool, style: Style) {
    let width = usize::from(width);
    let text = fit(text, width);
    let pad = width.saturating_sub(text.width());
    let padded = if right {
        format!("{:pad$}{text}", "")
    } else {
        format!("{text}{:pad$}", "")
    };
    buffer.set_stringn(x, y, padded, width, style);
}

/// `text` when it fits in `width` cells, else its longest prefix that fits
/// with a trailing `…`.
pub(super) fn fit(text: &str, width: usize) -> Cow<'_, str> {
    elide(text, width, usize::MAX)
}

/// `text` when it fits in `width` cells, else at most `head` cells of its
/// start and as much of its end as still fits around one `…`. A `head` of
/// zero keeps only the end; `usize::MAX` keeps only the start.
pub(super) fn elide(text: &str, width: usize, head: usize) -> Cow<'_, str> {
    if text.width() <= width {
        return Cow::Borrowed(text);
    }
    if width == 0 {
        return Cow::Borrowed("");
    }
    let room = width - 1;
    let head = head.min(room);
    let mut used = 0;
    let mut kept = 0;
    for (index, ch) in text.char_indices() {
        let ch_width = ch.width().unwrap_or(0);
        if used + ch_width > head {
            break;
        }
        used += ch_width;
        kept = index + ch.len_utf8();
    }
    let tail_room = room - head;
    let mut start = text.len();
    let mut tail = 0;
    if tail_room > 0 {
        for (index, ch) in text[kept..].char_indices().rev() {
            let ch_width = ch.width().unwrap_or(0);
            if tail + ch_width > tail_room {
                break;
            }
            start = kept + index;
            tail += ch_width;
        }
    }
    Cow::Owned(format!("{}…{}", &text[..kept], &text[start..]))
}

#[cfg(test)]
mod tests {
    use ratatui::backend::TestBackend;
    use ratatui::crossterm::event::KeyCode;
    use ratatui::Terminal;

    use super::super::app::tests::{app, cities, grid, press, to_population};
    use super::*;

    fn screen(app: &mut App, width: u16, height: u16) -> Vec<String> {
        let mut terminal = Terminal::new(TestBackend::new(width, height)).expect("test terminal");
        terminal.draw(|frame| draw(frame, app)).expect("draw");
        let buffer = terminal.backend().buffer();
        (0..buffer.area.height)
            .map(|y| {
                (0..buffer.area.width)
                    .map(|x| buffer[(x, y)].symbol())
                    .collect::<String>()
            })
            .collect()
    }

    const LEDGER: &str = "\
Region|Rep|Amount|Note
West|Ann|$1,200.00|first quarter carry-over
East|Bob|950|
North|Dee|1,200|ok";

    #[test]
    fn first_screen_freezes_the_header_and_shows_status() {
        let mut app = app(LEDGER);
        assert_eq!(
            screen(&mut app, 48, 6),
            [
                "Region│Rep│Amount   │Note        │              ",
                "West  │Ann│$1,200.00│first quart…│              ",
                "East  │Bob│      950│            │              ",
                "North │Dee│    1,200│ok          │              ",
                "      │   │         │            │              ",
                "data.csv | Sheet1 [1/1] | R1/3 C1/4 A2 Region   ",
            ]
        );
    }

    #[test]
    fn long_values_are_cut_and_scrolled_rows_keep_the_header() {
        let mut table = String::from("Id|Text");
        for row in 1..=20 {
            table.push_str(&format!("\n{row}|abcdefghijklmnopqrstuvwxyz"));
        }
        let mut app = app(&table);
        screen(&mut app, 20, 5);
        press(&mut app, KeyCode::Char('G'));
        let lines = screen(&mut app, 20, 5);
        assert_eq!(lines[0], "Id│Text        │    ");
        assert_eq!(lines[1], "18│abcdefghijk…│    ");
        assert_eq!(lines[3], "20│abcdefghijk…│    ");
        assert_eq!(
            lines[4], "d… | R20/20 C1/2 A21",
            "the file gives way so the cursor and its address stay whole"
        );
    }

    #[test]
    fn narrow_status_line_keeps_the_cursor() {
        let mut app = App::new(
            "/reports/2026/a-very-long-workbook-name.csv".into(),
            vec!["Sheet1".into()],
            0,
            grid("Sheet1", LEDGER),
        );
        let lines = screen(&mut app, 24, 4);
        assert_eq!(lines[3], "a-ve….csv | R1/3 C1/4 A2");
        press(&mut app, KeyCode::Char('/'));
        for typed in "carry-over".chars() {
            press(&mut app, KeyCode::Char(typed));
        }
        let lines = screen(&mut app, 24, 4);
        assert_eq!(lines[3], "/…ry-over | R1/3 C1/4 A2");
    }

    const CITIES_PATH: &str =
        "engines/wolfxl/bindings/dotnet-facade/Checkers/Tables/TestData/DataTableCities.xlsx";

    /// The cities fixture under a long path, cursor on the total population
    /// after searching for it, then sorted by population: the total stays
    /// last, so the cursor stays on the search hit. Returns the screens
    /// before and after the sort, the searched total, and the sheet label.
    fn sorted_on_the_total(width: u16) -> (Vec<String>, Vec<String>, String, String) {
        let mut app = cities(CITIES_PATH);
        let population = to_population(&mut app);
        let total = app.view.grid.rows[12][population].text.clone();
        let sheet = format!(
            "{} [{}/{}]",
            app.view.grid.name,
            app.current + 1,
            app.names.len()
        );
        press(&mut app, KeyCode::Char('/'));
        for typed in total.chars() {
            press(&mut app, KeyCode::Char(typed));
        }
        press(&mut app, KeyCode::Enter);
        let searched = screen(&mut app, width, 8);
        press(&mut app, KeyCode::Char('s'));
        (searched, screen(&mut app, width, 8), total, sheet)
    }

    #[test]
    fn status_keeps_the_notice_cursor_and_file_name_at_common_widths() {
        let (searched, sorted, total, _) = sorted_on_the_total(60);
        assert_eq!(
            searched[7],
            format!("…/DataTableCities.xlsx | R13/13 C4/10 F15 | /{total} 1/1"),
            "without a notice the file name, cursor, and search hit all fit"
        );
        assert_eq!(
            sorted[7], "Dat…lsx | R13/13 C4/10 F15 | sorted by Population, ascending",
            "the whole notice and cursor come first"
        );

        let (_, sorted, _, _) = sorted_on_the_total(80);
        assert_eq!(
            sorted[7],
            "DataTableCities.xlsx | Tab… | R13/13 C4/10 F15 | sorted by Population, ascending"
        );

        let (_, sorted, total, sheet) = sorted_on_the_total(120);
        assert_eq!(
            sorted[7],
            format!(
                "…/TestData/DataTableCities.xlsx | {sheet} | R13/13 C4/10 F15 P… \
                 | /{total} 1/1 | sorted by Population, ascending"
            )
        );
        for (width, line) in [(60, &searched[7]), (120, &sorted[7])] {
            assert_eq!(line.width(), width, "{line}");
        }
    }

    #[test]
    fn long_notices_end_with_an_ellipsis() {
        let mut app = cities(CITIES_PATH);
        to_population(&mut app);
        press(&mut app, KeyCode::Char('s'));
        let lines = screen(&mut app, 40, 6);
        assert_eq!(lines[5], "R1/13 C4/10 F14 | sorted by Population,…");
    }

    #[test]
    fn frequency_popup_lists_counts() {
        let mut app = app("City\nOslo\nLima\nOslo\n");
        press(&mut app, KeyCode::Char('f'));
        let lines = screen(&mut app, 40, 8);
        assert!(
            lines[1].contains(" City: 2 distinct in 3 rows "),
            "{lines:?}"
        );
        assert!(lines[2].contains("2  Oslo"), "{lines:?}");
        assert!(lines[3].contains("1  Lima"), "{lines:?}");
        assert!(lines[4].contains(" any key closes "), "{lines:?}");
    }

    #[test]
    fn empty_sheet_draws_a_placeholder() {
        let mut app = app("");
        let lines = screen(&mut app, 40, 3);
        assert_eq!(lines[0].trim_end(), "(empty sheet)");
        assert!(
            lines[2].starts_with("data.csv | Sheet1 [1/1] | R0/0 C0/0"),
            "{lines:?}"
        );
    }

    #[test]
    fn fit_cuts_with_an_ellipsis() {
        assert_eq!(fit("abc", 3), "abc");
        assert_eq!(fit("abcd", 3), "ab…");
        assert_eq!(fit("abcd", 0), "");
        assert_eq!(fit("日本語", 4), "日…");
    }

    #[test]
    fn elide_keeps_the_start_the_end_or_both() {
        assert_eq!(elide("abc", 3, 0), "abc");
        assert_eq!(elide("abcdefgh", 5, 2), "ab…gh");
        assert_eq!(elide("abcdef", 4, 0), "…def");
        assert_eq!(elide("abcdef", 4, usize::MAX), "abc…");
        assert_eq!(elide("abcd", 0, 1), "");
        assert_eq!(elide("日本語テキスト", 6, 2), "日…ト");
    }
}
