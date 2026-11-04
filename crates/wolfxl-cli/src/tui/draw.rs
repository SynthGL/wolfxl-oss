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
fn fit(text: &str, width: usize) -> Cow<'_, str> {
    if text.width() <= width {
        return Cow::Borrowed(text);
    }
    if width == 0 {
        return Cow::Borrowed("");
    }
    let mut out = String::new();
    let mut used = 0;
    for ch in text.chars() {
        let ch_width = ch.width().unwrap_or(0);
        if used + ch_width > width - 1 {
            break;
        }
        out.push(ch);
        used += ch_width;
    }
    out.push('…');
    Cow::Owned(out)
}

#[cfg(test)]
mod tests {
    use ratatui::backend::TestBackend;
    use ratatui::crossterm::event::KeyCode;
    use ratatui::Terminal;

    use super::super::app::tests::{app, press};
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
                "data.csv | Sheet1 [1/1] | R1/3 C1/4 Region      ",
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
        assert!(lines[4].starts_with("data.csv | Sheet1 ["));
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
}
