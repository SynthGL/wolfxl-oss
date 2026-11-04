//! `wolfxl peek --interactive`: a full-screen grid over one sheet.
//!
//! [`App`] holds every piece of state and turns keys into [`app::Action`]s;
//! `draw` renders it; [`run`] owns the terminal and the event loop.
//! `ratatui::try_init` installs a panic hook that restores the terminal, and
//! [`run`] restores it on every return path.

mod app;
mod draw;
mod grid;

use std::io::IsTerminal;
use std::path::Path;

use anyhow::{bail, Context, Result};
use ratatui::crossterm::event::{self, Event};
use ratatui::DefaultTerminal;

pub(crate) use app::App;
pub(crate) use grid::Grid;

use app::Action;

/// Refuse `--interactive` before any input is read: the grid needs a
/// terminal on standard output, and standard input for keys.
pub(crate) fn preflight(file: &Path) -> Result<()> {
    if file == Path::new("-") {
        bail!(
            "--interactive reads keys from standard input, so it cannot also read the data there; pass a file path"
        );
    }
    if !std::io::stdout().is_terminal() {
        bail!("--interactive needs a terminal on standard output; omit it to print a preview");
    }
    Ok(())
}

/// Browse `app` until the user quits. `load` returns sheet `name` the first
/// time Tab reaches it.
pub(crate) fn run(mut app: App, mut load: impl FnMut(&str) -> Grid) -> Result<()> {
    let mut terminal = match ratatui::try_init() {
        Ok(terminal) => terminal,
        Err(error) => {
            ratatui::restore();
            return Err(error).context("starting the terminal grid");
        }
    };
    let outcome = event_loop(&mut terminal, &mut app, &mut load);
    ratatui::restore();
    outcome
}

fn event_loop(
    terminal: &mut DefaultTerminal,
    app: &mut App,
    load: &mut impl FnMut(&str) -> Grid,
) -> Result<()> {
    loop {
        terminal
            .draw(|frame| draw::draw(frame, app))
            .context("drawing the terminal grid")?;
        let Event::Key(key) = event::read().context("reading a key")? else {
            continue;
        };
        match app.handle_key(key) {
            Action::Continue => {}
            Action::Quit => return Ok(()),
            Action::Load(index) => {
                let grid = load(app.sheet_name(index));
                app.insert(index, grid);
            }
        }
    }
}
