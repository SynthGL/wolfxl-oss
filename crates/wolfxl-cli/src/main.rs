use std::path::PathBuf;
use std::process::ExitCode;

use clap::{Parser, ValueEnum};

mod commands;
#[allow(dead_code)]
mod eval;
mod render;

/// Fast, agent-friendly spreadsheet previews.
///
/// `wolfxl peek <file>` prints a styled, token-efficient view of a workbook or
/// delimited file: box / text / csv / json / markdown output, sheet
/// selection, row and column caps, and width caps.
/// `wolfxl map <file>` prints a one-page summary of every sheet.
/// `wolfxl agent <file> --max-tokens N` composes a token-budgeted briefing.
/// `wolfxl schema <file>` emits per-column type, cardinality, and format.
#[derive(Parser, Debug)]
#[command(name = "wolfxl", version, about, long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(clap::Subcommand, Debug)]
enum Command {
    /// Print a preview of a spreadsheet.
    Peek(PeekArgs),
    /// Print a one-page workbook overview (sheets, dims, headers, named ranges).
    Map(MapArgs),
    /// Compose a token-budgeted workbook briefing for an LLM context window.
    Agent(AgentArgs),
    /// Per-column schema: type, cardinality, null count, format category.
    Schema(SchemaArgs),
}

#[derive(clap::Args, Debug)]
struct MapArgs {
    /// Path to a spreadsheet (.xlsx/.xlsm/.xls/.xlsb/.ods/.csv/.tsv/.txt).
    file: PathBuf,

    /// Output format.
    #[arg(short = 'f', long = "format", default_value = "json")]
    format: MapFormat,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
pub enum MapFormat {
    Json,
    Text,
}

#[derive(clap::Args, Debug)]
struct AgentArgs {
    /// Path to a spreadsheet (.xlsx/.xlsm/.xls/.xlsb/.ods/.csv/.tsv/.txt).
    file: PathBuf,

    /// Token budget (cl100k_base). Output is composed greedily to fit.
    #[arg(short = 't', long = "max-tokens", default_value_t = 800)]
    max_tokens: usize,

    /// Sheet to focus on (default: largest data-class sheet, else first).
    #[arg(short = 's', long)]
    sheet: Option<String>,
}

#[derive(clap::Args, Debug)]
struct SchemaArgs {
    /// Path to a spreadsheet (.xlsx/.xlsm/.xls/.xlsb/.ods/.csv/.tsv/.txt).
    file: PathBuf,

    /// Sheet name. Omit to schema every sheet in the workbook.
    #[arg(short = 's', long)]
    sheet: Option<String>,

    /// Output format.
    #[arg(short = 'f', long = "format", default_value = "json")]
    format: SchemaFormat,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
pub enum SchemaFormat {
    Json,
    Text,
}

#[derive(clap::Args, Debug)]
struct PeekArgs {
    /// Path to a spreadsheet (.xlsx/.xlsm/.xls/.xlsb/.ods/.csv/.tsv/.txt).
    file: PathBuf,

    /// Sheet name (default: first sheet).
    #[arg(short = 's', long)]
    sheet: Option<String>,

    /// Maximum number of rows to display (0 = all).
    #[arg(short = 'n', long = "max-rows", default_value_t = 50)]
    max_rows: usize,

    /// Maximum number of columns in JSON output (0 = all).
    #[arg(long = "max-columns", default_value_t = 0)]
    max_columns: usize,

    /// Maximum column width in characters.
    #[arg(short = 'w', long = "max-width", default_value_t = 30)]
    max_width: usize,

    /// Export format. Omit for the boxed terminal preview.
    #[arg(short = 'e', long = "export")]
    export: Option<ExportFormat>,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
enum ExportFormat {
    Csv,
    Json,
    Markdown,
    Text,
    Box,
}

fn main() -> ExitCode {
    let cli = Cli::parse();
    let result = match cli.command {
        Command::Peek(args) => commands::peek::run(args),
        Command::Map(args) => commands::map::run(args.file, args.format),
        Command::Agent(args) => commands::agent::run(args.file, args.max_tokens, args.sheet),
        Command::Schema(args) => commands::schema::run(args.file, args.format, args.sheet),
    };
    match result {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("error: {e:#}");
            ExitCode::from(1)
        }
    }
}
