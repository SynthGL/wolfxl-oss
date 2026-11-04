use std::path::PathBuf;
use std::process::ExitCode;

use clap::{Parser, ValueEnum};

mod commands;
mod derive;
mod eval;
mod filter;
mod formula_data;
mod input;
mod mcp;
mod render;
mod tui;

/// Fast, agent-friendly spreadsheet previews.
///
/// `wolfxl peek <file>` prints a styled, token-efficient view of a workbook or
/// delimited file: box / text / csv / json / markdown output, sheet
/// selection, row and column caps, and width caps.
/// `wolfxl map <file>` prints a one-page summary of every sheet.
/// `wolfxl agent <file> --max-tokens N` composes a token-budgeted briefing.
/// `wolfxl schema <file>` emits per-column type, cardinality, and format.
/// `wolfxl calc <formula> [file]` evaluates a formula over stored cell values.
/// `wolfxl list-functions` lists supported formula functions or checks names.
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
    /// Evaluate one formula against stored values (0 value, 1 Excel error, 2 invalid request).
    Calc(CalcArgs),
    /// List supported formula functions, or check names (exit 1 if any are unsupported).
    ListFunctions(ListFunctionsArgs),
    /// Group rows by header fields and aggregate value fields, with grand totals.
    Pivot(PivotArgs),
    /// Print a style-independent content fingerprint (`wxf1:<sha256 hex>`).
    Fingerprint(FingerprintArgs),
    /// Check a file against a content fingerprint (0 match, 1 mismatch, 2 input failure).
    Verify(VerifyArgs),
    /// Reconcile two tables by a key column, with numeric tolerance.
    Diff(DiffArgs),
    /// Profile every column: types, distinct values, numeric summary, top values.
    Stats(StatsArgs),
    /// Serve file-based spreadsheet tools over MCP on standard input/output.
    Mcp(McpArgs),
}

#[derive(clap::Args, Debug)]
pub(crate) struct McpArgs {
    /// Confine all workbook paths to this directory; relative paths resolve here.
    #[arg(long)]
    root: Option<PathBuf>,

    /// Expose only inspection, calculation, and analysis tools.
    #[arg(long)]
    read_only: bool,
}

#[derive(clap::Args, Debug)]
pub(crate) struct CalcArgs {
    /// Formula to evaluate, with or without a leading `=`.
    formula: String,

    /// Spreadsheet path, or - for standard input (the default).
    file: Option<PathBuf>,

    /// Sheet the formula evaluates on (default: first sheet).
    #[arg(short = 's', long)]
    sheet: Option<String>,

    /// Input format. Required for standard input; overrides a path's extension.
    #[arg(long, value_enum)]
    from: Option<input::InputFormat>,

    /// Print {formula, sheet, type, value} as JSON.
    #[arg(long)]
    json: bool,
}

#[derive(clap::Args, Debug)]
pub(crate) struct ListFunctionsArgs {
    /// Only functions in this category, ignoring case (for example `math`).
    #[arg(long, conflicts_with = "check")]
    category: Option<String>,

    /// Only functions whose name contains this text, ignoring case.
    #[arg(long, conflicts_with = "check")]
    search: Option<String>,

    /// Report names as supported or unsupported; comma-separated or repeated.
    #[arg(long, value_delimiter = ',', value_name = "NAME")]
    check: Option<Vec<String>>,

    /// Print JSON instead of aligned text.
    #[arg(long)]
    json: bool,
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
    /// Spreadsheet path (.xlsx/.xlsm/.xls/.xlsb/.ods/.csv/.tsv/.txt), or - for standard input with --from.
    file: PathBuf,

    /// Sheet name (default: first sheet).
    #[arg(short = 's', long)]
    sheet: Option<String>,

    /// One-based sheet row containing headers; discard rows above it.
    #[arg(long, value_name = "N")]
    header_row: Option<u32>,

    /// Label previews with sheet column letters and source row numbers (not CSV/JSON).
    #[arg(long)]
    addresses: bool,

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

    /// Input format. Required for standard input; overrides a path's extension.
    #[arg(long, value_enum)]
    from: Option<input::InputFormat>,

    /// Append NAME=FORMULA as a column, filled down from the first data row.
    /// Repeat to add columns in order, before filtering and selection.
    #[arg(long, value_name = "NAME=FORMULA")]
    add: Vec<String>,

    /// Keep data rows where COLUMN<op>VALUE holds; op is one of = != < > <= >= ~
    /// (contains). Repeat to combine with AND.
    #[arg(long = "where", value_name = "EXPR")]
    filters: Vec<String>,

    /// Comma-separated headers to show, in the order given.
    #[arg(long, value_delimiter = ',', value_name = "HEADERS")]
    select: Vec<String>,

    /// Browse the (filtered and selected) sheet in a full-screen terminal grid:
    /// arrows/PageUp/PageDown/Home/End/g/G move, / n N search, s S sort,
    /// f value counts, Tab next sheet, q or Esc quit.
    #[arg(short = 'i', long, conflicts_with = "export")]
    interactive: bool,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
enum ExportFormat {
    Csv,
    Json,
    Markdown,
    Text,
    Box,
}

#[derive(clap::Args, Debug)]
pub(crate) struct PivotArgs {
    /// Spreadsheet path, or - for standard input.
    file: PathBuf,

    /// Header(s) to group rows by, comma-separated or repeated.
    #[arg(long, value_delimiter = ',', required = true)]
    rows: Vec<String>,

    /// Header whose distinct values become columns.
    #[arg(long)]
    column: Option<String>,

    /// Aggregations as AGG:FIELD (sum, count, avg, min, max), comma-separated or repeated.
    #[arg(long, value_delimiter = ',', required = true)]
    values: Vec<String>,

    /// Sheet to pivot (default: first sheet).
    #[arg(short = 's', long)]
    sheet: Option<String>,

    /// One-based sheet row containing headers; discard rows above it.
    #[arg(long, value_name = "N")]
    header_row: Option<u32>,

    /// Input format. Required for standard input; overrides a path's extension.
    #[arg(long, value_enum)]
    from: Option<input::InputFormat>,

    /// Output format.
    #[arg(short = 'e', long = "export", value_enum, default_value = "box")]
    export: PivotFormat,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
pub(crate) enum PivotFormat {
    Box,
    Csv,
    Json,
}

#[derive(clap::Args, Debug)]
pub(crate) struct FingerprintArgs {
    /// Path to a spreadsheet (.xlsx/.xlsm/.xlam/.xls/.xlsb/.ods/.csv/.tsv/.txt).
    file: PathBuf,

    /// Print {fingerprint, version, sheets, cells, formulas} as JSON.
    #[arg(long)]
    json: bool,
}

#[derive(clap::Args, Debug)]
pub(crate) struct VerifyArgs {
    /// Path to a spreadsheet (.xlsx/.xlsm/.xlam/.xls/.xlsb/.ods/.csv/.tsv/.txt).
    file: PathBuf,

    /// Expected fingerprint: `wxf1:` followed by 64 hex digits.
    #[arg(long)]
    fingerprint: String,
}

#[derive(clap::Args, Debug)]
pub(crate) struct DiffArgs {
    /// Left spreadsheet path, or - for standard input.
    left: PathBuf,

    /// Right spreadsheet path, or - for standard input.
    right: PathBuf,

    /// Header of the column that identifies a row on both sides.
    #[arg(long)]
    key: String,

    /// Columns to compare, comma-separated (default: every shared non-key column).
    #[arg(long, value_delimiter = ',')]
    compare: Vec<String>,

    /// Largest absolute numeric difference that still counts as within tolerance.
    #[arg(long, default_value_t = 0.0, allow_negative_numbers = true)]
    tolerance: f64,

    /// Key normalization applied on both sides before matching.
    #[arg(long, value_enum, default_value = "none")]
    key_transform: KeyTransform,

    /// Sheet compared on both sides (default: each side's first sheet).
    #[arg(short = 's', long)]
    sheet: Option<String>,

    /// Format of the standard-input side; with two paths, overrides both extensions.
    #[arg(long, value_enum)]
    from: Option<input::InputFormat>,

    /// Output format.
    #[arg(short = 'f', long = "format", value_enum, default_value = "json")]
    format: DiffFormat,

    /// Also exit 1 when differences fall within tolerance.
    #[arg(long)]
    strict_exit: bool,

    /// Print nothing; report through the exit code only.
    #[arg(short = 'q', long)]
    quiet: bool,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
pub(crate) enum KeyTransform {
    /// Match keys exactly as spelled.
    None,
    /// Strip surrounding whitespace.
    Trim,
    /// Strip surrounding whitespace and lowercase.
    Lower,
    /// Keep only ASCII digits.
    Digits,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
pub(crate) enum DiffFormat {
    Json,
    Text,
}

#[derive(clap::Args, Debug)]
pub(crate) struct StatsArgs {
    /// Spreadsheet path, or - for standard input.
    file: PathBuf,

    /// Sheet to profile (default: first sheet).
    #[arg(short = 's', long)]
    sheet: Option<String>,

    /// One-based sheet row containing headers; discard rows above it.
    #[arg(long, value_name = "N")]
    header_row: Option<u32>,

    /// Input format. Required for standard input; overrides a path's extension.
    #[arg(long, value_enum)]
    from: Option<input::InputFormat>,

    /// Keep data rows where COLUMN<op>VALUE holds; op is one of = != < > <= >= ~
    /// (contains). Repeat to combine with AND.
    #[arg(long = "where", value_name = "EXPR")]
    filters: Vec<String>,

    /// Comma-separated headers to profile, in the order given.
    #[arg(long, value_delimiter = ',', value_name = "HEADERS")]
    select: Vec<String>,

    /// Most frequent values listed per column (0 = every distinct value).
    #[arg(long, default_value_t = 5)]
    top: usize,

    /// Output format.
    #[arg(short = 'e', long = "export", value_enum, default_value = "box")]
    export: StatsFormat,
}

#[derive(Copy, Clone, Debug, ValueEnum)]
pub(crate) enum StatsFormat {
    Box,
    Csv,
    Json,
}

fn main() -> ExitCode {
    let cli = Cli::parse();
    match cli.command {
        Command::Peek(args) => commands::peek::run(args),
        Command::Map(args) => finish_legacy(commands::map::run(args.file, args.format)),
        Command::Agent(args) => {
            finish_legacy(commands::agent::run(args.file, args.max_tokens, args.sheet))
        }
        Command::Schema(args) => {
            finish_legacy(commands::schema::run(args.file, args.format, args.sheet))
        }
        Command::Calc(args) => commands::calc::run(args),
        Command::ListFunctions(args) => commands::functions::run(args),
        Command::Pivot(args) => commands::pivot::run(args),
        Command::Fingerprint(args) => commands::fingerprint::fingerprint(args),
        Command::Verify(args) => commands::fingerprint::verify(args),
        Command::Diff(args) => commands::diff::run(args),
        Command::Stats(args) => commands::stats::run(args),
        Command::Mcp(args) => finish_legacy(mcp::run(args)),
    }
}

fn finish_legacy(result: anyhow::Result<()>) -> ExitCode {
    match result {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("error: {error:#}");
            ExitCode::from(1)
        }
    }
}
