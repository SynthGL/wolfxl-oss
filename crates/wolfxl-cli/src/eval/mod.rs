//! A small, workbook-independent evaluator for scalar formulas and ranges.
//!
//! Coordinates are one-based. Explicit ranges retain their shape, while whole
//! row and column references are intersected with the sheet's used extent.
//! Operators work element by element when an operand is a range or array: a
//! scalar pairs with every element, a single row or column repeats to match
//! the other side, and positions outside both shapes are `#N/A`. Ranges are
//! not implicitly intersected. Array constants such as `{1,2;3,4}` evaluate
//! to arrays, and functions that take ranges (aggregates, criteria, lookups,
//! `NPV`, `IRR`) accept them too.
//! Structured references such as `Sales[Amount]` or `Sales[[#Totals],[Amount]]`
//! resolve through [`DataSource::table`] to the rectangle they select when the
//! formula is evaluated, so they work wherever a range does.

mod parser;
mod runtime;

#[cfg(test)]
mod table_tests;
#[cfg(test)]
mod tests;

use std::fmt;

/// A worksheet value. Array rows are in worksheet order.
#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    Number(f64),
    Text(String),
    Boolean(bool),
    Blank,
    Error(ExcelError),
    Array(Vec<Vec<Value>>),
}

/// The seven traditional worksheet errors.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ExcelError {
    Div0,
    Value,
    Ref,
    Name,
    Na,
    Num,
    Null,
}

impl ExcelError {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Div0 => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Na => "#N/A",
            Self::Num => "#NUM!",
            Self::Null => "#NULL!",
        }
    }

    const ALL: [Self; 7] = [
        Self::Div0,
        Self::Value,
        Self::Ref,
        Self::Name,
        Self::Na,
        Self::Num,
        Self::Null,
    ];

    /// The error spelled `text`, ignoring case.
    fn parse(text: &str) -> Option<Self> {
        Self::ALL
            .into_iter()
            .find(|error| error.as_str().eq_ignore_ascii_case(text))
    }
}

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

/// Last used row and column, inclusive. An empty sheet has zero for both.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct Extent {
    pub rows: u32,
    pub columns: u32,
}

/// Supplies cached cell values, not formulas that need recursive calculation.
///
/// Names passed to `cell_value` and `used_extent` are canonical names returned by
/// `sheet_names`. Missing cells must return `Value::Blank`, as must every cell
/// beyond the used extent. Cell values must be scalar; the evaluator constructs
/// arrays for ranges itself.
pub trait DataSource {
    fn sheet_names(&self) -> impl Iterator<Item = &str>;
    fn cell_value(&self, sheet: &str, row: u32, column: u32) -> Value;
    fn used_extent(&self, sheet: &str) -> Extent;

    /// Whether date serials count days from 1904-01-01 rather than from
    /// 1900-01-01. Date functions read and return serials in this system.
    fn date1904(&self) -> bool {
        false
    }

    /// Whether `row` is hidden. `SUBTOTAL` with a function number from 101
    /// to 111 skips hidden rows; sources that cannot tell report none.
    fn row_hidden(&self, _sheet: &str, _row: u32) -> bool {
        false
    }

    /// Whether the cell holds a formula that calls `SUBTOTAL`, which
    /// `SUBTOTAL` skips so nested subtotals are not counted twice. Sources
    /// without formula text report none.
    fn is_subtotal(&self, _sheet: &str, _row: u32, _column: u32) -> bool {
        false
    }

    /// The worksheet table named `name`, ignoring case as [`names_match`]
    /// does. Sources without tables have none.
    fn table(&self, _name: &str) -> Option<TableInfo<'_>> {
        None
    }
}

/// One-based bounds of a rectangle of cells, inclusive.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct CellArea {
    pub first_row: u32,
    pub first_column: u32,
    pub last_row: u32,
    pub last_column: u32,
}

/// A worksheet table that structured references select from.
#[derive(Clone, Copy, Debug)]
pub struct TableInfo<'a> {
    /// Canonical name, as `sheet_names` returns it, of the sheet holding the
    /// table.
    pub sheet: &'a str,
    /// The whole table, header and totals rows included.
    pub area: CellArea,
    /// Header rows at the top of `area`: 0 or 1.
    pub header_rows: u32,
    /// Totals rows at the bottom of `area`: 0 or 1.
    pub totals_rows: u32,
    /// Column names, left to right.
    pub columns: &'a [String],
}

/// Whether two table or column names are equal ignoring case, as Excel
/// compares them.
pub fn names_match(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}

/// The rectangle an A1 cell or range such as `C2:F15` covers, as table parts
/// store it; `None` for anything else.
pub fn cell_area(text: &str) -> Option<CellArea> {
    parser::cell_area(text)
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct FunctionInfo {
    pub name: &'static str,
    pub category: &'static str,
    /// Recalculate on each evaluation, independently of stored workbook values.
    pub volatile: bool,
}

/// Most arguments a worksheet function call accepts.
const MAX_ARGUMENTS: usize = 255;

/// Declares every supported function once: its variant, worksheet name,
/// category, and the fewest and most arguments a call may pass.
macro_rules! functions {
    ($($variant:ident $name:literal $category:literal $min:literal..=$max:expr;)*) => {
        /// Stable registry for discovery by callers, sorted by name.
        pub const FUNCTIONS: &[FunctionInfo] = &[
            $(FunctionInfo {
                name: $name,
                category: $category,
                volatile: matches!(Function::$variant, Function::Rand | Function::RandBetween | Function::Now | Function::Today | Function::Offset),
            },)*
        ];

        #[derive(Clone, Copy, Debug, Eq, PartialEq)]
        enum Function {
            $($variant,)*
        }

        impl Function {
            fn from_name(name: &str) -> Option<Self> {
                // Files store functions newer than Excel 2007 with `_xlfn.`,
                // and worksheet-only ones such as SORT also with `_xlws.`.
                let mut name = name;
                for marker in ["_xlfn.", "_xlws."] {
                    name = match name.get(..marker.len()) {
                        Some(prefix) if prefix.eq_ignore_ascii_case(marker) => &name[marker.len()..],
                        _ => name,
                    };
                }
                $(if name.eq_ignore_ascii_case($name) {
                    return Some(Self::$variant);
                })*
                None
            }

            const fn name(self) -> &'static str {
                match self {
                    $(Self::$variant => $name,)*
                }
            }

            /// Fewest and most arguments.
            const fn arity(self) -> (usize, usize) {
                match self {
                    $(Self::$variant => ($min, $max),)*
                }
            }
        }
    };
}

functions! {
    Abs "ABS" "math" 1..=1;
    Acos "ACOS" "math" 1..=1;
    And "AND" "logical" 1..=MAX_ARGUMENTS;
    Asin "ASIN" "math" 1..=1;
    Atan "ATAN" "math" 1..=1;
    Atan2 "ATAN2" "math" 2..=2;
    Average "AVERAGE" "statistical" 0..=MAX_ARGUMENTS;
    AverageIf "AVERAGEIF" "statistical" 2..=3;
    AverageIfs "AVERAGEIFS" "statistical" 3..=MAX_ARGUMENTS;
    Ceiling "CEILING" "math" 2..=2;
    Choose "CHOOSE" "lookup" 2..=MAX_ARGUMENTS;
    Concat "CONCAT" "text" 1..=MAX_ARGUMENTS;
    Concatenate "CONCATENATE" "text" 1..=MAX_ARGUMENTS;
    Cos "COS" "math" 1..=1;
    Count "COUNT" "statistical" 0..=MAX_ARGUMENTS;
    CountA "COUNTA" "statistical" 1..=MAX_ARGUMENTS;
    CountBlank "COUNTBLANK" "statistical" 1..=1;
    CountIf "COUNTIF" "statistical" 2..=2;
    CountIfs "COUNTIFS" "statistical" 2..=MAX_ARGUMENTS;
    Date "DATE" "date" 3..=3;
    Day "DAY" "date" 1..=1;
    Days "DAYS" "date" 2..=2;
    Db "DB" "financial" 4..=5;
    Degrees "DEGREES" "math" 1..=1;
    EDate "EDATE" "date" 2..=2;
    EoMonth "EOMONTH" "date" 2..=2;
    Exact "EXACT" "text" 2..=2;
    Exp "EXP" "math" 1..=1;
    Filter "FILTER" "lookup" 2..=3;
    Find "FIND" "text" 2..=3;
    Floor "FLOOR" "math" 2..=2;
    Fv "FV" "financial" 3..=5;
    HLookup "HLOOKUP" "lookup" 3..=4;
    Hour "HOUR" "date" 1..=1;
    If "IF" "logical" 2..=3;
    IfError "IFERROR" "logical" 2..=2;
    IfNa "IFNA" "logical" 2..=2;
    Ifs "IFS" "logical" 2..=254;
    Index "INDEX" "lookup" 2..=3;
    Int "INT" "math" 1..=1;
    Irr "IRR" "financial" 1..=2;
    IsBlank "ISBLANK" "information" 1..=1;
    IsError "ISERROR" "information" 1..=1;
    IsNa "ISNA" "information" 1..=1;
    IsNumber "ISNUMBER" "information" 1..=1;
    IsText "ISTEXT" "information" 1..=1;
    Large "LARGE" "statistical" 2..=2;
    Left "LEFT" "text" 1..=2;
    Len "LEN" "text" 1..=1;
    Ln "LN" "math" 1..=1;
    Log "LOG" "math" 1..=2;
    Log10 "LOG10" "math" 1..=1;
    Lower "LOWER" "text" 1..=1;
    Match "MATCH" "lookup" 2..=3;
    Max "MAX" "statistical" 0..=MAX_ARGUMENTS;
    MaxIfs "MAXIFS" "statistical" 3..=MAX_ARGUMENTS;
    Median "MEDIAN" "statistical" 1..=MAX_ARGUMENTS;
    Mid "MID" "text" 3..=3;
    Min "MIN" "statistical" 0..=MAX_ARGUMENTS;
    MinIfs "MINIFS" "statistical" 3..=MAX_ARGUMENTS;
    Minute "MINUTE" "date" 1..=1;
    Mod "MOD" "math" 2..=2;
    Mode "MODE" "statistical" 1..=MAX_ARGUMENTS;
    ModeSngl "MODE.SNGL" "statistical" 1..=MAX_ARGUMENTS;
    Month "MONTH" "date" 1..=1;
    Na "NA" "information" 0..=0;
    NetworkDays "NETWORKDAYS" "date" 2..=3;
    Not "NOT" "logical" 1..=1;
    Now "NOW" "date" 0..=0;
    Nper "NPER" "financial" 3..=5;
    Npv "NPV" "financial" 2..=MAX_ARGUMENTS;
    Offset "OFFSET" "lookup" 3..=5;
    Or "OR" "logical" 1..=MAX_ARGUMENTS;
    Percentile "PERCENTILE" "statistical" 2..=2;
    PercentileExc "PERCENTILE.EXC" "statistical" 2..=2;
    PercentileInc "PERCENTILE.INC" "statistical" 2..=2;
    Pi "PI" "math" 0..=0;
    Pmt "PMT" "financial" 3..=5;
    Power "POWER" "math" 2..=2;
    Product "PRODUCT" "math" 1..=MAX_ARGUMENTS;
    Pv "PV" "financial" 3..=5;
    Quartile "QUARTILE" "statistical" 2..=2;
    QuartileExc "QUARTILE.EXC" "statistical" 2..=2;
    QuartileInc "QUARTILE.INC" "statistical" 2..=2;
    Radians "RADIANS" "math" 1..=1;
    Rand "RAND" "math" 0..=0;
    RandBetween "RANDBETWEEN" "math" 2..=2;
    Rank "RANK" "statistical" 2..=3;
    RankAvg "RANK.AVG" "statistical" 2..=3;
    RankEq "RANK.EQ" "statistical" 2..=3;
    Rate "RATE" "financial" 3..=6;
    Rept "REPT" "text" 2..=2;
    Right "RIGHT" "text" 1..=2;
    Round "ROUND" "math" 2..=2;
    RoundDown "ROUNDDOWN" "math" 2..=2;
    RoundUp "ROUNDUP" "math" 2..=2;
    Second "SECOND" "date" 1..=1;
    Sign "SIGN" "math" 1..=1;
    Sin "SIN" "math" 1..=1;
    Sln "SLN" "financial" 3..=3;
    Small "SMALL" "statistical" 2..=2;
    Sort "SORT" "lookup" 1..=4;
    Sqrt "SQRT" "math" 1..=1;
    StDev "STDEV" "statistical" 1..=MAX_ARGUMENTS;
    StDevDotP "STDEV.P" "statistical" 1..=MAX_ARGUMENTS;
    StDevDotS "STDEV.S" "statistical" 1..=MAX_ARGUMENTS;
    StDevP "STDEVP" "statistical" 1..=MAX_ARGUMENTS;
    Substitute "SUBSTITUTE" "text" 3..=4;
    Subtotal "SUBTOTAL" "math" 2..=MAX_ARGUMENTS;
    Sum "SUM" "math" 0..=MAX_ARGUMENTS;
    SumIf "SUMIF" "math" 2..=3;
    SumIfs "SUMIFS" "math" 3..=MAX_ARGUMENTS;
    SumProduct "SUMPRODUCT" "math" 1..=MAX_ARGUMENTS;
    Switch "SWITCH" "logical" 3..=MAX_ARGUMENTS;
    Tan "TAN" "math" 1..=1;
    Text "TEXT" "text" 2..=2;
    TextJoin "TEXTJOIN" "text" 3..=MAX_ARGUMENTS;
    Today "TODAY" "date" 0..=0;
    Trim "TRIM" "text" 1..=1;
    Trunc "TRUNC" "math" 1..=2;
    Unique "UNIQUE" "lookup" 1..=3;
    Upper "UPPER" "text" 1..=1;
    NumberValue "VALUE" "text" 1..=1;
    Var "VAR" "statistical" 1..=MAX_ARGUMENTS;
    VarDotP "VAR.P" "statistical" 1..=MAX_ARGUMENTS;
    VarDotS "VAR.S" "statistical" 1..=MAX_ARGUMENTS;
    VarP "VARP" "statistical" 1..=MAX_ARGUMENTS;
    VLookup "VLOOKUP" "lookup" 3..=4;
    WorkDay "WORKDAY" "date" 2..=3;
    Xirr "XIRR" "financial" 2..=3;
    XLookup "XLOOKUP" "lookup" 3..=6;
    Xnpv "XNPV" "financial" 3..=3;
    Year "YEAR" "date" 1..=1;
}

impl Function {
    /// Why `count` arguments are wrong for this function, if they are.
    fn arity_error(self, count: usize) -> Option<String> {
        let (min, max) = self.arity();
        let name = self.name();
        if count < min || count > max {
            return Some(match (min, max) {
                (min, max) if min == max => format!("{name} expects {min} argument(s)"),
                (min, MAX_ARGUMENTS) => format!("{name} expects at least {min} argument(s)"),
                (min, max) => format!("{name} expects {min} to {max} arguments"),
            });
        }
        // Criteria come in (range, criteria) pairs, after the target range
        // for every conditional function but COUNTIFS; IFS takes
        // (condition, value) pairs.
        let (pairs, what) = match self {
            Self::CountIfs => (count, "criteria ranges and criteria"),
            Self::SumIfs | Self::AverageIfs | Self::MinIfs | Self::MaxIfs => {
                (count - 1, "criteria ranges and criteria")
            }
            Self::Ifs => (count, "conditions and values"),
            _ => return None,
        };
        (pairs % 2 != 0).then(|| format!("{name} expects {what} in pairs"))
    }
}

/// Parse failures are distinct from worksheet error values.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ParseError {
    Syntax { position: usize, message: String },
    Unsupported { function: String },
}

impl fmt::Display for ParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Syntax { position, message } => write!(f, "{message} at byte {position}"),
            Self::Unsupported { function } => write!(f, "unsupported function: {function}"),
        }
    }
}

impl std::error::Error for ParseError {}

/// An immutable parsed formula that can be evaluated against multiple sheets.
#[derive(Clone, Debug)]
pub struct Formula {
    expression: Expr,
}

impl Formula {
    pub fn evaluate(&self, source: &impl DataSource, current_sheet: &str) -> Value {
        runtime::evaluate(&self.expression, source, current_sheet)
    }
}

/// Parse formula text, accepting an optional leading equals sign.
pub fn parse(formula: &str) -> Result<Formula, ParseError> {
    parser::parse(formula)
}

/// Parse and evaluate once. Only syntax and unsupported functions return `Err`.
pub fn evaluate(
    formula: &str,
    source: &impl DataSource,
    current_sheet: &str,
) -> Result<Value, ParseError> {
    Ok(parse(formula)?.evaluate(source, current_sheet))
}

const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

#[derive(Clone, Debug)]
enum Expr {
    Literal(Value),
    Reference(Reference),
    /// A structured reference, resolved against the source's tables when it
    /// is evaluated.
    Table(TableReference),
    Unary {
        negative: bool,
        operand: Box<Expr>,
    },
    Percent(Box<Expr>),
    Binary {
        operator: Operator,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    Call {
        function: Function,
        arguments: Vec<Expr>,
    },
    /// An empty argument slot. It evaluates as a blank cell, except where a
    /// function documents that an empty argument takes its default.
    Missing,
}

#[derive(Clone, Copy, Debug)]
enum Operator {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Concat,
    Equal,
    NotEqual,
    Less,
    Greater,
    LessEqual,
    GreaterEqual,
}

#[derive(Clone, Debug)]
struct Reference {
    sheet: Option<String>,
    area: Area,
}

#[derive(Clone, Copy, Debug)]
enum Area {
    Cell {
        row: u32,
        column: u32,
    },
    Cells {
        first_row: u32,
        last_row: u32,
        first_column: u32,
        last_column: u32,
    },
    Columns {
        first: u32,
        last: u32,
    },
    Rows {
        first: u32,
        last: u32,
    },
}

/// A structured reference such as `Sales[[#Data],[Q1]:[Q4]]`.
#[derive(Clone, Debug, PartialEq)]
struct TableReference {
    table: String,
    rows: TableRows,
    /// First and last selected column names as written; `None` selects every
    /// column.
    columns: Option<(String, String)>,
}

/// The rows a structured reference selects: its special items.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TableRows {
    /// `[#All]`: header, data, and totals rows.
    All,
    /// `[#Data]`, or no special item: the data body.
    Data,
    /// `[#Headers]`.
    Headers,
    /// `[#Totals]`.
    Totals,
    /// `[#Headers],[#Data]`.
    HeadersData,
    /// `[#Data],[#Totals]`.
    DataTotals,
    /// `[#This Row]` or `@`, which needs the formula's own row.
    ThisRow,
}
