//! A small, workbook-independent evaluator for scalar formulas and ranges.
//!
//! Coordinates are one-based. Explicit ranges retain their shape, while whole
//! row and column references are intersected with the sheet's used extent.
//! Range operands are not implicitly intersected or broadcast by operators.

mod parser;
mod runtime;

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
/// `sheet_names`. Missing cells must return `Value::Blank`. Cell values must be
/// scalar; the evaluator constructs arrays for ranges itself.
pub trait DataSource {
    fn sheet_names(&self) -> impl Iterator<Item = &str>;
    fn cell_value(&self, sheet: &str, row: u32, column: u32) -> Value;
    fn used_extent(&self, sheet: &str) -> Extent;
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct FunctionInfo {
    pub name: &'static str,
    pub category: &'static str,
}

/// Stable registry for discovery by callers.
pub const FUNCTIONS: &[FunctionInfo] = &[
    FunctionInfo {
        name: "SUM",
        category: "math",
    },
    FunctionInfo {
        name: "AVERAGE",
        category: "statistical",
    },
    FunctionInfo {
        name: "COUNT",
        category: "statistical",
    },
    FunctionInfo {
        name: "MIN",
        category: "math",
    },
    FunctionInfo {
        name: "MAX",
        category: "math",
    },
    FunctionInfo {
        name: "IF",
        category: "logical",
    },
];

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

#[derive(Clone, Copy, Debug)]
enum Function {
    Sum,
    Average,
    Count,
    Min,
    Max,
    If,
}

impl Function {
    fn from_name(name: &str) -> Option<Self> {
        match FUNCTIONS
            .iter()
            .find(|info| info.name.eq_ignore_ascii_case(name))?
            .name
        {
            "SUM" => Some(Self::Sum),
            "AVERAGE" => Some(Self::Average),
            "COUNT" => Some(Self::Count),
            "MIN" => Some(Self::Min),
            "MAX" => Some(Self::Max),
            "IF" => Some(Self::If),
            _ => unreachable!(),
        }
    }
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
