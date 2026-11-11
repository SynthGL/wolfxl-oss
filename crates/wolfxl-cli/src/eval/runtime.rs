use std::cmp::Ordering;

use crate::filter::lenient_number;
use crate::formula_data::rounded;

use super::{
    names_match, Area, DataSource, ExcelError, Expr, Function, Operator, Reference, TableReference,
    TableRows, Value, MAX_COLUMN, MAX_ROW,
};

mod arrays;
mod criteria;
mod dates;
mod financial;
mod format;
mod information;
mod logic;
mod lookup;
mod math;
mod random;
mod statistics;
mod text;

/// A computed worksheet value or the error it evaluates to.
type Outcome = Result<Value, ExcelError>;

struct Context<'a, S> {
    source: &'a S,
    current_sheet: &'a str,
}

pub(super) fn evaluate(expression: &Expr, source: &impl DataSource, current_sheet: &str) -> Value {
    Context {
        source,
        current_sheet,
    }
    .value(expression)
}

/// A rectangle of cells, zero-based offsets from `first_row` and
/// `first_column`. Whole-row and whole-column references keep their full
/// worksheet size; reads past the used extent see blanks.
#[derive(Clone, Copy, Debug)]
struct Range<'a> {
    sheet: &'a str,
    first_row: u32,
    first_column: u32,
    rows: u32,
    columns: u32,
    /// Written as a single cell, so it evaluates to a scalar.
    single: bool,
    /// Written as whole rows or columns, so it evaluates to the used part.
    whole: bool,
}

/// What an argument evaluates to before ranges are read: a reference keeps
/// its position so functions such as `OFFSET` and `SUMIF` can use it.
enum Operand<'a> {
    Range(Range<'a>),
    Value(Value),
}

/// A rectangular argument: a range or an array of values.
#[derive(Clone)]
enum Grid<'a> {
    Range(Range<'a>),
    Array(Vec<Vec<Value>>),
}

impl Grid<'_> {
    /// Declared rows and columns.
    fn size(&self) -> (u32, u32) {
        match self {
            Self::Range(range) => (range.rows, range.columns),
            Self::Array(rows) => (
                rows.len() as u32,
                rows.first().map_or(0, |row| row.len() as u32),
            ),
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Statistic {
    Sum,
    Average,
    Count,
    CountA,
    Min,
    Max,
}

impl<'a, S: DataSource> Context<'a, S> {
    /// Evaluate to a value, reading ranges into arrays.
    fn value(&self, expression: &Expr) -> Value {
        let operand = self.operand(expression);
        self.read(operand)
    }

    /// Evaluate to a scalar: a one-value array yields that value and any
    /// larger array is `#VALUE!`.
    fn scalar(&self, expression: &Expr) -> Value {
        match self.operand(expression) {
            Operand::Range(range) if range.rows == 1 && range.columns == 1 => {
                self.cell(range.sheet, range.first_row, range.first_column)
            }
            Operand::Range(_) => Value::Error(ExcelError::Value),
            Operand::Value(Value::Array(rows)) => match <[Vec<Value>; 1]>::try_from(rows) {
                Ok([row]) => match <[Value; 1]>::try_from(row) {
                    Ok([value]) => value,
                    Err(_) => Value::Error(ExcelError::Value),
                },
                Err(_) => Value::Error(ExcelError::Value),
            },
            Operand::Value(value) => value,
        }
    }

    fn number(&self, expression: &Expr) -> Result<f64, ExcelError> {
        number(self.scalar(expression))
    }

    /// The numeric argument at `index`, or `default` when it is omitted. An
    /// empty argument is a blank, so it counts as 0.
    fn number_or(&self, arguments: &[Expr], index: usize, default: f64) -> Result<f64, ExcelError> {
        arguments
            .get(index)
            .map_or(Ok(default), |argument| self.number(argument))
    }

    /// Like [`Self::number_or`], but an empty argument also takes `default`.
    fn number_or_default(
        &self,
        arguments: &[Expr],
        index: usize,
        default: f64,
    ) -> Result<f64, ExcelError> {
        supplied(arguments, index).map_or(Ok(default), |argument| self.number(argument))
    }

    fn text(&self, expression: &Expr) -> Result<String, ExcelError> {
        text(self.scalar(expression))
    }

    /// The argument as a range or array; errors propagate and scalars become
    /// one-value arrays.
    fn grid(&self, expression: &Expr) -> Result<Grid<'a>, ExcelError> {
        match self.operand(expression) {
            Operand::Range(range) => Ok(Grid::Range(range)),
            Operand::Value(Value::Array(rows)) => Ok(Grid::Array(rows)),
            Operand::Value(Value::Error(error)) => Err(error),
            Operand::Value(value) => Ok(Grid::Array(vec![vec![value]])),
        }
    }

    /// The value at a zero-based offset; positions outside the grid are blank.
    fn at(&self, grid: &Grid<'_>, row: u32, column: u32) -> Value {
        match grid {
            Grid::Range(range) if row < range.rows && column < range.columns => self.cell(
                range.sheet,
                range.first_row + row,
                range.first_column + column,
            ),
            Grid::Range(_) => Value::Blank,
            Grid::Array(rows) => rows
                .get(row as usize)
                .and_then(|values| values.get(column as usize))
                .cloned()
                .unwrap_or(Value::Blank),
        }
    }

    /// Rows and columns that can hold values: a range is cut at the used
    /// extent, since every cell beyond it is blank.
    fn live(&self, grid: &Grid<'_>) -> (u32, u32) {
        match grid {
            Grid::Range(range) => self.live_range(range),
            array => array.size(),
        }
    }

    fn live_range(&self, range: &Range<'_>) -> (u32, u32) {
        let extent = self.source.used_extent(range.sheet);
        (
            range
                .rows
                .min(extent.rows.saturating_sub(range.first_row - 1)),
            range
                .columns
                .min(extent.columns.saturating_sub(range.first_column - 1)),
        )
    }

    /// Every live value of a grid in row-major order.
    fn values(
        &self,
        grid: &Grid<'_>,
        mut visit: impl FnMut(Value) -> Result<(), ExcelError>,
    ) -> Result<(), ExcelError> {
        let (rows, columns) = self.live(grid);
        for row in 0..rows {
            for column in 0..columns {
                visit(self.at(grid, row, column))?;
            }
        }
        Ok(())
    }

    fn read(&self, operand: Operand<'_>) -> Value {
        match operand {
            Operand::Value(value) => value,
            Operand::Range(range) if range.single => {
                self.cell(range.sheet, range.first_row, range.first_column)
            }
            Operand::Range(range) => {
                let (rows, columns) = if range.whole {
                    self.live_range(&range)
                } else {
                    (range.rows, range.columns)
                };
                if rows == 0 || columns == 0 {
                    return Value::Array(Vec::new());
                }
                Value::Array(
                    (0..rows)
                        .map(|row| {
                            (0..columns)
                                .map(|column| {
                                    self.cell(
                                        range.sheet,
                                        range.first_row + row,
                                        range.first_column + column,
                                    )
                                })
                                .collect()
                        })
                        .collect(),
                )
            }
        }
    }

    fn operand(&self, expression: &Expr) -> Operand<'a> {
        let value = match expression {
            Expr::Literal(value) => value.clone(),
            Expr::Reference(reference) => {
                return match self.resolve(reference) {
                    Ok(range) => Operand::Range(range),
                    Err(error) => Operand::Value(Value::Error(error)),
                }
            }
            Expr::Table(reference) => {
                return match self.resolve_table(reference) {
                    Ok(range) => Operand::Range(range),
                    Err(error) => Operand::Value(Value::Error(error)),
                }
            }
            Expr::Unary { negative, operand } => {
                let negative = *negative;
                elementwise(self.value(operand), |value| match number(value) {
                    Ok(value) => finite(if negative { -value } else { value }),
                    Err(error) => Value::Error(error),
                })
            }
            Expr::Percent(operand) => {
                elementwise(self.value(operand), |value| match number(value) {
                    Ok(value) => finite(value / 100.0),
                    Err(error) => Value::Error(error),
                })
            }
            Expr::Binary {
                operator,
                left,
                right,
            } => {
                let left = self.value(left);
                if matches!(left, Value::Error(_)) {
                    return Operand::Value(left);
                }
                let right = self.value(right);
                if matches!(right, Value::Error(_)) {
                    return Operand::Value(right);
                }
                combine(*operator, left, right)
            }
            Expr::Call {
                function,
                arguments,
            } => return self.call(*function, arguments),
            Expr::Missing => Value::Blank,
        };
        Operand::Value(value)
    }

    fn call(&self, function: Function, arguments: &[Expr]) -> Operand<'a> {
        use Function::*;
        let outcome = match function {
            Acos | Asin | Atan | Atan2 | Cos | Sin | Tan | Degrees | Radians | Pi | Exp | Ln
            | Log | Log10 | Ceiling | Floor | Trunc | Product => self.math(function, arguments),
            IsBlank | IsError | IsNa | IsNumber | IsText | Na => {
                self.information(function, arguments)
            }
            Rand | RandBetween => self.random(function, arguments),
            If => return self.branch(arguments),
            IfError => return self.if_error(arguments),
            Choose => return self.choose(arguments),
            Index => return self.index(arguments),
            Offset => return self.offset(arguments),
            XLookup => return self.xlookup(arguments),
            IfNa => return self.if_na(arguments),
            Ifs => return self.ifs(arguments),
            Switch => return self.switch(arguments),
            Filter => return self.filter(arguments),
            Sum => self.aggregate(Statistic::Sum, arguments),
            Average => self.aggregate(Statistic::Average, arguments),
            Count => self.aggregate(Statistic::Count, arguments),
            CountA => self.aggregate(Statistic::CountA, arguments),
            Min => self.aggregate(Statistic::Min, arguments),
            Max => self.aggregate(Statistic::Max, arguments),
            Median => self.median(arguments),
            Mode | ModeSngl => self.mode(arguments),
            StDev | StDevDotS | StDevP | StDevDotP | Var | VarDotS | VarP | VarDotP => {
                self.dispersion(function, arguments)
            }
            Percentile | PercentileInc => self.percentile(arguments, false),
            PercentileExc => self.percentile(arguments, true),
            Quartile | QuartileInc => self.quartile(arguments, false),
            QuartileExc => self.quartile(arguments, true),
            Rank | RankEq => self.rank(arguments, false),
            RankAvg => self.rank(arguments, true),
            Large => self.kth(arguments, true),
            Small => self.kth(arguments, false),
            CountBlank => self.count_blank(arguments),
            SumProduct => self.sum_product(arguments),
            Subtotal => self.subtotal(arguments),
            Unique => self.unique(arguments),
            Sort => self.sort(arguments),
            And | Or => self.logical(function == And, arguments),
            Not => condition(self.scalar(&arguments[0])).map(|value| Value::Boolean(!value)),
            Abs => self.number(&arguments[0]).map(|value| finite(value.abs())),
            Int => self
                .number(&arguments[0])
                .map(|value| finite(value.floor())),
            Sign => self.number(&arguments[0]).map(|value| {
                Value::Number(if value > 0.0 {
                    1.0
                } else if value < 0.0 {
                    -1.0
                } else {
                    0.0
                })
            }),
            Sqrt => self.number(&arguments[0]).map(|value| {
                if value < 0.0 {
                    Value::Error(ExcelError::Num)
                } else {
                    finite(value.sqrt())
                }
            }),
            Mod => self.modulo(arguments),
            Power => self.number(&arguments[0]).and_then(|base| {
                let exponent = self.number(&arguments[1])?;
                Ok(binary(
                    Operator::Power,
                    Value::Number(base),
                    Value::Number(exponent),
                ))
            }),
            Round => self.round(arguments, Rounding::Nearest),
            RoundDown => self.round(arguments, Rounding::Down),
            RoundUp => self.round(arguments, Rounding::Up),
            CountIf => self.count_if(arguments),
            CountIfs => self.conditional(function, arguments),
            SumIf => self.sum_if(arguments, false),
            AverageIf => self.sum_if(arguments, true),
            SumIfs | AverageIfs | MinIfs | MaxIfs => self.conditional(function, arguments),
            Match => self.match_position(arguments),
            VLookup => self.table_lookup(arguments, true),
            HLookup => self.table_lookup(arguments, false),
            Pmt => self.pmt(arguments),
            Pv => self.pv(arguments),
            Fv => self.fv(arguments),
            Npv => self.npv(arguments),
            Irr => self.irr(arguments),
            Sln => self.sln(arguments),
            Db => self.db(arguments),
            Xnpv => self.xnpv(arguments),
            Xirr => self.xirr(arguments),
            Rate => self.rate(arguments),
            Nper => self.nper(arguments),
            Concat => self.concat(arguments),
            TextJoin => self.text_join(arguments),
            NumberValue => self.number_value(arguments),
            Concatenate => self.concatenate(arguments),
            Exact => self.exact(arguments),
            Find => self.find(arguments),
            Left => self.left(arguments),
            Right => self.right(arguments),
            Mid => self.mid(arguments),
            Len => self.len(arguments),
            Lower => self
                .text(&arguments[0])
                .map(|text| Value::Text(text.to_lowercase())),
            Upper => self
                .text(&arguments[0])
                .map(|text| Value::Text(text.to_uppercase())),
            Rept => self.rept(arguments),
            Substitute => self.substitute(arguments),
            Trim => self.trim(arguments),
            Text => self.format_text(arguments),
            Date => self.date(arguments),
            Year | Month | Day => self.date_part(function, arguments),
            Hour | Minute | Second => self.time_part(function, arguments),
            Days => self.days(arguments),
            EDate => self.edate(arguments, false),
            EoMonth => self.edate(arguments, true),
            NetworkDays => self.network_days(arguments),
            WorkDay => self.work_day(arguments),
            Today => Ok(self.today(false)),
            Now => Ok(self.today(true)),
        };
        Operand::Value(outcome.unwrap_or_else(Value::Error))
    }

    /// A function result: a single cell or one-value array reads as its
    /// value, with blank as zero, while larger ranges and arrays stay whole.
    fn result(&self, operand: Operand<'a>) -> Operand<'a> {
        let value = match operand {
            Operand::Range(range) if range.single => {
                self.cell(range.sheet, range.first_row, range.first_column)
            }
            Operand::Value(Value::Array(mut rows)) if rows.len() == 1 && rows[0].len() == 1 => {
                rows[0].pop().expect("one value")
            }
            Operand::Value(value) => value,
            range => return range,
        };
        Operand::Value(match value {
            Value::Blank => Value::Number(0.0),
            value => value,
        })
    }

    fn branch(&self, arguments: &[Expr]) -> Operand<'a> {
        match condition(self.scalar(&arguments[0])) {
            Ok(selected) => match arguments.get(if selected { 1 } else { 2 }) {
                Some(branch) => self.result(self.operand(branch)),
                None => Operand::Value(Value::Boolean(false)),
            },
            Err(error) => Operand::Value(Value::Error(error)),
        }
    }

    fn if_error(&self, arguments: &[Expr]) -> Operand<'a> {
        let first = self.result(self.operand(&arguments[0]));
        if matches!(first, Operand::Value(Value::Error(_))) {
            self.result(self.operand(&arguments[1]))
        } else {
            first
        }
    }

    fn choose(&self, arguments: &[Expr]) -> Operand<'a> {
        let index = match self.number(&arguments[0]) {
            Ok(index) => index.trunc(),
            Err(error) => return Operand::Value(Value::Error(error)),
        };
        if index < 1.0 || index >= arguments.len() as f64 {
            return Operand::Value(Value::Error(ExcelError::Value));
        }
        self.result(self.operand(&arguments[index as usize]))
    }

    fn aggregate(&self, statistic: Statistic, arguments: &[Expr]) -> Outcome {
        let mut aggregate = Aggregate::new(statistic);
        for argument in arguments {
            match self.operand(argument) {
                // Read references cell by cell rather than materializing them.
                Operand::Range(range) => {
                    self.values(&Grid::Range(range), |value| aggregate.add(value, true))?
                }
                Operand::Value(value) => aggregate.add(value, false)?,
            }
        }
        Ok(aggregate.finish())
    }

    /// `AND` and `OR`: referenced text and blanks are skipped, and with no
    /// logical values at all the result is `#VALUE!`.
    fn logical(&self, all: bool, arguments: &[Expr]) -> Outcome {
        let mut seen = false;
        let mut result = all;
        let mut visit = |value: Value, referenced: bool| {
            let truth = match value {
                Value::Boolean(value) => value,
                Value::Number(value) => value != 0.0,
                Value::Error(error) => return Err(error),
                Value::Text(text) if !referenced => {
                    if text.eq_ignore_ascii_case("TRUE") {
                        true
                    } else if text.eq_ignore_ascii_case("FALSE") {
                        false
                    } else {
                        return Err(ExcelError::Value);
                    }
                }
                _ => return Ok(()),
            };
            seen = true;
            if all {
                result &= truth;
            } else {
                result |= truth;
            }
            Ok(())
        };
        for argument in arguments {
            match self.operand(argument) {
                Operand::Range(range) => {
                    self.values(&Grid::Range(range), |value| visit(value, true))?
                }
                Operand::Value(Value::Array(rows)) => {
                    for value in rows.into_iter().flatten() {
                        visit(value, true)?;
                    }
                }
                Operand::Value(value) => visit(value, false)?,
            }
        }
        if seen {
            Ok(Value::Boolean(result))
        } else {
            Err(ExcelError::Value)
        }
    }

    fn modulo(&self, arguments: &[Expr]) -> Outcome {
        let dividend = self.number(&arguments[0])?;
        let divisor = self.number(&arguments[1])?;
        if divisor == 0.0 {
            return Err(ExcelError::Div0);
        }
        Ok(finite(dividend - divisor * (dividend / divisor).floor()))
    }

    fn round(&self, arguments: &[Expr], mode: Rounding) -> Outcome {
        let number = self.number(&arguments[0])?;
        let digits = self.number(&arguments[1])?;
        Ok(finite(round_decimal(number, digits, mode)))
    }

    fn cell(&self, sheet: &str, row: u32, column: u32) -> Value {
        match self.source.cell_value(sheet, row, column) {
            Value::Number(number) => finite(number),
            Value::Array(_) => Value::Error(ExcelError::Value),
            value => value,
        }
    }

    fn resolve(&self, reference: &Reference) -> Result<Range<'a>, ExcelError> {
        let requested = reference.sheet.as_deref().unwrap_or(self.current_sheet);
        let sheet = self
            .source
            .sheet_names()
            .find(|name| name.eq_ignore_ascii_case(requested))
            .ok_or(ExcelError::Ref)?;
        let (first_row, first_column, rows, columns) = match reference.area {
            Area::Cell { row, column } => (row, column, 1, 1),
            Area::Cells {
                first_row,
                last_row,
                first_column,
                last_column,
            } => (
                first_row,
                first_column,
                last_row - first_row + 1,
                last_column - first_column + 1,
            ),
            Area::Columns { first, last } => (1, first, MAX_ROW, last - first + 1),
            Area::Rows { first, last } => (first, 1, last - first + 1, MAX_COLUMN),
        };
        Ok(Range {
            sheet,
            first_row,
            first_column,
            rows,
            columns,
            single: matches!(reference.area, Area::Cell { .. }),
            whole: matches!(reference.area, Area::Columns { .. } | Area::Rows { .. }),
        })
    }

    /// The rectangle a structured reference selects. Names match ignoring
    /// case. An unknown table or column, or a header or totals row the table
    /// lacks, is `#REF!`; `@` and `[#This Row]` are `#VALUE!`, since a
    /// formula evaluated here has no row of its own.
    fn resolve_table(&self, reference: &TableReference) -> Result<Range<'a>, ExcelError> {
        if reference.rows == TableRows::ThisRow {
            return Err(ExcelError::Value);
        }
        let table = self.source.table(&reference.table).ok_or(ExcelError::Ref)?;
        let area = table.area;
        // Row bounds as half-open intervals: headers, data, then totals.
        let data_first = area.first_row + table.header_rows;
        let totals_first = area.last_row + 1 - table.totals_rows;
        let end = area.last_row + 1;
        let (first_row, end_row) = match reference.rows {
            TableRows::All => (area.first_row, end),
            TableRows::Data => (data_first, totals_first),
            TableRows::Headers if table.header_rows > 0 => (area.first_row, data_first),
            TableRows::Totals if table.totals_rows > 0 => (totals_first, end),
            TableRows::HeadersData if table.header_rows > 0 => (area.first_row, totals_first),
            TableRows::DataTotals if table.totals_rows > 0 => (data_first, end),
            _ => return Err(ExcelError::Ref),
        };
        if first_row >= end_row {
            return Err(ExcelError::Ref);
        }
        let width = area.last_column - area.first_column + 1;
        let (first_column, columns) = match &reference.columns {
            None => (area.first_column, width),
            Some((first, last)) => {
                let index = |name: &str| {
                    table
                        .columns
                        .iter()
                        .take(width as usize)
                        .position(|column| names_match(column, name))
                        .map(|index| index as u32)
                        .ok_or(ExcelError::Ref)
                };
                let (first, last) = (index(first)?, index(last)?);
                (
                    area.first_column + first.min(last),
                    first.abs_diff(last) + 1,
                )
            }
        };
        let rows = end_row - first_row;
        Ok(Range {
            sheet: table.sheet,
            first_row,
            first_column,
            rows,
            columns,
            single: rows == 1 && columns == 1,
            whole: false,
        })
    }
}

fn finite(number: f64) -> Value {
    if number.is_finite() {
        Value::Number(number)
    } else {
        Value::Error(ExcelError::Num)
    }
}

/// The argument at `index`, or `None` when it is omitted or left empty.
fn supplied(arguments: &[Expr], index: usize) -> Option<&Expr> {
    arguments
        .get(index)
        .filter(|argument| !matches!(argument, Expr::Missing))
}

fn number(value: Value) -> Result<f64, ExcelError> {
    match value {
        Value::Number(number) => Ok(number),
        Value::Boolean(value) => Ok(if value { 1.0 } else { 0.0 }),
        Value::Blank => Ok(0.0),
        Value::Text(text) => lenient_number(&text).ok_or(ExcelError::Value),
        Value::Error(error) => Err(error),
        Value::Array(_) => Err(ExcelError::Value),
    }
}

fn condition(value: Value) -> Result<bool, ExcelError> {
    match value {
        Value::Boolean(value) => Ok(value),
        Value::Number(value) => Ok(value != 0.0),
        Value::Blank => Ok(false),
        Value::Error(error) => Err(error),
        Value::Text(_) | Value::Array(_) => Err(ExcelError::Value),
    }
}

fn text(value: Value) -> Result<String, ExcelError> {
    match value {
        Value::Text(text) => Ok(text),
        Value::Number(number) => Ok(general(number)),
        Value::Boolean(value) => Ok(if value { "TRUE" } else { "FALSE" }.to_owned()),
        Value::Blank => Ok(String::new()),
        Value::Error(error) => Err(error),
        Value::Array(_) => Err(ExcelError::Value),
    }
}

/// A number as text, with at most 15 significant digits.
fn general(number: f64) -> String {
    let number = rounded(number);
    if number == 0.0 {
        "0".to_owned()
    } else {
        number.to_string()
    }
}

/// Apply a unary operator to a value, or to each element of an array.
fn elementwise(value: Value, apply: impl Fn(Value) -> Value) -> Value {
    match value {
        Value::Array(rows) => elementwise_rows(rows, apply),
        value => apply(value),
    }
}

/// A binary operator over scalars, or element by element when either side
/// is an array. A scalar pairs with every element, and a single row or
/// column repeats to match the other side; positions outside both shapes
/// are `#N/A`, as in Excel.
fn combine(operator: Operator, left: Value, right: Value) -> Value {
    let (left, right) = match (left, right) {
        (Value::Array(left), Value::Array(right)) => (left, right),
        (Value::Array(left), right) => {
            return elementwise_rows(left, |value| binary(operator, value, right.clone()))
        }
        (left, Value::Array(right)) => {
            return elementwise_rows(right, |value| binary(operator, left.clone(), value))
        }
        (left, right) => return binary(operator, left, right),
    };
    let size = |rows: &[Vec<Value>]| (rows.len(), rows.first().map_or(0, Vec::len));
    let ((left_rows, left_columns), (right_rows, right_columns)) = (size(&left), size(&right));
    // The element at a result position: a single row or column repeats.
    let pick = |rows: &[Vec<Value>], count: usize, width: usize, row: usize, column: usize| {
        let row = if count == 1 { 0 } else { row };
        let column = if width == 1 { 0 } else { column };
        rows.get(row).and_then(|values| values.get(column)).cloned()
    };
    let rows = left_rows.max(right_rows);
    let columns = left_columns.max(right_columns);
    Value::Array(
        (0..rows)
            .map(|row| {
                (0..columns)
                    .map(|column| {
                        match (
                            pick(&left, left_rows, left_columns, row, column),
                            pick(&right, right_rows, right_columns, row, column),
                        ) {
                            (Some(left), Some(right)) => binary(operator, left, right),
                            _ => Value::Error(ExcelError::Na),
                        }
                    })
                    .collect()
            })
            .collect(),
    )
}

fn elementwise_rows(rows: Vec<Vec<Value>>, apply: impl Fn(Value) -> Value) -> Value {
    Value::Array(
        rows.into_iter()
            .map(|row| row.into_iter().map(&apply).collect())
            .collect(),
    )
}

fn binary(operator: Operator, left: Value, right: Value) -> Value {
    use Operator::*;
    match operator {
        Equal | NotEqual | Less | Greater | LessEqual | GreaterEqual => {
            match compare(&left, &right) {
                Ok(ordering) => Value::Boolean(match operator {
                    Equal => ordering == Ordering::Equal,
                    NotEqual => ordering != Ordering::Equal,
                    Less => ordering == Ordering::Less,
                    Greater => ordering == Ordering::Greater,
                    LessEqual => ordering != Ordering::Greater,
                    GreaterEqual => ordering != Ordering::Less,
                    _ => unreachable!(),
                }),
                Err(error) => Value::Error(error),
            }
        }
        Concat => match text(left).and_then(|mut left| {
            left.push_str(&text(right)?);
            Ok(left)
        }) {
            Ok(text) => Value::Text(text),
            Err(error) => Value::Error(error),
        },
        _ => match number(left).and_then(|left| number(right).map(|right| (left, right))) {
            Ok((left, right)) => match operator {
                Add => finite(left + right),
                Subtract => finite(left - right),
                Multiply => finite(left * right),
                Divide if right == 0.0 => Value::Error(ExcelError::Div0),
                Divide => finite(left / right),
                Power if left == 0.0 && right < 0.0 => Value::Error(ExcelError::Div0),
                Power if left == 0.0 && right == 0.0 => Value::Error(ExcelError::Num),
                Power => finite(left.powf(right)),
                _ => unreachable!(),
            },
            Err(error) => Value::Error(error),
        },
    }
}

fn compare(left: &Value, right: &Value) -> Result<Ordering, ExcelError> {
    use Value::*;
    let ordering = match (left, right) {
        (Error(error), _) | (_, Error(error)) => return Err(*error),
        (Array(_), _) | (_, Array(_)) => return Err(ExcelError::Value),
        (Blank, Blank) => Ordering::Equal,
        (Blank, Number(right)) => 0.0_f64.partial_cmp(right).ok_or(ExcelError::Num)?,
        (Number(left), Blank) => left.partial_cmp(&0.0).ok_or(ExcelError::Num)?,
        (Blank, Text(right)) => "".cmp(right.as_str()),
        (Text(left), Blank) => left.as_str().cmp(""),
        (Blank, Boolean(right)) => false.cmp(right),
        (Boolean(left), Blank) => left.cmp(&false),
        (Number(left), Number(right)) => left.partial_cmp(right).ok_or(ExcelError::Num)?,
        (Text(left), Text(right)) => left.to_lowercase().cmp(&right.to_lowercase()),
        (Boolean(left), Boolean(right)) => left.cmp(right),
        (Number(_), Text(_) | Boolean(_)) | (Text(_), Boolean(_)) => Ordering::Less,
        (Text(_) | Boolean(_), Number(_)) | (Boolean(_), Text(_)) => Ordering::Greater,
    };
    Ok(ordering)
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Rounding {
    /// Halves away from zero.
    Nearest,
    /// Toward zero.
    Down,
    /// Away from zero.
    Up,
}

/// Round to `digits` decimal places (negative digits round left of the
/// point) as Excel does: on the 15-significant-digit decimal form of the
/// number, so `2.675` rounds to `2.68` even though its binary value is lower.
fn round_decimal(number: f64, digits: f64, mode: Rounding) -> f64 {
    if number == 0.0 || !number.is_finite() {
        return number;
    }
    // Beyond these bounds every finite double is either kept or lost whole.
    let digits = digits.trunc().clamp(-400.0, 400.0) as i32;
    let decimal = format!("{:.14e}", number.abs());
    let (mantissa, exponent) = decimal.split_once('e').expect("exponent notation");
    let exponent: i32 = exponent.parse().expect("integer exponent");
    let significant: Vec<u8> = mantissa
        .bytes()
        .filter(u8::is_ascii_digit)
        .map(|digit| digit - b'0')
        .collect();
    // The leading digit has place value 10^exponent.
    let keep = exponent + 1 + digits;
    if keep >= significant.len() as i32 {
        return rounded(number);
    }
    let (kept, dropped) = significant.split_at(keep.max(0) as usize);
    let up = match mode {
        Rounding::Nearest => keep >= 0 && dropped[0] >= 5,
        Rounding::Down => false,
        Rounding::Up => dropped.iter().any(|&digit| digit != 0),
    };
    let mut integer = kept
        .iter()
        .fold(0_u64, |integer, &digit| integer * 10 + u64::from(digit));
    integer += u64::from(up);
    if integer == 0 {
        return 0.0;
    }
    let magnitude: f64 = format!("{integer}e{}", -digits)
        .parse()
        .expect("decimal number");
    magnitude.copysign(number)
}

struct Aggregate {
    statistic: Statistic,
    count: u64,
    sum: f64,
    minimum: f64,
    maximum: f64,
}

impl Aggregate {
    fn new(statistic: Statistic) -> Self {
        Self {
            statistic,
            count: 0,
            sum: 0.0,
            minimum: f64::INFINITY,
            maximum: f64::NEG_INFINITY,
        }
    }

    fn add(&mut self, value: Value, referenced: bool) -> Result<(), ExcelError> {
        if let Value::Array(rows) = value {
            for value in rows.into_iter().flatten() {
                self.add(value, true)?;
            }
            return Ok(());
        }
        let value = match (self.statistic, value) {
            (Statistic::CountA, Value::Blank) => return Ok(()),
            (Statistic::CountA, _) => {
                self.count += 1;
                return Ok(());
            }
            (Statistic::Count, Value::Error(_)) => return Ok(()),
            (_, Value::Error(error)) => return Err(error),
            (_, Value::Number(value)) => value,
            (_, Value::Blank) => return Ok(()),
            (_, Value::Text(_) | Value::Boolean(_)) if referenced => return Ok(()),
            (Statistic::Count, Value::Text(text)) => match lenient_number(&text) {
                Some(value) => value,
                None => return Ok(()),
            },
            (_, value) => number(value)?,
        };
        self.count += 1;
        // Do not compute unrelated statistics for COUNT, MIN, or MAX.
        match self.statistic {
            Statistic::Sum | Statistic::Average => self.sum += value,
            Statistic::Min => self.minimum = self.minimum.min(value),
            Statistic::Max => self.maximum = self.maximum.max(value),
            Statistic::Count | Statistic::CountA => {}
        }
        Ok(())
    }

    fn finish(self) -> Value {
        match self.statistic {
            Statistic::Count | Statistic::CountA => Value::Number(self.count as f64),
            Statistic::Sum => finite(self.sum),
            Statistic::Average if self.count == 0 => Value::Error(ExcelError::Div0),
            Statistic::Average => finite(self.sum / self.count as f64),
            Statistic::Min | Statistic::Max if self.count == 0 => Value::Number(0.0),
            Statistic::Min => finite(self.minimum),
            Statistic::Max => finite(self.maximum),
        }
    }
}
