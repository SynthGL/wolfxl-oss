use super::{
    Area, CellArea, ExcelError, Expr, Formula, Function, Operator, ParseError, Reference,
    TableReference, TableRows, Value, MAX_COLUMN, MAX_ROW,
};

#[derive(Clone, Debug, PartialEq)]
enum TokenKind<'a> {
    Number(f64),
    Word(&'a str),
    Text(String),
    Sheet(String),
    Error(ExcelError),
    Table(TableReference),
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Percent,
    Ampersand,
    Equal,
    NotEqual,
    Less,
    Greater,
    LessEqual,
    GreaterEqual,
    Open,
    Close,
    BraceOpen,
    BraceClose,
    Comma,
    Semicolon,
    Colon,
    Bang,
    End,
}

struct Token<'a> {
    kind: TokenKind<'a>,
    start: usize,
    end: usize,
}

struct Lexer<'a> {
    input: &'a str,
    position: usize,
}

fn syntax(position: usize, message: impl Into<String>) -> ParseError {
    ParseError::Syntax {
        position,
        message: message.into(),
    }
}

impl<'a> Lexer<'a> {
    fn next(&mut self) -> Result<Token<'a>, ParseError> {
        while let Some(ch) = self.input[self.position..].chars().next() {
            if !ch.is_whitespace() {
                break;
            }
            self.position += ch.len_utf8();
        }
        let start = self.position;
        let remaining = &self.input[start..];
        let Some(ch) = remaining.chars().next() else {
            return Ok(Token {
                kind: TokenKind::End,
                start,
                end: start,
            });
        };
        self.position += ch.len_utf8();
        let kind = match ch {
            '+' => TokenKind::Plus,
            '-' => TokenKind::Minus,
            '*' => TokenKind::Star,
            '/' => TokenKind::Slash,
            '^' => TokenKind::Caret,
            '%' => TokenKind::Percent,
            '&' => TokenKind::Ampersand,
            '=' => TokenKind::Equal,
            '(' => TokenKind::Open,
            ')' => TokenKind::Close,
            '{' => TokenKind::BraceOpen,
            '}' => TokenKind::BraceClose,
            ',' => TokenKind::Comma,
            ';' => TokenKind::Semicolon,
            ':' => TokenKind::Colon,
            '!' => TokenKind::Bang,
            '<' => {
                if self.consume('=') {
                    TokenKind::LessEqual
                } else if self.consume('>') {
                    TokenKind::NotEqual
                } else {
                    TokenKind::Less
                }
            }
            '>' => {
                if self.consume('=') {
                    TokenKind::GreaterEqual
                } else {
                    TokenKind::Greater
                }
            }
            '"' | '\'' => {
                let mut text = String::new();
                loop {
                    let Some(next) = self.input[self.position..].chars().next() else {
                        return Err(syntax(start, "unterminated quoted text"));
                    };
                    self.position += next.len_utf8();
                    if next == ch {
                        if self.consume(ch) {
                            text.push(ch);
                        } else {
                            break;
                        }
                    } else {
                        text.push(next);
                    }
                }
                if ch == '"' {
                    TokenKind::Text(text)
                } else {
                    TokenKind::Sheet(text)
                }
            }
            '#' => {
                let error = ExcelError::ALL
                    .into_iter()
                    .find(|error| {
                        remaining
                            .get(..error.as_str().len())
                            .is_some_and(|prefix| prefix.eq_ignore_ascii_case(error.as_str()))
                    })
                    .ok_or_else(|| syntax(start, "unknown error literal"))?;
                self.position = start + error.as_str().len();
                TokenKind::Error(error)
            }
            '0'..='9' | '.'
                if ch != '.' || remaining.as_bytes().get(1).is_some_and(u8::is_ascii_digit) =>
            {
                // Retain the source span so integer tokens can also be row references.
                while self
                    .input
                    .as_bytes()
                    .get(self.position)
                    .is_some_and(u8::is_ascii_digit)
                {
                    self.position += 1;
                }
                if ch != '.' && self.consume('.') {
                    while self
                        .input
                        .as_bytes()
                        .get(self.position)
                        .is_some_and(u8::is_ascii_digit)
                    {
                        self.position += 1;
                    }
                }
                if self.consume('e') || self.consume('E') {
                    if !self.consume('+') {
                        self.consume('-');
                    }
                    let exponent = self.position;
                    while self
                        .input
                        .as_bytes()
                        .get(self.position)
                        .is_some_and(u8::is_ascii_digit)
                    {
                        self.position += 1;
                    }
                    if exponent == self.position {
                        return Err(syntax(start, "missing exponent digits"));
                    }
                }
                let number = self.input[start..self.position]
                    .parse()
                    .map_err(|_| syntax(start, "invalid number"))?;
                TokenKind::Number(number)
            }
            _ if ch.is_alphabetic() || matches!(ch, '$' | '_' | '\\') => {
                while let Some(next) = self.input[self.position..].chars().next() {
                    if !(next.is_alphanumeric() || matches!(next, '$' | '_' | '.' | '\\')) {
                        break;
                    }
                    self.position += next.len_utf8();
                }
                let word = &self.input[start..self.position];
                if self.input[self.position..].starts_with('[') {
                    TokenKind::Table(self.table_reference(word)?)
                } else {
                    TokenKind::Word(word)
                }
            }
            '[' => {
                return Err(syntax(
                    start,
                    "a structured reference needs a table name, as in Table[Column]; \
                     external workbook references are not supported",
                ))
            }
            _ => return Err(syntax(start, format!("unexpected character {ch:?}"))),
        };
        Ok(Token {
            kind,
            start,
            end: self.position,
        })
    }

    fn consume(&mut self, ch: char) -> bool {
        if self.input[self.position..].starts_with(ch) {
            self.position += ch.len_utf8();
            true
        } else {
            false
        }
    }

    fn peek(&self) -> Option<char> {
        self.input[self.position..].chars().next()
    }

    fn skip_spaces(&mut self) {
        while self.peek() == Some(' ') {
            self.position += 1;
        }
    }

    /// The bracketed part of a structured reference after `table`, with the
    /// position at its opening `[`. Excel's forms are `Table[]`,
    /// `Table[Column]`, `Table[#Item]`, `Table[@Column]`, and a list of
    /// bracketed special items followed by one column or column range, as in
    /// `Table[[#Headers],[#Data],[First]:[Last]]`. Spaces may separate the
    /// list's parts.
    fn table_reference(&mut self, table: &str) -> Result<TableReference, ParseError> {
        let open = self.position;
        self.position += 1;
        let after_open = self.position;
        self.skip_spaces();
        let mut specials = Vec::new();
        let mut columns = None;
        match self.peek() {
            Some(']') => self.position += 1,
            Some('@') => {
                self.position += 1;
                specials.push(TableRows::ThisRow);
                self.skip_spaces();
                match self.peek() {
                    Some(']') => self.position += 1,
                    Some('[') => {
                        columns = Some(self.column_item()?);
                        self.skip_spaces();
                        if !self.consume(']') {
                            return Err(syntax(self.position, "expected ']'"));
                        }
                    }
                    _ => {
                        let name = self.table_name_part(open)?;
                        columns = Some((name.clone(), name));
                    }
                }
            }
            Some('#') => {
                self.position += 1;
                specials.push(self.special_item(open)?);
            }
            Some('[') => loop {
                self.skip_spaces();
                let item = self.position;
                if !self.consume('[') {
                    return Err(syntax(item, "expected '[' in a structured reference"));
                }
                if self.consume('#') {
                    if columns.is_some() {
                        return Err(syntax(item, "special items must come before the column"));
                    }
                    specials.push(self.special_item(open)?);
                } else if columns.is_some() {
                    return Err(syntax(
                        item,
                        "a structured reference selects one column or column range",
                    ));
                } else {
                    self.position = item;
                    columns = Some(self.column_item()?);
                }
                self.skip_spaces();
                if self.consume(']') {
                    break;
                }
                if !self.consume(',') {
                    return Err(syntax(
                        self.position,
                        "expected ',' or ']' in a structured reference",
                    ));
                }
            },
            _ => {
                self.position = after_open;
                let name = self.table_name_part(open)?;
                columns = Some((name.clone(), name));
            }
        }
        let rows = match specials.as_slice() {
            [] | [TableRows::Data] => TableRows::Data,
            [rows] => *rows,
            [TableRows::Headers, TableRows::Data] => TableRows::HeadersData,
            [TableRows::Data, TableRows::Totals] => TableRows::DataTotals,
            _ => {
                return Err(syntax(
                    open,
                    "table items combine only as [#Headers],[#Data] or [#Data],[#Totals]",
                ))
            }
        };
        Ok(TableReference {
            table: table.to_owned(),
            rows,
            columns,
        })
    }

    /// A bracketed column, or a column range `[First]:[Last]`, with the
    /// position at the first `[`.
    fn column_item(&mut self) -> Result<(String, String), ParseError> {
        let open = self.position;
        self.position += 1;
        let first = self.table_name_part(open)?;
        self.skip_spaces();
        if !self.consume(':') {
            return Ok((first.clone(), first));
        }
        self.skip_spaces();
        let second = self.position;
        if !self.consume('[') || self.peek() == Some('#') {
            return Err(syntax(second, "expected a bracketed column after ':'"));
        }
        Ok((first, self.table_name_part(second)?))
    }

    /// A special item's name after its `#`, through the closing `]`.
    fn special_item(&mut self, open: usize) -> Result<TableRows, ParseError> {
        let start = self.position;
        let name = self.table_name_part(open)?;
        [
            ("All", TableRows::All),
            ("Data", TableRows::Data),
            ("Headers", TableRows::Headers),
            ("Totals", TableRows::Totals),
            ("This Row", TableRows::ThisRow),
        ]
        .into_iter()
        .find(|(spelling, _)| spelling.eq_ignore_ascii_case(&name))
        .map(|(_, rows)| rows)
        .ok_or_else(|| syntax(start - 1, format!("unknown table item #{name}")))
    }

    /// Text through the next unescaped `]`, which is consumed. A `'` escapes
    /// the character after it, so `'[`, `']`, `'#`, and `''` stand for `[`,
    /// `]`, `#`, and `'` in a column name.
    fn table_name_part(&mut self, open: usize) -> Result<String, ParseError> {
        let mut name = String::new();
        loop {
            let Some(ch) = self.peek() else {
                return Err(syntax(open, "unterminated structured reference"));
            };
            self.position += ch.len_utf8();
            match ch {
                ']' => return Ok(name),
                '[' => {
                    return Err(syntax(
                        self.position - 1,
                        "write '[ for a '[' in a column name",
                    ))
                }
                '\'' => {
                    let Some(escaped) = self.peek() else {
                        return Err(syntax(open, "unterminated structured reference"));
                    };
                    self.position += escaped.len_utf8();
                    name.push(escaped);
                }
                _ => name.push(ch),
            }
        }
    }
}

struct Parser<'a> {
    lexer: Lexer<'a>,
    token: Token<'a>,
}

pub(super) fn parse(input: &str) -> Result<Formula, ParseError> {
    let mut lexer = Lexer { input, position: 0 };
    let token = lexer.next()?;
    let mut parser = Parser { lexer, token };
    if parser.token.kind == TokenKind::Equal {
        parser.advance()?;
    }
    let expression = parser.expression(0)?;
    if parser.token.kind != TokenKind::End {
        return Err(syntax(
            parser.token.start,
            "unexpected token after expression",
        ));
    }
    Ok(Formula { expression })
}

impl<'a> Parser<'a> {
    fn advance(&mut self) -> Result<Token<'a>, ParseError> {
        let next = self.lexer.next()?;
        Ok(std::mem::replace(&mut self.token, next))
    }

    fn expression(&mut self, minimum: u8) -> Result<Expr, ParseError> {
        let mut left = if matches!(self.token.kind, TokenKind::Plus | TokenKind::Minus) {
            let negative = self.advance()?.kind == TokenKind::Minus;
            Expr::Unary {
                negative,
                operand: Box::new(self.expression(7)?),
            }
        } else {
            self.primary()?
        };
        loop {
            if self.token.kind == TokenKind::Percent && minimum <= 6 {
                self.advance()?;
                left = Expr::Percent(Box::new(left));
                continue;
            }
            let Some((operator, precedence)) = binary_operator(&self.token.kind) else {
                break;
            };
            if precedence < minimum {
                break;
            }
            self.advance()?;
            // Excel's exponentiation, like its other binary operators, associates left.
            let right = self.expression(precedence + 1)?;
            left = Expr::Binary {
                operator,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn primary(&mut self) -> Result<Expr, ParseError> {
        let token = self.advance()?;
        let raw = &self.lexer.input[token.start..token.end];
        if matches!(
            token.kind,
            TokenKind::Word(_) | TokenKind::Number(_) | TokenKind::Sheet(_)
        ) && self.token.kind == TokenKind::Bang
        {
            let sheet = match token.kind {
                TokenKind::Sheet(sheet) => sheet,
                _ => raw.to_owned(),
            };
            self.advance()?;
            let reference = self.advance()?;
            if !matches!(reference.kind, TokenKind::Word(_) | TokenKind::Number(_)) {
                return Err(syntax(reference.start, "expected a sheet reference"));
            }
            let raw = &self.lexer.input[reference.start..reference.end];
            return self.reference(raw, Some(sheet), reference.start, true);
        }
        match token.kind {
            TokenKind::Number(number) if self.token.kind != TokenKind::Colon => {
                Ok(Expr::Literal(if number.is_finite() {
                    Value::Number(number)
                } else {
                    Value::Error(ExcelError::Num)
                }))
            }
            TokenKind::Number(_) => self.reference(raw, None, token.start, true),
            TokenKind::Word(name) => {
                if self.token.kind == TokenKind::Open {
                    return self.call(name, token.start);
                }
                if self.token.kind == TokenKind::Colon
                    || endpoint(name).is_some_and(|e| matches!(e, Endpoint::Cell(..)))
                {
                    return self.reference(name, None, token.start, false);
                }
                Ok(Expr::Literal(if name.eq_ignore_ascii_case("TRUE") {
                    Value::Boolean(true)
                } else if name.eq_ignore_ascii_case("FALSE") {
                    Value::Boolean(false)
                } else {
                    Value::Error(ExcelError::Name)
                }))
            }
            TokenKind::Text(text) => Ok(Expr::Literal(Value::Text(text))),
            TokenKind::Error(error) => Ok(Expr::Literal(Value::Error(error))),
            TokenKind::Table(reference) => Ok(Expr::Table(reference)),
            TokenKind::BraceOpen => self.array(token.start),
            TokenKind::Open => {
                let expression = self.expression(0)?;
                if self.token.kind != TokenKind::Close {
                    return Err(syntax(self.token.start, "expected ')'"));
                }
                self.advance()?;
                Ok(expression)
            }
            _ => Err(syntax(
                token.start,
                "expected a value, reference, or function call",
            )),
        }
    }

    fn reference(
        &mut self,
        first: &str,
        sheet: Option<String>,
        position: usize,
        qualified: bool,
    ) -> Result<Expr, ParseError> {
        let Some(first) = endpoint(first) else {
            return if qualified {
                Ok(Expr::Literal(Value::Error(ExcelError::Ref)))
            } else {
                Err(syntax(position, "invalid reference"))
            };
        };
        let area = if self.token.kind == TokenKind::Colon {
            self.advance()?;
            let last = self.advance()?;
            let raw = &self.lexer.input[last.start..last.end];
            if !matches!(last.kind, TokenKind::Number(_) | TokenKind::Word(_)) {
                return Err(syntax(last.start, "expected range endpoint"));
            }
            match (first, endpoint(raw)) {
                (Endpoint::Cell(r1, c1), Some(Endpoint::Cell(r2, c2))) => Area::Cells {
                    first_row: r1.min(r2),
                    last_row: r1.max(r2),
                    first_column: c1.min(c2),
                    last_column: c1.max(c2),
                },
                (Endpoint::Column(a), Some(Endpoint::Column(b))) => Area::Columns {
                    first: a.min(b),
                    last: a.max(b),
                },
                (Endpoint::Row(a), Some(Endpoint::Row(b))) => Area::Rows {
                    first: a.min(b),
                    last: a.max(b),
                },
                _ => {
                    return Err(syntax(
                        last.start,
                        "range endpoints must have the same kind",
                    ))
                }
            }
        } else if let Endpoint::Cell(row, column) = first {
            Area::Cell { row, column }
        } else {
            return Ok(Expr::Literal(Value::Error(ExcelError::Ref)));
        };
        let valid = match area {
            Area::Cell { row, column } => valid_row(row) && valid_column(column),
            Area::Cells {
                first_row,
                last_row,
                first_column,
                last_column,
            } => {
                valid_row(first_row)
                    && valid_row(last_row)
                    && valid_column(first_column)
                    && valid_column(last_column)
            }
            Area::Columns { first, last } => valid_column(first) && valid_column(last),
            Area::Rows { first, last } => valid_row(first) && valid_row(last),
        };
        Ok(if valid {
            Expr::Reference(Reference { sheet, area })
        } else {
            Expr::Literal(Value::Error(ExcelError::Ref))
        })
    }

    /// An array constant after its `{`: rows separated by `;`, elements by
    /// `,`, each a number with an optional sign, text, a boolean, or an error.
    fn array(&mut self, position: usize) -> Result<Expr, ParseError> {
        let mut rows: Vec<Vec<Value>> = vec![Vec::new()];
        loop {
            let element = self.advance()?;
            let value = match element.kind {
                TokenKind::Plus | TokenKind::Minus => {
                    let negative = element.kind == TokenKind::Minus;
                    let number = self.advance()?;
                    let TokenKind::Number(number) = number.kind else {
                        return Err(syntax(number.start, "expected a number after the sign"));
                    };
                    Value::Number(if negative { -number } else { number })
                }
                TokenKind::Number(number) => Value::Number(number),
                TokenKind::Text(text) => Value::Text(text),
                TokenKind::Error(error) => Value::Error(error),
                TokenKind::Word(word) if word.eq_ignore_ascii_case("TRUE") => Value::Boolean(true),
                TokenKind::Word(word) if word.eq_ignore_ascii_case("FALSE") => {
                    Value::Boolean(false)
                }
                _ => {
                    return Err(syntax(
                        element.start,
                        "array constants hold only numbers, text, booleans, and errors",
                    ))
                }
            };
            let value = match value {
                Value::Number(number) if !number.is_finite() => Value::Error(ExcelError::Num),
                value => value,
            };
            rows.last_mut().expect("a current row").push(value);
            let separator = self.advance()?;
            match separator.kind {
                TokenKind::Comma => {}
                TokenKind::Semicolon => rows.push(Vec::new()),
                TokenKind::BraceClose => break,
                _ => return Err(syntax(separator.start, "expected ',', ';', or '}'")),
            }
        }
        if rows.iter().any(|row| row.len() != rows[0].len()) {
            return Err(syntax(
                position,
                "array constant rows must have the same number of values",
            ));
        }
        Ok(Expr::Literal(Value::Array(rows)))
    }

    fn call(&mut self, name: &str, position: usize) -> Result<Expr, ParseError> {
        let function = Function::from_name(name).ok_or_else(|| ParseError::Unsupported {
            function: name.to_owned(),
        })?;
        self.advance()?;
        let mut arguments = Vec::new();
        if self.token.kind != TokenKind::Close {
            loop {
                // An empty slot, as in `XLOOKUP(x,a,b,,1)` or `IF(c,1,)`, is
                // an omitted argument.
                arguments.push(
                    if matches!(self.token.kind, TokenKind::Comma | TokenKind::Close) {
                        Expr::Missing
                    } else {
                        self.expression(0)?
                    },
                );
                if self.token.kind != TokenKind::Comma {
                    break;
                }
                self.advance()?;
            }
        }
        if self.token.kind != TokenKind::Close {
            return Err(syntax(self.token.start, "expected ',' or ')'"));
        }
        self.advance()?;
        if let Some(message) = function.arity_error(arguments.len()) {
            return Err(syntax(position, message));
        }
        Ok(Expr::Call {
            function,
            arguments,
        })
    }
}

fn binary_operator(token: &TokenKind<'_>) -> Option<(Operator, u8)> {
    Some(match token {
        TokenKind::Equal => (Operator::Equal, 1),
        TokenKind::NotEqual => (Operator::NotEqual, 1),
        TokenKind::Less => (Operator::Less, 1),
        TokenKind::Greater => (Operator::Greater, 1),
        TokenKind::LessEqual => (Operator::LessEqual, 1),
        TokenKind::GreaterEqual => (Operator::GreaterEqual, 1),
        TokenKind::Ampersand => (Operator::Concat, 2),
        TokenKind::Plus => (Operator::Add, 3),
        TokenKind::Minus => (Operator::Subtract, 3),
        TokenKind::Star => (Operator::Multiply, 4),
        TokenKind::Slash => (Operator::Divide, 4),
        TokenKind::Caret => (Operator::Power, 5),
        _ => return None,
    })
}

enum Endpoint {
    Cell(u32, u32),
    Column(u32),
    Row(u32),
}

fn endpoint(raw: &str) -> Option<Endpoint> {
    let raw = raw.strip_prefix('$').unwrap_or(raw);
    let letters = raw.bytes().take_while(u8::is_ascii_alphabetic).count();
    if letters == 0 {
        return (!raw.is_empty() && raw.bytes().all(|b| b.is_ascii_digit()))
            .then(|| Endpoint::Row(raw.parse().unwrap_or(u32::MAX)));
    }
    let column = raw[..letters].bytes().fold(0u32, |column, ch| {
        column
            .saturating_mul(26)
            .saturating_add(u32::from(ch.to_ascii_uppercase() - b'A' + 1))
    });
    let rest = &raw[letters..];
    if rest.is_empty() {
        return Some(Endpoint::Column(column));
    }
    let row = rest.strip_prefix('$').unwrap_or(rest);
    if row.is_empty() || !row.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    Some(Endpoint::Cell(row.parse().unwrap_or(u32::MAX), column))
}

fn valid_row(row: u32) -> bool {
    (1..=MAX_ROW).contains(&row)
}
fn valid_column(column: u32) -> bool {
    (1..=MAX_COLUMN).contains(&column)
}

/// The rectangle an A1 cell or range covers; `None` for whole rows or
/// columns, mixed endpoints, and positions outside the worksheet.
pub(super) fn cell_area(text: &str) -> Option<CellArea> {
    let (first, last) = text.split_once(':').unwrap_or((text, text));
    let (Endpoint::Cell(r1, c1), Endpoint::Cell(r2, c2)) = (endpoint(first)?, endpoint(last)?)
    else {
        return None;
    };
    let area = CellArea {
        first_row: r1.min(r2),
        first_column: c1.min(c2),
        last_row: r1.max(r2),
        last_column: c1.max(c2),
    };
    let valid = valid_row(area.first_row)
        && valid_row(area.last_row)
        && valid_column(area.first_column)
        && valid_column(area.last_column);
    valid.then_some(area)
}
