//! A1-syntax Excel formula tokenizer.
//!
//! Direct port of openpyxl's `formula/tokenizer.py` (which is itself a
//! port of Eric Bachtal's JavaScript tokenizer, see
//! <http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html>).
//!
//! The tokenizer's job is to split a formula string into a flat,
//! lossless sequence of [`Token`] objects so that:
//!
//! 1. String literals (`"..."`) are recognized and their contents
//!    are NOT mistaken for cell references.
//! 2. Bracket expressions (`[Col1]`, structured table refs) are kept
//!    intact.
//! 3. Whitespace, operators, and separators round-trip byte-for-byte.
//!
//! The tokenizer does NOT classify the *value* of `Operand/Range`
//! tokens further (e.g. it does not split `Sheet2!A1` into a sheet
//! prefix and a cell); that's the job of [`crate::reference`].

use std::fmt;

/// Coarse-grained token kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TokenKind {
    /// The whole formula was a literal (didn't start with `=`).
    Literal,
    /// An operand: cell reference, string, number, logical, or error.
    Operand,
    /// A function open token (`SUM(`).
    Func,
    /// An array open/close (`{` or `}`).
    Array,
    /// A parenthesis open/close (`(` or `)`).
    Paren,
    /// An argument or row separator (`,` or `;`) inside a function/array.
    Sep,
    /// Prefix unary operator (e.g. unary `-`).
    OpPre,
    /// Infix binary operator (`+ - * / ^ & = > < <= >= <> %`).
    OpIn,
    /// Postfix operator (`%`).
    OpPost,
    /// Whitespace token (preserved for round-trip).
    Wspace,
}

/// Operand subtype + open/close + arg/row separator subtype.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TokenSubKind {
    /// No subtype.
    None,
    /// Operand: string literal `"..."`.
    Text,
    /// Operand: numeric literal.
    Number,
    /// Operand: `TRUE` or `FALSE`.
    Logical,
    /// Operand: `#REF!`, `#N/A`, etc.
    Error,
    /// Operand: a cell, range, sheet-qualified ref, or structured-table ref.
    /// (Defined names also land here — disambiguation is in [`crate::reference`].)
    Range,
    /// Func/Array/Paren open token.
    Open,
    /// Func/Array/Paren close token.
    Close,
    /// Sep token: `,` (function-argument separator).
    Arg,
    /// Sep token: `;` (array row separator).
    Row,
}

/// A single token from a formula.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Token {
    /// The exact source text of this token, byte-for-byte.
    pub value: String,
    /// Coarse kind.
    pub kind: TokenKind,
    /// Subtype (or [`TokenSubKind::None`]).
    pub subkind: TokenSubKind,
}

impl Token {
    /// Construct an arbitrary token.
    pub fn new(value: impl Into<String>, kind: TokenKind, subkind: TokenSubKind) -> Self {
        Self {
            value: value.into(),
            kind,
            subkind,
        }
    }

    /// Build an `Operand` with the right subkind based on the value's
    /// shape. Mirrors `openpyxl.formula.tokenizer.Token.make_operand`.
    pub fn make_operand(value: impl Into<String>) -> Self {
        let v: String = value.into();
        let subkind = if v.starts_with('"') {
            TokenSubKind::Text
        } else if v.starts_with('#') {
            TokenSubKind::Error
        } else if v == "TRUE" || v == "FALSE" {
            TokenSubKind::Logical
        } else if is_numeric_literal(&v) {
            TokenSubKind::Number
        } else {
            TokenSubKind::Range
        };
        Self {
            value: v,
            kind: TokenKind::Operand,
            subkind,
        }
    }

    fn make_subexp(value: impl Into<String>) -> Self {
        let v: String = value.into();
        let last = v.chars().last().expect("subexp value cannot be empty");
        debug_assert!(matches!(last, '{' | '}' | '(' | ')'));
        let kind = if v == "{" || v == "}" {
            TokenKind::Array
        } else if v == "(" || v == ")" {
            TokenKind::Paren
        } else {
            // Trailing `(` or `)` with prefix → function call.
            TokenKind::Func
        };
        let subkind = if last == ')' || last == '}' {
            TokenSubKind::Close
        } else {
            TokenSubKind::Open
        };
        Self {
            value: v,
            kind,
            subkind,
        }
    }

    fn closer_for(opener: &Token) -> Token {
        debug_assert!(matches!(
            opener.kind,
            TokenKind::Func | TokenKind::Array | TokenKind::Paren
        ));
        debug_assert_eq!(opener.subkind, TokenSubKind::Open);
        let value = if opener.kind == TokenKind::Array {
            "}"
        } else {
            ")"
        };
        Self {
            value: value.to_string(),
            kind: opener.kind,
            subkind: TokenSubKind::Close,
        }
    }

    fn make_separator(value: char) -> Self {
        debug_assert!(value == ',' || value == ';');
        let subkind = if value == ',' {
            TokenSubKind::Arg
        } else {
            TokenSubKind::Row
        };
        Self {
            value: value.to_string(),
            kind: TokenKind::Sep,
            subkind,
        }
    }
}

fn is_numeric_literal(s: &str) -> bool {
    s.parse::<f64>().is_ok()
}

/// Errors emitted by the tokenizer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TokenizeError {
    /// Reached end of formula while inside a string literal.
    UnterminatedString {
        /// Byte offset of the opening quote.
        offset: usize,
    },
    /// Reached end of formula while inside a `[...]` block.
    UnterminatedBracket {
        /// Byte offset of the opening `[`.
        offset: usize,
    },
    /// `#` followed by something that's not one of `ERROR_CODES`.
    InvalidErrorCode {
        /// Byte offset of the `#`.
        offset: usize,
    },
    /// Mismatched `(` and `{` pairs.
    MismatchedSubexp {
        /// Byte offset of the offending closer.
        offset: usize,
    },
    /// Unbalanced subexpression at end-of-formula.
    UnclosedSubexp,
    /// `#`/`'`/`"` in an unexpected position.
    UnexpectedChar {
        /// Byte offset.
        offset: usize,
    },
}

impl fmt::Display for TokenizeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            TokenizeError::UnterminatedString { offset } => {
                write!(f, "unterminated string literal at offset {}", offset)
            }
            TokenizeError::UnterminatedBracket { offset } => {
                write!(f, "unterminated `[` at offset {}", offset)
            }
            TokenizeError::InvalidErrorCode { offset } => {
                write!(f, "invalid `#`-error code at offset {}", offset)
            }
            TokenizeError::MismatchedSubexp { offset } => {
                write!(f, "mismatched ( / {{ pair at offset {}", offset)
            }
            TokenizeError::UnclosedSubexp => {
                f.write_str("unclosed subexpression at end of formula")
            }
            TokenizeError::UnexpectedChar { offset } => {
                write!(f, "unexpected character at offset {}", offset)
            }
        }
    }
}

impl std::error::Error for TokenizeError {}

/// Excel literal-error codes recognized after `#`.
const ERROR_CODES: &[&str] = &[
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
];

/// Characters that terminate an in-progress operand token. Every whitespace
/// byte is one, so a line break after `B1` ends the operand like a space does.
const TOKEN_ENDERS: &[u8] = b",;}) \r\n+-*/^&=><%";

/// Tokenize a formula. Returns the flat list of tokens.
///
/// # Round-trip
///
/// `render(&tokenize(f)?) == f` for every well-formed formula `f`.
pub fn tokenize(formula: &str) -> Result<Vec<Token>, TokenizeError> {
    Tokenizer::new(formula).run()
}

/// Re-emit a token stream as a formula string, byte-identical to the
/// original input (modulo any token mutations performed by the caller).
///
/// If the first token is [`TokenKind::Literal`], its value is returned
/// verbatim. Otherwise the result is `=` + concatenation of every token's
/// `value`, mirroring `openpyxl.formula.tokenizer.Tokenizer.render`.
pub fn render(tokens: &[Token]) -> String {
    if tokens.is_empty() {
        return String::new();
    }
    if tokens[0].kind == TokenKind::Literal {
        return tokens[0].value.clone();
    }
    let cap = 1 + tokens.iter().map(|t| t.value.len()).sum::<usize>();
    let mut out = String::with_capacity(cap);
    out.push('=');
    for t in tokens {
        out.push_str(&t.value);
    }
    out
}

struct Tokenizer<'a> {
    formula: &'a str,
    bytes: &'a [u8],
    items: Vec<Token>,
    token_stack: Vec<usize>, // indices into `items` for opener tokens
    offset: usize,
    token: String,
}

impl<'a> Tokenizer<'a> {
    fn new(formula: &'a str) -> Self {
        Self {
            formula,
            bytes: formula.as_bytes(),
            items: Vec::new(),
            token_stack: Vec::new(),
            offset: 0,
            token: String::new(),
        }
    }

    fn run(mut self) -> Result<Vec<Token>, TokenizeError> {
        if self.formula.is_empty() {
            return Ok(self.items);
        }
        if self.bytes[0] != b'=' {
            self.items.push(Token::new(
                self.formula.to_string(),
                TokenKind::Literal,
                TokenSubKind::None,
            ));
            return Ok(self.items);
        }
        self.offset = 1;

        while self.offset < self.bytes.len() {
            if self.check_scientific_notation() {
                continue;
            }
            let c = self.bytes[self.offset];
            if TOKEN_ENDERS.contains(&c) {
                self.save_token();
            }
            match c {
                b'"' | b'\'' => {
                    let n = self.parse_string()?;
                    self.offset += n;
                }
                b'[' => {
                    let n = self.parse_brackets()?;
                    self.offset += n;
                }
                b'#' => {
                    let n = self.parse_error()?;
                    self.offset += n;
                }
                b' ' | b'\r' | b'\n' => {
                    let n = self.parse_whitespace();
                    self.offset += n;
                }
                b'+' | b'-' | b'*' | b'/' | b'^' | b'&' | b'=' | b'>' | b'<' | b'%' => {
                    let n = self.parse_operator();
                    self.offset += n;
                }
                b'{' | b'(' => {
                    let n = self.parse_opener()?;
                    self.offset += n;
                }
                b')' | b'}' => {
                    let n = self.parse_closer()?;
                    self.offset += n;
                }
                b';' | b',' => {
                    let n = self.parse_separator();
                    self.offset += n;
                }
                _ => {
                    let ch_end = next_char_boundary(self.formula, self.offset);
                    self.token.push_str(&self.formula[self.offset..ch_end]);
                    self.offset = ch_end;
                }
            }
        }
        self.save_token();

        if !self.token_stack.is_empty() {
            return Err(TokenizeError::UnclosedSubexp);
        }
        Ok(self.items)
    }

    fn parse_string(&mut self) -> Result<usize, TokenizeError> {
        self.assert_empty_token(Some(':'))?;
        let delim = self.bytes[self.offset];
        debug_assert!(delim == b'"' || delim == b'\'');
        let start = self.offset;
        let mut i = start + 1;
        loop {
            if i >= self.bytes.len() {
                return Err(TokenizeError::UnterminatedString { offset: start });
            }
            if self.bytes[i] == delim {
                if i + 1 < self.bytes.len() && self.bytes[i + 1] == delim {
                    i += 2;
                } else {
                    i += 1;
                    break;
                }
            } else {
                i += 1;
            }
        }
        let n = i - start;
        let slice = &self.formula[start..i];
        if delim == b'"' {
            self.items.push(Token::make_operand(slice));
        } else {
            self.token.push_str(slice);
        }
        Ok(n)
    }

    fn parse_brackets(&mut self) -> Result<usize, TokenizeError> {
        debug_assert_eq!(self.bytes[self.offset], b'[');
        let start = self.offset;
        let mut depth: i32 = 0;
        let mut i = start;
        while i < self.bytes.len() {
            match self.bytes[i] {
                b'[' => depth += 1,
                b']' => {
                    depth -= 1;
                    if depth == 0 {
                        let end = i + 1;
                        let slice = &self.formula[start..end];
                        self.token.push_str(slice);
                        return Ok(end - start);
                    }
                }
                _ => {}
            }
            i += 1;
        }
        Err(TokenizeError::UnterminatedBracket { offset: start })
    }

    fn parse_error(&mut self) -> Result<usize, TokenizeError> {
        self.assert_empty_token(Some('!'))?;
        debug_assert_eq!(self.bytes[self.offset], b'#');
        let rest = &self.formula[self.offset..];
        for code in ERROR_CODES {
            if rest.starts_with(code) {
                let mut value = std::mem::take(&mut self.token);
                value.push_str(code);
                self.items.push(Token::make_operand(value));
                return Ok(code.len());
            }
        }
        Err(TokenizeError::InvalidErrorCode {
            offset: self.offset,
        })
    }

    /// One whitespace token for a run of spaces and line breaks (LF, CR, or
    /// CRLF), kept byte for byte so rendering restores the run in place.
    fn parse_whitespace(&mut self) -> usize {
        debug_assert!(matches!(self.bytes[self.offset], b' ' | b'\r' | b'\n'));
        let start = self.offset;
        let end = self.bytes[start..]
            .iter()
            .position(|b| !matches!(b, b' ' | b'\r' | b'\n'))
            .map_or(self.bytes.len(), |n| start + n);
        self.items.push(Token::new(
            &self.formula[start..end],
            TokenKind::Wspace,
            TokenSubKind::None,
        ));
        end - start
    }

    fn parse_operator(&mut self) -> usize {
        if self.offset + 1 < self.bytes.len() {
            let two = &self.bytes[self.offset..self.offset + 2];
            if two == b">=" || two == b"<=" || two == b"<>" {
                self.items.push(Token::new(
                    std::str::from_utf8(two).unwrap().to_string(),
                    TokenKind::OpIn,
                    TokenSubKind::None,
                ));
                return 2;
            }
        }
        let c = self.bytes[self.offset] as char;
        debug_assert!(matches!(
            c,
            '%' | '*' | '/' | '^' | '&' | '=' | '>' | '<' | '+' | '-'
        ));
        let token = if c == '%' {
            Token::new("%", TokenKind::OpPost, TokenSubKind::None)
        } else if matches!(c, '*' | '/' | '^' | '&' | '=' | '>' | '<') {
            Token::new(c.to_string(), TokenKind::OpIn, TokenSubKind::None)
        } else if self.items.is_empty() {
            Token::new(c.to_string(), TokenKind::OpPre, TokenSubKind::None)
        } else {
            let prev = self
                .items
                .iter()
                .rev()
                .find(|t| t.kind != TokenKind::Wspace);
            let is_infix = match prev {
                Some(t) => {
                    t.subkind == TokenSubKind::Close
                        || t.kind == TokenKind::OpPost
                        || t.kind == TokenKind::Operand
                }
                None => false,
            };
            if is_infix {
                Token::new(c.to_string(), TokenKind::OpIn, TokenSubKind::None)
            } else {
                Token::new(c.to_string(), TokenKind::OpPre, TokenSubKind::None)
            }
        };
        self.items.push(token);
        1
    }

    fn parse_opener(&mut self) -> Result<usize, TokenizeError> {
        let c = self.bytes[self.offset];
        debug_assert!(c == b'(' || c == b'{');
        let token = if c == b'{' {
            self.assert_empty_token(None)?;
            Token::make_subexp("{")
        } else if !self.token.is_empty() {
            let mut v = std::mem::take(&mut self.token);
            v.push('(');
            Token::make_subexp(v)
        } else {
            Token::make_subexp("(")
        };
        let idx = self.items.len();
        self.items.push(token);
        self.token_stack.push(idx);
        Ok(1)
    }

    fn parse_closer(&mut self) -> Result<usize, TokenizeError> {
        let c = self.bytes[self.offset];
        debug_assert!(c == b')' || c == b'}');
        let opener_idx = self
            .token_stack
            .pop()
            .ok_or(TokenizeError::MismatchedSubexp {
                offset: self.offset,
            })?;
        let opener = self.items[opener_idx].clone();
        let closer = Token::closer_for(&opener);
        if (c == b')' && closer.value != ")") || (c == b'}' && closer.value != "}") {
            return Err(TokenizeError::MismatchedSubexp {
                offset: self.offset,
            });
        }
        self.items.push(closer);
        Ok(1)
    }

    fn parse_separator(&mut self) -> usize {
        let c = self.bytes[self.offset];
        debug_assert!(c == b',' || c == b';');
        let token = if c == b';' {
            Token::make_separator(';')
        } else {
            let inside_paren = self
                .token_stack
                .last()
                .map_or(true, |&idx| self.items[idx].kind == TokenKind::Paren);
            if inside_paren {
                Token::new(",", TokenKind::OpIn, TokenSubKind::None)
            } else {
                Token::make_separator(',')
            }
        };
        self.items.push(token);
        1
    }

    fn check_scientific_notation(&mut self) -> bool {
        if self.offset >= self.bytes.len() {
            return false;
        }
        let c = self.bytes[self.offset];
        if !(c == b'+' || c == b'-') || self.token.is_empty() {
            return false;
        }
        let s = self.token.as_bytes();
        if s.len() < 2 {
            return false;
        }
        if !matches!(s[0], b'1'..=b'9') {
            return false;
        }
        let last = *s.last().unwrap();
        if last != b'E' && last != b'e' {
            return false;
        }
        if s.len() > 2 {
            if s[1] != b'.' {
                return false;
            }
            for &b in &s[2..s.len() - 1] {
                if !b.is_ascii_digit() {
                    return false;
                }
            }
        }
        self.token.push(c as char);
        self.offset += 1;
        true
    }

    fn assert_empty_token(&self, can_follow: Option<char>) -> Result<(), TokenizeError> {
        if self.token.is_empty() {
            return Ok(());
        }
        if let Some(c) = can_follow {
            if let Some(last) = self.token.chars().next_back() {
                if last == c {
                    return Ok(());
                }
            }
        }
        Err(TokenizeError::UnexpectedChar {
            offset: self.offset,
        })
    }

    fn save_token(&mut self) {
        if self.token.is_empty() {
            return;
        }
        let v = std::mem::take(&mut self.token);
        self.items.push(Token::make_operand(v));
    }
}

fn next_char_boundary(s: &str, mut i: usize) -> usize {
    i += 1;
    while !s.is_char_boundary(i) {
        i += 1;
    }
    i
}

#[cfg(test)]
mod tokenizer_tests {
    use super::*;

    fn t(formula: &str) -> Vec<Token> {
        tokenize(formula).expect("tokenize ok")
    }

    #[test]
    fn empty_formula() {
        assert!(t("").is_empty());
    }

    #[test]
    fn literal_formula_no_equals() {
        let toks = t("hello world");
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].kind, TokenKind::Literal);
        assert_eq!(toks[0].value, "hello world");
        assert_eq!(render(&toks), "hello world");
    }

    #[test]
    fn simple_cell_ref() {
        let toks = t("=A1");
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].kind, TokenKind::Operand);
        assert_eq!(toks[0].subkind, TokenSubKind::Range);
        assert_eq!(toks[0].value, "A1");
        assert_eq!(render(&toks), "=A1");
    }

    #[test]
    fn binop_two_cells() {
        let toks = t("=A1+B2");
        assert_eq!(toks.len(), 3);
        assert_eq!(toks[1].kind, TokenKind::OpIn);
        assert_eq!(render(&toks), "=A1+B2");
    }

    #[test]
    fn function_call_and_range() {
        let toks = t("=SUM(A1:B5)");
        assert_eq!(toks.len(), 3);
        assert_eq!(toks[0].kind, TokenKind::Func);
        assert_eq!(toks[0].subkind, TokenSubKind::Open);
        assert_eq!(toks[0].value, "SUM(");
        assert_eq!(toks[1].subkind, TokenSubKind::Range);
        assert_eq!(toks[1].value, "A1:B5");
        assert_eq!(toks[2].kind, TokenKind::Func);
        assert_eq!(toks[2].subkind, TokenSubKind::Close);
        assert_eq!(render(&toks), "=SUM(A1:B5)");
    }

    #[test]
    fn quoted_sheet_with_apostrophe() {
        let f = "='O''Brien'!A1";
        let toks = t(f);
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].subkind, TokenSubKind::Range);
        assert_eq!(toks[0].value, "'O''Brien'!A1");
        assert_eq!(render(&toks), f);
    }

    #[test]
    fn string_literal_preserved() {
        let f = "=IF(A1=\"B5\",X1,Y1)";
        let toks = t(f);
        let text_tok = toks
            .iter()
            .find(|t| t.subkind == TokenSubKind::Text)
            .unwrap();
        assert_eq!(text_tok.value, "\"B5\"");
        assert_eq!(render(&toks), f);
    }

    #[test]
    fn doubled_quotes_in_string() {
        let f = "=\"a\"\"b\"";
        let toks = t(f);
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].subkind, TokenSubKind::Text);
        assert_eq!(toks[0].value, "\"a\"\"b\"");
        assert_eq!(render(&toks), f);
    }

    #[test]
    fn error_literal() {
        let toks = t("=#REF!");
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].subkind, TokenSubKind::Error);
        assert_eq!(toks[0].value, "#REF!");
    }

    #[test]
    fn number_and_logical() {
        let toks = t("=TRUE+1.5");
        assert_eq!(toks[0].subkind, TokenSubKind::Logical);
        assert_eq!(toks[2].subkind, TokenSubKind::Number);
        assert_eq!(render(&toks), "=TRUE+1.5");
    }

    #[test]
    fn scientific_notation_literal() {
        let toks = t("=1.5E+3");
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].subkind, TokenSubKind::Number);
        assert_eq!(toks[0].value, "1.5E+3");
    }

    #[test]
    fn array_constant() {
        let toks = t("={1,2;3,4}");
        assert_eq!(toks.first().unwrap().kind, TokenKind::Array);
        assert_eq!(toks.first().unwrap().subkind, TokenSubKind::Open);
        assert_eq!(toks.last().unwrap().kind, TokenKind::Array);
        assert_eq!(toks.last().unwrap().subkind, TokenSubKind::Close);
        let arg_seps: Vec<_> = toks.iter().filter(|t| t.kind == TokenKind::Sep).collect();
        assert!(arg_seps.iter().any(|t| t.subkind == TokenSubKind::Arg));
        assert!(arg_seps.iter().any(|t| t.subkind == TokenSubKind::Row));
        assert_eq!(render(&toks), "={1,2;3,4}");
    }

    #[test]
    fn whitespace_preserved() {
        let f = "=A1 + B1";
        let toks = t(f);
        assert!(toks.iter().any(|t| t.kind == TokenKind::Wspace));
        assert_eq!(render(&toks), f);
    }

    #[test]
    fn structured_table_ref_brackets() {
        let f = "=Table1[Col1]";
        let toks = t(f);
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].subkind, TokenSubKind::Range);
        assert_eq!(toks[0].value, "Table1[Col1]");
        assert_eq!(render(&toks), f);
    }

    #[test]
    fn nested_table_ref_brackets() {
        let f = "=Table1[[#This Row], [Col1]]";
        let toks = t(f);
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].subkind, TokenSubKind::Range);
        assert_eq!(toks[0].value, "Table1[[#This Row], [Col1]]");
        assert_eq!(render(&toks), f);
    }

    #[test]
    fn unary_minus_prefix() {
        let toks = t("=-A1");
        assert_eq!(toks[0].kind, TokenKind::OpPre);
        assert_eq!(toks[0].value, "-");
        assert_eq!(toks[1].subkind, TokenSubKind::Range);
    }

    #[test]
    fn comparison_two_char() {
        let toks = t("=A1>=B1");
        assert!(toks
            .iter()
            .any(|t| t.value == ">=" && t.kind == TokenKind::OpIn));
    }

    #[test]
    fn three_d_unquoted() {
        let f = "=Sheet2!A1";
        let toks = t(f);
        assert_eq!(toks.len(), 1);
        assert_eq!(toks[0].value, "Sheet2!A1");
        assert_eq!(render(&toks), f);
    }

    #[test]
    fn multi_char_function() {
        let toks = t("=VLOOKUP(A1,B1:C10,2,FALSE)");
        assert_eq!(toks[0].value, "VLOOKUP(");
        assert_eq!(toks[0].kind, TokenKind::Func);
        assert_eq!(render(&toks), "=VLOOKUP(A1,B1:C10,2,FALSE)");
    }

    #[test]
    fn external_book_ref_passthrough() {
        let f = "=[Book2.xlsx]Sheet1!A1";
        let toks = t(f);
        assert_eq!(render(&toks), f);
    }

    #[test]
    fn external_book_ref_quoted() {
        let f = "='C:\\path\\[Book2.xlsx]Sheet1'!A1";
        let toks = t(f);
        assert_eq!(render(&toks), f);
    }
}
