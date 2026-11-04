use quick_xml::{events::Event, Reader};
use wolfxl_formula::{reference::parse_ref, tokenizer::render, RefKind, TokenKind, TokenSubKind};

use crate::EditError;

/// Rewrite formula payloads without serializing any surrounding XML.
pub(crate) fn rewrite_xml(
    xml: &str,
    mut rewrite: impl FnMut(&str) -> Result<String, EditError>,
) -> Result<String, EditError> {
    let mut reader = Reader::from_str(xml);
    let mut changes = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(start)
                if matches!(
                    start.local_name().as_ref().as_bytes(),
                    b"f" | b"formula"
                        | b"formula1"
                        | b"formula2"
                        | b"definedName"
                        | b"calculatedColumnFormula"
                        | b"totalsRowFormula"
                ) =>
            {
                let begin = reader.buffer_position() as usize;
                let raw = reader.read_text(start.name())?;
                let end = begin + raw.len();
                let decoded =
                    quick_xml::escape::unescape(&raw).map_err(|e| EditError::Xml(e.to_string()))?;
                let changed = rewrite(&decoded)?;
                if changed != decoded {
                    changes.push((begin, end, quick_xml::escape::escape(&changed).into_owned()));
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let mut out = xml.to_owned();
    for (start, end, text) in changes.into_iter().rev() {
        out.replace_range(start..end, &text);
    }
    Ok(out)
}

fn is_target(value: &str, source: &str, target: &str) -> bool {
    let prefix = match parse_ref(value) {
        RefKind::Cell { sheet, .. }
        | RefKind::Range { sheet, .. }
        | RefKind::RowRange { sheet, .. }
        | RefKind::ColRange { sheet, .. } => sheet,
        _ => return false,
    };
    prefix.map_or_else(
        || source.eq_ignore_ascii_case(target),
        |s| s.name.eq_ignore_ascii_case(target),
    )
}

pub(crate) fn shift(
    formula: &str,
    source: &str,
    target: &str,
    plan: &wolfxl_structural::ShiftPlan,
) -> Result<String, EditError> {
    tokens(formula, |value| {
        if is_target(value, source, target) {
            wolfxl_structural::shift_formula(value, plan)
        } else {
            value.to_owned()
        }
    })
}

fn tokens(formula: &str, mut rewrite: impl FnMut(&str) -> String) -> Result<String, EditError> {
    let wrapped = format!("={formula}");
    let mut tokens = wolfxl_formula::tokenize(&wrapped)
        .map_err(|e| EditError::Unsupported(format!("cannot rewrite formula {formula:?}: {e}")))?;
    for token in &mut tokens {
        if token.kind == TokenKind::Operand && token.subkind == TokenSubKind::Range {
            token.value = rewrite(&token.value);
        }
    }
    Ok(render(&tokens)
        .strip_prefix('=')
        .unwrap_or(formula)
        .to_owned())
}

/// Hide foreign references from the structural crate's coordinate-only translator.
pub(crate) fn protect_foreign(
    xml: &str,
    source: &str,
    prefix: &str,
) -> Result<(String, Vec<String>), EditError> {
    let mut references = Vec::new();
    let protected = rewrite_xml(xml, |formula| {
        tokens(formula, |value| {
            if !is_target(value, source, source)
                && matches!(
                    parse_ref(value),
                    RefKind::Cell { .. }
                        | RefKind::Range { .. }
                        | RefKind::RowRange { .. }
                        | RefKind::ColRange { .. }
                )
            {
                let index = references.len();
                references.push(value.to_owned());
                format!("{prefix}{index}_")
            } else {
                value.to_owned()
            }
        })
    })?;
    Ok((protected, references))
}

pub(crate) fn unprotect(mut xml: String, prefix: &str, references: &[String]) -> String {
    for (index, reference) in references.iter().enumerate() {
        xml = xml.replace(
            &format!("{prefix}{index}_"),
            &quick_xml::escape::escape(reference),
        );
    }
    xml
}

pub(crate) fn rename(formula: &str, old: &str, new: &str) -> Result<String, EditError> {
    tokens(formula, |value| {
        let old_case = match parse_ref(value) {
            RefKind::Cell { sheet: Some(s), .. }
            | RefKind::Range { sheet: Some(s), .. }
            | RefKind::RowRange { sheet: Some(s), .. }
            | RefKind::ColRange { sheet: Some(s), .. }
                if s.name.eq_ignore_ascii_case(old) =>
            {
                s.name
            }
            _ => return value.to_owned(),
        };
        wolfxl_formula::rename_sheet(&format!("={value}"), &old_case, new)
            .trim_start_matches('=')
            .to_owned()
    })
}

pub(crate) fn shift_defined_names(
    xml: &str,
    sheet_names: &[String],
    target: &str,
    plan: &wolfxl_structural::ShiftPlan,
) -> Result<String, EditError> {
    let mut reader = Reader::from_str(xml);
    let mut scopes = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(start) if start.local_name().as_ref().as_bytes() == b"definedName" => {
                let scope = crate::ooxml_util::attr_value(&start, b"localSheetId")
                    .and_then(|id| id.parse::<usize>().ok())
                    .and_then(|id| sheet_names.get(id))
                    .cloned()
                    .unwrap_or_default();
                scopes.push(scope);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let mut scopes = scopes.iter();
    rewrite_xml(xml, |formula| {
        shift(
            formula,
            scopes.next().map_or("", String::as_str),
            target,
            plan,
        )
    })
}
