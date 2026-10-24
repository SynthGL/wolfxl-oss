// SPDX-License-Identifier: MIT
// Independently authored from ODF 1.3 Part 3, sections 1.5 and 3.18.
use std::collections::HashSet;

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::NsReader;

use crate::{ReaderError, Result};

pub(super) const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(super) const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(super) const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(super) const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
pub(super) const MANIFEST: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Name {
    pub uri: String,
    pub local: String,
}

impl Name {
    pub fn is(&self, uri: &str, local: &str) -> bool {
        self.uri == uri && self.local == local
    }
}

pub(super) struct Element {
    pub name: Name,
    attrs: Vec<(Name, String)>,
}

impl Element {
    pub fn get(&self, uri: &str, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(n, _)| n.is(uri, local))
            .map(|(_, v)| v.as_str())
    }

    pub fn required(&self, uri: &str, local: &str) -> Result<&str> {
        self.get(uri, local)
            .ok_or_else(|| malformed(format!("missing {local}")))
    }

    pub fn positive(&self, uri: &str, local: &str) -> Result<u32> {
        match self.get(uri, local) {
            None => Ok(1),
            Some(v) => v
                .trim()
                .parse::<u32>()
                .ok()
                .filter(|v| *v > 0)
                .ok_or_else(|| malformed(format!("invalid {local}"))),
        }
    }
}

pub(super) enum Token<'a> {
    Start(&'a Element),
    End(&'a Name),
    Text(&'a str),
}

pub(super) fn malformed(message: impl Into<String>) -> ReaderError {
    ReaderError::Xml(format!("ODS: {}", message.into()))
}

fn expanded(result: ResolveResult<'_>, local: &str) -> Result<Name> {
    let uri = match result {
        ResolveResult::Bound(ns) => ns.as_ref().to_owned(),
        ResolveResult::Unbound => String::new(),
        ResolveResult::Unknown(_) => return Err(malformed("undeclared namespace prefix")),
    };
    let local = local.to_owned();
    Ok(Name { uri, local })
}

#[derive(Clone, Copy)]
pub(super) struct Limits {
    pub bytes: usize,
    pub depth: usize,
    pub events: usize,
    pub attributes: usize,
    pub text: usize,
    pub cells: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            bytes: 64 * 1024 * 1024,
            depth: 128,
            events: 2_000_000,
            attributes: 256,
            text: 16 * 1024 * 1024,
            cells: 2_000_000,
        }
    }
}

fn valid_chars(value: &str) -> Result<()> {
    if value.chars().all(|c| matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
        Ok(())
    } else { Err(malformed("forbidden XML character")) }
}

// Streaming admission applies to every XML member, including unused package parts.
// No DTD, custom entities, network resolution, or unbounded element tree is used.
pub(super) fn walk(
    bytes: &[u8],
    limits: Limits,
    mut visit: impl FnMut(Token<'_>, &[Name]) -> Result<()>,
) -> Result<()> {
    if bytes.len() > limits.bytes {
        return Err(malformed("XML byte budget exceeded"));
    }
    let utf16 = if bytes.starts_with(b"\xff\xfe") || bytes.starts_with(b"<\0") {
        Some(true)
    } else if bytes.starts_with(b"\xfe\xff") || bytes.starts_with(b"\0<") {
        Some(false)
    } else {
        None
    };
    let decoded;
    let source = if let Some(little) = utf16 {
        let data = if bytes.starts_with(b"\xff\xfe") || bytes.starts_with(b"\xfe\xff") {
            &bytes[2..]
        } else {
            bytes
        };
        if data.len() % 2 != 0 {
            return Err(malformed("incomplete UTF-16 code unit"));
        }
        let units = data.chunks_exact(2).map(|b| {
            if little {
                u16::from_le_bytes([b[0], b[1]])
            } else {
                u16::from_be_bytes([b[0], b[1]])
            }
        });
        let mut text = String::new();
        for ch in char::decode_utf16(units) {
            text.push(ch.map_err(|e| malformed(e.to_string()))?);
            if text.len() > limits.bytes {
                return Err(malformed("decoded XML byte budget exceeded"));
            }
        }
        decoded = text;
        decoded.as_str()
    } else {
        std::str::from_utf8(bytes).map_err(|e| malformed(e.to_string()))?
    };
    valid_chars(source)?;
    // XML 1.0 end-of-line normalization precedes entity expansion, so literal
    // CR/LF differs correctly from an explicit &#13; character reference.
    let normalized;
    let source = if source.contains('\r') {
        normalized = source.replace("\r\n", "\n").replace('\r', "\n");
        normalized.as_str()
    } else {
        source
    };
    let mut reader = NsReader::from_reader(source.as_bytes());
    reader.config_mut().expand_empty_elements = true;
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut roots = 0usize;
    let mut events = 0usize;
    let mut text_bytes = 0usize;
    let mut declaration_seen = false;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| malformed("XML event overflow"))?;
        if events > limits.events {
            return Err(malformed("XML event budget exceeded"));
        }
        match reader
            .read_event_into(&mut buffer)
            .map_err(|e| malformed(e.to_string()))?
        {
            Event::Start(e) => {
                if stack.len() >= limits.depth {
                    return Err(malformed("XML depth budget exceeded"));
                }
                if stack.is_empty() {
                    roots += 1;
                    if roots != 1 {
                        return Err(malformed("multiple XML roots"));
                    }
                }
                let (resolved, local) = reader.resolver().resolve_element(e.name());
                let name = expanded(resolved, local.as_ref())?;
                let mut attrs = Vec::new();
                let mut seen = HashSet::new();
                let mut count = 0usize;
                for a in e.attributes() {
                    count += 1;
                    if count > limits.attributes {
                        return Err(malformed("XML attribute budget exceeded"));
                    }
                    let a = a.map_err(|e| malformed(e.to_string()))?;
                    if a.key.as_ref() == "xmlns" || a.key.as_ref().starts_with("xmlns:") {
                        continue;
                    }
                    let (resolved, local) = reader.resolver().resolve_attribute(a.key);
                    let name = expanded(resolved, local.as_ref())?;
                    if !seen.insert((name.uri.clone(), name.local.clone())) {
                        return Err(malformed("duplicate expanded attribute"));
                    }
                    let raw = a.value.as_ref();
                    let normalized;
                    let raw = if raw.contains(['\t', '\n']) {
                        normalized = raw.replace(['\t', '\n'], " ");
                        normalized.as_str()
                    } else {
                        raw
                    };
                    let value = quick_xml::escape::unescape(raw)
                        .map_err(|e| malformed(e.to_string()))?
                        .into_owned();
                    valid_chars(&value)?;
                    attrs.push((name, value));
                }
                let element = Element { name, attrs };
                visit(Token::Start(&element), &stack)?;
                stack.push(element.name);
            }
            Event::End(_) => {
                let name = stack.pop().ok_or_else(|| malformed("unbalanced XML end"))?;
                visit(Token::End(&name), &stack)?;
            }
            Event::Text(e) => {
                let value: &str = &e;
                valid_chars(&value)?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| malformed("XML text overflow"))?;
                if text_bytes > limits.text {
                    return Err(malformed("XML text budget exceeded"));
                }
                if stack.is_empty() && !value.trim().is_empty() {
                    return Err(malformed("text outside XML root"));
                }
                visit(Token::Text(value), &stack)?;
            }
            Event::GeneralRef(e) => {
                let mut utf8 = [0; 4];
                let value = crate::resolve_xml_reference(&e, &mut utf8)
                    .map_err(|e| malformed(e.to_string()))?;
                valid_chars(value)?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| malformed("XML text overflow"))?;
                if text_bytes > limits.text {
                    return Err(malformed("XML text budget exceeded"));
                }
                if stack.is_empty() {
                    return Err(malformed("reference outside XML root"));
                }
                visit(Token::Text(value), &stack)?;
            }
            Event::CData(e) => {
                let value: &str = &e;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| malformed("XML text overflow"))?;
                if text_bytes > limits.text {
                    return Err(malformed("XML text budget exceeded"));
                }
                if stack.is_empty() {
                    return Err(malformed("CDATA outside XML root"));
                }
                visit(Token::Text(value), &stack)?;
            }
            Event::DocType(_) => return Err(malformed("DTDs are not admitted")),
            Event::Decl(e) => {
                if declaration_seen || roots != 0 || !stack.is_empty() {
                    return Err(malformed("misplaced XML declaration"));
                }
                declaration_seen = true;
                if let Some(encoding) = e.encoding() {
                    let encoding = encoding.map_err(|e| malformed(e.to_string()))?;
                    let label = encoding.as_ref();
                    let matches = match utf16 {
                        None => label.eq_ignore_ascii_case("utf-8"),
                        Some(little) => {
                            label.eq_ignore_ascii_case("utf-16")
                                || label.eq_ignore_ascii_case(if little {
                                    "utf-16le"
                                } else {
                                    "utf-16be"
                                })
                        }
                    };
                    if !matches {
                        return Err(malformed(
                            "XML encoding declaration differs from byte encoding",
                        ));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() || roots != 1 {
        return Err(malformed("incomplete XML document"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn independently_fixed_xml_admission_boundaries() {
        let l = Limits {
            bytes: 40,
            depth: 2,
            events: 20,
            attributes: 2,
            text: 3,
            cells: 2,
        };
        assert!(walk(b"<a><b>abc</b></a>", l, |_, _| Ok(())).is_ok());
        for xml in [
            "<a><b><c/></b></a>",
            "<a>abcd</a>",
            "<a x='1' y='2' z='3'/>",
            "<!DOCTYPE a><a/>",
            "<a/><b/>",
            "<a>&unknown;</a>",
            "<a><b></a>",
            "<x:a/>",
            "<a>&#0;</a>",
            "<a x='&#0;'/>",
            "<a>\u{0}</a>",
        ] {
            assert!(walk(xml.as_bytes(), l, |_, _| Ok(())).is_err(), "{xml}");
        }
        assert!(walk(b"<a/>", Limits { events: 2, ..l }, |_, _| Ok(())).is_err());
        assert!(walk(
            b"<a>0123456789012345678901234567890123456789</a>",
            l,
            |_, _| Ok(())
        )
        .is_err());
    }

    #[test]
    fn references_preserve_text_and_enforce_admission_limits() {
        let mut text = String::new();
        walk(
            b"<a xmlns='urn:test' v='A &amp; B'>A &amp; B &#x1D11E; &amp;lt;\r\n&#13;<![CDATA[<x>]]></a>",
            Limits::default(),
            |token, _| {
                match token {
                    Token::Start(e) => {
                        assert!(e.name.is("urn:test", "a"));
                        assert_eq!(e.get("", "v"), Some("A & B"));
                    }
                    Token::Text(value) => text.push_str(value),
                    Token::End(_) => {}
                }
                Ok(())
            },
        )
        .unwrap();
        assert_eq!(text, "A & B \u{1D11E} &lt;\n\r<x>");
        for xml in [
            "<a>&unknown;</a>",
            "<a>&#0;</a>",
            "<a>&#xD800;</a>",
            "&#32;<a/>",
        ] {
            assert!(walk(xml.as_bytes(), Limits::default(), |_, _| Ok(())).is_err());
        }
        assert!(walk(
            b"<a>&#x1D11E;</a>",
            Limits { text: 3, ..Limits::default() },
            |_, _| Ok(()),
        )
        .is_err());
        assert!(walk(
            b"<a>&amp;&amp;&amp;</a>",
            Limits { events: 4, ..Limits::default() },
            |_, _| Ok(()),
        )
        .is_err());
    }
}
