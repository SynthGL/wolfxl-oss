//! Text-node accumulation for quick-xml 0.42 rewrite passes.
//!
//! quick-xml reports one XML text node as a sequence of `Event::Text`
//! chunks split by `Event::GeneralRef` entity references (`&amp;`,
//! `&#10;`, ...). Rewrite passes that transform a text node collect the
//! whole node here, then either rewrite the resolved text or emit the
//! original escaped form untouched.

use quick_xml::escape::{resolve_predefined_entity, EscapeError};
use quick_xml::events::{BytesRef, BytesText};

/// Resolve a general entity reference to the text it stands for.
/// Recognises the same references as `quick_xml::escape::unescape`:
/// character references and the predefined XML entities; anything else is
/// an unrecognised-entity error.
pub(crate) fn resolve_general_ref(r: &BytesRef<'_>) -> quick_xml::Result<std::borrow::Cow<'static, str>> {
    if let Some(ch) = r.resolve_char_ref()? {
        return Ok(std::borrow::Cow::Owned(ch.to_string()));
    }
    let name: &str = r;
    match resolve_predefined_entity(name) {
        Some(text) => Ok(std::borrow::Cow::Borrowed(text)),
        None => Err(EscapeError::UnrecognizedEntity(1..name.len() + 1, name.to_owned()).into()),
    }
}

/// One XML text node, held both in its original escaped form and resolved.
#[derive(Debug, Default)]
pub(crate) struct TextNode {
    raw: String,
    text: String,
    error: Option<quick_xml::Error>,
    started: bool,
}

impl TextNode {
    /// `true` until a text chunk or entity reference has been pushed.
    pub(crate) fn is_empty(&self) -> bool {
        !self.started
    }

    /// Append an escaped text chunk.
    pub(crate) fn push_text(&mut self, t: &BytesText<'_>) {
        self.started = true;
        let chunk: &str = t;
        self.raw.push_str(chunk);
        self.text.push_str(chunk);
    }

    /// Append an entity reference, recording the first resolution error.
    pub(crate) fn push_ref(&mut self, r: &BytesRef<'_>) {
        self.started = true;
        let name: &str = r;
        self.raw.push('&');
        self.raw.push_str(name);
        self.raw.push(';');
        match resolve_general_ref(r) {
            Ok(text) => self.text.push_str(&text),
            Err(err) => {
                if self.error.is_none() {
                    self.error = Some(err);
                }
            }
        }
    }

    /// The node's text with entity references resolved, or the first
    /// unresolvable reference.
    pub(crate) fn unescaped(&self) -> Result<&str, &quick_xml::Error> {
        match &self.error {
            Some(err) => Err(err),
            None => Ok(&self.text),
        }
    }

    /// The node exactly as it appeared in the source, still escaped.
    pub(crate) fn raw(&self) -> &str {
        &self.raw
    }

    /// Reset for the next text node.
    pub(crate) fn clear(&mut self) {
        self.raw.clear();
        self.text.clear();
        self.error = None;
        self.started = false;
    }
}
