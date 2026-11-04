//! Shared OOXML attribute decoding.
use quick_xml::events::BytesStart;

/// Read an attribute and decode XML entity references.
pub fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref().as_bytes() == key {
            if let Ok(v) = quick_xml::escape::unescape(&a.value) {
                return Some(v.to_string());
            }
            return Some(a.value.into_owned());
        }
    }
    None
}
