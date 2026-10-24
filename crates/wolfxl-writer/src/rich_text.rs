//! Rich-text run parser for OOXML.
//!
//! Both `xl/sharedStrings.xml` (`<si>...</si>`) and inline-string cells
//! (`<c t="inlineStr"><is>...</is></c>`) carry the *same* run grammar:
//!
//! ```xml
//! <r><rPr>
//!   <b/><i/><u val="single"/><strike/>
//!   <sz val="11"/>
//!   <color rgb="FFFF0000"/>
//!   <rFont val="Calibri"/>
//! </rPr><t xml:space="preserve">hello </t></r>
//! <r><t>world</t></r>
//! ```
//!
//! [`parse_runs`] consumes the *body* of the `<si>` / `<is>` element
//! (i.e. starting from the first `<r>` or `<t>` child) and returns a
//! [`Vec<RichTextRun>`].
//!
//! Sprint Ι Pod-α (RFC pending) — closes the Phase 3 rich-text-reads
//! gap and the implicit T3 rich-text-write deferral.

use std::borrow::Cow;

use quick_xml::escape::{resolve_predefined_entity, EscapeError};
use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesRef, Event};
use quick_xml::Reader as XmlReader;

/// One inline-font property block.  Mirrors the keyword arguments
/// accepted by `wolfxl.cell.rich_text.InlineFont` (and openpyxl's
/// `InlineFont`).  `None` means the attribute was absent — *not*
/// "explicitly false".
#[derive(Debug, Clone, Default, PartialEq)]
pub struct InlineFontProps {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub strike: Option<bool>,
    pub underline: Option<String>,
    pub size: Option<f64>,
    /// ARGB hex (8 chars, no leading `#`) when sourced from `rgb=`.
    /// May also carry a `theme=` / `indexed=` descriptor; we keep the
    /// raw string verbatim for round-trip fidelity.
    pub color: Option<String>,
    pub name: Option<String>,
    pub family: Option<i32>,
    pub charset: Option<i32>,
    pub vert_align: Option<String>,
    pub scheme: Option<String>,
}

/// One rich-text run.
///
/// `font` is `None` for plain runs (`<r><t>...</t></r>` with no
/// `<rPr>`) and for top-level `<t>` text outside any `<r>`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RichTextRun {
    pub text: String,
    pub font: Option<InlineFontProps>,
}

/// Parse the children of a `<si>` or `<is>` element from a full XML
/// fragment.  The fragment must include the wrapping element so
/// quick-xml can locate the run boundaries.
///
/// Returns:
/// * `Ok(None)` when the element contained only a single plain `<t>`
///   (no `<r>` runs — caller should treat the cell as plain).
/// * `Ok(Some(runs))` when at least one `<r>` was present, even if
///   only one (openpyxl flips the cell to `CellRichText` whenever the
///   on-disk shape uses runs).
/// * `Err` on malformed XML.
pub fn parse_runs_in_element(
    xml: &str,
    wrapper_tag: &[u8],
) -> Result<Option<Vec<RichTextRun>>, String> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut in_wrapper = false;
    let mut depth_under_wrapper = 0u32; // 1 inside <si>, 2 inside <r>, etc.

    let mut current_run: Option<RichTextRun> = None;
    let mut current_props: Option<InlineFontProps> = None;
    let mut in_t = false;
    let mut t_buf = String::new();

    let mut runs: Vec<RichTextRun> = Vec::new();
    // True when at least one `<r>` was seen — used to decide between
    // "single plain text" vs. "rich-text cell with one run".
    let mut saw_r = false;
    let mut plain_t_text = String::new();
    let mut plain_t_active = false;

    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(|e| format!("xml read: {e}"))?
        {
            Event::Start(e) => {
                let name = e.local_name();
                let tag = name.as_ref().as_bytes();
                if !in_wrapper && tag == wrapper_tag {
                    in_wrapper = true;
                    continue;
                }
                if !in_wrapper {
                    continue;
                }
                depth_under_wrapper += 1;
                match tag {
                    b"r" => {
                        saw_r = true;
                        current_run = Some(RichTextRun::default());
                        current_props = None;
                    }
                    b"rPr" => {
                        current_props = Some(InlineFontProps::default());
                    }
                    b"t" => {
                        in_t = true;
                        t_buf.clear();
                        // Inline-string single `<t>` (no `<r>` wrapper) —
                        // this is the "plain string" case.
                        if current_run.is_none() {
                            plain_t_active = true;
                            plain_t_text.clear();
                        }
                    }
                    other if current_props.is_some() => {
                        // Inside <rPr>: parse Start-tags that may
                        // carry text (none of the rPr children carry
                        // text — they're all empty/self-closing —
                        // but quick-xml still reports them as Start
                        // when written `<b></b>` instead of `<b/>`).
                        apply_rpr_attr(current_props.as_mut().unwrap(), other, e.attributes());
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let name = e.local_name();
                let tag = name.as_ref().as_bytes();
                if !in_wrapper {
                    continue;
                }
                if let Some(props) = current_props.as_mut() {
                    apply_rpr_attr(props, tag, e.attributes());
                }
            }
            Event::End(e) => {
                let name = e.local_name();
                let tag = name.as_ref().as_bytes();
                if in_wrapper && tag == wrapper_tag {
                    break;
                }
                if !in_wrapper {
                    continue;
                }
                match tag {
                    b"t" => {
                        if let Some(run) = current_run.as_mut() {
                            run.text.push_str(&t_buf);
                        } else if plain_t_active {
                            plain_t_text.push_str(&t_buf);
                            plain_t_active = false;
                        }
                        in_t = false;
                        t_buf.clear();
                    }
                    b"rPr" => {
                        if let (Some(run), Some(props)) =
                            (current_run.as_mut(), current_props.take())
                        {
                            run.font = Some(props);
                        }
                    }
                    b"r" => {
                        if let Some(run) = current_run.take() {
                            runs.push(run);
                        }
                    }
                    _ => {}
                }
                if depth_under_wrapper > 0 {
                    depth_under_wrapper -= 1;
                }
            }
            Event::Text(e) => {
                if in_t {
                    t_buf.push_str(&e);
                }
            }
            Event::GeneralRef(r) => {
                if in_t {
                    let resolved = resolve_general_ref(&r)
                        .map_err(|err| format!("xml text unescape: {err}"))?;
                    t_buf.push_str(&resolved);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if !saw_r {
        // No `<r>` runs — caller should treat as plain text and bypass
        // the rich-text path entirely.  We keep the plain text around
        // so the rare caller that wants to inspect "what would have
        // been the flat string" can still do so via a wrapping helper.
        let _ = plain_t_text; // silence the unused-variable lint
        Ok(None)
    } else {
        Ok(Some(runs))
    }
}

/// Resolve a general entity reference that quick-xml reports between text
/// runs (`&amp;`, `&#10;`, ...) to the text it stands for. Recognises the
/// same references as `quick_xml::escape::unescape`: character references
/// and the predefined XML entities; anything else is an unrecognised-entity
/// error.
fn resolve_general_ref(r: &BytesRef<'_>) -> quick_xml::Result<Cow<'static, str>> {
    if let Some(ch) = r.resolve_char_ref()? {
        return Ok(Cow::Owned(ch.to_string()));
    }
    let name: &str = r;
    match resolve_predefined_entity(name) {
        Some(text) => Ok(Cow::Borrowed(text)),
        None => Err(EscapeError::UnrecognizedEntity(1..name.len() + 1, name.to_owned()).into()),
    }
}

fn apply_rpr_attr(
    props: &mut InlineFontProps,
    tag: &[u8],
    mut attrs: quick_xml::events::attributes::Attributes<'_>,
) {
    // Each rPr child encodes one font property.  Most are self-closing
    // and may carry a `val=` attribute (e.g. `<b val="1"/>` or
    // `<sz val="12"/>`).  Collect every attribute into a small vec so
    // the helpers below can iterate it more than once without
    // re-parsing the XML stream.
    let mut all_attrs: Vec<(Vec<u8>, String)> = Vec::with_capacity(2);
    for a in attrs.with_checks(false) {
        if let Ok(Attribute { key, value }) = a {
            all_attrs.push((
                key.local_name().as_ref().as_bytes().to_vec(),
                value.into_owned(),
            ));
        }
    }
    let val = || -> Option<String> {
        all_attrs
            .iter()
            .rev()
            .find(|(k, _)| k.as_slice() == b"val")
            .map(|(_, v)| v.clone())
    };

    fn parse_bool_val(v: Option<String>) -> bool {
        // `<b/>` (no val) → true.
        // `<b val="1"/>` / `"true"` → true.
        // `<b val="0"/>` / `"false"` → false.
        match v.as_deref() {
            None => true,
            Some(s) => {
                let t = s.trim();
                !(t == "0" || t.eq_ignore_ascii_case("false"))
            }
        }
    }

    match tag {
        b"b" => props.bold = Some(parse_bool_val(val())),
        b"i" => props.italic = Some(parse_bool_val(val())),
        b"strike" => props.strike = Some(parse_bool_val(val())),
        b"u" => {
            // `<u/>` defaults to "single"; `<u val="double"/>` keeps the value.
            let v = val();
            props.underline = Some(v.unwrap_or_else(|| "single".to_string()));
        }
        b"sz" => {
            if let Some(v) = val() {
                if let Ok(n) = v.parse::<f64>() {
                    props.size = Some(n);
                }
            }
        }
        b"rFont" => props.name = val(),
        b"family" => {
            if let Some(v) = val() {
                if let Ok(n) = v.parse::<i32>() {
                    props.family = Some(n);
                }
            }
        }
        b"charset" => {
            if let Some(v) = val() {
                if let Ok(n) = v.parse::<i32>() {
                    props.charset = Some(n);
                }
            }
        }
        b"vertAlign" => props.vert_align = val(),
        b"scheme" => props.scheme = val(),
        b"color" => {
            // Color carries multiple possible attrs: rgb, theme, indexed, tint.
            // Keep the most-specific value; prefer rgb when present.
            let mut rgb: Option<String> = None;
            let mut theme: Option<String> = None;
            let mut indexed: Option<String> = None;
            for (k, v) in &all_attrs {
                match k.as_slice() {
                    b"rgb" => rgb = Some(v.clone()),
                    b"theme" => theme = Some(v.clone()),
                    b"indexed" => indexed = Some(v.clone()),
                    _ => {}
                }
            }
            if let Some(v) = rgb {
                props.color = Some(v);
            } else if let Some(v) = theme {
                // Encode as `theme:N` so callers can distinguish from rgb.
                props.color = Some(format!("theme:{v}"));
            } else if let Some(v) = indexed {
                props.color = Some(format!("indexed:{v}"));
            }
        }
        _ => {}
    }
}

/// Serialize a slice of [`RichTextRun`] back to OOXML run XML — i.e.
/// the *body* of an `<si>` or `<is>` element.  Caller wraps the
/// output in `<si>...</si>` (for SST writes) or `<is>...</is>` (for
/// inline-string cells).
pub fn emit_runs(runs: &[RichTextRun]) -> String {
    let mut out = String::new();
    for run in runs {
        out.push_str("<r>");
        if let Some(props) = &run.font {
            out.push_str("<rPr>");
            // Order mirrors openpyxl (which mirrors the schema): bold,
            // italic, strike, underline, size, color, font name,
            // family, charset, vertAlign, scheme.
            if matches!(props.bold, Some(true)) {
                out.push_str("<b/>");
            } else if matches!(props.bold, Some(false)) {
                out.push_str("<b val=\"0\"/>");
            }
            if matches!(props.italic, Some(true)) {
                out.push_str("<i/>");
            } else if matches!(props.italic, Some(false)) {
                out.push_str("<i val=\"0\"/>");
            }
            if matches!(props.strike, Some(true)) {
                out.push_str("<strike/>");
            } else if matches!(props.strike, Some(false)) {
                out.push_str("<strike val=\"0\"/>");
            }
            if let Some(u) = &props.underline {
                if u == "single" {
                    out.push_str("<u/>");
                } else {
                    out.push_str(&format!("<u val=\"{}\"/>", xml_attr_escape(u)));
                }
            }
            if let Some(sz) = props.size {
                out.push_str(&format!("<sz val=\"{}\"/>", format_num(sz)));
            }
            if let Some(c) = &props.color {
                if let Some(rest) = c.strip_prefix("theme:") {
                    out.push_str(&format!("<color theme=\"{}\"/>", xml_attr_escape(rest)));
                } else if let Some(rest) = c.strip_prefix("indexed:") {
                    out.push_str(&format!("<color indexed=\"{}\"/>", xml_attr_escape(rest)));
                } else {
                    out.push_str(&format!("<color rgb=\"{}\"/>", xml_attr_escape(c)));
                }
            }
            if let Some(name) = &props.name {
                out.push_str(&format!("<rFont val=\"{}\"/>", xml_attr_escape(name)));
            }
            if let Some(family) = props.family {
                out.push_str(&format!("<family val=\"{family}\"/>"));
            }
            if let Some(charset) = props.charset {
                out.push_str(&format!("<charset val=\"{charset}\"/>"));
            }
            if let Some(va) = &props.vert_align {
                out.push_str(&format!("<vertAlign val=\"{}\"/>", xml_attr_escape(va)));
            }
            if let Some(sc) = &props.scheme {
                out.push_str(&format!("<scheme val=\"{}\"/>", xml_attr_escape(sc)));
            }
            out.push_str("</rPr>");
        }
        // Preserve leading/trailing whitespace.
        let needs_preserve = run.text.chars().next().is_some_and(|c| c.is_whitespace())
            || run
                .text
                .chars()
                .next_back()
                .is_some_and(|c| c.is_whitespace());
        let escaped = xml_text_escape(&run.text);
        if needs_preserve {
            out.push_str(&format!("<t xml:space=\"preserve\">{escaped}</t>"));
        } else {
            out.push_str(&format!("<t>{escaped}</t>"));
        }
        out.push_str("</r>");
    }
    out
}

fn xml_attr_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn xml_text_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn format_num(n: f64) -> String {
    // Drop the trailing ".0" for whole numbers so that `<sz val="11"/>`
    // round-trips byte-identically (matches openpyxl's emitter).
    if n.fract() == 0.0 && n.is_finite() {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_single_plain_t_returns_none() {
        let xml = "<si><t>hello</t></si>";
        let runs = parse_runs_in_element(xml, b"si").unwrap();
        assert!(runs.is_none(), "single <t> should be plain, got {runs:?}");
    }

    #[test]
    fn parse_one_run_returns_some_one() {
        let xml = "<si><r><t>only</t></r></si>";
        let runs = parse_runs_in_element(xml, b"si").unwrap().unwrap();
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "only");
        assert!(runs[0].font.is_none());
    }

    #[test]
    fn parse_bold_run_extracts_font() {
        let xml = r#"<si><r><rPr><b/></rPr><t>bold</t></r></si>"#;
        let runs = parse_runs_in_element(xml, b"si").unwrap().unwrap();
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].text, "bold");
        let f = runs[0].font.as_ref().unwrap();
        assert_eq!(f.bold, Some(true));
        assert_eq!(f.italic, None);
    }

    #[test]
    fn parse_multiple_runs() {
        let xml = r#"<si><r><rPr><b/></rPr><t>bold</t></r><r><t> regular</t></r></si>"#;
        let runs = parse_runs_in_element(xml, b"si").unwrap().unwrap();
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "bold");
        assert_eq!(runs[0].font.as_ref().unwrap().bold, Some(true));
        assert_eq!(runs[1].text, " regular");
        assert!(runs[1].font.is_none());
    }

    #[test]
    fn parse_inline_string_wrapper() {
        let xml = r#"<is><r><rPr><i/></rPr><t>italic</t></r></is>"#;
        let runs = parse_runs_in_element(xml, b"is").unwrap().unwrap();
        assert_eq!(runs.len(), 1);
        assert_eq!(runs[0].font.as_ref().unwrap().italic, Some(true));
    }

    #[test]
    fn parse_color_size_name() {
        let xml = r#"<si><r><rPr><sz val="14"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t>x</t></r></si>"#;
        let runs = parse_runs_in_element(xml, b"si").unwrap().unwrap();
        let f = runs[0].font.as_ref().unwrap();
        assert_eq!(f.size, Some(14.0));
        assert_eq!(f.color.as_deref(), Some("FFFF0000"));
        assert_eq!(f.name.as_deref(), Some("Arial"));
    }

    #[test]
    fn parse_underline_default_single() {
        let xml = r#"<si><r><rPr><u/></rPr><t>x</t></r></si>"#;
        let runs = parse_runs_in_element(xml, b"si").unwrap().unwrap();
        assert_eq!(
            runs[0].font.as_ref().unwrap().underline.as_deref(),
            Some("single")
        );
    }

    #[test]
    fn parse_xml_entities_unescaped() {
        let xml = r#"<si><r><t>A &amp; B</t></r></si>"#;
        let runs = parse_runs_in_element(xml, b"si").unwrap().unwrap();
        assert_eq!(runs[0].text, "A & B");
    }

    #[test]
    fn parse_preserve_whitespace_run() {
        let xml = r#"<si><r><t xml:space="preserve">  hi  </t></r></si>"#;
        let runs = parse_runs_in_element(xml, b"si").unwrap().unwrap();
        assert_eq!(runs[0].text, "  hi  ");
    }

    #[test]
    fn emit_then_parse_round_trip_simple() {
        let original = vec![
            RichTextRun {
                text: "Bold".to_string(),
                font: Some(InlineFontProps {
                    bold: Some(true),
                    ..Default::default()
                }),
            },
            RichTextRun {
                text: " plain".to_string(),
                font: None,
            },
        ];
        let body = emit_runs(&original);
        let wrapped = format!("<si>{body}</si>");
        let runs = parse_runs_in_element(&wrapped, b"si").unwrap().unwrap();
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].text, "Bold");
        assert_eq!(runs[0].font.as_ref().unwrap().bold, Some(true));
        assert_eq!(runs[1].text, " plain");
        assert!(runs[1].font.is_none());
    }

    #[test]
    fn emit_full_font_properties() {
        let runs = vec![RichTextRun {
            text: "x".to_string(),
            font: Some(InlineFontProps {
                bold: Some(true),
                italic: Some(true),
                size: Some(11.0),
                color: Some("FF0000FF".to_string()),
                name: Some("Calibri".to_string()),
                ..Default::default()
            }),
        }];
        let body = emit_runs(&runs);
        assert!(body.contains("<b/>"), "{body}");
        assert!(body.contains("<i/>"), "{body}");
        assert!(body.contains("<sz val=\"11\"/>"), "{body}");
        assert!(body.contains("<color rgb=\"FF0000FF\"/>"), "{body}");
        assert!(body.contains("<rFont val=\"Calibri\"/>"), "{body}");
    }

    #[test]
    fn emit_xml_escape_in_text() {
        let runs = vec![RichTextRun {
            text: "A & B < C".to_string(),
            font: None,
        }];
        let body = emit_runs(&runs);
        assert!(body.contains("A &amp; B &lt; C"), "{body}");
    }

    #[test]
    fn emit_preserve_whitespace_attr_on_leading_space() {
        let runs = vec![RichTextRun {
            text: " leading".to_string(),
            font: None,
        }];
        let body = emit_runs(&runs);
        assert!(body.contains("xml:space=\"preserve\""), "{body}");
    }
}
