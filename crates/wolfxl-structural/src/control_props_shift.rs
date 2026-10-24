//! Form-control property shifts for `xl/ctrlProps/ctrlPropN.xml`.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use crate::axis::ShiftPlan;
use crate::shift_anchors::shift_anchor;

fn push_attr<'a>(e: &mut BytesStart<'a>, key: &str, val: &str) {
    e.push_attribute((key, val));
}

/// Rewrite form-control range attributes such as `fmlaRange`.
pub fn shift_control_props_xml(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }
    let xml_str = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };
    let mut reader = XmlReader::from_str(xml_str);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(std::io::Cursor::new(Vec::new()));
    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let _ = writer.write_event(Event::Start(rewrite_attrs(e, plan)));
            }
            Ok(Event::Empty(ref e)) => {
                let _ = writer.write_event(Event::Empty(rewrite_attrs(e, plan)));
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                let _ = writer.write_event(other);
            }
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner().into_inner()
}

fn rewrite_attrs<'a>(e: &BytesStart<'a>, plan: &ShiftPlan) -> BytesStart<'a> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    for attr_res in e.attributes().with_checks(false) {
        let Ok(attr) = attr_res else { continue };
        let key = attr.key.as_ref();
        let val = match quick_xml::escape::unescape(&attr.value) {
            Ok(v) => v.into_owned(),
            Err(_) => continue,
        };
        if key == "fmlaRange" {
            let shifted = shift_anchor(&val, plan);
            if shifted != "#REF!" {
                push_attr(&mut new_e, key, &shifted);
            } else {
                push_attr(&mut new_e, key, &val);
            }
        } else {
            push_attr(&mut new_e, key, &val);
        }
    }
    new_e
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Axis, ShiftPlan};

    #[test]
    fn shifts_fmla_range_attr() {
        let xml = br#"<formControlPr fmlaRange="$A$2:$A$6" objectType="List"/>"#;
        let out = shift_control_props_xml(xml, &ShiftPlan::delete(Axis::Row, 1, 1));
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"fmlaRange="$A$1:$A$5""#), "{s}");
    }
}
