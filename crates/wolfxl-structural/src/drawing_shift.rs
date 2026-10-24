//! DrawingML anchor shifts for worksheet DrawingML parts.

use quick_xml::events::{BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use crate::axis::{Axis, ShiftPlan};
use crate::xml_text::TextNode;

/// Rewrite an `xl/drawings/drawingN.xml` part by shifting zero-based
/// `<xdr:row>` / `<xdr:col>` marker values under `<xdr:from>` and `<xdr:to>`.
pub fn shift_drawing_xml(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
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
    let mut in_marker = false;
    let mut in_axis_coord = false;
    // Text of the current `<row>` / `<col>` coordinate, collected across
    // text and entity-reference events and rewritten as one node.
    let mut coord = TextNode::default();

    loop {
        let event = reader.read_event_into(&mut buf);
        if !coord.is_empty() && !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_))) {
            let shifted = coord
                .unescaped()
                .ok()
                .and_then(|s| s.trim().parse::<i64>().ok())
                .map(|n| shift_zero_based_point(n, plan).to_string());
            let _ = match shifted {
                Some(text) => writer.write_event(Event::Text(BytesText::new(&text))),
                None => writer.write_event(Event::Text(BytesText::from_escaped(coord.raw()))),
            };
            coord.clear();
        }
        match event {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name().as_ref().as_bytes().to_vec();
                match local.as_slice() {
                    b"from" | b"to" => in_marker = true,
                    b"row" if in_marker && plan.axis == Axis::Row => in_axis_coord = true,
                    b"col" if in_marker && plan.axis == Axis::Col => in_axis_coord = true,
                    _ => {}
                }
                let _ = writer.write_event(Event::Start(e.to_owned()));
            }
            Ok(Event::Empty(ref e)) => {
                let _ = writer.write_event(Event::Empty(e.to_owned()));
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name().as_ref().as_bytes().to_vec();
                match local.as_slice() {
                    b"from" | b"to" => in_marker = false,
                    b"row" | b"col" => in_axis_coord = false,
                    _ => {}
                }
                let _ = writer.write_event(Event::End(e.to_owned()));
            }
            Ok(Event::Text(ref t)) if in_axis_coord => coord.push_text(t),
            Ok(Event::GeneralRef(ref r)) if in_axis_coord => coord.push_ref(r),
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

fn shift_zero_based_point(zero_based: i64, plan: &ShiftPlan) -> i64 {
    if plan.is_insert() {
        if zero_based + 1 >= plan.idx as i64 {
            zero_based + plan.n as i64
        } else {
            zero_based
        }
    } else {
        let delete_start = plan.idx as i64 - 1;
        let delete_end = delete_start + plan.abs_n() as i64;
        if zero_based >= delete_start && zero_based < delete_end {
            delete_start
        } else if zero_based >= delete_end {
            zero_based + plan.n as i64
        } else {
            zero_based
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Axis, ShiftPlan};

    #[test]
    fn shifts_prefixed_two_cell_anchor_rows() {
        let xml = br#"<xdr:wsDr><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>6</xdr:row></xdr:to></xdr:twoCellAnchor></xdr:wsDr>"#;
        let out = shift_drawing_xml(xml, &ShiftPlan::delete(Axis::Row, 1, 1));
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("<xdr:row>0</xdr:row>"), "{s}");
        assert!(s.contains("<xdr:row>5</xdr:row>"), "{s}");
    }

    #[test]
    fn shifts_arbitrary_prefixed_two_cell_anchor_cols() {
        let xml = br#"<d:wsDr><d:twoCellAnchor><d:from><d:col>1</d:col><d:row>1</d:row></d:from><d:to><d:col>6</d:col><d:row>1</d:row></d:to></d:twoCellAnchor></d:wsDr>"#;
        let out = shift_drawing_xml(xml, &ShiftPlan::delete(Axis::Col, 1, 1));
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("<d:col>0</d:col>"), "{s}");
        assert!(s.contains("<d:col>5</d:col>"), "{s}");
    }

    #[test]
    fn deleted_row_anchor_markers_snap_to_delete_boundary() {
        let xml = br#"<xdr:wsDr><xdr:twoCellAnchor><xdr:from><xdr:row>5</xdr:row></xdr:from><xdr:to><xdr:row>8</xdr:row></xdr:to></xdr:twoCellAnchor></xdr:wsDr>"#;
        let out = shift_drawing_xml(xml, &ShiftPlan::delete(Axis::Row, 5, 3));
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("<xdr:row>4</xdr:row>"), "{s}");
        assert!(s.contains("<xdr:row>5</xdr:row>"), "{s}");
    }

    #[test]
    fn deleted_col_anchor_markers_snap_to_delete_boundary() {
        let xml = br#"<xdr:wsDr><xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col></xdr:from><xdr:to><xdr:col>8</xdr:col></xdr:to></xdr:twoCellAnchor></xdr:wsDr>"#;
        let out = shift_drawing_xml(xml, &ShiftPlan::delete(Axis::Col, 5, 3));
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("<xdr:col>4</xdr:col>"), "{s}");
        assert!(s.contains("<xdr:col>5</xdr:col>"), "{s}");
    }
}
