use std::collections::HashMap;
use std::fs::File;

use pyo3::exceptions::PyIOError;
use pyo3::prelude::*;
use zip::ZipArchive;

use crate::ooxml_util;

use super::{content_types, patcher_workbook, XlsxPatcher};

pub(super) fn apply_external_links_drop_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    if !patcher.drop_external_links {
        return Ok(());
    }

    let workbook_rels_path = "xl/_rels/workbook.xml.rels";
    let mut workbook_rels =
        patcher_workbook::current_or_empty_rels(patcher, zip, workbook_rels_path)?;

    let external_rids: Vec<wolfxl_rels::RelId> = workbook_rels
        .iter()
        .filter(|rel| rel.rel_type == wolfxl_rels::rt::EXTERNAL_LINK)
        .map(|rel| rel.id.clone())
        .collect();
    if external_rids.is_empty() {
        patcher.drop_external_links = false;
        return Ok(());
    }

    for rid in &external_rids {
        workbook_rels.remove(rid);
    }
    patcher
        .rels_patches
        .insert(workbook_rels_path.to_string(), workbook_rels);

    for i in 0..zip.len() {
        let name = zip
            .by_index(i)
            .map_err(|e| PyIOError::new_err(format!("zip entry read: {e}")))?
            .name()
            .map_err(|e| PyIOError::new_err(format!("zip entry name: {e}")))?
            .into_owned();
        if name.starts_with("xl/externalLinks/") {
            patcher.file_deletes.insert(name);
        }
    }
    patcher
        .file_adds
        .retain(|name, _| !name.starts_with("xl/externalLinks/"));
    file_patches.retain(|name, _| !name.starts_with("xl/externalLinks/"));
    patcher
        .rels_patches
        .retain(|name, _| !name.starts_with("xl/externalLinks/"));

    let wb_xml = match file_patches.get("xl/workbook.xml") {
        Some(bytes) => String::from_utf8_lossy(bytes).into_owned(),
        None => ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?,
    };
    let stripped = remove_external_references_block(&wb_xml);
    file_patches.insert("xl/workbook.xml".to_string(), stripped.into_bytes());

    patcher
        .queued_content_type_ops
        .entry("__external_links_drop__".to_string())
        .or_default()
        .push(content_types::ContentTypeOp::RemoveOverridePrefix(
            "/xl/externalLinks/".to_string(),
        ));
    patcher.drop_external_links = false;
    Ok(())
}

fn remove_external_references_block(workbook_xml: &str) -> String {
    let Some(start) = workbook_xml.find("<externalReferences") else {
        return workbook_xml.to_string();
    };
    let rest = &workbook_xml[start..];
    let Some(open_end_rel) = rest.find('>') else {
        return workbook_xml.to_string();
    };
    let open_end = start + open_end_rel + 1;
    let end = if rest[..open_end_rel + 1].trim_end().ends_with("/>") {
        open_end
    } else if let Some(close_rel) = workbook_xml[open_end..].find("</externalReferences>") {
        open_end + close_rel + "</externalReferences>".len()
    } else {
        return workbook_xml.to_string();
    };
    let mut out = String::with_capacity(workbook_xml.len());
    out.push_str(&workbook_xml[..start]);
    out.push_str(&workbook_xml[end..]);
    out
}

#[cfg(test)]
mod tests {
    use super::remove_external_references_block;

    #[test]
    fn removes_external_references_block() {
        let xml = r#"<workbook><sheets/><externalReferences><externalReference r:id="rId2"/></externalReferences><definedNames/></workbook>"#;
        let out = remove_external_references_block(xml);
        assert!(!out.contains("externalReferences"));
        assert!(out.contains("<sheets/>"));
        assert!(out.contains("<definedNames/>"));
    }

    #[test]
    fn removes_self_closing_external_references_block() {
        let xml = r#"<workbook><sheets/><externalReferences/></workbook>"#;
        assert_eq!(
            remove_external_references_block(xml),
            "<workbook><sheets/></workbook>"
        );
    }
}

