// SPDX-License-Identifier: MIT
// Independently authored from ODF 1.3 Part 2, sections 2 and 3.
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use zip::ZipArchive;

use super::xml::{self, Limits, Token};
use crate::{ReaderError, Result};

pub(super) fn read<R: Read + Seek>(input: R, limits: Limits) -> Result<HashMap<String, Vec<u8>>> {
    let mut archive = ZipArchive::new(input)?;
    if archive.len() > crate::max_zip_entries() {
        return Err(xml::malformed("ZIP entry budget exceeded"));
    }
    let mut names = HashSet::new();
    let mut declared_total = 0u64;
    // Admit every member's declaration before inflating anything.
    for i in 0..archive.len() {
        let file = archive.by_index(i)?;
        let name = file.name()?;
        crate::validate_zip_part_name(&name)?;
        if !names.insert(name.to_ascii_lowercase()) {
            return Err(xml::malformed("duplicate ZIP member"));
        }
        crate::validate_zip_entry_metadata(&name, file.size(), file.compressed_size())?;
        declared_total = declared_total
            .checked_add(file.size())
            .ok_or_else(|| xml::malformed("ZIP total overflow"))?;
        if declared_total > crate::max_zip_total_bytes() {
            return Err(xml::malformed("ZIP total budget exceeded"));
        }
    }
    let mut parts = HashMap::new();
    let mut total = 0u64;
    let mut chunk = [0u8; 16 * 1024];
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name()?.into_owned();
        let declared = file.size();
        if name == "mimetype"
            && declared != b"application/vnd.oasis.opendocument.spreadsheet".len() as u64
        {
            return Err(xml::malformed("invalid ODS mimetype length"));
        }
        let is_xml = name.to_ascii_lowercase().ends_with(".xml");
        let retain = is_xml || name == "mimetype";
        if is_xml && declared > limits.bytes as u64 {
            return Err(xml::malformed("XML byte budget exceeded"));
        }
        let mut bytes = Vec::new();
        let mut actual = 0u64;
        loop {
            let count = file.read(&mut chunk)?;
            if count == 0 {
                break;
            }
            actual = actual
                .checked_add(count as u64)
                .ok_or_else(|| xml::malformed("ZIP member overflow"))?;
            total = total
                .checked_add(count as u64)
                .ok_or_else(|| xml::malformed("ZIP total overflow"))?;
            if actual > declared
                || actual > crate::max_zip_entry_bytes()
                || total > crate::max_zip_total_bytes()
            {
                return Err(xml::malformed(
                    "ZIP actual size exceeds declaration or budget",
                ));
            }
            if retain {
                bytes.extend_from_slice(&chunk[..count]);
            }
        }
        if actual != declared {
            return Err(xml::malformed("ZIP actual size differs from declaration"));
        }
        // The workbook parser admits retained XML once, while decoding it.
        // Admit unused XML here so ignored members cannot bypass the budget.
        if is_xml && name != "content.xml" && name != "styles.xml" {
            xml::walk(&bytes, limits, |token, _| {
                if let Token::Start(e) = token {
                    if e.name.is(xml::MANIFEST, "encryption-data") {
                        return Err(ReaderError::Unsupported(
                            "encrypted ODS packages are not supported".into(),
                        ));
                    }
                }
                Ok(())
            })?;
        }
        if name == "content.xml" || name == "styles.xml" || name == "mimetype" {
            parts.insert(name, bytes);
        }
    }
    if total != declared_total {
        return Err(xml::malformed("ZIP total size mismatch"));
    }
    if let Some(mime) = parts.get("mimetype") {
        if mime != b"application/vnd.oasis.opendocument.spreadsheet" {
            return Err(ReaderError::Unsupported(
                "package is not an ODS spreadsheet".into(),
            ));
        }
    }
    Ok(parts)
}
