//! Byte-level `<row>` locator for the values-only streaming fast path.
//!
//! [`RowScanner`] walks worksheet XML at tag granularity outside rows and
//! yields each `<row>` element as one validated UTF-8 window. It commits only
//! constructs it fully understands: plain start/end/empty tags with balanced
//! names, comments, processing instructions, and whitespace or plain text.
//! Anything else (DOCTYPE, top-level CDATA, entity references, prefixed rows,
//! mismatched end tags, invalid UTF-8, read errors, end of input) stops the
//! scan. [`RowScanner::handoff`] then returns the unconsumed bytes, preceded
//! by start tags for the still-open elements, so quick-xml resumes at the same
//! point with the same nesting and reports the same events and errors it
//! would have without the scanner.

use std::io::{self, Cursor, Read};
use std::ops::Range;

use memchr::{memchr, memchr3, memmem};

/// Bytes inflated per refill.
const READ_BLOCK: usize = 128 * 1024;
/// A row window larger than this is left to the event reader instead of
/// being buffered whole.
const MAX_ROW_WINDOW: usize = 16 * 1024 * 1024;

/// Where the remaining sheet bytes come from.
enum ScanSource {
    Reader(Box<dyn Read>),
    /// A read failed; the event reader reports it once the buffered bytes
    /// before it are consumed.
    Failed(io::Error),
    Done,
}

/// One located `<row>` element. Ranges index [`RowScanner::text`].
pub(crate) struct RowWindow {
    /// Start-tag content between `<` and `>`, without a trailing `/`.
    pub tag: Range<usize>,
    /// Content between the start tag and `</row>`; empty for `<row/>`.
    pub body: Range<usize>,
    pub self_closing: bool,
    /// Offset just past the element.
    end: usize,
}

pub(crate) struct RowScanner {
    buf: Vec<u8>,
    /// Next unconsumed byte.
    pos: usize,
    /// `buf[..valid_upto]` is known to be UTF-8.
    valid_upto: usize,
    /// Bytes past `valid_upto` are not UTF-8; nothing beyond is committed.
    invalid: bool,
    source: ScanSource,
    /// Qualified names of the elements open at `pos`.
    open: Vec<Box<[u8]>>,
    read_block: usize,
}

impl RowScanner {
    /// Scan `initial` followed by whatever `source` yields.
    pub(crate) fn new(initial: Vec<u8>, source: Option<Box<dyn Read>>) -> Self {
        let mut scanner = Self {
            buf: initial,
            pos: 0,
            valid_upto: 0,
            invalid: false,
            source: source.map_or(ScanSource::Done, ScanSource::Reader),
            open: Vec::new(),
            read_block: READ_BLOCK,
        };
        scanner.validate();
        scanner
    }

    #[cfg(test)]
    pub(crate) fn with_read_block(mut self, read_block: usize) -> Self {
        self.read_block = read_block.max(1);
        self
    }

    /// The validated buffer, which every [`RowWindow`] range indexes.
    pub(crate) fn text(&self) -> &str {
        // SAFETY: `validate` checked `buf[..valid_upto]` is UTF-8, and the
        // buffer only changes through `fill` (append) and `compact` (drop a
        // prefix ending at a consumed tag boundary), both of which keep that
        // prefix valid. Ranges are cut at ASCII delimiters, so slices of this
        // view fall on character boundaries.
        unsafe { std::str::from_utf8_unchecked(&self.buf[..self.valid_upto]) }
    }

    /// Mark `window` consumed.
    pub(crate) fn consume(&mut self, window: &RowWindow) {
        self.pos = window.end;
    }

    /// Locate the next `<row>`, skipping the markup before it. `None` means
    /// the scanner cannot continue and the caller must hand off; read errors
    /// are deferred to the event reader the same way.
    pub(crate) fn next_row(&mut self) -> Option<RowWindow> {
        if self.pos == 0 && self.buf.is_empty() && !self.fill() {
            return None;
        }
        if self.pos == 0 && self.buf.starts_with(b"\xEF\xBB\xBF") {
            // Leave BOM handling to quick-xml.
            return None;
        }
        loop {
            self.compact();
            let lt = self.find_bounded(self.pos, |hay| memchr(b'<', hay))?;
            let text = &self.buf[self.pos..lt];
            if text.contains(&b'&')
                || (self.open.is_empty() && !text.iter().all(u8::is_ascii_whitespace))
            {
                return None;
            }
            self.pos = lt;
            if !self.ensure(lt + 2) {
                return None;
            }
            match self.buf[lt + 1] {
                b'?' => {
                    let end = self.find_seq(lt + 2, b"?>")?;
                    self.pos = end + 2;
                }
                b'!' => {
                    // A short comment can end before nine bytes are available.
                    self.ensure(lt + 9);
                    if !self.ensure(lt + 4) {
                        return None;
                    }
                    let rest = &self.buf[lt..];
                    let end = if rest.starts_with(b"<!--") {
                        self.find_seq(lt + 4, b"-->")? + 3
                    } else if rest.starts_with(b"<![CDATA[") && !self.open.is_empty() {
                        self.find_seq(lt + 9, b"]]>")? + 3
                    } else {
                        return None;
                    };
                    self.pos = end;
                }
                b'/' => {
                    let gt = self.find_bounded(lt + 2, |hay| memchr(b'>', hay))?;
                    let name = self.buf[lt + 2..gt].trim_ascii_end();
                    if self.open.last().map(AsRef::as_ref) != Some(name) {
                        return None;
                    }
                    self.open.pop();
                    self.pos = gt + 1;
                }
                _ => {
                    let gt = self.find_tag_end(lt + 1)?;
                    let self_closing = self.buf[gt - 1] == b'/';
                    let tag = lt + 1..gt - usize::from(self_closing);
                    let content = &self.buf[tag.clone()];
                    let name_len = content
                        .iter()
                        .position(u8::is_ascii_whitespace)
                        .unwrap_or(content.len());
                    let name = &content[..name_len];
                    if name.is_empty() {
                        return None;
                    }
                    if name == b"row" {
                        return self.row_window(tag, gt, self_closing);
                    }
                    if name.rsplit(|b| *b == b':').next() == Some(&b"row"[..]) {
                        return None;
                    }
                    if !self_closing {
                        self.open.push(name.into());
                    }
                    self.pos = gt + 1;
                }
            }
        }
    }

    /// Bytes the event reader should parse in place of the rest of the
    /// stream: start tags for the open elements, the unconsumed buffer,
    /// then the remaining source (or its pending read error).
    pub(crate) fn handoff(self) -> HandoffReader {
        let mut prefix = Vec::with_capacity(
            self.open.iter().map(|name| name.len() + 2).sum::<usize>() + self.buf.len() - self.pos,
        );
        for name in &self.open {
            prefix.push(b'<');
            prefix.extend_from_slice(name);
            prefix.push(b'>');
        }
        prefix.extend_from_slice(&self.buf[self.pos..]);
        let rest: Box<dyn Read> = match self.source {
            ScanSource::Reader(reader) => reader,
            ScanSource::Failed(err) => Box::new(FailedRead(Some(err))),
            ScanSource::Done => Box::new(io::empty()),
        };
        Cursor::new(prefix).chain(rest)
    }

    fn row_window(
        &mut self,
        tag: Range<usize>,
        gt: usize,
        self_closing: bool,
    ) -> Option<RowWindow> {
        let body_start = gt + 1;
        let (body, end) = if self_closing {
            (body_start..body_start, body_start)
        } else {
            let finder = memmem::Finder::new(b"</row>");
            let close = self.find_bounded(body_start, |hay| finder.find(hay))?;
            (body_start..close, close + b"</row>".len())
        };
        Some(RowWindow {
            tag,
            body,
            self_closing,
            end,
        })
    }

    /// End of the tag opened at `start - 1`: the first `>` outside quotes.
    fn find_tag_end(&mut self, start: usize) -> Option<usize> {
        let mut at = start;
        let mut quote: Option<u8> = None;
        loop {
            let found = match quote {
                Some(q) => memchr(q, &self.buf[at..self.valid_upto]),
                None => memchr3(b'>', b'"', b'\'', &self.buf[at..self.valid_upto]),
            };
            match found {
                Some(offset) => {
                    let i = at + offset;
                    match (quote, self.buf[i]) {
                        (Some(_), _) => quote = None,
                        (None, b'>') => return Some(i),
                        (None, q) => quote = Some(q),
                    }
                    at = i + 1;
                }
                None => {
                    at = self.valid_upto;
                    if at - self.pos > MAX_ROW_WINDOW || !self.fill() {
                        return None;
                    }
                }
            }
        }
    }

    fn find_seq(&mut self, start: usize, needle: &[u8]) -> Option<usize> {
        let finder = memmem::Finder::new(needle);
        self.find_bounded(start, |hay| finder.find(hay))
    }

    fn find_bounded(
        &mut self,
        start: usize,
        find: impl Fn(&[u8]) -> Option<usize>,
    ) -> Option<usize> {
        let mut start = start;
        loop {
            if let Some(offset) = find(&self.buf[start..self.valid_upto]) {
                return Some(start + offset);
            }
            if self.valid_upto - self.pos > MAX_ROW_WINDOW {
                return None;
            }
            // Re-search the tail in case the needle straddles the refill.
            let resume = self.valid_upto.saturating_sub(8).max(start);
            if !self.fill() {
                return None;
            }
            start = resume;
        }
    }

    /// Ensure `buf[..len]` is buffered and validated.
    fn ensure(&mut self, len: usize) -> bool {
        while self.valid_upto < len {
            if !self.fill() {
                return false;
            }
        }
        true
    }

    /// Drop the consumed prefix once it outweighs the live bytes. Runs only
    /// between scan steps, so no offset other than `pos` is live.
    fn compact(&mut self) {
        if self.pos >= self.read_block && self.pos * 2 >= self.buf.len() {
            self.buf.drain(..self.pos);
            self.valid_upto -= self.pos;
            self.pos = 0;
        }
    }

    /// Append another block; offsets stay valid. `false` once no more valid
    /// bytes can arrive.
    fn fill(&mut self) -> bool {
        if self.invalid {
            return false;
        }
        let ScanSource::Reader(reader) = &mut self.source else {
            return false;
        };
        let old = self.buf.len();
        self.buf.resize(old + self.read_block, 0);
        let read = loop {
            match reader.read(&mut self.buf[old..]) {
                Err(err) if err.kind() == io::ErrorKind::Interrupted => continue,
                other => break other,
            }
        };
        let n = match read {
            Ok(n) => n,
            Err(err) => {
                self.buf.truncate(old);
                self.source = ScanSource::Failed(err);
                return false;
            }
        };
        self.buf.truncate(old + n);
        if n == 0 {
            self.source = ScanSource::Done;
        }
        let before = self.valid_upto;
        self.validate();
        self.valid_upto > before || (n > 0 && !self.invalid)
    }

    fn validate(&mut self) {
        match std::str::from_utf8(&self.buf[self.valid_upto..]) {
            Ok(_) => self.valid_upto = self.buf.len(),
            Err(err) => {
                self.valid_upto += err.valid_up_to();
                // An incomplete trailing sequence may complete on refill.
                if err.error_len().is_some() || matches!(self.source, ScanSource::Done) {
                    self.invalid = true;
                }
            }
        }
    }
}

/// Bytes the event reader parses after a handoff.
pub(crate) type HandoffReader = io::Chain<Cursor<Vec<u8>>, Box<dyn Read>>;

/// Replays a read error deferred by the scanner.
struct FailedRead(Option<io::Error>);

impl Read for FailedRead {
    fn read(&mut self, _buf: &mut [u8]) -> io::Result<usize> {
        match self.0.take() {
            Some(err) => Err(err),
            None => Ok(0),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rows(xml: &[u8], block: usize) -> (Vec<(String, String)>, Vec<u8>) {
        let source: Box<dyn Read> = Box::new(Cursor::new(xml.to_vec()));
        let mut scanner = RowScanner::new(Vec::new(), Some(source)).with_read_block(block);
        let mut out = Vec::new();
        while let Some(window) = scanner.next_row() {
            let text = scanner.text();
            out.push((
                text[window.tag.clone()].to_owned(),
                text[window.body.clone()].to_owned(),
            ));
            scanner.consume(&window);
        }
        let mut rest = Vec::new();
        scanner.handoff().read_to_end(&mut rest).unwrap();
        (out, rest)
    }

    #[test]
    fn rows_survive_every_refill_boundary() {
        let xml = "<?xml version=\"1.0\"?>\n<worksheet a='x>y'><!-- <row r=\"9\"> -->\
                   <sheetData><row r=\"1\" spans=\"1:2\"><c r=\"A1\"><v>1</v></c></row>\
                   <row r=\"2\"/><row r=\"3\"><c r=\"A3\" t=\"s\"><v>0</v></c></row>\
                   </sheetData></worksheet>";
        for block in 1..=xml.len() {
            let (found, rest) = rows(xml.as_bytes(), block);
            assert_eq!(
                found,
                [
                    (
                        "row r=\"1\" spans=\"1:2\"".into(),
                        "<c r=\"A1\"><v>1</v></c>".into()
                    ),
                    ("row r=\"2\"".into(), String::new()),
                    (
                        "row r=\"3\"".into(),
                        "<c r=\"A3\" t=\"s\"><v>0</v></c>".into()
                    ),
                ],
                "block {block}"
            );
            // End of input hands off with the still-open context consumed.
            assert_eq!(rest, b"", "block {block}");
        }
    }

    #[test]
    fn unsupported_markup_hands_off_with_open_elements() {
        for (xml, rest) in [
            (
                "<worksheet><sheetData><x:row r=\"1\"/></sheetData></worksheet>",
                "<worksheet><sheetData><x:row r=\"1\"/></sheetData></worksheet>",
            ),
            (
                "<worksheet><sheetData><![CDATA[<row/>]]><row/></sheetData></worksheet>",
                "<worksheet><sheetData><row/></sheetData></worksheet>",
            ),
            ("<!DOCTYPE w><worksheet/>", "<!DOCTYPE w><worksheet/>"),
            ("<a><b></a>", "<a><b></a>"),
            ("<a>&amp;<row/></a>", "<a>&amp;<row/></a>"),
            ("\u{FEFF}<a><row/></a>", "\u{FEFF}<a><row/></a>"),
        ] {
            let (found, handed) = rows(xml.as_bytes(), 4);
            // CDATA inside an element is skipped; its row is still found.
            if xml.contains("CDATA") {
                assert_eq!(found.len(), 1, "{xml}");
                assert_eq!(handed, b"");
                continue;
            }
            assert!(found.is_empty(), "{xml}");
            assert_eq!(String::from_utf8(handed).unwrap(), rest, "{xml}");
        }
    }

    #[test]
    fn invalid_utf8_and_read_errors_reach_the_event_reader() {
        let mut xml = b"<a><row/><row r=\"\xFF\"/></a>".to_vec();
        let (found, rest) = rows(&xml, 3);
        assert_eq!(found.len(), 1);
        assert_eq!(rest, b"<a><row r=\"\xFF\"/></a>");

        struct Failing(Cursor<Vec<u8>>);
        impl Read for Failing {
            fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
                match self.0.read(buf)? {
                    0 => Err(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "Invalid checksum",
                    )),
                    n => Ok(n),
                }
            }
        }
        xml = b"<a><row/><row/>".to_vec();
        let source: Box<dyn Read> = Box::new(Failing(Cursor::new(xml)));
        let mut scanner = RowScanner::new(Vec::new(), Some(source)).with_read_block(5);
        let mut found = 0;
        while let Some(window) = scanner.next_row() {
            scanner.consume(&window);
            found += 1;
        }
        assert_eq!(found, 2);
        let mut handed = Vec::new();
        let err = scanner.handoff().read_to_end(&mut handed).unwrap_err();
        assert_eq!(err.to_string(), "Invalid checksum");
        assert_eq!(handed, b"<a>");
    }
}
