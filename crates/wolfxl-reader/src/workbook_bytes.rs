//! Workbook bytes kept alive by an external owner, and a seekable reader over
//! them that reports the input it consumed so file-mapped owners can release
//! resident pages instead of keeping the whole package resident.

use std::io::{self, Read, Seek, SeekFrom};
use std::ops::Range;

/// Read input reported to [`WorkbookBytes::release_consumed`] at a time, so
/// sequential reads of a large package cost one release call per window.
pub const RELEASE_WINDOW: usize = 8 * 1024 * 1024;

/// Read-only workbook storage whose owner keeps its address and length stable
/// for every reader, such as Python `bytes` or a read-only file mapping.
///
/// A mapping prevents writes through this view, not through other file handles.
/// Callers must prevent writes or truncation of the mapped file while readers
/// remain alive.
pub trait WorkbookBytes: AsRef<[u8]> + std::fmt::Debug + Send + Sync {
    /// Reports that `range` of [`AsRef::as_ref`] has been read. File-mapped
    /// owners may release the resident pages covering it; the bytes must stay
    /// readable for later and concurrent readers. The default keeps them.
    fn release_consumed(&self, _range: Range<usize>) {}
}

impl WorkbookBytes for Vec<u8> {}

/// Span read through one reader since its last release report.
///
/// The span grows to cover every range read, and is reported once at least
/// [`RELEASE_WINDOW`] bytes were read and when the reader finishes. Releasing
/// is advisory: a range read again later is simply read from its owner again.
#[derive(Debug, Default)]
pub struct ConsumedRanges {
    span: Option<Range<usize>>,
    read: usize,
}

impl ConsumedRanges {
    /// Record that `range` was read from `owner`.
    pub fn record<B: WorkbookBytes + ?Sized>(&mut self, owner: &B, range: Range<usize>) {
        if range.is_empty() {
            return;
        }
        self.read = self.read.saturating_add(range.len());
        let span = self.span.get_or_insert(range.clone());
        span.start = span.start.min(range.start);
        span.end = span.end.max(range.end);
        if self.read >= RELEASE_WINDOW {
            self.flush(owner);
        }
    }

    /// Report the span read since the last report, if any.
    pub fn flush<B: WorkbookBytes + ?Sized>(&mut self, owner: &B) {
        self.read = 0;
        if let Some(span) = self.span.take() {
            owner.release_consumed(span);
        }
    }
}

/// Independent `Read + Seek` position over borrowed workbook bytes.
///
/// Each reader keeps its own position, so readers over the same owner never
/// share one; consumed input is reported to the owner as it is read and when
/// the reader is dropped.
#[derive(Debug)]
pub struct WorkbookBytesCursor<'a, B: WorkbookBytes + ?Sized> {
    owner: &'a B,
    pos: u64,
    consumed: ConsumedRanges,
}

impl<'a, B: WorkbookBytes + ?Sized> WorkbookBytesCursor<'a, B> {
    /// Reader positioned at the start of `owner`'s bytes.
    pub fn new(owner: &'a B) -> Self {
        Self {
            owner,
            pos: 0,
            consumed: ConsumedRanges::default(),
        }
    }

    /// The owner this reader borrows.
    pub fn get_ref(&self) -> &'a B {
        self.owner
    }
}

impl<B: WorkbookBytes + ?Sized> Read for WorkbookBytesCursor<'_, B> {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        let data = self.owner.as_ref();
        // Like `std::io::Cursor`, reads at or past the end return 0 bytes.
        let start = usize::try_from(self.pos).map_or(data.len(), |pos| pos.min(data.len()));
        let n = buf.len().min(data.len() - start);
        buf[..n].copy_from_slice(&data[start..start + n]);
        self.consumed.record(self.owner, start..start + n);
        self.pos += n as u64;
        Ok(n)
    }
}

impl<B: WorkbookBytes + ?Sized> Seek for WorkbookBytesCursor<'_, B> {
    fn seek(&mut self, pos: SeekFrom) -> io::Result<u64> {
        let (base, offset) = match pos {
            SeekFrom::Start(pos) => {
                self.pos = pos;
                return Ok(pos);
            }
            SeekFrom::End(offset) => (self.owner.as_ref().len() as u64, offset),
            SeekFrom::Current(offset) => (self.pos, offset),
        };
        match base.checked_add_signed(offset) {
            Some(pos) => {
                self.pos = pos;
                Ok(pos)
            }
            None => Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "invalid seek to a negative or overflowing position",
            )),
        }
    }
}

impl<B: WorkbookBytes + ?Sized> Drop for WorkbookBytesCursor<'_, B> {
    fn drop(&mut self) {
        self.consumed.flush(self.owner);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;

    #[derive(Debug, Default)]
    struct Recorded {
        bytes: Vec<u8>,
        released: Mutex<Vec<Range<usize>>>,
    }

    impl AsRef<[u8]> for Recorded {
        fn as_ref(&self) -> &[u8] {
            &self.bytes
        }
    }

    impl WorkbookBytes for Recorded {
        fn release_consumed(&self, range: Range<usize>) {
            self.released.lock().unwrap().push(range);
        }
    }

    #[test]
    fn cursor_matches_std_cursor_and_reports_consumed_span_on_drop() {
        let owner = Recorded {
            bytes: (0..=255).collect(),
            ..Recorded::default()
        };
        {
            let mut cursor = WorkbookBytesCursor::new(&owner);
            let mut buf = [0u8; 4];
            assert_eq!(cursor.seek(SeekFrom::End(-4)).unwrap(), 252);
            cursor.read_exact(&mut buf).unwrap();
            assert_eq!(buf, [252, 253, 254, 255]);
            assert_eq!(cursor.read(&mut buf).unwrap(), 0);
            assert_eq!(cursor.seek(SeekFrom::Start(1000)).unwrap(), 1000);
            assert_eq!(cursor.read(&mut buf).unwrap(), 0);
            cursor.seek(SeekFrom::Start(10)).unwrap();
            assert_eq!(cursor.seek(SeekFrom::Current(-2)).unwrap(), 8);
            cursor.read_exact(&mut buf).unwrap();
            assert_eq!(buf, [8, 9, 10, 11]);
            assert!(cursor.seek(SeekFrom::Current(-13)).is_err());
            assert!(owner.released.lock().unwrap().is_empty());
        }
        assert_eq!(*owner.released.lock().unwrap(), [8..256]);
    }

    #[test]
    fn independent_cursors_keep_their_own_positions() {
        let owner = Recorded {
            bytes: (0..16).collect(),
            ..Recorded::default()
        };
        let mut first = WorkbookBytesCursor::new(&owner);
        let mut second = WorkbookBytesCursor::new(&owner);
        let mut buf = [0u8; 2];
        first.seek(SeekFrom::Start(6)).unwrap();
        second.read_exact(&mut buf).unwrap();
        assert_eq!(buf, [0, 1]);
        first.read_exact(&mut buf).unwrap();
        assert_eq!(buf, [6, 7]);
    }

    #[test]
    fn consumed_ranges_report_each_full_window() {
        let owner = Recorded::default();
        let mut consumed = ConsumedRanges::default();
        consumed.record(&owner, 0..RELEASE_WINDOW - 1);
        assert!(owner.released.lock().unwrap().is_empty());
        consumed.record(&owner, RELEASE_WINDOW - 1..RELEASE_WINDOW + 1);
        consumed.record(&owner, RELEASE_WINDOW + 1..RELEASE_WINDOW + 3);
        consumed.flush(&owner);
        consumed.flush(&owner);
        assert_eq!(
            *owner.released.lock().unwrap(),
            [
                0..RELEASE_WINDOW + 1,
                RELEASE_WINDOW + 1..RELEASE_WINDOW + 3
            ]
        );
    }
}
