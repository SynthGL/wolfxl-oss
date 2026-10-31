//! Workbook input shared with Python without copying: an immutable `bytes`
//! object, or a read-only `mmap.mmap` of the opened workbook file.
//!
//! Every native reader over one input holds a clone of the same owner, which
//! keeps the Python object alive and, for a mapping, holds a buffer export so
//! Python cannot close or resize the mapping while any reader remains. Readers
//! borrow the bytes only through their own cursors, so none shares a position.
//!
//! A mapping is not an immutable snapshot. It shows later writes to the same
//! file, and on POSIX systems reading a page past a later truncation of that
//! file raises a bus error that terminates the process. Consumed mapped pages
//! are released through `mmap.madvise(MADV_DONTNEED)` where Python provides
//! it; elsewhere (for example on Windows) read pages stay resident until the
//! operating system trims them.

use std::ops::Range;
use std::sync::Arc;

use pyo3::buffer::PyBuffer;
use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::pybacked::PyBackedBytes;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyBytes, PyType};
use wolfxl_reader::WorkbookBytes;

/// Cheaply cloneable owner of workbook bytes held by a Python object.
#[derive(Clone)]
pub(crate) struct PyWorkbookBytes(Arc<Owner>);

enum Owner {
    /// `bytes` storage is immutable and never moves while referenced.
    Bytes(PyBackedBytes),
    /// A read-only, C-contiguous export of an `mmap.mmap`, released on drop.
    Mapped {
        buffer: PyBuffer<u8>,
        advice: Option<PageRelease>,
    },
}

/// The mapping's own `madvise` method and the platform's `MADV_DONTNEED` and
/// page size, as published by Python's `mmap` module. Absent where Python
/// offers no `madvise` (for example on Windows).
struct PageRelease {
    madvise: Py<PyAny>,
    dontneed: i64,
    page_size: usize,
}

impl PyWorkbookBytes {
    /// Accept `bytes`, or an `mmap.mmap` whose buffer is read-only and
    /// C-contiguous. Writable mappings and other buffer objects are rejected,
    /// so no reader ever retains memory Python code can write through.
    pub(crate) fn new(source: &Bound<'_, PyAny>) -> PyResult<Self> {
        static MMAP_TYPE: PyOnceLock<Py<PyType>> = PyOnceLock::new();
        if let Ok(bytes) = source.cast::<PyBytes>() {
            return Ok(Self(Arc::new(Owner::Bytes(PyBackedBytes::from(
                bytes.clone(),
            )))));
        }
        if !source.is_instance(MMAP_TYPE.import(source.py(), "mmap", "mmap")?.as_any())? {
            return Err(PyTypeError::new_err(format!(
                "workbook source must be bytes or a read-only mmap.mmap, not {}",
                source.get_type().name()?
            )));
        }
        let buffer = PyBuffer::<u8>::get(source)?;
        if !buffer.readonly() {
            return Err(PyTypeError::new_err(
                "workbook mmap source must be read-only (mmap.ACCESS_READ)",
            ));
        }
        if !buffer.is_c_contiguous() {
            return Err(PyTypeError::new_err(
                "workbook mmap source must be C-contiguous",
            ));
        }
        let advice = PageRelease::lookup(source);
        Ok(Self(Arc::new(Owner::Mapped { buffer, advice })))
    }
}

impl<'a, 'py> FromPyObject<'a, 'py> for PyWorkbookBytes {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        Self::new(&obj)
    }
}

impl AsRef<[u8]> for PyWorkbookBytes {
    fn as_ref(&self) -> &[u8] {
        match &*self.0 {
            Owner::Bytes(bytes) => &**bytes,
            Owner::Mapped { buffer, .. } => {
                let len = buffer.len_bytes();
                if len == 0 {
                    return &[];
                }
                // SAFETY: the export is read-only and C-contiguous, so `len`
                // bytes start at `buf_ptr`. Holding the export keeps the
                // mapping alive at that address until `buffer` is dropped with
                // this owner, and no Python code can write through it.
                unsafe { std::slice::from_raw_parts(buffer.buf_ptr().cast::<u8>(), len) }
            }
        }
    }
}

impl std::fmt::Debug for PyWorkbookBytes {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let kind = match &*self.0 {
            Owner::Bytes(_) => "bytes",
            Owner::Mapped { .. } => "mmap",
        };
        f.debug_struct("PyWorkbookBytes")
            .field("kind", &kind)
            .field("len", &self.as_ref().len())
            .finish()
    }
}

impl WorkbookBytes for PyWorkbookBytes {
    /// Drop the resident pages of a consumed mapped range. The mapping is a
    /// shared read-only view of the file, so released pages are read again
    /// from the same file on the next access. `bytes` input is kept as is.
    fn release_consumed(&self, range: Range<usize>) {
        if let Owner::Mapped {
            advice: Some(advice),
            ..
        } = &*self.0
        {
            advice.release(range, self.as_ref().len());
        }
    }
}

impl PageRelease {
    fn lookup(source: &Bound<'_, PyAny>) -> Option<Self> {
        let module = source.py().import("mmap").ok()?;
        let madvise = source.getattr("madvise").ok()?;
        let dontneed = module
            .getattr("MADV_DONTNEED")
            .ok()?
            .extract::<i64>()
            .ok()?;
        let page_size = module.getattr("PAGESIZE").ok()?.extract::<usize>().ok()?;
        (page_size != 0).then(|| Self {
            madvise: madvise.unbind(),
            dontneed,
            page_size,
        })
    }

    fn release(&self, range: Range<usize>, len: usize) {
        // `madvise` needs a page-aligned start inside the mapping.
        let start = range.start - range.start % self.page_size;
        let end = range.end.min(len);
        if start >= end {
            return;
        }
        Python::attach(|py| {
            // Advisory only: on failure the pages simply stay resident.
            let _ = self.madvise.call1(py, (self.dontneed, start, end - start));
        });
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Read, Seek, SeekFrom, Write};

    use pyo3::types::PyDict;
    use wolfxl_reader::{WorkbookBytesCursor, RELEASE_WINDOW};

    use super::*;

    /// `mmap.mmap` over `file` with the given `mmap.ACCESS_*` mode.
    fn map<'py>(py: Python<'py>, file: &Bound<'py, PyAny>, access: &str) -> Bound<'py, PyAny> {
        let mmap = py.import("mmap").unwrap();
        let kwargs = PyDict::new(py);
        kwargs
            .set_item("access", mmap.getattr(access).unwrap())
            .unwrap();
        let fileno = file.call_method0("fileno").unwrap();
        mmap.getattr("mmap")
            .unwrap()
            .call((fileno, 0), Some(&kwargs))
            .unwrap()
    }

    #[test]
    fn bytes_are_shared_without_a_copy() {
        Python::initialize();
        Python::attach(|py| {
            let bytes = PyBytes::new(py, b"workbook package");
            let source = PyWorkbookBytes::new(bytes.as_any()).unwrap();
            assert_eq!(source.as_ref().as_ptr(), bytes.as_bytes().as_ptr());
            // `bytes` input never releases anything.
            source.release_consumed(0..bytes.as_bytes().len());
            assert_eq!(source.as_ref(), b"workbook package");
        });
    }

    #[test]
    fn rejects_writable_and_unmapped_buffers() {
        Python::initialize();
        let mut file = tempfile::NamedTempFile::new().unwrap();
        file.write_all(b"workbook package").unwrap();
        file.flush().unwrap();
        Python::attach(|py| {
            let builtins = py.import("builtins").unwrap();
            for value in [
                builtins
                    .getattr("bytearray")
                    .unwrap()
                    .call1((b"data".as_slice(),))
                    .unwrap(),
                builtins
                    .getattr("memoryview")
                    .unwrap()
                    .call1((PyBytes::new(py, b"data"),))
                    .unwrap(),
                py.None().into_bound(py),
            ] {
                let err = PyWorkbookBytes::new(&value).unwrap_err();
                assert!(err.is_instance_of::<PyTypeError>(py), "{err}");
            }
            let handle = builtins
                .getattr("open")
                .unwrap()
                .call1((file.path(), "rb"))
                .unwrap();
            let copied = map(py, &handle, "ACCESS_COPY");
            let err = PyWorkbookBytes::new(&copied).unwrap_err();
            assert!(err.to_string().contains("read-only"), "{err}");
            copied.call_method0("close").unwrap();
            handle.call_method0("close").unwrap();
        });
    }

    #[test]
    fn mapped_source_holds_its_export_and_stays_readable_after_release() {
        Python::initialize();
        let data: Vec<u8> = (0..RELEASE_WINDOW + 12_345)
            .map(|i| (i % 251) as u8)
            .collect();
        let mut file = tempfile::NamedTempFile::new().unwrap();
        file.write_all(&data).unwrap();
        file.flush().unwrap();
        Python::attach(|py| {
            let handle = py
                .import("builtins")
                .unwrap()
                .getattr("open")
                .unwrap()
                .call1((file.path(), "rb"))
                .unwrap();
            let mapped = map(py, &handle, "ACCESS_READ");
            let source = PyWorkbookBytes::new(&mapped).unwrap();
            assert_eq!(source.as_ref(), data.as_slice());
            #[cfg(unix)]
            assert!(matches!(
                &*source.0,
                Owner::Mapped {
                    advice: Some(_),
                    ..
                }
            ));

            // Reading past a full window releases consumed pages; independent
            // cursors still read the same bytes afterwards.
            let mut first = WorkbookBytesCursor::new(&source);
            let mut all = Vec::new();
            first.read_to_end(&mut all).unwrap();
            assert_eq!(all, data);
            let mut second = WorkbookBytesCursor::new(&source);
            second.seek(SeekFrom::Start(4097)).unwrap();
            let mut tail = [0u8; 16];
            second.read_exact(&mut tail).unwrap();
            assert_eq!(tail, data[4097..4113]);
            source.release_consumed(1..data.len() + 10);
            source.release_consumed(data.len()..data.len() + 1);
            assert_eq!(source.as_ref(), data.as_slice());
            drop((first, second));

            // The held export keeps Python from closing the mapping.
            let err = mapped.call_method0("close").unwrap_err();
            assert!(
                err.is_instance_of::<pyo3::exceptions::PyBufferError>(py),
                "{err}"
            );
            drop(source);
            mapped.call_method0("close").unwrap();
            handle.call_method0("close").unwrap();
        });
    }
}
