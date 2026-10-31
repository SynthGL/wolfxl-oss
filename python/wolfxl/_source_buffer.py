"""Read-only workbook source buffers.

An XLSX load keeps one source buffer that admission, the native reader,
streaming and package metadata reads all share. Path loads on POSIX map the
file descriptor opened by the load read-only instead of copying the package
into memory, so pages load on demand and can be discarded after use.

The map stays bound to the file opened at load time: replacing or renaming the
path afterwards, including an atomic rename over it, does not change what the
workbook reads. The map does not make the content immutable. Writing to or
truncating that same file in place while the workbook is open changes the
mapped bytes, and reading pages past a truncation can terminate the process
with ``SIGBUS``. Windows does not allow replacing, renaming or deleting a mapped
file, so Windows path loads keep an in-memory copy of the package instead.
"""

from __future__ import annotations

import io
import mmap
import os
import stat
import zipfile
from typing import Any, BinaryIO

from wolfxl._limits import LimitExceeded

#: Whether path loads map the source file rather than copying it into memory.
#: Windows keeps the copy so callers can still replace or delete the file.
MAP_PATH_SOURCES = os.name != "nt"

#: Immutable buffer types handed to native readers and package metadata reads.
SOURCE_BUFFERS = (bytes, mmap.mmap)

# Whole-source copies advance in page-aligned chunks and discard each chunk's
# mapped pages afterwards, so a copy never leaves the package resident.
_COPY_CHUNK = 8 * 1024 * 1024
_DISCARD = getattr(mmap, "MADV_DONTNEED", None)


class MappedSource(mmap.mmap):
    """Read-only map of one opened workbook file.

    ``file_id`` is the ``(st_dev, st_ino)`` pair of the mapped file, so saves can
    avoid rewriting that file in place while the workbook still maps it.
    """

    file_id: tuple[int, int]


def snapshot_file(source: BinaryIO, limit: int | None = None) -> bytes | MappedSource:
    """Return the source buffer for an opened workbook file.

    Regular non-empty files are mapped read-only when :data:`MAP_PATH_SOURCES`
    is set. Other files, and every file on Windows, are read into ``bytes``.
    ``limit`` bounds the accepted file size in bytes.
    """
    status = os.fstat(source.fileno())
    size = status.st_size
    if limit is not None and size > limit:
        raise _package_limit_error(size, limit)
    if MAP_PATH_SOURCES and size and stat.S_ISREG(status.st_mode):
        try:
            mapped = _map_file(source.fileno(), size)
        except (OverflowError, ValueError) as error:
            # The file shrank after fstat or exceeds the address space.
            raise OSError(str(error)) from error
        mapped.file_id = (status.st_dev, status.st_ino)
        return mapped
    cap = size + 1 if limit is None else min(size + 1, limit + 1)
    data = source.read(cap)
    if limit is not None and len(data) > limit:
        raise _package_limit_error(len(data), limit)
    return data


def _map_file(fd: int, size: int) -> MappedSource:
    try:
        # The map does not need its own descriptor once created.
        return MappedSource(fd, size, access=mmap.ACCESS_READ, trackfd=False)
    except TypeError:
        # ``trackfd`` requires Python 3.13.
        return MappedSource(fd, size, access=mmap.ACCESS_READ)


def _package_limit_error(observed: int, limit: int) -> LimitExceeded:
    return LimitExceeded(
        f"max_package_bytes exceeded: observed {observed}, configured "
        f"{limit}; adjust load_workbook(..., limits=Limits(max_package_bytes=...))"
    )


def as_source(data: Any) -> bytes | mmap.mmap:
    """Return ``data`` as an immutable source buffer.

    ``bytes`` and read-only maps are used as they are. Mutable buffers, which
    must never be retained, are copied once into ``bytes``.
    """
    if type(data) is bytes:
        return data
    if isinstance(data, mmap.mmap):
        with memoryview(data) as view:
            if view.readonly:
                return data
    return bytes(data)


def maps_file(source: Any, path: str) -> bool:
    """Return whether ``source`` maps the file currently at ``path``."""
    file_id = getattr(source, "file_id", None) if isinstance(source, MappedSource) else None
    if file_id is None:
        return False
    try:
        status = os.stat(path)
    except OSError:
        return False
    return (status.st_dev, status.st_ino) == file_id


class SourceCursor(io.RawIOBase):
    """Seekable binary reader over a source buffer.

    Each cursor keeps its own position, so concurrent ZIP readers over one map
    never share a file position, and reads copy only the requested range.
    """

    def __init__(self, source: bytes | mmap.mmap) -> None:
        super().__init__()
        self._view = memoryview(source)
        self._pos = 0

    def readable(self) -> bool:
        return True

    def seekable(self) -> bool:
        return True

    def tell(self) -> int:
        self._check_open()
        return self._pos

    def seek(self, offset: int, whence: int = io.SEEK_SET) -> int:
        self._check_open()
        if whence == io.SEEK_SET:
            position = offset
        elif whence == io.SEEK_CUR:
            position = self._pos + offset
        elif whence == io.SEEK_END:
            position = len(self._view) + offset
        else:
            raise ValueError(f"invalid whence ({whence})")
        if position < 0:
            raise ValueError(f"negative seek position {position}")
        self._pos = position
        return position

    def read(self, size: int | None = -1) -> bytes:
        self._check_open()
        start = min(self._pos, len(self._view))
        end = len(self._view) if size is None or size < 0 else min(start + size, len(self._view))
        self._pos = max(self._pos, end)
        return self._view[start:end].tobytes()

    def readall(self) -> bytes:
        return self.read()

    def readinto(self, buffer: Any) -> int:
        self._check_open()
        with memoryview(buffer) as target, target.cast("B") as out:
            start = min(self._pos, len(self._view))
            end = min(start + len(out), len(self._view))
            count = end - start
            out[:count] = self._view[start:end]
        self._pos = max(self._pos, end)
        return count

    def close(self) -> None:
        if not self.closed:
            self._view.release()
        super().close()

    def _check_open(self) -> None:
        if self.closed:
            raise ValueError("I/O operation on closed source cursor")


def source_file(source: bytes | mmap.mmap) -> SourceCursor:
    """Return an independent seekable file object over a source buffer."""
    return SourceCursor(source)


def open_zip(source: str | bytes | mmap.mmap) -> zipfile.ZipFile:
    """Open a source path or buffer as a read-only ZIP archive."""
    if isinstance(source, SOURCE_BUFFERS):
        return zipfile.ZipFile(SourceCursor(source))
    return zipfile.ZipFile(source)


def write_source(target: BinaryIO, source: bytes | mmap.mmap) -> None:
    """Write a whole source buffer to ``target`` without an in-memory copy."""
    if not isinstance(source, mmap.mmap):
        target.write(source)
        return
    with memoryview(source) as view:
        for start in range(0, len(view), _COPY_CHUNK):
            end = min(start + _COPY_CHUNK, len(view))
            target.write(view[start:end])
            _discard(source, start, end - start)


def _discard(source: mmap.mmap, start: int, length: int) -> None:
    madvise = getattr(source, "madvise", None)
    if madvise is None or _DISCARD is None:
        return
    try:
        madvise(_DISCARD, start, length)
    except (OSError, ValueError):
        # Advice is optional; the copy is already complete.
        pass


def release(source: Any) -> None:
    """Unmap ``source`` unless a live reader still uses it.

    Native readers and cursors keep the map alive while they exist, so the map
    is then released when the last of them is collected instead.
    """
    if not isinstance(source, mmap.mmap):
        return
    try:
        source.close()
    except BufferError:
        pass
