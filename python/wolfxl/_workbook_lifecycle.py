"""Workbook lifecycle helpers."""

from __future__ import annotations

from typing import Any, TypeVar

WorkbookT = TypeVar("WorkbookT")


def close_workbook(workbook: Any) -> None:
    """Release native handles and delete any temporary decrypted input."""
    archive = getattr(workbook, "_archive", None)
    if archive is not None:
        try:
            archive.close()
        except Exception:
            pass
    workbook._merged_border_metadata_cache = None
    _clear_worksheet_style_caches(workbook)
    workbook._rust_reader = None
    workbook._streaming_style_cache = None
    workbook._stream_shared_strings = None
    workbook._rust_writer = None
    workbook._rust_patcher = None
    _release_mapped_source(workbook)
    tmp_path = getattr(workbook, "_tempfile_path", None)
    if tmp_path is not None:
        finalizer = getattr(workbook, "_tempfile_finalizer", None)
        if finalizer is not None:
            finalizer()
        else:
            _remove_temporary_source(tmp_path)
        workbook._tempfile_path = None


def _release_mapped_source(workbook: Any) -> None:
    """Drop a mapped source once the workbook's readers are gone.

    Readers or row iterators that outlive ``close()`` keep the map alive until
    they are collected; byte sources stay in place as before.
    """
    import mmap

    from wolfxl._source_buffer import release

    source = getattr(workbook, "_source_bytes", None)
    if isinstance(source, mmap.mmap):
        workbook._source_bytes = None
        release(source)


def _remove_temporary_source(path: str) -> None:
    import os

    try:
        os.unlink(path)
    except OSError:
        pass


def _clear_worksheet_style_caches(workbook: Any) -> None:
    """Release read-side style windows before dropping the workbook reader."""
    # A style setter can replace the cache with a disabled sentinel before
    # close; weak registration still reaches an externally retained window.
    for window in getattr(workbook, "_style_payload_windows", ()) or ():
        window.close()
    workbook._style_payload_windows = None
    for worksheet in getattr(workbook, "_sheets", {}).values():
        cache = getattr(worksheet, "_style_payload_cache", None)
        close = getattr(cache, "close", None)
        if close is not None:
            close()
        if hasattr(worksheet, "_style_payload_cache"):
            worksheet._style_payload_cache = None
        if hasattr(worksheet, "_style_component_cache"):
            worksheet._style_component_cache = None


def enter_workbook(workbook: WorkbookT) -> WorkbookT:
    """Return this workbook for ``with`` statement use."""
    return workbook


def exit_workbook(workbook: Any, *args: object) -> None:
    """Close this workbook at the end of a ``with`` block."""
    workbook.close()


def repr_workbook(workbook: Any) -> str:
    """Return a compact debug representation for this workbook."""
    if workbook._rust_patcher is not None:
        mode = "modify"
    elif workbook._rust_reader is not None:
        mode = "read"
    else:
        mode = "write"
    return f"<Workbook [{mode}] sheets={workbook._sheet_names}>"
