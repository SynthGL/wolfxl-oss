"""Bounded ZIP/XML helpers for Python-only OOXML inspection paths."""

from __future__ import annotations

import os
import zipfile

from wolfxl._rust import load_limit_defaults

(
    DEFAULT_MAX_ZIP_TOTAL_BYTES,
    DEFAULT_MAX_ZIP_ENTRY_BYTES,
    _DEFAULT_MAX_CELLS,
    _DEFAULT_MAX_SHARED_STRING_BYTES,
    _DEFAULT_MAX_MATERIALIZED_STRING_BYTES,
    DEFAULT_MAX_ZIP_ENTRIES,
    DEFAULT_MAX_COMPRESSION_RATIO,
) = load_limit_defaults()


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, str(default)))
    except ValueError:
        return default


def _max_entries() -> int:
    return _env_int("WOLFXL_MAX_ZIP_ENTRIES", DEFAULT_MAX_ZIP_ENTRIES)


def _max_entry_bytes() -> int:
    return _env_int("WOLFXL_MAX_ZIP_ENTRY_BYTES", DEFAULT_MAX_ZIP_ENTRY_BYTES)


def _max_total_bytes() -> int:
    return _env_int("WOLFXL_MAX_ZIP_TOTAL_BYTES", DEFAULT_MAX_ZIP_TOTAL_BYTES)


def _max_ratio() -> int:
    return _env_int("WOLFXL_MAX_ZIP_COMPRESSION_RATIO", DEFAULT_MAX_COMPRESSION_RATIO)


def validate_zipfile(zf: zipfile.ZipFile) -> None:
    infos = zf.infolist()
    if len(infos) > _max_entries():
        raise ValueError(
            f"OOXML package has too many ZIP entries: {len(infos)} > {_max_entries()}"
        )
    total = 0
    seen: set[str] = set()
    for info in infos:
        _validate_part_name(info.filename)
        if info.filename in seen:
            raise ValueError(f"OOXML package contains duplicate ZIP entry: {info.filename}")
        seen.add(info.filename)
        _validate_info(info)
        total += info.file_size
        if total > _max_total_bytes():
            raise ValueError(
                f"OOXML package is too large: {total} > {_max_total_bytes()} uncompressed bytes"
            )


def read_entry(zf: zipfile.ZipFile, name: str) -> bytes:
    info = zf.getinfo(name)
    _validate_info(info)
    return zf.read(info)


def read_entry_optional(zf: zipfile.ZipFile, name: str) -> bytes | None:
    try:
        return read_entry(zf, name)
    except KeyError:
        return None


def _validate_part_name(name: str) -> None:
    invalid = (
        not name
        or name.startswith(("/", "\\"))
        or "\\" in name
        or any(part == ".." or ":" in part for part in name.split("/"))
    )
    if invalid:
        raise ValueError(f"unsafe OOXML package part path: {name}")


def _validate_info(info: zipfile.ZipInfo) -> None:
    if info.file_size > _max_entry_bytes():
        raise ValueError(
            f"OOXML package part {info.filename} is too large: "
            f"{info.file_size} > {_max_entry_bytes()} bytes"
        )
    if info.file_size > 0 and info.compress_size == 0:
        raise ValueError(f"OOXML package part {info.filename} has invalid compressed size")
    if info.compress_size > 0 and info.file_size > _max_ratio() * info.compress_size:
        raise ValueError(
            f"OOXML package part {info.filename} exceeds compression ratio limit"
        )
