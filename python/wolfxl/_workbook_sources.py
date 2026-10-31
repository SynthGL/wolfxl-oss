"""Workbook source-opening helpers."""

from __future__ import annotations

from pathlib import PurePosixPath
from typing import Any
from wolfxl._limits import LimitExceeded, Limits

from wolfxl._workbook_state import (
    build_xlsb_xls_wb,
    build_xlsx_wb,
    xlsb_xls_via_tempfile,
)


def open_workbook_source(
    cls: type,
    *,
    fmt: str,
    path: str | None,
    data: bytes | None,
    password: str | bytes | None,
    data_only: bool,
    keep_links: bool,
    keep_vba: bool,
    permissive: bool,
    modify: bool,
    read_only: bool,
    limits: Limits | None = None,
) -> Any:
    """Open a classified workbook source through the matching constructor.

    Args:
        cls: Workbook class to materialize.
        fmt: File format returned by ``_loader.classify_input``.
        path: Source path for path-backed inputs.
        data: Source bytes for in-memory inputs.
        password: Password for OOXML-encrypted XLSX inputs.
        data_only: Return cached formula values when available.
        permissive: Enable recoverable malformed-workbook fallbacks.
        modify: Open XLSX inputs in read-modify-write mode.
        read_only: Enable streaming XLSX row iteration.

    Returns:
        A workbook instance opened in the requested mode.
    """
    if fmt == "xlsx":
        limits = limits if limits is not None else Limits()
        # Own one immutable snapshot for admission, native loading and streaming.
        # Never reopen the caller's path after accepting its contents.
        if data is None:
            data = _snapshot_xlsx_path(path, limits)
            path = None
        if password is None:
            limits._validate(data if data is not None else path, permissive=permissive)
        if password is not None:
            return from_encrypted(
                cls,
                path=path,
                data=data,
                password=password,
                data_only=data_only,
                keep_links=keep_links,
                keep_vba=keep_vba,
                permissive=permissive,
                modify=modify,
                read_only=read_only,
                limits=limits,
            )
        return from_bytes(
            cls,
            data,
            data_only=data_only,
            keep_links=keep_links,
            keep_vba=keep_vba,
            permissive=permissive,
            modify=modify,
            read_only=read_only,
            limits=limits,
        )

    if fmt == "xlsb":
        return from_xlsb(
            cls,
            path=path,
            data=data,
            data_only=data_only,
            permissive=permissive,
        )

    if fmt == "xls":
        return from_xls(
            cls,
            path=path,
            data=data,
            data_only=data_only,
            permissive=permissive,
        )

    raise ValueError(f"unsupported workbook format: {fmt!r}")


def _snapshot_xlsx_path(path: str, limits: Limits) -> bytes:
    import os
    import sys

    try:
        with open(path, "rb") as source:
            size = os.fstat(source.fileno()).st_size
            if size > limits.max_package_bytes:
                raise LimitExceeded(
                    f"max_package_bytes exceeded: observed {size}, configured "
                    f"{limits.max_package_bytes}; adjust load_workbook(..., limits=Limits(max_package_bytes=...))"
                )
            data = source.read(min(size + 1, limits.max_package_bytes + 1, sys.maxsize))
    except OSError as error:
        raise OSError(f"native xlsx open failed: {error}") from error
    if len(data) > limits.max_package_bytes:
        raise LimitExceeded(
            f"max_package_bytes exceeded: observed {len(data)}, configured "
            f"{limits.max_package_bytes}; adjust load_workbook(..., limits=Limits(max_package_bytes=...))"
        )
    return data


def from_reader(
    cls: type,
    path: str,
    *,
    data_only: bool = False,
    keep_links: bool = True,
    keep_vba: bool = False,
    permissive: bool = False,
    read_only: bool = False,
) -> Any:
    """Open an existing .xlsx file in read mode."""
    from wolfxl import _rust

    original_path = path
    temp_path = None
    path = _normalize_nonstandard_workbook_part(path) or path
    if path != original_path:
        temp_path = path

    reader_cls = _xlsx_reader_class(
        _rust,
        modify=False,
        read_only=read_only,
        permissive=permissive,
    )
    wb = build_xlsx_wb(
        cls,
        rust_reader=reader_cls.open(path, permissive),
        rust_patcher=None,
        data_only=data_only,
        read_only=read_only,
        source_path=path,
        keep_links=keep_links,
        keep_vba=keep_vba,
    )
    if temp_path is not None:
        wb._tempfile_path = temp_path
    if read_only:
        _attach_read_only_archive(wb, path)
    return wb


def _normalize_nonstandard_workbook_part(path: str | bytes) -> str | None:
    """Return a temp XLSX path when the workbook part is not ``xl/workbook.xml``."""
    import os
    import io
    import tempfile
    import zipfile
    from xml.etree import ElementTree as ET

    tmp_path = None
    try:
        with zipfile.ZipFile(io.BytesIO(path) if isinstance(path, bytes) else path, "r") as src:
            names = set(src.namelist())
            if "xl/workbook.xml" in names:
                return None
            rels = ET.fromstring(src.read("_rels/.rels"))
            target = None
            for rel in rels:
                if rel.get("Type", "").endswith("/officeDocument"):
                    target = rel.get("Target")
                    break
            if target:
                target = target.lstrip("/")
            if not target or target not in names:
                return None
            tmp = tempfile.NamedTemporaryFile(
                prefix="wolfxl-normalized-", suffix=".xlsx", delete=False
            )
            tmp_path = tmp.name
            tmp.close()
            target_rels = _workbook_rels_for_part(target)
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst:
                for info in src.infolist():
                    member = info.filename
                    data = src.read(member)
                    if member == target:
                        member = "xl/workbook.xml"
                    elif member == target_rels:
                        member = "xl/_rels/workbook.xml.rels"
                    elif member == "_rels/.rels":
                        data = _rewrite_root_rels_target(data)
                    elif member == "[Content_Types].xml":
                        data = _rewrite_content_types_workbook_part(data, target)
                    info.filename = member
                    dst.writestr(info, data)
            return tmp_path
    except Exception:
        if tmp_path is not None:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
        return None


def _workbook_rels_for_part(target: str) -> str:
    part = PurePosixPath(target)
    return str(part.parent / "_rels" / f"{part.name}.rels")


def _rewrite_root_rels_target(data: bytes) -> bytes:
    from xml.etree import ElementTree as ET

    root = ET.fromstring(data)
    for rel in root:
        if rel.get("Type", "").endswith("/officeDocument"):
            rel.set("Target", "xl/workbook.xml")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _rewrite_content_types_workbook_part(data: bytes, old_target: str) -> bytes:
    from xml.etree import ElementTree as ET

    root = ET.fromstring(data)
    old_part = "/" + old_target.lstrip("/")
    for child in root:
        if child.get("PartName") == old_part:
            child.set("PartName", "/xl/workbook.xml")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _attach_read_only_archive(wb: Any, path: str) -> None:
    try:
        wb._archive = _ReadOnlyArchive(path)
    except Exception:
        wb._archive = None


class _ReadOnlyArchive:
    """ZipFile-shaped archive that does not keep the source path locked."""

    def __init__(self, path: str) -> None:
        self.filename = path
        self.mode = "r"
        self._closed = False

    def open(self, name: str, *args: Any, **kwargs: Any) -> Any:
        import zipfile

        self._check_open()
        archive = zipfile.ZipFile(self.filename, "r")
        try:
            member = archive.open(name, *args, **kwargs)
        except Exception:
            archive.close()
            raise
        return _ZipMemberHandle(member, archive)

    def read(self, name: str, *args: Any, **kwargs: Any) -> bytes:
        import zipfile

        self._check_open()
        with zipfile.ZipFile(self.filename, "r") as archive:
            return archive.read(name, *args, **kwargs)

    def namelist(self) -> list[str]:
        import zipfile

        self._check_open()
        with zipfile.ZipFile(self.filename, "r") as archive:
            return archive.namelist()

    def infolist(self) -> list[Any]:
        import zipfile

        self._check_open()
        with zipfile.ZipFile(self.filename, "r") as archive:
            return archive.infolist()

    def getinfo(self, name: str) -> Any:
        import zipfile

        self._check_open()
        with zipfile.ZipFile(self.filename, "r") as archive:
            return archive.getinfo(name)

    def close(self) -> None:
        self._closed = True

    @property
    def fp(self) -> object | None:
        return None if self._closed else self

    def __enter__(self) -> _ReadOnlyArchive:
        self._check_open()
        return self

    def __exit__(self, *args: object) -> None:
        self.close()

    def _check_open(self) -> None:
        if self._closed:
            raise ValueError("Attempt to use ZIP archive that was already closed")


class _ZipMemberHandle:
    """Close the owning ZipFile when a member stream is closed."""

    def __init__(self, member: Any, archive: Any) -> None:
        self._member = member
        self._archive = archive
        self._closed = False

    def close(self) -> None:
        if self._closed:
            return
        try:
            self._member.close()
        finally:
            self._archive.close()
            self._closed = True

    @property
    def closed(self) -> bool:
        return self._closed or bool(getattr(self._member, "closed", False))

    def __enter__(self) -> _ZipMemberHandle:
        return self

    def __exit__(self, *args: object) -> None:
        self.close()

    def __del__(self) -> None:
        self.close()

    def __iter__(self) -> Any:
        return iter(self._member)

    def __getattr__(self, name: str) -> Any:
        return getattr(self._member, name)


def from_encrypted(
    cls: type,
    path: str | None = None,
    *,
    data: bytes | bytearray | memoryview | None = None,
    password: str | bytes,
    data_only: bool = False,
    keep_links: bool = True,
    keep_vba: bool = False,
    permissive: bool = False,
    modify: bool = False,
    read_only: bool = False,
    limits: Limits | None = None,
) -> Any:
    """Open an OOXML-encrypted .xlsx via msoffcrypto-tool."""
    if (path is None) == (data is None):
        raise TypeError("_from_encrypted requires exactly one of path / data")

    limits = limits if limits is not None else Limits()
    if path is not None:
        data = _snapshot_xlsx_path(path, limits)
        path = None
    is_plain_xlsx = bytes(data).startswith(b"PK")

    if is_plain_xlsx:
        limits._validate(bytes(data), permissive=permissive)
        return from_bytes(
            cls,
            bytes(data),
            data_only=data_only,
            keep_links=keep_links,
            keep_vba=keep_vba,
            permissive=permissive,
            modify=modify,
            read_only=read_only,
            limits=limits,
        )

    try:
        import msoffcrypto  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ImportError(
            "password reads require msoffcrypto-tool; install with: pip install wolfxl[encrypted]"
        ) from exc

    import io

    pw_str = password.decode("utf-8") if isinstance(password, bytes) else password
    if path is not None:
        src_fp = open(path, "rb")  # noqa: SIM115 - closed in finally
    else:
        src_fp = io.BytesIO(bytes(data))  # type: ignore[arg-type]

    try:
        office = msoffcrypto.OfficeFile(src_fp)
        try:
            is_encrypted = office.is_encrypted()
        except Exception:
            is_encrypted = False

        if not is_encrypted:
            limits._validate(bytes(data), permissive=permissive)
            return from_bytes(
                cls,
                bytes(data),
                data_only=data_only,
                keep_links=keep_links,
                keep_vba=keep_vba,
                permissive=permissive,
                modify=modify,
                read_only=read_only,
                limits=limits,
            )

        try:
            office.load_key(password=pw_str)
        except Exception as exc:
            raise ValueError(f"failed to load decryption key: {exc}") from exc

        buf = io.BytesIO()
        try:
            office.decrypt(buf)
        except Exception as exc:
            raise ValueError(f"failed to decrypt workbook (wrong password?): {exc}") from exc
        decrypted_bytes = buf.getvalue()
        (limits if limits is not None else Limits())._validate(
            decrypted_bytes, permissive=permissive
        )
    finally:
        src_fp.close()

    return from_bytes(
        cls,
        decrypted_bytes,
        data_only=data_only,
        keep_links=keep_links,
        keep_vba=keep_vba,
        permissive=permissive,
        modify=modify,
        read_only=read_only,
        limits=limits,
    )


def from_bytes(
    cls: type,
    data: bytes | bytearray | memoryview,
    *,
    data_only: bool = False,
    keep_links: bool = True,
    keep_vba: bool = False,
    permissive: bool = False,
    modify: bool = False,
    read_only: bool = False,
    limits: Limits | None = None,
) -> Any:
    """Open an .xlsx blob from memory."""
    from wolfxl import _rust

    # Normalize mutable buffers once. NativeXlsxBook retains this exact immutable
    # bytes object, also used by streaming and package metadata reads.
    data_bytes = bytes(data)
    normalized_path = _normalize_nonstandard_workbook_part(data_bytes)
    if normalized_path is not None:
        import os

        try:
            with open(normalized_path, "rb") as normalized:
                data_bytes = normalized.read()
        finally:
            os.unlink(normalized_path)
        (limits if limits is not None else Limits())._validate(
            data_bytes, permissive=permissive
        )
    reader_cls = _xlsx_reader_class(
        _rust,
        modify=modify,
        read_only=read_only,
        permissive=permissive,
    )
    bytes_open = getattr(reader_cls, "open_from_bytes", None)
    needs_tempfile = modify or bytes_open is None

    if needs_tempfile:
        import os
        import tempfile

        with tempfile.NamedTemporaryFile(prefix="wolfxl-", suffix=".xlsx", delete=False) as tmp:
            tmp.write(data_bytes)
            tmp_path = tmp.name

        try:
            if modify:
                wb = from_patcher(
                    cls,
                    tmp_path,
                    data_only=data_only,
                    keep_links=keep_links,
                    keep_vba=keep_vba,
                    permissive=permissive,
                )
            else:
                wb = from_reader(
                    cls,
                    tmp_path,
                    data_only=data_only,
                    keep_links=keep_links,
                    keep_vba=keep_vba,
                    permissive=permissive,
                    read_only=read_only,
                )
        except Exception:
            try:
                os.unlink(tmp_path)
            finally:
                raise
        wb._tempfile_path = tmp_path
        import weakref
        from wolfxl._workbook_lifecycle import _remove_temporary_source

        wb._tempfile_finalizer = weakref.finalize(wb, _remove_temporary_source, tmp_path)
        wb._source_bytes = data_bytes
        return wb

    return build_xlsx_wb(
        cls,
        rust_reader=bytes_open(data_bytes, permissive),
        rust_patcher=None,
        data_only=data_only,
        read_only=read_only,
        source_path=None,
        source_bytes=data_bytes,
        keep_links=keep_links,
        keep_vba=keep_vba,
    )


def from_patcher(
    cls: type,
    path: str,
    *,
    data_only: bool = False,
    keep_links: bool = True,
    keep_vba: bool = False,
    permissive: bool = False,
) -> Any:
    """Open an existing .xlsx file in modify mode."""
    from wolfxl import _rust

    reader_cls = _xlsx_reader_class(
        _rust,
        modify=True,
        read_only=False,
        permissive=permissive,
    )
    return build_xlsx_wb(
        cls,
        rust_reader=reader_cls.open(path, permissive),
        rust_patcher=_rust.XlsxPatcher.open(path, permissive),
        data_only=data_only,
        read_only=False,
        source_path=path,
        keep_links=keep_links,
        keep_vba=keep_vba,
    )


def _xlsx_reader_class(
    rust_module: Any,
    *,
    modify: bool,
    read_only: bool,
    permissive: bool,
) -> Any:
    """Return the active XLSX Rust reader class.

    Plain eager reads, permissive recovery, streaming bootstrap, and
    modify-mode bootstrap reads use WolfXL's native reader.
    """
    if hasattr(rust_module, "NativeXlsxBook"):
        return rust_module.NativeXlsxBook
    raise RuntimeError("wolfxl Rust extension is missing NativeXlsxBook")


def from_xlsb(
    cls: type,
    *,
    path: str | None,
    data: bytes | None,
    data_only: bool = False,
    permissive: bool = False,
) -> Any:
    """Open an .xlsb workbook via the native BIFF12 reader."""
    from wolfxl import _rust

    rust_cls = getattr(_rust, "NativeXlsbBook", None)
    if rust_cls is None:
        raise NotImplementedError(
            ".xlsb reads require the NativeXlsbBook backend from the wolfxl Rust extension."
        )

    if data is not None:
        rust_book, tmp_path = _open_binary_bytes(
            rust_cls, data, suffix=".xlsb", permissive=permissive
        )
        wb = build_xlsb_xls_wb(
            cls,
            rust_book=rust_book,
            fmt="xlsb",
            data_only=data_only,
            source_path=None,
        )
        if tmp_path is not None:
            wb._tempfile_path = tmp_path
        return wb

    rust_book = _open_binary_path(rust_cls, path, permissive=permissive)
    return build_xlsb_xls_wb(
        cls,
        rust_book=rust_book,
        fmt="xlsb",
        data_only=data_only,
        source_path=path,
    )


def from_xls(
    cls: type,
    *,
    path: str | None,
    data: bytes | None,
    data_only: bool = False,
    permissive: bool = False,
) -> Any:
    """Open a legacy .xls workbook via the independent native Community reader."""
    from wolfxl import _rust

    rust_cls = getattr(_rust, "NativeXlsBook", None)
    if rust_cls is None:
        raise NotImplementedError(
            ".xls reads require the NativeXlsBook backend from the wolfxl Rust extension."
        )

    if data is not None:
        rust_book, tmp_path = _open_binary_bytes(
            rust_cls, data, suffix=".xls", permissive=permissive
        )
        wb = build_xlsb_xls_wb(
            cls,
            rust_book=rust_book,
            fmt="xls",
            data_only=data_only,
            source_path=None,
        )
        if tmp_path is not None:
            wb._tempfile_path = tmp_path
        return wb

    rust_book = _open_binary_path(rust_cls, path, permissive=permissive)
    return build_xlsb_xls_wb(
        cls,
        rust_book=rust_book,
        fmt="xls",
        data_only=data_only,
        source_path=path,
    )


def _open_binary_bytes(
    rust_cls: Any,
    data: bytes,
    *,
    suffix: str,
    permissive: bool,
) -> tuple[Any, str | None]:
    bytes_open = getattr(rust_cls, "open_from_bytes", None)
    if bytes_open is None:
        rust_book, tmp_path = xlsb_xls_via_tempfile(
            rust_cls, data, suffix=suffix, permissive=permissive
        )
        return rust_book, tmp_path
    try:
        return bytes_open(data, permissive), None
    except TypeError:
        return bytes_open(data), None


def _open_binary_path(rust_cls: Any, path: str | None, *, permissive: bool) -> Any:
    opener = getattr(rust_cls, "open", None)
    if opener is None:
        raise NotImplementedError(
            f"{rust_cls.__name__}.open is not yet exposed by the Rust extension."
        )
    try:
        return opener(path, permissive)
    except TypeError:
        return opener(path)
