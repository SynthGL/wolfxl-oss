"""Workbook-level security classes — RFC-058.

Two openpyxl-shaped containers:

* :class:`WorkbookProtection` — toggles structure / window / revision
  locks plus optional SHA-512 spin-hashed passwords for the
  ``workbookPassword`` and ``revisionsPassword`` slots.
* :class:`FileSharing` — read-only-recommended flag plus optional
  user name and reservation password.

Both classes match openpyxl's attribute names (camelCase + snake_case
aliases) so existing code using ``wb.security`` / ``wb.fileSharing``
continues to work. Hash algorithm defaults to ``"SHA-512"`` with a
spin count of 100,000 (Excel 2013+ default).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._compat import _iter_openpyxl_attrs, _openpyxl_name_fallback
from wolfxl.utils.protection import (
    check_password as check_legacy_password,
    hash_password,
    hash_password_sha512,
    verify_password_sha512,
)
from wolfxl.xml.functions import Element

if TYPE_CHECKING:
    pass


# ---------------------------------------------------------------------------
# WorkbookProtection
# ---------------------------------------------------------------------------


def _protection_dicts_equal(left: dict[str, str], right: dict[str, str]) -> bool:
    if left == right:
        return True

    def equivalent_password(a: str | None, b: str | None) -> bool:
        if a == b:
            return True
        if a is None or b is None:
            return False
        return hash_password(a) == b or hash_password(b) == a

    password_keys = {"workbookPassword", "revisionsPassword"}
    for key in set(left) | set(right):
        if key in password_keys:
            if not equivalent_password(left.get(key), right.get(key)):
                return False
            continue
        if left.get(key) != right.get(key):
            return False
    return True


class WorkbookProtection:
    """Workbook-level protection block (``<workbookProtection>``).

    Two independent password slots:

    * **Workbook password** — protects the workbook structure (sheet
      add/remove/reorder, hidden-sheet visibility). Set via
      :meth:`set_workbook_password`.
    * **Revisions password** — protects revision tracking history.
      Set via :meth:`set_revisions_password`.

    Each password is stored as a SHA-512 spin hash plus a salt; the
    plaintext is never persisted. Passing ``workbook_password=`` /
    ``revisions_password=`` to the constructor is a convenience that
    routes through the corresponding ``set_*_password`` method.

    The three lock booleans (``lock_structure``, ``lock_windows``,
    ``lock_revision``) work independently of the password slots —
    a True flag with no password means "locked but anyone can unlock".
    """

    __slots__ = (
        "lock_structure",
        "lock_windows",
        "lock_revision",
        "workbook_password",
        "workbook_password_character_set",
        "workbook_algorithm_name",
        "workbook_hash_value",
        "workbook_salt_value",
        "workbook_spin_count",
        "revisions_password",
        "revisions_password_character_set",
        "revisions_algorithm_name",
        "revisions_hash_value",
        "revisions_salt_value",
        "revisions_spin_count",
    )

    __attrs__ = (
        "workbookPassword",
        "workbookPasswordCharacterSet",
        "revisionsPassword",
        "revisionsPasswordCharacterSet",
        "lockStructure",
        "lockWindows",
        "lockRevision",
        "revisionsAlgorithmName",
        "revisionsHashValue",
        "revisionsSaltValue",
        "revisionsSpinCount",
        "workbookAlgorithmName",
        "workbookHashValue",
        "workbookSaltValue",
        "workbookSpinCount",
    )

    tagname = "workbookPr"

    def __init__(
        self,
        *,
        workbook_password: str | None = None,
        revisions_password: str | None = None,
        lock_structure: bool | None = None,
        lock_windows: bool | None = None,
        lock_revision: bool | None = None,
        workbook_password_character_set: str | None = None,
        workbook_algorithm_name: str | None = None,
        workbook_hash_value: str | None = None,
        workbook_salt_value: str | None = None,
        workbook_spin_count: int | None = None,
        revisions_password_character_set: str | None = None,
        revisions_algorithm_name: str | None = None,
        revisions_hash_value: str | None = None,
        revisions_salt_value: str | None = None,
        revisions_spin_count: int | None = None,
        workbookPassword: str | None = None,  # noqa: N803
        revisionsPassword: str | None = None,  # noqa: N803
        lockStructure: bool | None = None,  # noqa: N803
        lockWindows: bool | None = None,  # noqa: N803
        lockRevision: bool | None = None,  # noqa: N803
        workbookPasswordCharacterSet: str | None = None,  # noqa: N803
        workbookAlgorithmName: str | None = None,  # noqa: N803
        workbookHashValue: str | None = None,  # noqa: N803
        workbookSaltValue: str | None = None,  # noqa: N803
        workbookSpinCount: int | None = None,  # noqa: N803
        revisionsPasswordCharacterSet: str | None = None,  # noqa: N803
        revisionsAlgorithmName: str | None = None,  # noqa: N803
        revisionsHashValue: str | None = None,  # noqa: N803
        revisionsSaltValue: str | None = None,  # noqa: N803
        revisionsSpinCount: int | None = None,  # noqa: N803
    ) -> None:
        if workbookPassword is not None:
            workbook_password = workbookPassword
        if revisionsPassword is not None:
            revisions_password = revisionsPassword
        if lockStructure is not None:
            lock_structure = lockStructure
        if lockWindows is not None:
            lock_windows = lockWindows
        if lockRevision is not None:
            lock_revision = lockRevision
        if workbookPasswordCharacterSet is not None:
            workbook_password_character_set = workbookPasswordCharacterSet
        if workbookAlgorithmName is not None:
            workbook_algorithm_name = workbookAlgorithmName
        if workbookHashValue is not None:
            workbook_hash_value = workbookHashValue
        if workbookSaltValue is not None:
            workbook_salt_value = workbookSaltValue
        if workbookSpinCount is not None:
            workbook_spin_count = workbookSpinCount
        if revisionsPasswordCharacterSet is not None:
            revisions_password_character_set = revisionsPasswordCharacterSet
        if revisionsAlgorithmName is not None:
            revisions_algorithm_name = revisionsAlgorithmName
        if revisionsHashValue is not None:
            revisions_hash_value = revisionsHashValue
        if revisionsSaltValue is not None:
            revisions_salt_value = revisionsSaltValue
        if revisionsSpinCount is not None:
            revisions_spin_count = revisionsSpinCount

        self.lock_structure: bool | None = lock_structure
        self.lock_windows: bool | None = lock_windows
        self.lock_revision: bool | None = lock_revision

        self.workbook_password: str | None = None
        self.workbook_password_character_set: str | None = workbook_password_character_set
        self.workbook_algorithm_name: str | None = workbook_algorithm_name
        self.workbook_hash_value: str | None = workbook_hash_value
        self.workbook_salt_value: str | None = workbook_salt_value
        self.workbook_spin_count: int | None = workbook_spin_count

        self.revisions_password: str | None = None
        self.revisions_password_character_set: str | None = revisions_password_character_set
        self.revisions_algorithm_name: str | None = revisions_algorithm_name
        self.revisions_hash_value: str | None = revisions_hash_value
        self.revisions_salt_value: str | None = revisions_salt_value
        self.revisions_spin_count: int | None = revisions_spin_count

        if workbook_password is not None:
            self.set_workbook_password(workbook_password)
        if revisions_password is not None:
            self.set_revisions_password(revisions_password)

    # ------------------------------------------------------------------
    # Workbook password
    # ------------------------------------------------------------------

    def set_workbook_password(
        self,
        plaintext: str | None = None,
        algorithm: str | None = None,
        *,
        salt: bytes | None = None,
        spin_count: int = 100_000,
        value: str | None = None,
        already_hashed: bool = False,
    ) -> None:
        """Hash ``plaintext`` and store it on the workbook-password slot.

        The openpyxl-compatible default stores the legacy 16-bit XOR
        hash on ``workbook_password`` / ``workbookPassword``. Passing
        modern hash options (``algorithm``, ``salt``, or a custom
        ``spin_count``) also populates the SHA hash attribute group.
        """
        plaintext = "" if plaintext is None and value is None else (
            value if value is not None else plaintext
        )
        if already_hashed:
            self.workbook_password = plaintext
            return

        self.workbook_password = hash_password(plaintext)
        use_modern_hash = algorithm is not None or salt is not None or spin_count != 100_000
        if not use_modern_hash:
            self.workbook_algorithm_name = None
            self.workbook_hash_value = None
            self.workbook_salt_value = None
            self.workbook_spin_count = None
            return

        algorithm = algorithm or "SHA-512"
        h, s = hash_password_sha512(
            plaintext,
            salt=salt,
            spin_count=spin_count,
            algorithm=algorithm,
        )
        self.workbook_algorithm_name = algorithm
        self.workbook_hash_value = h
        self.workbook_salt_value = s
        self.workbook_spin_count = spin_count

    def check_workbook_password(self, plaintext: str) -> bool:
        """Return ``True`` iff ``plaintext`` matches the stored hash.

        Prefer the modern SHA hash group when present, otherwise verify
        the openpyxl-compatible legacy ``workbookPassword`` hash.
        """
        if (
            self.workbook_hash_value is not None
            and self.workbook_salt_value is not None
            and self.workbook_algorithm_name is not None
            and self.workbook_spin_count is not None
        ):
            return verify_password_sha512(
                plaintext,
                self.workbook_hash_value,
                self.workbook_salt_value,
                spin_count=self.workbook_spin_count,
                algorithm=self.workbook_algorithm_name,
            )
        if self.workbook_password is None:
            return False
        return check_legacy_password(plaintext, self.workbook_password)

    # ------------------------------------------------------------------
    # Revisions password
    # ------------------------------------------------------------------

    def set_revisions_password(
        self,
        plaintext: str | None = None,
        algorithm: str | None = None,
        *,
        salt: bytes | None = None,
        spin_count: int = 100_000,
        value: str | None = None,
        already_hashed: bool = False,
    ) -> None:
        """Hash ``plaintext`` and store it on the revisions-password slot."""
        plaintext = "" if plaintext is None and value is None else (
            value if value is not None else plaintext
        )
        if already_hashed:
            self.revisions_password = plaintext
            return

        self.revisions_password = hash_password(plaintext)
        use_modern_hash = algorithm is not None or salt is not None or spin_count != 100_000
        if not use_modern_hash:
            self.revisions_algorithm_name = None
            self.revisions_hash_value = None
            self.revisions_salt_value = None
            self.revisions_spin_count = None
            return

        algorithm = algorithm or "SHA-512"
        h, s = hash_password_sha512(
            plaintext,
            salt=salt,
            spin_count=spin_count,
            algorithm=algorithm,
        )
        self.revisions_algorithm_name = algorithm
        self.revisions_hash_value = h
        self.revisions_salt_value = s
        self.revisions_spin_count = spin_count

    def check_revisions_password(self, plaintext: str) -> bool:
        """Return ``True`` iff ``plaintext`` matches the stored hash."""
        if (
            self.revisions_hash_value is not None
            and self.revisions_salt_value is not None
            and self.revisions_algorithm_name is not None
            and self.revisions_spin_count is not None
        ):
            return verify_password_sha512(
                plaintext,
                self.revisions_hash_value,
                self.revisions_salt_value,
                spin_count=self.revisions_spin_count,
                algorithm=self.revisions_algorithm_name,
            )
        if self.revisions_password is None:
            return False
        return check_legacy_password(plaintext, self.revisions_password)

    # ------------------------------------------------------------------
    # Dict serialisation (RFC-058 §10)
    # ------------------------------------------------------------------

    def to_dict(self) -> dict[str, object]:
        """Return the patcher/writer-side flat dict (RFC-058 §10).

        Every slot in declaration order, the three lock flags as booleans.
        """
        flat: dict[str, object] = {
            name: getattr(self, name) for name in WorkbookProtection.__slots__
        }
        for lock in ("lock_structure", "lock_windows", "lock_revision"):
            flat[lock] = bool(flat[lock])
        return flat

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self, self.__attrs__)

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> Any:
        node = Element(tagname or self.tagname)
        for name, value in self:
            node.set(name, value)
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "WorkbookProtection":
        attrs = _attrs_from_node(node, cls.__attrs__)
        workbook_password = attrs.pop("workbookPassword", None)
        revisions_password = attrs.pop("revisionsPassword", None)
        protection = cls(**attrs)
        if workbook_password is not None:
            protection.set_workbook_password(workbook_password, already_hashed=True)
        if revisions_password is not None:
            protection.set_revisions_password(revisions_password, already_hashed=True)
        return protection

    def __eq__(self, other: object) -> bool:
        try:
            return _protection_dicts_equal(dict(self), dict(other))  # type: ignore[arg-type]
        except Exception:
            return NotImplemented

    def __repr__(self) -> str:  # pragma: no cover - debug aid
        return (
            "WorkbookProtection("
            f"lock_structure={self.lock_structure}, "
            f"lock_windows={self.lock_windows}, "
            f"lock_revision={self.lock_revision}, "
            f"workbook_password_set={self.workbook_hash_value is not None}, "
            f"revisions_password_set={self.revisions_hash_value is not None})"
        )

    # ------------------------------------------------------------------
    # CamelCase aliases for openpyxl compatibility.
    # ------------------------------------------------------------------

    @property
    def lockStructure(self) -> bool:  # noqa: N802
        return self.lock_structure

    @lockStructure.setter
    def lockStructure(self, value: bool) -> None:  # noqa: N802
        self.lock_structure = bool(value)

    @property
    def lockWindows(self) -> bool:  # noqa: N802
        return self.lock_windows

    @lockWindows.setter
    def lockWindows(self, value: bool) -> None:  # noqa: N802
        self.lock_windows = bool(value)

    @property
    def lockRevision(self) -> bool:  # noqa: N802
        return self.lock_revision

    @lockRevision.setter
    def lockRevision(self, value: bool) -> None:  # noqa: N802
        self.lock_revision = bool(value)

    @property
    def workbookPassword(self) -> str | None:  # noqa: N802
        return self.workbook_password

    @workbookPassword.setter
    def workbookPassword(self, value: str | None) -> None:  # noqa: N802
        self.workbook_password = value

    @property
    def workbookPasswordCharacterSet(self) -> str | None:  # noqa: N802
        return self.workbook_password_character_set

    @workbookPasswordCharacterSet.setter
    def workbookPasswordCharacterSet(self, value: str | None) -> None:  # noqa: N802
        self.workbook_password_character_set = value

    @property
    def workbookAlgorithmName(self) -> str | None:  # noqa: N802
        return self.workbook_algorithm_name

    @workbookAlgorithmName.setter
    def workbookAlgorithmName(self, value: str | None) -> None:  # noqa: N802
        self.workbook_algorithm_name = value

    @property
    def workbookHashValue(self) -> str | None:  # noqa: N802
        return self.workbook_hash_value

    @workbookHashValue.setter
    def workbookHashValue(self, value: str | None) -> None:  # noqa: N802
        self.workbook_hash_value = value

    @property
    def workbookSaltValue(self) -> str | None:  # noqa: N802
        return self.workbook_salt_value

    @workbookSaltValue.setter
    def workbookSaltValue(self, value: str | None) -> None:  # noqa: N802
        self.workbook_salt_value = value

    @property
    def workbookSpinCount(self) -> int | None:  # noqa: N802
        return self.workbook_spin_count

    @workbookSpinCount.setter
    def workbookSpinCount(self, value: int | None) -> None:  # noqa: N802
        self.workbook_spin_count = value

    @property
    def revisionsPassword(self) -> str | None:  # noqa: N802
        return self.revisions_password

    @revisionsPassword.setter
    def revisionsPassword(self, value: str | None) -> None:  # noqa: N802
        self.revisions_password = value

    @property
    def revisionsPasswordCharacterSet(self) -> str | None:  # noqa: N802
        return self.revisions_password_character_set

    @revisionsPasswordCharacterSet.setter
    def revisionsPasswordCharacterSet(self, value: str | None) -> None:  # noqa: N802
        self.revisions_password_character_set = value

    @property
    def revisionsAlgorithmName(self) -> str | None:  # noqa: N802
        return self.revisions_algorithm_name

    @revisionsAlgorithmName.setter
    def revisionsAlgorithmName(self, value: str | None) -> None:  # noqa: N802
        self.revisions_algorithm_name = value

    @property
    def revisionsHashValue(self) -> str | None:  # noqa: N802
        return self.revisions_hash_value

    @revisionsHashValue.setter
    def revisionsHashValue(self, value: str | None) -> None:  # noqa: N802
        self.revisions_hash_value = value

    @property
    def revisionsSaltValue(self) -> str | None:  # noqa: N802
        return self.revisions_salt_value

    @revisionsSaltValue.setter
    def revisionsSaltValue(self, value: str | None) -> None:  # noqa: N802
        self.revisions_salt_value = value

    @property
    def revisionsSpinCount(self) -> int | None:  # noqa: N802
        return self.revisions_spin_count

    @revisionsSpinCount.setter
    def revisionsSpinCount(self, value: int | None) -> None:  # noqa: N802
        self.revisions_spin_count = value


# ---------------------------------------------------------------------------
# FileSharing
# ---------------------------------------------------------------------------


class FileSharing:
    """``<fileSharing>`` block — read-only-recommended + reservation password.

    The ``read_only_recommended`` flag suggests Excel show the workbook
    in read-only mode by default. The reservation password (set via
    :meth:`set_reservation_password`) gates write access — without it,
    Excel falls back to read-only.
    """

    __slots__ = (
        "read_only_recommended",
        "user_name",
        "reservation_password",
        "algorithm_name",
        "hash_value",
        "salt_value",
        "spin_count",
    )

    __attrs__ = (
        "readOnlyRecommended",
        "userName",
        "reservationPassword",
        "algorithmName",
        "hashValue",
        "saltValue",
        "spinCount",
    )

    tagname = "fileSharing"

    def __init__(
        self,
        *,
        read_only_recommended: bool | None = None,
        user_name: str | None = None,
        reservation_password: str | None = None,
        algorithm_name: str | None = None,
        hash_value: str | None = None,
        salt_value: str | None = None,
        spin_count: int | None = None,
        readOnlyRecommended: bool | None = None,  # noqa: N803
        userName: str | None = None,  # noqa: N803
        reservationPassword: str | None = None,  # noqa: N803
        algorithmName: str | None = None,  # noqa: N803
        hashValue: str | None = None,  # noqa: N803
        saltValue: str | None = None,  # noqa: N803
        spinCount: int | None = None,  # noqa: N803
    ) -> None:
        if readOnlyRecommended is not None:
            read_only_recommended = readOnlyRecommended
        if userName is not None:
            user_name = userName
        if reservationPassword is not None:
            reservation_password = reservationPassword
        if algorithmName is not None:
            algorithm_name = algorithmName
        if hashValue is not None:
            hash_value = hashValue
        if saltValue is not None:
            salt_value = saltValue
        if spinCount is not None:
            spin_count = spinCount

        self.read_only_recommended: bool | None = read_only_recommended
        self.user_name: str | None = user_name
        self.reservation_password: str | None = reservation_password
        self.algorithm_name: str | None = algorithm_name
        self.hash_value: str | None = hash_value
        self.salt_value: str | None = salt_value
        self.spin_count: int | None = spin_count

    def set_reservation_password(
        self,
        plaintext: str,
        algorithm: str = "SHA-512",
        *,
        salt: bytes | None = None,
        spin_count: int = 100_000,
    ) -> None:
        """Hash ``plaintext`` and store it on the reservation-password slot."""
        h, s = hash_password_sha512(
            plaintext,
            salt=salt,
            spin_count=spin_count,
            algorithm=algorithm,
        )
        self.algorithm_name = algorithm
        self.hash_value = h
        self.salt_value = s
        self.spin_count = spin_count
        self.reservation_password = hash_password(plaintext)

    def check_reservation_password(self, plaintext: str) -> bool:
        """Return ``True`` iff ``plaintext`` matches the stored hash."""
        if (
            self.hash_value is not None
            and self.salt_value is not None
            and self.algorithm_name is not None
            and self.spin_count is not None
        ):
            return verify_password_sha512(
                plaintext,
                self.hash_value,
                self.salt_value,
                spin_count=self.spin_count,
                algorithm=self.algorithm_name,
            )
        if self.reservation_password is None:
            return False
        return check_legacy_password(plaintext, self.reservation_password)

    def to_dict(self) -> dict[str, object]:
        """Return the patcher/writer-side flat dict (RFC-058 §10).

        Every slot in declaration order, ``read_only_recommended`` as a boolean.
        """
        flat: dict[str, object] = {
            name: getattr(self, name) for name in FileSharing.__slots__
        }
        flat["read_only_recommended"] = bool(flat["read_only_recommended"])
        return flat

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self, self.__attrs__)

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> Any:
        node = Element(tagname or self.tagname)
        for name, value in self:
            node.set(name, value)
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "FileSharing":
        return cls(**_attrs_from_node(node, cls.__attrs__))

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, FileSharing):
            return NotImplemented
        return dict(self) == dict(other)

    def __repr__(self) -> str:  # pragma: no cover - debug aid
        return (
            "FileSharing("
            f"read_only_recommended={self.read_only_recommended}, "
            f"user_name={self.user_name!r}, "
            f"reservation_password_set={self.hash_value is not None})"
        )

    @property
    def readOnlyRecommended(self) -> bool | None:  # noqa: N802
        return self.read_only_recommended

    @readOnlyRecommended.setter
    def readOnlyRecommended(self, value: bool | None) -> None:  # noqa: N802
        self.read_only_recommended = None if value is None else bool(value)

    @property
    def userName(self) -> str | None:  # noqa: N802
        return self.user_name

    @userName.setter
    def userName(self, value: str | None) -> None:  # noqa: N802
        self.user_name = value

    @property
    def reservationPassword(self) -> str | None:  # noqa: N802
        return self.reservation_password

    @reservationPassword.setter
    def reservationPassword(self, value: str | None) -> None:  # noqa: N802
        self.reservation_password = value

    @property
    def algorithmName(self) -> str | None:  # noqa: N802
        return self.algorithm_name

    @algorithmName.setter
    def algorithmName(self, value: str | None) -> None:  # noqa: N802
        self.algorithm_name = value

    @property
    def hashValue(self) -> str | None:  # noqa: N802
        return self.hash_value

    @hashValue.setter
    def hashValue(self, value: str | None) -> None:  # noqa: N802
        self.hash_value = value

    @property
    def saltValue(self) -> str | None:  # noqa: N802
        return self.salt_value

    @saltValue.setter
    def saltValue(self, value: str | None) -> None:  # noqa: N802
        self.salt_value = value

    @property
    def spinCount(self) -> int | None:  # noqa: N802
        return self.spin_count

    @spinCount.setter
    def spinCount(self, value: int | None) -> None:  # noqa: N802
        self.spin_count = value


DocumentSecurity = WorkbookProtection


def _attrs_from_node(node: Any, names: tuple[str, ...]) -> dict[str, Any]:
    attrs: dict[str, Any] = {}
    for name in names:
        value = node.get(name)
        if value is None:
            continue
        attrs[name] = _typed_attr(name, value)
    return attrs


def _typed_attr(name: str, value: str) -> Any:
    if name in {
        "lockStructure",
        "lockWindows",
        "lockRevision",
        "readOnlyRecommended",
    }:
        return value.lower() in {"1", "true"}
    if name in {"workbookSpinCount", "revisionsSpinCount", "spinCount"}:
        return int(value)
    return value


__all__ = ["WorkbookProtection", "FileSharing", "DocumentSecurity"]

__getattr__ = _openpyxl_name_fallback(globals())
