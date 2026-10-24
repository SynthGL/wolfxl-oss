"""Person registry entry for threaded comments (G08 step 2).

Excel 365 stores threaded-comment authors in a workbook-scoped registry at
``xl/persons/personList.xml``. Each entry has a stable GUID that the
``threadedComment`` payload references via ``personId``. WolfXL mirrors
openpyxl's surface: ``Person(name=..., user_id=..., provider_id=...)``
with an auto-allocated ``id`` GUID at registration time.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator


def _new_guid() -> str:
    """Return a lowercase brace-wrapped GUID matching Excel's threadedComments style."""
    import uuid

    return "{" + str(uuid.uuid4()).upper() + "}"


@dataclass
class Person:
    """A threaded-comment author.

    The ``id`` field is auto-allocated when ``None`` so callers can build
    a ``Person`` without thinking about GUIDs; the registry assigns one
    on first ``add()``.
    """

    name: str
    user_id: str = ""
    provider_id: str = "None"
    id: str | None = None  # noqa: A003 - mirror openpyxl public name


class PersonRegistry:
    """Workbook-scoped, insertion-ordered Person registry.

    ``add()`` is idempotent on ``(user_id, provider_id)`` when both are
    non-empty — the second call returns the original ``Person`` rather
    than allocating a new GUID. This matches the RFC-068 mitigation for
    drift across reload.
    """

    def __init__(self) -> None:
        self._items: list[Person] = []

    def add(
        self,
        *,
        name: str,
        user_id: str = "",
        provider_id: str = "None",
    ) -> Person:
        if user_id and provider_id:
            for existing in self._items:
                if existing.user_id == user_id and existing.provider_id == provider_id:
                    return existing
        person = Person(
            name=name,
            user_id=user_id,
            provider_id=provider_id,
            id=_new_guid(),
        )
        self._items.append(person)
        return person

    def __iter__(self) -> Iterator[Person]:
        return iter(self._items)

    def __len__(self) -> int:
        return len(self._items)

    def __getitem__(self, index: int) -> Person:
        return self._items[index]

    def by_id(self, person_id: str) -> Person | None:
        for p in self._items:
            if p.id == person_id:
                return p
        return None

    def _seed(self, person: Person) -> None:
        """Append a ``Person`` with its pre-existing GUID intact.

        Used by the reader hydration path to preserve byte-stable round
        trips: if the file already has GUIDs, we reuse them rather than
        allocating fresh ones via ``add()``. Idempotent on ``id`` so a
        second call with the same ``Person.id`` is a no-op.
        """
        if person.id is None:
            return
        if any(existing.id == person.id for existing in self._items):
            return
        self._items.append(person)
