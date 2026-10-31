#!/usr/bin/env python3
"""Write a pip hash-checking requirements file for locally built artifacts.

Usage:

    python .github/scripts/hashed_artifact_requirements.py \
        --output PATH ARTIFACT_GLOB [ARTIFACT_GLOB ...]

Each ARTIFACT_GLOB may be a quoted glob pattern (expanded here, so the
behaviour is the same on every shell and platform) or an already expanded
path. Every pattern must match at least one regular file. Matches are
deduplicated by resolved path and written in sorted order as:

    file:///<absolute artifact path> --hash=sha256:<digest>

The output is meant for ``pip install --require-hashes --no-deps -r PATH``.
Nothing is written when any pattern fails to match or names a non-regular
file. Standard library only.
"""

from __future__ import annotations

import argparse
import glob
import hashlib
import os
import stat
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence

CHUNK_SIZE = 1024 * 1024


class ArtifactError(Exception):
    """Raised when the artifact patterns do not describe a usable file set."""


def sha256_file(path: Path, chunk_size: int = CHUNK_SIZE) -> str:
    """Return the hex SHA-256 digest of ``path``, read in bounded chunks."""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(chunk_size)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def expand_patterns(patterns: Iterable[str]) -> List[Path]:
    """Expand glob patterns and literal paths into sorted, unique files.

    Raises ArtifactError listing every pattern that matched nothing and every
    match that is not a regular file.
    """
    errors: List[str] = []
    resolved: Dict[str, Path] = {}
    for pattern in patterns:
        matches = glob.glob(pattern, recursive=True)
        if not matches and os.path.lexists(pattern):
            # Literal paths containing glob metacharacters such as "[".
            matches = [pattern]
        if not matches:
            errors.append(f"pattern matched no files: {pattern}")
            continue
        for match in matches:
            path = Path(match)
            try:
                real = path.resolve(strict=True)
                mode = os.stat(real).st_mode
            except OSError as exc:
                errors.append(f"cannot access {match}: {exc}")
                continue
            if not stat.S_ISREG(mode):
                errors.append(f"not a regular file: {match}")
                continue
            resolved[real.as_uri()] = real
    if errors:
        raise ArtifactError("\n".join(errors))
    return [resolved[uri] for uri in sorted(resolved)]


def requirement_lines(paths: Sequence[Path]) -> List[str]:
    """Return one hash-pinned direct-URL requirement line per artifact."""
    return [f"{path.as_uri()} --hash=sha256:{sha256_file(path)}" for path in paths]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Write a pip --require-hashes requirements file for locally "
            "built wheel and sdist artifacts."
        )
    )
    parser.add_argument(
        "--output",
        required=True,
        type=Path,
        help="requirements file to write",
    )
    parser.add_argument(
        "patterns",
        nargs="+",
        metavar="ARTIFACT_GLOB",
        help="artifact path or quoted glob pattern; each must match a file",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        paths = expand_patterns(args.patterns)
    except ArtifactError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    lines = requirement_lines(paths)
    with args.output.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write("".join(f"{line}\n" for line in lines))
    for path in paths:
        print(f"hashed {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
