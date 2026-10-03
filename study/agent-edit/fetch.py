"""Fetch upstream study inputs into a local, ignored cache; never rehost them."""

import hashlib
import http.client
import io
import sys
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CACHE = ROOT / ".cache"
CONTOSO_URL = (
    "https://download.microsoft.com/download/b/e/c/"
    "becf5873-6b88-4920-9096-2c10ba98de60/ContosoPnL_Excel2013.zip"
)
CONTOSO_SHA256 = "d9a77819ed43a93cda82ab7ba08dd6406981d9d3eae7cfb7cae4720a40f87ac5"
CONTOSO = CACHE / "real-excel-powerpivot-contoso-pnl.xlsx"
SKILLS_COMMIT = "33375500bcea98d610eb30ce10ac4e59b89c390d"
RECALC = CACHE / "reference" / "recalc.py"
REFERENCE_FILES = {
    "recalc.py": "66f6e726d5a53cfabea0b6471d6669deb95e54e96cb7913b687e4388da1bb3d9",
    "office/soffice.py": "df2c8d7249c132a4512d744482e4b5ed0ecbe885f0bcfb1295e7e6a7ced0432c",
}


def verify(data, expected, label):
    actual = hashlib.sha256(data).hexdigest()
    if actual != expected:
        raise ValueError(
            f"SHA-256 mismatch for {label}: expected {expected}, got {actual}"
        )
    return data


def download(url):
    with urllib.request.urlopen(url, timeout=60) as response:
        return response.read()


def fetch_reference():
    """Download the exact upstream recalculation code used by the original study."""
    for name, digest in REFERENCE_FILES.items():
        target = RECALC.parent / name
        if target.exists():
            verify(target.read_bytes(), digest, name)
            continue
        url = f"https://raw.githubusercontent.com/anthropics/skills/{SKILLS_COMMIT}/skills/xlsx/scripts/{name}"
        data = verify(download(url), digest, name)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)


def fetch_contoso():
    """Verify even cached workbooks; a mismatch is fatal, not a seven-file fallback."""
    if CONTOSO.exists():
        verify(CONTOSO.read_bytes(), CONTOSO_SHA256, CONTOSO.name)
        return CONTOSO
    try:
        data = download(CONTOSO_URL)
    except (urllib.error.URLError, OSError, http.client.IncompleteRead) as exc:
        print(
            f"NOTE: Microsoft Contoso download unavailable ({exc}); running on 7 files.",
            file=sys.stderr,
        )
        return None
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        candidates = [n for n in archive.namelist() if n.lower().endswith(".xlsx")]
        if len(candidates) != 1:
            raise ValueError(
                f"Expected one workbook in Microsoft archive, got {candidates}"
            )
        workbook = verify(archive.read(candidates[0]), CONTOSO_SHA256, candidates[0])
    CACHE.mkdir(parents=True, exist_ok=True)
    CONTOSO.write_bytes(workbook)
    return CONTOSO


if __name__ == "__main__":
    fetch_reference()
    workbook = fetch_contoso()
    print("Reference recalculator verified:", SKILLS_COMMIT)
    print("Contoso:", "verified" if workbook else "unavailable; 7-file study only")
