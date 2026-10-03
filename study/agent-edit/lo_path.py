"""WolfXL modify edit followed by the reference LibreOffice recalc.py; same census as run.py."""

import hashlib
import json
import shutil
import sys

import wolfxl

from fetch import CONTOSO_URL, fetch_reference
from run import (
    CELL,
    FORMULA,
    IGNORED_PARTS,
    ROOT,
    environment,
    failed,
    parts,
    recalculate,
    sources,
    summarize,
)


def main():
    fetch_reference()
    inputs = sources()
    out_dir = ROOT / "out_lo"
    out_dir.mkdir(exist_ok=True)
    res = {
        **environment(),
        "wolfxl": wolfxl.__version__,
        "edit": f"{CELL} {FORMULA}",
        "files": {},
        "source_sha256": {
            src.name: hashlib.sha256(src.read_bytes()).hexdigest() for src in inputs
        },
        "contoso_url": CONTOSO_URL,
        "contoso_available": len(inputs) == 8,
    }
    for src in inputs:
        out = out_dir / f"{src.stem}.wolfxl-lo{src.suffix}"
        try:
            shutil.copy(src, out)
            wb = wolfxl.load_workbook(
                str(out), modify=True, keep_vba=out.suffix == ".xlsm"
            )
            sheet = wb.sheetnames[0]
            wb[sheet][CELL] = FORMULA
            wb.save(str(out))
            wb.close()
            edit_missing = sorted(set(parts(src)) - set(parts(out)) - IGNORED_PARTS)
            row = summarize(
                src,
                out,
                sheet,
                {"edit_step_parts_missing": edit_missing, "recalc": recalculate(out)},
            )
        except (
            Exception
        ) as exc:  # Preserve the failure in the receipt, then exit nonzero.
            row = {"error": f"{type(exc).__name__}: {exc}"[:400]}
        res["files"][src.name] = row
        if "error" in row:
            print(f"{src.name:55} ERROR {row['error']}")
            continue
        print(
            f"{src.name:55} edit_missing={len(edit_missing)} lo_missing={len(row['parts_missing']):3} "
            f"lost={row['families_lost']} changed={row['sheet_features_changed']} "
            f"regressions={row['cached_value_regressions']} cached={row['cached_value']}"
        )
    (ROOT / "results_lo.json").write_text(json.dumps(res, indent=2, default=str))
    return int(any(failed(row) for row in res["files"].values()))


if __name__ == "__main__":
    sys.exit(main())
