"""Extract cached native Excel results after a successful native save.

Usage: python extract_math_information.py CASES.json RECEIPT.json OUT.json
Keep the native save output beside RECEIPT.json during extraction.
Only portable provenance and formula results enter OUT.json.
"""

import hashlib
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def main():
    cases_path, receipt_path, output_path = map(Path, sys.argv[1:])
    cases = json.loads(cases_path.read_text())
    receipt = json.loads(receipt_path.read_text())
    (record,) = receipt["results"]
    if record["status"] != "saved":
        raise ValueError(f"Native workbook was not saved: {record['status']}")
    saved = receipt_path.parent / record["saved_file"]
    digest = hashlib.sha256(saved.read_bytes()).hexdigest()
    if digest != record["saved_sha256"]:
        raise ValueError("Saved workbook digest differs from native receipt")
    namespace = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(saved) as book:
        sheet = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
    cells = {cell.attrib["r"]: cell for cell in sheet.findall(".//s:c", namespace)}
    for index, case in enumerate(cases, 1):
        cell = cells[f"B{index}"]
        formula = cell.find("s:f", namespace)
        cached = cell.find("s:v", namespace)
        if formula is None or cached is None or cached.text is None:
            raise ValueError(f"Missing native formula/cache for B{index}")
        if formula.text.replace("E+", "E") != case["formula"][1:].replace("E+", "E"):
            raise ValueError(f"Native formula changed for B{index}: {formula.text}")
        kind = cell.get("t", "n")
        if kind == "n":
            expected = {"type": "number", "value": float(cached.text)}
        elif kind == "b":
            expected = {"type": "boolean", "value": cached.text == "1"}
        elif kind == "e":
            expected = {"type": "error", "value": cached.text}
        else:
            raise ValueError(f"Unrecognized cache type for B{index}: {kind}")
        case["expected"] = expected
    result = {
        "source": {
            "application": receipt["app"],
            "version": receipt["app_version"],
            "captured_at": receipt["finished_at"],
            "input_sha256": record["sha256"],
            "saved_sha256": digest,
            "runner_sha256": receipt["job"]["runner_sha256"],
        },
        "cases": cases,
    }
    output_path.write_text(json.dumps(result, indent=2, allow_nan=False) + "\n")
    print(f"Extracted {len(cases)} native cached results to {output_path}")


if __name__ == "__main__":
    main()
