"""Generate the math/information workbook for native Excel recalculation.

Run: uv run --with openpyxl==3.1.5 python capture_math_information.py OUT.xlsx.
Pass an output workbook path. Cells A1:A6
match the evaluator's MemorySource fixture. B contains the oracle formulas.
"""

import json
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.workbook.properties import CalcProperties

CASES = {
    "ACOS": ["ACOS(0.5)", "ACOS(1)", "ACOS(-1)", "ACOS(2)"],
    "ASIN": ["ASIN(0.5)", "ASIN(1)", "ASIN(-1)", "ASIN(2)"],
    "ATAN": ["ATAN(1)", "ATAN(0)", "ATAN(-1)"],
    "ATAN2": ["ATAN2(1,1)", "ATAN2(-1,0)", "ATAN2(0,-1)", "ATAN2(0,0)"],
    "COS": ["COS(1)", "COS(0)", "COS(-1)", "COS(134217727)", "COS(134217728)"],
    "SIN": ["SIN(1)", "SIN(0)", "SIN(-1)", "SIN(134217727)", "SIN(134217728)"],
    "TAN": ["TAN(1)", "TAN(0)", "TAN(-1)", "TAN(134217727)", "TAN(134217728)"],
    "DEGREES": ["DEGREES(1)", "DEGREES(0)", "DEGREES(-1)"],
    "RADIANS": ["RADIANS(180)", "RADIANS(0)", "RADIANS(-90)"],
    "PI": ["PI()"],
    "EXP": ["EXP(1)", "EXP(0)", "EXP(-1)", "EXP(710)"],
    "LN": ["LN(2)", "LN(1)", "LN(0)", "LN(-1)"],
    "LOG": [
        "LOG(100)",
        "LOG(8,2)",
        "LOG(1)",
        "LOG(0)",
        "LOG(1,0)",
        "LOG(2,1)",
        "LOG(1,-2)",
        "LOG(100,)",
    ],
    "LOG10": ["LOG10(100)", "LOG10(1)", "LOG10(0)", "LOG10(-1)"],
    "CEILING": [
        "CEILING(2.3,0.2)",
        "CEILING(-2.3,0.2)",
        "CEILING(-2.3,-0.2)",
        "CEILING(2.3,-0.2)",
        "CEILING(0,0)",
        "CEILING(2,0)",
        "CEILING(0,-2)",
        "CEILING(0.3,0.1)",
    ],
    "FLOOR": [
        "FLOOR(2.3,0.2)",
        "FLOOR(-2.3,0.2)",
        "FLOOR(-2.3,-0.2)",
        "FLOOR(2.3,-0.2)",
        "FLOOR(0,0)",
        "FLOOR(2,0)",
        "FLOOR(0,-2)",
        "FLOOR(0.3,0.1)",
    ],
    "TRUNC": [
        "TRUNC(8.9)",
        "TRUNC(-8.9)",
        "TRUNC(1.239,2)",
        "TRUNC(123,-1)",
        "TRUNC(0,400)",
        "TRUNC(1,400)",
        "TRUNC(123,-400)",
    ],
    "PRODUCT": [
        "PRODUCT(2,3)",
        "PRODUCT(A1:A6)",
        'PRODUCT(TRUE,"3",2)',
        'PRODUCT({2,TRUE,"3"})',
        "PRODUCT(A4)",
        "PRODUCT(,5)",
        "PRODUCT(0,5)",
        "PRODUCT(9E307,2)",
    ],
    "RAND": ["AND(RAND()>=0,RAND()<1)"],
    "RANDBETWEEN": [
        "RANDBETWEEN(4,4)",
        "RANDBETWEEN(1.1,1.9)",
        "RANDBETWEEN(-2.9,-1.1)",
        "RANDBETWEEN(2,1)",
        "AND(RANDBETWEEN(-2,2)>=-2,RANDBETWEEN(-2,2)<=2)",
    ],
    "ISBLANK": [
        "ISBLANK(A4)",
        "ISBLANK(0)",
        'ISBLANK("")',
        "ISBLANK(#N/A)",
        "ISBLANK(A1:A4)",
        'ISBLANK({1,"";TRUE,#N/A})',
    ],
    "ISERROR": [
        "ISERROR(#N/A)",
        "ISERROR(1/0)",
        'ISERROR("#N/A")',
        "ISERROR(A4)",
        "ISERROR({1,#N/A;#VALUE!,#REF!})",
    ],
    "ISNA": [
        "ISNA(#N/A)",
        "ISNA(#NUM!)",
        'ISNA("#N/A")',
        "ISNA(A4)",
        "ISNA({#N/A,1;#DIV/0!,#N/A})",
    ],
    "ISNUMBER": [
        "ISNUMBER(1)",
        'ISNUMBER("1")',
        "ISNUMBER(TRUE)",
        "ISNUMBER(A4)",
        "ISNUMBER(#N/A)",
        "ISNUMBER(A1:A4)",
    ],
    "ISTEXT": [
        'ISTEXT("x")',
        'ISTEXT("")',
        "ISTEXT(1)",
        "ISTEXT(TRUE)",
        "ISTEXT(A4)",
        "ISTEXT(#N/A)",
        "ISTEXT(A1:A4)",
    ],
    "NA": ["NA()"],
}


def rows():
    result = []
    for name, formulas in CASES.items():
        for formula in formulas:
            # SUM preserves array results as a single observable scalar in the
            # source workbook; array shapes get their own runtime assertions.
            if name.startswith("IS") and ("{" in formula or ":" in formula):
                formula = "SUM(--" + formula + ")"
            result.append({"function": name, "formula": "=" + formula})
        if name not in {"PI", "RAND", "NA"} and not name.startswith("IS"):
            for argument in ['"x"', "#N/A", "#VALUE!", "#NUM!", "#DIV/0!"]:
                suffix = ",2" if name in {"ATAN2", "CEILING", "FLOOR", "RANDBETWEEN"} else ""
                result.append({"function": name, "formula": f"={name}({argument}{suffix})"})
    for name in ["ISBLANK", "ISERROR", "ISNA", "ISNUMBER", "ISTEXT"]:
        for argument, rows, columns in [('{1,"";TRUE,#N/A}', 2, 2), ("A1:A4", 4, 1)]:
            array_formula = f"={name}({argument})"
            for row in range(1, rows + 1):
                for column in range(1, columns + 1):
                    result.append(
                        {
                            "function": name,
                            "formula": f"=INDEX({array_formula[1:]},{row},{column})",
                            "array_formula": array_formula,
                            "row": row,
                            "column": column,
                            "rows": rows,
                            "columns": columns,
                        }
                    )
    return result


def main():
    target = Path(sys.argv[1])
    manifest = target.with_suffix(".cases.json")
    cases = rows()
    if len(sys.argv) > 2:
        cases = cases[: int(sys.argv[2])]
    manifest.write_text(json.dumps(cases, indent=2) + "\n")
    book = Workbook()
    sheet = book.active
    sheet.title = "Sheet1"
    for index, value in [(1, 2), (2, "3"), (3, True), (5, 8), (6, "not a number")]:
        sheet.cell(index, 1, value)
    for index, case in enumerate(cases, 1):
        formula = case["formula"]
        if case["function"].startswith("IS") and (":" in formula or "{" in formula):
            # Declare array evaluation explicitly. A legacy scalar workbook
            # formula may implicitly intersect its argument in native Excel.
            sheet.cell(index, 2, ArrayFormula(ref=f"B{index}", text=formula))
        else:
            sheet.cell(index, 2, formula)
    book.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True)
    book.save(target)
    print(f"Generated {len(cases)} cases: {target}")


if __name__ == "__main__":
    main()
