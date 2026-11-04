"""Build tests/fixtures/data-table-cities.xlsx: two side-by-side Excel tables
under title rows, totals rows of SUBTOTAL formulas, and a sheet of structured
references, every formula carrying the cached value Excel would store.

Run: uv run --with openpyxl==3.1.5 python generate_data_table_cities.py OUT.xlsx
The output is byte-for-byte deterministic.
"""

import io
import re
import sys
import unicodedata
import zipfile
from datetime import datetime

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

BIG = [
    ("Shanghai, China", 31.23, 121.5, 24256800),
    ("Karachi, Pakistan", 24.86, 67.01, 23500000),
    ("Beijing, China", 39.93, 116.4, 21516000),
    ("São Paulo, Brazil", -23.53, -46.63, 21292900),
    ("Delhi, India", 28.67, 77.21, 16788000),
    ("Lagos, Nigeria", 6.45, 3.47, 16060000),
    ("Istanbul, Turkey", 41.1, 29, 14657000),
    ("Tokyo, Japan", 35.67, 139.8, 13298000),
    ("Mumbai, India", 18.96, 72.82, 12478000),
    ("Moscow, Russia", 55.75, 37.62, 12198000),
    ("Guangzhou, China", 23.12, 113.3, 12081000),
    ("Shenzhen, China", 22.55, 114.1, 10780000),
]
SMALL = [
    ("Pyongyang, North Korea", 39.02, 125.74, 2581000),
    ("Buenos Aires, Argentina", -34.6, -58.38, 2890000),
    ("Seattle, Washington, USA", 47.61, -122.33, 609000),
    ("Toronto, Canada", 43.7, -79.4, 2615000),
    ("Auckland, New Zealand", -36.84, 174.74, 1454000),
    ("Miami, Florida, USA", 25.78, -80.21, 400000),
    ("Havana, Cuba", 23.13, -82.38, 2106000),
    ("Fairbanks, Alaska, USA", 64.84, -147.72, 32000),
    ("Longyearbyen, Svalbard", 78.22, 15.55, 2600),
    ("Johannesburg, South Africa", -26.2, 28, 957000),
    ("Cancun, Mexico", 21.16, -86.85, 722800),
    ("Oahu, Hawaii, USA", 21.47, -157.98, 953000),
]
HEADERS = ["City", "Latitude", "Longitude", "Population"]
FIXED = datetime(2026, 11, 3, 12, 0, 0)


# The averages Excel stores for these columns. Excel's summation rounds
# differently from Python's left-to-right float sum in the last bits, and a
# fixture should hold what Excel writes.
EXCEL_AVERAGES = {
    "BigCity": (25.396666666666672, 70.46666666666665),
    "SmallCity": (22.274166666666662, -39.26833333333334),
}


def totals(name, rows):
    latitude, longitude = EXCEL_AVERAGES[name]
    return latitude, longitude, sum(r[3] for r in rows)


def excel_collation(row):
    # Excel sorts pivot row labels ignoring accents and case: "São" < "Sha".
    folded = unicodedata.normalize("NFKD", row[0])
    return "".join(c for c in folded if not unicodedata.combining(c)).casefold()


def add_table(ws, name, first_col, title, rows):
    letters = [chr(ord(first_col) + i) for i in range(4)]
    ws[f"{letters[0]}1"] = title
    for letter, header in zip(letters, HEADERS):
        ws[f"{letter}2"] = header
    for offset, row in enumerate(rows, start=3):
        for letter, value in zip(letters, row):
            ws[f"{letter}{offset}"] = value
    last = 3 + len(rows)
    ws[f"{letters[0]}{last}"] = "Total"
    ws[f"{letters[1]}{last}"] = f"=SUBTOTAL(101,{name}[Latitude])"
    ws[f"{letters[2]}{last}"] = f"=SUBTOTAL(101,{name}[Longitude])"
    ws[f"{letters[3]}{last}"] = f"=SUBTOTAL(109,{name}[Population])"
    table = Table(
        displayName=name,
        name=name,
        ref=f"{letters[0]}2:{letters[3]}{last}",
        totalsRowCount=1,
    )
    table.tableColumns = [
        TableColumn(id=1, name="City", totalsRowLabel="Total"),
        TableColumn(id=2, name="Latitude", totalsRowFunction="average"),
        TableColumn(id=3, name="Longitude", totalsRowFunction="average"),
        TableColumn(id=4, name="Population", totalsRowFunction="sum"),
    ]
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    lat, lon, pop = totals(name, rows)
    return {
        f"{letters[1]}{last}": lat,
        f"{letters[2]}{last}": lon,
        f"{letters[3]}{last}": pop,
    }


def build():
    wb = Workbook()
    formula = wb.active
    formula.title = "Formula"
    table_ws = wb.create_sheet("Table")
    pivot = wb.create_sheet("PivotTable")

    cached = {"Table": {}, "Formula": {}}
    cached["Table"].update(add_table(table_ws, "BigCity", "C", "Big Cities", BIG))
    cached["Table"].update(add_table(table_ws, "SmallCity", "I", "Small Cities", SMALL))

    for top, name, label in (
        (3, "BigCity", "world's 12 largest cities"),
        (7, "SmallCity", "12 random cities"),
    ):
        formula[f"A{top}"] = (
            "Big City Average Location"
            if name == "BigCity"
            else "Small City Average Location"
        )
        prefix = (
            "Total city proper population in world's "
            if name == "BigCity"
            else "Total city population in "
        )
        suffix = " largest cities" if name == "BigCity" else " random cities"
        formula[f"D{top}"] = f'="{prefix}"&COUNTA({name}[City])&"{suffix}"'
        formula[f"A{top + 1}"] = "Latitude"
        formula[f"B{top + 1}"] = "Longitude"
        formula[f"A{top + 2}"] = f"={name}[[#Totals],[Latitude]]"
        formula[f"B{top + 2}"] = f"={name}[[#Totals],[Longitude]]"
        formula[f"D{top + 2}"] = f"={name}[[#Totals],[Population]]"
        rows = BIG if name == "BigCity" else SMALL
        lat, lon, pop = totals(name, rows)
        cached["Formula"].update(
            {
                f"D{top}": f"{prefix}12{suffix}",
                f"A{top + 2}": lat,
                f"B{top + 2}": lon,
                f"D{top + 2}": pop,
            }
        )
    formula["D4"] = "Source: city proper populations as listed on Wikipedia"

    pivot["B5"] = "Row Labels"
    pivot["C5"] = "Sum of Latitude"
    pivot["D5"] = "Sum of Longitude"
    pivot["E5"] = "Sum of Population"
    for offset, row in enumerate(sorted(BIG, key=excel_collation), start=6):
        (
            pivot[f"B{offset}"],
            pivot[f"C{offset}"],
            pivot[f"D{offset}"],
            pivot[f"E{offset}"],
        ) = row
    grand = 6 + len(BIG)
    pivot[f"B{grand}"] = "Grand Total"
    pivot[f"C{grand}"] = sum(r[1] for r in BIG)
    # The longitude total Excel stores; a left-to-right float sum gives 845.6.
    pivot[f"D{grand}"] = 845.6000000000001
    pivot[f"E{grand}"] = sum(r[3] for r in BIG)

    wb.properties.creator = "wolfxl-cli tests"
    wb.properties.created = FIXED
    wb.properties.modified = FIXED
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), cached, wb.sheetnames


def number_text(value):
    return repr(float(value)) if not float(value).is_integer() else str(int(value))


def inject(xml, values):
    def replace(match):
        ref = match.group("ref")
        if ref not in values:
            return match.group(0)
        value = values[ref]
        attrs, formula_text = match.group("attrs"), match.group("f")
        attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
        if isinstance(value, str):
            escaped = (
                value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            )
            return (
                f'<c r="{ref}"{attrs} t="str"><f>{formula_text}</f><v>{escaped}</v></c>'
            )
        return (
            f'<c r="{ref}"{attrs}><f>{formula_text}</f><v>{number_text(value)}</v></c>'
        )

    pattern = re.compile(
        r'<c r="(?P<ref>[A-Z]+\d+)"(?P<attrs>[^>]*)><f>(?P<f>[^<]*)</f>(?:<v\s*/>|<v></v>|<v>[^<]*</v>)?</c>'
    )
    return pattern.sub(replace, xml)


def main(path):
    raw, cached, names = build()
    source = zipfile.ZipFile(io.BytesIO(raw))
    sheet_paths = {
        name: f"xl/worksheets/sheet{index}.xml"
        for index, name in enumerate(names, start=1)
    }
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "docProps/core.xml":
                # openpyxl stamps the save time; pin it so the bytes repeat.
                data = re.sub(
                    rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:)",
                    rb"\g<1>2026-11-03T12:00:00Z\g<2>",
                    data,
                )
            for sheet, values in cached.items():
                if info.filename == sheet_paths[sheet]:
                    xml = inject(data.decode("utf-8"), values)
                    for ref in values:
                        if not re.search(rf'<c r="{ref}"[^>]*><f>[^<]*</f><v>', xml):
                            raise SystemExit(
                                f"no cached value injected for {sheet}!{ref}"
                            )
                    data = xml.encode("utf-8")
            entry = zipfile.ZipInfo(info.filename, date_time=(2026, 11, 3, 12, 0, 0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.external_attr = 0o644 << 16
            target.writestr(entry, data)
    with open(path, "wb") as handle:
        handle.write(output.getvalue())


if __name__ == "__main__":
    main(sys.argv[1])
