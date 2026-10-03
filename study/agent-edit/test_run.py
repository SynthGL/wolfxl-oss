"""Offline verdict checks using the public Excel fixtures and synthetic x14 probe."""

import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import run

FIXTURES = Path(__file__).resolve().parents[2] / "tests/fixtures/external_oracle"
CHART = FIXTURES / "real-excel-chart-cf-basic.xlsx"
VALIDATION = FIXTURES / "real-excel-p1-comments-validation-protection.xlsx"
X14_PROBE = Path(__file__).resolve().parent / "x14-validation-sparkline.xlsx"
SHEET = "xl/worksheets/sheet1.xml"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def rewrite_package(source, output, replacements=None, omitted=()):
    replacements = replacements or {}
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(output, "w") as after:
        for item in before.infolist():
            if item.filename not in omitted:
                after.writestr(
                    item, replacements.get(item.filename, before.read(item.filename))
                )
        for name, data in replacements.items():
            if name not in before.namelist():
                after.writestr(name, data)


class VerdictTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.output = Path(self.directory.name) / "output.xlsx"

    def test_worksheet_parts_resolves_public_fixture_relationships(self):
        with zipfile.ZipFile(CHART) as archive:
            self.assertEqual(run.worksheet_parts(archive), {"Data": SHEET})

    def test_worksheet_parts_uses_absolute_relationship_target_not_sheet_order(self):
        with zipfile.ZipFile(CHART) as archive:
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            for rel in rels:
                if rel.get("Id") == "rId1":
                    rel.set("Target", "/xl/worksheets/reordered.xml")
            worksheet = archive.read(SHEET)
        rewrite_package(
            CHART,
            self.output,
            {
                "xl/_rels/workbook.xml.rels": ET.tostring(rels),
                "xl/worksheets/reordered.xml": worksheet,
            },
            omitted=(SHEET,),
        )
        with zipfile.ZipFile(self.output) as archive:
            self.assertEqual(
                run.worksheet_parts(archive), {"Data": "xl/worksheets/reordered.xml"}
            )
        self.assertEqual(run.scan(self.output)[0]["cf"], 3)

    def test_worksheet_parts_excludes_missing_and_nonworksheet_targets(self):
        with zipfile.ZipFile(CHART) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            sheets = workbook.find(f"{{{run.MAIN}}}sheets")
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        for name, target, relationship in (
            ("Chart only", "chartsheets/chart.xml", "rIdChart"),
            ("Missing", "worksheets/missing.xml", "rIdMissing"),
        ):
            ET.SubElement(
                sheets,
                f"{{{run.MAIN}}}sheet",
                {
                    "name": name,
                    f"{{{run.REL}}}id": relationship,
                },
            )
            ET.SubElement(
                rels,
                f"{{{PACKAGE_REL}}}Relationship",
                {
                    "Id": relationship,
                    "Target": target,
                },
            )
        rewrite_package(
            CHART,
            self.output,
            {
                "xl/workbook.xml": ET.tostring(workbook),
                "xl/_rels/workbook.xml.rels": ET.tostring(rels),
                "xl/chartsheets/chart.xml": b"<chartsheet/>",
            },
        )
        with zipfile.ZipFile(self.output) as archive:
            self.assertEqual(run.worksheet_parts(archive), {"Data": SHEET})

    def test_scan_counts_features_and_only_formula_caches(self):
        features, caches = run.scan(VALIDATION)
        self.assertEqual(
            (features["cf"], features["dv"], features["x14_ext"]), (0, 1, 0)
        )
        self.assertEqual(caches, {("P1 Surface", "F2"): (None, "60")})
        self.assertEqual(run.scan(CHART)[0]["cf"], 3)

    def test_scan_detects_changed_x14_reference_even_when_counts_are_unchanged(self):
        features, caches = run.scan(X14_PROBE)
        self.assertEqual(
            (features["x14_ext"], features["cf"], features["dv"]), (8, 0, 0)
        )
        self.assertEqual(caches, {})
        with zipfile.ZipFile(X14_PROBE) as archive:
            worksheet = ET.fromstring(archive.read(SHEET))
        worksheet.find(f".//{{{run.XM}}}f").text = "Inputs!$A$1:$A$4"
        rewrite_package(X14_PROBE, self.output, {SHEET: ET.tostring(worksheet)})
        changed, _ = run.scan(self.output)
        self.assertEqual((changed["x14_ext"], changed["cf"], changed["dv"]), (8, 0, 0))
        self.assertNotEqual(changed["x14_digest"], features["x14_digest"])

    def test_untouched_copy_reports_no_preservation_losses(self):
        shutil.copyfile(CHART, self.output)
        row = run.summarize(CHART, self.output, "Data", {})
        self.assertEqual(row["parts_missing"], [])
        self.assertEqual(row["families_lost"], {})
        self.assertEqual(row["sheet_features_changed"], [])
        self.assertEqual(row["cached_value_regressions"], 0)

    def test_removed_chart_is_reported_as_part_and_family_loss(self):
        with zipfile.ZipFile(CHART) as archive:
            drawing = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
        drawing.clear()
        rewrite_package(
            CHART,
            self.output,
            {"xl/drawings/drawing1.xml": ET.tostring(drawing)},
            omitted=("xl/charts/chart1.xml",),
        )
        row = run.summarize(CHART, self.output, "Data", {})
        self.assertEqual(row["parts_missing"], ["xl/charts/chart1.xml"])
        self.assertEqual(row["families_lost"], {"chart": [1, 0]})
        self.assertEqual(row["sheet_features_changed"], [])

    def test_formula_cache_error_transition_is_reported_as_regression(self):
        with zipfile.ZipFile(VALIDATION) as archive:
            worksheet = ET.fromstring(archive.read(SHEET))
        cell = worksheet.find(f".//{{{run.MAIN}}}c[@r='F2']")
        cell.set("t", "e")
        cell.find(f"{{{run.MAIN}}}v").text = "#REF!"
        rewrite_package(VALIDATION, self.output, {SHEET: ET.tostring(worksheet)})
        row = run.summarize(VALIDATION, self.output, "P1 Surface", {})
        self.assertEqual(row["formula_cells_checked"], 1)
        self.assertEqual(row["cached_value_regressions"], 1)
        self.assertEqual(
            row["cached_value_regression_samples"],
            [
                {
                    "sheet": "P1 Surface",
                    "cell": "F2",
                    "before": "60",
                    "after": "#REF!",
                }
            ],
        )
        self.assertEqual(row["parts_missing"], [])

    def test_missing_formula_cache_is_not_counted_as_error_regression(self):
        with zipfile.ZipFile(VALIDATION) as archive:
            worksheet = ET.fromstring(archive.read(SHEET))
        cell = worksheet.find(f".//{{{run.MAIN}}}c[@r='F2']")
        cell.remove(cell.find(f"{{{run.MAIN}}}v"))
        rewrite_package(VALIDATION, self.output, {SHEET: ET.tostring(worksheet)})
        row = run.summarize(VALIDATION, self.output, "P1 Surface", {})
        self.assertEqual(row["formula_cells_checked"], 1)
        self.assertEqual(row["cached_value_regressions"], 0)
        self.assertEqual(row["cached_value_regression_samples"], [])

    def test_failed_distinguishes_execution_failures_from_feature_observations(self):
        success = {
            "cached_value": 2,
            "recalc": {"status": "errors_found", "total_errors": 1},
            "parts_missing": ["xl/charts/chart1.xml"],
            "families_lost": {"chart": [1, 0]},
        }
        self.assertFalse(run.failed(success))
        for label, changes in (
            ("workflow exception", {"error": "failed to save"}),
            ("recalc exception", {"recalc": {"error": "timeout"}}),
            ("absent cache", {"cached_value": None}),
            ("wrong calculated value", {"cached_value": 3}),
        ):
            with self.subTest(label=label):
                self.assertTrue(run.failed({**success, **changes}))

    def test_community_native_workflow_rejects_before_editing(self):
        with self.assertRaisesRegex(
            RuntimeError, r"Commercial 2\.2\.0\+.*persists formula caches"
        ):
            run.wolfxl_workflow(CHART, self.output)
        self.assertFalse(self.output.exists())


if __name__ == "__main__":
    unittest.main()
