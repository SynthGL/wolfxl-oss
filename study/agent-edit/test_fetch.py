"""Offline checks for corpus integrity and the explicit seven-file fallback."""

import contextlib
import hashlib
import io
import tempfile
import unittest
import urllib.error
import zipfile
from pathlib import Path
from unittest.mock import patch

import fetch


class FetchTests(unittest.TestCase):
    def test_cached_workbook_mismatch_stops_without_downloading(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.xlsx"
            path.write_bytes(b"corrupted cached workbook")
            with (
                patch.object(fetch, "CONTOSO", path),
                patch.object(fetch, "download") as download,
            ):
                with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
                    fetch.fetch_contoso()
                download.assert_not_called()

    def test_downloaded_workbook_mismatch_is_not_skipped_or_cached(self):
        archive = io.BytesIO()
        with zipfile.ZipFile(archive, "w") as zipped:
            zipped.writestr("sample.xlsx", b"different workbook")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.xlsx"
            with (
                patch.object(fetch, "CONTOSO", path),
                patch.object(fetch, "download", return_value=archive.getvalue()),
            ):
                with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
                    fetch.fetch_contoso()
                self.assertFalse(path.exists())

    def test_verified_workbook_extracted_without_archive_path_traversal(self):
        workbook = b"verified workbook bytes"
        archive = io.BytesIO()
        with zipfile.ZipFile(archive, "w") as zipped:
            zipped.writestr("../outside.xlsx", workbook)
        with tempfile.TemporaryDirectory() as directory:
            cache = Path(directory) / "cache"
            path = cache / "sample.xlsx"
            with (
                patch.object(fetch, "CONTOSO", path),
                patch.object(fetch, "CACHE", cache),
                patch.object(
                    fetch, "CONTOSO_SHA256", hashlib.sha256(workbook).hexdigest()
                ),
                patch.object(fetch, "download", return_value=archive.getvalue()),
            ):
                self.assertEqual(fetch.fetch_contoso().read_bytes(), workbook)
                self.assertFalse((Path(directory) / "outside.xlsx").exists())

    def test_unavailable_download_reports_seven_file_fallback(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.xlsx"
            stderr = io.StringIO()
            with (
                patch.object(fetch, "CONTOSO", path),
                patch.object(
                    fetch, "download", side_effect=urllib.error.URLError("unavailable")
                ),
                contextlib.redirect_stderr(stderr),
            ):
                self.assertIsNone(fetch.fetch_contoso())
            self.assertIn("running on 7 files", stderr.getvalue())
            self.assertFalse(path.exists())


if __name__ == "__main__":
    unittest.main()
