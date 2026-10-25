"""Fresh-interpreter coverage for optional upstream compatibility bindings."""

from __future__ import annotations

import subprocess
import sys
import textwrap

import pytest


def _run(source: str) -> None:
    result = subprocess.run(
        [sys.executable, "-c", textwrap.dedent(source)],
        capture_output=True,
        text=True,
        timeout=60,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_import_does_not_load_openpyxl_and_first_shim_borrows_attributes() -> None:
    pytest.importorskip("openpyxl")
    _run("""
        import sys
        import wolfxl
        assert not any(n == 'openpyxl' or n.startswith('openpyxl.') for n in sys.modules)
        from wolfxl._compat import _make_serialisable
        shim = _make_serialisable('PhoneticProperties', module_name='wolfxl.cell.text')
        from openpyxl.cell.text import PhoneticProperties
        assert shim.tagname == PhoneticProperties.tagname == 'phoneticPr'
        assert shim.__attrs__ == PhoneticProperties.__attrs__
        obj = shim(fontId=1)
        assert obj.to_tree().tag == 'phoneticPr'
        assert obj.to_tree().attrib['fontId'] == '1'
    """)


@pytest.mark.parametrize("upstream_first", [False, True])
def test_exception_identity_survives_either_import_order(upstream_first: bool) -> None:
    pytest.importorskip("openpyxl")
    _run(f"""
        if {upstream_first!r}:
            import openpyxl
        import wolfxl
        from wolfxl.utils.exceptions import CellCoordinatesException as before
        instance = before('bad coordinate')
        from openpyxl.utils.exceptions import CellCoordinatesException as upstream
        from wolfxl.utils.exceptions import CellCoordinatesException as after
        assert before is after
        assert isinstance(instance, (ValueError, before))
        assert isinstance(instance, upstream)
        try:
            raise instance
        except upstream:
            pass
    """)


@pytest.mark.parametrize("upstream_first", [False, True])
def test_alias_preserves_real_snapshot_and_pandas_version_probe(upstream_first: bool) -> None:
    pytest.importorskip("openpyxl")
    pytest.importorskip("pandas")
    _run(f"""
        if {upstream_first!r}:
            import openpyxl
        import wolfxl
        wolfxl.install_as_openpyxl()
        import openpyxl
        from wolfxl import _compat
        from wolfxl.openpyxl_compat import OPENPYXL_API_VERSION
        from pandas.compat import _optional
        assert openpyxl is wolfxl
        assert _optional.get_version(openpyxl) == OPENPYXL_API_VERSION
        real = _compat._REAL_OPENPYXL
        assert real.__name__ == 'openpyxl'
        assert real is not wolfxl
        upstream = _compat._resolve_openpyxl_class('wolfxl.cell.text', 'PhoneticProperties')
        assert upstream is real.cell.text.PhoneticProperties
        shim = _compat._make_serialisable('PhoneticProperties', module_name='wolfxl.cell.text')
        assert shim.tagname == upstream.tagname
        assert shim.__attrs__ == upstream.__attrs__
    """)


def test_partial_upstream_import_binds_after_module_execution() -> None:
    pytest.importorskip("openpyxl")
    _run("""
        import importlib.abc
        import importlib.machinery
        import sys
        class Loader:
            def __init__(self, loader):
                self.loader = loader
            def create_module(self, spec):
                return self.loader.create_module(spec)
            def exec_module(self, module):
                import wolfxl
                self.loader.exec_module(module)
        class Finder(importlib.abc.MetaPathFinder):
            def find_spec(self, fullname, path=None, target=None):
                if fullname == 'openpyxl':
                    spec = importlib.machinery.PathFinder.find_spec(fullname, path)
                    spec.loader = Loader(spec.loader)
                    return spec
        sys.meta_path.insert(0, Finder())
        import openpyxl
        from wolfxl.utils.exceptions import CellCoordinatesException
        assert issubclass(CellCoordinatesException, openpyxl.utils.exceptions.CellCoordinatesException)
        import wolfxl
        wolfxl.install_as_openpyxl()
        from wolfxl._compat import _REAL_OPENPYXL
        assert _REAL_OPENPYXL.cell.text.PhoneticProperties.tagname == 'phoneticPr'
    """)


def test_no_openpyxl_installation_keeps_standalone_shims() -> None:
    _run("""
        import importlib.abc
        import sys
        class MissingOpenpyxl(importlib.abc.MetaPathFinder):
            def find_spec(self, fullname, path=None, target=None):
                if fullname == 'openpyxl' or fullname.startswith('openpyxl.'):
                    raise ModuleNotFoundError(fullname)
        sys.meta_path.insert(0, MissingOpenpyxl())
        import wolfxl
        from wolfxl._compat import _make_serialisable, _REAL_OPENPYXL
        from wolfxl.utils.exceptions import CellCoordinatesException
        assert _REAL_OPENPYXL is None
        assert issubclass(CellCoordinatesException, ValueError)
        shim = _make_serialisable('AbsentModel')
        assert shim(value=3).value == 3
        wolfxl.install_as_openpyxl()
        import openpyxl
        assert openpyxl is wolfxl
    """)


def test_lazy_drawing_exports_keep_upstream_identity() -> None:
    pytest.importorskip("openpyxl")
    _run("""
        import wolfxl
        import wolfxl.drawing as drawing
        assert set(drawing.__all__).issubset(dir(drawing))
        from wolfxl.drawing import AnchorMarker, Drawing, Image
        from wolfxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
        from wolfxl.chartsheet.chartsheet import SpreadsheetDrawing as chartsheet_drawing
        from wolfxl.drawing.image import Image as original_image
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker as upstream
        assert AnchorMarker is upstream
        assert drawing.AnchorMarker is AnchorMarker
        assert Drawing is SpreadsheetDrawing
        assert chartsheet_drawing is SpreadsheetDrawing
        assert Image is original_image
        assert isinstance(AnchorMarker(), upstream)
    """)
