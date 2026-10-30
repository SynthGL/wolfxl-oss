from __future__ import annotations

import importlib
import inspect
import os
import pkgutil
import subprocess
import sys
import traceback
from dataclasses import FrozenInstanceError, is_dataclass
from importlib.metadata import version
from importlib.util import find_spec
from pathlib import Path
from tempfile import TemporaryDirectory

import wolfxl
from wolfxl import Workbook, load_workbook


def calculate_roundtrip(directory: Path) -> None:
    source = directory / "report.xlsx"
    original = Workbook()
    sheet = original.active
    sheet.title = "Summary"
    sheet["B2"] = 1000
    sheet["B3"] = 950
    sheet["B4"] = "=SUM(B2:B3)"
    sheet["B2"].number_format = "#,##0.00"
    sheet.merge_cells("D1:E1")
    sheet["D1"] = "Merged heading"
    original.save(source)
    original.close()

    for modify in (False, True):
        workbook = load_workbook(source, modify=modify)
        workbook["Summary"]["B2"] = 1500
        assert workbook.calculate()["Summary!B4"] == 2450
        output = directory / f"calculated-{modify}.xlsx"
        workbook.save(output)
        workbook.close()

        for data_only in (False, True):
            saved = load_workbook(output, data_only=data_only)
            sheet = saved["Summary"]
            assert sheet["B2"].value == 1500
            assert sheet["B3"].value == 950
            assert sheet["B4"].value == (2450 if data_only else "=SUM(B2:B3)")
            assert sheet["B2"].number_format == "#,##0.00"
            assert sheet["D1"].value == "Merged heading"
            assert sheet["E1"].value is None
            saved.close()


def import_all_installed_modules() -> int:
    """Import every installed wolfxl module, private modules included."""
    failures: dict[str, BaseException] = {}

    def record_traversal_error(name: str) -> None:
        error = sys.exc_info()[1]
        assert error is not None, name
        failures.setdefault(name, error)

    imported = []
    for module in pkgutil.walk_packages(
        wolfxl.__path__, wolfxl.__name__ + ".", onerror=record_traversal_error
    ):
        try:
            importlib.import_module(module.name)
        except Exception as error:
            failures.setdefault(module.name, error)
        else:
            imported.append(module.name)

    if failures:
        names = sorted(failures)
        details = "\n".join(
            f"{name}: {failures[name]!r}\n"
            + "".join(
                traceback.format_exception(
                    type(failures[name]), failures[name], failures[name].__traceback__
                )
            )
            for name in names
        )
        raise AssertionError(
            f"failed to import {len(names)} installed wolfxl module(s): {', '.join(names)}\n{details}"
        ) from failures[names[0]]
    assert imported, "pkgutil.walk_packages discovered no installed wolfxl modules"
    return len(imported)


def test_drawing_text_bullet_compatibility() -> None:
    from wolfxl.drawing.text import (
        AutoNumberedBullet,
        CharacterBullet,
        PictureBullet,
        _BulletStyle,
    )

    style_names = ("color", "size_percent", "size_points", "font")
    image = b"\x89PNG\r\n\x1a\n"
    cases = (
        (
            _BulletStyle,
            {"color": "#1f4e79", "size_percent": 150, "font": "Arial"},
            {"color": "1F4E79", "size_percent": 150.0, "size_points": None, "font": "Arial"},
        ),
        (
            CharacterBullet,
            {"character": "-", "color": "#1f4e79", "size_percent": 150, "font": "Arial"},
            {
                "character": "-",
                "color": "1F4E79",
                "size_percent": 150.0,
                "size_points": None,
                "font": "Arial",
            },
        ),
        (
            CharacterBullet,
            {},
            {
                "character": "\u2022",
                "color": None,
                "size_percent": None,
                "size_points": None,
                "font": None,
            },
        ),
        (
            AutoNumberedBullet,
            {"scheme": "romanUcPeriod", "start_at": 3, "color": "C00000", "size_points": 12.5},
            {
                "scheme": "romanUcPeriod",
                "start_at": 3,
                "color": "C00000",
                "size_percent": None,
                "size_points": 12.5,
                "font": None,
            },
        ),
        (
            AutoNumberedBullet,
            {},
            {
                "scheme": "arabicPeriod",
                "start_at": 1,
                "color": None,
                "size_percent": None,
                "size_points": None,
                "font": None,
            },
        ),
        (
            PictureBullet,
            {"image": bytearray(image), "image_format": "png", "size_points": 10},
            {
                "image": image,
                "image_format": "png",
                "color": None,
                "size_percent": None,
                "size_points": 10.0,
                "font": None,
            },
        ),
    )
    instances = []
    for cls, kwargs, fields in cases:
        instance = cls(**kwargs)
        assert is_dataclass(instance), cls.__name__
        actual = {name: getattr(instance, name) for name in fields}
        assert actual == fields, (cls.__name__, actual)
        assert instance == cls(**kwargs), cls.__name__
        assert hash(instance) == hash(cls(**kwargs)), cls.__name__
        for name in fields:
            try:
                setattr(instance, name, None)
            except FrozenInstanceError:
                pass
            else:
                raise AssertionError(f"{cls.__name__}.{name} accepted assignment")
            try:
                delattr(instance, name)
            except FrozenInstanceError:
                pass
            else:
                raise AssertionError(f"{cls.__name__}.{name} accepted deletion")
        assert {name: getattr(instance, name) for name in fields} == fields, cls.__name__
        instances.append(instance)

    style, character, default_character, numbered, default_numbered, picture = instances
    assert type(picture.image) is bytes
    assert style._style_payload("character") == {
        "kind": "character",
        "color": "1F4E79",
        "size_percent": 150_000,
        "size_points": None,
        "font": "Arial",
    }
    assert character.to_payload() == {
        "kind": "character",
        "color": "1F4E79",
        "size_percent": 150_000,
        "size_points": None,
        "font": "Arial",
        "character": "-",
    }
    assert default_character.to_payload() == {
        "kind": "character",
        "color": None,
        "size_percent": None,
        "size_points": None,
        "font": None,
        "character": "\u2022",
    }
    assert numbered.to_payload() == {
        "kind": "auto_number",
        "color": "C00000",
        "size_percent": None,
        "size_points": 1_250,
        "font": None,
        "scheme": "romanUcPeriod",
        "start_at": 3,
    }
    assert default_numbered.to_payload() == {
        "kind": "auto_number",
        "color": None,
        "size_percent": None,
        "size_points": None,
        "font": None,
        "scheme": "arabicPeriod",
        "start_at": 1,
    }
    assert AutoNumberedBullet(start_at=None).to_payload()["start_at"] is None
    assert picture.to_payload() == {
        "kind": "picture",
        "color": None,
        "size_percent": None,
        "size_points": 1_000,
        "font": None,
        "image": image,
        "image_format": "png",
    }
    for cls, kwargs in (
        (CharacterBullet, {"size_percent": 150, "size_points": 12}),
        (PictureBullet, {}),
    ):
        try:
            cls(**kwargs)
        except ValueError:
            pass
        else:
            raise AssertionError(f"{cls.__name__}({kwargs!r}) was accepted")

    if sys.version_info >= (3, 10):
        keyword_only = inspect.Parameter.KEYWORD_ONLY
        positional = inspect.Parameter.POSITIONAL_OR_KEYWORD
        style_parameters = [(name, keyword_only, None) for name in style_names]
        for cls, prefix in (
            (_BulletStyle, ()),
            (CharacterBullet, (("character", "\u2022"),)),
            (AutoNumberedBullet, (("scheme", "arabicPeriod"), ("start_at", 1))),
            (PictureBullet, (("image", b""), ("image_format", "png"))),
        ):
            actual = [
                (parameter.name, parameter.kind, parameter.default)
                for parameter in inspect.signature(cls).parameters.values()
            ]
            expected = [(name, positional, default) for name, default in prefix]
            assert actual == expected + style_parameters, (cls.__name__, actual)


def test_clean_installed_python39_wheel(tmp_path: Path) -> None:
    """Explicit remote qualification node; excluded from default test discovery."""
    project = Path(__file__).resolve().parents[1]
    dist = tmp_path / "dist"
    # The pinned gate image's interpreter directory is read-only.
    environment_variables = {
        **os.environ,
        "UV_PYTHON_INSTALL_DIR": str(tmp_path / "interpreters"),
    }
    subprocess.run(
        ["uv", "build", "--wheel", "--python", "3.9", "--out-dir", str(dist)],
        cwd=project,
        env=environment_variables,
        check=True,
    )
    wheels = list(dist.glob("wolfxl-2.0.31-cp39-cp39-*.whl"))
    assert len(wheels) == 1, wheels
    environment = tmp_path / "clean-python39"
    subprocess.run(
        ["uv", "venv", "--python", "3.9", str(environment)],
        env=environment_variables,
        check=True,
    )
    python = environment / "bin" / "python"
    subprocess.run(["uv", "pip", "install", "--python", str(python), str(wheels[0])], check=True)
    subprocess.run(
        [str(python), "-I", str(Path(__file__).resolve())],
        cwd=tmp_path,
        check=True,
    )


def main() -> None:
    assert version("wolfxl") == "2.0.31"
    assert wolfxl.__version__ == "2.0.31"
    for module_name in (
        "wolfxl.operations",
        "wolfxl.render",
        "wolfxl.integrations",
        "wolfxl._conversion",
        "wolfxl._power_query",
        "wolfxl._vba_runtime",
    ):
        assert find_spec(module_name) is None, module_name

    module_count = import_all_installed_modules()
    test_drawing_text_bullet_compatibility()
    print(
        f"community artifact smoke: imported {module_count} installed wolfxl modules;"
        " bullet keyword contract: ok"
    )
    with TemporaryDirectory() as directory:
        calculate_roundtrip(Path(directory))
    print("community artifact smoke: calculate=2450; normal/modify caches and formulas: ok")


if __name__ == "__main__":
    main()
