"""Runtime backend metadata smoke tests."""

from __future__ import annotations


def test_build_info_names_native_xlsx_and_binary_compat() -> None:
    from wolfxl import _rust

    info = _rust.build_info()

    assert info["package"] == "wolfxl"
    assert "native-xlsx" in info["enabled_backends"]
    assert "native-xlsb" in info["enabled_backends"]
    assert "native-xls" in info["enabled_backends"]
    assert "native-ods" not in info["enabled_backends"]
    assert not any(name.startswith("calamine") for name in info["enabled_backends"])
    assert not any(name.startswith("calamine") for name in info["backend_versions"])
