from pathlib import Path

from excelmgr.core.passwords import resolve_password


def test_resolve_password_prefers_exact_path(tmp_path: Path) -> None:
    file_a = tmp_path / "a.xlsx"
    mapping = {
        str(file_a): "alpha",
        file_a.name: "beta",
        "other.xlsx": "gamma",
    }

    assert resolve_password(str(file_a), None, mapping) == "alpha"
    assert resolve_password(str(tmp_path / "b.xlsx"), "fallback", mapping) == "fallback"
    nested = tmp_path / "nested" / "a.xlsx"
    nested.parent.mkdir()
    nested.touch()
    assert resolve_password(str(nested), None, mapping) == "beta"
