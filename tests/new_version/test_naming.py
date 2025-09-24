from src.core.naming import sanitize_sheet_name, dedupe

def test_sanitize_basic():
    assert sanitize_sheet_name("  /Weird:Name*  ") == "Weird_Name_"
    assert sanitize_sheet_name("") == "Empty"

def test_dedupe():
    s = set()
    a = dedupe("Data", s)
    b = dedupe("Data", s)
    assert a == "Data"
    assert b.startswith("Data_")


def test_dedupe_truncates_when_max_length():
    s: set[str] = set()
    base = "x" * 31
    first = dedupe(base, s)
    second = dedupe(base, s)
    assert first == base
    assert second.endswith("_2")
    assert len(second) <= 31
