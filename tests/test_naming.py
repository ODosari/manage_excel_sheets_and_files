from excelmgr.core.naming import dedupe, sanitize_sheet_name


def test_sanitize_basic():
    assert sanitize_sheet_name("  /Weird:Name*  ") == "Weird_Name_"
    assert sanitize_sheet_name("") == "Empty"


def test_sanitize_illegal_characters():
    result = sanitize_sheet_name("Name:?[]*/\\")
    assert result == "Name_______"

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


def test_dedupe_handles_multiple_full_length_names():
    s: set[str] = set()
    base = "y" * 31
    names = [dedupe(base, s) for _ in range(3)]
    assert names[0] == base
    assert names[1].endswith("_2")
    assert names[2].endswith("_3")
    assert all(len(n) <= 31 for n in names)


def test_dedupe_without_max_length_allows_long_names():
    s: set[str] = set()
    base = "abc" * 20
    first = dedupe(base, s, max_length=None)
    second = dedupe(base, s, max_length=None)
    assert first == base
    assert second == f"{base}_2"
