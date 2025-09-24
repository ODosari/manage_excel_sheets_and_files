from NewVersion.excelmgr.core.naming import sanitize_sheet_name, dedupe

def test_sanitize_basic():
    assert sanitize_sheet_name("  /Weird:Name*  ") == "Weird_Name_"
    assert sanitize_sheet_name("") == "Empty"

def test_dedupe():
    s = set()
    a = dedupe("Data", s)
    b = dedupe("Data", s)
    assert a == "Data"
    assert b.startswith("Data_")
