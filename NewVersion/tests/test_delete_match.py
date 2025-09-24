from NewVersion.excelmgr.core.models import DeleteSpec
from NewVersion.excelmgr.core.delete_cols import _match_columns

def test_match_names_exact():
    cols = ["ID", "Name", "Notes", "CustomerID"]
    spec = DeleteSpec(path="x", targets=["Notes","CustomerID"], match_kind="names", strategy="exact", all_sheets=True, inplace=True, on_missing="ignore", dry_run=True)
    remove, missing = _match_columns(cols, spec)
    assert set(remove) == {"Notes","CustomerID"}
    assert not missing

def test_match_index():
    cols = ["A", "B", "C", "D"]
    spec = DeleteSpec(path="x", targets=[1,3], match_kind="index", strategy="exact", all_sheets=True, inplace=True, on_missing="ignore", dry_run=True)
    remove, missing = _match_columns(cols, spec)
    assert set(remove) == {"A","C"}
    assert not missing
