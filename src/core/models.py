from dataclasses import dataclass
from typing import Literal, Sequence, Optional, Mapping

ModeCombine = Literal["one_sheet", "multi_sheets"]
ModeSplitTo = Literal["sheets", "files"]
MatchKind = Literal["names", "index"]
NameMatchStrategy = Literal["exact", "ci", "contains", "startswith", "endswith", "regex"]

@dataclass(frozen=True)
class SheetSpec:
    name_or_index: str | int  # flexible addressing

@dataclass(frozen=True)
class CombinePlan:
    inputs: Sequence[str]                  # files or dirs
    glob: Optional[str] = None
    recursive: bool = False
    mode: ModeCombine = "one_sheet"
    include_sheets: Sequence[SheetSpec] | Literal["all"] = "all"
    output_path: str = "combined.xlsx"
    add_source_column: bool = False
    password: Optional[str] = None
    password_map: Optional[Mapping[str, str]] = None
    output_format: Literal["xlsx", "csv", "parquet"] = "xlsx"
    dry_run: bool = False

@dataclass(frozen=True)
class SplitPlan:
    input_file: str
    sheet: SheetSpec | Literal["active"] = "active"
    by_column: str | int = "Category"
    to: ModeSplitTo = "files"
    include_nan: bool = False
    output_dir: str = "out"
    password: Optional[str] = None
    password_map: Optional[Mapping[str, str]] = None
    output_format: Literal["xlsx", "csv", "parquet"] = "xlsx"
    dry_run: bool = False

@dataclass(frozen=True)
class DeleteSpec:
    path: str                       # file or dir
    targets: Sequence[str] | Sequence[int]
    match_kind: MatchKind = "names"
    strategy: NameMatchStrategy = "exact"
    all_sheets: bool = False
    sheet_selector: Optional[str | int] = None
    inplace: bool = False
    on_missing: Literal["ignore", "error"] = "ignore"
    dry_run: bool = False
    glob: Optional[str] = None
    recursive: bool = False
    password: Optional[str] = None
    password_map: Optional[Mapping[str, str]] = None
