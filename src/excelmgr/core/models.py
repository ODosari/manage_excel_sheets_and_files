from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from typing import Literal

ModeCombine = Literal["one_sheet", "multi_sheets"]
ModeSplitTo = Literal["sheets", "files"]
MatchKind = Literal["names", "index"]
NameMatchStrategy = Literal["exact", "ci", "contains", "startswith", "endswith", "regex"]

@dataclass(frozen=True)
class SheetSpec:
    name_or_index: str | int  # flexible addressing

@dataclass(frozen=True)
class DatabaseDestination:
    uri: str
    table: str
    mode: Literal["replace", "append"] = "replace"
    options: Mapping[str, object] = field(default_factory=dict)


@dataclass(frozen=True)
class CloudDestination:
    root: str
    key: str
    format: Literal["csv", "parquet", "xlsx"] = "parquet"
    options: Mapping[str, object] = field(default_factory=dict)


Destination = DatabaseDestination | CloudDestination


@dataclass(frozen=True)
class CombinePlan:
    inputs: Sequence[str]                  # files or dirs
    glob: str | None = None
    recursive: bool = False
    mode: ModeCombine = "one_sheet"
    include_sheets: Sequence[SheetSpec] | Literal["all"] = "all"
    output_path: str = "combined.xlsx"
    add_source_column: bool = False
    password: str | None = None
    password_map: Mapping[str, str] | None = None
    output_format: Literal["xlsx", "csv", "parquet"] = "xlsx"
    dry_run: bool = False
    destination: Destination | None = None


@dataclass(frozen=True)
class SplitPlan:
    input_file: str
    sheet: SheetSpec | Literal["active"] = "active"
    by_column: str | int = "Category"
    to: ModeSplitTo = "files"
    include_nan: bool = False
    output_dir: str = "out"
    password: str | None = None
    password_map: Mapping[str, str] | None = None
    output_format: Literal["xlsx", "csv", "parquet"] = "xlsx"
    dry_run: bool = False
    destination: Destination | None = None


@dataclass(frozen=True)
class DeleteSpec:
    path: str                       # file or dir
    targets: Sequence[str] | Sequence[int]
    match_kind: MatchKind = "names"
    strategy: NameMatchStrategy = "exact"
    all_sheets: bool = False
    sheet_selector: str | int | None = None
    inplace: bool = False
    on_missing: Literal["ignore", "error"] = "ignore"
    dry_run: bool = False
    glob: str | None = None
    recursive: bool = False
    password: str | None = None
    password_map: Mapping[str, str] | None = None


@dataclass(frozen=True)
class PreviewPlan:
    path: str
    password: str | None = None
    password_map: Mapping[str, str] | None = None
    limit: int | None = None

