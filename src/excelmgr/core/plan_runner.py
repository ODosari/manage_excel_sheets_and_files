from __future__ import annotations

import json
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from excelmgr.adapters.cloud import LocalCloudObjectWriter
from excelmgr.adapters.database import SQLiteTableWriter
from excelmgr.core.combine import combine
from excelmgr.core.delete_cols import delete_columns
from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.models import (
    CloudDestination,
    CombinePlan,
    DatabaseDestination,
    DeleteSpec,
    Destination,
    PreviewPlan,
    SheetSpec,
    SplitPlan,
)
from excelmgr.core.password_maps import load_password_map
from excelmgr.core.preview import preview
from excelmgr.core.progress import ProgressHook
from excelmgr.core.split import split
from excelmgr.util.text import read_text
from excelmgr.ports.readers import WorkbookReader
from excelmgr.ports.writers import CloudObjectWriter, TableWriter, WorkbookWriter


PlanType = Literal["combine", "split", "delete", "preview"]


@dataclass
class PlanOperation:
    type: PlanType
    plan: CombinePlan | SplitPlan | DeleteSpec | PreviewPlan
    name: str | None = None


def _ensure_sequence(value, field: str) -> Sequence:
    if not isinstance(value, Sequence) or isinstance(value, (str, bytes)):
        raise ExcelMgrError(f"Plan field '{field}' must be a sequence.")
    return value


def _resolve_path(base_dir: Path, value: object) -> str:
    path = Path(str(value)).expanduser()
    if path.is_absolute():
        return str(path.resolve())
    return str((base_dir / path).resolve())


def _parse_sheet_specs(raw: object) -> Sequence[SheetSpec] | Literal["all"]:
    if raw in (None, "all"):
        return "all"
    items = _ensure_sequence(raw, "include_sheets")
    specs: list[SheetSpec] = []
    for item in items:
        if isinstance(item, Mapping):
            if "index" in item:
                specs.append(SheetSpec(int(item["index"])))
            elif "name" in item:
                specs.append(SheetSpec(str(item["name"])))
            continue
        if isinstance(item, int):
            specs.append(SheetSpec(item))
            continue
        if isinstance(item, str) and item.isdigit():
            specs.append(SheetSpec(int(item)))
        else:
            specs.append(SheetSpec(str(item)))
    return specs or "all"


def _parse_sheet_selector(value: object) -> SheetSpec | Literal["active"]:
    if value in (None, "active"):
        return "active"
    if isinstance(value, Mapping) and "index" in value:
        return SheetSpec(int(value["index"]))
    if isinstance(value, int):
        return SheetSpec(value)
    if isinstance(value, str):
        cleaned = value.strip()
        if not cleaned:
            raise ExcelMgrError("Plan sheet selector cannot be empty.")
        if cleaned.lower().startswith("index:"):
            _, _, rest = cleaned.partition(":")
            if not rest or not rest.strip().isdigit():
                raise ExcelMgrError("Sheet selector index must be numeric.")
            return SheetSpec(int(rest.strip()))
        if cleaned.isdigit():
            return SheetSpec(int(cleaned))
        if cleaned == "active":
            return "active"
        return SheetSpec(cleaned)
    raise ExcelMgrError("Unsupported sheet selector format in plan file.")


def _parse_destination(raw: object, base_dir: Path) -> Destination | None:
    if raw is None:
        return None
    if not isinstance(raw, Mapping):
        raise ExcelMgrError("Destination entry must be a mapping.")
    kind = raw.get("kind")
    if kind == "database":
        uri = raw.get("uri")
        table = raw.get("table")
        if not uri or not table:
            raise ExcelMgrError("Database destination requires 'uri' and 'table'.")
        mode = str(raw.get("mode", "replace"))
        options = raw.get("options")
        if options is not None and not isinstance(options, Mapping):
            raise ExcelMgrError("Database destination 'options' must be a mapping if provided.")
        resolved_uri = _resolve_path(base_dir, uri)
        return DatabaseDestination(
            uri=resolved_uri,
            table=str(table),
            mode="append" if mode == "append" else "replace",
            options=dict(options or {}),
        )
    if kind == "cloud":
        root = raw.get("root")
        key = raw.get("key")
        if not root or not key:
            raise ExcelMgrError("Cloud destination requires 'root' and 'key'.")
        fmt = raw.get("format")
        options = raw.get("options")
        if options is not None and not isinstance(options, Mapping):
            raise ExcelMgrError("Cloud destination 'options' must be a mapping if provided.")
        resolved_root = _resolve_path(base_dir, root)
        return CloudDestination(
            root=resolved_root,
            key=str(key),
            format=str(fmt) if fmt else "parquet",
            options=dict(options or {}),
        )
    raise ExcelMgrError("Destination kind must be 'database' or 'cloud'.")


def _load_password_map(raw: object, base_dir: Path) -> dict[str, str] | None:
    if raw is None:
        return None
    if isinstance(raw, Mapping):
        return {str(k): str(v) for k, v in raw.items()}
    return load_password_map(str(raw), base_dir=base_dir)


def _build_combine_plan(entry: Mapping, base_dir: Path) -> CombinePlan:
    inputs = entry.get("inputs")
    if inputs is None:
        raise ExcelMgrError("Combine operation requires an 'inputs' list.")
    paths = [_resolve_path(base_dir, item) for item in _ensure_sequence(inputs, "inputs")]
    include = _parse_sheet_specs(entry.get("include_sheets"))
    password_map = _load_password_map(entry.get("password_map"), base_dir)
    destination = _parse_destination(entry.get("destination"), base_dir)
    mode = entry.get("mode", "one_sheet")
    if mode not in {"one_sheet", "multi_sheets"}:
        raise ExcelMgrError("Combine 'mode' must be 'one_sheet' or 'multi_sheets'.")
    output_path = entry.get("output_path", "combined.xlsx")
    resolved_output = _resolve_path(base_dir, output_path)
    sheet_name = entry.get("sheet_name") or entry.get("output_sheet_name")
    return CombinePlan(
        inputs=paths,
        glob=entry.get("glob"),
        recursive=bool(entry.get("recursive", False)),
        mode=mode,  # type: ignore[arg-type]
        include_sheets=include,
        output_path=resolved_output,
        output_sheet_name=sheet_name or "Data",
        add_source_column=bool(entry.get("add_source_column", False)),
        password=entry.get("password"),
        password_map=password_map,
        output_format=str(entry.get("output_format", "xlsx")),
        csv_add_bom=bool(entry.get("csv_add_bom", False)),
        dry_run=bool(entry.get("dry_run", False)),
        destination=destination,
    )


def _build_split_plan(entry: Mapping, base_dir: Path) -> SplitPlan:
    input_file = entry.get("input") or entry.get("input_file")
    if input_file is None:
        raise ExcelMgrError("Split operation requires an 'input' path.")
    resolved_input = _resolve_path(base_dir, input_file)
    password_map = _load_password_map(entry.get("password_map"), base_dir)
    destination = _parse_destination(entry.get("destination"), base_dir)
    sheet = _parse_sheet_selector(entry.get("sheet"))
    by_column = entry.get("by") or entry.get("by_column") or "Category"
    to = entry.get("to", "files")
    output_dir = entry.get("output_dir") or entry.get("out") or "out"
    resolved_out = _resolve_path(base_dir, output_dir)
    out_file = entry.get("output_filename") or entry.get("out_file")
    sheet_name = entry.get("sheet_name") or entry.get("output_sheet_name")
    return SplitPlan(
        input_file=resolved_input,
        sheet=sheet,
        by_column=by_column,
        to=to,  # type: ignore[arg-type]
        include_nan=bool(entry.get("include_nan", False)),
        output_dir=resolved_out,
        output_filename=out_file,
        output_sheet_name=sheet_name or "Data",
        password=entry.get("password"),
        password_map=password_map,
        output_format=str(entry.get("output_format", "xlsx")),
        csv_add_bom=bool(entry.get("csv_add_bom", False)),
        dry_run=bool(entry.get("dry_run", False)),
        destination=destination,
    )


def _build_delete_plan(entry: Mapping, base_dir: Path) -> DeleteSpec:
    path = entry.get("path")
    if path is None:
        raise ExcelMgrError("Delete operation requires a 'path'.")
    resolved_path = _resolve_path(base_dir, path)
    targets = entry.get("targets")
    if targets is None:
        raise ExcelMgrError("Delete operation requires 'targets'.")
    target_seq = _ensure_sequence(targets, "targets")
    match_kind = entry.get("match", entry.get("match_kind", "names"))
    if match_kind not in {"names", "index"}:
        raise ExcelMgrError("Delete operation 'match' must be 'names' or 'index'.")
    if match_kind == "index":
        parsed_targets = [int(item) for item in target_seq]
    else:
        parsed_targets = [str(item).strip() for item in target_seq]
    password_map = _load_password_map(entry.get("password_map"), base_dir)
    sheet_selector = entry.get("sheet")
    selector_value: str | int | None
    if sheet_selector is None:
        selector_value = None
    elif isinstance(sheet_selector, int):
        selector_value = sheet_selector
    elif isinstance(sheet_selector, str) and sheet_selector.lower().startswith("index:"):
        _, _, rest = sheet_selector.partition(":")
        selector_value = int(rest.strip())
    elif isinstance(sheet_selector, str) and sheet_selector.isdigit():
        selector_value = int(sheet_selector)
    else:
        selector_value = sheet_selector
    return DeleteSpec(
        path=resolved_path,
        targets=parsed_targets,
        match_kind=match_kind,  # type: ignore[arg-type]
        strategy=str(entry.get("strategy", "exact")),
        all_sheets=bool(entry.get("all_sheets", False)),
        sheet_selector=selector_value,
        inplace=bool(entry.get("inplace", False)),
        on_missing=str(entry.get("on_missing", "ignore")),
        dry_run=bool(entry.get("dry_run", False)),
        glob=entry.get("glob"),
        recursive=bool(entry.get("recursive", False)),
        password=entry.get("password"),
        password_map=password_map,
    )


def _build_preview_plan(entry: Mapping, base_dir: Path) -> PreviewPlan:
    path = entry.get("path")
    if path is None:
        raise ExcelMgrError("Preview operation requires a 'path'.")
    resolved_path = _resolve_path(base_dir, path)
    password_map = _load_password_map(entry.get("password_map"), base_dir)
    limit = entry.get("limit")
    limit_value = int(limit) if isinstance(limit, (int, str)) and str(limit).isdigit() else None
    return PreviewPlan(
        path=resolved_path,
        password=entry.get("password"),
        password_map=password_map,
        limit=limit_value,
    )


def load_plan_file(path: str) -> list[PlanOperation]:
    plan_path = Path(path).expanduser().resolve()
    if not plan_path.exists():
        raise ExcelMgrError(f"Plan file not found: {path}")
    base_dir = plan_path.parent

    data: object
    if plan_path.suffix.lower() in {".yaml", ".yml"}:
        try:
            import yaml
        except ImportError as exc:  # pragma: no cover - optional dependency
            raise ExcelMgrError("YAML plan support requires the 'PyYAML' package.") from exc
        data = yaml.safe_load(read_text(plan_path))
    else:
        data = json.loads(read_text(plan_path))

    if data is None:
        return []

    if isinstance(data, Mapping):
        operations = data.get("operations")
        if operations is None:
            raise ExcelMgrError("Plan file must include an 'operations' list.")
    elif isinstance(data, Sequence):
        operations = data
    else:
        raise ExcelMgrError("Plan file must be a list of operations or contain an 'operations' list.")

    result: list[PlanOperation] = []
    for idx, raw in enumerate(operations):
        if not isinstance(raw, Mapping):
            raise ExcelMgrError(f"Operation #{idx + 1} must be a mapping.")
        op_type = raw.get("type")
        if op_type not in {"combine", "split", "delete", "preview"}:
            raise ExcelMgrError("Operation type must be one of combine, split, delete, preview.")
        options = raw.get("options", raw)
        if not isinstance(options, Mapping):
            raise ExcelMgrError(f"Operation #{idx + 1} options must be a mapping.")

        if op_type == "combine":
            plan = _build_combine_plan(options, base_dir)
        elif op_type == "split":
            plan = _build_split_plan(options, base_dir)
        elif op_type == "delete":
            plan = _build_delete_plan(options, base_dir)
        else:
            plan = _build_preview_plan(options, base_dir)

        result.append(PlanOperation(type=op_type, plan=plan, name=raw.get("name")))
    return result


def _make_database_writer(destination: Destination | None) -> TableWriter | None:
    if isinstance(destination, DatabaseDestination):
        return SQLiteTableWriter()
    return None


def _make_cloud_writer(destination: Destination | None) -> CloudObjectWriter | None:
    if isinstance(destination, CloudDestination):
        return LocalCloudObjectWriter(destination.root)
    return None


def execute_plan(
    operations: Iterable[PlanOperation],
    reader: WorkbookReader,
    writer: WorkbookWriter,
    *,
    progress_hooks: Iterable[ProgressHook] | None = None,
) -> list[dict[str, object]]:
    results: list[dict[str, object]] = []
    hooks = tuple(progress_hooks or ())
    for op in operations:
        destination = getattr(op.plan, "destination", None)
        db_writer = _make_database_writer(destination)
        cloud_writer = _make_cloud_writer(destination)
        if op.type == "combine":
            result = combine(
                op.plan,  # type: ignore[arg-type]
                reader,
                writer,
                database_writer=db_writer,
                cloud_writer=cloud_writer,
                progress_hooks=hooks,
            )
        elif op.type == "split":
            result = split(
                op.plan,  # type: ignore[arg-type]
                reader,
                writer,
                cloud_writer=cloud_writer,
                progress_hooks=hooks,
            )
        elif op.type == "delete":
            result = delete_columns(
                op.plan,  # type: ignore[arg-type]
                reader,
                writer,
                progress_hooks=hooks,
            )
        else:
            result = preview(op.plan, reader)  # type: ignore[arg-type]
        results.append({"type": op.type, "name": op.name, "result": result})
    return results
