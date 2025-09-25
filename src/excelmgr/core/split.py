from pathlib import Path

import pandas as pd

from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.models import SplitPlan
from excelmgr.core.naming import dedupe, sanitize_sheet_name
from excelmgr.core.passwords import resolve_password
from excelmgr.core.sinks import csv_sink, parquet_sink
from excelmgr.ports.readers import WorkbookReader
from excelmgr.ports.writers import WorkbookWriter


def split(plan: SplitPlan, reader: WorkbookReader, writer: WorkbookWriter) -> dict:
    sheet_ref = plan.sheet.name_or_index if plan.sheet != "active" else 0
    pw = resolve_password(plan.input_file, plan.password, plan.password_map)
    df = reader.read_sheet(plan.input_file, sheet_ref, pw)
    col = plan.by_column
    try:
        if isinstance(col, int):
            key_series = df.iloc[:, col]
            key_name = df.columns[col]
        else:
            key_series = df[col]
            key_name = col
    except IndexError as exc:
        raise ExcelMgrError(
            f"Column index {col} is out of range for sheet {sheet_ref!r}."
        ) from exc
    except KeyError as exc:
        raise ExcelMgrError(
            f"Column '{col}' was not found in sheet {sheet_ref!r}."
        ) from exc

    if not plan.include_nan:
        parts = df[~key_series.isna()].groupby(key_series, dropna=True)
    else:
        parts = df.groupby(key_series, dropna=False)

    if plan.to == "sheets":
        mapping: dict[str, pd.DataFrame] = {}
        seen: set[str] = set()
        for k, g in parts:
            name = sanitize_sheet_name(str(k))
            name = dedupe(name, seen)
            mapping[name] = g
        base = Path(plan.output_dir).expanduser()
        if base.suffix.lower() == ".xlsx":
            out_path = base
        else:
            derived = f"{Path(plan.input_file).stem}_split.xlsx"
            out_path = base / derived
        if not plan.dry_run:
            writer.write_multi_sheets(mapping, str(out_path))
        return {
            "to": "sheets",
            "sheets": list(mapping.keys()),
            "out": str(out_path),
            "by": key_name,
            "dry_run": plan.dry_run,
        }

    # to files
    outputs: list[str] = []
    base_dir = Path(plan.output_dir).expanduser()
    seen_files: set[str] = set()
    ext_map = {"xlsx": ".xlsx", "csv": ".csv", "parquet": ".parquet"}
    for k, g in parts:
        name = sanitize_sheet_name(str(k)) or "Empty"
        unique = dedupe(name, seen_files, max_length=None)
        suffix = ext_map[plan.output_format]
        out_path = base_dir / f"{unique}{suffix}"
        if not plan.dry_run:
            if plan.output_format == "xlsx":
                writer.write_single_sheet(g, str(out_path), sheet_name="Data")
            elif plan.output_format == "csv":
                with csv_sink(str(out_path)) as sink:
                    sink.append(g)
            elif plan.output_format == "parquet":
                with parquet_sink(str(out_path)) as sink:
                    sink.append(g)
        outputs.append(str(out_path))
    return {
        "to": "files",
        "count": len(outputs),
        "out_dir": str(base_dir),
        "by": key_name,
        "format": plan.output_format,
        "dry_run": plan.dry_run,
    }
