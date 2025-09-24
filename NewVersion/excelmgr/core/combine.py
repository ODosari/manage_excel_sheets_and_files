import os
from contextlib import nullcontext
from typing import Dict, Iterable, List

import pandas as pd

from excelmgr.core.models import CombinePlan
from excelmgr.core.naming import sanitize_sheet_name, dedupe
from excelmgr.core.passwords import resolve_password
from excelmgr.core.sinks import csv_sink, parquet_sink
from excelmgr.ports.readers import WorkbookReader
from excelmgr.ports.writers import WorkbookWriter


def _resolve_sheets(reader: WorkbookReader, f: str, include, password: str | None):
    if include == "all":
        return reader.sheet_names(f, password)
    resolved: List[str | int] = []
    for spec in include:
        resolved.append(spec.name_or_index)
    return resolved

def _iter_input_files(reader: WorkbookReader, inputs: Iterable[str], glob: str | None, recursive: bool) -> Iterable[str]:
    for item in inputs:
        if os.path.isdir(item):
            for path in reader.iter_files(item, glob, recursive):
                yield path
        else:
            yield item

class _NullSink:
    def append(self, df: pd.DataFrame) -> None:  # pragma: no cover - trivial
        return


def combine(plan: CombinePlan, reader: WorkbookReader, writer: WorkbookWriter) -> dict:
    sheet_frames: Dict[str, pd.DataFrame] | None = {} if plan.mode == "multi_sheets" and not plan.dry_run else None
    sheet_names: list[str] = []
    combined_rows = 0
    sheet_name_set: set[str] = set()
    files_processed = 0

    if plan.mode == "one_sheet":
        if plan.dry_run:
            sink_cm = nullcontext(_NullSink())
        else:
            if plan.output_format == "xlsx":
                sink_cm = writer.stream_single_sheet(plan.output_path, sheet_name="Data")
            elif plan.output_format == "csv":
                sink_cm = csv_sink(plan.output_path)
            elif plan.output_format == "parquet":
                sink_cm = parquet_sink(plan.output_path)
            else:  # pragma: no cover - guarded by CLI validation
                sink_cm = nullcontext(_NullSink())
    else:
        sink_cm = nullcontext(_NullSink())

    with sink_cm as sink_obj:
        sink = sink_obj or _NullSink()
        for f in _iter_input_files(reader, plan.inputs, plan.glob, plan.recursive):
            files_processed += 1
            pw = resolve_password(f, plan.password, plan.password_map)
            sheets = _resolve_sheets(reader, f, plan.include_sheets, pw)
            for s in sheets:
                df = reader.read_sheet(f, s, pw)
                if plan.add_source_column:
                    df = df.copy()
                    df.insert(0, "source_file", f)
                if plan.mode == "one_sheet":
                    combined_rows += len(df)
                    sink.append(df)
                else:
                    name = sanitize_sheet_name(str(s))
                    name = dedupe(name, sheet_name_set)
                    if sheet_frames is not None:
                        sheet_frames[name] = df
                    sheet_names.append(name)

    if plan.mode == "one_sheet":
        return {
            "mode": "one_sheet",
            "rows": combined_rows,
            "files": files_processed,
            "out": plan.output_path,
            "format": plan.output_format,
            "dry_run": plan.dry_run,
        }

    sheets_out = list(sheet_frames.keys()) if sheet_frames is not None else sheet_names
    if not plan.dry_run and sheet_frames is not None:
        writer.write_multi_sheets(sheet_frames, plan.output_path)
    return {
        "mode": "multi_sheets",
        "sheets": sheets_out,
        "files": files_processed,
        "out": plan.output_path,
        "dry_run": plan.dry_run,
    }
