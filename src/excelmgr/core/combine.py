import os
from collections.abc import Iterable
from contextlib import nullcontext

import pandas as pd

from excelmgr.core.models import CombinePlan
from excelmgr.core.naming import dedupe, sanitize_sheet_name
from excelmgr.core.passwords import resolve_password
from excelmgr.core.sinks import csv_sink, parquet_sink
from excelmgr.ports.readers import WorkbookReader
from excelmgr.ports.writers import WorkbookWriter


def _resolve_sheets(reader: WorkbookReader, f: str, include, password: str | None):
    if include == "all":
        return reader.sheet_names(f, password)
    resolved: list[str | int] = []
    for spec in include:
        resolved.append(spec.name_or_index)
    return resolved

def _iter_input_files(reader: WorkbookReader, inputs: Iterable[str], glob: str | None, recursive: bool) -> Iterable[str]:
    for item in inputs:
        if os.path.isdir(item):
            yield from reader.iter_files(item, glob, recursive)
        else:
            yield item

class _NullSink:
    def append(self, df: pd.DataFrame) -> None:  # pragma: no cover - trivial
        return


class _NullMultiSink:
    def write_sheet(self, name: str, df: pd.DataFrame) -> None:  # pragma: no cover - trivial
        return


def _report(plan: CombinePlan, event: str, **payload) -> None:
    if plan.progress is not None:
        plan.progress(event, payload)


def combine(plan: CombinePlan, reader: WorkbookReader, writer: WorkbookWriter) -> dict:
    sheet_names: list[str] = []
    combined_rows = 0
    sheet_name_set: set[str] = set()
    files_processed = 0

    if plan.mode == "one_sheet":
        if plan.dry_run:
            sink_cm = nullcontext(_NullSink())
        else:
            if plan.output_format == "xlsx":
                sink_cm = writer.stream_single_sheet(plan.output_path, sheet_name=plan.sheet_name)
            elif plan.output_format == "csv":
                sink_cm = csv_sink(plan.output_path)
            elif plan.output_format == "parquet":
                sink_cm = parquet_sink(plan.output_path)
            else:  # pragma: no cover - guarded by CLI validation
                sink_cm = nullcontext(_NullSink())
    else:
        if plan.dry_run:
            sink_cm = nullcontext(_NullMultiSink())
        else:
            sink_cm = writer.stream_multi_sheets(plan.output_path)

    with sink_cm as sink_obj:
        sink = sink_obj or (_NullSink() if plan.mode == "one_sheet" else _NullMultiSink())
        for f in _iter_input_files(reader, plan.inputs, plan.glob, plan.recursive):
            files_processed += 1
            _report(plan, "combine_file_started", file=f, index=files_processed)
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
                    _report(
                        plan,
                        "combine_sheet_appended",
                        file=f,
                        sheet=str(s),
                        rows=len(df),
                        index=files_processed,
                    )
                else:
                    name = sanitize_sheet_name(str(s))
                    name = dedupe(name, sheet_name_set)
                    sheet_names.append(name)
                    sink.write_sheet(name, df)
                    _report(
                        plan,
                        "combine_sheet_written",
                        file=f,
                        sheet=str(s),
                        output_sheet=name,
                        rows=len(df),
                        index=files_processed,
                    )

    if plan.mode == "one_sheet":
        return {
            "mode": "one_sheet",
            "rows": combined_rows,
            "files": files_processed,
            "out": plan.output_path,
            "format": plan.output_format,
            "dry_run": plan.dry_run,
        }

    return {
        "mode": "multi_sheets",
        "sheets": sheet_names,
        "files": files_processed,
        "out": plan.output_path,
        "dry_run": plan.dry_run,
    }
