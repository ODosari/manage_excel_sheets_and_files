import os
from collections.abc import Iterable
from contextlib import nullcontext

import pandas as pd

from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.models import CloudDestination, CombinePlan, DatabaseDestination
from excelmgr.core.naming import dedupe, sanitize_sheet_name
from excelmgr.core.passwords import resolve_password
from excelmgr.core.sinks import csv_sink, parquet_sink
from excelmgr.ports.readers import WorkbookReader
from excelmgr.ports.writers import CloudObjectWriter, TableWriter, WorkbookWriter


def _resolve_sheets(reader: WorkbookReader, f: str, include, password: str | None):
    if include == "all":
        return reader.sheet_names(f, password)
    resolved: list[str | int] = []
    for spec in include:
        resolved.append(spec.name_or_index)
    return resolved

def _iter_input_files(reader: WorkbookReader, inputs: Iterable[str], glob: str | None, recursive: bool) -> Iterable[str]:
    for item in inputs:
        if os.path.basename(item).startswith("~$"):
            continue
        if os.path.isdir(item):
            yield from reader.iter_files(item, glob, recursive)
        else:
            yield item

class _NullSink:
    def append(self, df: pd.DataFrame) -> None:  # pragma: no cover - trivial
        return


def _cloud_sink_cm(
    plan: CombinePlan,
    destination: CloudDestination,
    *,
    cloud_writer: CloudObjectWriter | None,
):
    if plan.mode != "one_sheet":
        raise ExcelMgrError("Cloud destinations are only supported for one-sheet combine mode.")
    if plan.dry_run:
        return nullcontext(_NullSink())
    if cloud_writer is None:
        raise ExcelMgrError("Cloud destination requested but no cloud writer was provided.")
    return cloud_writer.stream_object(destination.key, destination.format or plan.output_format)


def _database_state(destination: DatabaseDestination | None):
    if destination is None:
        return None
    return {
        "destination": destination,
        "first": True,
    }


def combine(
    plan: CombinePlan,
    reader: WorkbookReader,
    writer: WorkbookWriter,
    *,
    database_writer: TableWriter | None = None,
    cloud_writer: CloudObjectWriter | None = None,
) -> dict:
    sheet_frames: dict[str, pd.DataFrame] | None = {} if plan.mode == "multi_sheets" and not plan.dry_run else None
    sheet_names: list[str] = []
    combined_rows = 0
    sheet_name_set: set[str] = set()
    files_processed = 0

    destination = plan.destination
    database_state = _database_state(destination if isinstance(destination, DatabaseDestination) else None)

    if isinstance(destination, DatabaseDestination) and plan.mode != "one_sheet":
        raise ExcelMgrError("Database destinations are only supported for one-sheet combine mode.")

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
        if isinstance(destination, CloudDestination):
            cloud_cm = _cloud_sink_cm(plan, destination, cloud_writer=cloud_writer)
        else:
            cloud_cm = nullcontext(_NullSink())
    else:
        sink_cm = nullcontext(_NullSink())
        cloud_cm = nullcontext(_NullSink())

    with sink_cm as sink_obj, cloud_cm as cloud_obj:
        sink = sink_obj or _NullSink()
        cloud_sink = cloud_obj or _NullSink()
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
                    cloud_sink.append(df)
                    if database_state and not plan.dry_run:
                        destination = database_state["destination"]
                        if database_writer is None:
                            raise ExcelMgrError(
                                "Database destination requested but no database writer was provided."
                            )
                        mode = destination.mode if database_state["first"] else "append"
                        database_writer.write_dataframe(
                            df,
                            destination.table,
                            mode=mode,
                            options=destination.options,
                            uri=destination.uri,
                        )
                        database_state["first"] = False
                else:
                    name = sanitize_sheet_name(str(s))
                    name = dedupe(name, sheet_name_set)
                    if sheet_frames is not None:
                        sheet_frames[name] = df
                    sheet_names.append(name)

    if plan.mode == "one_sheet":
        result = {
            "mode": "one_sheet",
            "rows": combined_rows,
            "files": files_processed,
            "out": plan.output_path,
            "format": plan.output_format,
            "dry_run": plan.dry_run,
        }
        if destination is not None:
            if isinstance(destination, DatabaseDestination):
                result["destination"] = {"kind": "database", "uri": destination.uri, "table": destination.table}
            elif isinstance(destination, CloudDestination):
                result["destination"] = {"kind": "cloud", "key": destination.key, "root": destination.root}
        return result

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
