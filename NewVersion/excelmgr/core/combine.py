import os
from typing import Dict, Iterable, List

import pandas as pd

from excelmgr.core.models import CombinePlan
from excelmgr.core.naming import sanitize_sheet_name, dedupe
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

def combine(plan: CombinePlan, reader: WorkbookReader, writer: WorkbookWriter) -> dict:
    # Use separate data structures for the two different modes to maintain type safety.
    dfs_to_concat: List[pd.DataFrame] = []  # For "one_sheet" mode
    sheet_frames: Dict[str, pd.DataFrame] = {}  # For "multi_sheets" mode
    combined_rows = 0
    sheet_name_set: set[str] = set()
    files_processed = 0

    for f in _iter_input_files(reader, plan.inputs, plan.glob, plan.recursive):
        files_processed += 1
        sheets = _resolve_sheets(reader, f, plan.include_sheets, plan.password)
        for s in sheets:
            df = reader.read_sheet(f, s, plan.password)
            if plan.add_source_column:
                df = df.copy()
                df.insert(0, "source_file", f)
            if plan.mode == "one_sheet":
                combined_rows += len(df)
                dfs_to_concat.append(df)
            else:
                name = sanitize_sheet_name(str(s))
                name = dedupe(name, sheet_name_set)
                sheet_frames[name] = df

    if plan.mode == "one_sheet":
        final = pd.concat(dfs_to_concat, ignore_index=True) if dfs_to_concat else pd.DataFrame()
        writer.write_single_sheet(final, plan.output_path, sheet_name="Data")
        return {"mode": "one_sheet", "rows": len(final), "files": files_processed, "out": plan.output_path}
    else:
        writer.write_multi_sheets(sheet_frames, plan.output_path)
        return {"mode": "multi_sheets", "sheets": list(sheet_frames.keys()), "files": files_processed, "out": plan.output_path}
