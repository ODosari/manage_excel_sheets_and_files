from __future__ import annotations

import os
import re
from typing import Dict, List

import pandas as pd

from NewVersion.excelmgr.adapters.local_storage import iter_files
from NewVersion.excelmgr.core.models import DeleteSpec
from NewVersion.excelmgr.ports.readers import WorkbookReader
from NewVersion.excelmgr.ports.writers import WorkbookWriter


def _match_columns(columns: List[str], spec: DeleteSpec) -> tuple[list[str], list[str]]:
    cols_norm = [str(c).strip() for c in columns]
    to_remove: list[str] = []
    not_found: list[str] = []

    if spec.match_kind == "index":
        idx_targets = [int(x) for x in spec.targets]  # 1-based
        for i in idx_targets:
            pos = i - 1
            if 0 <= pos < len(cols_norm):
                to_remove.append(cols_norm[pos])
            else:
                not_found.append(str(i))
        # dedupe
        seen = set()
        to_remove = [x for x in to_remove if not (x in seen or seen.add(x))]
        return to_remove, not_found

    # names
    wanted = [str(t).strip() for t in spec.targets]
    for w in wanted:
        matched = []
        for c in cols_norm:
            c_cmp = c
            w_cmp = w
            if spec.strategy in ("ci", "case_insensitive"):
                c_cmp, w_cmp = c.lower(), w.lower()
            if spec.strategy in ("exact", "ci", "case_insensitive"):
                ok = c_cmp == w_cmp
            elif spec.strategy == "contains":
                ok = w_cmp in c_cmp
            elif spec.strategy == "startswith":
                ok = c_cmp.startswith(w_cmp)
            elif spec.strategy == "endswith":
                ok = c_cmp.endswith(w_cmp)
            elif spec.strategy == "regex":
                ok = re.search(w, c) is not None
            else:
                ok = c_cmp == w_cmp
            if ok:
                matched.append(c)
        if matched:
            to_remove.extend(matched)
        else:
            not_found.append(w)
    # dedupe
    seen = set()
    to_remove = [x for x in to_remove if not (x in seen or seen.add(x))]
    return to_remove, not_found

def _apply(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    return df.drop(columns=cols, errors="ignore")

def delete_columns(spec: DeleteSpec, reader: WorkbookReader, writer: WorkbookWriter) -> dict:
    # collect paths
    if os.path.isdir(spec.path):
        paths = list(iter_files(spec.path, spec.glob, spec.recursive))
    else:
        paths = [spec.path]

    summary = []
    for p in paths:
        sheets = reader.sheet_names(p, spec.password)
        target_sheets = sheets if spec.all_sheets or spec.sheet_selector is None else [spec.sheet_selector]
        # read once
        mapping: Dict[str, pd.DataFrame] = {}
        per_sheet = []
        for s in target_sheets:
            df = reader.read_sheet(p, s if s is not None else 0, spec.password)
            remove, missing = _match_columns(list(df.columns), spec)
            new_df = _apply(df, remove)
            mapping[str(s)] = new_df
            per_sheet.append({"sheet": str(s), "removed": remove, "missing": missing, "final_columns": list(new_df.columns)})

        if not spec.dry_run:
            out = p if spec.inplace else _build_out_path(p)
            # For untouched sheets (when not all_sheets), keep originals
            if spec.sheet_selector is not None and not spec.all_sheets:
                for s in sheets:
                    if str(s) not in mapping and s != spec.sheet_selector:
                        mapping[str(s)] = reader.read_sheet(p, s, spec.password)
            writer.write_multi_sheets(mapping, out)
            summary.append({"path": p, "out": out, "sheets": per_sheet})
        else:
            summary.append({"path": p, "out": None, "sheets": per_sheet})
    return {"updated": len(summary), "items": summary}

def _build_out_path(path: str) -> str:
    from pathlib import Path
    p = Path(path)
    return str(p.with_name(p.stem + ".cleaned" + p.suffix))
