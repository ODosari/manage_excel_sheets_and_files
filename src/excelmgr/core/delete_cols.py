from __future__ import annotations

import os
import re
from collections.abc import Iterable

import pandas as pd

from excelmgr.adapters.local_storage import iter_files
from excelmgr.core.errors import MissingColumnsError
from excelmgr.core.models import DeleteSpec
from excelmgr.core.passwords import resolve_password
from excelmgr.core.progress import ProgressHook, emit_progress
from excelmgr.ports.readers import WorkbookReader
from excelmgr.ports.writers import WorkbookWriter


def _plan_target_sheets(sheet_names: list[str], spec: DeleteSpec) -> list[tuple[str | int, str]]:
    if not sheet_names:
        return []

    if spec.all_sheets:
        return [(name, name) for name in sheet_names]

    if spec.sheet_selector is None:
        first = sheet_names[0]
        return [(first, first)]

    selector = spec.sheet_selector
    if isinstance(selector, int):
        if 0 <= selector < len(sheet_names):
            display = sheet_names[selector]
        else:
            display = str(selector)
        return [(selector, display)]

    cleaned = str(selector)
    return [(cleaned, cleaned)]


def _match_columns(columns: list[str], spec: DeleteSpec) -> tuple[list[str], list[str]]:
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

def delete_columns(
    spec: DeleteSpec,
    reader: WorkbookReader,
    writer: WorkbookWriter,
    *,
    progress_hooks: Iterable[ProgressHook] | None = None,
) -> dict:
    hooks = tuple(progress_hooks or ())
    emit_progress(
        hooks,
        "delete_start",
        path=spec.path,
        recursive=spec.recursive,
        glob=spec.glob,
        dry_run=spec.dry_run,
    )
    # collect paths
    if os.path.isdir(spec.path):
        paths = list(iter_files(spec.path, spec.glob, spec.recursive))
    else:
        if os.path.basename(spec.path).startswith("~$"):
            paths = []
        else:
            paths = [spec.path]

    summary = []
    missing_records: list[dict] = []
    for index, p in enumerate(paths, 1):
        emit_progress(hooks, "delete_workbook", index=index, path=p)
        pw = resolve_password(p, spec.password, spec.password_map)
        sheet_names = reader.sheet_names(p, pw)
        sheet_cache: dict[str, pd.DataFrame] = {}
        sheets = list(sheet_names)
        targets = _plan_target_sheets(sheets, spec)
        mapping: dict[str, pd.DataFrame] = {}
        per_sheet = []
        for position, (lookup, sheet_name) in enumerate(targets, 1):
            emit_progress(
                hooks,
                "delete_sheet",
                workbook=p,
                sheet=sheet_name,
                index=position,
            )
            cache_key = str(sheet_name)
            df = sheet_cache.get(cache_key)
            if df is None:
                df = reader.read_sheet(p, lookup, pw)
                sheet_cache[cache_key] = df
            remove, missing = _match_columns(list(df.columns), spec)
            new_df = _apply(df, remove)
            mapping[sheet_name] = new_df
            per_sheet.append(
                {
                    "sheet": sheet_name,
                    "removed": remove,
                    "missing": missing,
                    "final_columns": list(new_df.columns),
                }
            )
            if spec.on_missing == "error" and missing:
                missing_records.append({"path": p, "sheet": sheet_name, "missing": missing})

        if missing_records and spec.on_missing == "error":
            break

        if not spec.dry_run:
            out = p if spec.inplace else _build_out_path(p)
            if spec.all_sheets:
                final_mapping = mapping
            else:
                final_mapping = {}
                for name in sheets:
                    if name in mapping:
                        final_mapping[name] = mapping[name]
                    else:
                        cache_key = str(name)
                        original = sheet_cache.get(cache_key)
                        if original is None:
                            original = reader.read_sheet(p, name, pw)
                            sheet_cache[cache_key] = original
                        final_mapping[name] = original
            writer.write_multi_sheets(final_mapping, out)
            summary.append({"path": p, "out": out, "sheets": per_sheet})
        else:
            summary.append({"path": p, "out": None, "sheets": per_sheet})
        emit_progress(
            hooks,
            "delete_workbook_complete",
            path=p,
            sheets=len(per_sheet),
            removed=sum(len(s["removed"]) for s in per_sheet),
        )
    if missing_records and spec.on_missing == "error":
        details = "; ".join(
            f"{item['path']}[{item['sheet']}] missing {', '.join(item['missing'])}" for item in missing_records
        )
        raise MissingColumnsError(f"Columns not found: {details}")

    missing_total = sum(len(sheet["missing"]) for item in summary for sheet in item["sheets"])
    removed_total = sum(len(sheet["removed"]) for item in summary for sheet in item["sheets"])
    result = {
        "updated": len(summary),
        "items": summary,
        "removed_total": removed_total,
        "missing_total": missing_total,
        "dry_run": spec.dry_run,
    }
    emit_progress(
        hooks,
        "delete_complete",
        updated=len(summary),
        removed=removed_total,
        missing=missing_total,
    )
    return result

def _build_out_path(path: str) -> str:
    from pathlib import Path
    p = Path(path)
    return str(p.with_name(p.stem + ".cleaned" + p.suffix))
