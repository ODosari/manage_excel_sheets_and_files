from typing import Dict
import pandas as pd
from NewVersion.excelmgr.core.models import SplitPlan
from NewVersion.excelmgr.core.naming import sanitize_sheet_name, dedupe
from NewVersion.excelmgr.ports.readers import WorkbookReader
from NewVersion.excelmgr.ports.writers import WorkbookWriter

def split(plan: SplitPlan, reader: WorkbookReader, writer: WorkbookWriter) -> dict:
    df = reader.read_sheet(plan.input_file, plan.sheet.name_or_index if plan.sheet != "active" else 0, plan.password)
    col = plan.by_column
    if isinstance(col, int):
        key_series = df.iloc[:, col]
        key_name = df.columns[col]
    else:
        key_series = df[col]
        key_name = col

    if not plan.include_nan:
        parts = df[~key_series.isna()].groupby(key_series, dropna=True)
    else:
        parts = df.groupby(key_series, dropna=False)

    if plan.to == "sheets":
        mapping: Dict[str, pd.DataFrame] = {}
        seen: set[str] = set()
        for k, g in parts:
            name = sanitize_sheet_name(str(k))
            name = dedupe(name, seen)
            mapping[name] = g
        out = plan.output_dir if plan.output_dir.lower().endswith(".xlsx") else f"{plan.output_dir.rstrip('/')}/split.xlsx"
        writer.write_multi_sheets(mapping, out)
        return {"to": "sheets", "sheets": list(mapping.keys()), "out": out, "by": key_name}

    # to files
    outputs = []
    for k, g in parts:
        name = sanitize_sheet_name(str(k)) or "Empty"
        out = f"{plan.output_dir.rstrip('/')}/{name}.xlsx"
        writer.write_single_sheet(g, out, sheet_name="Data")
        outputs.append(out)
    return {"to": "files", "count": len(outputs), "out_dir": plan.output_dir, "by": key_name}
