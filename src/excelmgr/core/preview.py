from __future__ import annotations

from excelmgr.core.models import PreviewPlan
from excelmgr.core.passwords import resolve_password
from excelmgr.ports.readers import WorkbookReader


def preview(plan: PreviewPlan, reader: WorkbookReader) -> dict:
    password = resolve_password(plan.path, plan.password, plan.password_map)
    names = reader.sheet_names(plan.path, password)
    limit = plan.limit
    sheets: list[dict] = []
    for name in names:
        df = reader.read_sheet(plan.path, name, password)
        info: dict[str, object] = {
            "name": name,
            "rows": len(df),
            "columns": len(df.columns),
            "headers": [str(col) for col in df.columns],
        }
        if limit is not None and limit > 0:
            info["sample"] = df.head(limit).to_dict(orient="records")
        sheets.append(info)
    return {
        "path": plan.path,
        "sheets": sheets,
    }
