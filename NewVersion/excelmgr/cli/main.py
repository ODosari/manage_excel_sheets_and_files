from __future__ import annotations

import json
import logging
import sys
from typing import Annotated, Literal

import typer
from rich import print

from NewVersion.excelmgr.adapters.json_logger import JsonLogger
from NewVersion.excelmgr.adapters.pandas_io import PandasReader, PandasWriter
from NewVersion.excelmgr.config.settings import settings
from NewVersion.excelmgr.core.combine import combine as combine_command
from NewVersion.excelmgr.core.delete_cols import delete_columns as delete_columns_command
from NewVersion.excelmgr.core.errors import ExcelMgrError
from NewVersion.excelmgr.core.models import CombinePlan, DeleteSpec, SheetSpec, SplitPlan
from NewVersion.excelmgr.core.split import split as split_command

app = typer.Typer(no_args_is_help=True, add_completion=False)


def _make_logger(fmt: str, level: str, file: str | None):
    level_num = getattr(logging, level.upper(), logging.INFO)
    return JsonLogger(level=level_num, fmt=fmt, file=file)


@app.callback()
def main(
    # settings provide defaults; IDEs can’t infer they match the Literal set
    log_format: Annotated[Literal["json", "text"], typer.Option("--log")] = settings.log_format,  # type: ignore[assignment]
    log_level: Annotated[Literal["DEBUG", "INFO", "WARN", "ERROR"], typer.Option("--log-level")] = settings.log_level,  # type: ignore[assignment]
    log_file: Annotated[str | None, typer.Option("--log-file")] = None,
):
    app.state = {"logger": _make_logger(log_format, log_level, log_file)}


@app.command(help="Combine Excel files into one workbook (one sheet or multi-sheets).")
def combine(
    inputs: Annotated[list[str], typer.Argument(help="Files or directories.")],
    mode: Annotated[Literal["one-sheet", "multi-sheets"], typer.Option("--mode")] = "one-sheet",
    glob: Annotated[str | None, typer.Option("--glob")] = None,
    recursive: Annotated[bool, typer.Option("--recursive")] = False,
    sheets: Annotated[str, typer.Option("--sheets")] = "all",
    out: Annotated[str, typer.Option("--out")] = "combined.xlsx",
    add_source_column: Annotated[bool, typer.Option("--add-source-column")] = False,
    password: Annotated[str | None, typer.Option("--password")] = None,
    password_env: Annotated[str | None, typer.Option("--password-env")] = None,
    password_file: Annotated[str | None, typer.Option("--password-file")] = None,
):
    logger = app.state["logger"]
    from .options import read_secret

    try:
        pw = read_secret(password, password_env, password_file)

        # parse sheet selection
        include: list[SheetSpec] | Literal["all"]
        if sheets == "all":
            include = "all"
        else:
            include = []
            for token in sheets.split(","):
                token = token.strip()
                if token.startswith("index:"):
                    for t in token.split(":", 1)[1].split(","):
                        include.append(SheetSpec(int(t)))
                else:
                    include.append(SheetSpec(int(token) if token.isdigit() else token))

        mode_map = {"one-sheet": "one_sheet", "multi-sheets": "multi_sheets"}
        plan = CombinePlan(
            inputs=inputs,
            glob=glob,
            recursive=recursive,
            mode=mode_map[mode],
            include_sheets=include,
            output_path=out,
            add_source_column=add_source_column,
            password=pw,
        )
        result = combine_command(plan, PandasReader(), PandasWriter())
        logger.info("combine_completed", **result)
        print(json.dumps(result, indent=2))
    except ExcelMgrError as e:
        logger.error("combine_failed", error=str(e))
        raise typer.Exit(code=2)
    except Exception as e:
        logger.error("combine_crash", error=str(e))
        raise typer.Exit(code=1)


@app.command(help="Split a sheet by a column into many sheets or files.")
def split(
    input_file: Annotated[str, typer.Argument(help="Input workbook (.xlsx).")],
    sheet: Annotated[str, typer.Option("--sheet")] = "active",
    by: Annotated[str, typer.Option("--by")] = ...,
    to: Annotated[Literal["files", "sheets"], typer.Option("--to")] = "files",
    include_nan: Annotated[bool, typer.Option("--include-nan")] = False,
    out: Annotated[str, typer.Option("--out")] = "out",
    password: Annotated[str | None, typer.Option("--password")] = None,
    password_env: Annotated[str | None, typer.Option("--password-env")] = None,
    password_file: Annotated[str | None, typer.Option("--password-file")] = None,
):
    logger = app.state["logger"]
    from .options import read_secret

    try:
        pw = read_secret(password, password_env, password_file)
        sheet_spec = "active" if sheet == "active" else SheetSpec(int(sheet) if sheet.isdigit() else sheet)
        by_col: str | int = int(by) if by.isdigit() else by

        plan = SplitPlan(
            input_file=input_file,
            sheet=sheet_spec,
            by_column=by_col,
            to=to,  # already Literal-aligned
            include_nan=include_nan,
            output_dir=out,
            password=pw,
        )
        result = split_command(plan, PandasReader(), PandasWriter())
        logger.info("split_completed", **result)
        print(json.dumps(result, indent=2))
    except ExcelMgrError as e:
        logger.error("split_failed", error=str(e))
        raise typer.Exit(code=2)
    except Exception as e:
        logger.error("split_crash", error=str(e))
        raise typer.Exit(code=1)


@app.command("delete-cols", help="Delete columns across files/sheets.")
def delete_cols(
    path: Annotated[str, typer.Argument(help="File or directory.")],
    targets: Annotated[str, typer.Option("--targets")] = ...,
    match: Annotated[Literal["names", "index"], typer.Option("--match")] = "names",
    strategy: Annotated[
        Literal["exact", "ci", "contains", "startswith", "endswith", "regex"],
        typer.Option("--strategy"),
    ] = "exact",
    all_sheets: Annotated[bool, typer.Option("--all-sheets")] = False,
    sheet: Annotated[str | None, typer.Option("--sheet")] = None,
    inplace: Annotated[bool, typer.Option("--inplace")] = False,
    on_missing: Annotated[Literal["ignore", "error"], typer.Option("--on-missing")] = "ignore",
    dry_run: Annotated[bool, typer.Option("--dry-run")] = False,
    glob: Annotated[str | None, typer.Option("--glob")] = None,
    recursive: Annotated[bool, typer.Option("--recursive")] = False,
    password: Annotated[str | None, typer.Option("--password")] = None,
    password_env: Annotated[str | None, typer.Option("--password-env")] = None,
    password_file: Annotated[str | None, typer.Option("--password-file")] = None,
    yes: Annotated[bool, typer.Option("--yes")] = False,
):
    logger = app.state["logger"]
    from .options import read_secret

    try:
        pw = read_secret(password, password_env, password_file)

        if (not yes) and (not dry_run) and (not inplace):
            typer.confirm("Write cleaned copies next to originals?", default=True, abort=False)

        # targets -> Sequence[str] | Sequence[int]
        if match == "index":
            targets_list: list[int] = [int(x) for x in targets.split(",") if x.strip()]
        else:
            targets_list: list[str] = [t.strip() for t in targets.split(",") if t.strip()]

        sheet_selector = None
        if sheet is not None:
            sheet_selector = int(sheet) if sheet.isdigit() else sheet

        spec = DeleteSpec(
            path=path,
            targets=targets_list,
            match_kind=match,
            strategy=strategy,
            all_sheets=all_sheets,
            sheet_selector=sheet_selector,
            inplace=inplace,
            on_missing=on_missing,
            dry_run=dry_run,
            glob=glob,
            recursive=recursive,
            password=pw,
        )
        result = delete_columns_command(spec, PandasReader(), PandasWriter())
        logger.info("delete_cols_completed", **result)
        print(json.dumps(result, indent=2))
    except ExcelMgrError as e:
        logger.error("delete_cols_failed", error=str(e))
        raise typer.Exit(code=2)
    except Exception as e:
        logger.error("delete_cols_crash", error=str(e))
        raise typer.Exit(code=1)


@app.command(help="Print environment diagnostics.")
def diagnose():
    import platform
    import openpyxl
    import pandas

    info = {
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "pandas": pandas.__version__,
        "openpyxl": openpyxl.__version__,
        "settings": {
            "glob": settings.glob,
            "recursive": settings.recursive,
            "log_format": settings.log_format,
            "macro_policy": settings.macro_policy,
        },
    }
    print(json.dumps(info, indent=2))


@app.command(help="Show version.")
def version():
    from NewVersion.excelmgr import __version__

    print(__version__)
