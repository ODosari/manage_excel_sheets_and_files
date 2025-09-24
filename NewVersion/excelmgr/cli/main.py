from __future__ import annotations

import json
import logging
import sys
from typing import Annotated, Literal

import typer
from typer.core import TyperOption
from rich import print

from excelmgr.adapters.json_logger import JsonLogger
from excelmgr.adapters.pandas_io import PandasReader, PandasWriter
from excelmgr.config.settings import settings
from excelmgr.core.combine import combine as combine_command
from excelmgr.core.delete_cols import delete_columns as delete_columns_command
from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.models import CombinePlan, DeleteSpec, SheetSpec, SplitPlan
from excelmgr.core.split import split as split_command

def _patch_typer_option() -> None:
    original = TyperOption.make_metavar

    def _make_metavar(self, ctx=None, orig=original):
        if ctx is None:
            name = self.name or "OPTION"
            return (self.metavar or name.upper())
        return orig(self, ctx)

    if original.__code__.co_argcount == 2:
        TyperOption.make_metavar = _make_metavar  # type: ignore[assignment]


_patch_typer_option()


app = typer.Typer(no_args_is_help=True, add_completion=False, rich_markup_mode=None)


def _make_logger(fmt: str, level: str, file: str | None):
    level_num = getattr(logging, level.upper(), logging.INFO)
    return JsonLogger(level=level_num, fmt=fmt, file=file)


def _parse_sheet_list(raw: str) -> list[SheetSpec]:
    tokens = [token.strip() for token in raw.split(",")]
    if any(not token for token in tokens):
        raise ExcelMgrError("Sheet selector contains empty entries. Use comma-separated values without blanks.")

    include: list[SheetSpec] = []
    for token in tokens:
        if token.lower().startswith("index:"):
            _, _, rest = token.partition(":")
            rest = rest.strip()
            if not rest:
                raise ExcelMgrError("index: specifier must include at least one sheet index.")
            for idx in (part.strip() for part in rest.split(",")):
                if not idx:
                    continue
                if not idx.isdigit():
                    raise ExcelMgrError(f"Invalid sheet index '{idx}'.")
                include.append(SheetSpec(int(idx)))
            continue

        if token.isdigit():
            include.append(SheetSpec(int(token)))
        else:
            include.append(SheetSpec(token))

    if not include:
        raise ExcelMgrError("No sheets were specified; provide at least one name or index.")

    return include


def _parse_sheet_option(value: str) -> SheetSpec | Literal["active"]:
    if value == "active":
        return "active"
    cleaned = value.strip()
    if not cleaned:
        raise ExcelMgrError("--sheet cannot be empty.")
    if cleaned.lower().startswith("index:"):
        _, _, rest = cleaned.partition(":")
        rest = rest.strip()
        if not rest or not rest.isdigit():
            raise ExcelMgrError("--sheet index specifier must be numeric, e.g. index:2")
        return SheetSpec(int(rest))
    return SheetSpec(int(cleaned) if cleaned.isdigit() else cleaned)


@app.callback()
def main(
    # settings provide defaults; IDEs can’t infer they match the Literal set
    log_format: Annotated[str, typer.Option("--log")] = settings.log_format,
    log_level: Annotated[str, typer.Option("--log-level")] = settings.log_level,
    log_file: Annotated[str | None, typer.Option("--log-file")] = None,
):
    if log_format not in {"json", "text"}:
        raise typer.BadParameter("--log must be either 'json' or 'text'.")
    if log_level not in {"DEBUG", "INFO", "WARN", "ERROR"}:
        raise typer.BadParameter("--log-level must be DEBUG, INFO, WARN, or ERROR.")
    app.state = {"logger": _make_logger(log_format, log_level, log_file)}


@app.command(help="Combine Excel files into one workbook (one sheet or multi-sheets).")
def combine(
    inputs: Annotated[list[str], typer.Argument(help="Files or directories.")],
    mode: Annotated[str, typer.Option("--mode")] = "one-sheet",
    glob: Annotated[str | None, typer.Option("--glob")] = None,
    recursive: Annotated[bool, typer.Option("--recursive")] = False,
    sheets: Annotated[str, typer.Option("--sheets")] = "all",
    out: Annotated[str, typer.Option("--out")] = "combined.xlsx",
    add_source_column: Annotated[bool, typer.Option("--add-source-column")] = False,
    password: Annotated[str | None, typer.Option("--password")] = None,
    password_env: Annotated[str | None, typer.Option("--password-env")] = None,
    password_file: Annotated[str | None, typer.Option("--password-file")] = None,
    password_map: Annotated[str | None, typer.Option("--password-map")] = None,
    out_format: Annotated[str, typer.Option("--format")] = "xlsx",
    dry_run: Annotated[bool, typer.Option("--dry-run")] = False,
):
    logger = app.state["logger"]
    from .options import read_password_map, read_secret

    try:
        pw = read_secret(password, password_env, password_file)
        pw_map = read_password_map(password_map)

        # parse sheet selection
        include: list[SheetSpec] | Literal["all"]
        if sheets == "all":
            include = "all"
        else:
            include = _parse_sheet_list(sheets)

        mode_map = {"one-sheet": "one_sheet", "multi-sheets": "multi_sheets"}
        if mode not in mode_map:
            raise ExcelMgrError("--mode must be 'one-sheet' or 'multi-sheets'.")
        format_normalized = out_format.lower()
        if format_normalized not in {"xlsx", "csv", "parquet"}:
            raise ExcelMgrError("--format must be xlsx, csv, or parquet.")
        if mode_map[mode] == "multi_sheets" and format_normalized != "xlsx":
            raise ExcelMgrError("Multi-sheet combine output must use Excel format.")

        plan = CombinePlan(
            inputs=inputs,
            glob=glob,
            recursive=recursive,
            mode=mode_map[mode],
            include_sheets=include,
            output_path=out,
            add_source_column=add_source_column,
            password=pw,
            password_map=pw_map,
            output_format=format_normalized,  # type: ignore[arg-type]
            dry_run=dry_run,
        )
        result = combine_command(plan, PandasReader(), PandasWriter())
        logger.info("combine_completed", **result)
        print(json.dumps(result, indent=2))
    except typer.Exit:
        raise
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
    to: Annotated[str, typer.Option("--to")] = "files",
    include_nan: Annotated[bool, typer.Option("--include-nan")] = False,
    out: Annotated[str, typer.Option("--out")] = "out",
    password: Annotated[str | None, typer.Option("--password")] = None,
    password_env: Annotated[str | None, typer.Option("--password-env")] = None,
    password_file: Annotated[str | None, typer.Option("--password-file")] = None,
    password_map: Annotated[str | None, typer.Option("--password-map")] = None,
    out_format: Annotated[str, typer.Option("--format")] = "xlsx",
    dry_run: Annotated[bool, typer.Option("--dry-run")] = False,
):
    logger = app.state["logger"]
    from .options import read_password_map, read_secret

    try:
        pw = read_secret(password, password_env, password_file)
        pw_map = read_password_map(password_map)
        sheet_spec = _parse_sheet_option(sheet)
        by_clean = by.strip()
        if not by_clean:
            raise ExcelMgrError("--by cannot be empty.")
        by_col: str | int = int(by_clean) if by_clean.isdigit() else by_clean
        if to not in {"files", "sheets"}:
            raise ExcelMgrError("--to must be either 'files' or 'sheets'.")
        format_normalized = out_format.lower()
        if format_normalized not in {"xlsx", "csv", "parquet"}:
            raise ExcelMgrError("--format must be xlsx, csv, or parquet.")
        if to == "sheets" and format_normalized != "xlsx":
            raise ExcelMgrError("Sheet mode output must be Excel format.")

        plan = SplitPlan(
            input_file=input_file,
            sheet=sheet_spec,
            by_column=by_col,
            to=to,  # already Literal-aligned
            include_nan=include_nan,
            output_dir=out,
            password=pw,
            password_map=pw_map,
            output_format=format_normalized,  # type: ignore[arg-type]
            dry_run=dry_run,
        )
        result = split_command(plan, PandasReader(), PandasWriter())
        logger.info("split_completed", **result)
        print(json.dumps(result, indent=2))
    except typer.Exit:
        raise
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
    match: Annotated[str, typer.Option("--match")] = "names",
    strategy: Annotated[
        str,
        typer.Option("--strategy"),
    ] = "exact",
    all_sheets: Annotated[bool, typer.Option("--all-sheets")] = False,
    sheet: Annotated[str | None, typer.Option("--sheet")] = None,
    inplace: Annotated[bool, typer.Option("--inplace")] = False,
    on_missing: Annotated[str, typer.Option("--on-missing")] = "ignore",
    dry_run: Annotated[bool, typer.Option("--dry-run")] = False,
    glob: Annotated[str | None, typer.Option("--glob")] = None,
    recursive: Annotated[bool, typer.Option("--recursive")] = False,
    password: Annotated[str | None, typer.Option("--password")] = None,
    password_env: Annotated[str | None, typer.Option("--password-env")] = None,
    password_file: Annotated[str | None, typer.Option("--password-file")] = None,
    password_map: Annotated[str | None, typer.Option("--password-map")] = None,
    yes: Annotated[bool, typer.Option("--yes")] = False,
):
    logger = app.state["logger"]
    from .options import read_password_map, read_secret

    try:
        pw = read_secret(password, password_env, password_file)
        pw_map = read_password_map(password_map)
        if match not in {"names", "index"}:
            raise ExcelMgrError("--match must be 'names' or 'index'.")
        allowed_strategies = {"exact", "ci", "contains", "startswith", "endswith", "regex"}
        if strategy not in allowed_strategies:
            raise ExcelMgrError("--strategy value is invalid.")
        if on_missing not in {"ignore", "error"}:
            raise ExcelMgrError("--on-missing must be 'ignore' or 'error'.")

        if (not yes) and (not dry_run) and (not inplace):
            proceed = typer.confirm("Write cleaned copies next to originals?", default=True, abort=False)
            if not proceed:
                logger.info("delete_cols_aborted", reason="user_declined_confirmation")
                raise typer.Exit(code=0)

        # targets -> Sequence[str] | Sequence[int]
        if match == "index":
            try:
                targets_list = [int(x) for x in targets.split(",") if x.strip()]
            except ValueError as exc:
                raise ExcelMgrError("--targets must be integers when --match index is used.") from exc
        else:
            targets_list: list[str] = [t.strip() for t in targets.split(",") if t.strip()]

        if not targets_list:
            raise ExcelMgrError("--targets must include at least one entry.")

        sheet_selector = None
        if sheet is not None:
            cleaned_sheet = sheet.strip()
            if not cleaned_sheet:
                raise ExcelMgrError("--sheet cannot be empty when provided.")
            if cleaned_sheet.lower().startswith("index:"):
                _, _, rest = cleaned_sheet.partition(":")
                rest = rest.strip()
                if not rest or not rest.isdigit():
                    raise ExcelMgrError("--sheet index specifier must be numeric, e.g. index:2")
                sheet_selector = int(rest)
            elif cleaned_sheet.isdigit():
                sheet_selector = int(cleaned_sheet)
            else:
                sheet_selector = cleaned_sheet

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
            password_map=pw_map,
        )
        result = delete_columns_command(spec, PandasReader(), PandasWriter())
        logger.info("delete_cols_completed", **result)
        print(json.dumps(result, indent=2))
    except typer.Exit:
        raise
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
    from excelmgr import __version__

    print(__version__)
