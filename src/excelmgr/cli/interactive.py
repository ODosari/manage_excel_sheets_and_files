"""Interactive entry point for the Excel Manager CLI.

This module provides a guided, menu-driven interface that mirrors the
non-interactive Typer commands. It reuses the same validators, orchestration
functions, and logging hooks so that behaviour is consistent regardless of how
the CLI is invoked.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Callable, Sequence

import typer

from excelmgr.adapters.pandas_io import PandasReader, PandasWriter
from excelmgr.cli import main as cli_main
from excelmgr.cli.options import read_password_map, read_secret
from excelmgr.core.combine import combine as combine_command
from excelmgr.core.delete_cols import delete_columns as delete_columns_command
from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.models import CombinePlan, DeleteSpec, PreviewPlan, SplitPlan
from excelmgr.core.plan_runner import execute_plan, load_plan_file
from excelmgr.core.preview import preview as preview_command
from excelmgr.core.split import split as split_command


class BackRequested(Exception):
    """Raised when the user asks to go back to the previous menu."""


@dataclass
class MenuChoice:
    value: str
    label: str
    aliases: tuple[str, ...] = ()


def _get_logger():
    return cli_main.app.state["logger"]


def prompt_menu(title: str, choices: Sequence[MenuChoice]) -> MenuChoice:
    typer.echo()
    typer.echo(title)
    for idx, option in enumerate(choices, start=1):
        typer.echo(f"  {idx}. {option.label}")

    while True:
        raw = typer.prompt("Select an option").strip()
        if not raw:
            typer.secho("Please choose an option.", fg=typer.colors.YELLOW)
            continue
        lowered = raw.lower()
        if raw.isdigit():
            idx = int(raw)
            if 1 <= idx <= len(choices):
                return choices[idx - 1]
        for choice in choices:
            if lowered == choice.value:
                return choice
            if lowered == choice.label.lower():
                return choice
            if lowered in choice.aliases:
                return choice
        typer.secho("Invalid selection. Try again.", fg=typer.colors.RED)


def _ensure_not_back(value: str) -> str:
    if value.lower() == "back":
        raise BackRequested
    return value


def _prompt_text(prompt: str, *, default: str | None = None, allow_empty: bool = False) -> str:
    while True:
        if default is None:
            raw = typer.prompt(prompt)
        else:
            raw = typer.prompt(prompt, default=default)
        text = _ensure_not_back(raw.strip())
        if text or allow_empty:
            return text
        typer.secho("Value cannot be empty. Type 'back' to return to the previous menu.", fg=typer.colors.RED)


def _prompt_optional_text(prompt: str) -> str | None:
    raw = typer.prompt(prompt, default="")
    text = _ensure_not_back(raw.strip())
    return text or None


def _prompt_password_inputs() -> tuple[str | None, dict[str, str] | None]:
    password_source: dict[str, str | None] = {"password": None, "password_env": None, "password_file": None}
    password_map_path: str | None = None

    def _show_summary() -> None:
        summary_parts: list[str] = []
        if password_source["password"]:
            summary_parts.append("manual password set")
        if password_source["password_env"]:
            summary_parts.append(f"env:{password_source['password_env']}")
        if password_source["password_file"]:
            summary_parts.append(f"file:{password_source['password_file']}")
        if password_map_path:
            summary_parts.append(f"map:{password_map_path}")
        if summary_parts:
            typer.echo("Current password settings: " + ", ".join(summary_parts))
        else:
            typer.echo("No password settings configured.")

    while True:
        _show_summary()
        choice = prompt_menu(
            "Password options (type 'back' to return)",
            [
                MenuChoice("manual", "Type password manually"),
                MenuChoice("env", "Read password from environment variable"),
                MenuChoice("file", "Load password from file"),
                MenuChoice("map", "Use password map file"),
                MenuChoice("done", "Done"),
                MenuChoice("back", "Back"),
            ],
        )
        if choice.value == "manual":
            password_source = {"password": typer.prompt("Password", hide_input=True), "password_env": None, "password_file": None}
        elif choice.value == "env":
            env_name = _prompt_text("Environment variable name")
            password_source = {"password": None, "password_env": env_name, "password_file": None}
        elif choice.value == "file":
            file_path = _prompt_text("Path to password file")
            password_source = {"password": None, "password_env": None, "password_file": file_path}
        elif choice.value == "map":
            try:
                password_map_path = _prompt_text("Path to password map file")
                password_map = read_password_map(password_map_path)
                typer.echo("Password map loaded.")
            except BackRequested:
                password_map_path = None
                continue
            except typer.BadParameter as exc:
                typer.secho(str(exc), fg=typer.colors.RED)
                password_map_path = None
                continue
            else:
                return read_secret(**password_source), password_map
        elif choice.value == "done":
            try:
                secret = read_secret(**password_source)
                password_map = read_password_map(password_map_path)
                return secret, password_map
            except typer.BadParameter as exc:
                typer.secho(str(exc), fg=typer.colors.RED)
        elif choice.value == "back":
            raise BackRequested


def _prompt_try_again(action: str) -> bool:
    follow_up = prompt_menu(
        f"{action} - what would you like to do next?",
        [
            MenuChoice("again", "Try again", aliases=("retry",)),
            MenuChoice("back", "Back to main menu", aliases=("exit", "quit")),
        ],
    )
    return follow_up.value == "again"


def _prompt_paths_list(message: str) -> list[str]:
    while True:
        raw = _prompt_text(message)
        parts = [part.strip() for part in raw.split(",") if part.strip()]
        if parts:
            return parts
        typer.secho("Provide at least one entry.", fg=typer.colors.RED)


def _prompt_combine_plan() -> CombinePlan:
    typer.echo("\nCombine workflow (type 'back' at any prompt to return).")
    inputs = _prompt_paths_list("Input files or directories (comma separated)")

    mode_choice = prompt_menu(
        "Select combine mode",
        [
            MenuChoice("one_sheet", "One sheet", aliases=("one-sheet", "1")),
            MenuChoice("multi_sheets", "Multi sheets", aliases=("multi-sheets", "2")),
        ],
    )
    glob = _prompt_optional_text("Glob pattern (leave blank for none)")
    recursive = typer.confirm("Search recursively?", default=False)

    while True:
        sheets_raw = _prompt_text("Sheets to include ('all' or comma separated)", default="all")
        if sheets_raw.lower() == "all":
            include: list | str = "all"
            break
        try:
            include = cli_main._parse_sheet_list(sheets_raw)
            break
        except ExcelMgrError as exc:
            typer.secho(str(exc), fg=typer.colors.RED)

    output_path = _prompt_text("Output path", default="combined.xlsx")
    sheet_name = _prompt_text("Output sheet name", default="Data")
    add_source_column = typer.confirm("Add source column?", default=False)

    format_choice = prompt_menu(
        "Select output format",
        [
            MenuChoice("xlsx", "Excel (.xlsx)"),
            MenuChoice("csv", "CSV"),
            MenuChoice("parquet", "Parquet"),
        ],
    )
    dry_run = typer.confirm("Dry run (no files written)?", default=False)

    password, password_map = _prompt_password_inputs()

    mode = mode_choice.value
    if mode == "multi_sheets" and format_choice.value != "xlsx":
        typer.secho("Multi-sheet mode requires Excel output format.", fg=typer.colors.RED)
        if _prompt_try_again("Combine configuration invalid"):
            return _prompt_combine_plan()
        raise BackRequested

    plan = CombinePlan(
        inputs=inputs,
        glob=glob,
        recursive=recursive,
        mode=mode,
        include_sheets=include,  # type: ignore[arg-type]
        output_path=output_path,
        output_sheet_name=sheet_name,
        add_source_column=add_source_column,
        password=password,
        password_map=password_map,
        output_format=format_choice.value,  # type: ignore[arg-type]
        dry_run=dry_run,
    )
    return plan


def _prompt_split_plan() -> SplitPlan:
    typer.echo("\nSplit workflow (type 'back' at any prompt to return).")
    input_file = _prompt_text("Input workbook path")

    while True:
        sheet_raw = _prompt_text("Sheet to split (name, index, or 'active')", default="active")
        try:
            sheet_spec = cli_main._parse_sheet_option(sheet_raw)
            break
        except ExcelMgrError as exc:
            typer.secho(str(exc), fg=typer.colors.RED)

    by_column = _prompt_text("Column to split by")
    by_clean = by_column.strip()
    by_value: str | int = int(by_clean) if by_clean.isdigit() else by_clean

    destination = prompt_menu(
        "Split destination",
        [
            MenuChoice("files", "Create files"),
            MenuChoice("sheets", "Create sheets"),
        ],
    )

    include_nan = typer.confirm("Include rows with missing split values?", default=False)
    output_dir = _prompt_text("Output directory", default="out")
    output_file = _prompt_optional_text("Explicit output filename (leave blank to auto-generate)")
    output_sheet_name = _prompt_text("Output sheet name", default="Data")

    format_choice = prompt_menu(
        "Select output format",
        [
            MenuChoice("xlsx", "Excel (.xlsx)"),
            MenuChoice("csv", "CSV"),
            MenuChoice("parquet", "Parquet"),
        ],
    )

    if destination.value == "sheets" and format_choice.value != "xlsx":
        typer.secho("Sheet destination requires Excel output format.", fg=typer.colors.RED)
        if _prompt_try_again("Split configuration invalid"):
            return _prompt_split_plan()
        raise BackRequested

    dry_run = typer.confirm("Dry run (no files written)?", default=False)

    password, password_map = _prompt_password_inputs()

    plan = SplitPlan(
        input_file=input_file,
        sheet=sheet_spec,
        by_column=by_value,
        to=destination.value,  # type: ignore[arg-type]
        include_nan=include_nan,
        output_dir=output_dir,
        output_filename=output_file,
        output_sheet_name=output_sheet_name,
        password=password,
        password_map=password_map,
        output_format=format_choice.value,  # type: ignore[arg-type]
        dry_run=dry_run,
    )
    return plan


def _prompt_preview_plan() -> PreviewPlan:
    typer.echo("\nPreview workflow (type 'back' at any prompt to return).")
    path = _prompt_text("Workbook path")
    limit_raw = _prompt_optional_text("Sample row limit per sheet (leave blank for unlimited)")
    limit_value: int | None = None
    if limit_raw is not None:
        if not limit_raw.isdigit():
            typer.secho("Limit must be a positive integer.", fg=typer.colors.RED)
            if _prompt_try_again("Preview configuration invalid"):
                return _prompt_preview_plan()
            raise BackRequested
        limit_value = int(limit_raw)

    password, password_map = _prompt_password_inputs()

    plan = PreviewPlan(path=path, password=password, password_map=password_map, limit=limit_value)
    return plan


def _prompt_delete_spec() -> DeleteSpec:
    typer.echo("\nDelete columns workflow (type 'back' at any prompt to return).")
    path = _prompt_text("Workbook path")

    match_choice = prompt_menu(
        "Match columns by",
        [
            MenuChoice("names", "Names"),
            MenuChoice("index", "Index"),
        ],
    )
    targets_raw = _prompt_text("Columns to delete (comma separated)")

    if match_choice.value == "index":
        try:
            targets = [int(t.strip()) for t in targets_raw.split(",") if t.strip()]
        except ValueError:
            typer.secho("Targets must be integers when matching by index.", fg=typer.colors.RED)
            if _prompt_try_again("Delete columns configuration invalid"):
                return _prompt_delete_spec()
            raise BackRequested
    else:
        targets = [t.strip() for t in targets_raw.split(",") if t.strip()]

    strategy_choice = prompt_menu(
        "Matching strategy",
        [
            MenuChoice("exact", "Exact"),
            MenuChoice("ci", "Case-insensitive"),
            MenuChoice("contains", "Contains"),
            MenuChoice("startswith", "Starts with"),
            MenuChoice("endswith", "Ends with"),
            MenuChoice("regex", "Regular expression"),
        ],
    )

    all_sheets = typer.confirm("Apply to all sheets?", default=False)
    sheet_selector_raw = _prompt_optional_text("Specific sheet (leave blank for default)")
    sheet_selector: str | int | None
    if sheet_selector_raw is None:
        sheet_selector = None
    elif sheet_selector_raw.lower().startswith("index:"):
        _, _, rest = sheet_selector_raw.partition(":")
        if not rest.strip().isdigit():
            typer.secho("Sheet index specifier must be numeric, e.g. index:2", fg=typer.colors.RED)
            if _prompt_try_again("Delete columns configuration invalid"):
                return _prompt_delete_spec()
            raise BackRequested
        sheet_selector = int(rest.strip())
    elif sheet_selector_raw.isdigit():
        sheet_selector = int(sheet_selector_raw)
    else:
        sheet_selector = sheet_selector_raw

    inplace = typer.confirm("Modify files in-place?", default=False)
    on_missing_choice = prompt_menu(
        "Missing columns handling",
        [
            MenuChoice("ignore", "Ignore"),
            MenuChoice("error", "Error"),
        ],
    )
    dry_run = typer.confirm("Dry run (no files written)?", default=False)
    glob = _prompt_optional_text("Glob pattern (leave blank for none)")
    recursive = typer.confirm("Search recursively?", default=False)

    password, password_map = _prompt_password_inputs()

    if (not dry_run) and (not inplace):
        proceed = typer.confirm("Write cleaned copies next to originals?", default=True, abort=False)
        if not proceed:
            raise BackRequested

    spec = DeleteSpec(
        path=path,
        targets=targets,
        match_kind=match_choice.value,  # type: ignore[arg-type]
        strategy=strategy_choice.value,  # type: ignore[arg-type]
        all_sheets=all_sheets,
        sheet_selector=sheet_selector,
        inplace=inplace,
        on_missing=on_missing_choice.value,  # type: ignore[arg-type]
        dry_run=dry_run,
        glob=glob,
        recursive=recursive,
        password=password,
        password_map=password_map,
    )
    return spec


def _run_with_retry(name: str, builder: Callable[[], object], executor: Callable[[object], None]) -> None:
    while True:
        try:
            payload = builder()
        except BackRequested:
            return

        try:
            executor(payload)
        except ExcelMgrError as exc:
            typer.secho(str(exc), fg=typer.colors.RED)
            if not _prompt_try_again(f"{name} failed"):
                return
        except typer.BadParameter as exc:
            typer.secho(str(exc), fg=typer.colors.RED)
            if not _prompt_try_again(f"{name} failed"):
                return
        else:
            if not _prompt_try_again(f"{name} completed"):
                return


def _execute_combine(plan: CombinePlan) -> None:
    logger = _get_logger()
    hook = cli_main._make_progress_hook(logger)
    result = combine_command(plan, PandasReader(), PandasWriter(), progress_hooks=[hook])
    logger.info("combine_completed", **result)
    typer.echo(json.dumps(result, indent=2))


def _execute_split(plan: SplitPlan) -> None:
    logger = _get_logger()
    hook = cli_main._make_progress_hook(logger)
    result = split_command(plan, PandasReader(), PandasWriter(), progress_hooks=[hook])
    logger.info("split_completed", **result)
    typer.echo(json.dumps(result, indent=2))


def _execute_preview(plan: PreviewPlan) -> None:
    logger = _get_logger()
    result = preview_command(plan, PandasReader())
    logger.info("preview_completed", path=plan.path, sheets=len(result.get("sheets", [])))
    typer.echo(json.dumps(result, indent=2))


def _execute_delete(spec: DeleteSpec) -> None:
    logger = _get_logger()
    hook = cli_main._make_progress_hook(logger)
    result = delete_columns_command(spec, PandasReader(), PandasWriter(), progress_hooks=[hook])
    logger.info("delete_cols_completed", **result)
    typer.echo(json.dumps(result, indent=2))


def _execute_plan(path: str) -> None:
    logger = _get_logger()
    operations = load_plan_file(path)
    hook = cli_main._make_progress_hook(logger)
    results = execute_plan(operations, PandasReader(), PandasWriter(), progress_hooks=[hook])
    logger.info("plan_completed", operations=len(results))
    typer.echo(json.dumps({"operations": results}, indent=2))


def _run_plan_execution() -> None:
    def _builder() -> str:
        typer.echo("\nPlan execution (type 'back' at any prompt to return).")
        return _prompt_text("Plan file path")

    _run_with_retry("Plan execution", _builder, _execute_plan)


def _run_diagnostics() -> None:
    typer.echo("\nDiagnostics")
    cli_main.diagnose()


def _run_version() -> None:
    typer.echo("\nVersion")
    cli_main.version()


def main() -> None:
    typer.echo("Excel Manager interactive mode. Type 'back' while answering prompts to return to the previous menu.")

    actions: dict[str, Callable[[], None]] = {
        "combine": lambda: _run_with_retry("Combine", _prompt_combine_plan, _execute_combine),
        "split": lambda: _run_with_retry("Split", _prompt_split_plan, _execute_split),
        "preview": lambda: _run_with_retry("Preview", _prompt_preview_plan, _execute_preview),
        "delete": lambda: _run_with_retry("Delete columns", _prompt_delete_spec, _execute_delete),
        "plan": _run_plan_execution,
        "diagnostics": _run_diagnostics,
        "version": _run_version,
    }

    menu_options = [
        MenuChoice("combine", "Combine"),
        MenuChoice("split", "Split"),
        MenuChoice("preview", "Preview"),
        MenuChoice("delete", "Delete columns", aliases=("delete columns",)),
        MenuChoice("plan", "Plan execution", aliases=("plan execution",)),
        MenuChoice("diagnostics", "Diagnostics"),
        MenuChoice("version", "Version"),
        MenuChoice("exit", "Exit", aliases=("quit", "q")),
    ]

    while True:
        choice = prompt_menu("Main menu", menu_options)
        if choice.value == "exit":
            typer.echo("Goodbye!")
            return
        action = actions.get(choice.value)
        if action is not None:
            try:
                action()
            except BackRequested:
                continue

