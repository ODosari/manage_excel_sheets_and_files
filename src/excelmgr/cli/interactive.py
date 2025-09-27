"""Interactive entry point for the Excel Manager CLI with enhanced UX."""

from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import sys
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Sequence

import typer
from rich.progress import (
    BarColumn,
    Progress,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)

from excelmgr.adapters.pandas_io import PandasReader, PandasWriter
from excelmgr.cli import main as cli_main
from excelmgr.cli.options import read_password_map, read_secret
from excelmgr.core.combine import combine as combine_command
from excelmgr.core.delete_cols import delete_columns as delete_columns_command
from excelmgr.core.errors import ExcelMgrError
from excelmgr.core.models import CombinePlan, DeleteSpec, PreviewPlan, SheetSpec, SplitPlan
from excelmgr.core.passwords import resolve_password
from excelmgr.core.plan_runner import execute_plan, load_plan_file
from excelmgr.core.preview import preview as preview_command
from excelmgr.core.split import split as split_command
from excelmgr.util.text import write_text

try:
    import pyperclip  # type: ignore[import-not-found]
except ImportError:  # pragma: no cover - optional dependency
    pyperclip = None  # type: ignore[assignment]


class BackRequested(Exception):
    """Raised when the user requests to return to the previous menu."""


@dataclass
class MenuItem:
    key: str
    label: str
    aliases: tuple[str, ...] = ()


@dataclass
class OperationOutcome:
    result: dict
    paths: list[Path]
    dry_run: bool = False


ILLEGAL_FILENAME = re.compile(r'[\x00-\x1F\x7F<>:"/\\|?*]')
BACK_KEYWORDS = {"back", "b", ".."}


def _normalize_token(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", value)
    return normalized


def _simplify(value: str) -> str:
    nfkd = unicodedata.normalize("NFKD", value)
    stripped = "".join(ch for ch in nfkd if not unicodedata.combining(ch))
    return stripped.casefold()


def _parse_numeric_token(value: str) -> int:
    normalized = unicodedata.normalize("NFKC", value)
    digits: list[str] = []
    for ch in normalized:
        if ch.isdigit():
            try:
                digits.append(str(unicodedata.digit(ch)))
            except (TypeError, ValueError):
                digits.append(ch)
        elif ch in {"+", "-"} and not digits:
            digits.append(ch)
        else:
            raise ValueError
    if not digits:
        raise ValueError
    return int("".join(digits))


BACK_KEYWORDS_NORMALIZED = {_simplify(word) for word in BACK_KEYWORDS}


def _get_logger():
    return cli_main.app.state["logger"]


def _progress_events_enabled(logger) -> bool:
    toggle = bool(cli_main.app.state.get("interactive_show_events", False))
    if toggle:
        return True
    fmt = getattr(logger, "_fmt", "json")
    if fmt != "json":
        return True
    log_impl = getattr(logger, "_logger", None)
    if log_impl and any(isinstance(handler, logging.FileHandler) for handler in getattr(log_impl, "handlers", [])):
        return True
    return False


def _sanitize_token(token: str | None) -> str:
    if not token:
        return "data"
    normalized = _normalize_token(str(token))
    cleaned = ILLEGAL_FILENAME.sub("_", normalized)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"_+", "_", cleaned)
    cleaned = cleaned.strip("_")
    cleaned = unicodedata.normalize("NFC", cleaned)
    return cleaned or "data"


def _current_timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _find_data_root(reference: Path | None) -> Path:
    if reference is not None:
        ref = reference.resolve()
        search_candidates = []
        if ref.is_dir():
            search_candidates.append(ref)
        search_candidates.extend(ref.parents)
        for candidate in search_candidates:
            if candidate.name == "data":
                return candidate
    return (Path.cwd() / "data").resolve()


def _operation_out_dir(reference: Path | None, op: str) -> Path:
    root = _find_data_root(reference)
    out_dir = root / op / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def _build_name_tokens(*, base: str | None, src: Path | None, sheet: str | None, extra: str | None) -> list[str]:
    tokens: list[str] = []
    if base:
        tokens.append(base)
    elif src is not None:
        tokens.append(src.stem)
    if sheet:
        tokens.append(sheet)
    if extra:
        tokens.append(extra)

    cleaned: list[str] = []
    for token in tokens:
        sanitized = _sanitize_token(token)
        if not sanitized:
            continue
        if not cleaned or cleaned[-1] != sanitized:
            cleaned.append(sanitized)
    return cleaned


def _compose_filename(*, base: str | None, src: Path | None, sheet: str | None, extra: str | None, timestamp: str, suffix: str) -> str:
    tokens = _build_name_tokens(base=base, src=src, sheet=sheet, extra=extra)
    tokens.append(timestamp)
    stem = "__".join(tokens)
    return f"{stem}{suffix}"


def _preview_filename(*, base: str | None, src: Path | None, sheet: str | None, extra: str | None, timestamp_placeholder: str, suffix: str, extra_placeholder_display: str | None = None) -> str:
    tokens = _build_name_tokens(base=base, src=src, sheet=sheet, extra=extra)
    tokens.append(timestamp_placeholder)
    stem = "__".join(tokens)
    if extra and extra_placeholder_display:
        stem = stem.replace(_sanitize_token(extra), extra_placeholder_display)
    return f"{stem}{suffix}"


def _preview_output_destination(out_dir: Path, sample_name: str) -> None:
    typer.echo()
    typer.secho("Outputs will be written under:", fg=typer.colors.BLUE)
    typer.echo(f"  {out_dir.resolve()}")
    typer.echo(f"  e.g., {sample_name}")


def _common_output_dir(paths: Sequence[Path]) -> Path | None:
    if not paths:
        return None
    resolved = [path.resolve() for path in paths]
    roots = [p if p.is_dir() else p.parent for p in resolved]
    try:
        common = os.path.commonpath([str(root) for root in roots])
    except ValueError:
        return roots[0]
    return Path(common)


def _open_output_folder(paths: Sequence[Path]) -> None:
    target = _common_output_dir(paths)
    if target is None:
        typer.secho("No output paths available to open.", fg=typer.colors.YELLOW)
        return
    if not target.exists():
        typer.secho(f"Path does not exist yet: {target}", fg=typer.colors.YELLOW)
        return
    command: list[str]
    if sys.platform.startswith("darwin"):
        command = ["open", str(target)]
    elif os.name == "nt":
        command = ["cmd", "/c", "start", "", str(target)]
    else:
        command = ["xdg-open", str(target)]
    try:
        completed = subprocess.run(command, check=False)
        if completed.returncode == 0:
            typer.secho(f"Opened {target}", fg=typer.colors.GREEN)
        else:
            typer.secho(
                f"Open command exited with status {completed.returncode}.",
                fg=typer.colors.YELLOW,
            )
    except FileNotFoundError:
        typer.secho("Unable to open folder: command not found.", fg=typer.colors.RED)


def _copy_paths_to_clipboard(paths: Sequence[Path]) -> bool:
    if not pyperclip:
        return False
    try:
        joined = "\n".join(str(path.resolve()) for path in paths)
        pyperclip.copy(joined)
    except pyperclip.PyperclipException:  # type: ignore[attr-defined]
        return False
    return True


def _toggle_progress_logging() -> None:
    current = bool(cli_main.app.state.get("interactive_show_events", False))
    cli_main.app.state["interactive_show_events"] = not current
    status = "enabled" if not current else "disabled"
    typer.secho(f"JSON progress logging {status}.", fg=typer.colors.GREEN)
    if not current:
        typer.echo("Structured event output will appear on the next run.")


def prompt_menu(
    title: str,
    items: Sequence[MenuItem],
    *,
    show_exit: bool = False,
    context: str | None = None,
    back_label: str = "Back",
) -> str:
    while True:
        typer.echo()
        if context:
            typer.secho(context, fg=typer.colors.CYAN)
        typer.secho(title, bold=True)
        typer.echo(f"  0) {back_label}")
        for index, item in enumerate(items, start=1):
            typer.echo(f"  {index}) {item.label}")
        exit_index = None
        if show_exit:
            exit_index = len(items) + 1
            typer.echo(f"  {exit_index}) Exit")
        raw = typer.prompt("Select an option").strip()
        if not raw:
            typer.secho("Please choose an option.", fg=typer.colors.YELLOW)
            continue
        simplified = _simplify(raw)
        normalized_raw = _normalize_token(raw)
        if simplified in BACK_KEYWORDS_NORMALIZED or simplified == "0":
            return "back"
        if show_exit and exit_index is not None and simplified in {
            _simplify("exit"),
            _simplify("quit"),
            _simplify("q"),
        }:
            return "exit"
        if normalized_raw.isdigit():
            idx = _parse_numeric_token(normalized_raw)
            if idx == 0:
                return "back"
            if 1 <= idx <= len(items):
                return items[idx - 1].key
            if show_exit and exit_index is not None and idx == exit_index:
                return "exit"
        matches: list[str] = []
        for item in items:
            candidates = [
                _simplify(item.key),
                _simplify(item.label),
                *[_simplify(alias) for alias in item.aliases],
            ]
            if simplified in candidates:
                matches = [item.key]
                break
            if any(candidate.startswith(simplified) for candidate in candidates):
                matches.append(item.key)
        if len(dict.fromkeys(matches)) == 1:
            return matches[0]
        typer.secho("Invalid selection. Try again.", fg=typer.colors.RED)


def _ensure_not_back(value: str) -> str:
    if _simplify(value) in BACK_KEYWORDS_NORMALIZED:
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


def _prompt_confirm(message: str, *, default: bool = True) -> bool:
    suffix = "Y/n" if default else "y/N"
    while True:
        raw = typer.prompt(f"{message} [{suffix}]").strip()
        if not raw:
            return default
        lowered = raw.lower()
        if lowered in BACK_KEYWORDS:
            raise BackRequested
        if lowered in {"y", "yes"}:
            return True
        if lowered in {"n", "no"}:
            return False
        typer.secho("Please answer yes or no.", fg=typer.colors.YELLOW)


def _prompt_password_inputs(*, context: str | None = None) -> tuple[str | None, dict[str, str] | None]:
    items = [
        MenuItem("none", "No password (open)", aliases=("open",)),
        MenuItem("manual", "Type password manually"),
        MenuItem("env", "Read password from environment variable", aliases=("environment", "environment variable")),
        MenuItem("file", "Load password from file"),
        MenuItem("map", "Use password map file"),
    ]
    while True:
        choice = prompt_menu(
            "Choose a password source",
            items,
            context=context,
        )
        if choice == "back":
            raise BackRequested
        if choice == "none":
            typer.echo("Password: <none>")
            return None, None
        if choice == "manual":
            password = _prompt_text("Password", allow_empty=False)
            typer.echo("Password: typed manually")
            return password, None
        if choice == "env":
            env_name = _prompt_text("Environment variable name", allow_empty=False)
            try:
                resolved = read_secret(None, env_name, None)
            except typer.BadParameter as exc:
                typer.secho(str(exc), fg=typer.colors.RED)
                continue
            if resolved is None:
                typer.secho(f"Environment variable {env_name} is not set.", fg=typer.colors.YELLOW)
            typer.echo(f"Password: environment variable {env_name}")
            return resolved, None
        if choice == "file":
            file_path = Path(_prompt_text("Path to password file", allow_empty=False)).expanduser()
            try:
                resolved = read_secret(None, None, str(file_path))
            except typer.BadParameter as exc:
                typer.secho(str(exc), fg=typer.colors.RED)
                continue
            typer.echo(f"Password: file {file_path.resolve()}")
            return resolved, None
        if choice == "map":
            map_path = Path(_prompt_text("Path to password map file", allow_empty=False)).expanduser()
            try:
                password_map = read_password_map(str(map_path))
            except typer.BadParameter as exc:
                typer.secho(str(exc), fg=typer.colors.RED)
                continue
            typer.echo(f"Password map: {map_path.resolve()}")
            return None, password_map


def _resolve_password(path: str, password: str | None, password_map: dict[str, str] | None) -> str | None:
    try:
        return resolve_password(path, password, password_map)
    except ExcelMgrError as exc:
        typer.secho(str(exc), fg=typer.colors.RED)
        raise


def _list_sheet_names(path: str, password: str | None) -> list[str]:
    reader = PandasReader()
    return reader.sheet_names(path, password)


def _parse_index_list(raw: str, total: int) -> list[int]:
    result: list[int] = []
    for chunk in raw.split(","):
        token = chunk.strip()
        if not token:
            continue
        if "-" in token:
            start_str, _, end_str = token.partition("-")
            start = _parse_numeric_token(start_str.strip())
            end = _parse_numeric_token(end_str.strip())
            if start < 1 or end < 1 or end < start:
                raise ValueError
            for value in range(start, end + 1):
                if value > total:
                    raise ValueError
                result.append(value)
        else:
            index = _parse_numeric_token(token)
            if not (1 <= index <= total):
                raise ValueError
            result.append(index)
    deduped: list[int] = []
    for idx in result:
        if idx not in deduped:
            deduped.append(idx)
    if not deduped:
        raise ValueError
    return deduped


def pick_sheets(path: str, *, password: str | None, allow_multi: bool, allow_all: bool = True) -> list[str] | str:
    names = _list_sheet_names(path, password)
    if not names:
        raise ExcelMgrError("No sheets found in the workbook.")
    if len(names) == 1:
        only = names[0]
        if _prompt_confirm(f'Found 1 sheet (index 1): "{only}". Use this sheet?', default=True):
            return [only]
    while True:
        typer.echo()
        typer.secho(f"Sheets in {Path(path).name}", bold=True)
        typer.echo("  0) Back")
        for index, name in enumerate(names, start=1):
            typer.echo(f"  {index}) {name}")
        suffix = " (comma/range or 'all' allowed)" if allow_multi and allow_all else ""
        raw_input = typer.prompt(f"Select a sheet by number{suffix}").strip()
        simplified = _simplify(raw_input)
        if simplified in BACK_KEYWORDS_NORMALIZED or simplified == "0":
            raise BackRequested
        if allow_multi and allow_all and simplified == _simplify("all"):
            return "__ALL__"
        try:
            normalized_input = _normalize_token(raw_input)
            if allow_multi and any(sep in raw_input for sep in {",", "-"}):
                indexes = _parse_index_list(normalized_input.lower(), len(names))
                selected = [names[i - 1] for i in indexes]
                typer.echo("Selected sheets: " + ", ".join(selected))
                return selected
            index = _parse_numeric_token(normalized_input)
        except (ValueError, TypeError):
            matches = [
                name
                for name in names
                if _simplify(name).startswith(simplified)
            ]
            if len(matches) == 1:
                return [matches[0]]
            typer.secho("Invalid selection. Try again.", fg=typer.colors.RED)
            continue
        if 1 <= index <= len(names):
            return [names[index - 1]]
        typer.secho("Invalid selection. Try again.", fg=typer.colors.RED)


class VisualProgress:
    stages = ["Reading", "Planning", "Executing", "Writing"]

    def __init__(self, logger) -> None:
        self._logger = logger
        self._hook = cli_main._make_progress_hook(logger)
        self._show_events = _progress_events_enabled(logger)
        self._progress = Progress(
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TimeElapsedColumn(),
            TimeRemainingColumn(),
            transient=True,
        )
        self._stage_task: int | None = None
        self._stage_index = -1
        self._partition_label = ""
        self._partition_count = 0
        self._partition_total: int | None = None

    def __enter__(self) -> "VisualProgress":
        self._progress.__enter__()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._stage_task is not None:
            self._progress.update(self._stage_task, completed=len(self.stages))
        self._progress.__exit__(exc_type, exc, tb)

    def _ensure_task(self) -> None:
        if self._stage_task is None:
            self._stage_task = self._progress.add_task(self.stages[0], total=len(self.stages))

    def _update_description(self, stage: str) -> None:
        description = stage
        if self._partition_label:
            description = f"{stage} — {self._partition_label}"
        if self._stage_task is not None:
            self._progress.update(self._stage_task, description=description)

    def _advance_stage(self, stage: str) -> None:
        self._ensure_task()
        index = self.stages.index(stage)
        if index <= self._stage_index:
            self._update_description(stage)
            return
        self._stage_index = index
        if self._stage_task is not None:
            self._progress.update(self._stage_task, completed=index + 1)
        self._update_description(stage)

    def hook(self, event: str, payload: dict[str, object]) -> None:
        if self._show_events:
            self._hook(event, payload)
        if event.endswith("_start"):
            self._advance_stage("Reading")
            self._advance_stage("Planning")
        elif event.endswith("_partition") or event in {"combine_file", "combine_sheet", "delete_sheet", "delete_workbook"}:
            self._advance_stage("Executing")
            self._partition_count += 1
            if "partitions" in payload and isinstance(payload["partitions"], int):
                self._partition_total = int(payload["partitions"])
            elif "index" in payload and isinstance(payload["index"], int):
                self._partition_total = max(self._partition_total or 0, int(payload["index"]))
            if "total" in payload and isinstance(payload["total"], int):
                self._partition_total = int(payload["total"])
            label = f"Partition {self._partition_count}"
            if self._partition_total:
                label = f"Partition {self._partition_count}/{self._partition_total}"
            self._partition_label = label
            self._update_description("Executing")
        elif event.endswith("_complete"):
            if "partitions" in payload and isinstance(payload["partitions"], int):
                self._partition_total = int(payload["partitions"])
                self._partition_label = f"Partition {self._partition_total}/{self._partition_total}"
            elif "files" in payload and isinstance(payload["files"], int):
                total = int(payload["files"])
                self._partition_label = f"Partition {total}/{total}"
            elif "sheets" in payload and isinstance(payload["sheets"], int):
                total = int(payload["sheets"])
                self._partition_label = f"Partition {total}/{total}"
            self._advance_stage("Writing")

    def manual_cycle(self) -> None:
        self._advance_stage("Reading")
        self._advance_stage("Planning")
        self._advance_stage("Executing")
        self._advance_stage("Writing")


@dataclass
class CombinePayload:
    plan: CombinePlan
    output_path: Path


@dataclass
class SplitPayload:
    plan: SplitPlan
    source_path: Path
    sheet_name: str | None
    output_dir: Path
    timestamp: str


@dataclass
class PreviewPayload:
    plan: PreviewPlan
    output_path: Path


@dataclass
class DeletePayload:
    spec: DeleteSpec
    source_paths: list[Path]
    output_dir: Path
    timestamp: str


@dataclass
class PlanPayload:
    plan_path: str
    output_path: Path


def _prompt_paths_list(message: str) -> list[str]:
    while True:
        raw = _prompt_text(message)
        entries = [entry.strip() for entry in raw.split(",") if entry.strip()]
        if entries:
            return entries
        typer.secho("Provide at least one entry.", fg=typer.colors.RED)


def _sample_path(paths: Sequence[str]) -> Path | None:
    for raw in paths:
        candidate = Path(raw).expanduser()
        if candidate.is_file():
            return candidate
        if candidate.is_dir():
            files = list(candidate.glob("*.xlsx"))
            if files:
                return files[0]
    return None


def _prompt_combine_plan() -> CombinePayload:
    typer.secho("\nCombine workflow", bold=True)
    inputs = _prompt_paths_list("Input files or directories (comma separated)")
    mode_choice = prompt_menu(
        "Select combine mode",
        [
            MenuItem("one_sheet", "One sheet"),
            MenuItem("multi_sheets", "Multi sheets"),
        ],
        context="Combine ▸ Mode",
    )
    glob_pattern = _prompt_optional_text("Glob pattern (leave blank for none)")
    recursive = _prompt_confirm("Search recursively?", default=False)
    password, password_map = _prompt_password_inputs(context="Combine ▸ Password")
    sample = _sample_path(inputs)
    include: Sequence[SheetSpec] | str
    if sample is not None:
        pw = _resolve_password(str(sample), password, password_map)
        selection = pick_sheets(str(sample), password=pw, allow_multi=True)
        if selection == "__ALL__":
            include = "all"
        else:
            include = [SheetSpec(name) for name in selection]  # type: ignore[arg-type]
    else:
        include = "all"
    output_format_choice = prompt_menu(
        "Select output format",
        [
            MenuItem("xlsx", "Excel (.xlsx)"),
            MenuItem("csv", "CSV"),
            MenuItem("parquet", "Parquet"),
        ],
        context="Combine ▸ Format",
    )
    add_bom = False
    if output_format_choice == "csv":
        add_bom = _prompt_confirm("Add UTF-8 BOM for compatibility?", default=False)
    if mode_choice == "multi_sheets" and output_format_choice != "xlsx":
        typer.secho("Multi-sheet mode requires Excel output format.", fg=typer.colors.RED)
        raise BackRequested
    add_source = _prompt_confirm("Add source column?", default=False)
    sheet_name = _prompt_text("Output sheet name", default="Data")
    dry_run = _prompt_confirm("Dry run (no files written)?", default=False)
    reference = sample or (Path(inputs[0]).expanduser() if inputs else None)
    out_dir = _operation_out_dir(reference, "combine")
    timestamp = _current_timestamp()
    suffix_map = {"xlsx": ".xlsx", "csv": ".csv", "parquet": ".parquet"}
    suffix = suffix_map[output_format_choice]
    base_name = "Combined" if (not sample or not sample.is_file() or len(inputs) > 1) else sample.stem
    filename = _compose_filename(base=base_name, src=sample, sheet=None, extra=None, timestamp=timestamp, suffix=suffix)
    output_path = out_dir / filename
    _preview_output_destination(out_dir, filename)
    plan = CombinePlan(
        inputs=inputs,
        glob=glob_pattern,
        recursive=recursive,
        mode=mode_choice,  # type: ignore[arg-type]
        include_sheets=include,  # type: ignore[arg-type]
        output_path=str(output_path),
        output_sheet_name=sheet_name,
        add_source_column=add_source,
        password=password,
        password_map=password_map,
        output_format=output_format_choice,  # type: ignore[arg-type]
        csv_add_bom=add_bom,
        dry_run=dry_run,
    )
    return CombinePayload(plan=plan, output_path=output_path)


def _prompt_split_plan() -> SplitPayload:
    typer.secho("\nSplit workflow", bold=True)
    input_file = _prompt_text("Input workbook path")
    password, password_map = _prompt_password_inputs(context="Split ▸ Password")
    pw = _resolve_password(input_file, password, password_map)
    selection = pick_sheets(input_file, password=pw, allow_multi=False, allow_all=False)
    sheet_name = selection[0]
    sheet_spec = SheetSpec(sheet_name)
    column = _prompt_text("Column to split by")
    destination_choice = prompt_menu(
        "Split destination",
        [
            MenuItem("files", "Create files (one file per partition)"),
            MenuItem("sheets", "Create sheets (one sheet per partition in a single file)"),
        ],
        context="Split ▸ Destination",
    )
    include_nan = _prompt_confirm("Include rows with missing split values?", default=False)
    output_format_choice = prompt_menu(
        "Select output format",
        [
            MenuItem("xlsx", "Excel (.xlsx)"),
            MenuItem("csv", "CSV"),
            MenuItem("parquet", "Parquet"),
        ],
        context="Split ▸ Format",
    )
    add_bom = False
    if output_format_choice == "csv":
        add_bom = _prompt_confirm("Add UTF-8 BOM for compatibility?", default=False)
    if destination_choice == "sheets" and output_format_choice != "xlsx":
        typer.secho("Sheet destination requires Excel output format.", fg=typer.colors.RED)
        raise BackRequested
    dry_run = _prompt_confirm("Dry run (no files written)?", default=False)
    out_dir = _operation_out_dir(Path(input_file), "split")
    timestamp = _current_timestamp()
    output_filename: str | None = None
    suffix_map = {"xlsx": ".xlsx", "csv": ".csv", "parquet": ".parquet"}
    suffix = suffix_map[output_format_choice]
    if destination_choice == "sheets":
        output_filename = _compose_filename(
            base=None,
            src=Path(input_file),
            sheet=sheet_name,
            extra="split",
            timestamp=timestamp,
            suffix=".xlsx",
        )
        _preview_output_destination(out_dir, output_filename)
    plan = SplitPlan(
        input_file=input_file,
        sheet=sheet_spec,
        by_column=column,
        to=destination_choice,  # type: ignore[arg-type]
        include_nan=include_nan,
        output_dir=str(out_dir),
        output_filename=output_filename,
        output_sheet_name="Data",
        password=password,
        password_map=password_map,
        output_format=output_format_choice,  # type: ignore[arg-type]
        csv_add_bom=add_bom,
        dry_run=dry_run,
    )
    if destination_choice == "files":
        sample = _preview_filename(
            base=None,
            src=Path(input_file),
            sheet=sheet_name,
            extra="PARTITION",
            timestamp_placeholder="YYYYMMDD_HHMMSS",
            suffix=suffix,
            extra_placeholder_display="<partition>",
        )
        _preview_output_destination(out_dir, sample)
    return SplitPayload(plan=plan, source_path=Path(input_file), sheet_name=sheet_name, output_dir=out_dir, timestamp=timestamp)


def _prompt_preview_plan() -> PreviewPayload:
    typer.secho("\nPreview workflow", bold=True)
    path = _prompt_text("Workbook path")
    limit_raw = _prompt_optional_text("Sample row limit per sheet (leave blank for unlimited)")
    if limit_raw is not None and not limit_raw.isdigit():
        typer.secho("Limit must be a positive integer.", fg=typer.colors.RED)
        raise BackRequested
    limit = int(limit_raw) if limit_raw else None
    password, password_map = _prompt_password_inputs(context="Preview ▸ Password")
    plan = PreviewPlan(path=path, password=password, password_map=password_map, limit=limit)
    out_dir = _operation_out_dir(Path(path), "preview")
    timestamp = _current_timestamp()
    filename = _compose_filename(
        base=Path(path).stem,
        src=Path(path),
        sheet=None,
        extra="preview",
        timestamp=timestamp,
        suffix=".json",
    )
    output_path = out_dir / filename
    _preview_output_destination(out_dir, filename)
    return PreviewPayload(plan=plan, output_path=output_path)


def _prompt_delete_spec() -> DeletePayload:
    typer.secho("\nDelete columns workflow", bold=True)
    path = _prompt_text("Workbook path or directory")
    match_choice = prompt_menu(
        "Match columns by",
        [MenuItem("names", "Names"), MenuItem("index", "Index")],
        context="Delete ▸ Match by",
    )
    targets_raw = _prompt_text("Columns to delete (comma separated)")
    if match_choice == "index":
        try:
            targets = [int(item.strip()) for item in targets_raw.split(",") if item.strip()]
        except ValueError:
            typer.secho("Targets must be integers when matching by index.", fg=typer.colors.RED)
            raise BackRequested
    else:
        targets = [item.strip() for item in targets_raw.split(",") if item.strip()]
    strategy = prompt_menu(
        "Matching strategy",
        [
            MenuItem("exact", "Exact"),
            MenuItem("ci", "Case-insensitive"),
            MenuItem("contains", "Contains"),
            MenuItem("startswith", "Starts with"),
            MenuItem("endswith", "Ends with"),
            MenuItem("regex", "Regular expression"),
        ],
        context="Delete ▸ Strategy",
    )
    all_sheets = _prompt_confirm("Apply to all sheets?", default=False)
    sheet_selector: str | int | None = None
    if not all_sheets:
        sheet_choice = _prompt_optional_text("Specific sheet (leave blank for default)")
        if sheet_choice:
            if sheet_choice.isdigit():
                sheet_selector = int(sheet_choice)
            elif sheet_choice.lower().startswith("index:"):
                _, _, rest = sheet_choice.partition(":")
                rest = rest.strip()
                if not rest.isdigit():
                    typer.secho("Sheet index specifier must be numeric, e.g. index:2", fg=typer.colors.RED)
                    raise BackRequested
                sheet_selector = int(rest)
            else:
                sheet_selector = sheet_choice
    inplace = _prompt_confirm("Modify files in-place?", default=False)
    on_missing = prompt_menu(
        "Missing columns handling",
        [MenuItem("ignore", "Ignore"), MenuItem("error", "Error")],
        context="Delete ▸ Missing columns",
    )
    dry_run = _prompt_confirm("Dry run (no files written)?", default=False)
    glob_pattern = _prompt_optional_text("Glob pattern (leave blank for none)")
    recursive = _prompt_confirm("Search recursively?", default=False)
    password, password_map = _prompt_password_inputs(context="Delete ▸ Password")
    spec = DeleteSpec(
        path=path,
        targets=targets,  # type: ignore[list-item]
        match_kind=match_choice,  # type: ignore[arg-type]
        strategy=strategy,  # type: ignore[arg-type]
        all_sheets=all_sheets,
        sheet_selector=sheet_selector,
        inplace=inplace,
        on_missing=on_missing,  # type: ignore[arg-type]
        dry_run=dry_run,
        glob=glob_pattern,
        recursive=recursive,
        password=password,
        password_map=password_map,
    )
    if not inplace and not dry_run:
        _prompt_confirm("Write cleaned copies under managed output directory?", default=True)
    base_path = Path(path).expanduser()
    if base_path.is_file():
        sources = [base_path]
    elif base_path.is_dir():
        sources = [base_path]
    else:
        sources = []
    out_dir = _operation_out_dir(base_path if sources else None, "delete")
    timestamp = _current_timestamp()
    if not inplace:
        sample_src = sources[0] if sources else base_path
        suffix = sample_src.suffix if sample_src and sample_src.suffix else ".xlsx"
        sample_name = _preview_filename(
            base=None,
            src=sample_src if sample_src and sample_src.exists() else base_path,
            sheet=None,
            extra="cleaned",
            timestamp_placeholder="YYYYMMDD_HHMMSS",
            suffix=suffix,
            extra_placeholder_display=None,
        )
        _preview_output_destination(out_dir, sample_name)
    return DeletePayload(spec=spec, source_paths=sources, output_dir=out_dir, timestamp=timestamp)


def _prompt_plan_payload() -> PlanPayload:
    typer.secho("\nPlan execution", bold=True)
    path = _prompt_text("Plan file path")
    out_dir = _operation_out_dir(Path(path), "plan")
    timestamp = _current_timestamp()
    filename = _compose_filename(base=Path(path).stem, src=Path(path), sheet=None, extra="plan", timestamp=timestamp, suffix=".json")
    return PlanPayload(plan_path=path, output_path=out_dir / filename)


def _execute_combine(payload: CombinePayload) -> OperationOutcome:
    logger = _get_logger()
    reader = PandasReader()
    writer = PandasWriter()
    with VisualProgress(logger) as progress:
        result = combine_command(
            payload.plan,
            reader,
            writer,
            progress_hooks=[progress.hook],
        )
    logger.info("combine_completed", **result)
    paths: list[Path] = []
    if not payload.plan.dry_run:
        paths = [payload.output_path.resolve()]
    return OperationOutcome(result=result, paths=paths, dry_run=payload.plan.dry_run)


def _execute_split(payload: SplitPayload) -> OperationOutcome:
    logger = _get_logger()
    reader = PandasReader()
    writer = PandasWriter()
    with VisualProgress(logger) as progress:
        result = split_command(
            payload.plan,
            reader,
            writer,
            progress_hooks=[progress.hook],
        )
    if payload.plan.dry_run:
        logger.info("split_completed", **result)
        return OperationOutcome(result=result, paths=[], dry_run=True)
    paths: list[Path] = []
    if payload.plan.to == "sheets":
        out_path = Path(result.get("out", payload.plan.output_dir)).resolve()
        final_path = payload.output_dir / _compose_filename(
            base=None,
            src=payload.source_path,
            sheet=payload.sheet_name,
            extra="split",
            timestamp=payload.timestamp,
            suffix=out_path.suffix,
        )
        if out_path != final_path:
            out_path.replace(final_path)
        progress.hook(
            "split_partition_finalized",
            {"original": str(out_path), "final": str(final_path)},
        )
        result["out"] = str(final_path)
        paths.append(final_path.resolve())
        progress.hook(
            "split_outputs_finalized",
            {"outputs": [str(final_path)], "output_dir": str(payload.output_dir.resolve())},
        )
    else:
        renamed: list[str] = []
        for original in result.get("outputs", []):
            orig_path = Path(original).resolve()
            suffix = orig_path.suffix or {
                "xlsx": ".xlsx",
                "csv": ".csv",
                "parquet": ".parquet",
            }.get(payload.plan.output_format, "")
            final = payload.output_dir / _compose_filename(
                base=None,
                src=payload.source_path,
                sheet=payload.sheet_name,
                extra=orig_path.stem,
                timestamp=payload.timestamp,
                suffix=suffix,
            )
            if orig_path != final:
                final.parent.mkdir(parents=True, exist_ok=True)
                orig_path.replace(final)
                progress.hook(
                    "split_partition_finalized",
                    {"original": str(orig_path), "final": str(final)},
                )
            else:
                progress.hook(
                    "split_partition_finalized",
                    {"original": str(orig_path), "final": str(final)},
                )
            renamed.append(str(final))
            paths.append(final.resolve())
        if renamed:
            result["outputs"] = renamed
        progress.hook(
            "split_outputs_finalized",
            {"outputs": result.get("outputs", []), "output_dir": str(payload.output_dir.resolve())},
        )
    logger.info("split_completed", **result)
    return OperationOutcome(result=result, paths=paths, dry_run=False)


def _execute_preview(payload: PreviewPayload) -> OperationOutcome:
    logger = _get_logger()
    with VisualProgress(logger) as progress:
        progress.manual_cycle()
        result = preview_command(payload.plan, PandasReader())
    logger.info("preview_completed", path=payload.plan.path, sheets=len(result.get("sheets", [])))
    write_text(payload.output_path, json.dumps(result, indent=2, ensure_ascii=False))
    return OperationOutcome(result=result, paths=[payload.output_path.resolve()], dry_run=False)


def _execute_delete(payload: DeletePayload) -> OperationOutcome:
    logger = _get_logger()
    reader = PandasReader()
    writer = PandasWriter()
    with VisualProgress(logger) as progress:
        result = delete_columns_command(
            payload.spec,
            reader,
            writer,
            progress_hooks=[progress.hook],
        )
    if payload.spec.dry_run or payload.spec.inplace:
        logger.info("delete_cols_completed", **result)
        return OperationOutcome(result=result, paths=[], dry_run=payload.spec.dry_run)
    final_paths: list[Path] = []
    renamed_items: list[dict] = []
    for item in result.get("items", []):
        out_path = item.get("out")
        if not out_path:
            renamed_items.append(item)
            continue
        orig = Path(out_path).resolve()
        src = Path(item.get("path", payload.spec.path)).resolve()
        final = payload.output_dir / _compose_filename(
            base=None,
            src=src,
            sheet=None,
            extra="cleaned",
            timestamp=payload.timestamp,
            suffix=orig.suffix,
        )
        if orig != final:
            final.parent.mkdir(parents=True, exist_ok=True)
            orig.replace(final)
        new_item = dict(item)
        new_item["out"] = str(final)
        renamed_items.append(new_item)
        final_paths.append(final.resolve())
    if renamed_items:
        result["items"] = renamed_items
    logger.info("delete_cols_completed", **result)
    return OperationOutcome(result=result, paths=final_paths, dry_run=False)


def _execute_plan(payload: PlanPayload) -> OperationOutcome:
    logger = _get_logger()
    operations = load_plan_file(payload.plan_path)
    with VisualProgress(logger) as progress:
        results = execute_plan(
            operations,
            PandasReader(),
            PandasWriter(),
            progress_hooks=[progress.hook],
        )
    summary = {"operations": results}
    logger.info("plan_completed", operations=len(results))
    write_text(payload.output_path, json.dumps(summary, indent=2, ensure_ascii=False))
    return OperationOutcome(result=summary, paths=[payload.output_path.resolve()], dry_run=False)


def _show_outcome(name: str, outcome: OperationOutcome) -> None:
    typer.echo()
    header = f"✅ {name} complete" if not outcome.dry_run else f"ℹ️ {name} dry run complete"
    typer.secho(header, bold=True, fg=typer.colors.GREEN if not outcome.dry_run else typer.colors.BLUE)
    if outcome.paths:
        typer.echo()
        typer.echo("Wrote:")
        for path in outcome.paths:
            typer.echo(f"  • {path}")
    elif outcome.dry_run:
        typer.echo("No files written (dry run).")
    typer.echo()
    typer.echo(json.dumps(outcome.result, indent=2, ensure_ascii=False))


def _prompt_next_action(message: str, *, outcome: OperationOutcome | None = None) -> str:
    items: list[MenuItem] = [MenuItem("again", "Run again")]
    if outcome and outcome.paths:
        items.append(MenuItem("open", "Open output folder"))
        if pyperclip:
            items.append(MenuItem("copy", "Copy output paths to clipboard"))
    items.append(MenuItem("main", "Back to main menu"))
    choice = prompt_menu(message, items, back_label="Main menu")
    if choice == "back":
        return "main"
    return choice


def _run_with_retry(name: str, builder: Callable[[], object], executor: Callable[[object], OperationOutcome]) -> None:
    while True:
        try:
            payload = builder()
        except BackRequested:
            return
        try:
            outcome = executor(payload)
        except ExcelMgrError as exc:
            typer.secho(str(exc), fg=typer.colors.RED)
            action = _prompt_next_action(f"{name} failed. What next?")
            if action != "again":
                return
        except typer.BadParameter as exc:
            typer.secho(str(exc), fg=typer.colors.RED)
            action = _prompt_next_action(f"{name} failed. What next?")
            if action != "again":
                return
        else:
            _show_outcome(name, outcome)
            while True:
                action = _prompt_next_action("What next?", outcome=outcome)
                if action == "again":
                    break
                if action == "open":
                    _open_output_folder(outcome.paths)
                    continue
                if action == "copy":
                    if _copy_paths_to_clipboard(outcome.paths):
                        typer.secho("Output paths copied to clipboard.", fg=typer.colors.GREEN)
                    else:
                        typer.secho("Clipboard copy unavailable.", fg=typer.colors.YELLOW)
                    continue
                return


def _run_plan_execution() -> None:
    _run_with_retry("Plan execution", _prompt_plan_payload, _execute_plan)


def _run_diagnostics() -> None:
    typer.echo("\nDiagnostics")
    cli_main.diagnose()


def _run_version() -> None:
    typer.echo("\nVersion")
    cli_main.version()


def run_interactive() -> None:
    state = cli_main.app.state
    state.setdefault("interactive_show_events", False)
    if not state.get("shown_welcome", False):
        typer.echo("Welcome to Excel Manager — a guided CLI for combining, splitting, previewing, and fixing Excel files.")
        typer.echo("Use the numbered menus; type “back” anytime to return to the previous step.")
        state["shown_welcome"] = True
    actions: dict[str, Callable[[], None]] = {
        "combine": lambda: _run_with_retry("Combine", _prompt_combine_plan, _execute_combine),
        "split": lambda: _run_with_retry("Split", _prompt_split_plan, _execute_split),
        "preview": lambda: _run_with_retry("Preview", _prompt_preview_plan, _execute_preview),
        "delete": lambda: _run_with_retry("Delete columns", _prompt_delete_spec, _execute_delete),
        "plan": _run_plan_execution,
        "diagnostics": _run_diagnostics,
        "version": _run_version,
        "toggle_logs": _toggle_progress_logging,
    }
    base_menu = [
        MenuItem("combine", "Combine"),
        MenuItem("split", "Split"),
        MenuItem("preview", "Preview"),
        MenuItem("delete", "Delete columns"),
        MenuItem("plan", "Plan execution"),
        MenuItem("diagnostics", "Diagnostics"),
        MenuItem("version", "Version"),
    ]
    while True:
        show_events = bool(cli_main.app.state.get("interactive_show_events", False))
        toggle_label = "Show JSON progress log" if not show_events else "Hide JSON progress log"
        menu_items = [*base_menu, MenuItem("toggle_logs", toggle_label)]
        choice = prompt_menu(
            "Select an action",
            menu_items,
            show_exit=True,
            context="Excel Manager — interactive mode",
            back_label="Refresh",
        )
        if choice in {"back", "exit"}:
            if choice == "exit":
                typer.echo("Goodbye!")
                return
            continue
        action = actions.get(choice)
        if action is None:
            typer.secho("Invalid selection. Try again.", fg=typer.colors.RED)
            continue
        try:
            action()
        except BackRequested:
            continue


def main() -> None:
    run_interactive()

