# excelmgr

Enterprise-grade CLI to **combine** Excel files, **split** sheets by a column, and **delete columns** across files/sheets. Secure, typed, with structured logging, config, tests, and CI.

## Install
```bash
python -m venv .venv && source .venv/bin/activate
pip install -U pip
pip install .[dev]
```

## Commands
## Interactive mode (default)
Run `excelmgr` with no arguments to launch the interactive menu:

```bash
excelmgr
```

This guided mode mirrors all existing commands (combine, split, preview, delete columns, diagnostics, version), reuses the same
validators, and honors environment/configuration and password handling. Non-interactive usage remains available for scripting,
for example `excelmgr combine ...`.

### Combine
```bash
excelmgr combine ./data --mode one-sheet --glob "*.xlsx,*.xlsm" --recursive --out combined.xlsx --add-source-column
```
### Split
```bash
excelmgr split ./input.xlsx --sheet Data --by Customer --to files --out out_dir
```
### Delete columns
```bash
excelmgr delete-cols ./data --targets "Notes,CustomerID" --match names --strategy ci --all-sheets --inplace --yes
# Index mode (1-based):
excelmgr delete-cols ./file.xlsx --targets "1,3,7" --match index --sheet Data --dry-run
```
`--inplace` edits the original workbook. Omit it to create `*.cleaned.xlsx` siblings. `--yes` skips the safety prompt—leave it
off to confirm before files are written.

## Config
Defaults are read from environment variables (prefix `EXCELMGR_`) or a `.env` file. CLI flags always override the
environment so you can temporarily change behavior without editing configuration files.
- `EXCELMGR_GLOB="*.xlsx,*.xlsm"`
- `EXCELMGR_RECURSIVE=false`
- `EXCELMGR_LOG="json"`
- `EXCELMGR_LOG_LEVEL="INFO"`
- `EXCELMGR_MACRO_POLICY="warn"`  # warn|forbid|ignore
- `EXCELMGR_TEMP_DIR` sets the directory used for temporary files (defaults next to the destination workbook)

## Security & passwords
- Use `--password-env` or `--password-file` over `--password` to avoid shell history leak.
- Encrypted workbooks require `msoffcrypto-tool` to decrypt to a temp stream before reading.
- Install it via `pip install msoffcrypto-tool` when working with password-protected files.
- Logs never include cell data; only shapes and counts.

## Macro safety
Writing `.xlsm` drops macros with pandas/openpyxl. Policy controlled via `EXCELMGR_MACRO_POLICY`:
- `warn` (default): log a warning.
- `forbid`: refuse to write `.xlsm` outputs.
- `ignore`: allow it without warnings.

## Logging
Choose log format and destination:
```bash
excelmgr --log text --log-level INFO --log-file excelmgr.log combine ./data --out combined.xlsx
```

## Encoding
- All text I/O defaults to UTF-8. Readers transparently accept UTF-8 with or without BOMs so headers such as `Müşteri Adı` or `年` stay intact.
- Interactive mode preserves Unicode in filenames, sheet names, column headers, previews, and JSON logs (no `\uXXXX` escaping).
- When exporting CSV/TSV, files are written as UTF-8 without a BOM by default. Enable the optional **“Add UTF-8 BOM for compatibility?”** toggle in the format menu if a legacy consumer requires it.

## Exit codes
- `0` success
- `2` known error (e.g., sheet not found, decryption)
- `1` unexpected crash

## Development
- Lint: `ruff check .`
- Type-check: `mypy .`
- Tests: `pytest -q`

## CI
GitHub Actions runs lint + type + tests and builds a wheel.

## License
MIT
