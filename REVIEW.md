# Review Notes

## Suggested Enhancements

1. **Expose CSV BOM toggle in CLI**
   - The execution paths already honor a `csv_add_bom` flag on `CombinePlan` and `SplitPlan`, but there is no CLI option to toggle it. Adding something like `--csv-add-bom/--no-csv-add-bom` would let users match the feature that the plan runner supports.
2. **Normalize log-level configuration**
   - `ExcelMgrSettings.log_level` is declared as a plain string and the CLI only accepts uppercase values. Automatically uppercasing config/env inputs before validation would make `excelmgr --log-level info` or `EXCELMGR_LOG_LEVEL=info` work without surprising errors.
3. **Implement database destinations for splits**
   - `SplitPlan.destination` allows a `DatabaseDestination`, but `split()` currently rejects it outright. Supporting writing partition results to a table (append/replace) would make split plans feature-parity with combine.

