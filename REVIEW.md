# Review Notes

## Suggested Enhancements

1. **Reuse the interactive column picker for combine plans**
   - The new `pick_column` helper is only wired into `_prompt_split_plan`, so the combine workflow still lacks any column listing or numeric selection even though the UX request explicitly called for it. Consider invoking the same picker when building `CombinePlan`s so users can choose their combine-by field via the numbered menu instead of free text. 【F:src/excelmgr/cli/interactive.py†L905-L919】【F:src/excelmgr/cli/interactive.py†L931-L937】

2. **Extend CombinePlan to capture the chosen column**
   - `CombinePlan` currently has no slot for a combine-by column, making it impossible for both the CLI and plan-runner to persist or honor the user’s choice even if the picker were presented. Adding a `by_column` attribute (mirroring `SplitPlan`) would let the executor honor the same selection path in both interactive and non-interactive modes. 【F:src/excelmgr/core/models.py†L34-L48】
1. **Expose CSV BOM toggle in CLI**
   - The execution paths already honor a `csv_add_bom` flag on `CombinePlan` and `SplitPlan`, but there is no CLI option to toggle it. Adding something like `--csv-add-bom/--no-csv-add-bom` would let users match the feature that the plan runner supports.
2. **Normalize log-level configuration**
   - `ExcelMgrSettings.log_level` is declared as a plain string and the CLI only accepts uppercase values. Automatically uppercasing config/env inputs before validation would make `excelmgr --log-level info` or `EXCELMGR_LOG_LEVEL=info` work without surprising errors.
3. **Implement database destinations for splits**
   - `SplitPlan.destination` allows a `DatabaseDestination`, but `split()` currently rejects it outright. Supporting writing partition results to a table (append/replace) would make split plans feature-parity with combine.

