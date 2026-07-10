# P4V Wrapper

`ExcelMergeP4VDiff.exe` makes P4V spreadsheet diffs stable and editable.

- P4V passes temporary files that can disappear after startup. The wrapper copies both sides to `%LOCALAPPDATA%\ExcelSmartDiff\p4v-diff\cache` before opening ExcelMerge.
- The wrapper recognizes P4V's `%1 %2 --open` argument pattern.
- For an editable `.xls` or `.xlsx` right side, it passes the original path to ExcelMerge so `Save` can copy the edited working file back after a backup is created.
- Cache folders older than three days are removed. When more than 100 sessions exist, only the newest 50 are retained.

Build with the command in the root README. Put the wrapper under `p4v-diff/` next to an `app/` folder containing the ExcelMerge Release output.
