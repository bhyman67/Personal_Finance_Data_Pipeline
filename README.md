# Personal Finance Data Pipeline

Personal transaction normalization pipeline for consolidating exports from multiple financial sources into one analytics-ready CSV.

## What This Project Does

- Consolidates transaction exports from multiple providers into a single unified file.
- Normalizes key fields such as dates, amounts, account names, and transaction type.
- Applies optional rule tables for categorization, exclusions, and manual overrides.
- Produces an `All Transactions.csv` output with a stable column layout.
- Writes rule-excluded transactions and transactions with unparseable post dates to a companion `All Transactions - Excluded.csv` file with an exclusion reason.

Designed for repeatable monthly or ad hoc transaction consolidation.

### Supported Source Families

- Multiple banking and cash-management exports
- Card transaction exports
- Income/payment platform exports

## Windows Setup

Install Python 3.10 or newer with the Windows Python launcher, then run this from the cloned repository in PowerShell:

```powershell
.\setup.ps1
```

The script creates `.venv` if needed and installs the package and its dependencies. Run it again after pulling dependency changes. If script execution is restricted, run `powershell.exe -ExecutionPolicy Bypass -File .\setup.ps1` instead; this applies only to that PowerShell process.

No activation is needed if you run `.venv\Scripts\pf-build-all-transactions.exe` directly.

## Workbook-Driven CLI

This pipeline supports command-line runs configured from either CSV rule files or tables in a saved Excel workbook. Existing integrations can also pass a live workbook object.

In workbook-driven usage, the `--workbook` option reads the workbook's named tables directly, without requiring the xlwings Run button or a workbook-local runner. The PathRegistry table can provide explicit source file paths, set the `All Transactions` and `Excluded Transactions` output paths, and provide a search root; any missing source files can still be discovered via search-root fallback.

Path resolution order is explicit source-path mappings first (when provided by workbook integration), then recursive search under a configured search root.

If required source files are still missing after resolution, the run fails with a clear missing-files error.

This keeps workbook-based configuration while allowing the pipeline to be launched and maintained entirely from the repository.

Run a workbook-configured build from the repository environment:

```powershell
.\.venv\Scripts\pf-build-all-transactions.exe --workbook "D:\path\to\Money_Management.xlsx"
```

Alternatively, run `.\.venv\Scripts\python.exe -m personal_finance_data_pipeline.cli --workbook "<workbook path>"`. Save the workbook before running; the command reads the saved file. `--search-root`, `--output`, and `--excluded-output` override their PathRegistry values. With `--workbook`, rule configuration comes from the workbook tables; CSV rule options are for runs without `--workbook`.

The excluded-transactions file uses the `Excluded Transactions` path in PathRegistry when configured. Otherwise, it is written next to the main output as `<output name> - Excluded.csv`. For command-line runs, use `--excluded-output` to choose a different path. It records transactions excluded by the description-exclude rules or transaction-exclude rules, as well as transactions with unparseable post dates that reach the final date filter. Transactions excluded solely because they precede the minimum post date are omitted; rule-excluded transactions are still recorded regardless of their date. The output message reports how many date-floor exclusions were omitted. Rows discarded inside source-specific parsing (for example, invalid dates or amounts) are not included.

## Output Schema

The unified output contains:

- Post Date
- Transaction Date
- Account
- Amount
- Description
- Type
- Income or Expense
- Description Category

## Why This Exists

This is a custom personal-finance ETL tool built for repeatable transaction consolidation and category rule application across heterogeneous export formats.

## Privacy Note

This public repository intentionally excludes local machine paths, private notes, and personal environment details.
