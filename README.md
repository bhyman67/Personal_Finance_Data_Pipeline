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

This pipeline runs from a saved Excel workbook. The workbook supplies rule tables, while `PATHS.json` supplies the workbook, source-file, and output paths. Existing integrations can also pass a live workbook object.

The `--workbook` option reads the workbook's named rule tables directly, without requiring the xlwings Run button or a workbook-local runner. All required source files must have explicit entries in `PATHS.json`; the pipeline does not scan folders to discover them.

This keeps workbook-based configuration while allowing the pipeline to be launched and maintained entirely from the repository.

Run a workbook-configured build from the repository environment:

```powershell
.\.venv\Scripts\pf-build-all-transactions.exe --workbook "D:\path\to\Money_Management.xlsx"
```

Alternatively, run `.\.venv\Scripts\python.exe -m personal_finance_data_pipeline.cli --workbook "<workbook path>"`. Save the workbook before running; the command reads the saved file. The workbook path can be omitted when `Workbook` is configured in `PATHS.json`. `--output` and `--excluded-output` override the corresponding paths from `PATHS.json`.

The excluded-transactions file uses the `Excluded Transactions` path in `PATHS.json` when configured. Otherwise, it is written next to the main output as `<output name> - Excluded.csv`. For command-line runs, use `--excluded-output` to choose a different path. It records transactions excluded by the description-exclude rules or transaction-exclude rules, as well as transactions with unparseable post dates that reach the final date check. Rows discarded inside source-specific parsing (for example, invalid dates or amounts) are not included.

There is no global minimum post date. The PNC Spending and Savings loader retains its separate July 1, 2026 cutoff.

## Shared Path Configuration

If a `PATHS.json` file exists in the folder named by the `ONEDRIVE` environment variable, the CLI reads it so the same configuration works on every computer with its own OneDrive location. Its `personal_financial_data_pipeline` section lists each path relative to that folder: `Workbook`, `All Transactions`, `Excluded Transactions`, and every required source file. There is no search-root setting; missing or invalid source paths cause the run to fail.

Command-line options override the corresponding `PATHS.json` values. Paths are not read from the workbook, so the workbook needs no PathRegistry table. Rule tables are always read from the workbook. Entries marked `Inactive` are still read: those sources are no longer updated, but their historical data is still included.

Source labels and their loaders are defined together in `SOURCE_LOADERS`. To add a source, register its loader there and add the matching path entry to `PATHS.json`.

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
