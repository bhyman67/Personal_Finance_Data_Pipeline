# Personal Finance Data Pipeline

Personal transaction normalization pipeline for consolidating exports from multiple financial sources into one analytics-ready CSV.

## What This Project Does

- Consolidates transaction exports from multiple providers into a single unified file.
- Normalizes key fields such as dates, amounts, account names, and transaction type.
- Applies optional rule tables for categorization, exclusions, and manual overrides.
- Produces an `All Transactions.csv` output with a stable column layout.
- Writes rule-excluded and pre-2022 transactions to a companion `All Transactions - Excluded.csv` file with an exclusion reason.

Designed for repeatable monthly or ad hoc transaction consolidation.

### Supported Source Families

- Multiple banking and cash-management exports
- Card transaction exports
- Income/payment platform exports

## xlwings Integration

This pipeline is designed to work both as a CLI process and as a workbook-driven process via xlwings.

In workbook-driven usage, a workbook object can be passed into the pipeline entrypoint so table-based configuration can control file-path behavior. The PathRegistry table can provide explicit source file paths, and any missing files can still be discovered via search-root fallback.

Path resolution order is explicit source-path mappings first (when provided by workbook integration), then recursive search under a configured search root.

If required source files are still missing after resolution, the run fails with a clear missing-files error.

This keeps workbook automation flexible while preserving the same normalization logic used by CLI runs.

The excluded-transactions file is written next to the main output by default. For command-line runs, use `--excluded-output` to choose a different path. It records transactions excluded by the description-exclude rules, transaction-exclude rules, and minimum post-date filter; rows discarded inside source-specific parsing (for example, invalid dates or amounts) are not included.

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
