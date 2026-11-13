import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter
import pandas as pd

from personal_finance_data_pipeline import money_manager_pipeline as pipeline


def add_table(sheet, name, headers, rows, start_row):
    for row_index, values in enumerate([headers, *rows], start=start_row):
        for column_index, value in enumerate(values, start=1):
            sheet.cell(row=row_index, column=column_index, value=value)

    end_row = start_row + len(rows)
    end_column = len(headers)
    table = Table(
        displayName=name,
        ref=f"A{start_row}:{get_column_letter(end_column)}{end_row}",
    )
    sheet.add_table(table)
    return end_row + 2


class WorkbookCliTests(unittest.TestCase):
    def create_workbook(self, path, search_root):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Config"
        row = 1
        row = add_table(
            sheet,
            "PathRegistry",
            ["Data Source", "Path"],
            [
                ["Search Root", str(search_root)],
                ["All Transactions", str(path.with_name("Configured.csv"))],
                ["Excluded Transactions", str(path.with_name("Excluded.csv"))],
            ],
            row,
        )
        row = add_table(
            sheet,
            "TransactionExcludes",
            ["Date", "Amount", "Description", "Income or Expense"],
            [],
            row,
        )
        row = add_table(
            sheet,
            "DescriptionCategoryLookup",
            ["Description", "Category"],
            [["Coffee", "Dining"]],
            row,
        )
        row = add_table(sheet, "DescriptionExcludes", ["Description"], [], row)
        row = add_table(
            sheet,
            "EventReplacements",
            pipeline.OUTPUT_COLUMNS,
            [],
            row,
        )
        add_table(
            sheet,
            "ManualCategoryOverride",
            ["Date", "Amount", "Description", "Description Category"],
            [],
            row,
        )
        workbook.save(path)

    def test_workbook_cli_reads_tables_and_delegates_to_pipeline(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            workbook_path = root / "Budget.xlsx"
            output_path = root / "All Transactions.csv"
            search_root = root / "sources"
            search_root.mkdir()
            self.create_workbook(workbook_path, search_root)

            with (
                patch.object(
                    sys,
                    "argv",
                    [
                        "pf-build-all-transactions",
                        "--workbook",
                        str(workbook_path),
                        "--search-root",
                        str(search_root),
                        "--output",
                        str(output_path),
                    ],
                ),
                patch.object(pipeline, "build_all_transactions_from_workbook") as build,
            ):
                pipeline.main()

            workbook_data = build.call_args.args[0]
            self.assertEqual(workbook_data.fullname, workbook_path.resolve())
            self.assertEqual(
                pipeline.load_description_lookup_from_workbook(workbook_data),
                {"Coffee": "Dining"},
            )
            self.assertEqual(build.call_args.kwargs["search_root"], search_root)
            self.assertEqual(build.call_args.kwargs["output_path"], output_path)

    def test_workbook_pipeline_uses_registry_outputs_and_search_fallback(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            workbook_path = root / "Budget.xlsx"
            search_root = root / "sources"
            search_root.mkdir()
            self.create_workbook(workbook_path, search_root)
            workbook_data = pipeline.load_workbook_table_data(workbook_path)
            source_paths = {name: root / name for name in pipeline.REQUIRED_SOURCES}
            transactions = pd.DataFrame(
                [
                    {
                        "Post Date": "2024-01-03",
                        "Transaction Date": "2024-01-02",
                        "Account": "Checking",
                        "Amount": 10.0,
                        "Description": "Coffee Shop",
                        "Type": "Debit",
                        "Income or Expense": "Expense",
                        "Description Category": "",
                    }
                ]
            )

            with (
                patch.object(pipeline, "resolve_source_paths", return_value=source_paths),
                patch.object(pipeline, "build_all_transactions", return_value=transactions),
                patch("sys.stdout", new_callable=io.StringIO) as stdout,
            ):
                result = pipeline.build_all_transactions_from_workbook(
                    workbook_data,
                    search_root=search_root,
                )

            output_path = root / "Configured.csv"
            excluded_output_path = root / "Excluded.csv"
            self.assertTrue(output_path.is_file())
            self.assertTrue(excluded_output_path.is_file())
            self.assertEqual(len(result), 1)
            self.assertEqual(len(pd.read_csv(output_path)), 1)
            self.assertEqual(len(pd.read_csv(excluded_output_path)), 0)
            self.assertIn("(date-floor rows not written: 0)", stdout.getvalue())

    def test_date_floor_exclusions_are_not_written_in_either_cli_mode(self):
        rows = [
            ("2021-12-31", "Before floor"),
            ("2022-01-01", "At floor"),
            ("2024-01-03", "Recent"),
            ("2021-12-30", "Reject by description"),
            ("2023-01-02", "Reject by transaction"),
            ("not-a-date", "Invalid date"),
        ]
        transactions = pd.DataFrame(
            [
                {
                    "Post Date": post_date,
                    "Transaction Date": post_date,
                    "Account": "Checking",
                    "Amount": 10.0,
                    "Description": description,
                    "Type": "Debit",
                    "Income or Expense": "Expense",
                    "Description Category": "",
                }
                for post_date, description in rows
            ]
        )
        transaction_excludes = pd.DataFrame(
            [
                {
                    "Date": "2023-01-02",
                    "Amount": 10.0,
                    "Description": "Reject by transaction",
                    "Income or Expense": "Expense",
                }
            ]
        )

        for workbook_mode in (False, True):
            with self.subTest(workbook_mode=workbook_mode), tempfile.TemporaryDirectory() as temp_dir:
                root = Path(temp_dir)
                output_path = root / "All Transactions.csv"
                excluded_path = root / "Excluded.csv"
                args = [
                    "pf-build-all-transactions",
                    "--search-root",
                    str(root),
                    "--output",
                    str(output_path),
                    "--excluded-output",
                    str(excluded_path),
                ]
                if workbook_mode:
                    workbook_path = root / "Budget.xlsx"
                    self.create_workbook(workbook_path, root)
                    args.extend(["--workbook", str(workbook_path)])

                with (
                    patch.object(sys, "argv", args),
                    patch.object(pipeline, "resolve_source_paths", return_value={}),
                    patch.object(pipeline, "build_all_transactions", return_value=transactions),
                    patch.object(pipeline, "load_description_excludes", return_value=["by description"]),
                    patch.object(pipeline, "load_description_excludes_from_workbook", return_value=["by description"]),
                    patch.object(pipeline, "load_transaction_excludes", return_value=transaction_excludes),
                    patch.object(pipeline, "load_transaction_excludes_from_workbook", return_value=transaction_excludes),
                    patch("sys.stdout", new_callable=io.StringIO) as stdout,
                ):
                    pipeline.main()

                self.assertIn(
                    f"Wrote 3 excluded rows to {excluded_path.resolve()} (date-floor rows not written: 1)",
                    stdout.getvalue(),
                )
                included = pd.read_csv(output_path)
                excluded = pd.read_csv(excluded_path).set_index("Description")
                self.assertEqual(set(included["Description"]), {"At floor", "Recent"})
                self.assertEqual(
                    set(excluded.index),
                    {"Reject by description", "Reject by transaction", "Invalid date"},
                )
                self.assertTrue(excluded.loc["Reject by description", "Exclusion Reason"].startswith("Description exclude:"))
                self.assertTrue(excluded.loc["Reject by transaction", "Exclusion Reason"].startswith("Transaction exclude rule:"))
                self.assertEqual(excluded.loc["Invalid date", "Exclusion Reason"], "Post Date could not be parsed")


if __name__ == "__main__":
    unittest.main()
