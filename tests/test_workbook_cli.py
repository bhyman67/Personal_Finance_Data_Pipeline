import io
import json
import os
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
    def setUp(self):
        isolated = tempfile.TemporaryDirectory()
        self.addCleanup(isolated.cleanup)
        env = patch.dict(os.environ, {"ONEDRIVE": isolated.name})
        env.start()
        self.addCleanup(env.stop)

    def create_workbook(
        self,
        path,
        transaction_excludes_rows=None,
        description_excludes_rows=None,
    ):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Config"
        row = 1
        row = add_table(
            sheet,
            "TransactionExcludes",
            ["Date", "Amount", "Description", "Income or Expense"],
            transaction_excludes_rows or [],
            row,
        )
        row = add_table(
            sheet,
            "DescriptionCategoryLookup",
            ["Description", "Category"],
            [["Coffee", "Dining"]],
            row,
        )
        row = add_table(
            sheet,
            "DescriptionExcludes",
            ["Description"],
            description_excludes_rows or [],
            row,
        )
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
            self.create_workbook(workbook_path)

            with (
                patch.object(
                    sys,
                    "argv",
                    [
                        "pf-build-all-transactions",
                        "--workbook",
                        str(workbook_path),
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
            self.assertNotIn("search_root", build.call_args.kwargs)
            self.assertEqual(build.call_args.kwargs["output_path"], output_path)
            self.assertEqual(build.call_args.kwargs["source_paths"], {})

    def test_workbook_pipeline_writes_requested_outputs(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            workbook_path = root / "Budget.xlsx"
            self.create_workbook(workbook_path)
            workbook_data = pipeline.load_workbook_table_data(workbook_path)
            source_paths = {name: root / name for name in pipeline.REQUIRED_SOURCES}
            for path in source_paths.values():
                path.touch()
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
                patch.object(pipeline, "build_all_transactions", return_value=transactions),
                patch("sys.stdout", new_callable=io.StringIO) as stdout,
            ):
                result = pipeline.build_all_transactions_from_workbook(
                    workbook_data,
                    output_path=root / "Configured.csv",
                    excluded_output_path=root / "Excluded.csv",
                    source_paths=source_paths,
                )

            output_path = root / "Configured.csv"
            excluded_output_path = root / "Excluded.csv"
            self.assertTrue(output_path.is_file())
            self.assertTrue(excluded_output_path.is_file())
            self.assertEqual(len(result), 1)
            self.assertEqual(len(pd.read_csv(output_path)), 1)
            self.assertEqual(len(pd.read_csv(excluded_output_path)), 0)
            self.assertIn(f"Wrote 0 excluded rows to {excluded_output_path.resolve()}", stdout.getvalue())

    def test_paths_json_supplies_workbook_and_paths(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            onedrive = Path(temp_dir)
            (onedrive / "Finance").mkdir()
            workbook_path = onedrive / "Finance" / "Budget.xlsx"
            self.create_workbook(workbook_path)
            source_name = "PNC Savings Transactions"
            (onedrive / "PATHS.json").write_text(
                json.dumps(
                    {
                        pipeline.PATHS_CONFIG_SECTION: {
                            "Workbook": {"path": "Finance\\Budget.xlsx"},
                            "All Transactions": {"path": "Finance\\Out.csv"},
                            "Excluded Transactions": {"path": "Finance\\Out - Excluded.csv"},
                            source_name: {"path": "Finance\\Savings.csv"},
                        }
                    }
                ),
                encoding="utf-8",
            )

            with (
                patch.dict(os.environ, {"ONEDRIVE": str(onedrive)}),
                patch.object(sys, "argv", ["pf-build-all-transactions"]),
                patch.object(pipeline, "build_all_transactions_from_workbook") as build,
            ):
                pipeline.main()

            kwargs = build.call_args.kwargs
            self.assertEqual(build.call_args.args[0].fullname, workbook_path.resolve())
            self.assertNotIn("search_root", kwargs)
            self.assertEqual(kwargs["output_path"], onedrive / "Finance" / "Out.csv")
            self.assertEqual(
                kwargs["excluded_output_path"], onedrive / "Finance" / "Out - Excluded.csv"
            )
            self.assertEqual(
                kwargs["source_paths"],
                {
                    pipeline.PATHS_SOURCE_MAP[source_name]: onedrive
                    / "Finance"
                    / "Savings.csv"
                },
            )

    def test_older_transactions_are_included_and_invalid_dates_are_excluded(self):
        rows = [
            ("2021-12-31", "Before floor"),
            ("2022-01-01", "At floor"),
            ("2024-01-03", "Recent"),
            ("2021-12-30", "Reject by description"),
            ("1/2/2023", "Reject by transaction"),
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
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            workbook_path = root / "Budget.xlsx"
            self.create_workbook(
                workbook_path,
                transaction_excludes_rows=[
                    ["2023-01-02", 10.0, "Reject by transaction", "Expense"]
                ],
                description_excludes_rows=[["Reject by description"]],
            )
            config = {
                "Workbook": {"path": "Budget.xlsx"},
                "All Transactions": {"path": "All Transactions.csv"},
                "Excluded Transactions": {"path": "Excluded.csv"},
            }
            for label, filename in pipeline.PATHS_SOURCE_MAP.items():
                source_path = root / "sources" / filename
                source_path.parent.mkdir(parents=True, exist_ok=True)
                source_path.touch()
                config[label] = {"path": str(source_path.relative_to(root))}
            (root / "PATHS.json").write_text(
                json.dumps({pipeline.PATHS_CONFIG_SECTION: config}),
                encoding="utf-8",
            )

            output_path = root / "All Transactions.csv"
            excluded_path = root / "Excluded.csv"
            with (
                patch.dict(os.environ, {"ONEDRIVE": str(root)}),
                patch.object(
                    sys,
                    "argv",
                    [
                        "pf-build-all-transactions",
                        "--output",
                        str(output_path),
                        "--excluded-output",
                        str(excluded_path),
                    ],
                ),
                patch.object(pipeline, "build_all_transactions", return_value=transactions),
                patch("sys.stdout", new_callable=io.StringIO) as stdout,
            ):
                pipeline.main()

            self.assertIn(
                f"Wrote 3 excluded rows to {excluded_path.resolve()}",
                stdout.getvalue(),
            )
            included = pd.read_csv(output_path)
            excluded = pd.read_csv(excluded_path).set_index("Description")
            self.assertEqual(
                list(included["Description"]), ["Recent", "At floor", "Before floor"]
            )
            self.assertEqual(
                set(excluded.index),
                {"Reject by description", "Reject by transaction", "Invalid date"},
            )
            self.assertTrue(excluded.loc["Reject by description", "Exclusion Reason"].startswith("Description exclude:"))
            self.assertTrue(excluded.loc["Reject by transaction", "Exclusion Reason"].startswith("Transaction exclude rule:"))
            self.assertEqual(excluded.loc["Invalid date", "Exclusion Reason"], "Post Date could not be parsed")

    def test_cli_requires_workbook_and_rejects_retired_options(self):
        with (
            patch.object(sys, "argv", ["pf-build-all-transactions"]),
            patch.object(pipeline, "load_paths_config", return_value={}),
            self.assertRaisesRegex(ValueError, "No workbook specified"),
        ):
            pipeline.main()

        for option in ("--search-root", "--description-lookup"):
            with self.subTest(option=option), patch("sys.stderr", new_callable=io.StringIO):
                with self.assertRaises(SystemExit):
                    pipeline.parse_args([option, "unused"])


if __name__ == "__main__":
    unittest.main()
