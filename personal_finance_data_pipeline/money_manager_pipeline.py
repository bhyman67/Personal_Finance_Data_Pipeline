r"""
Build a unified All Transactions CSV from FirstBank, Robinhood, Upwork, and PNC exports.

Output columns:
- Post Date
- Transaction Date
- Account
- Amount
- Description
- Type
- Income or Expense
- Description Category

"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

MIN_POST_DATE = pd.Timestamp("2022-01-01")
MIN_PNC_POST_DATE = pd.Timestamp("2026-07-01")

OUTPUT_COLUMNS = [
    "Post Date",
    "Transaction Date",
    "Account",
    "Amount",
    "Description",
    "Type",
    "Income or Expense",
    "Description Category",
]

INCOME_UPWORK_TYPES = {"Bonus", "Fixed-price", "Hourly", "Expense reimbursement"}

# Source files required for this export.
REQUIRED_SOURCES: Sequence[str] = [
    "FirstBank - All Accounts - Transactions.csv",
    "FirstBank - Gold Visa - Transactions.csv",
    "PNC - accountActivityExport - Spending.csv",
    "PNC - accountActivityExport - Savings.csv",
    "PNC - Cash Rewards Visa - Transactions.csv",
    "Upwork - Transactions.csv",
    "Robinhood - Spending Settled - Transactions.csv",
    "Robinhood - Unified Transfers - Transactions.csv",
    "Robinhood - Historical Activities - Transactions.csv",
    "Robinhood - Sweeps - Transactions.csv",
    "Robinhood - Dividends - Transactions.csv",
    "Robinhood - Boost Payout - Transactions.csv",
    "Robinhood - Subscription Fees - Transactions.csv",
    "Robinhood - Checking - Transactions.csv",
    "Robinhood - Savings - Transactions.csv",
]

OPTIONAL_RULE_FILES: Dict[str, Sequence[str]] = {
    "description_lookup": (
        "Description Category Lookup.csv",
        "Description Category Lookup - Table.csv",
    ),
    "description_excludes": (
        "Description Exclude List.csv",
        "Description Excludes.csv",
    ),
    "manual_categories": (
        "Manual Description Categorizations.csv",
        "Manual Description Categorization.csv",
    ),
    "transaction_excludes": (
        "Transaction Excludes.csv",
    ),
    "event_replacements": (
        "Event Replacements.csv",
    ),
}

WORKBOOK_RULE_TABLES: Dict[str, str] = {
    "transaction_excludes": "TransactionExcludes",
    "description_lookup": "DescriptionCategoryLookup",
    "description_excludes": "DescriptionExcludes",
    "event_replacements": "EventReplacements",
    "manual_category_override": "ManualCategoryOverride",
    "path_registry": "PathRegistry",
}

# Maps PathRegistry "Data Source" column values to REQUIRED_SOURCES keys.
# Update this map if PathRegistry row names ever change.
PATH_REGISTRY_SOURCE_MAP: Dict[str, str] = {
    "FirstBank All Accounts Transactions": "FirstBank - All Accounts - Transactions.csv",
    "FirstBank Gold Visa Transactions": "FirstBank - Gold Visa - Transactions.csv",
    "PNC Spending Transactions": "PNC - accountActivityExport - Spending.csv",
    "PNC Savings Transactions": "PNC - accountActivityExport - Savings.csv",
    "PNC Cash Rewards Visa": "PNC - Cash Rewards Visa - Transactions.csv",
    "Upwork Transactions": "Upwork - Transactions.csv",
    "Robinhood Spending Settled Transactions": "Robinhood - Spending Settled - Transactions.csv",
    "Robinhood Unified Transfers Transactions": "Robinhood - Unified Transfers - Transactions.csv",
    "Robinhood Historical Activities Transactions": "Robinhood - Historical Activities - Transactions.csv",
    "Robinhood Sweeps Transactions": "Robinhood - Sweeps - Transactions.csv",
    "Robinhood Dividends Transactions": "Robinhood - Dividends - Transactions.csv",
    "Robinhood Boost Payout Transactions": "Robinhood - Boost Payout - Transactions.csv",
    "Robinhood Subscription Fees Transactions": "Robinhood - Subscription Fees - Transactions.csv",
    "Robinhood Checking Transactions": "Robinhood - Checking - Transactions.csv",
    "Robinhood Savings Transactions": "Robinhood - Savings - Transactions.csv",
}

PATH_REGISTRY_OUTPUT_KEY = "All Transactions"
PATH_REGISTRY_SEARCH_ROOT_KEY = "Search Root"


def remove_visa(txn_description: str) -> str:
    if pd.isna(txn_description):
        return ""
    return re.sub(r"\bVISA \b", "", str(txn_description)).strip()


def parse_money(value: object) -> Optional[float]:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text or text.upper() == "PENDING":
        return None

    # Handles values like "- $20.42", "$456.68", "($42.00)", "-300".
    sign = -1 if text.startswith("-") or "(" in text else 1
    cleaned = re.sub(r"[^0-9.]", "", text)
    if not cleaned:
        return None
    return sign * float(cleaned)


def parse_date(value: object) -> pd.Timestamp:
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, (int, float)):
        # Excel serial dates are days since 1899-12-30.
        if 20000 <= float(value) <= 80000:
            return pd.to_datetime(float(value), unit="D", origin="1899-12-30", errors="coerce")
        return pd.to_datetime(value, errors="coerce")

    text = str(value).strip()
    if not text:
        return pd.NaT

    # Handle numeric strings that represent Excel serial dates.
    if re.fullmatch(r"\d+(\.\d+)?", text):
        num = float(text)
        if 20000 <= num <= 80000:
            return pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
        return pd.to_datetime(text, errors="coerce")

    # PNC includes entries like "PENDING - 07/11/2026".
    if text.upper().startswith("PENDING"):
        date_part = text.split("-")[-1].strip()
        return pd.to_datetime(date_part, errors="coerce")

    # Infer dates with mixed formats (YYYY-MM-DD, mm/dd/yyyy, d-MMM-yy).
    return pd.to_datetime(text, errors="coerce")


def extract_and_remove_date(description: str, post_date: pd.Timestamp) -> Tuple[str, str]:
    """
    Mirrors the pipeline behavior:
    - Parse MM-DD from description when present
    - Infer year from post_date
    - Remove full "ON MM-DD HHMM" phrase from description
    """
    if pd.isna(post_date):
        return "", description or ""

    description = "" if pd.isna(description) else str(description)
    full_pattern = r"\bON\s\d{2}-\d{2}\s\d{4}\b"
    date_pattern = r"\s(\d{2})-(\d{2})"
    match = re.search(date_pattern, description)

    if match:
        month, day = map(int, match.groups())
        try:
            txn_date = post_date.replace(month=month, day=day)
            if (txn_date - post_date).days > 0:
                txn_date = txn_date.replace(year=txn_date.year - 1)
            cleaned_description = re.sub(full_pattern, "", description).strip()
            return format_mmddyyyy(txn_date), cleaned_description
        except ValueError:
            pass

    return format_mmddyyyy(post_date), description.strip()


def format_mmddyyyy(value: pd.Timestamp) -> str:
    if pd.isna(value):
        return ""
    return f"{value.month}/{value.day}/{value.year}"


def assign_credit_debit_ind(amount: float) -> str:
    return "Credit" if amount >= 0 else "Debit"


def assign_income_expense(account: str, credit_debit_ind: str) -> str:
    account_u = (account or "").upper()
    is_credit_card = "VISA" in account_u or "CREDIT CARD" in account_u

    if is_credit_card:
        return "Expense" if credit_debit_ind == "Credit" else "Income"

    return "Income" if credit_debit_ind == "Credit" else "Expense"


def normalize_firstbank_account_name(account: object) -> str:
    account_name = normalize_text(account)
    account_upper = account_name.upper()

    if "VISA" in account_upper:
        return account_name

    if "CHECKING" in account_upper:
        return "FirstBank Checking"

    if "SAVING" in account_upper or "LIQUID ASSET" in account_upper:
        return "FirstBank Savings"

    if account_name:
        return f"FirstBank {account_name}"

    return "FirstBank"


def parse_pnc_description_and_type(raw_description: object) -> Tuple[str, str]:
    description = normalize_text(raw_description)
    if not description:
        return "", "PNC"

    upper_description = description.upper()
    debit_marker = " DEBIT CARD PURCHASE"
    pos_marker = " POS PURCHASE"
    recurring_marker = " RECURRING DEBIT CARD"

    if recurring_marker in upper_description:
        marker_idx = upper_description.index(recurring_marker)
        merchant = description[:marker_idx].strip()
        merchant = re.sub(r"\s+x{2,}\d{2,4}$", "", merchant, flags=re.IGNORECASE).strip()
        return merchant, "RECURRING DEBIT CARD"

    if debit_marker in upper_description:
        marker_idx = upper_description.index(debit_marker)
        merchant = description[:marker_idx].strip()
        merchant = re.sub(r"\s+x{2,}\d{2,4}$", "", merchant, flags=re.IGNORECASE).strip()
        return merchant, "DEBIT CARD PURCHASE"

    if pos_marker in upper_description:
        marker_idx = upper_description.index(pos_marker)
        merchant = description[:marker_idx].strip()
        merchant = re.sub(r"\s+POSx+\d+\s+x+\d+$", "", merchant, flags=re.IGNORECASE).strip()
        return merchant, "POS PURCHASE"

    if upper_description.startswith("ONLINE TRANSFER"):
        return description, "TRANSFER"

    if upper_description.startswith("ATM"):
        return description, "ATM"

    return description, "PNC"


def parse_rh_bank_description_and_type(raw_description: object) -> Tuple[str, str]:
    description = normalize_text(raw_description)
    if not description:
        return "", "TRANSFER"

    upper_description = description.upper()
    debit_prefix = "DEBIT CARD TRANSACTION AT "

    if upper_description.startswith(debit_prefix):
        merchant = description[len(debit_prefix):].strip()
        merchant = re.sub(r"\s{2,}", " ", merchant)
        return merchant, "DEBIT CARD PURCHASE"

    if upper_description.startswith("INTEREST PAYMENT"):
        return description, "INTEREST"

    if upper_description.startswith("ACH DEPOSIT"):
        return description, "DEPOSIT"

    if "PAYROLL" in upper_description:
        return description, "PAYROLL"

    if upper_description.startswith("INTERNAL TRANSFER") or upper_description.startswith("INTER-ENTITY TRANSFER"):
        return description, "TRANSFER"

    return description, "TRANSFER"


def to_standard_frame(df: pd.DataFrame) -> pd.DataFrame:
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[OUTPUT_COLUMNS].copy()


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_amount(value: object) -> Optional[float]:
    if pd.isna(value):
        return None
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return None


def find_first_existing_file(search_root: Path, candidate_names: Sequence[str]) -> Optional[Path]:
    for candidate in candidate_names:
        matches = list(search_root.rglob(candidate))
        if matches:
            return matches[0]
    return None


def select_first_column_by_keywords(columns: Sequence[str], keywords: Sequence[str]) -> Optional[str]:
    lowered = {c: c.lower() for c in columns}
    for col, col_lower in lowered.items():
        if all(k in col_lower for k in keywords):
            return col
    return None


def select_column_by_keyword_groups(
    columns: Sequence[str], keyword_groups: Sequence[Sequence[str]]
) -> Optional[str]:
    for keywords in keyword_groups:
        col = select_first_column_by_keywords(columns, keywords)
        if col:
            return col
    return None


def find_workbook_table(workbook: Any, table_name: str) -> Tuple[str, Any]:
    for sheet in workbook.sheets:
        try:
            table = sheet.tables[table_name]
            return sheet.name, table
        except Exception:
            continue
    raise KeyError(f"Workbook table not found: {table_name}")


def read_workbook_table_as_dataframe(workbook: Any, table_name: str) -> pd.DataFrame:
    _, table = find_workbook_table(workbook, table_name)

    df = table.range.options(pd.DataFrame, header=1, index=False).value
    if df is None:
        return pd.DataFrame()

    if isinstance(df, pd.Series):
        df = df.to_frame()

    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame()

    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    out = out.dropna(how="all")
    return out.reset_index(drop=True)


def load_description_lookup_from_workbook(
    workbook: Any,
    table_name: str = WORKBOOK_RULE_TABLES["description_lookup"],
) -> Dict[str, str]:
    df = read_workbook_table_as_dataframe(workbook, table_name)
    if df.empty:
        return {}

    key_col = select_column_by_keyword_groups(
        df.columns,
        (
            ("description", "substring"),
            ("description",),
        ),
    ) or df.columns[0]
    val_col = select_column_by_keyword_groups(df.columns, (("category",),)) or (
        df.columns[1] if len(df.columns) > 1 else None
    )
    if val_col is None:
        return {}

    lookup: Dict[str, str] = {}
    for _, row in df.iterrows():
        key = normalize_text(row.get(key_col))
        value = normalize_text(row.get(val_col))
        if key:
            lookup[key] = value

    return lookup


def load_description_excludes_from_workbook(
    workbook: Any,
    table_name: str = WORKBOOK_RULE_TABLES["description_excludes"],
) -> List[str]:
    df = read_workbook_table_as_dataframe(workbook, table_name)
    if df.empty:
        return []

    col = select_column_by_keyword_groups(df.columns, (("description", "exclude"), ("exclude",))) or df.columns[0]
    return [normalize_text(v) for v in df[col].tolist() if normalize_text(v)]


def load_manual_categories_from_workbook(
    workbook: Any,
    table_name: str = WORKBOOK_RULE_TABLES["manual_category_override"],
) -> pd.DataFrame:
    df = read_workbook_table_as_dataframe(workbook, table_name)
    if df.empty:
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Description Category"])

    date_col = select_column_by_keyword_groups(df.columns, (("date",),)) or df.columns[0]
    amount_col = select_column_by_keyword_groups(df.columns, (("amount",),)) or (
        df.columns[1] if len(df.columns) > 1 else df.columns[0]
    )
    description_col = select_column_by_keyword_groups(df.columns, (("description",),)) or (
        df.columns[2] if len(df.columns) > 2 else df.columns[0]
    )
    category_col = select_column_by_keyword_groups(df.columns, (("category",),)) or (
        df.columns[3] if len(df.columns) > 3 else None
    )
    if category_col is None:
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Description Category"])

    out = pd.DataFrame(
        {
            "Date": df[date_col].apply(parse_date).apply(format_mmddyyyy),
            "Amount": df[amount_col].apply(parse_money),
            "Description": df[description_col].apply(normalize_text),
            "Description Category": df[category_col].apply(normalize_text),
        }
    )
    out["Amount"] = out["Amount"].apply(normalize_amount)
    out = out[out["Date"].astype(str).str.len() > 0]
    out = out[out["Amount"].notna()]
    out = out[out["Description"].str.len() > 0]
    return out.reset_index(drop=True)


def load_transaction_excludes_from_workbook(
    workbook: Any,
    table_name: str = WORKBOOK_RULE_TABLES["transaction_excludes"],
) -> pd.DataFrame:
    df = read_workbook_table_as_dataframe(workbook, table_name)
    if df.empty:
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Income or Expense"])

    date_col = select_column_by_keyword_groups(df.columns, (("date",),)) or df.columns[0]
    amount_col = select_column_by_keyword_groups(df.columns, (("amount",),)) or (
        df.columns[1] if len(df.columns) > 1 else df.columns[0]
    )
    description_col = select_column_by_keyword_groups(df.columns, (("description",),)) or (
        df.columns[2] if len(df.columns) > 2 else df.columns[0]
    )
    inc_exp_col = select_column_by_keyword_groups(
        df.columns,
        (("income", "expense"), ("income",), ("expense",)),
    ) or (df.columns[3] if len(df.columns) > 3 else None)
    if inc_exp_col is None:
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Income or Expense"])

    out = pd.DataFrame(
        {
            "Date": df[date_col].apply(parse_date).apply(format_mmddyyyy),
            "Amount": df[amount_col].apply(parse_money),
            "Description": df[description_col].apply(normalize_text),
            "Income or Expense": df[inc_exp_col].apply(normalize_text),
        }
    )
    out["Amount"] = out["Amount"].apply(normalize_amount)
    out = out[out["Date"].astype(str).str.len() > 0]
    out = out[out["Amount"].notna()]
    out = out[out["Description"].str.len() > 0]
    out = out[out["Income or Expense"].str.len() > 0]
    return out.reset_index(drop=True)


def load_source_paths_from_workbook(workbook: Any) -> Dict[str, Path]:
    """Read source CSV paths from the PathRegistry workbook table.

    Returns a dict keyed by REQUIRED_SOURCES names (same keys used by
    build_all_transactions), with resolved Path values from the Path column.
    Rows whose Data Source name is not in PATH_REGISTRY_SOURCE_MAP are ignored.
    """
    df = read_workbook_table_as_dataframe(workbook, WORKBOOK_RULE_TABLES["path_registry"])
    if df.empty:
        return {}

    source_col = select_column_by_keyword_groups(df.columns, (("data", "source"), ("source",), ("data",))) or df.columns[0]
    path_col = select_column_by_keyword_groups(df.columns, (("path",),)) or (df.columns[1] if len(df.columns) > 1 else None)

    if path_col is None:
        return {}

    result: Dict[str, Path] = {}
    for _, row in df.iterrows():
        data_source = normalize_text(row.get(source_col))
        raw_path = normalize_text(row.get(path_col))
        if not data_source or not raw_path:
            continue
        source_key = PATH_REGISTRY_SOURCE_MAP.get(data_source)
        if source_key:
            result[source_key] = Path(raw_path)

    return result


def load_output_path_from_workbook(workbook: Any) -> Optional[Path]:
    """Read the All Transactions output path from the PathRegistry workbook table."""
    df = read_workbook_table_as_dataframe(workbook, WORKBOOK_RULE_TABLES["path_registry"])
    if df.empty:
        return None

    source_col = select_column_by_keyword_groups(df.columns, (("data", "source"), ("source",), ("data",))) or df.columns[0]
    path_col = select_column_by_keyword_groups(df.columns, (("path",),)) or (df.columns[1] if len(df.columns) > 1 else None)

    if path_col is None:
        return None

    for _, row in df.iterrows():
        data_source = normalize_text(row.get(source_col))
        raw_path = normalize_text(row.get(path_col))
        if data_source == PATH_REGISTRY_OUTPUT_KEY and raw_path:
            return Path(raw_path)

    return None


def load_search_root_from_workbook(workbook: Any) -> Optional[Path]:
    """Read the Search Root path from the PathRegistry workbook table."""
    df = read_workbook_table_as_dataframe(workbook, WORKBOOK_RULE_TABLES["path_registry"])
    if df.empty:
        return None

    source_col = select_column_by_keyword_groups(df.columns, (("data", "source"), ("source",), ("data",))) or df.columns[0]
    path_col = select_column_by_keyword_groups(df.columns, (("path",),)) or (df.columns[1] if len(df.columns) > 1 else None)

    if path_col is None:
        return None

    for _, row in df.iterrows():
        data_source = normalize_text(row.get(source_col))
        raw_path = normalize_text(row.get(path_col))
        if data_source == PATH_REGISTRY_SEARCH_ROOT_KEY and raw_path:
            return Path(raw_path)

    return None


def load_event_replacements_from_workbook(
    workbook: Any,
    table_name: str = WORKBOOK_RULE_TABLES["event_replacements"],
) -> pd.DataFrame:
    df = read_workbook_table_as_dataframe(workbook, table_name)
    if df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    post_date_col = select_column_by_keyword_groups(df.columns, (("post", "date"), ("date",))) or df.columns[0]
    txn_date_col = select_column_by_keyword_groups(df.columns, (("transaction", "date"),)) or post_date_col
    account_col = select_column_by_keyword_groups(df.columns, (("account",),)) or (
        df.columns[1] if len(df.columns) > 1 else df.columns[0]
    )
    amount_col = select_column_by_keyword_groups(df.columns, (("amount",),)) or (
        df.columns[2] if len(df.columns) > 2 else df.columns[0]
    )
    description_col = select_column_by_keyword_groups(df.columns, (("description",),)) or (
        df.columns[3] if len(df.columns) > 3 else df.columns[0]
    )
    type_col = select_column_by_keyword_groups(df.columns, (("type",),)) or (
        df.columns[4] if len(df.columns) > 4 else None
    )
    inc_exp_col = select_column_by_keyword_groups(
        df.columns,
        (("income", "expense"), ("income",), ("expense",)),
    ) or (df.columns[5] if len(df.columns) > 5 else None)
    category_col = select_column_by_keyword_groups(df.columns, (("category",),)) or (
        df.columns[6] if len(df.columns) > 6 else None
    )

    out = pd.DataFrame(
        {
            "Post Date": df[post_date_col].apply(parse_date).apply(format_mmddyyyy),
            "Transaction Date": df[txn_date_col].apply(parse_date).apply(format_mmddyyyy),
            "Account": df[account_col].apply(normalize_text),
            "Amount": df[amount_col].apply(parse_money),
            "Description": df[description_col].apply(normalize_text),
            "Type": df[type_col].apply(normalize_text) if type_col else "MANUAL ADJUSTMENT",
            "Income or Expense": df[inc_exp_col].apply(normalize_text) if inc_exp_col else "Expense",
            "Description Category": df[category_col].apply(normalize_text) if category_col else "",
        }
    )
    out["Amount"] = out["Amount"].apply(normalize_amount)
    out = out[out["Post Date"].astype(str).str.len() > 0]
    out = out[out["Amount"].notna()]
    out = out[out["Description"].str.len() > 0]

    missing_txn_date = out["Transaction Date"].astype(str).str.len() == 0
    out.loc[missing_txn_date, "Transaction Date"] = out.loc[missing_txn_date, "Post Date"]

    return to_standard_frame(out.reset_index(drop=True))


def load_description_lookup(path: Optional[Path]) -> Dict[str, str]:
    if path is None or not path.exists():
        return {}

    df = pd.read_csv(path)
    if df.empty:
        return {}

    key_col = select_first_column_by_keywords(df.columns, ["description"]) or df.columns[0]
    val_col = (
        select_first_column_by_keywords(df.columns, ["category"])
        or (df.columns[1] if len(df.columns) > 1 else None)
    )
    if val_col is None:
        return {}

    lookup: Dict[str, str] = {}
    for _, row in df.iterrows():
        key = normalize_text(row.get(key_col))
        value = normalize_text(row.get(val_col))
        if key:
            lookup[key] = value
    return lookup


def load_description_excludes(path: Optional[Path]) -> List[str]:
    if path is None or not path.exists():
        return []

    df = pd.read_csv(path)
    if df.empty:
        return []

    col = df.columns[0]
    return [normalize_text(v) for v in df[col].tolist() if normalize_text(v)]


def load_manual_categories(path: Optional[Path]) -> pd.DataFrame:
    if path is None or not path.exists():
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Description Category"])

    df = pd.read_csv(path)
    if df.empty:
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Description Category"])

    date_col = select_first_column_by_keywords(df.columns, ["date"]) or df.columns[0]
    amount_col = select_first_column_by_keywords(df.columns, ["amount"]) or (
        df.columns[1] if len(df.columns) > 1 else df.columns[0]
    )
    description_col = select_first_column_by_keywords(df.columns, ["description"]) or (
        df.columns[2] if len(df.columns) > 2 else df.columns[0]
    )
    category_col = select_first_column_by_keywords(df.columns, ["category"]) or (
        df.columns[3] if len(df.columns) > 3 else None
    )

    if category_col is None:
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Description Category"])

    out = pd.DataFrame(
        {
            "Date": df[date_col].apply(parse_date).apply(format_mmddyyyy),
            "Amount": df[amount_col].apply(parse_money),
            "Description": df[description_col].apply(normalize_text),
            "Description Category": df[category_col].apply(normalize_text),
        }
    )
    out["Amount"] = out["Amount"].apply(normalize_amount)
    out = out[out["Date"].astype(str).str.len() > 0]
    out = out[out["Amount"].notna()]
    out = out[out["Description"].str.len() > 0]
    return out.reset_index(drop=True)


def load_transaction_excludes(path: Optional[Path]) -> pd.DataFrame:
    if path is None or not path.exists():
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Income or Expense"])

    df = pd.read_csv(path)
    if df.empty:
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Income or Expense"])

    date_col = select_first_column_by_keywords(df.columns, ["date"]) or df.columns[0]
    amount_col = select_first_column_by_keywords(df.columns, ["amount"]) or (
        df.columns[1] if len(df.columns) > 1 else df.columns[0]
    )
    description_col = select_first_column_by_keywords(df.columns, ["description"]) or (
        df.columns[2] if len(df.columns) > 2 else df.columns[0]
    )
    inc_exp_col = select_first_column_by_keywords(df.columns, ["income", "expense"]) or (
        df.columns[3] if len(df.columns) > 3 else None
    )

    if inc_exp_col is None:
        return pd.DataFrame(columns=["Date", "Amount", "Description", "Income or Expense"])

    out = pd.DataFrame(
        {
            "Date": df[date_col].apply(parse_date).apply(format_mmddyyyy),
            "Amount": df[amount_col].apply(parse_money),
            "Description": df[description_col].apply(normalize_text),
            "Income or Expense": df[inc_exp_col].apply(normalize_text),
        }
    )
    out["Amount"] = out["Amount"].apply(normalize_amount)
    out = out[out["Date"].astype(str).str.len() > 0]
    out = out[out["Amount"].notna()]
    out = out[out["Description"].str.len() > 0]
    out = out[out["Income or Expense"].str.len() > 0]
    return out.reset_index(drop=True)


def load_event_replacements(path: Optional[Path]) -> pd.DataFrame:
    if path is None or not path.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    df = pd.read_csv(path)
    if df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    post_date_col = select_first_column_by_keywords(df.columns, ["post", "date"]) or (
        select_first_column_by_keywords(df.columns, ["date"]) or df.columns[0]
    )
    txn_date_col = select_first_column_by_keywords(df.columns, ["transaction", "date"]) or post_date_col
    account_col = select_first_column_by_keywords(df.columns, ["account"]) or (
        df.columns[1] if len(df.columns) > 1 else df.columns[0]
    )
    amount_col = select_first_column_by_keywords(df.columns, ["amount"]) or (
        df.columns[2] if len(df.columns) > 2 else df.columns[0]
    )
    description_col = select_first_column_by_keywords(df.columns, ["description"]) or (
        df.columns[3] if len(df.columns) > 3 else df.columns[0]
    )
    type_col = select_first_column_by_keywords(df.columns, ["type"]) or (
        df.columns[4] if len(df.columns) > 4 else None
    )
    inc_exp_col = select_first_column_by_keywords(df.columns, ["income", "expense"]) or (
        df.columns[5] if len(df.columns) > 5 else None
    )
    category_col = select_first_column_by_keywords(df.columns, ["category"]) or (
        df.columns[6] if len(df.columns) > 6 else None
    )

    out = pd.DataFrame(
        {
            "Post Date": df[post_date_col].apply(parse_date).apply(format_mmddyyyy),
            "Transaction Date": df[txn_date_col].apply(parse_date).apply(format_mmddyyyy),
            "Account": df[account_col].apply(normalize_text),
            "Amount": df[amount_col].apply(parse_money),
            "Description": df[description_col].apply(normalize_text),
            "Type": df[type_col].apply(normalize_text) if type_col else "MANUAL ADJUSTMENT",
            "Income or Expense": df[inc_exp_col].apply(normalize_text) if inc_exp_col else "Expense",
            "Description Category": df[category_col].apply(normalize_text) if category_col else "",
        }
    )

    out["Amount"] = out["Amount"].apply(normalize_amount)
    out = out[out["Post Date"].astype(str).str.len() > 0]
    out = out[out["Amount"].notna()]
    out = out[out["Description"].str.len() > 0]

    # Default transaction date to post date if omitted.
    missing_txn_date = out["Transaction Date"].astype(str).str.len() == 0
    out.loc[missing_txn_date, "Transaction Date"] = out.loc[missing_txn_date, "Post Date"]

    return to_standard_frame(out.reset_index(drop=True))


def categorize_description(description: str, lookup: Dict[str, str]) -> str:
    text = normalize_text(description)
    for substring, category in lookup.items():
        if substring.upper() in text.upper():
            return category
    return ""


def apply_original_rules(
    all_txns: pd.DataFrame,
    description_lookup: Dict[str, str],
    description_excludes: List[str],
    manual_categories: pd.DataFrame,
    transaction_excludes: pd.DataFrame,
) -> pd.DataFrame:
    df = to_standard_frame(all_txns).copy()

    df["Description"] = df["Description"].apply(normalize_text)
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    df = df[df["Amount"].notna()].copy()
    df["Amount Match"] = df["Amount"].apply(normalize_amount)

    # Equivalent to Income_Expense_Exclude filtering by description substring.
    if description_excludes:
        pattern = "|".join(re.escape(item) for item in description_excludes if item)
        if pattern:
            df = df[~df["Description"].str.contains(pattern, case=False, regex=True, na=False)].copy()

    # Auto-categorize by substring lookup.
    if description_lookup:
        df["Description Category"] = df["Description"].apply(
            lambda x: categorize_description(x, description_lookup)
        )

    # Manual category overrides.
    if not manual_categories.empty:
        for _, row in manual_categories.iterrows():
            mask = (
                (df["Post Date"] == row["Date"])
                & (df["Amount Match"] == row["Amount"])
                & (df["Description"] == row["Description"])
            )
            if mask.any():
                df.loc[mask, "Description Category"] = row["Description Category"]

    # Transaction-level exclusion rules.
    if not transaction_excludes.empty:
        for _, row in transaction_excludes.iterrows():
            mask = (
                (df["Post Date"] == row["Date"])
                & (df["Amount Match"] == row["Amount"])
                & (df["Description"] == row["Description"])
                & (df["Income or Expense"] == row["Income or Expense"])
            )
            if mask.any():
                df = df[~mask].copy()

    df = df.drop(columns=["Amount Match"])
    return to_standard_frame(df)


def load_firstbank(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)

    df["Date"] = df["Date"].apply(parse_date)
    df = df[df["Date"].notna()].copy()

    if "Credit_Debit_Ind" in df.columns:
        df["Credit_Debit_Ind"] = df["Credit_Debit_Ind"].fillna("").astype(str).str.capitalize()
    else:
        df["Credit_Debit_Ind"] = df["Amount"].astype(float).apply(assign_credit_debit_ind)

    df["Normalized Account"] = df["Account"].apply(normalize_firstbank_account_name)

    df["Income or Expense"] = df.apply(
        lambda row: assign_income_expense(str(row.get("Normalized Account", "")), str(row["Credit_Debit_Ind"])),
        axis=1,
    )

    txn_info = df.apply(
        lambda row: extract_and_remove_date(str(row.get("Description", "")), row["Date"]),
        axis=1,
        result_type="expand",
    )

    out = pd.DataFrame(
        {
            "Post Date": df["Date"].apply(format_mmddyyyy),
            "Transaction Date": txn_info[0],
            "Account": df["Normalized Account"].astype(str),
            "Amount": df["Amount"].astype(float).abs(),
            "Description": txn_info[1].apply(remove_visa),
            "Type": df["Type"].astype(str),
            "Income or Expense": df["Income or Expense"],
            "Description Category": "",
        }
    )
    return to_standard_frame(out)


def load_upwork(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df = df[df["Transaction Type"].isin(INCOME_UPWORK_TYPES)].copy()

    df["Date"] = df["Date"].apply(parse_date)
    df = df[df["Date"].notna()].copy()

    out = pd.DataFrame(
        {
            "Post Date": df["Date"].apply(format_mmddyyyy),
            "Transaction Date": df["Date"].apply(format_mmddyyyy),
            "Account": "Upwork",
            "Amount": pd.to_numeric(df["Amount $"], errors="coerce").abs(),
            "Description": df["Transaction Summary"].fillna("").astype(str),
            "Type": df["Transaction Type"].fillna("").astype(str),
            "Income or Expense": "Income",
            "Description Category": "Upwork Income",
        }
    )

    out = out[out["Amount"].notna()].copy()
    return to_standard_frame(out)


def load_pnc(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)

    account = "PNC Spending" if "Spending" in path.name else "PNC Savings"

    df["PostRaw"] = df["Transaction Date"].astype(str)
    # Drop pending rows to keep the output settled-only.
    df = df[~df["PostRaw"].str.upper().str.startswith("PENDING")].copy()

    df["PostDate"] = df["PostRaw"].apply(parse_date)
    df["AmountNum"] = df["Amount"].apply(parse_money)

    df = df[df["PostDate"].notna() & df["AmountNum"].notna()].copy()
    df = df[df["PostDate"] >= MIN_PNC_POST_DATE].copy()
    df["Credit_Debit_Ind"] = df["AmountNum"].apply(assign_credit_debit_ind)
    df["Income or Expense"] = df["Credit_Debit_Ind"].apply(
        lambda x: "Income" if x == "Credit" else "Expense"
    )

    parsed = df["Transaction Description"].apply(parse_pnc_description_and_type)
    df["Clean Description"] = parsed.apply(lambda x: x[0])
    df["PNC Type"] = parsed.apply(lambda x: x[1])

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["PostDate"].apply(format_mmddyyyy),
            "Account": account,
            "Amount": df["AmountNum"].abs(),
            "Description": df["Clean Description"],
            "Type": df["PNC Type"],
            "Income or Expense": df["Income or Expense"],
            "Description Category": "",
        }
    )
    return to_standard_frame(out)


def parse_credit_debit_indicator(value: object) -> str:
    text = normalize_text(value).upper()
    if not text:
        return ""

    if "DEBIT" in text or text in {"D", "DR"}:
        return "Debit"
    if "CREDIT" in text or text in {"C", "CR"}:
        return "Credit"
    return ""


def infer_positive_is_expense(amounts: pd.Series) -> bool:
    positives = int((amounts > 0).sum())
    negatives = int((amounts < 0).sum())
    return positives >= negatives


def load_pnc_cash_rewards_visa(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    if df.empty:
        return to_standard_frame(pd.DataFrame(columns=OUTPUT_COLUMNS))

    post_col = select_column_by_keyword_groups(
        df.columns,
        (("post", "date"), ("posting", "date"), ("transaction", "date"), ("date",)),
    ) or df.columns[0]
    txn_col = select_column_by_keyword_groups(
        df.columns,
        (("transaction", "date"), ("purchase", "date"), ("date",)),
    ) or post_col
    amount_col = select_column_by_keyword_groups(df.columns, (("amount",),))
    description_col = select_column_by_keyword_groups(
        df.columns,
        (("description",), ("merchant",), ("payee",), ("memo",)),
    ) or (df.columns[2] if len(df.columns) > 2 else df.columns[0])
    type_col = select_column_by_keyword_groups(df.columns, (("type",), ("category",)))
    credit_debit_col = select_column_by_keyword_groups(
        df.columns,
        (("credit", "debit"), ("debit", "credit"), ("debit", "credit", "indicator"), ("direction",)),
    )

    if amount_col is None:
        raise ValueError(f"Unable to identify amount column for source: {path.name}")

    df["PostDate"] = df[post_col].apply(parse_date)
    df["TxnDate"] = df[txn_col].apply(parse_date)
    df["AmountSigned"] = df[amount_col].apply(parse_money)
    df["DescriptionClean"] = df[description_col].apply(normalize_text)

    df = df[df["PostDate"].notna() & df["AmountSigned"].notna()].copy()

    post_text = df[post_col].astype(str).str.upper()
    df = df[~post_text.str.startswith("PENDING")].copy()

    if credit_debit_col:
        df["Credit_Debit_Ind"] = df[credit_debit_col].apply(parse_credit_debit_indicator)
        missing_cd = df["Credit_Debit_Ind"].str.len() == 0
        df.loc[missing_cd, "Credit_Debit_Ind"] = df.loc[missing_cd, "AmountSigned"].apply(assign_credit_debit_ind)
        positive_is_expense = True
    else:
        positive_is_expense = infer_positive_is_expense(df["AmountSigned"])
        df["Credit_Debit_Ind"] = df["AmountSigned"].apply(assign_credit_debit_ind)

    if positive_is_expense:
        df["Income or Expense"] = df["Credit_Debit_Ind"].apply(
            lambda x: "Expense" if x == "Credit" else "Income"
        )
    else:
        df["Income or Expense"] = df["Credit_Debit_Ind"].apply(
            lambda x: "Income" if x == "Credit" else "Expense"
        )

    missing_txn = df["TxnDate"].isna()
    df.loc[missing_txn, "TxnDate"] = df.loc[missing_txn, "PostDate"]

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["TxnDate"].apply(format_mmddyyyy),
            "Account": "PNC Cash Rewards Visa",
            "Amount": df["AmountSigned"].abs(),
            "Description": df["DescriptionClean"],
            "Type": df[type_col].apply(normalize_text) if type_col else "CREDIT CARD",
            "Income or Expense": df["Income or Expense"],
            "Description Category": "",
        }
    )

    out = out[out["Description"].str.len() > 0].copy()
    return to_standard_frame(out)


def load_rh_spending_settled(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    date_col = "post_date" if "post_date" in df.columns else "initiated_at"

    df["PostDate"] = df[date_col].apply(parse_date)
    df = df[df["PostDate"].notna()].copy()
    df["Credit_Debit_Ind"] = df["direction"].fillna("").astype(str).str.capitalize()

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["PostDate"].apply(format_mmddyyyy),
            "Account": "Robinhood Cash Card",
            "Amount": pd.to_numeric(df["amount.amount"], errors="coerce").abs(),
            "Description": df["merchant_description"].fillna("").astype(str),
            "Type": "CASH CARD",
            "Income or Expense": df["Credit_Debit_Ind"].apply(
                lambda x: "Income" if x == "Credit" else "Expense"
            ),
            "Description Category": "",
        }
    )
    out = out[out["Amount"].notna()].copy()
    return to_standard_frame(out)


def load_rh_historical_activities(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)

    if "is_visible" in df.columns:
        df = df[df["is_visible"].astype(str).str.lower() == "true"].copy()

    df["PostDate"] = df["created_at"].apply(parse_date)
    df["AmountNum"] = pd.to_numeric(df["metadata.amount.amount"], errors="coerce")
    df = df[df["PostDate"].notna() & df["AmountNum"].notna()].copy()

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["PostDate"].apply(format_mmddyyyy),
            "Account": "Robinhood Cash Card",
            "Amount": df["AmountNum"].abs(),
            "Description": df["metadata.title"].fillna("CASH BACK").astype(str),
            "Type": "CASH BACK",
            "Income or Expense": "Income",
            "Description Category": "",
        }
    )
    return to_standard_frame(out)


def load_rh_sweeps(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["PostDate"] = df["pay_date"].apply(parse_date)
    df["AmountNum"] = pd.to_numeric(df["amount.amount"], errors="coerce")
    df = df[df["PostDate"].notna() & df["AmountNum"].notna()].copy()

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["PostDate"].apply(format_mmddyyyy),
            "Account": "Robinhood Brokerage",
            "Amount": df["AmountNum"].abs(),
            "Description": df["reason"].fillna("INTEREST").astype(str),
            "Type": "INTEREST",
            "Income or Expense": "Income",
            "Description Category": "Interest",
        }
    )
    return to_standard_frame(out)


def load_rh_dividends(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["PostDate"] = df["paid_at"].apply(parse_date)
    df["AmountNum"] = pd.to_numeric(df["amount"], errors="coerce")
    df = df[df["PostDate"].notna() & df["AmountNum"].notna()].copy()

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["PostDate"].apply(format_mmddyyyy),
            "Account": "Robinhood Brokerage",
            "Amount": df["AmountNum"].abs(),
            "Description": "DIVIDEND",
            "Type": "DIVIDEND",
            "Income or Expense": "Income",
            "Description Category": "",
        }
    )
    return to_standard_frame(out)


def load_rh_boost(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["PostDate"] = df["created_at"].apply(parse_date)
    df["AmountNum"] = pd.to_numeric(df["amount"], errors="coerce")
    df = df[df["PostDate"].notna() & df["AmountNum"].notna()].copy()

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["PostDate"].apply(format_mmddyyyy),
            "Account": "Robinhood Cash Management",
            "Amount": df["AmountNum"].abs(),
            "Description": df["title"].fillna("Cash Rewards").astype(str),
            "Type": "Cash Rewards",
            "Income or Expense": "Income",
            "Description Category": "",
        }
    )
    return to_standard_frame(out)


def load_rh_subscription_fees(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["PostDate"] = df["date"].apply(parse_date)
    df["AmountNum"] = pd.to_numeric(df["amount"], errors="coerce")
    df = df[df["PostDate"].notna() & df["AmountNum"].notna()].copy()

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["PostDate"].apply(format_mmddyyyy),
            "Account": "Robinhood Cash Management",
            "Amount": df["AmountNum"].abs(),
            "Description": "Subscription Fee",
            "Type": "SUBSCRIPTION",
            "Income or Expense": "Expense",
            "Description Category": "",
        }
    )
    return to_standard_frame(out)


def load_rh_unified_transfers(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)

    # Keep the payroll/cashout/interest-style entries from the original pipeline intent.
    desc = df.get("details.description", pd.Series(index=df.index, dtype="object")).fillna("").astype(str)
    originator = df.get("details.originator_name", pd.Series(index=df.index, dtype="object")).fillna("").astype(str)

    mask = (
        desc.str.contains(r"PAYROLL|INDIVIDUAL|CASHOUT|interest_payment|PAYMENTS", case=False, regex=True)
        | originator.str.contains("CorePower", case=False, regex=False)
    )
    filtered = df[mask].copy()

    if filtered.empty:
        return to_standard_frame(pd.DataFrame(columns=OUTPUT_COLUMNS))

    date_series = filtered.get("details.settlement_date", pd.Series(index=filtered.index, dtype="object"))
    if date_series.isna().all():
        date_series = filtered.get("created_at", pd.Series(index=filtered.index, dtype="object"))

    filtered["PostDate"] = date_series.apply(parse_date)
    filtered["AmountNum"] = pd.to_numeric(filtered.get("amount"), errors="coerce")

    # Fall back to transfer direction if details.direction is unavailable.
    details_direction = filtered.get("details.direction", pd.Series(index=filtered.index, dtype="object")).fillna("")
    transfer_direction = filtered.get("direction", pd.Series(index=filtered.index, dtype="object")).fillna("")

    credit_debit = details_direction.astype(str)
    credit_debit = credit_debit.where(credit_debit.str.len() > 0, transfer_direction.map({"push": "Credit", "pull": "Debit"}))
    credit_debit = credit_debit.fillna("").astype(str).str.capitalize()

    filtered = filtered[filtered["PostDate"].notna() & filtered["AmountNum"].notna()].copy()

    transfer_type = filtered.get("transfer_type", pd.Series(index=filtered.index, dtype="object")).fillna("TRANSFER")
    description = (
        filtered.get("details.description", pd.Series(index=filtered.index, dtype="object")).fillna("").astype(str)
        + " "
        + filtered.get("details.originator_name", pd.Series(index=filtered.index, dtype="object")).fillna("").astype(str)
    ).str.strip()

    out = pd.DataFrame(
        {
            "Post Date": filtered["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": filtered["PostDate"].apply(format_mmddyyyy),
            "Account": "Robinhood Cash Management",
            "Amount": filtered["AmountNum"].abs(),
            "Description": description,
            "Type": transfer_type.astype(str),
            "Income or Expense": credit_debit.loc[filtered.index].apply(
                lambda x: "Income" if x == "Credit" else "Expense"
            ),
            "Description Category": "",
        }
    )
    return to_standard_frame(out)


def load_rh_account_csv(path: Path, account_name: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    df["PostDate"] = df["Date"].apply(parse_date)
    df["AmountNum"] = pd.to_numeric(df["Amount"], errors="coerce")
    df = df[df["PostDate"].notna() & df["AmountNum"].notna()].copy()

    df["Credit_Debit_Ind"] = df["AmountNum"].apply(assign_credit_debit_ind)
    parsed = df["Description"].apply(parse_rh_bank_description_and_type)
    df["Clean Description"] = parsed.apply(lambda x: x[0])
    df["RH Type"] = parsed.apply(lambda x: x[1])

    out = pd.DataFrame(
        {
            "Post Date": df["PostDate"].apply(format_mmddyyyy),
            "Transaction Date": df["PostDate"].apply(format_mmddyyyy),
            "Account": account_name,
            "Amount": df["AmountNum"].abs(),
            "Description": df["Clean Description"],
            "Type": df["RH Type"],
            "Income or Expense": df["Credit_Debit_Ind"].apply(
                lambda x: "Income" if x == "Credit" else "Expense"
            ),
            "Description Category": "",
        }
    )
    return to_standard_frame(out)


def find_source_files(search_root: Path, required_sources: Iterable[str]) -> Dict[str, Path]:
    source_lookup: Dict[str, Path] = {}
    names = set(required_sources)

    for path in search_root.rglob("*.csv"):
        if path.name in names and path.name not in source_lookup:
            source_lookup[path.name] = path

    return source_lookup


def build_all_transactions(source_paths: Dict[str, Path]) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []

    frames.append(load_firstbank(source_paths["FirstBank - All Accounts - Transactions.csv"]))
    frames.append(load_firstbank(source_paths["FirstBank - Gold Visa - Transactions.csv"]))
    frames.append(load_pnc(source_paths["PNC - accountActivityExport - Spending.csv"]))
    frames.append(load_pnc(source_paths["PNC - accountActivityExport - Savings.csv"]))
    frames.append(load_pnc_cash_rewards_visa(source_paths["PNC - Cash Rewards Visa - Transactions.csv"]))
    frames.append(load_upwork(source_paths["Upwork - Transactions.csv"]))

    frames.append(load_rh_spending_settled(source_paths["Robinhood - Spending Settled - Transactions.csv"]))
    frames.append(load_rh_unified_transfers(source_paths["Robinhood - Unified Transfers - Transactions.csv"]))
    frames.append(load_rh_historical_activities(source_paths["Robinhood - Historical Activities - Transactions.csv"]))
    frames.append(load_rh_sweeps(source_paths["Robinhood - Sweeps - Transactions.csv"]))
    frames.append(load_rh_dividends(source_paths["Robinhood - Dividends - Transactions.csv"]))
    frames.append(load_rh_boost(source_paths["Robinhood - Boost Payout - Transactions.csv"]))
    frames.append(load_rh_subscription_fees(source_paths["Robinhood - Subscription Fees - Transactions.csv"]))
    frames.append(load_rh_account_csv(source_paths["Robinhood - Checking - Transactions.csv"], "Robinhood Checking"))
    frames.append(load_rh_account_csv(source_paths["Robinhood - Savings - Transactions.csv"], "Robinhood Savings"))

    all_txns = pd.concat(frames, ignore_index=True)

    all_txns["Amount"] = pd.to_numeric(all_txns["Amount"], errors="coerce")
    all_txns = all_txns[all_txns["Amount"].notna()].copy()

    all_txns["Post Date Sort"] = pd.to_datetime(all_txns["Post Date"], errors="coerce")
    all_txns = all_txns.sort_values(by="Post Date Sort", ascending=False).drop(columns=["Post Date Sort"])

    return to_standard_frame(all_txns)


def resolve_source_paths(
    required_sources: Iterable[str],
    search_root: Path,
    explicit_source_paths: Optional[Dict[str, Path]] = None,
) -> Dict[str, Path]:
    """Resolve required source files from explicit paths first, then search-root scan."""
    source_paths: Dict[str, Path] = {}

    for name in required_sources:
        explicit = (explicit_source_paths or {}).get(name)
        if explicit and explicit.exists():
            source_paths[name] = explicit

    missing = [name for name in required_sources if name not in source_paths]
    if missing:
        discovered = find_source_files(search_root, missing)
        source_paths.update(discovered)

    final_missing = [name for name in required_sources if name not in source_paths]
    if final_missing:
        raise FileNotFoundError(
            "Could not find all required CSV files. Missing:\n- " + "\n- ".join(final_missing)
        )

    return source_paths


def resolve_search_root(cli_search_root: str, workbook: Optional[Any] = None) -> Path:
    """Resolve search root from CLI override, workbook PathRegistry, then cwd."""
    if cli_search_root:
        return Path(cli_search_root)

    if workbook is not None:
        workbook_search_root = load_search_root_from_workbook(workbook)
        if workbook_search_root:
            return workbook_search_root

    return Path.cwd()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build All Transactions.csv from source exports")
    parser.add_argument(
        "--search-root",
        type=str,
        default="",
        help="Optional root folder to recursively search for required source CSV files.",
    )
    parser.add_argument(
        "--output",
        type=str,
        default="All Transactions.csv",
        help="Output CSV path (default: ./All Transactions.csv)",
    )
    parser.add_argument(
        "--description-lookup",
        type=str,
        default="",
        help="Optional CSV path for description substring to category mapping.",
    )
    parser.add_argument(
        "--description-excludes",
        type=str,
        default="",
        help="Optional CSV path for description substrings to exclude.",
    )
    parser.add_argument(
        "--manual-categories",
        type=str,
        default="",
        help="Optional CSV path for manual description category overrides.",
    )
    parser.add_argument(
        "--transaction-excludes",
        type=str,
        default="",
        help="Optional CSV path for transaction exclusion rules.",
    )
    parser.add_argument(
        "--event-replacements",
        type=str,
        default="",
        help="Optional CSV path for one-off replacement transactions to append to output.",
    )
    return parser.parse_args()


def main(workbook: Optional[Any] = None) -> None:
    args = parse_args()
    search_root = resolve_search_root(args.search_root, workbook)
    output_path = Path(args.output)

    if not search_root.exists():
        raise FileNotFoundError(f"Search root does not exist: {search_root}")

    explicit_source_paths = load_source_paths_from_workbook(workbook) if workbook is not None else None
    source_paths = resolve_source_paths(REQUIRED_SOURCES, search_root, explicit_source_paths)

    all_txns = build_all_transactions(source_paths)

    desc_lookup_path = (
        Path(args.description_lookup)
        if args.description_lookup
        else find_first_existing_file(search_root, OPTIONAL_RULE_FILES["description_lookup"])
    )
    desc_excludes_path = (
        Path(args.description_excludes)
        if args.description_excludes
        else find_first_existing_file(search_root, OPTIONAL_RULE_FILES["description_excludes"])
    )
    manual_categories_path = (
        Path(args.manual_categories)
        if args.manual_categories
        else find_first_existing_file(search_root, OPTIONAL_RULE_FILES["manual_categories"])
    )
    txn_excludes_path = (
        Path(args.transaction_excludes)
        if args.transaction_excludes
        else find_first_existing_file(search_root, OPTIONAL_RULE_FILES["transaction_excludes"])
    )
    event_replacements_path = (
        Path(args.event_replacements)
        if args.event_replacements
        else find_first_existing_file(Path.cwd(), OPTIONAL_RULE_FILES["event_replacements"])
    )

    description_lookup = load_description_lookup(desc_lookup_path)
    description_excludes = load_description_excludes(desc_excludes_path)
    manual_categories = load_manual_categories(manual_categories_path)
    transaction_excludes = load_transaction_excludes(txn_excludes_path)
    event_replacements = load_event_replacements(event_replacements_path)

    all_txns = apply_original_rules(
        all_txns,
        description_lookup,
        description_excludes,
        manual_categories,
        transaction_excludes,
    )

    if not event_replacements.empty:
        all_txns = pd.concat([all_txns, event_replacements], ignore_index=True)

    all_txns["Post Date Sort"] = pd.to_datetime(all_txns["Post Date"], errors="coerce")
    all_txns = all_txns[all_txns["Post Date Sort"] >= MIN_POST_DATE].copy()
    all_txns = all_txns.sort_values(by="Post Date Sort", ascending=False).drop(columns=["Post Date Sort"])

    all_txns.to_csv(output_path, index=False)

    print(f"Wrote {len(all_txns):,} rows to {output_path.resolve()}")
    print(f"Description lookup rows applied: {len(description_lookup):,}")
    print(f"Description excludes loaded: {len(description_excludes):,}")
    print(f"Manual category rules loaded: {len(manual_categories):,}")
    print(f"Transaction excludes loaded: {len(transaction_excludes):,}")
    print(f"Event replacements loaded: {len(event_replacements):,}")
    print(f"Minimum Post Date included: {MIN_POST_DATE.strftime('%Y-%m-%d')}")


if __name__ == "__main__":
    main()
