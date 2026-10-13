import json
import logging
import os
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from pathlib import PurePosixPath
from urllib.parse import unquote_plus
import boto3
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
# ============================================================
# Logging
# ============================================================
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
# ============================================================
# Configuration
# ============================================================
AWS_REGION = os.getenv(
    "AWS_REGION",
    "us-east-2",
)
RAW_BUCKET = os.getenv(
    "RAW_BUCKET",
    "schools-raw",
).strip()
DEST_BUCKET = os.getenv(
    "DEST_BUCKET",
    "schools-transformed",
).strip()
# ============================================================
# Destination Prefixes
# ============================================================
# ------------------------------------------------------------
# Cincinnati Public Expense Detail
# ------------------------------------------------------------
EXPENSE_DETAIL_DEST_PREFIX = os.getenv(
    "CINCINNATI_PUBLIC_EXPENSE_DETAIL_DEST_PREFIX",
    "data_from_clients/cincinnati_public_expense_detail/",
).strip("/")
# ------------------------------------------------------------
# Cincinnati Public Non-Payroll Expenses
#
# NOTE:
# "cincinnat" is intentional because that is the requested
# table name.
# ------------------------------------------------------------
NON_PAYROLL_DEST_PREFIX = os.getenv(
    "CINCINNATI_PUBLIC_NON_PAYROLL_EXPENSES_DEST_PREFIX",
    "data_from_clients/cincinnati_public_non_payroll_expenses/",
).strip("/")
# ============================================================
# Report 1 - Cincinnati Public Expense Detail
# ============================================================
EXPENSE_DETAIL_FILENAME = (
    "1. Ohio Checkbook Q1 FY25 - OK.xlsx"
)
EXPENSE_DETAIL_SHEET_NAME = "Q1-FY25"
EXPENSE_DETAIL_FISCAL_YEAR = "2025"
EXPENSE_DETAIL_REPORT_TYPE = (
    "cincinnati_public_expense_detail"
)
# ============================================================
# Report 2 - Cincinnati Public Non-Payroll Expenses
# ============================================================
NON_PAYROLL_FILENAME = (
    "Exp=Non-Payroll_FY25_Checks_List_for_Auditors "
    "(from Adam_08-21-25 - format added).xlsx"
)
NON_PAYROLL_SHEET_NAME = "Data"
NON_PAYROLL_FISCAL_YEAR = "2025"
NON_PAYROLL_REPORT_TYPE = (
    "cincinnati_public_non_payroll_expenses"
)
# ============================================================
# Expense Detail Columns
# ============================================================
EXPENSE_DETAIL_SOURCE_COLUMNS = [
    "transaction_date",
    "vendor_name",
    "amount",
    "check_number",
    "usas_fund",
    "usas_fund_description",
    "usas_function",
    "usas_func_description",
    "usas_object",
    "usas_object_description",
]
EXPENSE_DETAIL_FINAL_COLUMNS = (
    EXPENSE_DETAIL_SOURCE_COLUMNS
    + ["fiscal_year"]
)
# ============================================================
# Non-Payroll Columns
# ============================================================
NON_PAYROLL_SOURCE_COLUMNS = [
    "oh_pe_id",
    "oh_pe_name",
    "invoice_number",
    "invoice_line_amount",
    "po_number",
    "pay_type",
    "oh_check_no",
    "oh_ck_dt",
    "key_fqa",
    "oh_gl_obj",
    "fqa",
    "check_amount",
    "oh_check_id",
]
NON_PAYROLL_FINAL_COLUMNS = (
    NON_PAYROLL_SOURCE_COLUMNS
    + ["fiscal_year"]
)
# ============================================================
# Normalize Column Names
# ============================================================
def normalize_column_name(name):
    """
    Convert Excel headers to lowercase snake_case.
    This normalization intentionally supports the way headers
    such as InvoiceLineAmount and PONUmber are currently being
    interpreted by the Cincinnati workbooks.
    """
    s = str(name).strip()
    # Split lower/digit followed by uppercase.
    s = re.sub(
        r"([a-z0-9])([A-Z])",
        r"\1_\2",
        s,
    )
    # Split acronym followed by normal capitalized word.
    s = re.sub(
        r"([A-Z]+)([A-Z][a-z])",
        r"\1_\2",
        s,
    )
    s = s.lower()
    # Replace punctuation/spaces with underscore.
    s = re.sub(
        r"[^a-z0-9]+",
        "_",
        s,
    )
    # Collapse multiple underscores.
    s = re.sub(
        r"_+",
        "_",
        s,
    )
    return s.strip("_")
# ============================================================
# Expense Detail Column Aliases
# ============================================================
EXPENSE_DETAIL_COLUMN_ALIASES = {
    "transaction_date":
        "transaction_date",
    "transactiondate":
        "transaction_date",
    "vendor_name":
        "vendor_name",
    "vendorname":
        "vendor_name",
    "amount":
        "amount",
    "check_number":
        "check_number",
    "checknumber":
        "check_number",
    "usas_fund":
        "usas_fund",
    "usasfund":
        "usas_fund",
    "usas_fund_description":
        "usas_fund_description",
    "usasfunddescription":
        "usas_fund_description",
    "usas_function":
        "usas_function",
    "usasfunction":
        "usas_function",
    "usas_func_description":
        "usas_func_description",
    "usasfuncdescription":
        "usas_func_description",
    "usas_object":
        "usas_object",
    "usasobject":
        "usas_object",
    "usas_object_description":
        "usas_object_description",
    "usasobjectdescription":
        "usas_object_description",
}
# ============================================================
# Non-Payroll Column Aliases
#
# Includes the exact normalized column names shown in the
# Lambda error:
#
#     invoiceline_amount
#     pon_umber
#
# ============================================================
NON_PAYROLL_COLUMN_ALIASES = {
    # --------------------------------------------------------
    # PE ID
    # --------------------------------------------------------
    "oh_pe_id":
        "oh_pe_id",
    "ohpeid":
        "oh_pe_id",
    # --------------------------------------------------------
    # PE Name
    # --------------------------------------------------------
    "oh_pe_name":
        "oh_pe_name",
    "ohpename":
        "oh_pe_name",
    # --------------------------------------------------------
    # Invoice Number
    # --------------------------------------------------------
    "invoice_number":
        "invoice_number",
    "invoicenumber":
        "invoice_number",
    # --------------------------------------------------------
    # Invoice Line Amount
    #
    # Actual normalized workbook value:
    # invoiceline_amount
    # --------------------------------------------------------
    "invoice_line_amount":
        "invoice_line_amount",
    "invoiceline_amount":
        "invoice_line_amount",
    "invoicelineamount":
        "invoice_line_amount",
    # --------------------------------------------------------
    # PO Number
    #
    # Actual normalized workbook value:
    # pon_umber
    # --------------------------------------------------------
    "po_number":
        "po_number",
    "pon_umber":
        "po_number",
    "ponumber":
        "po_number",
    # --------------------------------------------------------
    # Pay Type
    # --------------------------------------------------------
    "pay_type":
        "pay_type",
    "paytype":
        "pay_type",
    # --------------------------------------------------------
    # Check Number
    # --------------------------------------------------------
    "oh_check_no":
        "oh_check_no",
    "ohcheckno":
        "oh_check_no",
    # --------------------------------------------------------
    # Check Date
    # --------------------------------------------------------
    "oh_ck_dt":
        "oh_ck_dt",
    "ohckdt":
        "oh_ck_dt",
    # --------------------------------------------------------
    # Key FQA
    # --------------------------------------------------------
    "key_fqa":
        "key_fqa",
    "keyfqa":
        "key_fqa",
    # --------------------------------------------------------
    # GL Object
    # --------------------------------------------------------
    "oh_gl_obj":
        "oh_gl_obj",
    "ohglobj":
        "oh_gl_obj",
    # --------------------------------------------------------
    # FQA
    # --------------------------------------------------------
    "fqa":
        "fqa",
    # --------------------------------------------------------
    # Check Amount
    # --------------------------------------------------------
    "check_amount":
        "check_amount",
    "checkamount":
        "check_amount",
    # --------------------------------------------------------
    # Check ID
    # --------------------------------------------------------
    "oh_check_id":
        "oh_check_id",
    "ohcheckid":
        "oh_check_id",
}
# ============================================================
# Currency Conversion
# ============================================================
def coerce_currency_decimal(value):
    """
    Convert currency/numeric values to Decimal(20,2).
    Handles:
        1234.56
        1,234.56
        $1,234.56
        (1,234.56)
        -1234.56
        scientific notation
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    cleaned = str(value).strip()
    if not cleaned:
        return None
    # Accounting negative.
    if (
        cleaned.startswith("(")
        and cleaned.endswith(")")
    ):
        cleaned = (
            "-"
            + cleaned[1:-1]
        )
    # Remove currency formatting.
    cleaned = re.sub(
        r"[$,\s]",
        "",
        cleaned,
    )
    try:
        return Decimal(
            cleaned
        ).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
    except (
        InvalidOperation,
        ValueError,
    ) as exc:
        raise ValueError(
            "Invalid currency amount: "
            f"{value!r}"
        ) from exc
# ============================================================
# Text Cleanup
# ============================================================
def clean_text_series(series):
    return series.map(
        lambda value: (
            None
            if pd.isna(value)
            else str(value).strip()
        )
    ).astype("string")
# ============================================================
# Identifier Cleanup
# ============================================================
def clean_identifier_series(series):
    """
    Keep identifiers as strings and remove Excel's trailing .0
    when applicable.
    """
    cleaned = clean_text_series(
        series
    )
    cleaned = cleaned.str.replace(
        r"\.0$",
        "",
        regex=True,
    )
    return cleaned
# ============================================================
# Validate Cincinnati Public Source Path
# ============================================================
def validate_source_path(key):
    normalized_key = (
        key.strip("/").lower()
    )
    valid_prefixes = [
        "data_from_clients/cincinnati_public/",
        "data-from-clients/cincinnati_public/",
    ]
    if not any(
        normalized_key.startswith(prefix)
        for prefix in valid_prefixes
    ):
        raise ValueError(
            "Unexpected Cincinnati Public source path: "
            f"{key}. "
            f"Expected one of: {valid_prefixes}"
        )
# ============================================================
# Determine Report Type
# ============================================================
# Payroll: preserve all rows, including duplicate transactions, across every tab.
PAYROLL_FILENAME_PATTERN = re.compile(
    r"^Payroll Trans FY(?P<year>\d{2}|\d{4}) Q[1-4]\.xlsx$", re.IGNORECASE
)
PAYROLL_TABLE_NAME = "cincinnati_public_payroll_expenses"
PAYROLL_DEST_PREFIX = os.getenv(
    "CINCINNATI_PUBLIC_PAYROLL_EXPENSES_DEST_PREFIX",
    "data_from_clients/cincinnati_public_payroll_expenses/",
).strip("/")
PAYROLL_BATCH_ROWS = int(os.getenv("PAYROLL_BATCH_ROWS", "5000"))
PAYROLL_SOURCE_HEADERS = [
    "EEID", "EmployeeName", "CheckDt", "PayPeriod", "DebitAmt", "CreditAmt",
    "CDH", "CDHName", "FundCd", "FunctionCd", "ObjectCd", "ObjectCdDesc",
    "SpecialCostCenterCd", "SubjectCd", "OperatingUnitCd", "IntructionalLevel",
    "BudgetUnit",
]
PAYROLL_COLUMNS = [
    "eeid", "employee_name", "check_dt", "pay_period", "debit_amt", "credit_amt",
    "cdh", "cdh_name", "fund_cd", "function_cd", "object_cd", "object_cd_desc",
    "special_cost_center_cd", "subject_cd", "operating_unit_cd", "instructional_level",
    "budget_unit",
]
def get_payroll_schema():
    fields = []
    for col in PAYROLL_COLUMNS:
        dtype = (pa.decimal128(20, 2) if col in {"debit_amt", "credit_amt"}
                 else pa.timestamp("ns") if col == "check_dt" else pa.string())
        fields.append(pa.field(col, dtype))
    return pa.schema(fields + [pa.field("source_sheet", pa.string()),
                              pa.field("source_row", pa.int64()),
                              pa.field("fiscal_year", pa.string())])
def payroll_record(values, sheet_name, row_number, fiscal_year):
    from datetime import datetime, date, time
    from openpyxl.utils.datetime import from_excel
    result = {}
    for col, value in zip(PAYROLL_COLUMNS, values):
        if col in {"debit_amt", "credit_amt"}:
            value = coerce_currency_decimal(value)
        elif col == "check_dt":
            if value in (None, ""):
                value = None
            elif isinstance(value, datetime):
                pass
            elif isinstance(value, date):
                value = datetime.combine(value, time())
            elif isinstance(value, (int, float)):
                value = from_excel(value)
            else:
                value = datetime.strptime(str(value).strip(), "%m/%d/%Y")
        elif value is not None:
            # Keep original text, including leading zeros and description spacing.
            value = str(value)
        result[col] = value
    result.update(source_sheet=sheet_name, source_row=row_number, fiscal_year=fiscal_year)
    return result
def convert_payroll_s3_object(bucket, key, s3):
    """Stream all worksheets into one Parquet file without combining giant frames."""
    import tempfile
    from pathlib import Path
    from openpyxl import load_workbook
    if PAYROLL_BATCH_ROWS < 1:
        raise ValueError("PAYROLL_BATCH_ROWS must be positive")
    match = PAYROLL_FILENAME_PATTERN.fullmatch(PurePosixPath(key).name)
    if match is None:
        raise ValueError(f"Unrecognized payroll filename: {key}")
    year = match.group("year")
    fiscal_year = "20" + year if len(year) == 2 else year
    schema = get_payroll_schema()
    dest_key = f"{PAYROLL_DEST_PREFIX}/{PurePosixPath(key).stem}.parquet"
    counts = {}
    with tempfile.TemporaryDirectory(prefix="cincinnati-payroll-") as directory:
        source_path = str(Path(directory) / "source.xlsx")
        output_path = str(Path(directory) / "payroll.parquet")
        s3.download_file(bucket, key, source_path)
        # Read cached formula results; source workbook contains transaction values.
        workbook = load_workbook(source_path, read_only=True, data_only=True)
        try:
            with pq.ParquetWriter(output_path, schema, compression="snappy") as writer:
                for sheet in workbook:
                    rows = sheet.iter_rows(values_only=True)
                    headers = next(rows, None)
                    if headers is None:
                        counts[sheet.title] = 0
                        continue
                    if list(headers) != PAYROLL_SOURCE_HEADERS:
                        raise ValueError(f"Unexpected payroll headers in {sheet.title!r}: {headers!r}")
                    batch = []
                    count = 0
                    for row_number, values in enumerate(rows, start=2):
                        if all(value is None for value in values):
                            continue
                        try:
                            batch.append(payroll_record(values, sheet.title, row_number, fiscal_year))
                        except Exception as exc:
                            raise ValueError(f"Invalid payroll value in sheet {sheet.title!r}, row {row_number}") from exc
                        count += 1
                        if len(batch) >= PAYROLL_BATCH_ROWS:
                            writer.write_table(pa.Table.from_pylist(batch, schema=schema))
                            batch.clear()
                    if batch:
                        writer.write_table(pa.Table.from_pylist(batch, schema=schema))
                    counts[sheet.title] = count
                    logger.info("Payroll sheet %s: %s rows written", sheet.title, count)
        finally:
            workbook.close()
        s3.upload_file(output_path, DEST_BUCKET, dest_key)
    return {"source": f"s3://{bucket}/{key}",
            "destination": f"s3://{DEST_BUCKET}/{dest_key}",
            "rows": sum(counts.values()), "sheet_rows": counts,
            "fiscal_year": fiscal_year, "report_type": PAYROLL_TABLE_NAME,
            "table_name": PAYROLL_TABLE_NAME}
# Cincinnati Public Revenue Ledger: the fund pivot is a summary, not transactions.
REVENUE_LEDGER_FILENAME = "RECLED_Date (GL1020) Revenue Ledger (generated 08-17-25) (Amt column formatted).xlsx"
REVENUE_LEDGER_TABLE_NAME = "cincinnati_public_revenue_ledger"
REVENUE_LEDGER_DEST_PREFIX = os.getenv(
    "CINCINNATI_PUBLIC_REVENUE_LEDGER_DEST_PREFIX",
    "data_from_clients/cincinnati_public_revenue_ledger/",
).strip("/")
REVENUE_LEDGER_COLUMNS = [
    "reference", "unique_id", "date", "description", "fund", "function",
    "object", "scc", "subject", "opu", "il", "budu", "amount",
]
REVENUE_LEDGER_HEADERS = [
    "Reference", "Unique ID", "Date", "Description", "Fund", "Function",
    "Object", "SCC", "Subject", "OPU", "IL", "BUDU", "Amount (Column Formatted)",
]
def get_revenue_ledger_schema():
    return pa.schema([
        pa.field(col, pa.decimal128(20, 2) if col == "amount"
                 else pa.timestamp("ns") if col == "date" else pa.string())
        for col in REVENUE_LEDGER_COLUMNS
    ] + [pa.field("fiscal_year", pa.string())])
def convert_revenue_ledger_excel(content):
    df = pd.read_excel(BytesIO(content), sheet_name="2025", dtype=object,
                       engine="openpyxl")
    headers = [str(col).strip() for col in df.columns]
    if headers != REVENUE_LEDGER_HEADERS:
        raise ValueError(f"Unexpected Revenue Ledger columns: {headers!r}; "
                         f"expected {REVENUE_LEDGER_HEADERS!r}")
    df.columns = REVENUE_LEDGER_COLUMNS
    df = df.dropna(how="all").copy()
    for col in REVENUE_LEDGER_COLUMNS:
        if col == "amount":
            df[col] = df[col].map(coerce_currency_decimal)
        elif col == "date":
            # Reject malformed dates rather than silently losing source values.
            df[col] = pd.to_datetime(df[col], errors="raise")
        else:
            df[col] = clean_text_series(df[col])
    df["fiscal_year"] = "2025"
    return df
# Revenue Ledger by Object and Fund; Pivot contains summaries only.
REVENUE_OBJ_FND_FILENAME = "REVLED (GL1015) Revenue Ledger by Object and Fund (generated 08-17-25).xlsx"
REVENUE_OBJ_FND_TABLE_NAME = "cincinnati_public_revenue_ledger_obj_fnd"
REVENUE_OBJ_FND_DEST_PREFIX = os.getenv(
    "CINCINNATI_PUBLIC_REVENUE_LEDGER_OBJ_FND_DEST_PREFIX",
    "data_from_clients/cincinnati_public_revenue_ledger_obj_fnd/",
).strip("/")
REVENUE_OBJ_FND_HEADERS = [
    "Object", "Fund", "Function", "SCC", "Subj", "OPU", "IL", "BUDU",
    "Date", "Description", "Debit", "Credit",
]
REVENUE_OBJ_FND_COLUMNS = [
    "object", "fund", "function", "scc", "subj", "opu", "il", "budu",
    "date", "description", "debit", "credit",
]
def get_revenue_obj_fnd_schema():
    return pa.schema([
        pa.field(col, pa.decimal128(20, 2) if col in {"debit", "credit"}
                 else pa.timestamp("ns") if col == "date" else pa.string())
        for col in REVENUE_OBJ_FND_COLUMNS
    ] + [pa.field("fiscal_year", pa.string())])
def convert_revenue_obj_fnd_excel(content):
    df = pd.read_excel(BytesIO(content), sheet_name="2025", dtype=object,
                       engine="openpyxl")
    headers = [str(col).strip() for col in df.columns]
    if headers != REVENUE_OBJ_FND_HEADERS:
        raise ValueError(f"Unexpected Revenue Ledger Object/Fund columns: {headers!r}; "
                         f"expected {REVENUE_OBJ_FND_HEADERS!r}")
    df.columns = REVENUE_OBJ_FND_COLUMNS
    df = df.dropna(how="all").copy()
    for col in REVENUE_OBJ_FND_COLUMNS:
        if col in {"debit", "credit"}:
            df[col] = df[col].map(coerce_currency_decimal)
        elif col == "date":
            df[col] = pd.to_datetime(df[col], errors="raise")
        else:
            df[col] = clean_text_series(df[col])
    df["fiscal_year"] = "2025"
    return df
# Cincinnati Public Cash Fund Balances.
CASH_FUND_BAL_FILENAME = "Cash Fund Balances_FY25_GL1011-All Funds-Cash Summary by Fund (generated 08-17-25).xlsx"
CASH_FUND_BAL_TABLE_NAME = "cincinnati_public_cash_fund_bal"
CASH_FUND_BAL_DEST_PREFIX = os.getenv(
    "CINCINNATI_PUBLIC_CASH_FUND_BAL_DEST_PREFIX",
    "data_from_clients/cincinnati_public_cash_fund_bal/",
).strip("/")
CASH_FUND_BAL_HEADERS = [
    "Fund", "Fund Description", "Begining Balance", "YTD Receipts",
    "YTD Expenditures", "Current Balance", "Total Encumbrances",
    "Unencumbered Balance",
]
CASH_FUND_BAL_COLUMNS = [
    "fund", "fund_description", "beginning_balance", "ytd_receipts",
    "ytd_expenditures", "current_balance", "total_encumbrances",
    "unencumbered_balance",
]
def get_cash_fund_bal_schema():
    return pa.schema([
        pa.field(col, pa.string() if col in {"fund", "fund_description"}
                 else pa.decimal128(20, 2))
        for col in CASH_FUND_BAL_COLUMNS
    ] + [pa.field("fiscal_year", pa.string())])
def convert_cash_fund_bal_excel(content):
    df = pd.read_excel(BytesIO(content), sheet_name="2025", dtype=object,
                       engine="openpyxl")
    headers = [str(col).strip() for col in df.columns]
    if headers != CASH_FUND_BAL_HEADERS:
        raise ValueError(f"Unexpected Cash Fund Balances columns: {headers!r}; "
                         f"expected {CASH_FUND_BAL_HEADERS!r}")
    df.columns = CASH_FUND_BAL_COLUMNS
    df = df.dropna(how="all").copy()
    # The workbook's grand total has no fund or description. Exclude it
    # so that aggregating the exported fund rows does not count it twice.
    totals = df["fund"].isna() & df["fund_description"].isna()
    logger.info("Excluded %s Cash Fund Balances aggregate rows", int(totals.sum()))
    df = df.loc[~totals].copy()
    df["fund"] = clean_identifier_series(df["fund"]).str.zfill(3)
    if not df["fund"].str.fullmatch(r"[0-9]{3}", na=False).all():
        raise ValueError("Cash Fund Balances contains an invalid fund code")
    df["fund_description"] = clean_text_series(df["fund_description"])
    for col in CASH_FUND_BAL_COLUMNS[2:]:
        df[col] = df[col].map(coerce_currency_decimal)
    df["fiscal_year"] = "2025"
    return df
# All Ohio Checkbook quarters share one table; stream large workbooks in batches.
CHECK_BOOK_TABLE_NAME = "cincinnati_public_check_book"
CHECK_BOOK_DEST_PREFIX = os.getenv(
    "CINCINNATI_PUBLIC_CHECK_BOOK_DEST_PREFIX",
    "data_from_clients/cincinnati_public_check_book/",
).strip("/")
CHECK_BOOK_FILENAME_PATTERN = re.compile(
    r"^(?:[0-9]+\.\s\*)?Ohio Checkbook Q[1-4] FY(?:[0-9]{4}|[0-9]{2})(?:.\*)\.xlsx$",
    re.IGNORECASE,
)
CHECK_BOOK_SHEET_PATTERN = re.compile(r"^Q([1-4])[- _]\*FY([0-9]{4}|[0-9]{2})$", re.IGNORECASE)
CHECK_BOOK_HEADERS = [
    "TransactionDate", "VendorName", "Amount", "CheckNumber", "USASFund",
    "USASFundDescription", "USASFunction", "USASFuncDescription",
    "USASObject", "USASObjectDescription",
]
def get_check_book_schema():
    return pa.schema(list(get_expense_detail_schema()) + [
        pa.field("quarter", pa.string()), pa.field("source_sheet", pa.string()),
        pa.field("source_row", pa.int64()),
    ])
def check_book_period(sheet_name, filename):
    # Tab labels vary: Q1-FY25, FY25 Q1, Q1, or Sheet1.
    def quarter(text):
        match = re.search(r"(?:Q|QUARTER)\s\*([1-4])", text, re.IGNORECASE)
        return match.group(1) if match else None
    def year(text):
        match = re.search(r"FY\s\*[-_ ]\*([0-9]{4}|[0-9]{2})(?![0-9])", text, re.IGNORECASE)
        return match.group(1) if match else None
    q = quarter(sheet_name) or quarter(filename)
    y = year(sheet_name) or year(filename)
    if not q or not y:
        raise ValueError(f"Cannot determine Checkbook quarter/year from tab {sheet_name!r} "
                         f"and filename {filename!r}")
    return "Q" + q, "20" + y if len(y) == 2 else y
def check_book_record(values, sheet_name, row_number, filename=""):
    quarter, fiscal_year = check_book_period(sheet_name, filename)
    record = dict(zip(EXPENSE_DETAIL_SOURCE_COLUMNS, values))
    for col in EXPENSE_DETAIL_SOURCE_COLUMNS:
        value = record[col]
        if col == "amount":
            value = coerce_currency_decimal(value)
        elif col == "transaction_date":
            if value is None or value == "":
                value = None
            elif isinstance(value, (int, float)):
                from openpyxl.utils.datetime import from_excel
                value = from_excel(value)
            else:
                value = pd.to_datetime(value, errors="raise").to_pydatetime()
        elif value is not None:
            value = str(value).strip()
            if col in {"check_number", "usas_fund", "usas_function", "usas_object"}:
                value = re.sub(r"\.0$", "", value)
        record[col] = value
    record.update(fiscal_year=fiscal_year,
                  quarter=quarter, source_sheet=sheet_name,
                  source_row=row_number)
    return record
def convert_check_book_s3_object(bucket, key, s3):
    import tempfile
    from pathlib import Path
    from openpyxl import load_workbook
    if PAYROLL_BATCH_ROWS < 1:
        raise ValueError("PAYROLL_BATCH_ROWS must be positive")
    schema = get_check_book_schema()
    # Stable source-key hash prevents collisions between similarly named uploads.
    import hashlib
    suffix = hashlib.sha256(key.encode("utf-8")).hexdigest()[:16]
    dest_key = f"{CHECK_BOOK_DEST_PREFIX}/{PurePosixPath(key).stem}_{suffix}.parquet"
    counts = {}
    with tempfile.TemporaryDirectory(prefix="cincinnati-check-book-") as directory:
        source_path = str(Path(directory) / "source.xlsx")
        output_path = str(Path(directory) / "check_book.parquet")
        s3.download_file(bucket, key, source_path)
        workbook = load_workbook(source_path, read_only=True, data_only=True)
        try:
            logger.info("Checkbook workbook tabs: %s", workbook.sheetnames)
            with pq.ParquetWriter(output_path, schema, compression="snappy") as writer:
                for sheet in workbook:
                    rows = sheet.iter_rows(values_only=True)
                    header_row = None
                    # Allow titles and blank rows ahead of the actual headers.
                    for row_number, headers in enumerate(rows, start=1):
                        normalized = [normalize_column_name(v) if v is not None else ""
                                      for v in headers]
                        expected = [normalize_column_name(v) for v in CHECK_BOOK_HEADERS]
                        # Ignore trailing empty formatted columns.
                        while normalized and normalized[-1] == "":
                            normalized.pop()
                        if normalized == expected:
                            header_row = row_number
                            break
                        if row_number >= 25:
                            break
                    if header_row is None:
                        logger.info("Skipping tab %s: Checkbook headers not found in first 25 rows", sheet.title)
                        continue
                    check_book_period(sheet.title, PurePosixPath(key).name)
                    batch = []
                    count = 0
                    for row_number, values in enumerate(rows, start=header_row + 1):
                        if all(value is None for value in values):
                            continue
                        if any(value is not None for value in values[len(CHECK_BOOK_HEADERS):]):
                            raise ValueError(f"Unexpected extra Checkbook values in {sheet.title!r}, row {row_number}")
                        try:
                            batch.append(check_book_record(values[:len(CHECK_BOOK_HEADERS)], sheet.title, row_number, PurePosixPath(key).name))
                        except Exception as exc:
                            raise ValueError(f"Invalid Checkbook value in {sheet.title!r}, row {row_number}") from exc
                        count += 1
                        if len(batch) >= PAYROLL_BATCH_ROWS:
                            writer.write_table(pa.Table.from_pylist(batch, schema=schema))
                            batch.clear()
                    if batch:
                        writer.write_table(pa.Table.from_pylist(batch, schema=schema))
                    counts[sheet.title] = count
                    logger.info("Checkbook sheet %s: %s rows written", sheet.title, count)
        finally:
            workbook.close()
        if not counts:
            raise ValueError(f"No Checkbook data headers found. Tabs: {workbook.sheetnames!r}")
        s3.upload_file(output_path, DEST_BUCKET, dest_key)
    return {"source": f"s3://{bucket}/{key}",
            "destination": f"s3://{DEST_BUCKET}/{dest_key}",
            "rows": sum(counts.values()), "sheet_rows": counts,
            "report_type": CHECK_BOOK_TABLE_NAME, "table_name": CHECK_BOOK_TABLE_NAME}
def report_type_for_key(key):
    filename = PurePosixPath(
        key
    ).name
    filename_lower = (
        filename.lower()
    )
    if CHECK_BOOK_FILENAME_PATTERN.fullmatch(filename):
        return "check_book"
    if filename_lower == CASH_FUND_BAL_FILENAME.lower():
        return "cash_fund_bal"
    if filename_lower == REVENUE_OBJ_FND_FILENAME.lower():
        return "revenue_ledger_obj_fnd"
    if filename_lower == REVENUE_LEDGER_FILENAME.lower():
        return "revenue_ledger"
    if PAYROLL_FILENAME_PATTERN.fullmatch(filename):
        return "payroll"
    # --------------------------------------------------------
    # Ignore temporary Excel files
    # --------------------------------------------------------
    if filename.startswith("~$"):
        return None
    # --------------------------------------------------------
    # Expense Detail
    # --------------------------------------------------------
    if (
        filename_lower
        == EXPENSE_DETAIL_FILENAME.lower()
    ):
        return "expense_detail"
    # --------------------------------------------------------
    # Non-Payroll Expenses
    # --------------------------------------------------------
    if (
        filename_lower
        == NON_PAYROLL_FILENAME.lower()
    ):
        return "non_payroll"
    return None
# ============================================================
# Read Expense Detail Excel
# ============================================================
def read_expense_detail_excel(content):
    excel_file = pd.ExcelFile(
        BytesIO(content),
        engine="openpyxl",
    )
    logger.info(
        "Expense Detail workbook sheets: %s",
        excel_file.sheet_names,
    )
    if (
        EXPENSE_DETAIL_SHEET_NAME
        not in excel_file.sheet_names
    ):
        raise ValueError(
            "Cincinnati Public Expense Detail is "
            "missing worksheet "
            f"{EXPENSE_DETAIL_SHEET_NAME!r}. "
            "Available sheets: "
            f"{excel_file.sheet_names}"
        )
    df = pd.read_excel(
        BytesIO(content),
        sheet_name=EXPENSE_DETAIL_SHEET_NAME,
        header=0,
        dtype=object,
        engine="openpyxl",
    )
    logger.info(
        "Expense Detail raw rows: %s",
        len(df),
    )
    logger.info(
        "Expense Detail incoming columns: %s",
        list(df.columns),
    )
    # --------------------------------------------------------
    # Remove completely blank rows
    # --------------------------------------------------------
    df = df.dropna(
        how="all"
    ).copy()
    # --------------------------------------------------------
    # Normalize headers
    # --------------------------------------------------------
    df.columns = [
        normalize_column_name(col)
        for col in df.columns
    ]
    logger.info(
        "Expense Detail normalized columns: %s",
        list(df.columns),
    )
    # --------------------------------------------------------
    # Apply aliases
    # --------------------------------------------------------
    rename_map = {}
    for col in df.columns:
        if (
            col
            in EXPENSE_DETAIL_COLUMN_ALIASES
        ):
            rename_map[col] = (
                EXPENSE_DETAIL_COLUMN_ALIASES[
                    col
                ]
            )
    df = df.rename(
        columns=rename_map
    )
    logger.info(
        "Expense Detail standardized columns: %s",
        list(df.columns),
    )
    # --------------------------------------------------------
    # Validate required columns
    # --------------------------------------------------------
    missing_columns = [
        col
        for col
        in EXPENSE_DETAIL_SOURCE_COLUMNS
        if col not in df.columns
    ]
    if missing_columns:
        raise ValueError(
            "Cincinnati Public Expense Detail is "
            "missing required columns: "
            f"{missing_columns}. "
            f"Actual columns: {list(df.columns)}"
        )
    # --------------------------------------------------------
    # Reject unexpected columns
    # --------------------------------------------------------
    unexpected_columns = [
        col
        for col in df.columns
        if col
        not in EXPENSE_DETAIL_SOURCE_COLUMNS
    ]
    if unexpected_columns:
        raise ValueError(
            "Cincinnati Public Expense Detail contains "
            "unexpected columns: "
            f"{unexpected_columns}. "
            f"Actual columns: {list(df.columns)}"
        )
    # --------------------------------------------------------
    # Final source column order
    # --------------------------------------------------------
    df = df[
        EXPENSE_DETAIL_SOURCE_COLUMNS
    ].copy()
    return df
# ============================================================
# Convert Expense Detail
# ============================================================
def convert_expense_detail_excel(content):
    df = read_expense_detail_excel(
        content
    )
    logger.info(
        "Expense Detail rows before conversion: %s",
        len(df),
    )
    # --------------------------------------------------------
    # Transaction Date
    # --------------------------------------------------------
    df["transaction_date"] = pd.to_datetime(
        df["transaction_date"],
        errors="coerce",
    )
    invalid_dates = int(
        df["transaction_date"]
        .isna()
        .sum()
    )
    if invalid_dates:
        logger.warning(
            "%s Expense Detail rows contain blank "
            "or invalid transaction dates.",
            invalid_dates,
        )
    # --------------------------------------------------------
    # Amount
    # --------------------------------------------------------
    df["amount"] = (
        df["amount"]
        .map(coerce_currency_decimal)
    )
    # --------------------------------------------------------
    # Text columns
    # --------------------------------------------------------
    for col in [
        "vendor_name",
        "usas_fund_description",
        "usas_func_description",
        "usas_object_description",
    ]:
        df[col] = clean_text_series(
            df[col]
        )
    # --------------------------------------------------------
    # Identifier columns
    # --------------------------------------------------------
    for col in [
        "check_number",
        "usas_fund",
        "usas_function",
        "usas_object",
    ]:
        df[col] = clean_identifier_series(
            df[col]
        )
    # --------------------------------------------------------
    # Fiscal Year
    # --------------------------------------------------------
    df["fiscal_year"] = (
        EXPENSE_DETAIL_FISCAL_YEAR
    )
    # --------------------------------------------------------
    # Final column order
    # --------------------------------------------------------
    df = df[
        EXPENSE_DETAIL_FINAL_COLUMNS
    ].copy()
    logger.info(
        "Final Expense Detail rows: %s",
        len(df),
    )
    return df.reset_index(
        drop=True
    )
# ============================================================
# Expense Detail Parquet Schema
# ============================================================
def get_expense_detail_schema():
    return pa.schema(
        [
            pa.field(
                "transaction_date",
                pa.timestamp("ns"),
            ),
            pa.field(
                "vendor_name",
                pa.string(),
            ),
            pa.field(
                "amount",
                pa.decimal128(20, 2),
            ),
            pa.field(
                "check_number",
                pa.string(),
            ),
            pa.field(
                "usas_fund",
                pa.string(),
            ),
            pa.field(
                "usas_fund_description",
                pa.string(),
            ),
            pa.field(
                "usas_function",
                pa.string(),
            ),
            pa.field(
                "usas_func_description",
                pa.string(),
            ),
            pa.field(
                "usas_object",
                pa.string(),
            ),
            pa.field(
                "usas_object_description",
                pa.string(),
            ),
            pa.field(
                "fiscal_year",
                pa.string(),
            ),
        ]
    )
# ============================================================
# Read Non-Payroll Excel
# ============================================================
def read_non_payroll_excel(content):
    excel_file = pd.ExcelFile(
        BytesIO(content),
        engine="openpyxl",
    )
    logger.info(
        "Non-Payroll workbook sheets: %s",
        excel_file.sheet_names,
    )
    if (
        NON_PAYROLL_SHEET_NAME
        not in excel_file.sheet_names
    ):
        raise ValueError(
            "Cincinnati Public Non-Payroll workbook "
            "is missing worksheet "
            f"{NON_PAYROLL_SHEET_NAME!r}. "
            "Available sheets: "
            f"{excel_file.sheet_names}"
        )
    df = pd.read_excel(
        BytesIO(content),
        sheet_name=NON_PAYROLL_SHEET_NAME,
        header=0,
        dtype=object,
        engine="openpyxl",
    )
    logger.info(
        "Non-Payroll raw rows: %s",
        len(df),
    )
    logger.info(
        "Non-Payroll incoming columns: %s",
        list(df.columns),
    )
    # --------------------------------------------------------
    # Remove completely blank rows
    # --------------------------------------------------------
    df = df.dropna(
        how="all"
    ).copy()
    # --------------------------------------------------------
    # Normalize headers
    # --------------------------------------------------------
    df.columns = [
        normalize_column_name(col)
        for col in df.columns
    ]
    logger.info(
        "Non-Payroll normalized columns: %s",
        list(df.columns),
    )
    # --------------------------------------------------------
    # Drop the NOTE column
    #
    # Original:
    #
    # NOTE: Check Amount duplicates for each check and PO
    # paid on that check.
    # --------------------------------------------------------
    note_columns = [
        col
        for col in df.columns
        if col.startswith(
            "note_check_amount_duplicates"
        )
    ]
    if note_columns:
        logger.info(
            "Dropping Non-Payroll note columns: %s",
            note_columns,
        )
        df = df.drop(
            columns=note_columns
        )
    # --------------------------------------------------------
    # Apply aliases
    # --------------------------------------------------------
    rename_map = {}
    for col in df.columns:
        if (
            col
            in NON_PAYROLL_COLUMN_ALIASES
        ):
            rename_map[col] = (
                NON_PAYROLL_COLUMN_ALIASES[
                    col
                ]
            )
    df = df.rename(
        columns=rename_map
    )
    logger.info(
        "Non-Payroll standardized columns: %s",
        list(df.columns),
    )
    # --------------------------------------------------------
    # Validate required columns
    # --------------------------------------------------------
    missing_columns = [
        col
        for col
        in NON_PAYROLL_SOURCE_COLUMNS
        if col not in df.columns
    ]
    if missing_columns:
        raise ValueError(
            "Cincinnati Public Non-Payroll workbook "
            "is missing required columns: "
            f"{missing_columns}. "
            f"Actual columns: {list(df.columns)}"
        )
    # --------------------------------------------------------
    # Reject unexpected columns
    # --------------------------------------------------------
    unexpected_columns = [
        col
        for col in df.columns
        if col
        not in NON_PAYROLL_SOURCE_COLUMNS
    ]
    if unexpected_columns:
        raise ValueError(
            "Cincinnati Public Non-Payroll workbook "
            "contains unexpected columns: "
            f"{unexpected_columns}. "
            f"Actual columns: {list(df.columns)}"
        )
    # --------------------------------------------------------
    # Final source column order
    # --------------------------------------------------------
    df = df[
        NON_PAYROLL_SOURCE_COLUMNS
    ].copy()
    return df
# ============================================================
# Convert Non-Payroll Excel
# ============================================================
def convert_non_payroll_excel(content):
    df = read_non_payroll_excel(
        content
    )
    logger.info(
        "Non-Payroll rows before conversion: %s",
        len(df),
    )
    # --------------------------------------------------------
    # Check Date
    # --------------------------------------------------------
    df["oh_ck_dt"] = pd.to_datetime(
        df["oh_ck_dt"],
        errors="coerce",
    )
    invalid_dates = int(
        df["oh_ck_dt"]
        .isna()
        .sum()
    )
    if invalid_dates:
        logger.warning(
            "%s Non-Payroll rows contain blank "
            "or invalid check dates.",
            invalid_dates,
        )
    # --------------------------------------------------------
    # Money Columns
    # --------------------------------------------------------
    for col in [
        "invoice_line_amount",
        "check_amount",
    ]:
        df[col] = (
            df[col]
            .map(coerce_currency_decimal)
        )
    # --------------------------------------------------------
    # Text Columns
    # --------------------------------------------------------
    for col in [
        "oh_pe_name",
        "pay_type",
        "key_fqa",
        "fqa",
    ]:
        df[col] = clean_text_series(
            df[col]
        )
    # --------------------------------------------------------
    # Identifier Columns
    # --------------------------------------------------------
    for col in [
        "oh_pe_id",
        "invoice_number",
        "po_number",
        "oh_check_no",
        "oh_gl_obj",
        "oh_check_id",
    ]:
        df[col] = clean_identifier_series(
            df[col]
        )
    # --------------------------------------------------------
    # Fiscal Year
    # --------------------------------------------------------
    df["fiscal_year"] = (
        NON_PAYROLL_FISCAL_YEAR
    )
    # --------------------------------------------------------
    # Final Column Order
    # --------------------------------------------------------
    df = df[
        NON_PAYROLL_FINAL_COLUMNS
    ].copy()
    logger.info(
        "Final Non-Payroll rows: %s",
        len(df),
    )
    return df.reset_index(
        drop=True
    )
# ============================================================
# Non-Payroll Parquet Schema
# ============================================================
def get_non_payroll_schema():
    return pa.schema(
        [
            pa.field(
                "oh_pe_id",
                pa.string(),
            ),
            pa.field(
                "oh_pe_name",
                pa.string(),
            ),
            pa.field(
                "invoice_number",
                pa.string(),
            ),
            pa.field(
                "invoice_line_amount",
                pa.decimal128(20, 2),
            ),
            pa.field(
                "po_number",
                pa.string(),
            ),
            pa.field(
                "pay_type",
                pa.string(),
            ),
            pa.field(
                "oh_check_no",
                pa.string(),
            ),
            pa.field(
                "oh_ck_dt",
                pa.timestamp("ns"),
            ),
            pa.field(
                "key_fqa",
                pa.string(),
            ),
            pa.field(
                "oh_gl_obj",
                pa.string(),
            ),
            pa.field(
                "fqa",
                pa.string(),
            ),
            pa.field(
                "check_amount",
                pa.decimal128(20, 2),
            ),
            pa.field(
                "oh_check_id",
                pa.string(),
            ),
            pa.field(
                "fiscal_year",
                pa.string(),
            ),
        ]
    )
# ============================================================
# Write Parquet
# ============================================================
def write_parquet(
    df,
    schema,
):
    output = BytesIO()
    table = pa.Table.from_pandas(
        df,
        schema=schema,
        preserve_index=False,
        safe=True,
    )
    pq.write_table(
        table,
        output,
        compression="snappy",
    )
    return output.getvalue()
# ============================================================
# Process One S3 Object
# ============================================================

# General importer: recognize schemas, not exact filenames or sheet titles.
def infer_fiscal_year(sheet_name, filename, value=None):
    if value is not None and str(value).strip():
        text = re.sub(r"\.0$", "", str(value).strip())
        if re.fullmatch(r"20[0-9]{2}", text):
            return text
        raise ValueError(f"Invalid fiscal_year value: {value!r}")
    for text in (sheet_name, filename):
        match = re.search(r"FY\s*[-_ ]*(20[0-9]{2}|[0-9]{2})(?![0-9])", text, re.I)
        if match:
            year = match.group(1)
            return "20" + year if len(year) == 2 else year
        match = re.search(r"(?<![0-9])(20[0-9]{2})(?![0-9])", text)
        if text == sheet_name and match:
            return match.group(1)
    raise ValueError(f"Fiscal year not found in tab {sheet_name!r} or filename {filename!r}; "
                     "include FY24/FY2024 in the filename or a fiscal_year column")


def import_definitions():
    return [
        (CHECK_BOOK_TABLE_NAME, CHECK_BOOK_DEST_PREFIX, CHECK_BOOK_HEADERS,
         EXPENSE_DETAIL_SOURCE_COLUMNS, get_check_book_schema()),
        (PAYROLL_TABLE_NAME, PAYROLL_DEST_PREFIX, PAYROLL_SOURCE_HEADERS,
         PAYROLL_COLUMNS, get_payroll_schema()),
        (REVENUE_LEDGER_TABLE_NAME, REVENUE_LEDGER_DEST_PREFIX, REVENUE_LEDGER_HEADERS,
         REVENUE_LEDGER_COLUMNS, get_revenue_ledger_schema()),
        (REVENUE_OBJ_FND_TABLE_NAME, REVENUE_OBJ_FND_DEST_PREFIX, REVENUE_OBJ_FND_HEADERS,
         REVENUE_OBJ_FND_COLUMNS, get_revenue_obj_fnd_schema()),
        (CASH_FUND_BAL_TABLE_NAME, CASH_FUND_BAL_DEST_PREFIX, CASH_FUND_BAL_HEADERS,
         CASH_FUND_BAL_COLUMNS, get_cash_fund_bal_schema()),
        (NON_PAYROLL_REPORT_TYPE, NON_PAYROLL_DEST_PREFIX, NON_PAYROLL_SOURCE_COLUMNS,
         NON_PAYROLL_SOURCE_COLUMNS, get_non_payroll_schema()),
    ]


def match_import_headers(headers, definitions):
    normalized = [normalize_column_name(v) if v is not None else "" for v in headers]
    while normalized and not normalized[-1]:
        normalized.pop()
    for table, prefix, source, target, schema in definitions:
        aliases = dict(zip([normalize_column_name(v) for v in source], target))
        aliases.update({col: col for col in target})
        if table == NON_PAYROLL_REPORT_TYPE:
            aliases.update(NON_PAYROLL_COLUMN_ALIASES)
        mapped = [aliases.get(v, v) for v in normalized]
        if table == NON_PAYROLL_REPORT_TYPE:
            # Ignore explanatory NOTE columns, retaining their positions so
            # a note inserted between data fields cannot shift row values.
            mapped = [None if v == "note" or v.startswith("note_") else v
                      for v in mapped]
        active = [v for v in mapped if v is not None]
        data = [v for v in active if v != "fiscal_year"]
        if len(data) == len(target) and set(data) == set(target) and len(active) == len(set(active)):
            return (table, prefix, target, schema), mapped
    return None


def flexible_record(values, columns, definition, sheet_name, row_number, filename):
    from datetime import date, datetime, time
    from openpyxl.utils.datetime import from_excel
    table, prefix, target, schema = definition
    row = {col: value for col, value in zip(columns, values) if col is not None}
    if table == CASH_FUND_BAL_TABLE_NAME and row.get("fund") in (None, "") and row.get("fund_description") in (None, ""):
        return None  # Separate grand total, not a fund record.
    result = {}
    for field in schema:
        col = field.name
        value = row.get(col)
        if col == "fiscal_year":
            value = infer_fiscal_year(sheet_name, filename, value)
        elif col == "source_sheet":
            value = sheet_name
        elif col == "source_row":
            value = row_number
        elif col == "quarter":
            match = re.search(r"(?:Q|QUARTER)\s*([1-4])", sheet_name, re.I) or re.search(r"(?:Q|QUARTER)\s*([1-4])", filename, re.I)
            value = "Q" + match.group(1) if match else None
        elif pa.types.is_decimal(field.type):
            value = coerce_currency_decimal(value)
        elif pa.types.is_timestamp(field.type):
            if value is None or value == "":
                value = None
            elif isinstance(value, (int, float)):
                value = from_excel(value)
            elif isinstance(value, datetime):
                pass
            elif isinstance(value, date):
                value = datetime.combine(value, time())
            else:
                value = pd.to_datetime(value, errors="raise").to_pydatetime()
        elif value is not None:
            value = str(value).strip()
            if col not in {"description", "vendor_name", "employee_name"}:
                value = re.sub(r"\.0$", "", value)
            if table == CASH_FUND_BAL_TABLE_NAME and col == "fund":
                value = value.zfill(3)
        result[col] = value
    return result



def prepare_import_storage(s3, bucket, key):
    """Clean only this importer's stale folders and check download capacity."""
    import shutil
    import tempfile
    from pathlib import Path
    root = Path(tempfile.gettempdir())
    # A Lambda execution environment handles one invocation at a time. A killed
    # invocation can leave TemporaryDirectory contents for the next invocation.
    prefixes = ("cincinnati-import-", "cincinnati-check-book-", "cincinnati-payroll-")
    for child in root.iterdir():
        if child.is_dir() and not child.is_symlink() and child.name.startswith(prefixes):
            logger.info("Removing stale importer folder: %s", child.name)
            shutil.rmtree(child)
    size = s3.head_object(Bucket=bucket, Key=key)["ContentLength"]
    usage = shutil.disk_usage(root)
    logger.info("Import storage: source_bytes=%s total_bytes=%s free_bytes=%s key=%s",
                size, usage.total, usage.free, key)
    # Leave a small working margin; output size is unknown until conversion.
    margin = 64 * 1024 * 1024
    if size + margin > usage.free:
        raise ValueError(
            f"Insufficient temporary disk for {key!r}: source={size / 1024**2:.1f} MB, "
            f"free={usage.free / 1024**2:.1f} MB, total={usage.total / 1024**2:.1f} MB. "
            "Set this Lambda's Ephemeral storage to 10240 MB. If already set, "
            "the source exceeds available disk and needs splitting or processing outside Lambda."
        )

def convert_flexible_s3_object(bucket, key, s3=None):
    import csv
    import hashlib
    import tempfile
    from pathlib import Path
    from openpyxl import load_workbook
    validate_source_path(key)
    if bucket != RAW_BUCKET:
        raise ValueError(f"Unexpected source bucket: {bucket!r}")
    filename = PurePosixPath(key).name
    extension = PurePosixPath(key).suffix.lower()
    if filename.startswith("~$") or extension not in {".xlsx", ".csv"}:
        logger.info("Skipping unsupported file: %s", filename)
        return None
    if PAYROLL_BATCH_ROWS < 1:
        raise ValueError("PAYROLL_BATCH_ROWS must be positive")
    s3 = s3 or boto3.client("s3", region_name=AWS_REGION)
    prepare_import_storage(s3, bucket, key)
    definitions = import_definitions()
    results = []
    with tempfile.TemporaryDirectory(prefix="cincinnati-import-") as directory:
        source_path = str(Path(directory) / ("source" + extension))
        s3.download_file(bucket, key, source_path)
        workbook = None
        csv_file = None
        if extension == ".xlsx":
            workbook = load_workbook(source_path, read_only=True, data_only=True)
            sheets = [(sheet.title, sheet.iter_rows(values_only=True)) for sheet in workbook]
        else:
            # Validate encoding without loading the CSV into memory.
            encoding = "utf-8-sig"
            try:
                with open(source_path, encoding=encoding) as probe:
                    for line in probe:
                        pass
            except UnicodeDecodeError:
                encoding = "cp1252"
            csv_file = open(source_path, encoding=encoding, newline="")
            sample = csv_file.read(65536)
            csv_file.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            except csv.Error:
                dialect = csv.excel
            sheets = [("CSV", csv.reader(csv_file, dialect))]
        writers = {}
        outputs = {}
        observed = {}
        try:
            for sheet_name, rows in sheets:
                matched = None
                preview = []
                for row_number, headers in enumerate(rows, 1):
                    if len(preview) < 3:
                        preview.append([str(v) for v in headers])
                    matched = match_import_headers(headers, definitions)
                    if matched or row_number >= 25:
                        break
                if not matched:
                    observed[sheet_name] = preview
                    logger.info("Skipping non-data tab %s: %s", sheet_name, preview)
                    continue
                definition, columns = matched
                table, prefix, target, schema = definition
                header_row = row_number
                if table not in writers:
                    path = str(Path(directory) / (table + ".parquet"))
                    writers[table] = pq.ParquetWriter(path, schema, compression="snappy")
                    outputs[table] = {"path": path, "prefix": prefix, "rows": 0, "sheet_rows": {}}
                batch = []
                count = 0
                for row_number, values in enumerate(rows, header_row + 1):
                    if all(v is None or v == "" for v in values):
                        continue
                    if any(v is not None and v != "" for v in values[len(columns):]):
                        raise ValueError(f"Unexpected extra values in {sheet_name!r}, row {row_number}")
                    try:
                        record = flexible_record(values, columns, definition, sheet_name, row_number, filename)
                    except Exception as exc:
                        raise ValueError(f"Invalid {table} record in tab {sheet_name!r}, row {row_number}: {exc}") from exc
                    if record is None:
                        continue
                    batch.append(record)
                    count += 1
                    if len(batch) >= PAYROLL_BATCH_ROWS:
                        writers[table].write_table(pa.Table.from_pylist(batch, schema=schema))
                        batch.clear()
                if batch:
                    writers[table].write_table(pa.Table.from_pylist(batch, schema=schema))
                outputs[table]["rows"] += count
                outputs[table]["sheet_rows"][sheet_name] = count
                logger.info("Imported %s tab %s: %s rows", table, sheet_name, count)
        finally:
            for writer in writers.values():
                writer.close()
            if workbook:
                workbook.close()
            if csv_file:
                csv_file.close()
        if not outputs:
            raise ValueError(f"No supported report columns found in {filename!r}. Observed tabs/headers: {observed!r}")
        suffix = hashlib.sha256(key.encode("utf-8")).hexdigest()[:16]
        for table, output in outputs.items():
            dest_key = f"{output['prefix']}/{PurePosixPath(key).stem}_{suffix}.parquet"
            s3.upload_file(output["path"], DEST_BUCKET, dest_key)
            results.append({"source": f"s3://{bucket}/{key}",
                            "destination": f"s3://{DEST_BUCKET}/{dest_key}",
                            "table_name": table, "report_type": table,
                            "rows": output["rows"], "sheet_rows": output["sheet_rows"]})
    return results[0] if len(results) == 1 else {"source": f"s3://{bucket}/{key}", "exports": results}

def legacy_convert_s3_object(
    bucket,
    key,
    s3=None,
):
    # --------------------------------------------------------
    # Validate source bucket
    # --------------------------------------------------------
    if bucket != RAW_BUCKET:
        raise ValueError(
            f"Source bucket {bucket!r} does not match "
            f"RAW_BUCKET {RAW_BUCKET!r}"
        )
    # --------------------------------------------------------
    # Validate Cincinnati Public source path
    # --------------------------------------------------------
    validate_source_path(
        key
    )
    filename = PurePosixPath(
        key
    ).name
    # --------------------------------------------------------
    # Skip temporary Excel files
    # --------------------------------------------------------
    if filename.startswith("~$"):
        logger.info(
            "Skipping temporary Excel file: "
            "s3://%s/%s",
            bucket,
            key,
        )
        return None
    # --------------------------------------------------------
    # Determine report type
    # --------------------------------------------------------
    report_type = report_type_for_key(
        key
    )
    if report_type is None:
        logger.info(
            "Skipping unrecognized Cincinnati Public "
            "file: %r",
            filename,
        )
        return None
    logger.info(
        "Recognized Cincinnati Public report type: %s",
        report_type,
    )
    # --------------------------------------------------------
    # Create S3 client
    # --------------------------------------------------------
    if s3 is None:
        s3 = boto3.client(
            "s3",
            region_name=AWS_REGION,
        )
    # --------------------------------------------------------
    # Download source
    # --------------------------------------------------------
    if report_type == "check_book":
        return convert_check_book_s3_object(bucket, key, s3)
    if report_type == "payroll":
        return convert_payroll_s3_object(bucket, key, s3)
    response = s3.get_object(
        Bucket=bucket,
        Key=key,
    )
    content = response[
        "Body"
    ].read()
    logger.info(
        "Downloaded source: s3://%s/%s",
        bucket,
        key,
    )
    # ========================================================
    # EXPENSE DETAIL
    # ========================================================
    if report_type == "expense_detail":
        logger.info(
            "Processing Cincinnati Public "
            "Expense Detail."
        )
        df = convert_expense_detail_excel(
            content
        )
        schema = (
            get_expense_detail_schema()
        )
        dest_prefix = (
            EXPENSE_DETAIL_DEST_PREFIX
        )
        fiscal_year = (
            EXPENSE_DETAIL_FISCAL_YEAR
        )
        output_report_type = (
            EXPENSE_DETAIL_REPORT_TYPE
        )
        table_name = (
            "cincinnati_public_expense_detail"
        )
    # ========================================================
    # NON-PAYROLL EXPENSES
    # ========================================================
    elif report_type == "non_payroll":
        logger.info(
            "Processing Cincinnati Public "
            "Non-Payroll Expenses."
        )
        df = convert_non_payroll_excel(
            content
        )
        schema = (
            get_non_payroll_schema()
        )
        dest_prefix = (
            NON_PAYROLL_DEST_PREFIX
        )
        fiscal_year = (
            NON_PAYROLL_FISCAL_YEAR
        )
        output_report_type = (
            NON_PAYROLL_REPORT_TYPE
        )
        table_name = (
            "cincinnati_public_non_payroll_expenses"
        )
    elif report_type == "revenue_ledger":
        df = convert_revenue_ledger_excel(content)
        schema = get_revenue_ledger_schema()
        dest_prefix = REVENUE_LEDGER_DEST_PREFIX
        fiscal_year = "2025"
        output_report_type = REVENUE_LEDGER_TABLE_NAME
        table_name = REVENUE_LEDGER_TABLE_NAME
    elif report_type == "revenue_ledger_obj_fnd":
        df = convert_revenue_obj_fnd_excel(content)
        schema = get_revenue_obj_fnd_schema()
        dest_prefix = REVENUE_OBJ_FND_DEST_PREFIX
        fiscal_year = "2025"
        output_report_type = REVENUE_OBJ_FND_TABLE_NAME
        table_name = REVENUE_OBJ_FND_TABLE_NAME
    elif report_type == "cash_fund_bal":
        df = convert_cash_fund_bal_excel(content)
        schema = get_cash_fund_bal_schema()
        dest_prefix = CASH_FUND_BAL_DEST_PREFIX
        fiscal_year = "2025"
        output_report_type = CASH_FUND_BAL_TABLE_NAME
        table_name = CASH_FUND_BAL_TABLE_NAME
    else:
        raise ValueError(
            "Unhandled Cincinnati Public report type: "
            f"{report_type}"
        )
    # --------------------------------------------------------
    # Destination filename
    # --------------------------------------------------------
    dest_filename = (
        f"{PurePosixPath(key).stem}.parquet"
    )
    dest_key = (
        f"{dest_prefix}/"
        f"{dest_filename}"
    )
    # --------------------------------------------------------
    # Convert to Parquet
    # --------------------------------------------------------
    parquet_bytes = write_parquet(
        df=df,
        schema=schema,
    )
    logger.info(
        "Generated Parquet size: %s bytes",
        len(parquet_bytes),
    )
    # --------------------------------------------------------
    # Upload to transformed bucket
    # --------------------------------------------------------
    s3.put_object(
        Bucket=DEST_BUCKET,
        Key=dest_key,
        Body=parquet_bytes,
    )
    logger.info(
        "Successfully processed Cincinnati Public file."
    )
    logger.info(
        "Rows written: %s",
        len(df),
    )
    logger.info(
        "Destination: s3://%s/%s",
        DEST_BUCKET,
        dest_key,
    )
    logger.info(
        "Athena/Glue table: %s",
        table_name,
    )
    # --------------------------------------------------------
    # Return result
    # --------------------------------------------------------
    return {
        "source": (
            f"s3://{bucket}/{key}"
        ),
        "destination": (
            f"s3://{DEST_BUCKET}/{dest_key}"
        ),
        "rows": len(df),
        "fiscal_year": fiscal_year,
        "report_type": output_report_type,
        "table_name": table_name,
    }
# ============================================================
# Lambda Handler
#
# Direct S3 Trigger
# ============================================================
def convert_s3_object(bucket, key, s3=None):
    return convert_flexible_s3_object(bucket, key, s3)

# Direct continuation via asynchronous Lambda invocation; no SQS.
# One chunk is uploaded before another invocation is requested.
PART_ROWS = int(os.getenv("IMPORT_PART_ROWS", "50000"))
MIN_REMAINING_MS = int(os.getenv("IMPORT_MIN_REMAINING_MS", "180000"))


def process_partitioned_workbook(bucket, key, context, cursor=None, s3=None):
    """Convert one bounded chunk, upload it, then resume with direct Lambda invocation.

    Cursor is a worksheet index and absolute Excel row number. No transaction
    deduplication occurs. An identical retry overwrites the same deterministic part.
    """
    import hashlib
    import tempfile
    from pathlib import Path
    from openpyxl import load_workbook

    validate_source_path(key)
    if bucket != RAW_BUCKET:
        raise ValueError(f"Unexpected source bucket: {bucket}")
    if PART_ROWS < 1:
        raise ValueError("IMPORT_PART_ROWS must be positive")
    s3 = s3 or boto3.client("s3", region_name=AWS_REGION)
    head = s3.head_object(Bucket=bucket, Key=key)
    version = head.get("VersionId") or head.get("ETag", "").strip('"')
    source_hash = hashlib.sha256((key + "|" + version).encode()).hexdigest()[:20]
    filename = PurePosixPath(key).name
    sheet_index = int((cursor or {}).get("sheet_index", 0))
    start_row = int((cursor or {}).get("start_row", 1))
    if cursor and cursor.get("version") != version:
        raise ValueError("Source workbook changed during import; restart from S3 event")

    with tempfile.TemporaryDirectory(prefix="cincinnati-import-") as directory:
        source_path = str(Path(directory) / "source.xlsx")
        s3.download_file(bucket, key, source_path)
        wb = load_workbook(source_path, read_only=True, data_only=True)
        try:
            definitions = import_definitions()
            while sheet_index < len(wb.worksheets):
                sheet = wb.worksheets[sheet_index]
                iterator = enumerate(sheet.iter_rows(values_only=True), start=1)
                matched = None
                for row_num, headers in iterator:
                    matched = match_import_headers(headers, definitions)
                    if matched or row_num >= 25:
                        break
                if not matched:
                    logger.info("Skipping unsupported worksheet %s", sheet.title)
                    sheet_index += 1
                    start_row = 1
                    continue
                definition, columns = matched
                table, prefix, _, schema = definition
                first_data_row = row_num + 1
                effective_start = max(first_data_row, start_row)
                # Stable key per source, worksheet, and row interval. Parts from
                # different source files do not collide in one Athena table.
                records = []
                last_row = effective_start - 1
                for row_number, values in iterator:
                    if row_number < effective_start:
                        continue
                    last_row = row_number
                    if all(v is None or v == "" for v in values):
                        continue
                    if any(v is not None and v != "" for v in values[len(columns):]):
                        raise ValueError(f"Extra data in {sheet.title} row {row_number}")
                    record = flexible_record(values, columns, definition, sheet.title, row_number, filename)
                    if record is not None:
                        records.append(record)
                    if len(records) >= PART_ROWS:
                        break
                    if context and context.get_remaining_time_in_millis() < MIN_REMAINING_MS:
                        break
                if records:
                    # Use stable part names; retries overwrite the same part.
                    dest_key = (f"{prefix}/{source_hash}/"
                                f"sheet_{sheet_index:03d}_rows_{effective_start:08d}_{last_row:08d}.parquet")
                    local_out = str(Path(directory) / "part.parquet")
                    pq.write_table(pa.Table.from_pylist(records, schema=schema), local_out, compression="snappy")
                    s3.upload_file(local_out, DEST_BUCKET, dest_key)
                    logger.info("Uploaded %s rows from %s to s3://%s/%s", len(records), sheet.title, DEST_BUCKET, dest_key)
                # If no records remain, advance worksheet. If we hit a chunk
                # limit, continue at the next Excel row (including blank rows).
                if len(records) >= PART_ROWS or (context and context.get_remaining_time_in_millis() < MIN_REMAINING_MS):
                    next_cursor = {"sheet_index": sheet_index, "start_row": last_row + 1, "version": version}
                else:
                    next_cursor = {"sheet_index": sheet_index + 1, "start_row": 1, "version": version}
                if next_cursor["sheet_index"] >= len(wb.worksheets):
                    logger.info("Completed import of s3://%s/%s", bucket, key)
                    return {"status": "complete", "last_destination": dest_key if records else None}
                # Invoke after successful upload, not before.
                fn = os.environ.get("AWS_LAMBDA_FUNCTION_NAME")
                if not fn:
                    raise ValueError("AWS_LAMBDA_FUNCTION_NAME is required for continuation")
                client = boto3.client("lambda", region_name=AWS_REGION)
                response = client.invoke(FunctionName=fn, InvocationType="Event",
                    Payload=json.dumps({"continuation": {"bucket": bucket, "key": key, "cursor": next_cursor}}).encode())
                if response["StatusCode"] != 202:
                    raise RuntimeError(f"Continuation invoke failed: {response}")
                logger.info("Queued direct Lambda continuation: %s", next_cursor)
                return {"status": "continuing", "cursor": next_cursor}
            logger.info("Completed import of s3://%s/%s", bucket, key)
            return {"status": "complete"}
        finally:
            wb.close()


def lambda_handler(event, context):
    """S3 upload or direct continuation. No SQS needed."""
    if "continuation" in event:
        job = event["continuation"]
        return process_partitioned_workbook(job["bucket"], job["key"], context, job["cursor"])
    results = []
    for record in event.get("Records", []):
        if record.get("eventSource") != "aws:s3" or "s3" not in record:
            continue
        bucket = record["s3"]["bucket"]["name"]
        key = unquote_plus(record["s3"]["object"]["key"])
        if key.lower().endswith(".xlsx") and not PurePosixPath(key).name.startswith("~$"):
            results.append(process_partitioned_workbook(bucket, key, context))
        else:
            results.append(convert_flexible_s3_object(bucket, key))
    return {"statusCode": 200, "body": json.dumps(results, default=str)}
