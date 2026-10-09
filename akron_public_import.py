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

AWS_REGION = os.getenv("AWS_REGION", "us-east-2")
RAW_BUCKET = os.getenv("RAW_BUCKET", "schools-raw").strip()
DEST_BUCKET = os.getenv("DEST_BUCKET", "schools-transformed").strip()

EXPENSE_DEST_PREFIX = os.getenv(
    "AKRON_PUBLIC_EXPENSE_DETAIL_DEST_PREFIX",
    "data_from_clients/akron_public_expense_detail/",
).strip("/")

REVENUE_DEST_PREFIX = os.getenv(
    "AKRON_PUBLIC_REVENUE_DETAIL_DEST_PREFIX",
    "data_from_clients/akron_public_revenue_detail/",
).strip("/")

CASH_DEST_PREFIX = os.getenv(
    "AKRON_PUBLIC_CASH_REPORT_DEST_PREFIX",
    "data_from_clients/akron_public_cash_report/",
).strip("/")

SHEET_NAME = "ACCOUNT DETAIL"

CASH_COLUMNS = [
    "page",
    "fund",
    "account",
    "description",
    "row_type",
    "period",
    "beginning_balance",
    "receipts",
    "expenses",
    "encumbrances",
    "ending_cash_balance",
]

CASH_MONEY_COLUMNS = [
    "beginning_balance",
    "receipts",
    "expenses",
    "encumbrances",
    "ending_cash_balance",
]

# ============================================================
# EXPENSE CONFIGURATION
# ============================================================
#
# Optional source columns:
#   spc_cost_ctr
#   encumbrances
#   src
#   voucher
#   doc_number
#   ar_ref
#
# Missing optional columns are added as blank columns so all
# expense Parquet files have the same schema.
# ============================================================

EXPENSE_COLUMNS = [
    "fund",
    "spc_cost_ctr",
    "ti_function",
    "object",
    "account",
    "account_description",
    "original_approp",
    "tranfrs_adjsmts",
    "revised_budget",
    "ytd_expended",
    "encumbrances",
    "available_budget",
    "percent_used",
    "yr_per_jnl",
    "eff_date",
    "src",
    "amount",
    "vendor_ref_1",
    "ref_2",
    "ref_3",
    "ref_4",
    "po_number",
    "voucher",
    "invoice_number",
    "doc_number",
    "check_number",
    "vendor_name",
    "comment",
    "ar_ref",
    "item_description",
]

EXPENSE_REQUIRED_COLUMNS = [
    "fund",
    "ti_function",
    "object",
    "account",
    "account_description",
    "original_approp",
    "tranfrs_adjsmts",
    "revised_budget",
    "ytd_expended",
    "available_budget",
    "percent_used",
    "yr_per_jnl",
    "eff_date",
    "amount",
    "vendor_ref_1",
    "ref_2",
    "ref_3",
    "ref_4",
    "po_number",
    "invoice_number",
    "check_number",
    "vendor_name",
    "comment",
    "item_description",
]

EXPENSE_OPTIONAL_COLUMNS = [
    "spc_cost_ctr",
    "encumbrances",
    "src",
    "voucher",
    "doc_number",
    "ar_ref",
]

EXPENSE_ACCOUNT_COLUMNS = [
    "fund",
    "spc_cost_ctr",
    "ti_function",
    "object",
    "account",
    "account_description",
]

EXPENSE_SUMMARY_COLUMNS = [
    "original_approp",
    "tranfrs_adjsmts",
    "revised_budget",
    "ytd_expended",
    "encumbrances",
    "available_budget",
    "percent_used",
]

EXPENSE_FILL_DOWN_COLUMNS = EXPENSE_ACCOUNT_COLUMNS + EXPENSE_SUMMARY_COLUMNS

EXPENSE_MONEY_COLUMNS = [
    "original_approp",
    "tranfrs_adjsmts",
    "revised_budget",
    "ytd_expended",
    "encumbrances",
    "available_budget",
    "amount",
]

EXPENSE_TEXT_COLUMNS = [
    "fund",
    "spc_cost_ctr",
    "ti_function",
    "object",
    "account",
    "account_description",
    "yr_per_jnl",
    "src",
    "vendor_ref_1",
    "ref_2",
    "ref_3",
    "ref_4",
    "po_number",
    "voucher",
    "invoice_number",
    "doc_number",
    "check_number",
    "vendor_name",
    "comment",
    "ar_ref",
    "item_description",
]

# ============================================================
# REVENUE CONFIGURATION
# ============================================================
#
# Optional source columns:
#   spc_cost_ctr
#   obj
#   src
#   ar_ref
#
# Missing optional columns are added as blank columns so all
# revenue Parquet files have the same schema.
# ============================================================

REVENUE_COLUMNS = [
    "fund",
    "spc_cost_ctr",
    "ti_function",
    "obj",
    "account",
    "account_description",
    "original_estim_rev",
    "estim_rev_adj",
    "revised_estim_rev",
    "actual_ytd_revenue",
    "remaining_revenue",
    "percent_coll",
    "yr_per_jnl",
    "eff_date",
    "src",
    "amount",
    "vendor_ref_1",
    "ref_2",
    "ref_3",
    "ref_4",
    "invoice_number",
    "check_number",
    "vendor_name",
    "comment",
    "ar_ref",
    "item_description",
]

REVENUE_REQUIRED_COLUMNS = [
    "fund",
    "ti_function",
    "account",
    "account_description",
    "original_estim_rev",
    "estim_rev_adj",
    "revised_estim_rev",
    "actual_ytd_revenue",
    "remaining_revenue",
    "percent_coll",
    "yr_per_jnl",
    "eff_date",
    "amount",
    "vendor_ref_1",
    "ref_2",
    "ref_3",
    "ref_4",
    "invoice_number",
    "check_number",
    "vendor_name",
    "comment",
    "item_description",
]

REVENUE_OPTIONAL_COLUMNS = [
    "spc_cost_ctr",
    "obj",
    "src",
    "ar_ref",
]

REVENUE_ACCOUNT_COLUMNS = [
    "fund",
    "spc_cost_ctr",
    "ti_function",
    "obj",
    "account",
    "account_description",
]

REVENUE_SUMMARY_COLUMNS = [
    "original_estim_rev",
    "estim_rev_adj",
    "revised_estim_rev",
    "actual_ytd_revenue",
    "remaining_revenue",
    "percent_coll",
]

REVENUE_FILL_DOWN_COLUMNS = REVENUE_ACCOUNT_COLUMNS + REVENUE_SUMMARY_COLUMNS

REVENUE_MONEY_COLUMNS = [
    "original_estim_rev",
    "estim_rev_adj",
    "revised_estim_rev",
    "actual_ytd_revenue",
    "remaining_revenue",
    "amount",
]

REVENUE_TEXT_COLUMNS = [
    "fund",
    "spc_cost_ctr",
    "ti_function",
    "obj",
    "account",
    "account_description",
    "yr_per_jnl",
    "src",
    "vendor_ref_1",
    "ref_2",
    "ref_3",
    "ref_4",
    "invoice_number",
    "check_number",
    "vendor_name",
    "comment",
    "ar_ref",
    "item_description",
]

# ============================================================
# Normalize Excel column names
# ============================================================


def normalize_report_col_name(name):
    s = str(name).strip().lower()
    s = s.replace("%", "percent")
    s = s.replace("#", "number")
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_")


# ============================================================
# Exact money conversion
# ============================================================


def coerce_currency_decimal(value):
    """
    Convert monetary values to Decimal(20,2).

    Handles:
        1,234.56
        $1,234.56
        (1,234.56)
        -1234.56
        1.3942830474E8
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

    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = "-" + cleaned[1:-1]

    cleaned = re.sub(r"[$,\s]", "", cleaned)

    try:
        return Decimal(cleaned).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"Invalid Akron Public monetary value: {value!r}") from exc


# ============================================================
# Numeric conversion
# ============================================================


def coerce_numeric_series(series):
    cleaned = series.astype("string").str.strip()

    cleaned = cleaned.str.replace(
        r"^\((.*)\)$",
        r"-\1",
        regex=True,
    )

    cleaned = cleaned.str.replace(
        r"[$,%\s,]",
        "",
        regex=True,
    )

    return pd.to_numeric(
        cleaned,
        errors="coerce",
    ).astype("float64")


# ============================================================
# Text cleanup
# ============================================================


def clean_text_series(series):
    return series.map(
        lambda value: None if pd.isna(value) else str(value).strip()
    ).astype("string")


# ============================================================
# Identifier cleanup
# ============================================================


def clean_identifier_series(series):
    cleaned = clean_text_series(series)
    cleaned = cleaned.str.replace(
        r"\.0$",
        "",
        regex=True,
    )
    return cleaned


# ============================================================
# Fiscal year from filename
# ============================================================


def fiscal_year_for_key(key):
    """
    Supports:
        FY 2025
        FY2025
        FY 25
        FY25
    """
    filename = PurePosixPath(key).name

    match = re.search(
        r"\bFY\s*(20\d{2})\b",
        filename,
        flags=re.IGNORECASE,
    )
    if match:
        return match.group(1)

    match = re.search(
        r"\bFY\s*(\d{2})\b",
        filename,
        flags=re.IGNORECASE,
    )
    if match:
        return "20" + match.group(1)

    raise ValueError(
        "Could not determine fiscal year from Akron "
        f"Public filename: {filename}"
    )


# ============================================================
# Validate source path
# ============================================================


def validate_source_path(key):
    """
    Accept both versions of the Akron source path.

    Current S3 path:
        data-from-clients/akron_public/

    Legacy/alternate path:
        data_from_clients/akron_public/
    """
    normalized_key = key.strip("/").lower()
    valid_prefixes = [
        "data-from-clients/akron_public/",
        "data_from_clients/akron_public/",
    ]

    if not any(normalized_key.startswith(prefix) for prefix in valid_prefixes):
        raise ValueError(
            "Unexpected Akron Public source path: "
            f"{key}. Expected one of: {valid_prefixes}"
        )


# ============================================================
# Determine report type
# ============================================================


def report_type_for_key(key):
    """
    Supported reports:
        Detailed Expense Report
        Detailed Revenue Report
        Extracted Cash Report CSV
    """
    filename = PurePosixPath(key).name.lower()
    normalized = filename.replace("_", " ").replace("-", " ")
    normalized = re.sub(r"\s+", " ", normalized).strip()

    if filename.endswith(".csv") and (
        "cash report" in normalized
        or "revenues expenses extracted" in normalized
    ):
        return "cash_report"

    if "detailed expense report" in normalized:
        return "expense_detail"

    if "detailed revenue report" in normalized:
        return "revenue_detail"

    raise ValueError(
        "Unrecognized Akron Public report filename: "
        f"{key}"
    )


# ============================================================
# Read ACCOUNT DETAIL worksheet
# ============================================================


def read_account_detail_excel(
    content,
    expected_columns,
    report_name,
    required_columns,
    optional_columns,
):
    """
    Required columns must exist.
    Optional columns may or may not exist.
    Missing optional columns are created as blank columns.
    Unexpected columns still raise an error so that genuinely
    new report layouts do not silently pass unnoticed.
    """
    df = pd.read_excel(
        BytesIO(content),
        sheet_name=SHEET_NAME,
        header=1,
        dtype=object,
        engine="openpyxl",
    )

    df.columns = [
        normalize_report_col_name(col)
        for col in df.columns
    ]

    actual_columns = list(df.columns)
    logger.info(
        "%s incoming columns: %s",
        report_name,
        actual_columns,
    )

    missing_required = [
        col
        for col in required_columns
        if col not in actual_columns
    ]
    if missing_required:
        raise ValueError(
            f"Akron Public {report_name} is missing "
            f"required columns: {missing_required}. "
            f"Actual columns: {actual_columns}"
        )

    allowed_columns = set(required_columns + optional_columns)
    unexpected_columns = [
        col
        for col in actual_columns
        if col not in allowed_columns
    ]
    if unexpected_columns:
        raise ValueError(
            f"Akron Public {report_name} contains "
            f"unexpected columns: {unexpected_columns}. "
            f"Actual columns: {actual_columns}"
        )

    for col in optional_columns:
        if col not in df.columns:
            df[col] = None
            logger.info(
                "%s does not contain optional column "
                "%s; added blank column.",
                report_name,
                col,
            )

    df = df[expected_columns].copy()
    df = df.dropna(how="all").copy()
    return df


# ============================================================
# Convert Expense Excel
# ============================================================


def convert_expense_excel(content):
    df = read_account_detail_excel(
        content=content,
        expected_columns=EXPENSE_COLUMNS,
        report_name="Detailed Expense Report",
        required_columns=EXPENSE_REQUIRED_COLUMNS,
        optional_columns=EXPENSE_OPTIONAL_COLUMNS,
    )

    logger.info(
        "Expense rows before transaction filtering: %s",
        len(df),
    )

    df[EXPENSE_FILL_DOWN_COLUMNS] = df[EXPENSE_FILL_DOWN_COLUMNS].ffill()

    transaction_mask = (
        df["yr_per_jnl"].notna()
        & df["yr_per_jnl"].astype("string").str.strip().ne("")
    )

    df = df.loc[transaction_mask].copy()

    logger.info(
        "Expense transaction rows after filtering: %s",
        len(df),
    )

    for col in EXPENSE_MONEY_COLUMNS:
        df[col] = df[col].map(coerce_currency_decimal)

    df["percent_used"] = coerce_numeric_series(df["percent_used"])

    df["eff_date"] = pd.to_datetime(
        df["eff_date"],
        errors="coerce",
    )

    for col in EXPENSE_TEXT_COLUMNS:
        df[col] = clean_text_series(df[col])

    for col in [
        "fund",
        "spc_cost_ctr",
        "object",
        "po_number",
        "voucher",
        "invoice_number",
        "doc_number",
        "check_number",
    ]:
        df[col] = clean_identifier_series(df[col])

    df = df[EXPENSE_COLUMNS].copy()

    logger.info(
        "Final Akron expense rows: %s",
        len(df),
    )

    return df.reset_index(drop=True)


# ============================================================
# Convert Revenue Excel
# ============================================================


def convert_revenue_excel(content):
    df = read_account_detail_excel(
        content=content,
        expected_columns=REVENUE_COLUMNS,
        report_name="Detailed Revenue Report",
        required_columns=REVENUE_REQUIRED_COLUMNS,
        optional_columns=REVENUE_OPTIONAL_COLUMNS,
    )

    logger.info(
        "Revenue rows before transaction filtering: %s",
        len(df),
    )

    df[REVENUE_FILL_DOWN_COLUMNS] = df[REVENUE_FILL_DOWN_COLUMNS].ffill()

    transaction_mask = (
        df["yr_per_jnl"].notna()
        & df["yr_per_jnl"].astype("string").str.strip().ne("")
    )

    df = df.loc[transaction_mask].copy()

    logger.info(
        "Revenue transaction rows after filtering: %s",
        len(df),
    )

    for col in REVENUE_MONEY_COLUMNS:
        df[col] = df[col].map(coerce_currency_decimal)

    df["percent_coll"] = coerce_numeric_series(df["percent_coll"])

    df["eff_date"] = pd.to_datetime(
        df["eff_date"],
        errors="coerce",
    )

    for col in REVENUE_TEXT_COLUMNS:
        df[col] = clean_text_series(df[col])

    for col in [
        "fund",
        "spc_cost_ctr",
        "obj",
        "invoice_number",
        "check_number",
    ]:
        df[col] = clean_identifier_series(df[col])

    df = df[REVENUE_COLUMNS].copy()

    logger.info(
        "Final Akron revenue rows: %s",
        len(df),
    )

    return df.reset_index(drop=True)


# ============================================================
# Cash report CSV conversion and Parquet schema
# ============================================================


def convert_cash_csv(content):
    try:
        raw = pd.read_csv(BytesIO(content), dtype=str, encoding="utf-8-sig")
    except UnicodeDecodeError:
        raw = pd.read_csv(BytesIO(content), dtype=str, encoding="cp1252")

    raw.columns = [normalize_report_col_name(c) for c in raw.columns]

    missing = [c for c in CASH_COLUMNS if c not in raw.columns]
    if missing:
        raise ValueError(
            f"Akron cash CSV missing columns: {missing}; "
            f"actual: {list(raw.columns)}"
        )

    df = raw[CASH_COLUMNS].copy()
    df = df.dropna(how="all")

    for col in ["fund", "account", "description", "row_type", "period"]:
        df[col] = clean_text_series(df[col]).replace("", None)

    df["fund"] = df["fund"].str.replace(r"\.0$", "", regex=True)
    df["account"] = df["account"].str.replace(r"\.0$", "", regex=True)
    df["period"] = df["period"].str.upper()
    df["row_type"] = df["row_type"].str.lower()

    if not df["period"].isin(["MTD", "YTD"]).all():
        raise ValueError("Cash report has unexpected period values")

    if not df["row_type"].isin(["account", "fund_total", "grand_total"]).all():
        raise ValueError("Cash report has unexpected row_type values")

    df["page"] = pd.to_numeric(df["page"], errors="raise").astype("int32")

    for col in CASH_MONEY_COLUMNS:
        df[col] = df[col].map(coerce_currency_decimal)

    logger.info("Cash report rows: %s", len(df))
    return df.reset_index(drop=True)


def get_cash_schema():
    fields = []

    for col in CASH_COLUMNS + ["fiscal_year"]:
        if col == "page":
            typ = pa.int32()
        elif col in CASH_MONEY_COLUMNS:
            typ = pa.decimal128(20, 2)
        else:
            typ = pa.string()

        fields.append(pa.field(col, typ))

    return pa.schema(fields)


# ============================================================
# Expense Parquet schema
# ============================================================


def get_expense_schema():
    fields = []
    final_columns = EXPENSE_COLUMNS + ["fiscal_year"]

    for col in final_columns:
        if col in EXPENSE_MONEY_COLUMNS:
            fields.append(
                pa.field(
                    col,
                    pa.decimal128(20, 2),
                )
            )
        elif col == "eff_date":
            fields.append(
                pa.field(
                    col,
                    pa.timestamp("ns"),
                )
            )
        elif col == "percent_used":
            fields.append(
                pa.field(
                    col,
                    pa.float64(),
                )
            )
        else:
            fields.append(
                pa.field(
                    col,
                    pa.string(),
                )
            )

    return pa.schema(fields)


# ============================================================
# Revenue Parquet schema
# ============================================================


def get_revenue_schema():
    fields = []
    final_columns = REVENUE_COLUMNS + ["fiscal_year"]

    for col in final_columns:
        if col in REVENUE_MONEY_COLUMNS:
            fields.append(
                pa.field(
                    col,
                    pa.decimal128(20, 2),
                )
            )
        elif col == "eff_date":
            fields.append(
                pa.field(
                    col,
                    pa.timestamp("ns"),
                )
            )
        elif col == "percent_coll":
            fields.append(
                pa.field(
                    col,
                    pa.float64(),
                )
            )
        else:
            fields.append(
                pa.field(
                    col,
                    pa.string(),
                )
            )

    return pa.schema(fields)


# ============================================================
# Write Parquet
# ============================================================


def write_parquet(df, output, schema):
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


# ============================================================
# Process one S3 object
# ============================================================


def convert_s3_object(bucket, key, s3=None):
    if bucket != RAW_BUCKET:
        raise ValueError(
            f"Source bucket {bucket!r} does not match "
            f"RAW_BUCKET {RAW_BUCKET!r}"
        )

    validate_source_path(key)
    filename = PurePosixPath(key).name

    if filename.startswith("~$"):
        logger.info(
            "Skipping temporary Excel file: "
            "s3://%s/%s",
            bucket,
            key,
        )
        return None

    extension = PurePosixPath(filename).suffix.lower()
    if extension not in (".xlsx", ".csv"):
        logger.info("Skipping unsupported Akron file: s3://%s/%s", bucket, key)
        return None

    report_type = report_type_for_key(key)
    if report_type == "cash_report" and extension != ".csv":
        raise ValueError("Cash report must be the extracted CSV")
    if report_type != "cash_report" and extension != ".xlsx":
        raise ValueError("Expense/revenue reports must be XLSX")

    fiscal_year = fiscal_year_for_key(key)

    logger.info(
        "Detected Akron Public report: %s",
        report_type,
    )
    logger.info(
        "Detected fiscal year: %s",
        fiscal_year,
    )

    if s3 is None:
        s3 = boto3.client(
            "s3",
            region_name=AWS_REGION,
        )

    response = s3.get_object(
        Bucket=bucket,
        Key=key,
    )

    content = response["Body"].read()

    logger.info(
        "Downloaded source: s3://%s/%s",
        bucket,
        key,
    )

    if report_type == "expense_detail":
        logger.info(
            "Processing Akron Public Detailed Expense "
            "Report FY %s.",
            fiscal_year,
        )

        df = convert_expense_excel(content)
        df["fiscal_year"] = str(fiscal_year)
        final_columns = EXPENSE_COLUMNS + ["fiscal_year"]
        df = df[final_columns].copy()
        schema = get_expense_schema()
        dest_key = (
            f"{EXPENSE_DEST_PREFIX}/"
            f"{PurePosixPath(key).stem}.parquet"
        )

    elif report_type == "revenue_detail":
        logger.info(
            "Processing Akron Public Detailed Revenue "
            "Report FY %s.",
            fiscal_year,
        )

        df = convert_revenue_excel(content)
        df["fiscal_year"] = str(fiscal_year)
        final_columns = REVENUE_COLUMNS + ["fiscal_year"]
        df = df[final_columns].copy()
        schema = get_revenue_schema()
        dest_key = (
            f"{REVENUE_DEST_PREFIX}/"
            f"{PurePosixPath(key).stem}.parquet"
        )

    elif report_type == "cash_report":
        df = convert_cash_csv(content)
        df["fiscal_year"] = str(fiscal_year)
        df = df[CASH_COLUMNS + ["fiscal_year"]].copy()
        schema = get_cash_schema()
        dest_key = (
            f"{CASH_DEST_PREFIX}/"
            f"{PurePosixPath(key).stem}.parquet"
        )

    else:
        raise ValueError(
            "Unsupported Akron Public report type: "
            f"{report_type}"
        )

    output = BytesIO()

    write_parquet(
        df=df,
        output=output,
        schema=schema,
    )

    parquet_bytes = output.getvalue()

    logger.info(
        "Generated Parquet size: %s bytes",
        len(parquet_bytes),
    )

    s3.put_object(
        Bucket=DEST_BUCKET,
        Key=dest_key,
        Body=parquet_bytes,
    )

    logger.info(
        "Successfully processed Akron Public %s.",
        report_type,
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

    return {
        "source": f"s3://{bucket}/{key}",
        "destination": f"s3://{DEST_BUCKET}/{dest_key}",
        "rows": len(df),
        "fiscal_year": fiscal_year,
        "report_type": report_type,
    }


# ============================================================
# Lambda Handler
#
# Direct S3 trigger only
# ============================================================


def lambda_handler(event, context):
    logger.info(
        "Received event: %s",
        json.dumps(event),
    )

    results = []
    records = event.get("Records", [])

    if not records:
        logger.warning("Event contained no S3 records.")
        return {
            "statusCode": 200,
            "body": json.dumps([]),
        }

    for record in records:
        if record.get("eventSource") != "aws:s3":
            logger.warning("Skipping non-S3 event record.")
            continue

        if "s3" not in record:
            logger.warning("Skipping record without S3 information.")
            continue

        bucket = record["s3"]["bucket"]["name"]
        key = unquote_plus(record["s3"]["object"]["key"])

        logger.info(
            "S3 Object Created: s3://%s/%s",
            bucket,
            key,
        )

        result = convert_s3_object(
            bucket=bucket,
            key=key,
        )

        if result:
            results.append(result)

    return {
        "statusCode": 200,
        "body": json.dumps(results),
    }