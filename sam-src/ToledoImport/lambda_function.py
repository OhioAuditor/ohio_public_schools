import json
import logging
import os
import re
from io import BytesIO
from pathlib import PurePosixPath
from urllib.parse import unquote_plus

import pandas as pd


# ============================================================
# Logging
# ============================================================

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


# ============================================================
# Configuration
# ============================================================

AWS_REGION = os.getenv("AWS_REGION", "us-east-2")

RAW_BUCKET = os.getenv(
    "RAW_BUCKET",
    "schools-raw",
).strip()

DEST_BUCKET = os.getenv(
    "DEST_BUCKET",
    "schools-transformed",
).strip()


# ============================================================
# Toledo Public Destination Prefixes
# ============================================================
#
# Source:
#   s3://schools-raw/data-from-clients/toledo_public/
#
# Destinations:
#   s3://schools-transformed/data_from_clients/
#       toledo_public_expense_detail/
#
#   s3://schools-transformed/data_from_clients/
#       toledo_public_receipt_detail/
#
# Environment variables can override these defaults.
# ============================================================

EXPENSE_DEST_PREFIX = os.getenv(
    "TOLEDO_PUBLIC_EXPENSE_DETAIL_DEST_PREFIX",
    "data_from_clients/toledo_public_expense_detail/",
).strip("/")

RECEIPT_DEST_PREFIX = os.getenv(
    "TOLEDO_PUBLIC_RECEIPT_DETAIL_DEST_PREFIX",
    "data_from_clients/toledo_public_receipt_detail/",
).strip("/")


# ============================================================
# Report Configuration
# ============================================================

REPORTS = {
    "expense_detail": (
        "expense detail report",
        EXPENSE_DEST_PREFIX,
    ),
    "receipt_detail": (
        "receipt detail report",
        RECEIPT_DEST_PREFIX,
    ),
}


# ============================================================
# Expected Toledo Columns
# ============================================================

TRANSACTION_COLUMNS = [
    "obj_type",
    "fund_desc",
    "cc_desc",
    "fqa_acct",
    "fqa_desc",
    "ref_1",
    "ref_2",
    "ref_2_dt",
    "vendor",
    "trans_date",
    "encumb_amt",
    "check_no",
    "check_dt",
    "trans_amt",
    "state_line",
    "fiscal_yr",
]


# ============================================================
# Column Name Normalization
# ============================================================

def normalize_report_col_name(name):
    """
    Normalize Excel column names to lowercase snake_case.
    """

    s = str(name).strip().lower().replace("#", "number")

    return re.sub(
        r"_+",
        "_",
        re.sub(
            r"[^a-z0-9]+",
            "_",
            s,
        ),
    ).strip("_")


# ============================================================
# Numeric Conversion
# ============================================================

def coerce_numeric_series(series):
    """
    Convert currency/numeric values to float64.

    Handles:
        $1,234.56
        1,234.56
        (1,234.56)
        25%
        blanks
    """

    cleaned = (
        series
        .astype("string")
        .str.strip()
    )

    # Convert parentheses negatives:
    # (123.45) -> -123.45
    cleaned = cleaned.str.replace(
        r"^\((.*)\)$",
        r"-\1",
        regex=True,
    )

    # Remove currency symbols, commas,
    # percent signs, and whitespace.
    cleaned = cleaned.str.replace(
        r"[$,%\s]",
        "",
        regex=True,
    )

    return pd.to_numeric(
        cleaned,
        errors="coerce",
    ).astype("float64")


# ============================================================
# Determine Report Type
# ============================================================

def report_type_for_key(key):
    """
    Determine whether the Toledo file is an expense
    or revenue/receipt report based on the filename.
    """

    name = PurePosixPath(key).name.lower()

    name = (
        name
        .replace("_", " ")
        .replace("-", " ")
    )

    name = re.sub(
        r"\s+",
        " ",
        name,
    ).strip()

    # Examples:
    # Transaction Details Report - FY26 Expenses Fund 001.xlsx
    if re.search(r"\bexpenses?\b", name):
        return "expense_detail"

    # Supports revenue, revenues, receipt, receipts.
    if re.search(
        r"\b(revenues?|receipts?)\b",
        name,
    ):
        return "receipt_detail"

    raise ValueError(
        "Unrecognized Toledo Public report filename: "
        f"{key}"
    )


# ============================================================
# Excel Conversion
# ============================================================

def convert_excel(content, report_type):
    """
    Read a Toledo Public Excel report and clean the
    data before writing it to Parquet.
    """

    df = pd.read_excel(
        BytesIO(content),
        sheet_name=0,
        dtype=object,
        engine="openpyxl",
    )

    # --------------------------------------------------------
    # Normalize column names
    # --------------------------------------------------------

    df.columns = [
        normalize_report_col_name(col)
        for col in df.columns
    ]

    actual_columns = list(df.columns)

    # --------------------------------------------------------
    # Validate report type
    # --------------------------------------------------------

    if report_type not in REPORTS:
        raise ValueError(
            f"Unsupported report type: {report_type}"
        )

    # --------------------------------------------------------
    # Validate columns
    # --------------------------------------------------------

    if (
        len(actual_columns) != len(TRANSACTION_COLUMNS)
        or set(actual_columns) != set(TRANSACTION_COLUMNS)
    ):
        raise ValueError(
            f"Unexpected Toledo columns: "
            f"{actual_columns}; "
            f"expected {TRANSACTION_COLUMNS}"
        )

    # Put columns into expected order.
    df = (
        df.loc[:, TRANSACTION_COLUMNS]
        .dropna(how="all")
        .copy()
    )

    # --------------------------------------------------------
    # Numeric columns
    # --------------------------------------------------------

    amount_columns = [
        "encumb_amt",
        "trans_amt",
    ]

    for col in amount_columns:
        df[col] = (
            coerce_numeric_series(df[col])
            .round(2)
        )

    # --------------------------------------------------------
    # Date columns
    # --------------------------------------------------------

    date_columns = [
        "ref_2_dt",
        "trans_date",
        "check_dt",
    ]

    for col in date_columns:

        # Convert to strings first so invalid placeholders
        # can safely become null dates.
        date_values = (
            df[col]
            .astype("string")
            .str.strip()
        )

        # Parse each unique value once.
        parsed_dates = {
            value: pd.to_datetime(
                value,
                errors="coerce",
            )
            for value in date_values.dropna().unique()
        }

        df[col] = date_values.map(
            parsed_dates
        )

        df[col] = pd.to_datetime(
            df[col],
            errors="coerce",
        )

    # --------------------------------------------------------
    # Remaining columns -> strings
    # --------------------------------------------------------

    for col in df.columns:

        if (
            col not in amount_columns
            and col not in date_columns
        ):
            df[col] = df[col].map(
                lambda value: (
                    None
                    if pd.isna(value)
                    else str(value).strip()
                )
            ).astype("string")

    return df


# ============================================================
# Write Parquet
# ============================================================

def write_parquet(df, output):
    """
    Write DataFrame to Snappy-compressed Parquet.
    """

    df.to_parquet(
        output,
        engine="pyarrow",
        index=False,
        compression="snappy",
    )


# ============================================================
# Convert S3 Object
# ============================================================

def convert_s3_object(bucket, key, s3=None):
    """
    Download a Toledo Excel report from schools-raw,
    convert it to Parquet, and upload it to
    schools-transformed.
    """

    logger.info(
        "Processing source: s3://%s/%s",
        bucket,
        key,
    )

    # --------------------------------------------------------
    # Validate source bucket
    # --------------------------------------------------------

    if bucket != RAW_BUCKET:
        raise ValueError(
            f"Source bucket {bucket!r} does not match "
            f"RAW_BUCKET {RAW_BUCKET!r}"
        )

    filename = PurePosixPath(key).name

    # --------------------------------------------------------
    # Ignore non-XLSX files and temporary Excel files
    # --------------------------------------------------------

    if (
        not filename.lower().endswith(".xlsx")
        or filename.startswith("~$")
    ):
        logger.info(
            "Skipping non-report file: s3://%s/%s",
            bucket,
            key,
        )

        return None

    # --------------------------------------------------------
    # Determine report type
    # --------------------------------------------------------

    report_type = report_type_for_key(key)

    logger.info(
        "Detected report type: %s",
        report_type,
    )

    # --------------------------------------------------------
    # Create S3 client
    # --------------------------------------------------------

    if s3 is None:
        import boto3

        s3 = boto3.client(
            "s3",
            region_name=AWS_REGION,
        )

    # --------------------------------------------------------
    # Download Excel file
    # --------------------------------------------------------

    logger.info(
        "Downloading: s3://%s/%s",
        bucket,
        key,
    )

    content = s3.get_object(
        Bucket=bucket,
        Key=key,
    )["Body"].read()

    # --------------------------------------------------------
    # Convert Excel
    # --------------------------------------------------------

    df = convert_excel(
        content,
        report_type,
    )

    logger.info(
        "Converted %s rows from Excel.",
        len(df),
    )

    # --------------------------------------------------------
    # Create Parquet
    # --------------------------------------------------------

    output = BytesIO()

    write_parquet(
        df,
        output,
    )

    # --------------------------------------------------------
    # Build destination key
    # --------------------------------------------------------

    stem = PurePosixPath(key).stem

    dest_prefix = REPORTS[
        report_type
    ][1]

    dest_key = (
        f"{dest_prefix}/"
        f"{stem}.parquet"
    )

    # --------------------------------------------------------
    # Log destination BEFORE PutObject
    # --------------------------------------------------------

    logger.info(
        "Destination prefix: %s",
        dest_prefix,
    )

    logger.info(
        "Uploading to: s3://%s/%s",
        DEST_BUCKET,
        dest_key,
    )

    # --------------------------------------------------------
    # Upload Parquet
    # --------------------------------------------------------

    s3.put_object(
        Bucket=DEST_BUCKET,
        Key=dest_key,
        Body=output.getvalue(),
    )

    logger.info(
        "Successfully converted %s rows: "
        "s3://%s/%s",
        len(df),
        DEST_BUCKET,
        dest_key,
    )

    return {
        "source": (
            f"s3://{bucket}/{key}"
        ),
        "destination": (
            f"s3://{DEST_BUCKET}/{dest_key}"
        ),
        "report_type": report_type,
        "rows": len(df),
    }


# ============================================================
# Parse Lambda Events
# ============================================================

def iter_s3_records(event):
    """
    Support:
        - Direct Lambda test events
        - S3 events
        - SQS-wrapped S3 events
        - SNS-wrapped S3 events
    """

    if isinstance(event, str):
        event = json.loads(event)

    # --------------------------------------------------------
    # Direct test invocation
    # --------------------------------------------------------

    if (
        "Records" not in event
        and "bucket" in event
        and "key" in event
    ):
        yield (
            event["bucket"],
            event["key"],
        )

    # --------------------------------------------------------
    # AWS event records
    # --------------------------------------------------------

    for record in event.get(
        "Records",
        [],
    ):

        # Direct S3 event
        if "s3" in record:

            yield (
                record["s3"]["bucket"]["name"],
                unquote_plus(
                    record["s3"]["object"]["key"]
                ),
            )

        # SQS event
        elif "body" in record:

            yield from iter_s3_records(
                record["body"]
            )

        # SNS event
        elif "Sns" in record:

            yield from iter_s3_records(
                record["Sns"]["Message"]
            )


# ============================================================
# Lambda Handler
# ============================================================

def lambda_handler(event, context):

    # --------------------------------------------------------
    # Log runtime configuration
    # --------------------------------------------------------

    logger.info(
        "RAW_BUCKET=%s",
        RAW_BUCKET,
    )

    logger.info(
        "DEST_BUCKET=%s",
        DEST_BUCKET,
    )

    logger.info(
        "EXPENSE_DEST_PREFIX=%s",
        EXPENSE_DEST_PREFIX,
    )

    logger.info(
        "RECEIPT_DEST_PREFIX=%s",
        RECEIPT_DEST_PREFIX,
    )

    logger.info(
        "Received event: %s",
        json.dumps(event),
    )

    # --------------------------------------------------------
    # Process files
    # --------------------------------------------------------

    results = []

    for bucket, key in iter_s3_records(event):

        result = convert_s3_object(
            bucket,
            key,
        )

        if result:
            results.append(result)

    # --------------------------------------------------------
    # Return results
    # --------------------------------------------------------

    return {
        "statusCode": 200,
        "body": json.dumps(
            results,
            default=str,
        ),
    }