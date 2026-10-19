import json
import logging
import os
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from pathlib import PurePosixPath
from urllib.parse import unquote_plus

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

AWS_REGION = os.getenv("AWS_REGION", "us-east-2")
RAW_BUCKET = os.getenv("RAW_BUCKET", "schools-raw").strip()
DEST_BUCKET = os.getenv("DEST_BUCKET", "schools-transformed").strip()

CASH_DEST_PREFIX = os.getenv(
    "DAYTON_PUBLIC_CASH_POSITION_DEST_PREFIX",
    "data_from_clients/dayton_public_cash_position/",
).strip("/")

EXPENSE_DEST_PREFIX = os.getenv(
    "DAYTON_PUBLIC_EXPENSE_DETAIL_DEST_PREFIX",
    "data_from_clients/dayton_public_expense_detail/",
).strip("/")

RECEIPT_DEST_PREFIX = os.getenv(
    "DAYTON_PUBLIC_RECEIPT_DETAIL_DEST_PREFIX",
    "data_from_clients/dayton_public_receipt_detail/",
).strip("/")

REPORTS = {
    "cash_position": ("cash position report", CASH_DEST_PREFIX),
    "expense_detail": ("expense detail report", EXPENSE_DEST_PREFIX),
    "receipt_detail": ("receipt detail report", RECEIPT_DEST_PREFIX),
}

DETAIL_COLUMNS = [
    "fund",
    "special_cost_center",
    "function",
    "opu",
    "account",
    "gl_reference",
    "description",
    "posting_date",
    "ss",
    "ref2",
    "peid",
    "ck_id",
    "ck_no",
    "debit",
    "credit",
    "balance",
]

# Some Receipt Detail files contain an additional "exp/rev" column.
RECEIPT_COLUMNS_WITH_EXP_REV = DETAIL_COLUMNS.copy()
RECEIPT_COLUMNS_WITH_EXP_REV.insert(
    RECEIPT_COLUMNS_WITH_EXP_REV.index("gl_reference"),
    "exp_rev",
)

CASH_COLUMNS = [
    "fund",
    "scc",
    "opu",
    "desc",
    "beginning_balance",
    "receipts_ytd",
    "receipts_mtd",
    "expenditures_ytd",
    "expenditures_mtd",
    "current_balance",
    "encumbrances",
    "unencumbered_cash",
    "as_of",
]

CASH_AMOUNTS = CASH_COLUMNS[4:12]


def normalize_report_col_name(name):
    s = str(name).strip().lower().replace("#", "number")
    return re.sub(
        r"_+",
        "_",
        re.sub(r"[^a-z0-9]+", "_", s),
    ).strip("_")


def coerce_numeric_series(series):
    """Existing numeric handling for Expense and Receipt reports."""
    cleaned = series.astype("string").str.strip()
    cleaned = cleaned.str.replace(
        r"^\((.*)\)$",
        r"-\1",
        regex=True,
    )
    cleaned = cleaned.str.replace(
        r"[$,%\s]",
        "",
        regex=True,
    )
    return pd.to_numeric(
        cleaned,
        errors="coerce",
    ).astype("float64")


def coerce_currency_decimal(value):
    """Convert an Excel Cash Position amount to exact cents."""
    if pd.isna(value):
        return None

    cleaned = str(value).strip()
    if not cleaned:
        return None

    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = "-" + cleaned[1:-1]

    cleaned = re.sub(r"[$,%\s]", "", cleaned)

    try:
        return Decimal(cleaned).quantize(
            Decimal("0.01"),
            rounding=ROUND_HALF_UP,
        )
    except InvalidOperation as exc:
        raise ValueError(
            f"Invalid Cash Position amount: {value!r}"
        ) from exc


def report_type_for_key(key):
    name = PurePosixPath(key).name.lower()
    name = name.replace("_", " ").replace("-", " ")
    name = re.sub(r"\s+", " ", name)

    for report_type, (phrase, _) in REPORTS.items():
        if phrase in name:
            return report_type

    raise ValueError(
        f"Unrecognized Dayton Public report filename: {key}"
    )


def fiscal_year_for_key(key):
    filename = PurePosixPath(key).name

    match = re.match(
        r"^[A-Za-z]+ \d{1,2} (20\d{2}) ",
        filename,
        flags=re.IGNORECASE,
    )

    if not match:
        raise ValueError(
            f"Could not find fiscal year in report filename: {filename}"
        )

    return match.group(1)


def convert_excel(content, report_type):
    df = pd.read_excel(
        BytesIO(content),
        sheet_name=0,
        dtype=object,
        engine="openpyxl",
    )

    df.columns = [
        normalize_report_col_name(col)
        for col in df.columns
    ]

    actual_columns = list(df.columns)

    if report_type == "cash_position":
        valid_columns = [CASH_COLUMNS]
    elif report_type == "receipt_detail":
        valid_columns = [
            DETAIL_COLUMNS,
            RECEIPT_COLUMNS_WITH_EXP_REV,
        ]
    else:
        valid_columns = [DETAIL_COLUMNS]

    if actual_columns not in valid_columns:
        raise ValueError(
            f"Unexpected {report_type} columns: "
            f"{actual_columns}; expected one of {valid_columns}"
        )

    df = df.dropna(how="all").copy()

    if report_type == "cash_position":
        for col in CASH_AMOUNTS:
            df[col] = df[col].map(coerce_currency_decimal)

        amounts = CASH_AMOUNTS
    else:
        amounts = ["debit", "credit", "balance"]

        for col in amounts:
            df[col] = coerce_numeric_series(df[col]).round(2)

        df["posting_date"] = pd.to_datetime(
            df["posting_date"],
            errors="coerce",
        )

    text_cols = [
        col
        for col in df.columns
        if col not in amounts and col != "posting_date"
    ]

    for col in text_cols:
        df[col] = df[col].map(
            lambda value: (
                None if pd.isna(value) else str(value).strip()
            )
        ).astype("string")

    return df


def write_parquet(df, report_type, output):
    if report_type == "cash_position":
        schema = pa.schema([
            pa.field(
                col,
                pa.decimal128(20, 2)
                if col in CASH_AMOUNTS
                else pa.string(),
            )
            for col in CASH_COLUMNS + ["fiscal_year"]
        ])

        table = pa.Table.from_pandas(
            df,
            schema=schema,
            preserve_index=False,
        )

        pq.write_table(
            table,
            output,
            compression="snappy",
        )
    else:
        df.to_parquet(
            output,
            engine="pyarrow",
            index=False,
            compression="snappy",
        )


def convert_s3_object(bucket, key, s3=None):
    if bucket != RAW_BUCKET:
        raise ValueError(
            f"Source bucket {bucket!r} does not match "
            f"RAW_BUCKET {RAW_BUCKET!r}"
        )

    filename = PurePosixPath(key).name

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

    report_type = report_type_for_key(key)
    fiscal_year = fiscal_year_for_key(key)

    if s3 is None:
        import boto3

        s3 = boto3.client("s3", region_name=AWS_REGION)

    content = s3.get_object(
        Bucket=bucket,
        Key=key,
    )["Body"].read()

    df = convert_excel(content, report_type)
    df["fiscal_year"] = fiscal_year

    output = BytesIO()
    write_parquet(df, report_type, output)

    stem = PurePosixPath(key).stem
    dest_prefix = REPORTS[report_type][1]
    dest_key = f"{dest_prefix}/{stem}.parquet"

    s3.put_object(
        Bucket=DEST_BUCKET,
        Key=dest_key,
        Body=output.getvalue(),
    )

    logger.info(
        "Converted %s rows for fiscal year %s: s3://%s/%s",
        len(df),
        fiscal_year,
        DEST_BUCKET,
        dest_key,
    )

    return {
        "source": f"s3://{bucket}/{key}",
        "destination": f"s3://{DEST_BUCKET}/{dest_key}",
        "rows": len(df),
        "fiscal_year": fiscal_year,
    }


def iter_s3_records(event):
    if isinstance(event, str):
        event = json.loads(event)

    if "Records" not in event and "bucket" in event and "key" in event:
        yield event["bucket"], event["key"]

    for record in event.get("Records", []):
        if "s3" in record:
            yield (
                record["s3"]["bucket"]["name"],
                unquote_plus(record["s3"]["object"]["key"]),
            )
        elif "body" in record:
            yield from iter_s3_records(record["body"])
        elif "Sns" in record:
            yield from iter_s3_records(record["Sns"]["Message"])


def lambda_handler(event, context):
    results = []

    for bucket, key in iter_s3_records(event):
        result = convert_s3_object(bucket, key)
        if result:
            results.append(result)

    return {
        "statusCode": 200,
        "body": json.dumps(results),
    }