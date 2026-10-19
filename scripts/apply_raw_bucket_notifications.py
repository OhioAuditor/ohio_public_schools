"""Apply the Ohio school import S3 notification rules to the raw bucket.

This script merges the four school-import Lambda notifications into the
existing bucket notification configuration instead of replacing unrelated
entries.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from dataclasses import dataclass
from typing import Any, Dict, List, Sequence

import boto3


DEFAULT_BUCKET = "schools-raw"
DEFAULT_REGION = os.getenv("AWS_REGION", "us-east-2")
DEFAULT_ACCOUNT_ID = os.getenv("AWS_ACCOUNT_ID", "000000000000")
DEFAULT_NOTIFICATION_FILE = Path(__file__).resolve().parents[1] / "lambda-backup-20261015-105154" / "schools-raw-notifications.json"


@dataclass(frozen=True)
class NotificationRule:
    function_name: str
    prefix: str
    suffix: str | None = None


RULES: Sequence[NotificationRule] = (
    NotificationRule("akron_public_import", "data-from-clients/akron_public/"),
    NotificationRule("cincinnati_public_import", "data-from-clients/cincinnati_public/"),
    NotificationRule("dayton_public_schools_import", "data-from-clients/dayton_public/", ".xlsx"),
    NotificationRule("toledo_public_schools_import", "data-from-clients/toledo_public/", ".xlsx"),
)


def load_bucket_config(s3_client: Any, bucket: str) -> Dict[str, Any]:
    response = s3_client.get_bucket_notification_configuration(Bucket=bucket)
    return {
        key: value
        for key, value in response.items()
        if key != "ResponseMetadata" and value
    }


def load_local_config(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def account_id_from_config(config: Dict[str, Any]) -> str:
    for entry in config.get("LambdaFunctionConfigurations", []):
        arn = entry.get("LambdaFunctionArn", "")
        parts = arn.split(":")
        if len(parts) >= 5 and parts[0] == "arn" and parts[1] == "aws" and parts[2] == "lambda":
            return parts[4]
    return DEFAULT_ACCOUNT_ID


def build_function_arn(region: str, account_id: str, function_name: str) -> str:
    return f"arn:aws:lambda:{region}:{account_id}:function:{function_name}"


def ensure_lambda_permission(lambda_client: Any, function_name: str, bucket_arn: str, account_id: str) -> str:
    function = lambda_client.get_function(FunctionName=function_name)
    function_arn = function["Configuration"]["FunctionArn"]

    try:
        lambda_client.add_permission(
            FunctionName=function_name,
            StatementId=f"s3-{function_name}",
            Action="lambda:InvokeFunction",
            Principal="s3.amazonaws.com",
            SourceArn=bucket_arn,
            SourceAccount=account_id,
        )
    except lambda_client.exceptions.ResourceConflictException:
        pass

    return function_arn


def build_lambda_config(function_arn: str, prefix: str, suffix: str | None) -> Dict[str, Any]:
    rules = [{"Name": "prefix", "Value": prefix}]
    if suffix:
        rules.append({"Name": "suffix", "Value": suffix})

    return {
        "Id": f"ohio-public-schools-{prefix.strip('/').replace('/', '-')}",
        "LambdaFunctionArn": function_arn,
        "Events": ["s3:ObjectCreated:*"],
        "Filter": {
            "Key": {
                "FilterRules": rules,
            }
        },
    }


def merge_notification_config(existing: Dict[str, Any], updates: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    existing_configs = existing.get("LambdaFunctionConfigurations", [])

    update_keys = {
        (entry["Filter"]["Key"]["FilterRules"][0]["Value"], entry["LambdaFunctionArn"])
        for entry in updates
    }

    preserved = [
        entry
        for entry in existing_configs
        if (
            entry.get("Filter", {})
            .get("Key", {})
            .get("FilterRules", [{}])[0]
            .get("Value"),
            entry.get("LambdaFunctionArn"),
        )
        not in update_keys
    ]

    merged = dict(existing)
    merged["LambdaFunctionConfigurations"] = [*preserved, *updates]
    return merged


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bucket", default=DEFAULT_BUCKET)
    parser.add_argument("--region", default=DEFAULT_REGION)
    parser.add_argument("--account-id", default=DEFAULT_ACCOUNT_ID)
    parser.add_argument("--notifications-file", default=str(DEFAULT_NOTIFICATION_FILE))
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    bucket_arn = f"arn:aws:s3:::{args.bucket}"

    if args.dry_run:
        existing = load_local_config(Path(args.notifications_file))
        account_id = account_id_from_config(existing)
    else:
        session = boto3.Session(region_name=args.region)
        s3_client = session.client("s3")
        lambda_client = session.client("lambda")
        sts_client = session.client("sts")

        account_id = sts_client.get_caller_identity()["Account"]
        existing = load_bucket_config(s3_client, args.bucket)

    updates: List[Dict[str, Any]] = []
    for rule in RULES:
        if args.dry_run:
            function_arn = build_function_arn(args.region, account_id, rule.function_name)
        else:
            function_arn = ensure_lambda_permission(
                lambda_client=lambda_client,
                function_name=rule.function_name,
                bucket_arn=bucket_arn,
                account_id=account_id,
            )
        updates.append(build_lambda_config(function_arn, rule.prefix, rule.suffix))

    merged = merge_notification_config(existing, updates)

    print(json.dumps(merged, indent=2, sort_keys=True))

    if not args.dry_run:
        s3_client.put_bucket_notification_configuration(
            Bucket=args.bucket,
            NotificationConfiguration=merged,
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())