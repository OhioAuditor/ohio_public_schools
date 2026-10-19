# Notification helpers

`apply_raw_bucket_notifications.py` merges the four Ohio school import S3 event rules into the existing raw bucket notification configuration.

Example:

```bash
python scripts/apply_raw_bucket_notifications.py --bucket schools-raw --region us-east-2 --dry-run
python scripts/apply_raw_bucket_notifications.py --bucket schools-raw --region us-east-2
```