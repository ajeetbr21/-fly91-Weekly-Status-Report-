"""
S3 collector module for bucket inventory and storage metrics.

Lists all S3 buckets, determines their region, and retrieves CloudWatch
storage metrics (BucketSizeBytes, NumberOfObjects).
"""

from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class S3Collector(BaseCollector):
    """Collects S3 bucket inventory and CloudWatch storage metrics.

    Important S3/CloudWatch quirks handled:
    - ``get_bucket_location`` returns ``None`` for us-east-1.
    - S3 CloudWatch metrics are daily (period must be 86400).
    - The ``StorageType`` dimension is **required**.
    - Looks back at least 3 days to ensure a datapoint is found.
    """

    _PERIOD_DAILY = 86400  # S3 metrics are only published daily

    def collect(self) -> List[Dict[str, Any]]:
        """Collect S3 bucket inventory and storage metrics.

        Returns:
            List of bucket detail dicts sorted by size descending.
        """
        self.logger.info("Collecting S3 bucket data…")

        s3 = self._get_client("s3")

        try:
            response = s3.list_buckets()
            buckets = response.get("Buckets", [])
        except (ClientError, BotoCoreError) as exc:
            self.logger.error("Failed to list S3 buckets: %s", exc)
            return []

        self.logger.info("Found %d S3 buckets; collecting metrics…", len(buckets))

        results: List[Dict[str, Any]] = []
        for bucket in buckets:
            bucket_name = bucket.get("Name", "")
            bucket_data = self._collect_bucket_data(s3, bucket_name)
            if bucket_data is not None:
                # Filter to only include us-east-1 buckets as requested
                if bucket_data.get("region") == "us-east-1":
                    results.append(bucket_data)

        # Sort by size descending (treat None as 0)
        results.sort(key=lambda b: b.get("total_size_bytes") or 0, reverse=True)

        self.logger.info("S3 collection complete — %d buckets in us-east-1.", len(results))
        return results

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _get_bucket_region(self, s3, bucket_name: str) -> str:
        """Determine the region of an S3 bucket.

        ``get_bucket_location`` returns ``None`` for us-east-1 buckets.
        """
        try:
            response = s3.get_bucket_location(Bucket=bucket_name)
            location = response.get("LocationConstraint")
            return location if location else "us-east-1"
        except (ClientError, BotoCoreError) as exc:
            self.logger.debug(
                "Could not determine region for bucket %s: %s", bucket_name, exc
            )
            return "unknown"

    def _collect_bucket_data(self, s3, bucket_name: str) -> Optional[Dict[str, Any]]:
        """Collect region and CloudWatch metrics for a single bucket."""
        region = self._get_bucket_region(s3, bucket_name)

        data: Dict[str, Any] = {
            "name": bucket_name,
            "region": region,
            "total_size_bytes": None,
            "size_gb": None,
            "total_objects": None,
        }

        # CloudWatch S3 metrics need a lookback of at least 3 days
        end_time = datetime.combine(
            self.end_date + timedelta(days=1), datetime.min.time()
        )
        start_time = end_time - timedelta(days=3)

        # --- BucketSizeBytes ---
        size_datapoints = self._get_cloudwatch_metric(
            namespace="AWS/S3",
            metric_name="BucketSizeBytes",
            dimensions=[
                {"Name": "BucketName", "Value": bucket_name},
                {"Name": "StorageType", "Value": "StandardStorage"},
            ],
            statistics=["Average"],
            period=self._PERIOD_DAILY,
            start_time=start_time,
            end_time=end_time,
        )
        if size_datapoints:
            # Take the most recent datapoint
            latest = size_datapoints[-1]
            size_bytes = latest.get("Average", 0)
            data["total_size_bytes"] = round(size_bytes, 2)
            data["size_gb"] = round(size_bytes / (1024 ** 3), 4)

        # --- NumberOfObjects ---
        obj_datapoints = self._get_cloudwatch_metric(
            namespace="AWS/S3",
            metric_name="NumberOfObjects",
            dimensions=[
                {"Name": "BucketName", "Value": bucket_name},
                {"Name": "StorageType", "Value": "AllStorageTypes"},
            ],
            statistics=["Average"],
            period=self._PERIOD_DAILY,
            start_time=start_time,
            end_time=end_time,
        )
        if obj_datapoints:
            latest = obj_datapoints[-1]
            data["total_objects"] = int(latest.get("Average", 0))

        return data
