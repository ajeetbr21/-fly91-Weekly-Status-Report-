"""
Base collector module providing abstract base class for all AWS data collectors.

All concrete collectors inherit from BaseCollector and implement the collect() method.
Provides shared utilities for creating boto3 clients with retry logic and
querying CloudWatch metrics.
"""

from abc import ABC, abstractmethod
from datetime import datetime, date, timedelta
from typing import Any, Dict, List, Optional, Tuple

import boto3
from botocore.config import Config
from botocore.exceptions import ClientError, BotoCoreError


# Default retry configuration for all boto3 clients
DEFAULT_RETRY_CONFIG = Config(
    retries={
        "max_attempts": 3,
        "mode": "adaptive",
    },
    connect_timeout=10,
    read_timeout=30,
)


class BaseCollector(ABC):
    """Abstract base class for all AWS resource collectors.

    Provides shared functionality including:
    - boto3 client creation with retry configuration
    - CloudWatch metric querying helpers
    - Consistent error handling patterns

    Attributes:
        session: An authenticated boto3.Session.
        region: The AWS region to collect data from.
        config: Application configuration dictionary.
        start_date: Start date of the reporting period.
        end_date: End date of the reporting period.
        logger: Logger instance for structured logging.
    """

    def __init__(
        self,
        session: boto3.Session,
        region: str,
        config: dict,
        start_date: date,
        end_date: date,
        logger: Any,
    ) -> None:
        self.session = session
        self.region = region
        self.config = config
        self.start_date = start_date
        self.end_date = end_date
        self.logger = logger

    @abstractmethod
    def collect(self) -> dict:
        """Collect data from the AWS service.

        Must be implemented by all concrete collectors.

        Returns:
            A dictionary containing the collected data, structure varies
            by collector type.
        """

    # ------------------------------------------------------------------
    # Client helpers
    # ------------------------------------------------------------------

    def _get_client(self, service_name: str, region_override: Optional[str] = None):
        """Create a boto3 client for the given service with retry configuration.

        Args:
            service_name: AWS service name (e.g. 'ec2', 'ce', 's3').
            region_override: Optional region override; defaults to self.region.

        Returns:
            A boto3 client for the requested service.
        """
        target_region = region_override or self.region
        try:
            client = self.session.client(
                service_name,
                region_name=target_region,
                config=DEFAULT_RETRY_CONFIG,
            )
            return client
        except (ClientError, BotoCoreError) as exc:
            self.logger.error(
                "Failed to create %s client in %s: %s",
                service_name,
                target_region,
                exc,
            )
            raise

    # ------------------------------------------------------------------
    # CloudWatch helpers
    # ------------------------------------------------------------------

    def _get_cloudwatch_metric(
        self,
        namespace: str,
        metric_name: str,
        dimensions: List[Dict[str, str]],
        statistics: List[str],
        period: Optional[int] = None,
        unit: Optional[str] = None,
        start_time: Optional[datetime] = None,
        end_time: Optional[datetime] = None,
    ) -> List[dict]:
        """Retrieve CloudWatch metric statistics for a given metric.

        Uses self.start_date and self.end_date to define the time range
        unless explicit start_time/end_time are provided.

        Args:
            namespace: CloudWatch metric namespace (e.g. 'AWS/EC2').
            metric_name: Name of the metric.
            dimensions: List of dimension dicts with 'Name' and 'Value'.
            statistics: List of statistics to retrieve (e.g. ['Average']).
            period: Aggregation period in seconds.  Defaults to the total
                    seconds in the reporting date range.
            unit: Optional CloudWatch unit filter.
            start_time: Override for the query start time.
            end_time: Override for the query end time.

        Returns:
            Sorted list of CloudWatch datapoints (oldest first).
        """
        cw = self._get_client("cloudwatch")

        if start_time is None:
            from datetime import timezone
            start_time = datetime.combine(self.start_date, datetime.min.time(), tzinfo=timezone.utc)
        if end_time is None:
            from datetime import timezone
            # end_date is inclusive, so go to start of the next day
            end_time = datetime.combine(
                self.end_date + timedelta(days=1), datetime.min.time(), tzinfo=timezone.utc
            )

        if period is None:
            total_seconds = int((end_time - start_time).total_seconds())
            period = max(total_seconds, 60)  # CloudWatch minimum is 60s

        try:
            total_seconds = (end_time - start_time).total_seconds()
            max_chunk_seconds = 1440 * period
            
            all_datapoints = []
            current_start = start_time
            
            while current_start < end_time:
                current_end = current_start + timedelta(seconds=max_chunk_seconds)
                if current_end > end_time:
                    current_end = end_time
                    
                params: Dict[str, Any] = {
                    "Namespace": namespace,
                    "MetricName": metric_name,
                    "Dimensions": dimensions,
                    "StartTime": current_start,
                    "EndTime": current_end,
                    "Period": period,
                    "Statistics": statistics,
                }
                if unit is not None:
                    params["Unit"] = unit
                    
                response = cw.get_metric_statistics(**params)
                all_datapoints.extend(response.get("Datapoints", []))
                
                current_start = current_end
                
            # Sort by timestamp ascending
            all_datapoints.sort(key=lambda dp: dp["Timestamp"])
            return all_datapoints
        except ClientError as exc:
            self.logger.warning(
                "CloudWatch query failed for %s/%s: %s",
                namespace,
                metric_name,
                exc,
            )
            return []
        except BotoCoreError as exc:
            self.logger.warning(
                "CloudWatch query error for %s/%s: %s",
                namespace,
                metric_name,
                exc,
            )
            return []

    # ------------------------------------------------------------------
    # Utility helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _safe_float(value: Any, default: float = 0.0) -> float:
        """Safely convert a value to float, returning default on failure."""
        try:
            return float(value)
        except (TypeError, ValueError):
            return default

    def _calculate_period_seconds(self) -> int:
        """Return the total seconds spanning start_date to end_date (inclusive)."""
        start_dt = datetime.combine(self.start_date, datetime.min.time())
        end_dt = datetime.combine(
            self.end_date + timedelta(days=1), datetime.min.time()
        )
        return max(int((end_dt - start_dt).total_seconds()), 60)
