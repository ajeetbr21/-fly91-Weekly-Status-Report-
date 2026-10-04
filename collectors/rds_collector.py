"""
RDS collector module for database instance inventory and performance metrics.

Retrieves all RDS instances and their CloudWatch metrics including CPU
utilisation, free storage space, and database connection counts.
"""

from typing import Any, Dict, List

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class RDSCollector(BaseCollector):
    """Collects RDS instance data and CloudWatch metrics.

    For each DB instance:
    - Configuration: identifier, engine, instance class
    - CPU utilisation (average) from AWS/RDS
    - Free storage space (average, converted to GB)
    - Database connections (average)

    Returns an empty list if no RDS instances are found.
    """

    def collect(self) -> List[Dict[str, Any]]:
        """Collect RDS instance inventory and metrics.

        Returns:
            List of RDS instance detail dicts, or an empty list.
        """
        self.logger.info("Collecting RDS data for region %s…", self.region)

        rds = self._get_client("rds")
        instances: List[dict] = []

        try:
            paginator = rds.get_paginator("describe_db_instances")
            for page in paginator.paginate():
                instances.extend(page.get("DBInstances", []))
        except (ClientError, BotoCoreError) as exc:
            self.logger.error("Failed to describe RDS instances: %s", exc)
            return []

        if not instances:
            self.logger.info("No RDS instances found in %s.", self.region)
            return []

        self.logger.info(
            "Found %d RDS instances; collecting metrics…", len(instances)
        )

        results: List[Dict[str, Any]] = []
        for db in instances:
            db_data = self._collect_db_metrics(db)
            results.append(db_data)

        self.logger.info("RDS collection complete — %d instances.", len(results))
        return results

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _collect_db_metrics(self, db: dict) -> Dict[str, Any]:
        """Collect CloudWatch metrics for a single RDS instance."""
        identifier = db.get("DBInstanceIdentifier", "")
        dims = [{"Name": "DBInstanceIdentifier", "Value": identifier}]

        data: Dict[str, Any] = {
            "db_identifier": identifier,
            "engine": db.get("Engine", ""),
            "engine_version": db.get("EngineVersion", ""),
            "instance_class": db.get("DBInstanceClass", ""),
            "status": db.get("DBInstanceStatus", "unknown"),
            "multi_az": db.get("MultiAZ", False),
            "storage_type": db.get("StorageType", ""),
            "allocated_storage_gb": db.get("AllocatedStorage", 0),
        }

        # Lookup total memory (GB) based on instance class
        mem_map = {
            "db.t3.medium": 4,
            "db.t3.large": 8,
            "db.r7g.large": 16,
            "db.r5.large": 16,
            "db.m5.large": 8,
        }
        data["total_memory_gb"] = mem_map.get(data["instance_class"])

        # --- CPU Utilisation (Min, Max, Avg of 5-min Maxes) ---
        cpu_dp = self._get_cloudwatch_metric(
            namespace="AWS/RDS",
            metric_name="CPUUtilization",
            dimensions=dims,
            statistics=["Maximum"],
            period=300,
        )
        if cpu_dp:
            max_vals = [dp.get("Maximum", 0.0) for dp in cpu_dp if dp.get("Maximum") is not None]
            if max_vals:
                data["cpu_min"] = round(min(max_vals), 2)
                data["cpu_max"] = round(max(max_vals), 2)
                data["cpu_avg"] = round(sum(max_vals) / len(max_vals), 2)
        else:
            data["cpu_min"], data["cpu_max"], data["cpu_avg"] = None, None, None

        # --- Freeable Memory (Max of 5-min Maxes) ---
        mem_dp = self._get_cloudwatch_metric(
            namespace="AWS/RDS",
            metric_name="FreeableMemory",
            dimensions=dims,
            statistics=["Maximum"],
            period=300,
        )
        if mem_dp:
            max_vals = [dp.get("Maximum", 0.0) for dp in mem_dp if dp.get("Maximum") is not None]
            if max_vals:
                data["free_memory_bytes"] = max(max_vals)
            else:
                data["free_memory_bytes"] = None
        else:
            data["free_memory_bytes"] = None

        # --- Free Storage Space (Current/Latest value) ---
        storage_dp = self._get_cloudwatch_metric(
            namespace="AWS/RDS",
            metric_name="FreeStorageSpace",
            dimensions=dims,
            statistics=["Average"],
            period=300,
        )
        if storage_dp:
            # Get the chronologically last datapoint for "current"
            sorted_dp = sorted(storage_dp, key=lambda x: x["Timestamp"])
            data["free_storage_bytes"] = sorted_dp[-1].get("Average", 0.0)
        else:
            data["free_storage_bytes"] = None

        # --- Network Transmit Throughput (Max of 5-min Maxes) ---
        tx_dp = self._get_cloudwatch_metric(
            namespace="AWS/RDS",
            metric_name="NetworkTransmitThroughput",
            dimensions=dims,
            statistics=["Maximum"],
            period=300,
        )
        if tx_dp:
            max_vals = [dp.get("Maximum", 0.0) for dp in tx_dp if dp.get("Maximum") is not None]
            data["network_tx_bytes_sec"] = max(max_vals) if max_vals else None
        else:
            data["network_tx_bytes_sec"] = None

        # --- Network Receive Throughput (Max of 5-min Maxes) ---
        rx_dp = self._get_cloudwatch_metric(
            namespace="AWS/RDS",
            metric_name="NetworkReceiveThroughput",
            dimensions=dims,
            statistics=["Maximum"],
            period=300,
        )
        if rx_dp:
            max_vals = [dp.get("Maximum", 0.0) for dp in rx_dp if dp.get("Maximum") is not None]
            data["network_rx_bytes_sec"] = max(max_vals) if max_vals else None
        else:
            data["network_rx_bytes_sec"] = None

        # --- Database Connections (Max of 5-min Maxes) ---
        conn_dp = self._get_cloudwatch_metric(
            namespace="AWS/RDS",
            metric_name="DatabaseConnections",
            dimensions=dims,
            statistics=["Maximum"],
            period=300,
        )
        if conn_dp:
            max_vals = [dp.get("Maximum", 0.0) for dp in conn_dp if dp.get("Maximum") is not None]
            data["db_connections"] = max(max_vals) if max_vals else None
        else:
            data["db_connections"] = None

        return data
