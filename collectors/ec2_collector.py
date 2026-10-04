"""
EC2 collector module for instance inventory and performance metrics.

Retrieves all EC2 instances with their configuration details, state, and
CloudWatch metrics (CPU, memory, disk, network) for running instances.
"""

from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class EC2Collector(BaseCollector):
    """Collects EC2 instance data and associated CloudWatch metrics.

    For each instance:
    - Configuration: name, ID, type, state, AZ, region
    - CPU utilisation (min / max / avg) from AWS/EC2
    - Memory usage from CWAgent (if CloudWatch Agent installed)
    - Disk usage from CWAgent (all mount points)
    - Network I/O from AWS/EC2

    Only RUNNING instances have metrics collected; stopped instances
    receive ``None`` for all metric fields.
    """

    def collect(self) -> dict:
        """Collect EC2 instance inventory and metrics.

        Returns:
            Dictionary with key ``instances`` containing a list of
            instance detail dicts.
        """
        self.logger.info("Collecting EC2 instance data for region %s…", self.region)

        ec2 = self._get_client("ec2")
        instances_data: List[Dict[str, Any]] = []

        try:
            paginator = ec2.get_paginator("describe_instances")
            for page in paginator.paginate():
                for reservation in page.get("Reservations", []):
                    for instance in reservation.get("Instances", []):
                        inst = self._parse_instance(instance)
                        instances_data.append(inst)
        except (ClientError, BotoCoreError) as exc:
            self.logger.error("Failed to describe EC2 instances: %s", exc)
            return []

        self.logger.info(
            "Found %d EC2 instances; collecting metrics for running instances…",
            len(instances_data),
        )

        for inst in instances_data:
            if inst["state"] == "running":
                self._collect_instance_metrics(inst)
            else:
                inst.update(
                    {
                        "cpu_min": None,
                        "cpu_max": None,
                        "cpu_avg": None,
                        "memory_min": None,
                        "memory_max": None,
                        "memory_avg": None,
                        "disk_utilization": None,
                        "network_in": None,
                        "network_out": None,
                        "network_packets_in": None,
                        "network_packets_out": None,
                    }
                )

        self.logger.info("EC2 collection complete — %d instances.", len(instances_data))
        return instances_data

    # ------------------------------------------------------------------
    # Instance parsing
    # ------------------------------------------------------------------

    @staticmethod
    def _get_tag_value(tags: List[dict], key: str, default: str = "N/A") -> str:
        """Extract a tag value from an EC2 tag list."""
        if not tags:
            return default
        for tag in tags:
            if tag.get("Key") == key:
                return tag.get("Value", default)
        return default

    def _parse_instance(self, instance: dict) -> Dict[str, Any]:
        """Parse raw EC2 instance dict into a flat summary dict."""
        az = instance.get("Placement", {}).get("AvailabilityZone", "")
        region = az[:-1] if az else self.region

        return {
            "name": self._get_tag_value(instance.get("Tags", []), "Name"),
            "instance_id": instance.get("InstanceId", ""),
            "instance_type": instance.get("InstanceType", ""),
            "state": instance.get("State", {}).get("Name", "unknown"),
            "availability_zone": az,
            "region": region,
        }

    # ------------------------------------------------------------------
    # Metric collection
    # ------------------------------------------------------------------

    def _collect_instance_metrics(self, inst: Dict[str, Any]) -> None:
        """Populate *inst* in place with CloudWatch metrics."""
        instance_id = inst["instance_id"]
        ec2_dims = [{"Name": "InstanceId", "Value": instance_id}]
        period = self._calculate_period_seconds()

        # --- CPU ---
        cpu_data = self._get_cloudwatch_metric(
            namespace="AWS/EC2",
            metric_name="CPUUtilization",
            dimensions=ec2_dims,
            statistics=["Maximum"],
            period=300, # 5 minutes
        )
        if cpu_data:
            max_vals = [dp.get("Maximum", 0.0) for dp in cpu_data if dp.get("Maximum") is not None]
            if max_vals:
                inst["cpu_min"] = round(min(max_vals), 2)
                inst["cpu_max"] = round(max(max_vals), 2)
                inst["cpu_avg"] = round(sum(max_vals) / len(max_vals), 2)
            else:
                inst["cpu_min"] = None
                inst["cpu_max"] = None
                inst["cpu_avg"] = None
        else:
            inst["cpu_min"] = None
            inst["cpu_max"] = None
            inst["cpu_avg"] = None

        # --- Memory (CWAgent) ---
        self._collect_memory_metrics(inst)

        # --- Disk (CWAgent) ---
        inst["disk_utilization"] = self._get_disk_metric(instance_id)

        # --- Network ---
        self._collect_network_metrics(inst, ec2_dims)

        # --- Raw weekly time-series for CloudWatch-style charts ---
        self._collect_metric_series(inst, ec2_dims)

    def _collect_metric_series(
        self, inst: Dict[str, Any], ec2_dims: List[Dict[str, str]]
    ) -> None:
        """Attach raw weekly CloudWatch time-series to *inst* for charting.

        Populates ``inst['metric_series']`` with a list of series dicts of the
        shape consumed by ``report.metric_charts.render_metric_chart``:

            {'label', 'unit', 'timestamps': [datetime,...],
             'values': [float,...]}

        Collects CPUUtilization (AWS/EC2), NetworkIn/NetworkOut (AWS/EC2) and,
        when the CloudWatch Agent is present, mem_used_percent and
        disk_used_percent (CWAgent). Every call is guarded so a failure or an
        empty result simply yields no series for that metric rather than
        raising. Hourly aggregation (period=3600) keeps the series compact for
        a week-long window. Existing scalar fields are left untouched.
        """
        instance_id = inst["instance_id"]
        series: List[Dict[str, Any]] = []

        # CPU utilisation (percent).
        cpu_series = self._metric_series_from_datapoints(
            namespace="AWS/EC2",
            metric_name="CPUUtilization",
            dimensions=ec2_dims,
            statistic="Average",
            label="CPUUtilization",
            unit="%",
        )
        if cpu_series:
            series.append(cpu_series)

        # Network throughput (bytes).
        for metric_name, label in (
            ("NetworkIn", "NetworkIn"),
            ("NetworkOut", "NetworkOut"),
        ):
            net_series = self._metric_series_from_datapoints(
                namespace="AWS/EC2",
                metric_name=metric_name,
                dimensions=ec2_dims,
                statistic="Average",
                label=label,
                unit="Bytes",
            )
            if net_series:
                series.append(net_series)

        # Memory (CWAgent) - only when datapoints exist.
        mem_series = self._cwagent_series(
            instance_id, "mem_used_percent", "mem_used_percent", "%"
        )
        if mem_series:
            series.append(mem_series)

        # Disk (CWAgent, primary mount) - only when datapoints exist.
        disk_series = self._cwagent_series(
            instance_id, "disk_used_percent", "disk_used_percent", "%",
            prefer_paths=("/", "C:", "D:"),
        )
        if disk_series:
            series.append(disk_series)

        inst["metric_series"] = series

    def _metric_series_from_datapoints(
        self,
        namespace: str,
        metric_name: str,
        dimensions: List[Dict[str, str]],
        statistic: str,
        label: str,
        unit: str,
    ) -> Optional[Dict[str, Any]]:
        """Query one CloudWatch metric and return a chart series dict or None."""
        try:
            datapoints = self._get_cloudwatch_metric(
                namespace=namespace,
                metric_name=metric_name,
                dimensions=dimensions,
                statistics=[statistic],
                period=3600,  # hourly
            )
        except Exception as exc:  # defensive: never raise from charting path
            self.logger.debug(
                "metric_series query failed for %s/%s: %s",
                namespace, metric_name, exc,
            )
            return None
        return self._build_series(datapoints, statistic, label, unit)

    def _cwagent_series(
        self,
        instance_id: str,
        metric_name: str,
        label: str,
        unit: str,
        prefer_paths: Optional[tuple] = None,
    ) -> Optional[Dict[str, Any]]:
        """Return a CWAgent metric series (first matching dimension set)."""
        cw = self._get_client("cloudwatch")
        try:
            response = cw.list_metrics(
                Namespace="CWAgent",
                MetricName=metric_name,
                Dimensions=[{"Name": "InstanceId", "Value": instance_id}],
            )
        except (ClientError, BotoCoreError) as exc:
            self.logger.debug(
                "list_metrics failed for CWAgent %s on %s: %s",
                metric_name, instance_id, exc,
            )
            return None

        metrics = response.get("Metrics", [])
        if not metrics:
            return None

        chosen_dims = metrics[0].get("Dimensions", [])
        if prefer_paths:
            for metric in metrics:
                dims = metric.get("Dimensions", [])
                # Linux CWAgent exposes the volume under the 'path' dimension
                # (/, /var); Windows CWAgent uses the 'instance' dimension
                # (C:, D:). Check both so the preferred volume is picked on
                # either platform instead of falling back to metrics[0]
                # (which may be _Total). Mirrors _get_disk_metric.
                volume = self._extract_dimension(dims, "instance") or \
                    self._extract_dimension(dims, "path")
                if volume in prefer_paths:
                    chosen_dims = dims
                    break

        try:
            datapoints = self._get_cloudwatch_metric(
                namespace="CWAgent",
                metric_name=metric_name,
                dimensions=chosen_dims,
                statistics=["Average"],
                period=3600,
            )
        except Exception as exc:  # defensive
            self.logger.debug(
                "CWAgent metric_series query failed for %s on %s: %s",
                metric_name, instance_id, exc,
            )
            return None
        return self._build_series(datapoints, "Average", label, unit)

    @staticmethod
    def _build_series(
        datapoints: List[dict], statistic: str, label: str, unit: str
    ) -> Optional[Dict[str, Any]]:
        """Convert CloudWatch datapoints into a chart series dict (or None)."""
        if not datapoints:
            return None
        timestamps = []
        values = []
        for dp in datapoints:
            ts = dp.get("Timestamp")
            val = dp.get(statistic)
            if ts is None or val is None:
                continue
            timestamps.append(ts)
            values.append(float(val))
        if not timestamps:
            return None
        return {
            "label": label,
            "unit": unit,
            "timestamps": timestamps,
            "values": values,
        }

    def _collect_memory_metrics(self, inst: Dict[str, Any]) -> None:
        """Get Min, Max, and Avg memory usage from CWAgent metrics."""
        instance_id = inst["instance_id"]
        cw = self._get_client("cloudwatch")

        for metric_name in ("mem_used_percent", "Memory % Committed Bytes In Use"):
            try:
                response = cw.list_metrics(
                    Namespace="CWAgent",
                    MetricName=metric_name,
                    Dimensions=[{"Name": "InstanceId", "Value": instance_id}],
                )
                metrics = response.get("Metrics", [])
                if metrics:
                    discovered_dims = metrics[0].get("Dimensions", [])
                    datapoints = self._get_cloudwatch_metric(
                        namespace="CWAgent",
                        metric_name=metric_name,
                        dimensions=discovered_dims,
                        statistics=["Maximum"],
                        period=300, # 5 minutes
                    )
                    if datapoints:
                        max_vals = [dp.get("Maximum", 0.0) for dp in datapoints if dp.get("Maximum") is not None]
                        if max_vals:
                            inst["memory_max"] = round(sum(max_vals) / len(max_vals), 2)
                        else:
                            inst["memory_max"] = None
                        return
            except (ClientError, BotoCoreError) as exc:
                self.logger.debug(
                    "list_metrics failed for memory metric %s on %s: %s",
                    metric_name,
                    instance_id,
                    exc,
                )
        inst["memory_max"] = None

    def _get_disk_metric(self, instance_id: str) -> Optional[str]:
        """Get disk usage from CWAgent for all mount points, including Min, Max, and Avg.

        Supports Linux (disk_used_percent with 'path' dimension) and
        Windows (LogicalDisk % Free Space with 'instance' dimension).
        """
        cw = self._get_client("cloudwatch")

        for metric_name in ("disk_used_percent", "LogicalDisk % Free Space"):
            try:
                response = cw.list_metrics(
                    Namespace="CWAgent",
                    MetricName=metric_name,
                    Dimensions=[{"Name": "InstanceId", "Value": instance_id}],
                )
                metrics = response.get("Metrics", [])
                if not metrics:
                    continue

                path_usages: List[str] = []
                for metric in metrics:
                    dims = metric.get("Dimensions", [])
                    
                    # Windows uses 'instance' (C:, D:), Linux uses 'path' (/, /var)
                    label = self._extract_dimension(dims, "instance") or \
                            self._extract_dimension(dims, "path") or \
                            self._extract_dimension(dims, "device") or "?"
                    
                    if label == "_Total" or label == "?" or label.startswith(("/snap", "/run", "/dev", "/boot", "/sys")):
                        continue

                    # For Linux, strictly only / is considered primary. For Windows, C: and D: are common.
                    # We will only process /, C:, and D:
                    if label not in ("/", "C:", "D:"):
                        continue

                    datapoints = self._get_cloudwatch_metric(
                        namespace="CWAgent",
                        metric_name=metric_name,
                        dimensions=dims,
                        statistics=["Maximum"],
                        period=300, # 5 minutes
                    )
                    if datapoints:
                        max_vals = [dp.get("Maximum", 0.0) for dp in datapoints if dp.get("Maximum") is not None]
                        if max_vals:
                            d_max = round(max(max_vals), 1)
                            path_usages.append(f"{label} {d_max}%")

                if path_usages:
                    return " ".join(path_usages)
            except (ClientError, BotoCoreError) as exc:
                self.logger.debug(
                    "list_metrics failed for disk metric %s on %s: %s",
                    metric_name,
                    instance_id,
                    exc,
                )
        return None

    @staticmethod
    def _extract_dimension(
        dimensions: List[Dict[str, str]], name: str
    ) -> Optional[str]:
        """Extract a dimension value by name from a list of dimensions."""
        for dim in dimensions:
            if dim.get("Name") == name:
                return dim.get("Value")
        return None

    def _collect_network_metrics(
        self, inst: Dict[str, Any], ec2_dims: List[Dict[str, str]]
    ) -> None:
        """Collect network I/O metrics and populate *inst* in place.

        Queries 5-minute Maximums and finds the peak value.
        """
        network_metrics = {
            "NetworkIn": "network_in",
            "NetworkOut": "network_out",
            "NetworkPacketsIn": "network_packets_in",
            "NetworkPacketsOut": "network_packets_out",
        }

        for metric_name, key in network_metrics.items():
            datapoints = self._get_cloudwatch_metric(
                namespace="AWS/EC2",
                metric_name=metric_name,
                dimensions=ec2_dims,
                statistics=["Maximum"],
                period=300, # 5 minutes
            )
            if datapoints:
                peak = max(dp.get("Maximum", 0.0) for dp in datapoints if dp.get("Maximum") is not None)
                inst[key] = round(peak, 2)
            else:
                inst[key] = None
