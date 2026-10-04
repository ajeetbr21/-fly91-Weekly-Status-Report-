"""
ELB collector module for Application Load Balancer metrics.

Retrieves all ALBs and their CloudWatch metrics including request counts,
connections, processed bytes, response times, and consumed LCUs.
"""

from typing import Any, Dict, List

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class ELBCollector(BaseCollector):
    """Collects Application Load Balancer data and CloudWatch metrics.

    For each ALB:
    - Request count, active/new connections
    - Processed bytes, response time
    - Consumed LCUs
    - HTTP 3XX redirect count

    Network Load Balancers and Classic ELBs are currently excluded.
    """

    def collect(self) -> List[Dict[str, Any]]:
        """Collect ALB inventory and metrics.

        Returns:
            List of ALB detail dicts.
        """
        self.logger.info("Collecting ELB data for region %s…", self.region)

        elbv2 = self._get_client("elbv2")
        albs: List[Dict[str, Any]] = []

        try:
            paginator = elbv2.get_paginator("describe_load_balancers")
            for page in paginator.paginate():
                for lb in page.get("LoadBalancers", []):
                    if lb.get("Type") == "application":
                        albs.append(lb)
        except (ClientError, BotoCoreError) as exc:
            self.logger.error("Failed to describe load balancers: %s", exc)
            return []

        self.logger.info("Found %d ALBs; collecting metrics…", len(albs))

        results: List[Dict[str, Any]] = []
        for lb in albs:
            lb_data = self._collect_alb_metrics(lb)
            results.append(lb_data)

        self.logger.info("ELB collection complete — %d ALBs.", len(results))
        return results

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _extract_lb_dimension(arn: str) -> str:
        """Extract the CloudWatch dimension value from an ALB ARN.

        The dimension format is ``app/<name>/<id>`` extracted from the
        full ARN after ``':loadbalancer/'``.
        """
        try:
            return arn.split(":loadbalancer/")[1]
        except (IndexError, AttributeError):
            return ""

    def _collect_alb_metrics(self, lb: dict) -> Dict[str, Any]:
        """Collect all CloudWatch metrics for a single ALB."""
        name = lb.get("LoadBalancerName", "unknown")
        arn = lb.get("LoadBalancerArn", "")
        lb_dimension = self._extract_lb_dimension(arn)
        dims = [{"Name": "LoadBalancer", "Value": lb_dimension}]

        data: Dict[str, Any] = {
            "name": name,
            "arn": arn,
            "dns_name": lb.get("DNSName", ""),
            "state": lb.get("State", {}).get("Code", "unknown"),
        }

        # ----------------------------------------------------------
        # 5-minute Sum-based metrics
        # ----------------------------------------------------------
        sum_metrics = [
            ("RequestCount", "requests"),
            ("ActiveConnectionCount", "active_connections"),
            ("NewConnectionCount", "new_connections"),
            ("ConsumedLCUs", "consumed_lcus"),
            ("ProcessedBytes", "processed_bytes"),
            ("HTTPCode_ELB_3XX_Count", "http_redirect_count"),
        ]

        for metric_name, key in sum_metrics:
            datapoints = self._get_cloudwatch_metric(
                namespace="AWS/ApplicationELB",
                metric_name=metric_name,
                dimensions=dims,
                statistics=["Sum"],
                period=300, # 5 minutes
            )
            if datapoints:
                # The user explicitly wants the "Max" of the 5-minute "Sum" data
                peak_val = max(dp.get("Sum", 0.0) for dp in datapoints if dp.get("Sum") is not None)
                data[key] = round(peak_val, 2)
            else:
                data[key] = 0.0

        # ----------------------------------------------------------
        # 5-minute Average-based metrics (TargetResponseTime)
        # ----------------------------------------------------------
        avg_metrics = [
            ("TargetResponseTime", "target_response_time"),
        ]
        for metric_name, key in avg_metrics:
            datapoints = self._get_cloudwatch_metric(
                namespace="AWS/ApplicationELB",
                metric_name=metric_name,
                dimensions=dims,
                statistics=["Average"],
                period=300, # 5 minutes
            )
            if datapoints:
                avg_val = sum(dp.get("Average", 0.0) for dp in datapoints if dp.get("Average") is not None) / len(datapoints)
                data[key] = round(avg_val, 4)
            else:
                data[key] = None

        return data
