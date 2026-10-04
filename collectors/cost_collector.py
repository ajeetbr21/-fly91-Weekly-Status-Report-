"""
Cost collector module for AWS Cost Explorer data.

Retrieves current and previous week costs, daily trends, and per-service
breakdowns using the AWS Cost Explorer API.
"""

from datetime import timedelta
from typing import Any, Dict, List

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class CostCollector(BaseCollector):
    """Collects cost data from AWS Cost Explorer.

    Cost Explorer is a GLOBAL service — all calls are routed to us-east-1
    regardless of the configured region.

    Collects:
    - Current week total and daily costs
    - Previous week total for comparison
    - Per-service cost breakdown (top 10 and full list)
    - Week-over-week change metrics
    """

    # Cost Explorer is always in us-east-1
    _CE_REGION = "us-east-1"

    def collect(self) -> dict:
        """Collect cost data for the reporting period.

        Returns:
            Dictionary with keys:
                current_week_total, previous_week_total, difference,
                pct_change, avg_daily_cost, daily_costs, top_services,
                service_breakdown
        """
        self.logger.info("Collecting cost data from Cost Explorer…")

        ce = self._get_client("ce", region_override=self._CE_REGION)

        # ----------------------------------------------------------
        # Date ranges
        # ----------------------------------------------------------
        # CE end date is EXCLUSIVE — add one day
        current_start = self.start_date
        current_end = self.end_date + timedelta(days=1)

        previous_end = self.start_date  # exclusive, equals current start
        previous_start = previous_end - timedelta(days=7)

        # ----------------------------------------------------------
        # Current week — total (no GroupBy)
        # ----------------------------------------------------------
        current_week_total = self._get_total_cost(
            ce, current_start, current_end
        )

        # ----------------------------------------------------------
        # Previous week — total (no GroupBy)
        # ----------------------------------------------------------
        previous_week_total = self._get_total_cost(
            ce, previous_start, previous_end
        )

        # ----------------------------------------------------------
        # Daily cost trend (current week, no GroupBy)
        # ----------------------------------------------------------
        daily_costs = self._get_daily_costs(ce, current_start, current_end)

        # ----------------------------------------------------------
        # Service breakdown (current week)
        # ----------------------------------------------------------
        service_breakdown = self._get_service_breakdown(
            ce, current_start, current_end
        )

        # ----------------------------------------------------------
        # Service breakdown (previous week) to calculate changes
        # ----------------------------------------------------------
        prev_service_breakdown = self._get_service_breakdown(
            ce, previous_start, previous_end
        )
        prev_map = {s["service"]: s["cost"] for s in prev_service_breakdown}

        for s in service_breakdown:
            svc_name = s["service"]
            curr_cost = s["cost"]
            prev_cost = prev_map.get(svc_name, 0.0)
            diff = curr_cost - prev_cost
            if prev_cost > 0:
                s["pct_change"] = round((diff / prev_cost) * 100.0, 2)
            else:
                s["pct_change"] = 100.0 if curr_cost > 0 else 0.0
            s["difference"] = round(diff, 2)

        # ----------------------------------------------------------
        # Derived metrics
        # ----------------------------------------------------------
        difference = current_week_total - previous_week_total
        if previous_week_total > 0:
            pct_change = (difference / previous_week_total) * 100.0
        else:
            pct_change = 0.0 if current_week_total == 0 else 100.0

        num_days = max(len(daily_costs), 1)
        avg_daily_cost = current_week_total / num_days

        # Top 10 services by cost descending
        sorted_services = sorted(
            service_breakdown, key=lambda s: s["cost"], reverse=True
        )
        top_services = sorted_services[:10]

        result: Dict[str, Any] = {
            "current_week_total": round(current_week_total, 2),
            "previous_week_total": round(previous_week_total, 2),
            "difference": round(difference, 2),
            "pct_change": round(pct_change, 2),
            "avg_daily_cost": round(avg_daily_cost, 2),
            "daily_costs": daily_costs,
            "top_services": top_services,
            "service_breakdown": sorted_services,
        }

        self.logger.info(
            "Cost collection complete — current week: $%.2f, previous week: $%.2f",
            current_week_total,
            previous_week_total,
        )
        return result

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _get_total_cost(self, ce, start, end) -> float:
        """Get total unblended cost for a date range (no grouping)."""
        start_str = start.strftime("%Y-%m-%d")
        end_str = end.strftime("%Y-%m-%d")

        try:
            response = ce.get_cost_and_usage(
                TimePeriod={"Start": start_str, "End": end_str},
                Granularity="DAILY",
                Metrics=["UnblendedCost"],
            )
            total = 0.0
            for result in response.get("ResultsByTime", []):
                amount = result.get("Total", {}).get(
                    "UnblendedCost", {}
                ).get("Amount", "0")
                total += self._safe_float(amount)
            return total
        except (ClientError, BotoCoreError) as exc:
            self.logger.error("Failed to get total cost (%s→%s): %s", start_str, end_str, exc)
            return 0.0

    def _get_daily_costs(self, ce, start, end) -> List[Dict[str, Any]]:
        """Get daily cost breakdown for a date range."""
        start_str = start.strftime("%Y-%m-%d")
        end_str = end.strftime("%Y-%m-%d")

        try:
            response = ce.get_cost_and_usage(
                TimePeriod={"Start": start_str, "End": end_str},
                Granularity="DAILY",
                Metrics=["UnblendedCost"],
            )
            daily: List[Dict[str, Any]] = []
            for result in response.get("ResultsByTime", []):
                day_date = result["TimePeriod"]["Start"]
                amount = result.get("Total", {}).get(
                    "UnblendedCost", {}
                ).get("Amount", "0")
                daily.append(
                    {
                        "date": day_date,
                        "cost": round(self._safe_float(amount), 2),
                    }
                )
            return daily
        except (ClientError, BotoCoreError) as exc:
            self.logger.error("Failed to get daily costs: %s", exc)
            return []

    def _get_service_breakdown(self, ce, start, end) -> List[Dict[str, Any]]:
        """Get cost grouped by AWS service for a date range."""
        start_str = start.strftime("%Y-%m-%d")
        end_str = end.strftime("%Y-%m-%d")

        try:
            response = ce.get_cost_and_usage(
                TimePeriod={"Start": start_str, "End": end_str},
                Granularity="DAILY",
                Metrics=["UnblendedCost"],
                GroupBy=[{"Type": "DIMENSION", "Key": "SERVICE"}],
            )

            # Aggregate across days
            service_totals: Dict[str, float] = {}
            for result in response.get("ResultsByTime", []):
                for group in result.get("Groups", []):
                    service_name = group["Keys"][0]
                    amount = group["Metrics"]["UnblendedCost"]["Amount"]
                    service_totals[service_name] = service_totals.get(
                        service_name, 0.0
                    ) + self._safe_float(amount)

            breakdown: List[Dict[str, Any]] = [
                {"service": svc, "cost": round(cost, 2)}
                for svc, cost in service_totals.items()
            ]
            return breakdown
        except (ClientError, BotoCoreError) as exc:
            self.logger.error("Failed to get service breakdown: %s", exc)
            return []
