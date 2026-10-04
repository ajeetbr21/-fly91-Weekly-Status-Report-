"""
Amazon Inspector collector module for Inspector v2 (inspector2) findings.

Retrieves Amazon Inspector findings for the reporting window and
aggregates them by severity (CRITICAL / HIGH / MEDIUM / LOW /
INFORMATIONAL).  Each finding is summarised with its title, severity,
affected resource, finding type, CVE identifier, first-observed date,
and status.
"""

from datetime import datetime, time, timezone
from typing import Any, Dict, List, Optional, Tuple

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class InspectorCollector(BaseCollector):
    """Collects Amazon Inspector v2 findings for the reporting period.

    Queries the Inspector v2 ``list_findings`` API filtered to the
    reporting window (``self.start_date`` .. ``self.end_date``) and
    aggregates the results by severity.  Returns a dict with the
    date range, per-severity counts, total findings, and a detailed
    findings list.
    """

    SEVERITY_ORDER = ["CRITICAL", "HIGH", "MEDIUM", "LOW", "INFORMATIONAL"]

    def collect(self) -> Dict[str, Any]:
        """Collect Amazon Inspector findings for the reporting window.

        Returns:
            A dict with keys ``date_range``, ``severity_counts``,
            ``total_findings`` and ``findings``.
        """
        self.logger.info("Collecting Amazon Inspector findings…")

        severity_counts: Dict[str, int] = {sev: 0 for sev in self.SEVERITY_ORDER}
        findings: List[Dict[str, Any]] = []

        raw_findings, collection_status, collection_error = self._list_findings()
        for raw in raw_findings:
            finding = self._parse_finding(raw)
            severity = finding["severity"]
            if severity in severity_counts:
                severity_counts[severity] += 1
            else:
                severity_counts[severity] = 1
            findings.append(finding)

        result: Dict[str, Any] = {
            "date_range": self._format_date_range(),
            "severity_counts": severity_counts,
            "total_findings": len(findings),
            "findings": findings,
            # collection_status is "ok" when the API call succeeded (even
            # with zero findings) and "error" when it failed, so the sheet
            # can distinguish a clean week from a failed/denied query.
            "collection_status": collection_status,
            "collection_error": collection_error,
        }

        self.logger.info(
            "Inspector collection complete — %d findings.", len(findings)
        )
        return result

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _list_findings(self) -> Tuple[List[dict], str, Optional[str]]:
        """List Inspector v2 findings filtered to the reporting window.

        Returns a ``(findings, status, error)`` tuple.  ``status`` is
        ``"ok"`` when the API call completed (even if zero findings were
        returned) and ``"error"`` when it failed.  Failure is still
        fail-open (the report never crashes), but the status/error are
        surfaced so the sheet can render a distinct "collection failed"
        state instead of a silently-empty table.
        """
        inspector = self._get_client("inspector2")
        results: List[dict] = []

        start_dt = datetime.combine(self.start_date, time.min, tzinfo=timezone.utc)
        end_dt = datetime.combine(self.end_date, time.max, tzinfo=timezone.utc)

        filter_criteria = {
            "firstObservedAt": [
                {"startInclusive": start_dt, "endInclusive": end_dt}
            ]
        }

        try:
            params: Dict[str, Any] = {
                "filterCriteria": filter_criteria,
                "maxResults": 100,
            }
            while True:
                response = inspector.list_findings(**params)
                results.extend(response.get("findings", []))
                next_token = response.get("nextToken")
                if not next_token:
                    break
                params["nextToken"] = next_token
        except (ClientError, BotoCoreError) as exc:
            self.logger.warning("Failed to list Inspector findings: %s", exc)
            return results, "error", str(exc)
        return results, "ok", None

    def _parse_finding(self, raw: dict) -> Dict[str, Any]:
        """Normalise a raw Inspector v2 finding into the report shape."""
        severity = (raw.get("severity") or "INFORMATIONAL").upper()

        # Resource details (first resource reported for the finding)
        resources = raw.get("resources") or []
        resource_type = "-"
        resource_id = "-"
        if resources:
            resource_type = resources[0].get("type", "-")
            resource_id = resources[0].get("id", "-")

        # CVE / vulnerability identifier
        cve = "-"
        pkg_details = raw.get("packageVulnerabilityDetails") or {}
        if pkg_details.get("vulnerabilityId"):
            cve = pkg_details["vulnerabilityId"]

        first_observed = self._format_timestamp(raw.get("firstObservedAt"))

        return {
            "title": raw.get("title", "-"),
            "severity": severity,
            "resource_type": resource_type,
            "resource_id": resource_id,
            "finding_type": raw.get("type", "-"),
            "cve": cve,
            "first_observed": first_observed,
            "status": raw.get("status", "ACTIVE"),
        }

    def _format_timestamp(self, value: Any) -> str:
        """Format a boto3 datetime (or ISO string) as YYYY-MM-DD."""
        if value is None:
            return "-"
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d")
        try:
            return str(value)[:10]
        except (TypeError, ValueError):
            return "-"

    def _format_date_range(self) -> str:
        """Return the reporting window as 'YYYY-MM-DD to YYYY-MM-DD'."""
        return (
            f"{self.start_date.strftime('%Y-%m-%d')} to "
            f"{self.end_date.strftime('%Y-%m-%d')}"
        )
