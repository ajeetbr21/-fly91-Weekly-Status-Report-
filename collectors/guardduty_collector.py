"""
Amazon GuardDuty collector module.

Retrieves Amazon GuardDuty findings for the reporting window and
aggregates them by severity label (HIGH / MEDIUM / LOW).  GuardDuty
severity is reported as a numeric score (1.0-8.9); this collector maps
that score to a label:

    Low    : 1.0 - 3.9
    Medium : 4.0 - 6.9
    High   : 7.0 - 8.9
"""

from datetime import datetime, time, timezone
from typing import Any, Dict, List, Optional, Tuple

from botocore.exceptions import ClientError, BotoCoreError

from .base_collector import BaseCollector


class GuardDutyCollector(BaseCollector):
    """Collects Amazon GuardDuty findings for the reporting period.

    For each GuardDuty detector in the region, lists findings created
    within the reporting window, retrieves their detail, and aggregates
    them by severity label.  Returns a dict with the date range,
    per-severity counts, total findings, and a detailed findings list.
    """

    def collect(self) -> Dict[str, Any]:
        """Collect Amazon GuardDuty findings for the reporting window.

        Returns:
            A dict with keys ``date_range``, ``severity_counts``,
            ``total_findings`` and ``findings``.
        """
        self.logger.info("Collecting Amazon GuardDuty findings…")

        severity_counts: Dict[str, int] = {"HIGH": 0, "MEDIUM": 0, "LOW": 0}
        findings: List[Dict[str, Any]] = []

        raw_findings, collection_status, collection_error = self._list_findings()
        for raw in raw_findings:
            finding = self._parse_finding(raw)
            label = finding["severity_label"].upper()
            if label in severity_counts:
                severity_counts[label] += 1
            else:
                severity_counts[label] = 1
            findings.append(finding)

        result: Dict[str, Any] = {
            "date_range": self._format_date_range(),
            "severity_counts": severity_counts,
            "total_findings": len(findings),
            "findings": findings,
            # collection_status is "ok" when the API calls succeeded (even
            # with zero findings) and "error" when they failed, so the sheet
            # can distinguish a clean week from a failed/denied query.
            "collection_status": collection_status,
            "collection_error": collection_error,
        }

        self.logger.info(
            "GuardDuty collection complete — %d findings.", len(findings)
        )
        return result

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _list_findings(self) -> Tuple[List[dict], str, Optional[str]]:
        """List GuardDuty findings (with detail) for all detectors.

        Returns a ``(findings, status, error)`` tuple.  ``status`` is
        ``"ok"`` when the API calls completed (even if zero findings were
        returned) and ``"error"`` when they failed.  Behaviour remains
        fail-open (the report never crashes), but the status/error are
        surfaced so the sheet can render a distinct "collection failed"
        state rather than a silently-empty table.

        Time-window semantics: GuardDuty keeps a finding open and bumps
        ``updatedAt`` as the same activity recurs, so a threat created
        before the window but still firing during it has a ``createdAt``
        outside the window.  To capture findings *active during* the week
        (the intended "one week" meaning for a BAU security review) we
        filter on ``updatedAt`` rather than ``createdAt``; ``updatedAt``
        equals ``createdAt`` for findings that are both created and last
        active in the window, so new findings are still included.  The
        exact semantics could not be confirmed against the reference
        document, so this is a defensible choice documented here and in
        FEAT-003 findings.
        """
        guardduty = self._get_client("guardduty")
        results: List[dict] = []

        start_ms = int(
            datetime.combine(
                self.start_date, time.min, tzinfo=timezone.utc
            ).timestamp()
            * 1000
        )
        end_ms = int(
            datetime.combine(
                self.end_date, time.max, tzinfo=timezone.utc
            ).timestamp()
            * 1000
        )

        finding_criteria = {
            "Criterion": {
                "updatedAt": {"GreaterThanOrEqual": start_ms, "LessThanOrEqual": end_ms}
            }
        }

        try:
            detector_ids: List[str] = []
            det_params: Dict[str, Any] = {}
            while True:
                det_resp = guardduty.list_detectors(**det_params)
                detector_ids.extend(det_resp.get("DetectorIds", []))
                next_token = det_resp.get("NextToken")
                if not next_token:
                    break
                det_params["NextToken"] = next_token

            for detector_id in detector_ids:
                finding_ids: List[str] = []
                list_params: Dict[str, Any] = {
                    "DetectorId": detector_id,
                    "FindingCriteria": finding_criteria,
                }
                while True:
                    list_resp = guardduty.list_findings(**list_params)
                    finding_ids.extend(list_resp.get("FindingIds", []))
                    next_token = list_resp.get("NextToken")
                    if not next_token:
                        break
                    list_params["NextToken"] = next_token

                # get_findings accepts up to 50 IDs per request
                for chunk_start in range(0, len(finding_ids), 50):
                    chunk = finding_ids[chunk_start:chunk_start + 50]
                    get_resp = guardduty.get_findings(
                        DetectorId=detector_id, FindingIds=chunk
                    )
                    results.extend(get_resp.get("Findings", []))
        except (ClientError, BotoCoreError) as exc:
            self.logger.warning("Failed to list GuardDuty findings: %s", exc)
            return results, "error", str(exc)
        return results, "ok", None

    def _parse_finding(self, raw: dict) -> Dict[str, Any]:
        """Normalise a raw GuardDuty finding into the report shape."""
        score = self._safe_float(raw.get("Severity"), 0.0)
        label = self._severity_label(score)

        resource = raw.get("Resource") or {}
        resource_type = resource.get("ResourceType", "-")

        service = raw.get("Service") or {}
        count = int(service.get("Count", 1) or 1)
        first_seen = self._format_timestamp(service.get("EventFirstSeen"))
        last_seen = self._format_timestamp(service.get("EventLastSeen"))

        return {
            "title": raw.get("Title", "-"),
            "type": raw.get("Type", "-"),
            "severity_label": label,
            "severity_score": round(score, 1),
            "resource_type": resource_type,
            "region": raw.get("Region", self.region),
            "count": count,
            "first_seen": first_seen,
            "last_seen": last_seen,
        }

    @staticmethod
    def _severity_label(score: float) -> str:
        """Map a GuardDuty numeric severity score to a label."""
        if score >= 7.0:
            return "HIGH"
        if score >= 4.0:
            return "MEDIUM"
        return "LOW"

    def _format_timestamp(self, value: Any) -> str:
        """Format a GuardDuty timestamp (ISO string or epoch) as YYYY-MM-DD."""
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
