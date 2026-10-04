#!/usr/bin/env python3
"""
Standalone verification for the Word report's LIVE data path.

The mock run and scripts/verify_word_report.py exercise the full mock
``collected_data`` with the reference-aligned ``cost.per_service`` breakdown.
This script instead feeds ``WordReport`` a payload shaped like REAL collected
data WITHOUT the explicit ``per_service`` list (and without any legacy
``word_accounts``), so it drives the live fallbacks:

  * the Cost Summary Differences table derives its per-service rows from the
    live ``cost.service_breakdown`` (service / cost / difference), and still
    renders a Total Cost row from ``current_week_total`` / ``previous_week_total``;
  * the EC2 / RDS / ELB / WAF tables render from the live resource lists; and
  * the Amazon Inspector + Guard Duty summaries render from the live finding
    counts.

It then loads the generated .docx back with python-docx and asserts that all
the reference sections render. The rebuilt report is single-account scoped, so
the obsolete multi-account "Cost Summary Difference of All AWS Accounts" fleet
table must NOT appear.

Usage:
    python3 scripts/verify_word_live_fallback.py [--keep <out.docx>]

Runs on Python 3.9 (no PEP 604 unions). Requires python-docx.
"""

import argparse
import os
import sys
import tempfile
from datetime import date

# Make the repo root importable when run from anywhere.
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from docx import Document

from report.word_report import WordReport


def _build_fixture_config():
    """A config WITHOUT any word_report.accounts and with a custom
    disclaimer org, so the live path and the configurable disclaimer are both
    exercised."""
    return {
        "client_name": "JUST UDO AVIATION PRIVATE LIMITED (Fly91)",
        "aws_account_id": "674351849978",
        "thresholds": {"cpu_warning": 70, "cpu_critical": 90},
        "word_report": {
            "client_org": "JUST UDO AVIATION PRIVATE LIMITED (Fly91)",
            "submitter_org": "Greatworx",
            "activity_org": "Greatworx",
            "disclaimer_org": "Greatworx",
        },
    }


def _build_live_shaped_data():
    """A payload shaped like real collected data (no per_service, no
    word_accounts)."""
    return {
        "cost": {
            "current_week_total": 705.25,
            "previous_week_total": 728.33,
            "difference": -23.08,
            "pct_change": -3.17,
            "avg_daily_cost": 100.75,
            # Live-shaped service breakdown (service / cost / difference). The
            # Word report derives last-week = cost - difference per service.
            "service_breakdown": [
                {"service": "Elastic Load Balancing", "cost": 106.97, "difference": -13.65},
                {"service": "Relational Database Service", "cost": 73.68, "difference": -6.83},
                {"service": "CloudWatch", "cost": 77.64, "difference": -1.10},
                {"service": "WAF", "cost": 11.70, "difference": -0.96},
                {"service": "Inspector", "cost": 7.85, "difference": 0.00},
                {"service": "GuardDuty", "cost": 12.52, "difference": -0.12},
                {"service": "EC2-Other", "cost": 86.43, "difference": 0.24},
            ],
        },
        "ec2": [
            {
                "name": "FLY91- Production Website",
                "instance_id": "i-056b66d8b569a4a26",
                "instance_type": "r6a.xlarge",
                "state": "running",
                "region": "ap-south-1",
                "cpu_min": 2.79, "cpu_max": 77.11, "cpu_avg": 8.87,
                "memory_avg_display": "48.23", "disk_utilization": "49.0%",
                "network_in": 634 * 1024 * 1024, "network_out": 855 * 1024 * 1024,
                "network_packets_in": 446270, "network_packets_out": 169120,
            },
            {
                "name": "FLY91-Website-testing",
                "instance_id": "i-0d0653d04b9306f42",
                "instance_type": "r6a.large",
                "state": "stopped",
                "region": "ap-south-1",
                "cpu_min": None, "cpu_max": None, "cpu_avg": None,
                "disk_utilization": None,
                "network_in": None, "network_out": None,
                "network_packets_in": None, "network_packets_out": None,
            },
        ],
        "elb": [
            {
                "name": "Prod-Fly91-ALB",
                "requests": 25000, "active_connections": 4360,
                "new_connections": 8560, "consumed_lcus": 3.08,
                "http_redirect_count": 218,
                "processed_bytes": int(3.08 * 1024 * 1024 * 1024),
                "target_response_time": 0.3228,
            },
        ],
        "waf": [
            {
                "name": "Prod-WAF-Fly-91",
                "total_requests": 9150000, "blocked_requests": 88720,
                "allowed_requests": 9060000,
            },
        ],
        "rds": [
            {
                "name": "fly91-prod-psql-db",
                "down_time": "No", "instance_type": "db.t3.large",
                "min_utilization": "5.92%", "max_utilization": "100.0%",
                "avg_utilization": "43.76%", "free_memory": "4.35GB out of 8 GB",
                "free_storage": "9.0GB out of 100 GB",
                "network_transmit_throughput": "9.79 MB",
                "network_receive_throughput": "2.21 MB",
                "max_db_connections": 210,
            },
        ],
        "inspector": {
            "date_range": "2026-09-14 to 2026-09-20",
            "severity_counts": {"CRITICAL": 39, "HIGH": 200, "MEDIUM": 200,
                                "LOW": 15, "UNTRIAGED": 10},
            "total_findings": 464,
            "total_display": "400+",
            "severity_display": {"CRITICAL": "39", "HIGH": "200+",
                                 "MEDIUM": "200+", "LOW": "15",
                                 "UNTRIAGED": "10"},
        },
        "guardduty": {
            "date_range": "2026-09-14 to 2026-09-20",
            "severity_counts": {"HIGH": 0, "MEDIUM": 0, "LOW": 3},
            "total_findings": 66,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--keep", default=None,
                        help="Write the generated .docx to this path (otherwise a temp file is used and removed)")
    args = parser.parse_args()

    start_date = date(2026, 9, 14)
    end_date = date(2026, 9, 20)

    config = _build_fixture_config()
    collected_data = _build_live_shaped_data()

    gen = WordReport(config, start_date, end_date)

    out_path = args.keep
    tmp_handle = None
    if out_path is None:
        tmp_handle = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        tmp_handle.close()
        out_path = tmp_handle.name

    try:
        gen.generate(collected_data, out_path)
        doc = Document(out_path)

        all_para = "\n".join(p.text for p in doc.paragraphs)
        all_cell = "\n".join(
            cell.text
            for t in doc.tables for row in t.rows for cell in row.cells
        )
        full = all_para + "\n" + all_cell

        checks = [
            # Cost Summary Differences derived from live service_breakdown.
            ("Cost Summary Differences heading",
             any("Cost Summary Differences" in p.text for p in doc.paragraphs)),
            ("live cost table has Inspector row", "Inspector" in all_cell),
            ("live cost table has GuardDuty row", "GuardDuty" in all_cell),
            ("live cost table has Total Cost row", "Total Cost" in all_cell),
            ("account id present", "674351849978" in all_cell),

            # Disclaimer with configured org.
            ("Disclaimer heading",
             any(p.text.strip() == "Disclaimer" for p in doc.paragraphs)),
            ("configurable disclaimer org (Greatworx)",
             "confidential between Greatworx and Customer" in all_para),

            # BAU Matrix Overview + resource tables.
            ("BAU Matrix Overview heading",
             "BAU Matrix Overview of the resources" in all_para),
            ("EC2 table row (production website)",
             "i-056b66d8b569a4a26" in all_cell),
            ("RDS table row (prod psql)", "fly91-prod-psql-db" in all_cell),
            ("ELB table row (prod ALB)", "Prod-Fly91-ALB" in all_cell),
            ("WAF table row (prod WAF)", "Prod-WAF-Fly-91" in all_cell),

            # Inspector + GuardDuty summaries.
            ("Inspector findings summary",
             "findings for the last week" in all_para),
            ("GuardDuty findings summary",
             "new findings this week" in all_para),

            # Trailer.
            ("End of Document trailer", "-- End of Document --" in all_para),

            # NEGATIVE: obsolete layout must be gone.
            ("no multi-account fleet table",
             "Cost Summary Difference of All AWS Accounts" not in full),
            ("no Resource Utilization & Alarms",
             "Resource Utilization & Alarms" not in all_para),
        ]

        failed = [name for name, ok in checks if not ok]
        for name, ok in checks:
            print(("PASS" if ok else "FAIL"), name)

        if failed:
            print("\nFAILED CHECKS: %d" % len(failed), file=sys.stderr)
            return 1

        print("\nALL CHECKS PASSED  (live data path) tables=%d" % len(doc.tables))
        return 0
    finally:
        if tmp_handle is not None:
            try:
                os.unlink(out_path)
            except OSError:
                pass


if __name__ == "__main__":
    sys.exit(main())
