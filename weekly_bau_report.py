#!/usr/bin/env python3
"""
Weekly AWS BAU Report Generator
================================

Production-ready Python application that collects AWS infrastructure metrics
and generates a professional Weekly BAU Matrix Report in Excel format.

Client: Insync Analytics
AWS Account: 179787470151
Phase: 1 (Single Account)

Usage:
    python weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22

Requirements:
    - Python 3.8+
    - boto3 (pre-installed on AWS CloudShell)
    - openpyxl (install via: pip3 install openpyxl --user)

Author: Cloud Operations Team
"""

import argparse
import json
import math
import os
import sys
import time
from datetime import datetime, date, timedelta

try:
    import boto3
except ImportError:
    print("ERROR: boto3 is required. Install via: pip3 install boto3")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl is required. Install via: pip3 install openpyxl --user")
    sys.exit(1)

from utils.logger import setup_logger
from utils.helpers import (
    parse_date,
    validate_date_range,
    format_date_range,
    format_currency,
    format_percentage,
)
from collectors.cost_collector import CostCollector
from collectors.ec2_collector import EC2Collector
from collectors.elb_collector import ELBCollector
from collectors.waf_collector import WAFCollector
from collectors.s3_collector import S3Collector
from collectors.rds_collector import RDSCollector
from collectors.autoscaling_collector import AutoScalingCollector
from collectors.inspector_collector import InspectorCollector
from collectors.guardduty_collector import GuardDutyCollector
from report.excel_generator import ExcelGenerator
from report.word_report import WordReport, WordReportError


# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

DEFAULT_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def load_config(config_path: str) -> dict:
    """Load and validate the configuration file."""
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # Validate required fields
    required_fields = ['client_name', 'aws_account_id', 'regions']
    for field in required_fields:
        if field not in config:
            raise ValueError(f"Missing required config field: '{field}'")

    return config


# ──────────────────────────────────────────────────────────────────────────────
# CLI Argument Parsing
# ──────────────────────────────────────────────────────────────────────────────

def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Weekly AWS BAU Report Generator — Collects AWS metrics and "
                    "generates a professional Excel report.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22
  python weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22 --output my_report.xlsx
  python weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22 --config /path/to/config.json
        """
    )

    parser.add_argument(
        '--start-date',
        type=str,
        required=True,
        help='Report period start date (YYYY-MM-DD format)'
    )
    parser.add_argument(
        '--end-date',
        type=str,
        required=True,
        help='Report period end date (YYYY-MM-DD format)'
    )
    parser.add_argument(
        '--config',
        type=str,
        default=DEFAULT_CONFIG_PATH,
        help=f'Path to configuration file (default: {DEFAULT_CONFIG_PATH})'
    )
    parser.add_argument(
        '--output',
        type=str,
        default=None,
        help='Output file path (overrides config setting). If it ends in .docx it '
             'is used as the Word output path; otherwise it is the Excel output path.'
    )
    parser.add_argument(
        '--format',
        dest='output_format',
        type=str,
        choices=['docx', 'xlsx', 'both'],
        default='xlsx',
        help="Which report(s) to generate: 'xlsx' (default, Excel only), "
             "'docx' (Word only), or 'both'."
    )
    parser.add_argument(
        '--mock',
        action='store_true',
        help='Generate the report using mock data matching the exact report spec without querying AWS'
    )

    return parser.parse_args()


# ──────────────────────────────────────────────────────────────────────────────
# Data Collection Orchestrator
# ──────────────────────────────────────────────────────────────────────────────

def collect_all_data(session, config, start_date, end_date, logger) -> dict:
    """
    Run all data collectors and aggregate results.

    Args:
        session: boto3 Session
        config: Application configuration dict
        start_date: Report period start date
        end_date: Report period end date
        logger: Logger instance

    Returns:
        Dictionary containing all collected data
    """
    collected_data = {}
    regions = config.get('regions', ['us-east-1'])
    primary_region = regions[0]

    collectors = [
        ('cost', CostCollector, "Cost Explorer"),
        ('ec2', EC2Collector, "EC2 Instances & Metrics"),
        ('elb', ELBCollector, "Elastic Load Balancers"),
        ('waf', WAFCollector, "AWS WAF"),
        ('inspector', InspectorCollector, "Amazon Inspector Findings"),
        ('guardduty', GuardDutyCollector, "Amazon GuardDuty Findings"),
        ('s3', S3Collector, "S3 Buckets"),
        ('rds', RDSCollector, "RDS Instances"),
        ('autoscaling', AutoScalingCollector, "Auto Scaling Groups"),
    ]

    total = len(collectors)
    for idx, (key, collector_cls, description) in enumerate(collectors, 1):
        logger.info(f"[{idx}/{total}] Collecting {description}...")
        start_time = time.time()

        try:
            collector = collector_cls(
                session=session,
                region=primary_region,
                config=config,
                start_date=start_date,
                end_date=end_date,
                logger=logger
            )
            collected_data[key] = collector.collect()
            elapsed = time.time() - start_time
            logger.info(f"[{idx}/{total}] ✓ {description} collected in {elapsed:.1f}s")

        except Exception as e:
            elapsed = time.time() - start_time
            logger.error(f"[{idx}/{total}] ✗ {description} failed after {elapsed:.1f}s: {e}")
            collected_data[key] = None

    return collected_data


# ──────────────────────────────────────────────────────────────────────────────
# Report Summary (Console Output)
# ──────────────────────────────────────────────────────────────────────────────

def print_summary(collected_data: dict, start_date: date, end_date: date, logger):
    """Print a brief collection summary to the console."""
    logger.info("")
    logger.info("=" * 70)
    logger.info("  COLLECTION SUMMARY")
    logger.info("=" * 70)
    logger.info(f"  Report Period: {format_date_range(start_date, end_date)}")

    # Cost summary
    cost_data = collected_data.get('cost')
    if cost_data:
        logger.info(f"  Current Week Cost: {format_currency(cost_data.get('current_week_total'))}")
        logger.info(f"  Previous Week Cost: {format_currency(cost_data.get('previous_week_total'))}")
        pct = cost_data.get('pct_change')
        if pct is not None:
            direction = "↑" if pct > 0 else "↓" if pct < 0 else "→"
            logger.info(f"  Cost Change: {direction} {format_percentage(abs(pct))}")

    # EC2 summary
    ec2_data = collected_data.get('ec2')
    if ec2_data and isinstance(ec2_data, list):
        running = sum(1 for i in ec2_data if i.get('state') == 'running')
        stopped = sum(1 for i in ec2_data if i.get('state') == 'stopped')
        total = len(ec2_data)
        logger.info(f"  EC2 Instances: {total} total ({running} running, {stopped} stopped)")

    # ELB summary
    elb_data = collected_data.get('elb')
    if elb_data and isinstance(elb_data, list):
        logger.info(f"  Load Balancers: {len(elb_data)}")

    # WAF summary
    waf_data = collected_data.get('waf')
    if waf_data and isinstance(waf_data, list):
        logger.info(f"  WAF Web ACLs: {len(waf_data)}")

    # S3 summary
    s3_data = collected_data.get('s3')
    if s3_data and isinstance(s3_data, list):
        logger.info(f"  S3 Buckets: {len(s3_data)}")

    # RDS summary
    rds_data = collected_data.get('rds')
    if rds_data and isinstance(rds_data, list):
        logger.info(f"  RDS Instances: {len(rds_data)}" if rds_data else
                     "  RDS: No instances available")
    else:
        logger.info("  RDS: No instances available")

    logger.info("=" * 70)


# ──────────────────────────────────────────────────────────────────────────────
# Mock Data Generator
# ──────────────────────────────────────────────────────────────────────────────

# The reference per-service cost breakdown (Last Week / Current Week), taken
# verbatim from REFERENCE_FLY91_GOOGLEDOC.txt. Each tuple is
# (service, last_week, current_week). The reference-stated Tax-Excluded totals
# are $728.33 (last week) -> $705.25 (current week); per-row figures sum to the
# same within a rounding penny, so the explicit totals below are authoritative.
_REFERENCE_COST_ROWS = [
    ("Relational Database Service", 80.51, 73.68),
    ("Fortinet FortiGate VM Next-Generation Firewall", 171.36, 171.36),
    ("Savings Plans for Compute Usage", 84.00, 84.00),
    ("EC2-Other", 86.19, 86.43),
    ("Elastic Load Balancing", 120.62, 106.97),
    ("CloudWatch", 78.74, 77.64),
    ("VPC", 22.36, 21.93),
    ("EC2-Instances", 20.66, 20.01),
    ("Config", 27.98, 28.11),
    ("WAF", 12.66, 11.70),
    ("Inspector", 7.85, 7.85),
    ("GuardDuty", 12.64, 12.52),
    ("Route 53", 0.34, 0.35),
    ("CloudFront", 0.10, 0.24),
    ("Key Management Service", 0.30, 0.30),
    ("S3", 1.23, 1.21),
    ("Lambda", 0.59, 0.72),
    ("Secrets Manager", 0.10, 0.10),
    ("CloudTrail", 0.08, 0.08),
    ("CodePipeline", 0.00, 0.00),
    ("SNS", 0.03, 0.06),
    ("CloudWatch Events", 0.00, 0.00),
    ("SQS", 0.00, 0.00),
    ("CloudShell", 0.00, 0.00),
    ("Glue", 0.00, 0.00),
    ("DynamoDB", 0.00, 0.00),
    ("Tax", 0.00, None),
]

# Reference-stated Tax-Excluded weekly totals.
_REFERENCE_TOTAL_LAST_WEEK = 728.33
_REFERENCE_TOTAL_CURRENT_WEEK = 705.25


def _build_mock_cost(start_date: date, end_date: date) -> dict:
    """Build the mock ``cost`` dict matching the Fly91 reference document.

    Carries BOTH:
      * the per-service breakdown the Word report's "Cost Summary Differences"
        table and analysis bullets consume (``per_service`` list of
        {service, last_week, current_week}, plus explicit weekly totals), and
      * the fields the Excel Cost Analysis sheet already reads
        (current_week_total, previous_week_total, difference, pct_change,
        avg_daily_cost, daily_costs, top_services, service_breakdown).
    """
    per_service = [
        {"service": svc, "last_week": last, "current_week": curr}
        for (svc, last, curr) in _REFERENCE_COST_ROWS
    ]

    total_last = _REFERENCE_TOTAL_LAST_WEEK
    total_curr = _REFERENCE_TOTAL_CURRENT_WEEK
    difference = round(total_curr - total_last, 2)
    pct_change = round((difference / total_last) * 100.0, 2) if total_last else 0.0
    days = max((end_date - start_date).days + 1, 1)
    avg_daily = round(total_curr / days, 2)

    # Build daily cost entries that sum close to the current-week total so the
    # Excel daily-trend view stays coherent with the weekly total.
    daily_costs = []
    base = round(total_curr / days, 2)
    curr_date = start_date
    accumulated = 0.0
    i = 0
    n_days = days
    while curr_date <= end_date:
        if i == n_days - 1:
            day_cost = round(total_curr - accumulated, 2)
        else:
            # Small deterministic ripple around the daily average.
            ripple = [-3.1, 2.4, -1.8, 4.0, -2.2, 1.5, 0.0][i % 7]
            day_cost = round(base + ripple, 2)
            accumulated = round(accumulated + day_cost, 2)
        daily_costs.append({
            "date": curr_date.strftime("%Y-%m-%d"),
            "cost": day_cost,
        })
        curr_date += timedelta(days=1)
        i += 1

    # top_services / service_breakdown feed the Excel Cost Analysis sheet. Derive
    # them from the per-service current-week costs (skip the Tax row), largest
    # first, each with its week-over-week pct change for the change indicator.
    svc_entries = []
    for svc, last, curr in _REFERENCE_COST_ROWS:
        if svc == "Tax" or curr is None:
            continue
        if last:
            svc_pct = round(((curr - last) / last) * 100.0, 1)
        else:
            svc_pct = 0.0
        svc_entries.append({
            "service": svc,
            "cost": curr,
            "pct_change": svc_pct,
            "difference": round(curr - last, 2),
        })
    svc_entries.sort(key=lambda s: s["cost"], reverse=True)
    top_services = [{"service": s["service"], "cost": s["cost"]} for s in svc_entries]

    return {
        "current_week_total": total_curr,
        "previous_week_total": total_last,
        "difference": difference,
        "pct_change": pct_change,
        "avg_daily_cost": avg_daily,
        "daily_costs": daily_costs,
        "top_services": top_services,
        "service_breakdown": svc_entries,
        # Reference-aligned per-service breakdown + totals for the Word report.
        "per_service": per_service,
        "total_last_week": total_last,
        "total_current_week": total_curr,
        "total_last_week_label": "$728.33 Tax Excluded Cost",
        "total_current_week_label": "$705.25 Tax Excluded Cost",
    }


def get_mock_data(start_date: date, end_date: date) -> dict:
    """Generate realistic mock data matching the exact report spec."""
    return {
        "cost": _build_mock_cost(start_date, end_date),
        # The 9 reference EC2 rows (8 running + 1 stopped), verbatim from
        # REFERENCE_FLY91_GOOGLEDOC.txt. Byte/count values are chosen so that
        # utils.helpers.format_bytes / format_count render the reference's
        # human-readable figures (e.g. 1.20 GB, 952.35K). Memory and disk are
        # carried as display strings where the reference shows a formatted
        # value (e.g. "50.62", "91.6%", "C: 51.3%"); "-" where the reference
        # shows no data.
        "ec2": [
            {
                "name": "Fortinet-FW",
                "instance_id": "i-06a53ce120fca7a28",
                "instance_type": "c6i.xlarge",
                "state": "running",
                "availability_zone": "ap-south-1a",
                "region": "ap-south-1",
                "cpu_min": 0.59,
                "cpu_max": 6.19,
                "cpu_avg": 0.83,
                "memory_avg_display": "-",
                "disk_utilization": "-",
                "network_in": int(1.20 * 1024 * 1024 * 1024),
                "network_out": int(1.26 * 1024 * 1024 * 1024),
                "network_packets_in": 952350,
                "network_packets_out": 955080,
            },
            {
                "name": "FLY91- Staging Website",
                "instance_id": "i-041cbf2e1ea9a05ee",
                "instance_type": "t3a.xlarge",
                "state": "running",
                "availability_zone": "ap-south-1a",
                "region": "ap-south-1",
                "cpu_min": 4.5,
                "cpu_max": 35.46,
                "cpu_avg": 9.19,
                "memory_avg_display": "50.62",
                "disk_utilization": "91.6%",
                "network_in": int(1.31 * 1024 * 1024 * 1024),
                "network_out": int(1.56 * 1024 * 1024 * 1024),
                "network_packets_in": 1030000,
                "network_packets_out": 1100000,
            },
            {
                "name": "FLY91- Production Website",
                "instance_id": "i-056b66d8b569a4a26",
                "instance_type": "r6a.xlarge",
                "state": "running",
                "availability_zone": "ap-south-1b",
                "region": "ap-south-1",
                "cpu_min": 2.79,
                "cpu_max": 77.11,
                "cpu_avg": 8.87,
                "memory_avg_display": "48.23",
                "disk_utilization": "49.0%",
                "network_in": int(634.38 * 1024 * 1024),
                "network_out": int(855.01 * 1024 * 1024),
                "network_packets_in": 446270,
                "network_packets_out": 169120,
            },
            {
                "name": "Fly91-SFTP",
                "instance_id": "i-0b108151d07fe2bf8",
                "instance_type": "t3a.medium",
                "state": "running",
                "availability_zone": "ap-south-1b",
                "region": "ap-south-1",
                "cpu_min": 2.65,
                "cpu_max": 49.97,
                "cpu_avg": 3.89,
                "memory_avg_display": "25.35",
                "disk_utilization": "76.3%",
                "network_in": int(88.54 * 1024 * 1024),
                "network_out": int(14.32 * 1024 * 1024),
                "network_packets_in": 58770,
                "network_packets_out": 17500,
            },
            {
                "name": "Fly91-Desk",
                "instance_id": "i-0e24da06be3f86a40",
                "instance_type": "t3a.medium",
                "state": "running",
                "availability_zone": "ap-south-1c",
                "region": "ap-south-1",
                "cpu_min": 1.25,
                "cpu_max": 28.72,
                "cpu_avg": 1.8,
                "memory_avg_display": "73.97",
                "disk_utilization": "21.8%",
                "network_in": int(88.29 * 1024 * 1024),
                "network_out": int(424.86 * 1024),
                "network_packets_in": 58600,
                "network_packets_out": 3640,
            },
            {
                "name": "FLY91-Website-testing",
                "instance_id": "i-0d0653d04b9306f42",
                "instance_type": "r6a.large",
                "state": "stopped",
                "stopped_at_display": "30/10/2025 3:25 PM",
                "availability_zone": "ap-south-1c",
                "region": "ap-south-1",
                "cpu_min": None,
                "cpu_max": None,
                "cpu_avg": None,
                "memory_avg_display": "-",
                "disk_utilization": "-",
                "network_in": None,
                "network_out": None,
                "network_packets_in": None,
                "network_packets_out": None,
            },
            {
                "name": "Fly91-Development-Server-Windows",
                "instance_id": "i-00a0097da058e39a1",
                "instance_type": "t3a.xlarge",
                "state": "running",
                "availability_zone": "ap-south-1a",
                "region": "ap-south-1",
                "cpu_min": 9.93,
                "cpu_max": 58.83,
                "cpu_avg": 14.3,
                "memory_avg_display": "55.21",
                "disk_utilization": "C: 51.3%",
                "network_in": int(63.74 * 1024 * 1024),
                "network_out": int(10.83 * 1024 * 1024),
                "network_packets_in": 42440,
                "network_packets_out": 16410,
            },
            {
                "name": "Fly91-Development-Server-Ubuntu",
                "instance_id": "i-0233bf5f2a51870df",
                "instance_type": "t3a.xlarge",
                "state": "running",
                "availability_zone": "ap-south-1b",
                "region": "ap-south-1",
                "cpu_min": 2.49,
                "cpu_max": 42.45,
                "cpu_avg": 5.77,
                "memory_avg_display": "48.75",
                "disk_utilization": "87.7%",
                "network_in": int(186.64 * 1024 * 1024),
                "network_out": int(141.11 * 1024 * 1024),
                "network_packets_in": 136220,
                "network_packets_out": 97710,
            },
            {
                "name": "FLY91 - Data Science Production",
                "instance_id": "i-08b66727cdf8d5896",
                "instance_type": "t3a.large",
                "state": "running",
                "availability_zone": "ap-south-1c",
                "region": "ap-south-1",
                "cpu_min": 0.89,
                "cpu_max": 32.93,
                "cpu_avg": 5.15,
                "memory_avg_display": "9.22",
                "disk_utilization": "14.2%",
                "network_in": int(109.23 * 1024 * 1024),
                "network_out": int(48.32 * 1024 * 1024),
                "network_packets_in": 75590,
                "network_packets_out": 33800,
            },
        ],
        # The 3 reference ALBs (Sum-utilisation), verbatim from the reference.
        # Byte/count helpers render processed_bytes/requests/connections to the
        # reference's human-readable figures; target_response_time is seconds
        # (format_duration renders 1.109 s / 322.8 ms / 274.7 ms).
        "elb": [
            {
                "name": "Stage-Fly91-ALB",
                "requests": 1120,
                "active_connections": 2240,
                "new_connections": 1600,
                "consumed_lcus": 2.41,
                "http_redirect_count": 30,
                "processed_bytes": int(2.41 * 1024 * 1024 * 1024),
                "target_response_time": 1.109,
            },
            {
                "name": "Prod-Fly91-ALB",
                "requests": 25000,
                "active_connections": 4360,
                "new_connections": 8560,
                "consumed_lcus": 3.08,
                "http_redirect_count": 218,
                "processed_bytes": int(3.08 * 1024 * 1024 * 1024),
                "target_response_time": 0.3228,
            },
            {
                "name": "Fly91-Development-ALB",
                "requests": 2340,
                "active_connections": 843,
                "new_connections": 3370,
                "consumed_lcus": 0.04,
                "http_redirect_count": 2090,
                "processed_bytes": int(17.70 * 1024 * 1024),
                "target_response_time": 0.2747,
            },
        ],
        # The 2 reference WAF web ACLs (Mumbai region), verbatim.
        "waf": [
            {
                "name": "Prod-WAF-Fly-91",
                "total_requests": 9150000,
                "blocked_requests": 88720,
                "allowed_requests": 9060000,
            },
            {
                "name": "Stag-WAF-Fly-91",
                "total_requests": 246450,
                "blocked_requests": 5040,
                "allowed_requests": 241410,
            },
        ],
        "s3": [
            {"name": "databricks-workspace-stack-5fd33-bucket", "region": "us-east-1", "total_size_bytes": 5.9 * 1024 * 1024, "total_objects": 79},
            {"name": "databricks-workspace-stack-a07b6-bucket", "region": "us-east-1", "total_size_bytes": 33.0 * 1024 * 1024, "total_objects": 227},
            {"name": "frontend-istari.insyncanalytics.com", "region": "us-east-1", "total_size_bytes": 898.4 * 1024 * 1024, "total_objects": 12500},
            {"name": "indigo-insync-prebuilt", "region": "us-east-1", "total_size_bytes": 183.0 * 1024 * 1024 * 1024, "total_objects": 506200},
            {"name": "insync-prod-analytics-trail", "region": "us-east-1", "total_size_bytes": int(3.2 * 1024 * 1024 * 1024), "total_objects": 1300000},
            {"name": "insync5aws", "region": "us-east-1", "total_size_bytes": int(9.7 * 1024 * 1024 * 1024), "total_objects": 608},
            {"name": "insyncdatabricks", "region": "us-east-1", "total_size_bytes": 2.0 * 1024 * 1024 * 1024, "total_objects": 414},
            {"name": "istari-file-syncing", "region": "us-east-1", "total_size_bytes": int(259.8 * 1024 * 1024 * 1024), "total_objects": 135400},
            {"name": "istari-secure-folder", "region": "us-east-1", "total_size_bytes": int(228.1 * 1024 * 1024 * 1024), "total_objects": 134000},
            {"name": "weave-access-logs-lb", "region": "us-east-1", "total_size_bytes": 35.4 * 1024 * 1024, "total_objects": 20200},
            {"name": "insync4aws", "region": "us-east-1", "total_size_bytes": 0.0, "total_objects": 0},
            {"name": "insync6aws", "region": "us-east-1", "total_size_bytes": 15.0, "total_objects": 1},
            {"name": "istari-patching-logs", "region": "us-east-1", "total_size_bytes": 0.0, "total_objects": 0},
            {"name": "textractprogrammerbucket", "region": "us-east-1", "total_size_bytes": 0.0, "total_objects": 0},
        ],
        # The 6 reference RDS instances, verbatim. Utilisation / free-memory /
        # free-storage / throughput values are carried as display strings
        # because the reference shows pre-formatted values (e.g. "0.5 GB out of
        # 2GB", "817.54 KB"). down_time "No" matches the reference.
        "rds": [
            {
                "name": "fly91-data-science-production-db",
                "down_time": "No",
                "instance_type": "db.t4g.small",
                "min_utilization": "3.78%",
                "max_utilization": "30.61%",
                "avg_utilization": "8.74%",
                "free_memory": "0.5 GB out of 2GB",
                "free_storage": "15.8GB out of 20 GB",
                "network_transmit_throughput": "817.54 KB",
                "network_receive_throughput": "77.25 KB",
                "max_db_connections": 5,
            },
            {
                "name": "fly91-db",
                "down_time": "No",
                "instance_type": "db.r7g.large",
                "min_utilization": "2.9%",
                "max_utilization": "47.49%",
                "avg_utilization": "3.72%",
                "free_memory": "1.86GB out of 16 GB",
                "free_storage": "89.1GB out of 100 GB",
                "network_transmit_throughput": "24.70 MB",
                "network_receive_throughput": "9.52 MB",
                "max_db_connections": 12,
            },
            {
                "name": "fly91-dev-db",
                "down_time": "No",
                "instance_type": "db.t3.large",
                "min_utilization": "14.41%",
                "max_utilization": "100.0%",
                "avg_utilization": "46.17%",
                "free_memory": "0.86GB out of 8 GB",
                "free_storage": "21.4GB out of 30 GB",
                "network_transmit_throughput": "15.78 MB",
                "network_receive_throughput": "142.12 KB",
                "max_db_connections": 68,
            },
            {
                "name": "fly91-development-psql",
                "down_time": "No",
                "instance_type": "db.t3.medium",
                "min_utilization": "4.44%",
                "max_utilization": "79.77%",
                "avg_utilization": "13.32%",
                "free_memory": "2.90GB out of 4 GB",
                "free_storage": "0.8GB out of 77 GB",
                "network_transmit_throughput": "15.58 MB",
                "network_receive_throughput": "2.50 MB",
                "max_db_connections": 134,
            },
            {
                "name": "fly91-prod-psql-db",
                "down_time": "No",
                "instance_type": "db.t3.large",
                "min_utilization": "5.92%",
                "max_utilization": "100.0%",
                "avg_utilization": "43.76%",
                "free_memory": "4.35GB out of 8 GB",
                "free_storage": "9.0GB out of 100 GB",
                "network_transmit_throughput": "9.79 MB",
                "network_receive_throughput": "2.21 MB",
                "max_db_connections": 210,
            },
            {
                "name": "fly91-uat-psql-db",
                "down_time": "No",
                "instance_type": "db.t3.medium",
                "min_utilization": "5.65%",
                "max_utilization": "99.57%",
                "avg_utilization": "39.4%",
                "free_memory": "2.00GB out of 4 GB",
                "free_storage": "27.1GB out of 50 GB",
                "network_transmit_throughput": "18.92 MB",
                "network_receive_throughput": "772.02 KB",
                "max_db_connections": 193,
            },
        ],
        "inspector": {
            "date_range": f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
            # Reference: 400+ findings categorised as 39 Critical, 200+ High,
            # 200+ medium, 15 low and 10 untriaged.
            "severity_counts": {
                "CRITICAL": 39,
                "HIGH": 200,
                "MEDIUM": 200,
                "LOW": 15,
                "UNTRIAGED": 10,
            },
            "total_findings": 464,
            # Reference-shaped display strings for the Word summary sentence.
            # The numeric severity_counts above stay intact for the Excel
            # Inspector sheet; these only drive the Word wording so it reads
            # exactly like the reference ("400+ ... 200+ High, 200+ medium").
            "total_display": "400+",
            "severity_display": {
                "CRITICAL": "39",
                "HIGH": "200+",
                "MEDIUM": "200+",
                "LOW": "15",
                "UNTRIAGED": "10",
            },
            "collection_status": "ok",
            "collection_error": None,
            "findings": [
                {"title": "CVE-2024-3094 - xz-utils backdoor (liblzma)", "severity": "CRITICAL", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Nexus Server (i-09a1b5ff52381b04b)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2024-3094", "first_observed": start_date.strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2021-44228 - Apache Log4j2 RCE (Log4Shell)", "severity": "CRITICAL", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "ISTARI-DB Server (i-0bc7cb7733cccd23b)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2021-44228", "first_observed": start_date.strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2024-6387 - OpenSSH regreSSHion RCE", "severity": "HIGH", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync ClickHouse-Database (i-029867634022c4d78)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2024-6387", "first_observed": (start_date + timedelta(days=1)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2023-4911 - glibc Looney Tunables privilege escalation", "severity": "HIGH", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Nexus Server (i-09a1b5ff52381b04b)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2023-4911", "first_observed": (start_date + timedelta(days=1)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2022-0778 - OpenSSL infinite loop DoS", "severity": "HIGH", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Website Server (i-0a8f07365bce735f6)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2022-0778", "first_observed": (start_date + timedelta(days=2)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2023-38545 - curl SOCKS5 heap buffer overflow", "severity": "HIGH", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Weave-Server (i-05966c44f5be655cc)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2023-38545", "first_observed": (start_date + timedelta(days=2)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2024-1086 - Linux kernel nf_tables use-after-free", "severity": "HIGH", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync-Pipelines Server (i-08b99bf46815a7d38)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2024-1086", "first_observed": (start_date + timedelta(days=3)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2023-44487 - HTTP/2 Rapid Reset DoS", "severity": "HIGH", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Add-in Server (i-07c0fb23505c71dd8)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2023-44487", "first_observed": (start_date + timedelta(days=3)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2023-29491 - ncurses local privilege escalation", "severity": "MEDIUM", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "ISTARI-DB Server (i-0bc7cb7733cccd23b)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2023-29491", "first_observed": (start_date + timedelta(days=1)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2022-37434 - zlib heap buffer over-read", "severity": "MEDIUM", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync ClickHouse-Database (i-029867634022c4d78)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2022-37434", "first_observed": (start_date + timedelta(days=2)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2023-0286 - OpenSSL X.400 type confusion", "severity": "MEDIUM", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Nexus Server (i-09a1b5ff52381b04b)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2023-0286", "first_observed": (start_date + timedelta(days=2)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2023-27536 - curl GSSAPI credential reuse", "severity": "MEDIUM", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Website Server (i-0a8f07365bce735f6)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2023-27536", "first_observed": (start_date + timedelta(days=3)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2022-40674 - expat use-after-free", "severity": "MEDIUM", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Weave-Server (i-05966c44f5be655cc)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2022-40674", "first_observed": (start_date + timedelta(days=4)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2021-3711 - OpenSSL SM2 decryption buffer overflow", "severity": "MEDIUM", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync-Pipelines Server (i-08b99bf46815a7d38)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2021-3711", "first_observed": (start_date + timedelta(days=4)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2022-3602 - OpenSSL X.509 punycode buffer overflow", "severity": "MEDIUM", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "ISTARI-App-Server-ASG (i-06559919372821d45)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2022-3602", "first_observed": (start_date + timedelta(days=5)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2023-2650 - OpenSSL OBJECT IDENTIFIER DoS", "severity": "LOW", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Add-in Server (i-07c0fb23505c71dd8)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2023-2650", "first_observed": (start_date + timedelta(days=5)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2022-1292 - OpenSSL c_rehash command injection", "severity": "LOW", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync ClickHouse-Database (i-029867634022c4d78)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2022-1292", "first_observed": (start_date + timedelta(days=5)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2021-33560 - libgcrypt ElGamal side-channel", "severity": "LOW", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "ISTARI-DB Server (i-0bc7cb7733cccd23b)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2021-33560", "first_observed": (start_date + timedelta(days=6)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "CVE-2020-1971 - OpenSSL EDIPARTYNAME NULL deref", "severity": "LOW", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Nexus Server (i-09a1b5ff52381b04b)", "finding_type": "PACKAGE_VULNERABILITY", "cve": "CVE-2020-1971", "first_observed": (start_date + timedelta(days=6)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
                {"title": "EC2 instance has an outdated SSM Agent", "severity": "INFORMATIONAL", "resource_type": "AWS_EC2_INSTANCE", "resource_id": "InSync Website Server (i-0a8f07365bce735f6)", "finding_type": "SOFTWARE_PACKAGE", "cve": "-", "first_observed": (start_date + timedelta(days=6)).strftime('%Y-%m-%d'), "status": "ACTIVE"},
            ],
        },
        "guardduty": {
            "date_range": f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
            # Reference: 66 new findings this week, categorised as 3 lows,
            # 0 Medium and High.
            "severity_counts": {"HIGH": 0, "MEDIUM": 0, "LOW": 3},
            "total_findings": 66,
            "collection_status": "ok",
            "collection_error": None,
            "findings": [
                {"title": "Bitcoin-related domain name queried by i-09a1b5ff52381b04b", "type": "CryptoCurrency:EC2/BitcoinTool.B!DNS", "severity_label": "HIGH", "severity_score": 8.0, "resource_type": "Instance", "region": "us-east-1", "count": 12, "first_seen": start_date.strftime('%Y-%m-%d'), "last_seen": (start_date + timedelta(days=2)).strftime('%Y-%m-%d')},
                {"title": "EC2 instance i-0bc7cb7733cccd23b is the target of SSH brute force attacks", "type": "UnauthorizedAccess:EC2/SSHBruteForce", "severity_label": "HIGH", "severity_score": 7.5, "resource_type": "Instance", "region": "us-east-1", "count": 48, "first_seen": start_date.strftime('%Y-%m-%d'), "last_seen": (start_date + timedelta(days=5)).strftime('%Y-%m-%d')},
                {"title": "EC2 instance i-029867634022c4d78 is communicating with a known malware command and control server", "type": "Backdoor:EC2/C&CActivity.B!DNS", "severity_label": "HIGH", "severity_score": 7.2, "resource_type": "Instance", "region": "us-east-1", "count": 6, "first_seen": (start_date + timedelta(days=1)).strftime('%Y-%m-%d'), "last_seen": (start_date + timedelta(days=3)).strftime('%Y-%m-%d')},
                {"title": "Unprotected port on EC2 instance i-07c0fb23505c71dd8 is being probed", "type": "Recon:EC2/PortProbeUnprotectedPort", "severity_label": "MEDIUM", "severity_score": 5.0, "resource_type": "Instance", "region": "us-east-1", "count": 23, "first_seen": (start_date + timedelta(days=1)).strftime('%Y-%m-%d'), "last_seen": (start_date + timedelta(days=4)).strftime('%Y-%m-%d')},
                {"title": "API GeneratedFindingAPIName was invoked from a Tor exit node", "type": "UnauthorizedAccess:IAMUser/TorIPCaller", "severity_label": "MEDIUM", "severity_score": 5.5, "resource_type": "AccessKey", "region": "us-east-1", "count": 4, "first_seen": (start_date + timedelta(days=2)).strftime('%Y-%m-%d'), "last_seen": (start_date + timedelta(days=4)).strftime('%Y-%m-%d')},
                {"title": "EC2 instance i-05966c44f5be655cc is performing outbound port scans", "type": "Recon:EC2/Portscan", "severity_label": "MEDIUM", "severity_score": 4.5, "resource_type": "Instance", "region": "us-east-1", "count": 15, "first_seen": (start_date + timedelta(days=3)).strftime('%Y-%m-%d'), "last_seen": (start_date + timedelta(days=5)).strftime('%Y-%m-%d')},
                {"title": "Login to the console from an unusual geolocation", "type": "UnauthorizedAccess:IAMUser/ConsoleLoginSuccess.B", "severity_label": "LOW", "severity_score": 3.0, "resource_type": "AccessKey", "region": "us-east-1", "count": 2, "first_seen": (start_date + timedelta(days=4)).strftime('%Y-%m-%d'), "last_seen": (start_date + timedelta(days=5)).strftime('%Y-%m-%d')},
                {"title": "EC2 instance i-08b99bf46815a7d38 is querying a low-reputation domain", "type": "Recon:EC2/PortProbeEMRUnprotectedPort", "severity_label": "LOW", "severity_score": 2.0, "resource_type": "Instance", "region": "us-east-1", "count": 9, "first_seen": (start_date + timedelta(days=5)).strftime('%Y-%m-%d'), "last_seen": (start_date + timedelta(days=6)).strftime('%Y-%m-%d')},
            ],
        },
    }


# ──────────────────────────────────────────────────────────────────────────────
# Main Entry Point
# ──────────────────────────────────────────────────────────────────────────────

def main():
    """Main execution flow."""
    print()
    print("╔══════════════════════════════════════════════════════════════════╗")
    print("║       WEEKLY AWS INFRASTRUCTURE BAU MATRIX REPORT              ║")
    print("║       Phase 1 — Single Account Report Generator                ║")
    print("╚══════════════════════════════════════════════════════════════════╝")
    print()

    # ── Parse arguments ──
    args = parse_arguments()

    # ── Load configuration ──
    try:
        config = load_config(args.config)
    except (FileNotFoundError, ValueError, json.JSONDecodeError) as e:
        print(f"ERROR: Configuration error: {e}")
        sys.exit(1)

    # ── Setup logging ──
    log_config = config.get('logging', {})
    logger = setup_logger(
        name="bau_report",
        log_level=log_config.get('level', 'INFO'),
        log_file=log_config.get('file')
    )

    logger.info(f"Client: {config['client_name']}")
    logger.info(f"AWS Account: {config['aws_account_id']}")

    # ── Parse and validate dates ──
    try:
        start_date = parse_date(args.start_date)
        end_date = parse_date(args.end_date)
        validate_date_range(start_date, end_date)
    except ValueError as e:
        logger.error(f"Date validation failed: {e}")
        sys.exit(1)

    logger.info(f"Report Period: {format_date_range(start_date, end_date)}")

    # ── Determine output paths for each format ──
    output_format = args.output_format
    word_cfg = config.get('word_report', {}) or {}
    default_xlsx = config.get('output_filename', 'Weekly_AWS_BAU_Report.xlsx')
    default_docx = word_cfg.get('word_output_filename', 'Weekly_Status_Report.docx')

    explicit_output = args.output
    if not explicit_output:
        # No --output given: keep the config-default filenames per format.
        xlsx_output = default_xlsx
        docx_output = default_docx
    else:
        lower = explicit_output.lower()
        if lower.endswith('.docx'):
            # Explicit Word path: use it for docx, derive the xlsx sibling by
            # swapping the extension so --format both produces both files.
            docx_output = explicit_output
            xlsx_output = explicit_output[:-len('.docx')] + '.xlsx'
        elif lower.endswith('.xlsx'):
            # Explicit Excel path: use it for xlsx, derive the docx sibling.
            xlsx_output = explicit_output
            docx_output = explicit_output[:-len('.xlsx')] + '.docx'
        else:
            # No recognised extension: treat --output as a base stem and
            # append the correct extension for each format.
            xlsx_output = explicit_output + '.xlsx'
            docx_output = explicit_output + '.docx'

    logger.info(f"Output format: {output_format}")
    if output_format in ('xlsx', 'both'):
        logger.info(f"Excel output file: {xlsx_output}")
    if output_format in ('docx', 'both'):
        logger.info(f"Word output file: {docx_output}")

    overall_start = time.time()

    if args.mock:
        logger.info("Running in MOCK mode. Using high-fidelity sample datasets...")
        collected_data = get_mock_data(start_date, end_date)
    else:
        # ── Initialize AWS session ──
        try:
            session = boto3.Session()
            sts = session.client('sts')
            identity = sts.get_caller_identity()
            account_id = identity['Account']
            
            # Auto-detect current CloudShell region
            current_region = session.region_name
            if current_region:
                config['regions'] = [current_region]
                logger.info(f"Auto-detected AWS Region: {current_region}")
            
            logger.info(f"Authenticated as: {identity['Arn']}")
            logger.info(f"Account ID: {account_id}")

            # Verify account ID matches config
            if account_id != config['aws_account_id']:
                logger.warning(
                    f"Connected account ({account_id}) does not match "
                    f"configured account ({config['aws_account_id']})"
                )
        except Exception as e:
            logger.error(f"AWS authentication failed: {e}")
            logger.error("Ensure you are running this in AWS CloudShell or have valid AWS credentials.")
            sys.exit(1)

        # ── Collect data ──
        logger.info("")
        logger.info("Starting data collection from AWS APIs...")
        collected_data = collect_all_data(session, config, start_date, end_date, logger)

    overall_elapsed = time.time() - overall_start
    if not args.mock:
        logger.info(f"Data collection completed in {overall_elapsed:.1f}s")

    # ── Print summary ──
    print_summary(collected_data, start_date, end_date, logger)

    # ── Generate report(s) ──
    generated_files = []

    if output_format in ('xlsx', 'both'):
        logger.info("")
        logger.info("Generating Excel report...")
        report_start = time.time()
        try:
            generator = ExcelGenerator(config, start_date, end_date)
            generator.generate(collected_data, xlsx_output)
            report_elapsed = time.time() - report_start
            logger.info(f"✓ Excel report generated in {report_elapsed:.1f}s")
            logger.info(f"✓ Report saved to: {os.path.abspath(xlsx_output)}")
            generated_files.append(xlsx_output)
        except Exception as e:
            logger.error(f"Excel report generation failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            sys.exit(1)

    if output_format in ('docx', 'both'):
        logger.info("")
        logger.info("Generating Word report...")
        report_start = time.time()
        try:
            word_generator = WordReport(config, start_date, end_date)
            word_generator.generate(collected_data, docx_output)
            report_elapsed = time.time() - report_start
            logger.info(f"✓ Word report generated in {report_elapsed:.1f}s")
            logger.info(f"✓ Report saved to: {os.path.abspath(docx_output)}")
            generated_files.append(docx_output)
        except WordReportError as e:
            # python-docx unavailable. If Word was the only requested format,
            # fail clearly; otherwise warn and keep the Excel output.
            logger.error(f"Word report generation skipped: {e}")
            if output_format == 'docx':
                sys.exit(1)
        except Exception as e:
            logger.error(f"Word report generation failed: {e}")
            import traceback
            logger.error(traceback.format_exc())
            if output_format == 'docx':
                sys.exit(1)

    # ── Done ──
    total_elapsed = time.time() - overall_start
    files_line = ", ".join(generated_files) if generated_files else "(none)"
    print()
    print("╔══════════════════════════════════════════════════════════════════╗")
    print(f"║  ✓ Report generated successfully!                              ║")
    print("╚══════════════════════════════════════════════════════════════════╝")
    print(f"  Files: {files_line}")
    print(f"  Time:  {total_elapsed:.1f}s")
    print()


if __name__ == "__main__":
    main()
