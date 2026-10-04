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
        help='Output Excel file path (overrides config setting)'
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

def get_mock_data(start_date: date, end_date: date) -> dict:
    """Generate realistic mock data matching the exact report spec."""
    # Generate daily cost entries dynamically based on start_date and end_date
    daily_costs = []
    curr_date = start_date
    idx = 0
    mock_cost_pattern = [420.50, 435.10, 415.80, 460.25, 450.40, 470.15, 464.25]
    while curr_date <= end_date:
        daily_costs.append({
            "date": curr_date.strftime("%Y-%m-%d"),
            "cost": mock_cost_pattern[idx % len(mock_cost_pattern)]
        })
        curr_date += timedelta(days=1)
        idx += 1

    return {
        "cost": {
            "current_week_total": 3116.45,
            "previous_week_total": 3250.00,
            "difference": -133.55,
            "pct_change": -4.11,
            "avg_daily_cost": 445.21,
            "daily_costs": daily_costs,
            "top_services": [
                {"service": "Amazon Elastic Compute Cloud - Compute", "cost": 1850.20},
                {"service": "Amazon Relational Database Service", "cost": 720.50},
                {"service": "Amazon Simple Storage Service", "cost": 310.25},
                {"service": "Elastic Load Balancing", "cost": 120.40},
                {"service": "AWS WAF", "cost": 85.10},
                {"service": "Amazon CloudWatch", "cost": 30.00},
            ],
            "service_breakdown": [
                {"service": "Amazon Elastic Compute Cloud - Compute", "cost": 1850.20, "pct_change": -2.5, "difference": -47.40},
                {"service": "Amazon Relational Database Service", "cost": 720.50, "pct_change": 1.2, "difference": 8.50},
                {"service": "Amazon Simple Storage Service", "cost": 310.25, "pct_change": 4.5, "difference": 13.30},
                {"service": "Elastic Load Balancing", "cost": 120.40, "pct_change": -0.8, "difference": -1.00},
                {"service": "AWS WAF", "cost": 85.10, "pct_change": 12.0, "difference": 9.10},
                {"service": "Amazon CloudWatch", "cost": 30.00, "pct_change": 0.0, "difference": 0.00},
            ]
        },
        "ec2": [
            {
                "name": "InSync Add-in Server",
                "instance_id": "i-07c0fb23505c71dd8",
                "instance_type": "r6a.8xlarge",
                "state": "running",
                "availability_zone": "us-east-1a",
                "region": "us-east-1",
                "cpu_min": 0.96,
                "cpu_max": 18.00,
                "cpu_avg": 4.79,
                "memory_min": 56.5,
                "memory_max": 57.0,
                "memory_avg": 56.9,
                "disk_utilization": "C: Min 48.9% Max 49.0% Avg 49.0% | D: Min 12.3% Max 12.8% Avg 12.4%",
                "network_in": 390 * 1024 * 1024,
                "network_out": 810 * 1024 * 1024,
                "network_packets_in": 263000,
                "network_packets_out": 380000
            },
            {
                "name": "ISTARI-App Server",
                "instance_id": "i-0e68af64896637538",
                "instance_type": "m7i.large",
                "state": "stopped",
                "availability_zone": "us-east-1b",
                "region": "us-east-1",
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
                "network_packets_out": None
            },
            {
                "name": "ISTARI-DB Server",
                "instance_id": "i-0bc7cb7733cccd23b",
                "instance_type": "m7i.12xlarge",
                "state": "running",
                "availability_zone": "us-east-1c",
                "region": "us-east-1",
                "cpu_min": 0.16,
                "cpu_max": 1.86,
                "cpu_avg": 0.28,
                "memory_min": 11.5,
                "memory_max": 12.8,
                "memory_avg": 12.2,
                "disk_utilization": "/ Min 23.0% Max 24.0% Avg 23.6%",
                "network_in": 251 * 1024 * 1024,
                "network_out": 672 * 1024 * 1024,
                "network_packets_in": 128000,
                "network_packets_out": 292000
            },
            {
                "name": "InSync Website Server",
                "instance_id": "i-0a8f07365bce735f6",
                "instance_type": "t3a.medium",
                "state": "running",
                "availability_zone": "us-east-1d",
                "region": "us-east-1",
                "cpu_min": 1.42,
                "cpu_max": 54.40,
                "cpu_avg": 2.47,
                "memory_min": 30.5,
                "memory_max": 33.2,
                "memory_avg": 31.9,
                "disk_utilization": "/ Min 12.0% Max 13.5% Avg 12.8%",
                "network_in": 168 * 1024 * 1024,
                "network_out": 247 * 1024 * 1024,
                "network_packets_in": 24000,
                "network_packets_out": 102000
            },
            {
                "name": "InSync ClickHouse-Database",
                "instance_id": "i-029867634022c4d78",
                "instance_type": "c7i.12xlarge",
                "state": "running",
                "availability_zone": "us-east-1a",
                "region": "us-east-1",
                "cpu_min": 2.75,
                "cpu_max": 22.20,
                "cpu_avg": 6.02,
                "memory_min": 31.0,
                "memory_max": 35.0,
                "memory_avg": 33.2,
                "disk_utilization": "/ Min 39.5% Max 41.2% Avg 40.3%",
                "network_in": 694 * 1024 * 1024,
                "network_out": int(1.2 * 1024 * 1024 * 1024),
                "network_packets_in": 461000,
                "network_packets_out": 509000
            },
            {
                "name": "InSync Nexus Server",
                "instance_id": "i-09a1b5ff52381b04b",
                "instance_type": "m6i.8xlarge",
                "state": "running",
                "availability_zone": "us-east-1b",
                "region": "us-east-1",
                "cpu_min": 5.44,
                "cpu_max": 29.60,
                "cpu_avg": 7.63,
                "memory_min": 49.5,
                "memory_max": 52.5,
                "memory_avg": 51.0,
                "disk_utilization": "/ Min 71.0% Max 74.5% Avg 72.8%",
                "network_in": int(1.2 * 1024 * 1024 * 1024),
                "network_out": 421 * 1024 * 1024,
                "network_packets_in": 1400000,
                "network_packets_out": 108000
            },
            {
                "name": "InSync Weave-Server",
                "instance_id": "i-05966c44f5be655cc",
                "instance_type": "c7i.12xlarge",
                "state": "running",
                "availability_zone": "us-east-1c",
                "region": "us-east-1",
                "cpu_min": 6.12,
                "cpu_max": 10.70,
                "cpu_avg": 6.54,
                "memory_min": 19.5,
                "memory_max": 21.5,
                "memory_avg": 20.5,
                "disk_utilization": "/ Min 49.5% Max 52.0% Avg 50.9%",
                "network_in": 261 * 1024 * 1024,
                "network_out": 12 * 1024 * 1024,
                "network_packets_in": 273000,
                "network_packets_out": 22000
            },
            {
                "name": "InSync-Pipelines Server",
                "instance_id": "i-08b99bf46815a7d38",
                "instance_type": "m7i.large",
                "state": "running",
                "availability_zone": "us-east-1d",
                "region": "us-east-1",
                "cpu_min": 1.04,
                "cpu_max": 7.13,
                "cpu_avg": 1.46,
                "memory_min": 45.0,
                "memory_max": 50.0,
                "memory_avg": 48.1,
                "disk_utilization": "/ Min 61.5% Max 65.0% Avg 63.3%",
                "network_in": 2 * 1024 * 1024 * 1024,
                "network_out": 101 * 1024 * 1024,
                "network_packets_in": 98000,
                "network_packets_out": 47200
            },
            {
                "name": "ISTARI-App-Server-ASG",
                "instance_id": "i-06559919372821d45",
                "instance_type": "c7i.12xlarge",
                "state": "running",
                "availability_zone": "us-east-1a",
                "region": "us-east-1",
                "cpu_min": 0.57,
                "cpu_max": 4.28,
                "cpu_avg": 1.51,
                "memory_min": None,
                "memory_max": None,
                "memory_avg": None,
                "disk_utilization": "/ Min 13.5% Max 15.2% Avg 14.5%",
                "network_in": 575 * 1024 * 1024,
                "network_out": 241 * 1024 * 1024,
                "network_packets_in": 424000,
                "network_packets_out": 21600
            }
        ],
        "elb": [
            {
                "name": "ISTARI-ALB",
                "requests": 18900,
                "active_connections": 28000,
                "new_connections": 13100,
                "consumed_lcus": 0.017,
                "http_redirect_count": 409,
                "processed_bytes": int(1.70 * 1024 * 1024),
                "target_response_time": 1.74
            },
            {
                "name": "NEXUS-ALB",
                "requests": 34000,
                "active_connections": 45500,
                "new_connections": 18400,
                "consumed_lcus": 0.021,
                "http_redirect_count": 560,
                "processed_bytes": int(2.1 * 1024 * 1024),
                "target_response_time": 0.36
            },
            {
                "name": "Weave-ALB",
                "requests": 1640,
                "active_connections": 4600,
                "new_connections": 6130,
                "consumed_lcus": 0.001,
                "http_redirect_count": 758,
                "processed_bytes": 298 * 1024 * 1024,
                "target_response_time": 0.83
            }
        ],
        "waf": [
            {
                "name": "Istari-App-NewDeployment",
                "total_requests": 7660,
                "blocked_requests": 15,
                "allowed_requests": 7620,
                "captcha_requests": 4,
                "challenge_requests": 20
            }
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
        "rds": [],
        "inspector": {
            "date_range": f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}",
            "severity_counts": {
                "CRITICAL": 2,
                "HIGH": 6,
                "MEDIUM": 7,
                "LOW": 4,
                "INFORMATIONAL": 1,
            },
            "total_findings": 20,
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
            "severity_counts": {"HIGH": 3, "MEDIUM": 3, "LOW": 2},
            "total_findings": 8,
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

    # ── Determine output path ──
    output_path = args.output or config.get('output_filename', 'Weekly_AWS_BAU_Report.xlsx')
    logger.info(f"Output File: {output_path}")

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

    # ── Generate Excel report ──
    logger.info("")
    logger.info("Generating Excel report...")
    report_start = time.time()

    try:
        generator = ExcelGenerator(config, start_date, end_date)
        generator.generate(collected_data, output_path)
        report_elapsed = time.time() - report_start
        logger.info(f"✓ Excel report generated in {report_elapsed:.1f}s")
        logger.info(f"✓ Report saved to: {os.path.abspath(output_path)}")
    except Exception as e:
        logger.error(f"Report generation failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        sys.exit(1)

    # ── Done ──
    total_elapsed = time.time() - overall_start
    print()
    print("╔══════════════════════════════════════════════════════════════════╗")
    print(f"║  ✓ Report generated successfully!                              ║")
    print(f"║  File: {output_path:<56} ║")
    print(f"║  Time: {total_elapsed:.1f}s{' ' * (55 - len(f'{total_elapsed:.1f}s'))} ║")
    print("╚══════════════════════════════════════════════════════════════════╝")
    print()


if __name__ == "__main__":
    main()
