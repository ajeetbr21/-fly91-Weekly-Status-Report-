# Weekly AWS Infrastructure BAU Matrix Report

## Overview

A production-ready Python application that runs on **AWS CloudShell** to automatically collect AWS infrastructure metrics and generate a professional, management-ready Weekly BAU (Business As Usual) Matrix Report in Excel format.

**Client:** Insync Analytics  
**AWS Account:** 179787470151  
**Phase:** 1 (Single Account)

---

## Features

### Data Collection
- **AWS Cost Explorer** — Weekly cost summary, daily trends, top cost-contributing services
- **Amazon EC2** — Instance details, CPU/Memory/Disk utilization, network bandwidth
- **CloudWatch Agent** — Memory and disk metrics (requires CWAgent installed)
- **Elastic Load Balancer (ALB)** — Request counts, connections, processed bytes, response times
- **AWS WAF** — Web ACL request metrics (allowed, blocked, CAPTCHA, challenge)
- **Amazon S3** — Bucket sizes, object counts across all buckets
- **Amazon RDS** — Database CPU, storage, connections (if applicable)
- **Amazon Inspector** — Vulnerability findings for the reporting week, grouped by severity (Critical / High / Medium / Low / Informational)
- **Amazon GuardDuty** — Threat-detection findings for the reporting week, grouped by severity (High / Medium / Low)
- **Auto Scaling** — Group configurations and instance health

### Excel Report
- **Cover Page** — Professional branding with client details
- **Executive Summary** — Auto-generated key highlights and health indicators
- **Cost Summary** — Week-over-week comparison, daily trends, top services
- **EC2 Dashboard** — Server utilization and network bandwidth tables
- **Load Balancer Dashboard** — ALB metrics and performance data
- **WAF Dashboard** — Security request metrics
- **Inspector Dashboard** — Severity summary + detailed one-week findings table, plus a console-screenshot-style visual (see below)
- **GuardDuty Dashboard** — Severity summary + detailed one-week findings table, plus a console-screenshot-style visual (see below)
- **S3 Dashboard** — Storage distribution across buckets
- **RDS Dashboard** — Database performance (or "No RDS instances available")
- **Charts Dashboard** — 8 professional charts (cost trends, CPU, memory, network, etc.)

> The **Cost Summary** now also lists **Amazon Inspector** and **Amazon GuardDuty** as line items in the top-services and service-breakdown tables.

### Screenshot-style visuals

The Inspector and GuardDuty sheets each embed a **console-screenshot-style image**
of the one-week findings, mimicking the "Findings" view captured from the AWS
console. The image is rendered on the fly with [Pillow](https://pillow.readthedocs.io/)
(`report/screenshots.py`): a dark-blue title bar, column headers, zebra-striped
rows, and colour-coded severity chips (**High = red, Medium = orange, Low = blue**).

The visuals are generated **in memory** (via `BytesIO`) and embedded directly
into the workbook, so no PNG files are written to disk or committed. If Pillow is
not installed, or image rendering fails for any reason, the report **degrades
gracefully**: a warning is logged and the sheet is produced without the embedded
image (the tables are unaffected).

### Formatting
- Dark blue title bars with white text
- Alternating row colors for readability
- Conditional formatting (Green/Yellow/Red) for health indicators
- Currency, percentage, and storage unit formatting
- Frozen headers and auto-filters
- Auto-adjusted column widths

---

## Prerequisites

1. **AWS CloudShell** (recommended) or any environment with:
   - Python 3.8+
   - Valid AWS credentials with read access to the required services
2. **IAM Permissions** — The following permissions are required:
   - `ce:GetCostAndUsage`
   - `ec2:DescribeInstances`
   - `cloudwatch:GetMetricStatistics`
   - `cloudwatch:ListMetrics`
   - `elasticloadbalancing:DescribeLoadBalancers`
   - `wafv2:ListWebACLs`
   - `wafv2:GetWebACL`
   - `s3:ListAllMyBuckets`
   - `s3:GetBucketLocation`
   - `rds:DescribeDBInstances`
   - `autoscaling:DescribeAutoScalingGroups`
   - `sts:GetCallerIdentity`

---

## Installation

### On AWS CloudShell

```bash
# 1. Upload or clone the project
#    Option A: Upload via CloudShell "Actions → Upload file"
#    Option B: Clone from a repository

# 2. Navigate to the project directory
cd aws-weekly-bau-report

# 3. Install dependencies
pip3 install -r requirements.txt --user
```

> **Note:** `boto3` and `botocore` are pre-installed on AWS CloudShell.
> `Pillow` is used for the Inspector/GuardDuty screenshot-style visuals; if it
> cannot be installed, the report still generates (the images are skipped).

---

## Usage

### Basic Usage

```bash
python3 weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22
```

### Custom Output File

```bash
python3 weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22 --output my_report.xlsx
```

### Custom Configuration

```bash
python3 weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22 --config /path/to/config.json
```

### Help

```bash
python3 weekly_bau_report.py --help
```

---

## Configuration

Edit `config.json` to customize the report:

```json
{
    "client_name": "Insync Analytics",
    "aws_account_id": "179787470151",
    "prepared_by": "Cloud Operations Team",
    "report_title": "Weekly AWS Infrastructure BAU Matrix Report",
    "output_filename": "Weekly_AWS_BAU_Report.xlsx",
    "regions": ["us-east-1"],
    "thresholds": {
        "cpu_warning": 70,
        "cpu_critical": 90,
        "memory_warning": 75,
        "memory_critical": 90,
        "disk_warning": 75,
        "disk_critical": 90,
        "cost_increase_warning": 10,
        "cost_increase_critical": 25
    },
    "logging": {
        "level": "INFO",
        "file": "bau_report.log"
    }
}
```

### Configuration Fields

| Field | Description |
|---|---|
| `client_name` | Client/organization name displayed on the report |
| `aws_account_id` | AWS account ID for verification |
| `prepared_by` | Name/team shown on the cover page |
| `regions` | AWS regions to collect data from |
| `thresholds` | Warning/critical thresholds for conditional formatting |

---

## Output

The script generates:

| File | Description |
|---|---|
| `Weekly_AWS_BAU_Report.xlsx` | Professional Excel workbook with all dashboards |
| `bau_report.log` | Detailed execution log |

---

## Project Structure

```
aws-weekly-bau-report/
├── weekly_bau_report.py          # Main entry point (CLI)
├── config.json                   # Configuration
├── requirements.txt              # Dependencies
├── README.md                     # This file
├── collectors/                   # AWS data collectors
│   ├── __init__.py
│   ├── base_collector.py         # Abstract base with retry logic
│   ├── cost_collector.py         # Cost Explorer
│   ├── ec2_collector.py          # EC2 + CloudWatch metrics
│   ├── elb_collector.py          # Load Balancer metrics
│   ├── waf_collector.py          # WAF metrics
│   ├── s3_collector.py           # S3 bucket metrics
│   ├── rds_collector.py          # RDS metrics
│   └── autoscaling_collector.py  # Auto Scaling info
├── report/                       # Excel report generation
│   ├── __init__.py
│   ├── excel_generator.py        # Report orchestrator
│   ├── styles.py                 # Formatting & styles
│   ├── charts.py                 # Chart helpers
│   ├── cover_sheet.py            # Cover page
│   ├── executive_summary.py      # Executive summary
│   ├── cost_sheet.py             # Cost dashboard
│   ├── ec2_sheet.py              # EC2 dashboard
│   ├── elb_sheet.py              # ELB dashboard
│   ├── waf_sheet.py              # WAF dashboard
│   ├── inspector_sheet.py        # Amazon Inspector dashboard (+ screenshot)
│   ├── guardduty_sheet.py        # Amazon GuardDuty dashboard (+ screenshot)
│   ├── screenshots.py            # Pillow helper for console-style PNG visuals
│   ├── s3_sheet.py               # S3 dashboard
│   ├── rds_sheet.py              # RDS dashboard
│   └── dashboard_sheet.py        # Charts dashboard
└── utils/                        # Utilities
    ├── __init__.py
    ├── logger.py                 # Logging setup
    └── helpers.py                # Formatting helpers
```

---

## Troubleshooting

### "openpyxl not found"
```bash
pip3 install openpyxl --user
```

### "Access Denied" errors
Ensure your IAM role/user has the required permissions listed in Prerequisites.

### Memory/Disk metrics show "N/A"
The CloudWatch Agent must be installed and configured on your EC2 instances to report memory and disk utilization metrics. These are not available by default.

### No data for a specific period
CloudWatch retains:
- 1-minute data points for 15 days
- 5-minute data points for 63 days
- 1-hour data points for 455 days

### Cost Explorer shows $0
Cost Explorer data may take up to 24 hours to be available. Also ensure the Cost Explorer API is enabled in your AWS account.

---

## CloudWatch Agent Notes

To get Memory and Disk utilization metrics, install the CloudWatch Agent on your EC2 instances:

```bash
# Amazon Linux 2 / Amazon Linux 2023
sudo yum install amazon-cloudwatch-agent -y

# Configure the agent
sudo /opt/aws/amazon-cloudwatch-agent/bin/amazon-cloudwatch-agent-config-wizard

# Start the agent
sudo systemctl start amazon-cloudwatch-agent
sudo systemctl enable amazon-cloudwatch-agent
```

---

## Reference / Assumptions

This report's layout targets the structure of the reference Google Doc
("fly91-Weekly-Status-Report- 14th September - 20th September"): a service-wise
**Cost Summary** that includes Amazon Inspector and Amazon GuardDuty, EC2 Server
Utilization, WAF request metrics, and the **Amazon Inspector one-week Findings**
captured as console screenshots (with High / Medium / Low severity), plus the
GuardDuty findings.

**Important:** the reference Google Doc itself was **not machine-accessible**
during implementation — it requires Google sign-in, so only its title was
visible. The structure above therefore follows the description relayed with the
task, **not** the live document. As a result, some exact wording, column
ordering, and numeric values (including the Inspector ~$18.40 and GuardDuty
~$12.75 weekly costs, and the sample findings) are **assumptions** and should be
confirmed against the original document. Please review and let us know what, if
anything, is still missing or should be adjusted.

All AWS data shown when running with `--mock` is representative sample data; run
without `--mock` (with valid AWS credentials) to populate the report from the
live account.

---

## License

Internal use only — Insync Analytics.
