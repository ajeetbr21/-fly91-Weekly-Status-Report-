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
   - `cloudwatch:GetMetricData`
   - `cloudwatch:ListMetrics`
   - `elasticloadbalancing:DescribeLoadBalancers`
   - `wafv2:ListWebACLs`
   - `wafv2:GetWebACL`
   - `inspector2:ListFindings`
   - `guardduty:ListDetectors`
   - `guardduty:ListFindings`
   - `guardduty:GetFindings`
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
> `matplotlib` is used to render the per-server CloudWatch metric graphs in the
> Word report (headless `Agg` backend); if it cannot be installed, the Word
> report still generates with the graphs omitted.

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

### Choosing the Output Format (Excel and/or Word)

The tool can produce the existing Excel (`.xlsx`) workbook, a Microsoft Word
(`.docx`) **Weekly Status Report** that replicates the client reference
document, or both. Use the `--format` flag:

```bash
# Excel only (default — unchanged behaviour)
python3 weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22

# Word only
python3 weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22 --format docx

# Both Excel and Word
python3 weekly_bau_report.py --start-date 2026-06-16 --end-date 2026-06-22 --format both
```

Notes:

- `--format` accepts `xlsx` (default), `docx`, or `both`.
- When `--output` ends in `.docx`, it is used as the Word output path; otherwise
  it is treated as the Excel output path. The Word output path otherwise comes
  from `word_report.word_output_filename` in `config.json`
  (default `Weekly_Status_Report.docx`), and the Excel path from `output_filename`.
- If the `--output` suffix is incompatible with `--format` (for example
  `--format xlsx --output report.docx`, or `--format docx --output report.xlsx`),
  the tool logs a warning and the mismatched path is ignored; routing is
  otherwise unchanged.
- The cover-page report date is the submission date, computed as the period end
  date **+ 3 days** (e.g. a 14-20 Sep 2026 week shows `23/09/2026`).
- The Word report is scoped to a **single account** (JUST UDO AVIATION PRIVATE
  LIMITED (Fly91), account `674351849978`, prepared by Greatworx). It contains,
  in order: a cover page, a "security best practices" link table, a per-account
  Summary section (Billing & Cost Overview plus optional Resource Utilization &
  Alarms table), and an "-- End Of Document --" trailer. It does **not** include
  the multi-account "Cost Summary Difference of All AWS Accounts" fleet table.
  Cost date ranges and alarm dates are populated for whatever week you pass via
  `--start-date`/`--end-date`.
- Mock mode (`--mock`) synthesises a representative single-account dataset so the
  Word report is complete without AWS access. In live mode the configured
  account is populated from the real collected cost/EC2 data.
- The "Resource Utilization & Alarms" section now embeds AWS CloudWatch metric
  graphs per server (one chart image per Fly91 server). In production the charts
  are drawn from real CloudWatch time-series collected via `GetMetricStatistics`
  and `GetMetricData`; in `--mock` mode they are synthetic sample charts so the
  report is complete without AWS access. Charts are rendered with matplotlib
  using the headless `Agg` backend, so no display is required. Memory and disk
  metrics require the CloudWatch Agent to be installed on the instances in
  production (CPU and network metrics are native to AWS/EC2). Chart rendering
  degrades gracefully: if matplotlib is unavailable or the metric data is empty,
  the images are skipped and the surrounding text and alarm tables are still
  generated.

You can verify a generated Word report with the bundled checker:

```bash
python3 scripts/verify_word_report.py Weekly_Status_Report.docx

# Also assert that the report embeds at least one CloudWatch graph image:
python3 scripts/verify_word_report.py Weekly_Status_Report.docx --min-images 1
```

The live single-account fallback (no mock `word_accounts`, no
`word_report.accounts` in config) has its own standalone checker, which builds a
real-shaped `cost`/`ec2` payload and asserts the single-account Cost Summary and
CPU-proxy alarm row:

```bash
python3 scripts/verify_word_live_fallback.py
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
    "client_name": "JUST UDO AVIATION PRIVATE LIMITED (Fly91)",
    "aws_account_id": "674351849978",
    "prepared_by": "Cloud Operations Team",
    "report_title": "Weekly AWS Infrastructure BAU Matrix Report",
    "output_filename": "Weekly_AWS_BAU_Report.xlsx",
    "word_report": {
        "report_title": "Weekly Status Report",
        "client_org": "JUST UDO AVIATION PRIVATE LIMITED (Fly91)",
        "submitted_by_label": "Submitted By",
        "submitter_org": "Greatworx",
        "activity_org": "Greatworx",
        "word_output_filename": "Weekly_Status_Report.docx"
    },
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
| `output_filename` | Default Excel output filename |
| `word_report` | Optional. Branding/output settings for the Word (`.docx`) report (see below) |

#### `word_report` fields (all optional)

| Field | Default | Description |
|---|---|---|
| `report_title` | `Weekly Status Report` | Cover-page title of the Word report |
| `client_org` | falls back to `client_name` | Client organization name on the cover page |
| `submitted_by_label` | `Submitted By` | Label shown above the submitter org |
| `submitter_org` | `Operisoft Technologies Pvt Ltd` | Organization that prepared the report |
| `activity_org` | falls back to `submitter_org` | Org name used in "No Activity performed by &lt;org&gt;" lines |
| `word_output_filename` | `Weekly_Status_Report.docx` | Default Word output filename |

All `word_report` keys are optional; omitting the section (or any key) falls
back to the defaults above, so an older `config.json` keeps working. Swap these
strings to re-brand the report (e.g. Aptech / Operisoft / Fly91).

### Word report (multi-account mapping)

The client reference document is a **multi-account** report, while the live AWS
collectors in this project target a **single** account. The Word report is
therefore driven by an `accounts` list:

- In `--mock` mode the list is synthesised in `get_mock_word_accounts()` so the
  generated `.docx` mirrors the full reference layout; alarm dates are
  parameterised to the selected week.
- In live mode the single configured account (`client_name` / `aws_account_id`)
  is populated from the real collected cost and EC2 data, and appears as account
  #1. Remaining accounts can be supplied via an optional
  `word_report.accounts` list in `config.json` if you want the full multi-account
  table in production.
- The 48 per-account console screenshots embedded in the original reference
  `.docx` are account-specific captures that cannot be regenerated from mock
  data, so the generated report omits them and renders the data as text/tables.

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
