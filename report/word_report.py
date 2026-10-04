"""
Word (.docx) report generator for the AWS Weekly Status Report.

Replicates the client's reference Word document per-account layout (see
REFERENCE_DOC_TEXT.txt) using python-docx. The report is scoped to a SINGLE
account - JUST UDO AVIATION PRIVATE LIMITED (Fly91) - so the multi-account
"Cost Summary Difference of All AWS Accounts" fleet table is NOT produced. The
document is built in the following order:

    1. Cover page (report title, client org, "Submitted By", submitter org,
       report date, reporting period).
    2. "security best practices" two-column Content | Link table (8 reference
       links verbatim).
    3. Per-account "Summary" section (Billing and Cost Overview + optional
       Resource Utilization & Alarms table) for the single Fly91 account.
    4. Trailer line "-- End Of Document --".

The sections are driven by an `accounts` list supplied in
`collected_data["word_accounts"]` (a single-element list in single-account
scope). The generator mirrors the orchestration / graceful-degradation style of
report/excel_generator.py: each section is wrapped in try/except so a failure
in one section logs a warning and the rest of the document is still produced.

All type hints are Python 3.9 compatible (typing.Optional / typing.Union,
no PEP 604 "X | None" unions).
"""

import logging
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor
    _DOCX_AVAILABLE = True
    _DOCX_IMPORT_ERROR = None
except ImportError as exc:  # pragma: no cover - exercised only without python-docx
    Document = None
    WD_ALIGN_PARAGRAPH = None
    Pt = None
    RGBColor = None
    _DOCX_AVAILABLE = False
    _DOCX_IMPORT_ERROR = exc

from utils.helpers import (
    format_currency,
    format_date_range,
    get_previous_period,
)

logger = logging.getLogger(__name__)


class WordReportError(RuntimeError):
    """Raised when the Word report cannot be generated at all."""


# The 8 security best-practice links, verbatim from the reference document.
SECURITY_BEST_PRACTICES = [
    ("Best Practices for AWS root users",
     "https://docs.aws.amazon.com/accounts/latest/reference/best-practices-root-user.html"),
    ("Best Practices for AWS Access Keys",
     "https://docs.aws.amazon.com/accounts/latest/reference/credentials-access-keys-best-practices.html"),
    ("Shared Responsibility Model",
     "https://aws.amazon.com/compliance/shared-responsibility-model/"),
    ("AWS Cloudtrail",
     "https://aws.amazon.com/cloudtrail/"),
    ("Trusted Advisor",
     "https://aws.amazon.com/premiumsupport/trustedadvisor/"),
    ("Creating Billing alarms",
     "https://docs.aws.amazon.com/AmazonCloudWatch/latest/monitoring/"
     "gs_monitor_estimated_charges_with_cloudwatch.html#gs_creating_billing_alarm"),
    ("Enable MFA",
     "https://docs.aws.amazon.com/IAM/latest/UserGuide/id_credentials_mfa.html"),
    ("GIT Secrets",
     "https://github.com/awslabs/git-secrets"),
]


class WordReport:
    """
    Orchestrates generation of the Weekly Status Report Word document.

    Usage
    -----
    >>> gen = WordReport(config, start_date, end_date)
    >>> path = gen.generate(collected_data, "output/report.docx")
    """

    DEFAULT_WORD_CONFIG = {
        "report_title": "Weekly Status Report",
        "client_org": "JUST UDO AVIATION PRIVATE LIMITED (Fly91)",
        "submitted_by_label": "Submitted By",
        "submitter_org": "Greatworx",
        "activity_org": "Greatworx",
        "word_output_filename": "Weekly_Status_Report.docx",
    }

    def __init__(self, config: Optional[Dict[str, Any]] = None,
                 start_date: Optional[date] = None,
                 end_date: Optional[date] = None):
        config = config or {}
        self.config = config
        word_cfg = dict(self.DEFAULT_WORD_CONFIG)
        word_cfg.update(config.get("word_report", {}) or {})
        # Fall back to the client_name for the client org if not overridden.
        if not config.get("word_report", {}).get("client_org"):
            word_cfg["client_org"] = config.get("client_name", word_cfg["client_org"])
        # activity_org defaults to submitter_org if unset.
        if not config.get("word_report", {}).get("activity_org"):
            word_cfg["activity_org"] = word_cfg["submitter_org"]
        self.word_cfg = word_cfg
        self.start_date = start_date
        self.end_date = end_date

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def generate(self, collected_data: Optional[Dict[str, Any]],
                 output_path: Union[str, Path]) -> str:
        """
        Build the Word document and save it to *output_path*.

        Raises
        ------
        WordReportError
            If python-docx is not importable (so the CLI can report clearly).
        """
        if not _DOCX_AVAILABLE:
            raise WordReportError(
                "python-docx is not installed; cannot generate the Word report. "
                "Install it via 'pip install python-docx'. "
                f"(import error: {_DOCX_IMPORT_ERROR})"
            )

        collected_data = collected_data or {}
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        accounts = self._resolve_accounts(collected_data)

        logger.info("Creating Word document ...")
        doc = Document()

        doc.core_properties.title = self.word_cfg["report_title"]
        doc.core_properties.author = self.word_cfg["submitter_org"]
        doc.core_properties.company = self.word_cfg["client_org"]

        sections = [
            ("Cover page", self._build_cover, {"doc": doc}),
            ("Security best practices", self._build_security_best_practices,
             {"doc": doc}),
            ("Per-account summaries", self._build_account_summaries,
             {"doc": doc, "accounts": accounts}),
            ("Trailer", self._build_trailer, {"doc": doc}),
        ]

        for name, func, kwargs in sections:
            try:
                logger.info("Generating Word section: %s", name)
                func(**kwargs)
            except Exception as exc:  # graceful degradation per section
                logger.error("Failed to generate Word section '%s': %s",
                             name, exc, exc_info=True)
                try:
                    p = doc.add_paragraph()
                    run = p.add_run(f"[Section '{name}' could not be generated: {exc}]")
                    run.italic = True
                except Exception:
                    pass

        doc.save(str(output_path))
        logger.info("Word report saved to %s", output_path.resolve())
        return str(output_path.resolve())

    # ------------------------------------------------------------------
    # Date helpers
    # ------------------------------------------------------------------

    def _current_range_str(self) -> str:
        if self.start_date and self.end_date:
            return format_date_range(self.start_date, self.end_date)
        return "N/A"

    def _previous_range_str(self) -> str:
        if self.start_date and self.end_date:
            prev_start, prev_end = get_previous_period(self.start_date, self.end_date)
            return format_date_range(prev_start, prev_end)
        return "N/A"

    def _report_date_str(self) -> str:
        # Report date = submission date, three days after the current period
        # end, formatted dd/mm/yyyy to match the reference (period 14-20 Sep
        # 2026 -> cover date "23/09/2026").
        if self.end_date:
            from datetime import timedelta
            submit = self.end_date + timedelta(days=3)
            return submit.strftime("%d/%m/%Y")
        return date.today().strftime("%d/%m/%Y")

    def _period_days(self) -> int:
        if self.start_date and self.end_date:
            return (self.end_date - self.start_date).days + 1
        return 7

    # ------------------------------------------------------------------
    # Account resolution: real single account merged with mock/config list
    # ------------------------------------------------------------------

    def _resolve_accounts(self, collected_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Resolve the account(s) driving the per-account summary section.

        The report is single-account scoped, so this normally yields a
        one-element list. Priority:
            1. collected_data["word_accounts"] (mock mode supplies this).
            2. config["word_report"]["accounts"] if present.
            3. A single account synthesised from the real collected cost data
               for the configured account.
        """
        accounts = collected_data.get("word_accounts")
        if accounts:
            return accounts

        cfg_accounts = (self.config.get("word_report", {}) or {}).get("accounts")
        if cfg_accounts:
            return cfg_accounts

        # Fall back: build a single account from real collected cost data.
        cost = collected_data.get("cost") or {}
        current = cost.get("current_week_total")
        previous = cost.get("previous_week_total")
        avg_daily = cost.get("avg_daily_cost")
        services_text = "The cost remains same."
        diff = cost.get("difference")
        if isinstance(diff, (int, float)):
            if diff < 0:
                services_text = "The costs decreased compared to the previous week."
            elif diff > 0:
                services_text = "The costs increased compared to the previous week."

        alarms = self._alarms_from_ec2(collected_data.get("ec2"))

        return [{
            "no": 1,
            "account_name": self.config.get("client_name", "Primary Account"),
            "account_id": str(self.config.get("aws_account_id", "")),
            "last_week_cost": previous,
            "tax_cost": None,
            "current_week_cost": current,
            "services_text": services_text,
            "avg_daily_cost": avg_daily,
            "activity_note": None,
            "alarms": alarms,
        }]

    def _alarms_from_ec2(self, ec2_data: Any) -> List[Dict[str, Any]]:
        """
        Best-effort mapping of real EC2 utilisation data into alarm rows.

        The live collectors do not expose CloudWatch-alarm trigger counts, so
        this produces at most one representative "high utilisation" row per
        running instance that exceeds a basic CPU threshold, with the trigger
        column left as a note. In mock mode this path is not used (word_accounts
        already carries explicit alarm data).
        """
        alarms: List[Dict[str, Any]] = []
        if not isinstance(ec2_data, list):
            return alarms
        thresholds = self.config.get("thresholds", {}) or {}
        cpu_warn = thresholds.get("cpu_warning", 70)
        for inst in ec2_data:
            if not isinstance(inst, dict):
                continue
            cpu_max = inst.get("cpu_max")
            if cpu_max is None or cpu_max < cpu_warn:
                continue
            name = inst.get("name", "N/A")
            iid = inst.get("instance_id", "")
            region = inst.get("region", "")
            alarms.append({
                "server_name": f"{name}({iid})" if iid else name,
                "region": region,
                "metric": f"CPU {cpu_warn}%",
                "triggers": [f"Peak CPU {cpu_max:.1f}% during the week"],
            })
        return alarms

    # ------------------------------------------------------------------
    # Section builders
    # ------------------------------------------------------------------

    def _build_cover(self, doc) -> None:
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run(self.word_cfg["report_title"])
        run.bold = True
        run.font.size = Pt(28)

        org = doc.add_paragraph()
        org.alignment = WD_ALIGN_PARAGRAPH.CENTER
        org_run = org.add_run(self.word_cfg["client_org"])
        org_run.bold = True
        org_run.font.size = Pt(18)

        label = doc.add_paragraph()
        label.alignment = WD_ALIGN_PARAGRAPH.CENTER
        label.add_run(self.word_cfg["submitted_by_label"])

        submitter = doc.add_paragraph()
        submitter.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_run = submitter.add_run(self.word_cfg["submitter_org"])
        sub_run.bold = True

        date_p = doc.add_paragraph()
        date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_p.add_run(self._report_date_str())

        period_p = doc.add_paragraph()
        period_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        period_run = period_p.add_run(f"Reporting Period: {self._current_range_str()}")
        period_run.italic = True

        doc.add_page_break()

    def _build_cost_summary(self, doc, accounts: List[Dict[str, Any]]) -> None:
        # NOTE: Retained for reference only. The report is single-account
        # scoped, so this multi-account fleet table is NOT invoked by
        # generate(). Per-account cost figures are rendered in
        # _build_account_summaries instead.
        doc.add_heading("Cost Summary Difference of All AWS Accounts", level=1)

        headers = [
            "No",
            "Account Name",
            "Account ID",
            f"Last Week Cost ({self._previous_range_str()})",
            "Tax Cost",
            f"Current Week Cost ({self._current_range_str()})",
            "Services",
        ]
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = "Table Grid"
        self._set_header_row(table.rows[0], headers)

        total_last = 0.0
        total_tax = 0.0
        total_curr = 0.0
        for acct in accounts:
            last = acct.get("last_week_cost")
            tax = acct.get("tax_cost")
            curr = acct.get("current_week_cost")
            total_last += last if isinstance(last, (int, float)) else 0.0
            total_tax += tax if isinstance(tax, (int, float)) else 0.0
            total_curr += curr if isinstance(curr, (int, float)) else 0.0

            cells = table.add_row().cells
            cells[0].text = str(acct.get("no", ""))
            cells[1].text = str(acct.get("account_name", ""))
            cells[2].text = str(acct.get("account_id", ""))
            cells[3].text = format_currency(last)
            cells[4].text = format_currency(tax)
            cells[5].text = format_currency(curr)
            cells[6].text = str(acct.get("services_text", ""))

        # Total Cost row with up/down arrow indicators.
        diff = total_curr - total_last
        up = "\u2b06\ufe0f"    # ⬆️
        down = "\u2b07\ufe0f"  # ⬇️
        last_arrow = up if total_last >= total_curr else down
        curr_arrow = down if diff < 0 else (up if diff > 0 else "")

        total_cells = table.add_row().cells
        total_cells[0].text = ""
        total_cells[1].text = "Total Cost"
        total_cells[2].text = ""
        total_cells[3].text = f"{last_arrow}{format_currency(total_last)}"
        total_cells[4].text = format_currency(total_tax)
        total_cells[5].text = f"{curr_arrow}{format_currency(total_curr)}"
        total_cells[6].text = ""
        for idx in (1, 3, 4, 5):
            self._bold_cell(total_cells[idx])

        # Week-over-week difference note.
        note = doc.add_paragraph()
        if diff < 0:
            direction = "decreased"
        elif diff > 0:
            direction = "increased"
        else:
            direction = "remained the same"
        note.add_run(
            f"The billing for the current week ({self._current_range_str()}) has "
            f"{direction} compared to the previous week. "
            f"The cost difference is {format_currency(abs(diff))}."
        )

    def _build_security_best_practices(self, doc) -> None:
        doc.add_heading("security best practices", level=1)
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        self._set_header_row(table.rows[0], ["Content", "Link"])
        for content, link in SECURITY_BEST_PRACTICES:
            cells = table.add_row().cells
            cells[0].text = content
            cells[1].text = link

    def _build_account_summaries(self, doc, accounts: List[Dict[str, Any]]) -> None:
        activity_org = self.word_cfg["activity_org"]
        for acct in accounts:
            name = str(acct.get("account_name", ""))
            account_id = str(acct.get("account_id", ""))

            doc.add_heading("Summary", level=1)
            header = doc.add_paragraph()
            header_run = header.add_run(f"{name} - {account_id}".strip(" -"))
            header_run.bold = True

            doc.add_heading("Billing and Cost Overview", level=2)
            curr = acct.get("current_week_cost")
            avg_daily = acct.get("avg_daily_cost")
            if avg_daily is None and isinstance(curr, (int, float)):
                avg_daily = round(curr / max(self._period_days(), 1), 2)

            doc.add_paragraph(f"Total cost for the week: {format_currency(curr)}")
            doc.add_paragraph(f"Average Daily Cost: {format_currency(avg_daily)}")
            tax = acct.get("tax_cost")
            if isinstance(tax, (int, float)):
                doc.add_paragraph(f"Total Tax Cost: {format_currency(tax)}")
            services_text = acct.get("services_text")
            if services_text:
                doc.add_paragraph(str(services_text))

            activity_note = acct.get("activity_note")
            if not activity_note:
                activity_note = f"No Activity performed by {activity_org} in this account."
            doc.add_paragraph(str(activity_note))

            alarms = acct.get("alarms") or []
            if alarms:
                doc.add_heading("Resource Utilization & Alarms", level=2)
                doc.add_paragraph(
                    f"All EC2 Server Alarms triggered between "
                    f"({self._current_range_str()})"
                )
                table = doc.add_table(rows=1, cols=4)
                table.style = "Table Grid"
                self._set_header_row(
                    table.rows[0],
                    ["Server Name", "Region", "Memory/Disk/CPU",
                     "Alert date and no. of trigger"],
                )
                for alarm in alarms:
                    cells = table.add_row().cells
                    cells[0].text = str(alarm.get("server_name", ""))
                    cells[1].text = str(alarm.get("region", ""))
                    cells[2].text = str(alarm.get("metric", ""))
                    triggers = alarm.get("triggers") or []
                    cells[3].text = "".join(str(t) for t in triggers)
            else:
                doc.add_paragraph("No Resource Utilization & Alarms.")

    def _build_trailer(self, doc) -> None:
        trailer = doc.add_paragraph()
        trailer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = trailer.add_run("-- End Of Document --")
        run.bold = True

    # ------------------------------------------------------------------
    # Low-level cell helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _set_header_row(row, headers: List[str]) -> None:
        for idx, text in enumerate(headers):
            cell = row.cells[idx]
            cell.text = str(text)
            WordReport._bold_cell(cell)

    @staticmethod
    def _bold_cell(cell) -> None:
        for paragraph in cell.paragraphs:
            if not paragraph.runs:
                paragraph.add_run("")
            for run in paragraph.runs:
                run.bold = True
