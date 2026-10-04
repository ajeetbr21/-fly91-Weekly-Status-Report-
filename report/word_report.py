"""
Word (.docx) report generator for the Fly91 Weekly Status Report.

Reproduces the client's REAL reference Google-Doc layout (see
REFERENCE_FLY91_GOOGLEDOC.txt) using python-docx. The report is a
SINGLE-ACCOUNT, SERVICE-ORIENTED document for JUST UDO AVIATION PRIVATE
LIMITED (Fly91) / Greatworx - account 674351849978 - built in this order:

    1. "Cost Summary Differences" - a per-service cost table (one row per AWS
       service, including Inspector and GuardDuty) with Last Week / Current
       Week columns and a Total Cost row, followed by the cost-analysis
       bullets (decrease amount, percentage, average daily cost, the services
       that drove the decrease and the offsetting increases).
    2. "Disclaimer" - the confidentiality paragraph, with a configurable
       preparer org (default "Greatworx").
    3. "Contents" - an index of the report sections plus the date line.
    4. "BAU Matrix Overview of the resources." - "Uptime of the servers ..."
       with six subsections: 1 EC2, 2 RDS, 3 ELB, 4 AWS WAF, 5 Amazon
       Inspector (findings summary + embedded console screenshots) and
       6 Guard Duty (findings summary + embedded console screenshot).
    5. "-- End of Document --" trailer.

The sections are driven by the SAME ``collected_data`` the Excel report uses
(``cost`` / ``ec2`` / ``elb`` / ``waf`` / ``rds`` / ``inspector`` /
``guardduty``). The generator mirrors the orchestration / graceful-degradation
style of report/excel_generator.py: each section is wrapped in try/except so a
failure in one section logs a warning and the rest of the document is still
produced. Every embedded image (Inspector / GuardDuty console screenshots via
report/screenshots.py, optional EC2 metric charts via report/metric_charts.py)
returns ``None`` and is skipped if Pillow/matplotlib is unavailable or the
data is empty - the surrounding text and tables still generate.

All type hints are Python 3.9 compatible (typing.Optional / typing.Union,
no PEP 604 "X | None" unions).
"""

import logging
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Union

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.shared import Inches, Pt, RGBColor
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    _DOCX_AVAILABLE = True
    _DOCX_IMPORT_ERROR = None
except ImportError as exc:  # pragma: no cover - exercised only without python-docx
    Document = None
    WD_ALIGN_PARAGRAPH = None
    WD_ALIGN_VERTICAL = None
    Inches = None
    Pt = None
    RGBColor = None
    OxmlElement = None
    qn = None
    _DOCX_AVAILABLE = False
    _DOCX_IMPORT_ERROR = exc

# Console-screenshot-style renderer (Inspector / GuardDuty). Degrades
# gracefully if Pillow is unavailable - render_findings_table returns None.
from report import screenshots

# CloudWatch-style metric chart renderer (optional EC2 graphs). Degrades
# gracefully if matplotlib is unavailable - render_metric_chart returns None.
from report import metric_charts

from utils.helpers import (
    format_bytes,
    format_count,
    format_currency,
    format_date_range,
    get_previous_period,
)

logger = logging.getLogger(__name__)


class WordReportError(RuntimeError):
    """Raised when the Word report cannot be generated at all."""


# Order the per-service cost rows should appear in when the ``cost`` dict does
# not carry an explicit ``per_service`` list (live-mode fallback). The Word
# report renders whatever ``per_service`` rows it is given, so this only drives
# the fallback that derives rows from ``service_breakdown``.
_FALLBACK_SERVICE_ORDER = [
    "Relational Database Service",
    "Elastic Load Balancing",
    "CloudWatch",
    "WAF",
    "Inspector",
    "GuardDuty",
    "S3",
    "Lambda",
]


class WordReport:
    """
    Orchestrates generation of the Fly91 Weekly Status Report Word document.

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
        # Preparer org used in the Disclaimer. The reference sample reads
        # "Operisoft"; for Fly91/Greatworx the default is "Greatworx".
        "disclaimer_org": "Greatworx",
        "word_output_filename": "Weekly_Status_Report.docx",
    }

    # Identity row shown as the first body row of the Cost Summary table.
    _ACCOUNT_NAME = "JUST UDO AVIATION PRIVATE LIMITED (Fly91)/ Greatworx"
    _ACCOUNT_ID = "674351849978"

    # Default logical-name -> path mapping for the configurable asset images.
    # Any entry whose file does not exist on disk is skipped gracefully.
    DEFAULT_ASSETS = {
        "logo": "assets/logo.png",
        "cover_background": "assets/cover_background.png",
        "cost_overview": "assets/screenshots_cost_overview.png",
        "inspector_critical": "assets/screenshots_inspector_critical.png",
        "inspector_high": "assets/screenshots_inspector_high.png",
        "inspector_medium": "assets/screenshots_inspector_medium.png",
        "inspector_low": "assets/screenshots_inspector_low.png",
        "guardduty": "assets/screenshots_guardduty.png",
    }

    # Professional blue used for the cover accent bar.
    _ACCENT_BLUE = "1F4E79"

    # Repo root (two levels up from report/word_report.py) used to resolve
    # relative asset paths regardless of the process working directory.
    _REPO_ROOT = Path(__file__).resolve().parent.parent

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
        # Configurable asset image mapping (logical name -> path). Starts from
        # the defaults and is overridden by config['word_report']['assets'] so
        # a client can swap in their own logo / screenshots with no code change.
        assets = dict(self.DEFAULT_ASSETS)
        cfg_assets = (config.get("word_report", {}) or {}).get("assets", {}) or {}
        if isinstance(cfg_assets, dict):
            assets.update({k: v for k, v in cfg_assets.items() if v})
        self.assets = assets
        self.start_date = start_date
        self.end_date = end_date

    # ------------------------------------------------------------------
    # Asset resolution
    # ------------------------------------------------------------------

    def _asset_path(self, name: str) -> Optional[str]:
        """Resolve a configured asset to an on-disk path, or None.

        Returns the absolute path to the mapped asset if (and only if) the file
        actually exists, so a missing asset degrades gracefully (the caller
        simply skips the image) instead of crashing the report. Relative paths
        are resolved against the repo root.
        """
        rel = self.assets.get(name)
        if not rel:
            return None
        try:
            candidate = Path(rel)
            if not candidate.is_absolute():
                candidate = self._REPO_ROOT / candidate
            if candidate.is_file():
                return str(candidate)
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("Failed to resolve asset '%s' (%s): %s", name, rel, exc)
        return None

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

        logger.info("Creating Word document ...")
        doc = Document()

        self._apply_page_setup(doc)
        self._apply_default_font(doc)

        doc.core_properties.title = self.word_cfg["report_title"]
        doc.core_properties.author = self.word_cfg["submitter_org"]
        doc.core_properties.company = self.word_cfg["client_org"]

        sections = [
            ("Cover", self._build_cover, {"doc": doc}),
            ("Cost Summary Differences", self._build_cost_summary_differences,
             {"doc": doc, "cost": collected_data.get("cost")}),
            ("Disclaimer", self._build_disclaimer, {"doc": doc}),
            ("Contents", self._build_contents, {"doc": doc}),
            ("BAU Matrix Overview", self._build_bau_overview,
             {"doc": doc, "collected_data": collected_data}),
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

    def _iso_range_str(self) -> str:
        if self.start_date and self.end_date:
            return "%s to %s" % (self.start_date.isoformat(),
                                 self.end_date.isoformat())
        return "N/A"

    def _period_days(self) -> int:
        if self.start_date and self.end_date:
            return (self.end_date - self.start_date).days + 1
        return 7

    # ------------------------------------------------------------------
    # Page setup / default font
    # ------------------------------------------------------------------

    # Chosen reference-aligned margins (inches). The reference first-section
    # pgMar (twips) is top=1340 right=0 bottom=1160 left=360 header=0 footer=970.
    # We keep sane non-zero margins close to those while keeping the 6-column
    # cost table inside US Letter's printable width:
    #   left  = 0.25in  (~360 twips, matches reference left)
    #   right = 0.3in   (reference uses 0; keep a small gutter so text is not
    #                    clipped by printers)
    #   top   = 0.9in   (~1296 twips, close to reference top=1340)
    #   bottom= 0.8in   (~1152 twips, close to reference bottom=1160)
    _MARGIN_LEFT_IN = 0.25
    _MARGIN_RIGHT_IN = 0.3
    _MARGIN_TOP_IN = 0.9
    _MARGIN_BOTTOM_IN = 0.8

    def _apply_page_setup(self, doc) -> None:
        """Set page size to US Letter (12240x15840 twips) and tighten margins
        toward the reference section. Wrapped so a failure degrades gracefully.
        """
        try:
            section = doc.sections[0]
            section.page_width = Inches(8.5)   # 12240 twips
            section.page_height = Inches(11)   # 15840 twips
            section.left_margin = Inches(self._MARGIN_LEFT_IN)
            section.right_margin = Inches(self._MARGIN_RIGHT_IN)
            section.top_margin = Inches(self._MARGIN_TOP_IN)
            section.bottom_margin = Inches(self._MARGIN_BOTTOM_IN)
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("Failed to apply page setup: %s", exc)

    def _apply_default_font(self, doc) -> None:
        """Set a modest default body font/size approximating the reference."""
        try:
            normal = doc.styles["Normal"]
            normal.font.name = "Calibri"
            normal.font.size = Pt(10)
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("Failed to apply default font: %s", exc)

    # ------------------------------------------------------------------
    # Section 0: professional cover page
    # ------------------------------------------------------------------

    def _build_cover(self, doc) -> None:
        # (d) Optional full-page background photo behind the content. python-docx
        # has no true full-bleed API; we anchor a behindDoc floating drawing
        # sized to the page. If the asset is absent (the default - the reference
        # set ships NO cityscape) we render the cover without a photo.
        bg_path = self._asset_path("cover_background")
        if bg_path:
            try:
                self._add_full_page_background(doc, bg_path)
            except Exception as exc:
                logger.warning("Failed to add cover background: %s", exc)

        # (a) + (b): a top row with a blue left accent bar and the logo top-right.
        # Use a 1-row, 2-col borderless table: col0 = thin shaded blue bar,
        # col1 = right-aligned logo.
        try:
            header_tbl = doc.add_table(rows=1, cols=2)
            header_tbl.autofit = False
            bar_cell = header_tbl.rows[0].cells[0]
            logo_cell = header_tbl.rows[0].cells[1]
            # Thin blue accent bar.
            try:
                bar_cell.width = Inches(0.25)
                logo_cell.width = Inches(7.4)
            except Exception:
                pass
            self._shade_cell(bar_cell, self._ACCENT_BLUE)
            self._set_cell_vertical_bar_height(header_tbl.rows[0], bar_cell)

            logo_para = logo_cell.paragraphs[0]
            logo_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            logo_path = self._asset_path("logo")
            if logo_path:
                try:
                    run = logo_para.add_run()
                    run.add_picture(logo_path, width=Inches(1.8))
                except Exception as exc:
                    logger.warning("Failed to embed cover logo: %s", exc)
        except Exception as exc:
            logger.warning("Failed to build cover header: %s", exc)

        # (c) Title block.
        for _ in range(4):
            doc.add_paragraph()

        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run(self.word_cfg["report_title"])
        run.bold = True
        run.font.size = Pt(32)
        try:
            run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        except Exception:
            pass

        org = doc.add_paragraph()
        org.alignment = WD_ALIGN_PARAGRAPH.CENTER
        org_run = org.add_run(self.word_cfg["client_org"])
        org_run.bold = True
        org_run.font.size = Pt(16)

        label = doc.add_paragraph()
        label.alignment = WD_ALIGN_PARAGRAPH.CENTER
        label.add_run("%s %s" % (self.word_cfg["submitted_by_label"],
                                 self.word_cfg["submitter_org"]))

        period_p = doc.add_paragraph()
        period_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        period_run = period_p.add_run("Reporting Period: %s" % self._current_range_str())
        period_run.italic = True

        # Cover ends with a page break so Cost Summary starts on page 2.
        doc.add_page_break()

    def _add_full_page_background(self, doc, image_path: str) -> None:
        """Place *image_path* as a behind-text floating drawing sized to the
        full page. python-docx limitation: this is NOT a true full-bleed
        print background; it is an anchored wp:anchor drawing with behindDoc=1
        sized to the Letter page. Documented in the README.
        """
        # First embed the picture inline (so the image part + relationship are
        # created), then convert that inline drawing into a behindDoc anchor.
        para = doc.add_paragraph()
        run = para.add_run()
        pic = run.add_picture(image_path, width=Inches(8.5), height=Inches(11))
        inline = run._r.find(qn("w:drawing"))[0]  # wp:inline
        # Build a wp:anchor element reusing the inline's extent/graphic.
        anchor = OxmlElement("wp:anchor")
        for attr, val in (
            ("behindDoc", "1"), ("distT", "0"), ("distB", "0"),
            ("distL", "0"), ("distR", "0"), ("simplePos", "0"),
            ("locked", "0"), ("layoutInCell", "1"), ("allowOverlap", "1"),
            ("relativeHeight", "0"),
        ):
            anchor.set(attr, val)

        simple_pos = OxmlElement("wp:simplePos")
        simple_pos.set("x", "0")
        simple_pos.set("y", "0")
        anchor.append(simple_pos)

        pos_h = OxmlElement("wp:positionH")
        pos_h.set("relativeFrom", "page")
        align_h = OxmlElement("wp:align")
        align_h.text = "center"
        pos_h.append(align_h)
        anchor.append(pos_h)

        pos_v = OxmlElement("wp:positionV")
        pos_v.set("relativeFrom", "page")
        align_v = OxmlElement("wp:align")
        align_v.text = "center"
        pos_v.append(align_v)
        anchor.append(pos_v)

        # Reuse extent / docPr / graphic from the inline drawing.
        for tag in ("wp:extent", "wp:effectExtent", "wp:docPr",
                    "a:graphic"):
            child = inline.find(qn(tag))
            if child is not None:
                anchor.append(child)

        # wrapNone so it sits behind the text.
        wrap_none = OxmlElement("wp:wrapNone")
        # Insert wrapNone before docPr per schema order (after effectExtent).
        docpr = anchor.find(qn("wp:docPr"))
        if docpr is not None:
            docpr.addprevious(wrap_none)
        else:
            anchor.append(wrap_none)

        drawing = run._r.find(qn("w:drawing"))
        drawing.remove(inline)
        drawing.append(anchor)

    # ------------------------------------------------------------------
    # Section 1: Cost Summary Differences
    # ------------------------------------------------------------------

    def _per_service_rows(self, cost: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Return the per-service cost rows for the cost table.

        Prefers an explicit ``per_service`` list ({service, last_week,
        current_week}). Falls back to deriving rows from ``service_breakdown``
        (service / cost / difference) so the live path still renders a table.
        """
        per_service = cost.get("per_service")
        if isinstance(per_service, list) and per_service:
            return per_service

        rows: List[Dict[str, Any]] = []
        breakdown = cost.get("service_breakdown") or cost.get("top_services") or []
        for svc in breakdown:
            if not isinstance(svc, dict):
                continue
            curr = svc.get("cost")
            diff = svc.get("difference")
            last = None
            if isinstance(curr, (int, float)) and isinstance(diff, (int, float)):
                last = round(curr - diff, 2)
            rows.append({
                "service": svc.get("service", "Unknown"),
                "last_week": last,
                "current_week": curr,
            })
        return rows

    def _build_cost_summary_differences(self, doc, cost: Optional[Dict[str, Any]]) -> None:
        # Centered + underlined title (matches the reference screenshot).
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_p.add_run("Cost Summary Differences")
        title_run.bold = True
        title_run.underline = True
        title_run.font.size = Pt(16)

        cost = cost or {}

        prev_range = self._previous_range_str()
        curr_range = self._current_range_str()

        table = doc.add_table(rows=1, cols=6)
        table.style = "Table Grid"

        # Build the header row with rich (multi-line / bold) cells.
        hdr = table.rows[0].cells
        self._stacked_no_header(hdr[0])
        self._simple_bold_header(hdr[1], "Account Name")
        self._simple_bold_header(hdr[2], "Account ID")
        self._simple_bold_header(hdr[3], "Services")
        self._two_line_bold_header(hdr[4], "Last Week", prev_range)
        self._two_line_bold_header(hdr[5], "Current Week", curr_range)

        rows = self._per_service_rows(cost)

        # Build one row per service first; put the account-identity values only
        # in the FIRST service row, then vertically merge cols 0..2 down the
        # whole service list so the identity reads as one tall cell.
        service_cell_rows = []
        for row in rows:
            svc = str(row.get("service", ""))
            last = row.get("last_week")
            curr = row.get("current_week")
            cells = table.add_row().cells
            cells[3].text = svc
            cells[4].text = self._cost_cell(last)
            cells[5].text = self._cost_cell(curr)
            self._center_cell(cells[4])
            self._center_cell(cells[5])
            service_cell_rows.append(cells)

        if service_cell_rows:
            first_cells = service_cell_rows[0]
            first_cells[0].text = "1"
            first_cells[1].text = self._account_name()
            first_cells[2].text = self._account_id()
            # Vertical merge of the No / Account Name / Account ID columns.
            try:
                last_cells = service_cell_rows[-1]
                for col in (0, 1, 2):
                    if len(service_cell_rows) > 1:
                        first_cells[col].merge(last_cells[col])
                    self._center_cell(first_cells[col])
            except Exception as exc:
                logger.warning("Failed to vertically merge identity cells: %s", exc)

        # Total Cost row.
        total_last_label = cost.get("total_last_week_label")
        total_curr_label = cost.get("total_current_week_label")
        total_last = cost.get("total_last_week", cost.get("previous_week_total"))
        total_curr = cost.get("total_current_week", cost.get("current_week_total"))
        if not total_last_label:
            total_last_label = "(%s Tax Excluded Cost)" % format_currency(total_last)
        if not total_curr_label:
            total_curr_label = "(%s Tax Excluded Cost)" % format_currency(total_curr)

        total_cells = table.add_row().cells
        # Merge the leading columns (No..Services -> indices 0..3) into one cell.
        try:
            lead = total_cells[0]
            for idx in (1, 2, 3):
                lead = lead.merge(total_cells[idx])
            lead.text = "Total Cost"
            self._bold_cell(lead)
            self._center_cell(lead)
        except Exception as exc:
            logger.warning("Failed to merge Total Cost leading cells: %s", exc)
            total_cells[3].text = "Total Cost"
            self._bold_cell(total_cells[3])

        total_cells[4].text = str(total_last_label)
        total_cells[5].text = str(total_curr_label)
        for idx in (4, 5):
            self._bold_cell(total_cells[idx])
            self._center_cell(total_cells[idx])

        self._cost_analysis_bullets(doc, cost, total_last, total_curr, rows)

        # Optional Cost Explorer overview capture (configurable asset). Embedded
        # only if the asset exists on disk; skipped gracefully otherwise.
        overview = self._asset_path("cost_overview")
        if overview:
            try:
                doc.add_picture(overview, width=Inches(6))
                cap = doc.add_paragraph()
                cap_run = cap.add_run(
                    "Cost Explorer overview - %s" % self._current_range_str())
                cap_run.italic = True
                cap_run.font.size = Pt(8)
            except Exception as exc:
                logger.warning("Failed to embed cost overview asset: %s", exc)

    # ---- Cost-table header cell builders -----------------------------

    @staticmethod
    def _simple_bold_header(cell, text: str) -> None:
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        run.bold = True
        try:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        except Exception:
            pass

    @staticmethod
    def _stacked_no_header(cell) -> None:
        """Render the 'No' header stacked as 'N' over 'o' (a line break)."""
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run("N")
        run.bold = True
        run.add_break()
        run2 = para.add_run("o")
        run2.bold = True
        try:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        except Exception:
            pass

    @staticmethod
    def _two_line_bold_header(cell, line1: str, line2: str) -> None:
        """Bold header with the label on line 1 and the date range on line 2."""
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(line1)
        run.bold = True
        run.add_break()
        run2 = para.add_run("(%s)" % line2)
        run2.bold = True
        run2.font.size = Pt(9)
        try:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        except Exception:
            pass

    def _cost_analysis_bullets(self, doc, cost: Dict[str, Any],
                               total_last: Any, total_curr: Any,
                               rows: Sequence[Dict[str, Any]]) -> None:
        """Render the cost-analysis bullets below the Cost Summary table.

        Values are computed from the data where feasible (difference,
        percentage, average daily cost, per-service decreases/increases) so the
        bullets stay consistent with the table even if the numbers change.
        """
        difference = cost.get("difference")
        if not isinstance(difference, (int, float)):
            if isinstance(total_curr, (int, float)) and isinstance(total_last, (int, float)):
                difference = round(total_curr - total_last, 2)
            else:
                difference = None

        pct = cost.get("pct_change")
        if not isinstance(pct, (int, float)) and isinstance(difference, (int, float)) \
                and isinstance(total_last, (int, float)) and total_last:
            pct = round((difference / total_last) * 100.0, 2)

        avg_daily = cost.get("avg_daily_cost")
        if not isinstance(avg_daily, (int, float)) and isinstance(total_curr, (int, float)):
            avg_daily = round(total_curr / max(self._period_days(), 1), 2)

        decreased = isinstance(difference, (int, float)) and difference < 0
        direction = "decrease" if decreased else "increase"

        if isinstance(difference, (int, float)):
            doc.add_paragraph(
                "There is a %s of %s compared to the previous week."
                % (direction, format_currency(abs(difference))),
                style="List Bullet",
            )
        if isinstance(pct, (int, float)):
            doc.add_paragraph(
                "The percentage %s compared to the previous week is "
                "approximately %.2f%%." % (direction, abs(pct)),
                style="List Bullet",
            )
        if isinstance(avg_daily, (int, float)):
            doc.add_paragraph(
                "The Weekly Average Daily Cost for the current week is %s per day."
                % format_currency(avg_daily),
                style="List Bullet",
            )

        decreases, increases = self._service_movements(rows)
        if decreases:
            doc.add_paragraph(
                "The overall %s is primarily due to lower usage/cost in: %s"
                % (direction, ", ".join(decreases)),
                style="List Bullet",
            )
        if increases:
            doc.add_paragraph(
                "The %s was partially offset by higher costs in: %s"
                % (direction, ", ".join(increases)),
                style="List Bullet",
            )

    # The reference lists the six largest movers in each direction. Cap the
    # analysis bullets to this many so they match the reference prose rather
    # than enumerating every tiny per-service delta.
    _MAX_MOVEMENTS = 6

    @staticmethod
    def _service_movements(rows: Sequence[Dict[str, Any]]):
        """Split services into (decreases, increases) formatted strings.

        Each entry is "<service>: -$x.xx" / "+$x.xx", ordered by magnitude of
        change (largest first), skipping the Tax row and zero-change rows. Each
        list is capped to the ``_MAX_MOVEMENTS`` largest movers so the bullets
        match the reference's six-service decrease/increase lists.
        """
        movements = []
        for row in rows:
            svc = str(row.get("service", ""))
            if svc.lower() == "tax":
                continue
            last = row.get("last_week")
            curr = row.get("current_week")
            if not isinstance(last, (int, float)) or not isinstance(curr, (int, float)):
                continue
            delta = round(curr - last, 2)
            if abs(delta) < 0.005:
                continue
            movements.append((svc, delta))

        decreases = sorted((m for m in movements if m[1] < 0),
                           key=lambda m: m[1])[:WordReport._MAX_MOVEMENTS]
        increases = sorted((m for m in movements if m[1] > 0),
                           key=lambda m: m[1], reverse=True)[:WordReport._MAX_MOVEMENTS]

        def _fmt(svc, delta):
            sign = "-" if delta < 0 else "+"
            return "%s: %s$%.2f" % (svc, sign, abs(delta))

        return (
            [_fmt(s, d) for s, d in decreases],
            [_fmt(s, d) for s, d in increases],
        )

    def _account_name(self) -> str:
        return (self.config.get("word_report", {}) or {}).get(
            "cost_account_name", self._ACCOUNT_NAME)

    def _account_id(self) -> str:
        return str(self.config.get("aws_account_id") or self._ACCOUNT_ID)

    @staticmethod
    def _cost_cell(value: Any) -> str:
        """Render a per-service cost cell. None -> em dash (reference Tax row)."""
        if value is None:
            return "\u2014"  # em dash, matching the reference Tax current-week
        if isinstance(value, (int, float)):
            return format_currency(value)
        return str(value)

    # ------------------------------------------------------------------
    # Section 2: Disclaimer
    # ------------------------------------------------------------------

    def _build_disclaimer(self, doc) -> None:
        doc.add_heading("Disclaimer", level=1)
        org = self.word_cfg.get("disclaimer_org") or "Greatworx"
        doc.add_paragraph(
            "The contents of this document are based upon the information "
            "provided by the customer. This document is considered confidential "
            "between %s and Customer and may not be distributed to any third "
            "party without %s's prior written consent." % (org, org)
        )

    # ------------------------------------------------------------------
    # Section 3: Contents
    # ------------------------------------------------------------------

    def _build_contents(self, doc) -> None:
        doc.add_heading("Contents", level=1)

        entries = [
            ("Cost summary and projections", 5),
            ("BAU Matrix Overview of the resources", 6),
            ("Uptime of the servers and applications running on the infrastructure", 6),
            ("    1. Elastic Compute Cloud (EC2)", 6),
            ("    2. Relational Database Service (RDS)", 8),
            ("    3. Elastic Load Balancing (ELB)", 9),
            ("    4. AWS WAF", 9),
            ("    5. Amazon Inspector", 10),
            ("    6. Guard Duty", 11),
        ]
        for label, page in entries:
            p = doc.add_paragraph()
            p.add_run("%s %s %d" % (label, "." * max(3, 40 - len(label)), page))

        doc.add_paragraph(
            "Cost summary and projections:- The billing and the Charges for the "
            "Services are also as per the Budget set-up; There are increases in "
            "the estimated billing compared to the previous week.",
            style="List Bullet",
        )
        date_p = doc.add_paragraph()
        date_p.add_run(self._iso_range_str())

    # ------------------------------------------------------------------
    # Section 4: BAU Matrix Overview
    # ------------------------------------------------------------------

    def _build_bau_overview(self, doc, collected_data: Dict[str, Any]) -> None:
        doc.add_heading("BAU Matrix Overview of the resources.", level=1)
        subtitle = doc.add_paragraph()
        sub_run = subtitle.add_run(
            "Uptime of the servers and applications running on the infrastructure"
        )
        sub_run.bold = True

        self._build_ec2(doc, collected_data.get("ec2"))
        self._build_rds(doc, collected_data.get("rds"))
        self._build_elb(doc, collected_data.get("elb"))
        self._build_waf(doc, collected_data.get("waf"))
        self._build_inspector(doc, collected_data.get("inspector"))
        self._build_guardduty(doc, collected_data.get("guardduty"))

    # ---- 1. EC2 -------------------------------------------------------

    def _build_ec2(self, doc, ec2: Any) -> None:
        doc.add_heading("1. Elastic Compute Cloud (EC2)", level=2)
        instances = ec2 if isinstance(ec2, list) else []
        running = [i for i in instances if isinstance(i, dict)
                   and (i.get("state") or "").lower() == "running"]
        stopped = [i for i in instances if isinstance(i, dict)
                   and (i.get("state") or "").lower() == "stopped"]

        doc.add_paragraph(
            "We have %d EC2 Instances deployed in our AWS Account currently."
            % len(running),
            style="List Bullet",
        )
        for inst in stopped:
            name = inst.get("name", "N/A")
            iid = inst.get("instance_id", "")
            stopped_at = inst.get("stopped_at_display") or inst.get("stopped_at")
            if stopped_at:
                doc.add_paragraph(
                    "%s (%s) has been put to \"Stopped\" state from %s."
                    % (name, iid, stopped_at),
                    style="List Bullet",
                )
            else:
                doc.add_paragraph(
                    "%s (%s) has been put to \"Stopped\" state." % (name, iid),
                    style="List Bullet",
                )
        if stopped:
            names = ", ".join(
                "%s (%s)" % (i.get("name", "N/A"), i.get("instance_id", ""))
                for i in stopped
            )
            doc.add_paragraph(
                "In this last week there was no Incident of any downtime for the "
                "production servers, all the servers were up and in running "
                "state, except %s." % names,
                style="List Bullet",
            )
        else:
            doc.add_paragraph(
                "In this last week there was no Incident of any downtime for the "
                "production servers, all the servers were up and in running state.",
                style="List Bullet",
            )

        util_p = doc.add_paragraph()
        util_p.add_run(
            "All EC2 Server Utilization as mentioned in the table below (%s)"
            % self._current_range_str()
        ).bold = True

        util_headers = [
            "Name", "InstanceID", "Instance type",
            "Minimum CPU utilization", "Maximum CPU utilization",
            "Average CPU utilization", "Average Memory Utilization",
            "Current Disk Utilization",
        ]
        table = doc.add_table(rows=1, cols=len(util_headers))
        table.style = "Table Grid"
        self._set_header_row(table.rows[0], util_headers)
        for inst in instances:
            if not isinstance(inst, dict):
                continue
            cells = table.add_row().cells
            cells[0].text = str(inst.get("name", "-"))
            cells[1].text = str(inst.get("instance_id", "-"))
            cells[2].text = str(inst.get("instance_type", "-"))
            cells[3].text = self._num_cell(inst.get("cpu_min"))
            cells[4].text = self._num_cell(inst.get("cpu_max"))
            cells[5].text = self._num_cell(inst.get("cpu_avg"))
            cells[6].text = self._mem_cell(inst)
            cells[7].text = self._disk_cell(inst.get("disk_utilization"))

        net_p = doc.add_paragraph()
        net_p.add_run("Bandwidth and Network stats (Max)").bold = True

        net_headers = [
            "Name", "InstanceID", "Network in (bytes)", "Network Out (bytes)",
            "Network packets in (count)", "Network packets out (count)",
        ]
        net_table = doc.add_table(rows=1, cols=len(net_headers))
        net_table.style = "Table Grid"
        self._set_header_row(net_table.rows[0], net_headers)
        for inst in instances:
            if not isinstance(inst, dict):
                continue
            cells = net_table.add_row().cells
            cells[0].text = str(inst.get("name", "-"))
            cells[1].text = str(inst.get("instance_id", "-"))
            cells[2].text = self._bytes_cell(inst.get("network_in"))
            cells[3].text = self._bytes_cell(inst.get("network_out"))
            cells[4].text = self._count_cell(inst.get("network_packets_in"))
            cells[5].text = self._count_cell(inst.get("network_packets_out"))

        # Optional per-instance metric chart (only if metric_series present).
        self._embed_ec2_charts(doc, instances)

    def _embed_ec2_charts(self, doc, instances: Sequence[Dict[str, Any]]) -> None:
        """Embed an optional CloudWatch-style chart per instance that carries
        a non-empty ``metric_series``. The reference is tabular, so this is a
        no-op unless live data supplies time-series. Any failure is logged and
        skipped so the tables/text remain."""
        thresholds = self.config.get("thresholds", {}) or {}
        for inst in instances:
            if not isinstance(inst, dict):
                continue
            series_list = inst.get("metric_series") or []
            if not series_list:
                continue
            name = str(inst.get("name", ""))
            try:
                buf = metric_charts.render_metric_chart(
                    name, series_list, thresholds=thresholds
                )
                if buf is None:
                    continue
                doc.add_picture(buf, width=Inches(6))
                caption = doc.add_paragraph()
                cap_run = caption.add_run(
                    "CloudWatch metrics for %s - %s"
                    % (name, self._current_range_str())
                )
                cap_run.italic = True
                cap_run.font.size = Pt(8)
            except Exception as exc:
                logger.warning("Failed to embed EC2 metric chart for '%s': %s",
                               name, exc)

    # ---- 2. RDS -------------------------------------------------------

    def _build_rds(self, doc, rds: Any) -> None:
        doc.add_heading("2. Relational Database Service (RDS)", level=2)
        instances = rds if isinstance(rds, list) else []

        doc.add_paragraph(
            "We have been utilizing %d RDS Database Instances in our "
            "infrastructure (Production, Dev and UAT)." % len(instances),
            style="List Bullet",
        )
        doc.add_paragraph(
            "Throughout this whole week there was No downtime or any "
            "Unauthorized incident in the RDS Service.",
            style="List Bullet",
        )

        util_headers = [
            "RDS Name", "Down Time", "Instance type", "Minimum Utilization",
            "Maximum Utilization", "Average Utilization", "Free Memory",
            "Free Storage",
        ]
        table = doc.add_table(rows=1, cols=len(util_headers))
        table.style = "Table Grid"
        self._set_header_row(table.rows[0], util_headers)
        for inst in instances:
            if not isinstance(inst, dict):
                continue
            cells = table.add_row().cells
            cells[0].text = str(inst.get("name", "-"))
            cells[1].text = str(inst.get("down_time", "No"))
            cells[2].text = str(inst.get("instance_type", "-"))
            cells[3].text = self._str_cell(inst.get("min_utilization"))
            cells[4].text = self._str_cell(inst.get("max_utilization"))
            cells[5].text = self._str_cell(inst.get("avg_utilization"))
            cells[6].text = self._str_cell(inst.get("free_memory"))
            cells[7].text = self._str_cell(inst.get("free_storage"))

        net_p = doc.add_paragraph()
        net_p.add_run("Bandwidth and Network status").bold = True

        net_headers = [
            "RDS Name",
            "Network Transmit Throughput (Bytes per second)",
            "Network Receive Throughput (Bytes per second)",
            "Max Database Connection (Count)",
        ]
        net_table = doc.add_table(rows=1, cols=len(net_headers))
        net_table.style = "Table Grid"
        self._set_header_row(net_table.rows[0], net_headers)
        for inst in instances:
            if not isinstance(inst, dict):
                continue
            cells = net_table.add_row().cells
            cells[0].text = str(inst.get("name", "-"))
            cells[1].text = self._str_cell(inst.get("network_transmit_throughput"))
            cells[2].text = self._str_cell(inst.get("network_receive_throughput"))
            cells[3].text = self._str_cell(inst.get("max_db_connections"))

        doc.add_paragraph(
            "Due to recent changes in the AWS Console, we are currently unable "
            "to retrieve the maximum weekly data for Network Transmit/Receive "
            "Throughput.",
            style="List Bullet",
        )

    # ---- 3. ELB -------------------------------------------------------

    def _build_elb(self, doc, elb: Any) -> None:
        doc.add_heading("3. Elastic Load Balancing (ELB)", level=2)
        albs = elb if isinstance(elb, list) else []

        doc.add_paragraph(
            "We have been utilizing %d Load Balancer in our infrastructure "
            "(Production, Dev and Development). Utilization (Sum):" % len(albs),
            style="List Bullet",
        )

        headers = [
            "ALB Name", "Requests", "Active connection count",
            "New connection count", "Consumed LoadBalancer Capacity Units",
            "HTTP redirect count", "Processed Bytes",
            "Target Response Time (AVG)",
        ]
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = "Table Grid"
        self._set_header_row(table.rows[0], headers)
        for alb in albs:
            if not isinstance(alb, dict):
                continue
            cells = table.add_row().cells
            cells[0].text = str(alb.get("name", "-"))
            cells[1].text = self._count_cell(alb.get("requests"))
            cells[2].text = self._count_cell(alb.get("active_connections"))
            cells[3].text = self._count_cell(alb.get("new_connections"))
            cells[4].text = self._num_cell(alb.get("consumed_lcus"))
            cells[5].text = self._count_cell(alb.get("http_redirect_count"))
            cells[6].text = self._bytes_cell(alb.get("processed_bytes"))
            cells[7].text = self._response_time_cell(alb.get("target_response_time"))

    # ---- 4. WAF -------------------------------------------------------

    def _build_waf(self, doc, waf: Any) -> None:
        doc.add_heading("4. AWS WAF", level=2)
        acls = waf if isinstance(waf, list) else []

        doc.add_paragraph(
            "We have %d web ACL WAF in Mumbai region, 1 production, 1 staging."
            % len(acls),
            style="List Bullet",
        )
        req_p = doc.add_paragraph()
        req_p.add_run(
            "The requests data for the WAF is below (%s):" % self._current_range_str()
        )

        headers = ["WAF Name", "Total Request", "Blocked Request", "Allowed Request"]
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = "Table Grid"
        self._set_header_row(table.rows[0], headers)
        for acl in acls:
            if not isinstance(acl, dict):
                continue
            cells = table.add_row().cells
            cells[0].text = str(acl.get("name", "-"))
            cells[1].text = self._count_cell(acl.get("total_requests"))
            cells[2].text = self._count_cell(acl.get("blocked_requests"))
            cells[3].text = self._count_cell(acl.get("allowed_requests"))

    # ---- 5. Amazon Inspector -----------------------------------------

    def _build_inspector(self, doc, inspector: Any) -> None:
        doc.add_heading("5. Amazon Inspector", level=2)
        data = inspector if isinstance(inspector, dict) else {}
        counts = data.get("severity_counts", {}) or {}
        total = data.get("total_findings")

        # Prefer a reference-shaped summary sentence supplied by the data
        # (so the wording can read "400+ ... 200+ High, 200+ medium" exactly
        # like the reference) and fall back to deriving it from the numeric
        # severity_counts the Excel Inspector sheet consumes.
        summary = data.get("summary_text")
        if not summary:
            display = data.get("severity_display") or {}
            critical = self._severity_display(counts, display, "CRITICAL")
            high = self._severity_display(counts, display, "HIGH")
            medium = self._severity_display(counts, display, "MEDIUM")
            low = self._severity_display(counts, display, "LOW")
            untriaged = self._severity_display(counts, display, "UNTRIAGED")
            total_str = self._total_display(data, "total_display", total)
            summary = (
                "There are %s findings for the last week categorized as follows: "
                "%s Critical, %s High, %s medium and %s low and %s untriaged."
                % (total_str, critical, high, medium, low, untriaged)
            )
        doc.add_paragraph(summary)

        self._embed_inspector_screenshots(doc, data)

    @staticmethod
    def _severity_display(counts: Dict[str, Any], display: Dict[str, Any],
                          key: str) -> str:
        """Display string for one severity bucket.

        Prefers an explicit per-bucket display string carried in the data's
        ``severity_display`` map (so the reference's "200+" qualifier is
        honoured), otherwise renders the raw numeric count from
        ``severity_counts``."""
        if isinstance(display, dict) and display.get(key) is not None:
            return str(display[key])
        return str(int(counts.get(key, 0) or 0))

    @staticmethod
    def _total_display(data: Dict[str, Any], display_key: str, total: Any) -> str:
        """Display string for the grand total of findings.

        Prefers an explicit display string (reference "400+") carried on the
        data, otherwise falls back to "<n>+" derived from the raw total."""
        explicit = data.get(display_key)
        if explicit:
            return str(explicit)
        if isinstance(total, (int, float)) and total:
            return "%d+" % total
        return "many"

    def _embed_inspector_screenshots(self, doc, data: Dict[str, Any]) -> None:
        """Embed Inspector console-style screenshots grouped by severity.

        One image per non-empty severity group (High / Medium / Low /
        Critical). Falls back to a single all-findings image if no per-finding
        detail is available. Skips gracefully if Pillow is unavailable."""
        findings = data.get("findings") or []
        date_range = data.get("date_range", "")
        headers = ["Severity", "Title", "Resource", "CVE", "First Observed"]

        groups = [("CRITICAL", "Critical"), ("HIGH", "High"),
                  ("MEDIUM", "Medium"), ("LOW", "Low")]

        # Prefer the configured console-capture asset images (real screenshots)
        # when present, so the report looks like the original reference. Falls
        # back to the runtime Pillow renderer below if no asset is configured.
        asset_map = {
            "CRITICAL": "inspector_critical",
            "HIGH": "inspector_high",
            "MEDIUM": "inspector_medium",
            "LOW": "inspector_low",
        }
        asset_embedded = 0
        for sev_key, sev_label in groups:
            asset_path = self._asset_path(asset_map[sev_key])
            if not asset_path:
                continue
            try:
                doc.add_picture(asset_path, width=Inches(6))
                cap = doc.add_paragraph()
                cap_run = cap.add_run("%s Findings" % sev_label)
                cap_run.italic = True
                cap_run.font.size = Pt(8)
                asset_embedded += 1
            except Exception as exc:
                logger.warning("Failed to embed Inspector asset '%s': %s",
                               asset_path, exc)
        if asset_embedded:
            return

        embedded = 0
        if findings:
            for sev_key, sev_label in groups:
                rows = [
                    [
                        (f.get("severity", "-") or "-").upper(),
                        f.get("title", "-"),
                        f.get("resource_id", "-") or "-",
                        f.get("cve", "-"),
                        f.get("first_observed", "-"),
                    ]
                    for f in findings
                    if isinstance(f, dict)
                    and (f.get("severity", "") or "").upper() == sev_key
                ]
                if not rows:
                    continue
                title = "Amazon Inspector - %s Findings" % sev_label
                if date_range and date_range != "N/A":
                    title = "%s  |  %s" % (title, date_range)
                if self._embed_screenshot(doc, title, headers, rows, severity_col=0):
                    cap = doc.add_paragraph()
                    cap_run = cap.add_run("%s Findings" % sev_label)
                    cap_run.italic = True
                    cap_run.font.size = Pt(8)
                    embedded += 1

        if embedded == 0:
            # Fallback: a single synthetic summary screenshot so the section
            # still carries a console-style visual when per-finding detail is
            # unavailable (reference-count based).
            counts = data.get("severity_counts", {}) or {}
            rows = [
                ["CRITICAL", "Critical findings", "AWS account", "-", ""],
                ["HIGH", "High findings", "AWS account", "-", ""],
                ["MEDIUM", "Medium findings", "AWS account", "-", ""],
                ["LOW", "Low findings", "AWS account", "-", ""],
            ]
            # Only keep rows whose severity has a non-zero count.
            key_map = {"CRITICAL": "CRITICAL", "HIGH": "HIGH",
                       "MEDIUM": "MEDIUM", "LOW": "LOW"}
            rows = [r for r in rows if int(counts.get(key_map[r[0]], 0) or 0) > 0]
            if not rows:
                rows = [["LOW", "Findings", "AWS account", "-", ""]]
            title = "Amazon Inspector - Findings Overview"
            if date_range and date_range != "N/A":
                title = "%s  |  %s" % (title, date_range)
            self._embed_screenshot(doc, title, headers, rows, severity_col=0)

    # ---- 6. Guard Duty -----------------------------------------------

    def _build_guardduty(self, doc, guardduty: Any) -> None:
        doc.add_heading("6. Guard Duty", level=2)
        data = guardduty if isinstance(guardduty, dict) else {}
        counts = data.get("severity_counts", {}) or {}
        total = data.get("total_findings")

        high = int(counts.get("HIGH", 0) or 0)
        medium = int(counts.get("MEDIUM", 0) or 0)
        low = int(counts.get("LOW", 0) or 0)

        total_str = str(int(total)) if isinstance(total, (int, float)) else "several"
        doc.add_paragraph(
            "There are %s new findings this week, categorized as %d lows, "
            "%d Medium and %d High as displayed in the console."
            % (total_str, low, medium, high)
        )

        # Prefer the configured GuardDuty console-capture asset image when
        # present (looks like the original reference); otherwise fall back to
        # the runtime Pillow renderer.
        gd_asset = self._asset_path("guardduty")
        if gd_asset:
            try:
                doc.add_picture(gd_asset, width=Inches(6))
                return
            except Exception as exc:
                logger.warning("Failed to embed GuardDuty asset '%s': %s",
                               gd_asset, exc)

        findings = data.get("findings") or []
        date_range = data.get("date_range", "")
        headers = ["Severity", "Finding Type", "Title", "Resource", "Count"]
        rows = [
            [
                (f.get("severity_label", "-") or "-").upper(),
                f.get("type", "-"),
                f.get("title", "-"),
                f.get("resource_type", "-"),
                f.get("count", "-"),
            ]
            for f in findings if isinstance(f, dict)
        ]
        if not rows:
            rows = [["LOW", "GuardDuty finding", "Low severity finding",
                     "Instance", low or 1]]
        title = "Amazon GuardDuty - Findings"
        if date_range and date_range != "N/A":
            title = "%s  |  %s" % (title, date_range)
        self._embed_screenshot(doc, title, headers, rows, severity_col=0)

    def _embed_screenshot(self, doc, title: str, headers: Sequence[str],
                          rows: Sequence[Sequence[Any]],
                          severity_col: Optional[int] = None) -> bool:
        """Render a console-style screenshot via report.screenshots and embed
        it. Returns True if an image was embedded, False if it was skipped
        (Pillow unavailable, render failed, or embed failed)."""
        try:
            buf = screenshots.render_findings_table(
                title, headers, rows, severity_col=severity_col
            )
            if buf is None:
                return False
            doc.add_picture(buf, width=Inches(6))
            return True
        except Exception as exc:
            logger.warning("Failed to embed screenshot '%s': %s", title, exc)
            return False

    # ------------------------------------------------------------------
    # Trailer
    # ------------------------------------------------------------------

    def _build_trailer(self, doc) -> None:
        trailer = doc.add_paragraph()
        trailer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = trailer.add_run("-- End of Document --")
        run.bold = True

    # ------------------------------------------------------------------
    # Cell-value helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _num_cell(value: Any) -> str:
        if value is None:
            return "-"
        if isinstance(value, (int, float)):
            # Trim trailing zeros to match the reference (e.g. 0.59, 6.19, 1.8).
            text = ("%g" % round(float(value), 2))
            return text
        return str(value)

    @staticmethod
    def _str_cell(value: Any) -> str:
        if value is None or value == "":
            return "-"
        return str(value)

    @staticmethod
    def _disk_cell(value: Any) -> str:
        if value is None or value == "":
            return "-"
        return str(value)

    @staticmethod
    def _mem_cell(inst: Dict[str, Any]) -> str:
        """Average memory utilisation display.

        Prefers an explicit ``memory_avg_display`` string (reference figures
        such as "50.62"); falls back to a numeric ``memory_avg``."""
        display = inst.get("memory_avg_display")
        if display is not None and display != "":
            return str(display)
        mem = inst.get("memory_avg")
        if isinstance(mem, (int, float)):
            return "%g" % round(float(mem), 2)
        return "-"

    @staticmethod
    def _bytes_cell(value: Any) -> str:
        if value is None:
            return "-"
        return format_bytes(value, precision=2)

    @staticmethod
    def _count_cell(value: Any) -> str:
        if value is None:
            return "-"
        if isinstance(value, (int, float)):
            # Word tables use an uppercase thousands suffix ("952.35K") to
            # match the reference and stay consistent with the uppercase "M"
            # that format_count already emits. The shared helper keeps its
            # lowercase "k" so the Excel path is unchanged.
            return format_count(value, precision=2).replace("k", "K")
        return str(value)

    @staticmethod
    def _response_time_cell(value: Any) -> str:
        """Render target response time. Seconds in -> '1.109 s' / '322.8 ms'."""
        if value is None:
            return "-"
        if isinstance(value, (int, float)):
            if value < 1:
                return "%.1f ms" % (value * 1000.0)
            return "%.3f s" % value
        return str(value)

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

    @staticmethod
    def _center_cell(cell) -> None:
        """Center all paragraphs in a cell horizontally and vertically."""
        try:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("Failed to center cell: %s", exc)

    @staticmethod
    def _shade_cell(cell, hex_color: str) -> None:
        """Apply a solid background fill to a table cell via w:shd."""
        try:
            tc_pr = cell._tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), hex_color)
            tc_pr.append(shd)
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("Failed to shade cell: %s", exc)

    @staticmethod
    def _set_cell_vertical_bar_height(row, cell) -> None:
        """Give the accent-bar row a tall minimum height so the blue column
        reads as a full-height vertical bar down the left edge of the cover.

        python-docx row height is set via a ``w:trHeight`` on the row's
        ``w:trPr`` (val in twips, hRule=atLeast). The usable page height is the
        Letter page (11in) minus the top (0.9in) and bottom (0.8in) margins,
        i.e. ~9.3in -> 13392 twips (1in = 1440 twips). We also clear any run
        text in the bar cell so only the fill shows.
        """
        try:
            # Clear any text so only the fill shows; keep an empty paragraph.
            for paragraph in cell.paragraphs:
                for run in list(paragraph.runs):
                    run.text = ""
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("Failed to clear accent bar text: %s", exc)
        try:
            # ~9.3in usable height (11in page - 0.9in top - 0.8in bottom).
            usable_twips = int(9.3 * 1440)  # 13392 twips
            tr_pr = row._tr.get_or_add_trPr()
            tr_height = OxmlElement("w:trHeight")
            tr_height.set(qn("w:val"), str(usable_twips))
            tr_height.set(qn("w:hRule"), "atLeast")
            tr_pr.append(tr_height)
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("Failed to set accent bar row height: %s", exc)
