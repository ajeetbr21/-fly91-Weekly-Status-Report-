"""
Amazon Inspector sheet generator for the AWS Weekly BAU Report.

Displays all Amazon Inspector findings for the reporting week, including
a severity summary (Critical / High / Medium / Low / Informational) with
colour-coded counts and a detailed findings table.
"""

from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from report.styles import ReportStyles
from report import screenshots
from utils.logger import get_logger

logger = get_logger("bau_report")


class InspectorSheet:
    """Generates the Amazon Inspector worksheet."""

    SHEET_NAME = "Inspector"

    SUMMARY_HEADERS = ["Severity", "Count"]

    SEVERITY_ROWS = ["CRITICAL", "HIGH", "MEDIUM", "LOW", "INFORMATIONAL"]

    DETAIL_HEADERS = [
        "Title",
        "Severity",
        "Resource",
        "Finding Type",
        "CVE",
        "First Observed",
        "Status",
    ]

    def generate(self, wb, inspector_data, styles=None):
        """
        Create the Amazon Inspector sheet.

        Parameters
        ----------
        wb : openpyxl.Workbook
        inspector_data : dict
            Inspector data from the InspectorCollector.
        styles : ReportStyles class (optional)

        Returns
        -------
        openpyxl.worksheet.worksheet.Worksheet
        """
        styles = styles or ReportStyles
        ws = wb.create_sheet(self.SHEET_NAME)
        ws.sheet_properties.tabColor = ReportStyles.DARK_BLUE

        inspector_data = inspector_data or {}
        severity_counts = inspector_data.get("severity_counts", {}) or {}
        findings = inspector_data.get("findings", []) or []
        date_range = inspector_data.get("date_range", "N/A")
        collection_status = inspector_data.get("collection_status", "ok")
        collection_error = inspector_data.get("collection_error")

        total_cols = len(self.DETAIL_HEADERS)
        row = 1

        # ── Title bar ────────────────────────────────────────────────
        styles.apply_title_bar(
            ws, row, "Amazon Inspector - Findings (One Week)", total_cols
        )
        row += 2

        # ── Date-range section header ────────────────────────────────
        styles.apply_section_header(
            ws, row, f"Reporting Period: {date_range}", total_cols
        )
        row += 2

        # ── Collection-failure banner ────────────────────────────────
        # When the collector could not query Inspector (e.g. the service
        # is not enabled or the IAM role lacks inspector2:ListFindings),
        # make the failure visible instead of showing an empty table that
        # looks identical to a clean week.
        if collection_status == "error":
            row = self._write_collection_error(
                ws, row, total_cols, collection_error
            )

        # ── Severity summary table ────────────────────────────────────
        styles.apply_section_header(ws, row, "Severity Summary", total_cols)
        row += 1

        styles.apply_header_row(ws, row, self.SUMMARY_HEADERS)
        row += 1

        for i, severity in enumerate(self.SEVERITY_ROWS):
            even = i % 2 == 0
            count = int(severity_counts.get(severity, 0) or 0)
            label_cell = styles.apply_data_cell(
                ws, row, 1, severity.title(), even=even
            )
            count_cell = styles.apply_data_cell(ws, row, 2, count, even=even)
            self._apply_severity_fill(severity, label_cell, count_cell)
            row += 1

        # Total findings row
        total_label = styles.apply_data_cell(ws, row, 1, "TOTAL", even=False)
        total_label.font = ReportStyles.FONT_DATA_BOLD
        total_val = styles.apply_data_cell(
            ws, row, 2, int(inspector_data.get("total_findings", len(findings)) or 0),
            even=False,
        )
        total_val.font = ReportStyles.FONT_DATA_BOLD
        row += 2

        # ── Detailed findings table ───────────────────────────────────
        styles.apply_section_header(ws, row, "Detailed Findings", total_cols)
        row += 1

        detail_header_row = row
        styles.apply_header_row(ws, row, self.DETAIL_HEADERS)
        row += 1

        if not findings:
            ws.merge_cells(f"A{row}:{get_column_letter(total_cols)}{row}")
            if collection_status == "error":
                empty_msg = (
                    "Collection failed - findings unavailable (see warning above)."
                )
            else:
                empty_msg = "No Inspector findings for this week (clean)."
            cell = ws.cell(row=row, column=1, value=empty_msg)
            cell.font = ReportStyles.FONT_DATA
            cell.alignment = ReportStyles.ALIGN_CENTER
            row += 1
        else:
            for i, finding in enumerate(findings):
                even = i % 2 == 0
                resource = finding.get("resource_id", "-") or "-"
                row_data = [
                    finding.get("title", "-"),
                    (finding.get("severity", "-") or "-").title(),
                    resource,
                    finding.get("finding_type", "-"),
                    finding.get("cve", "-"),
                    finding.get("first_observed", "-"),
                    finding.get("status", "-"),
                ]
                styles.apply_data_row(ws, row, row_data, even=even)

                # Colour-code the severity cell (column 2)
                sev = (finding.get("severity", "") or "").upper()
                sev_cell = ws.cell(row=row, column=2)
                self._apply_severity_fill(sev, sev_cell)
                row += 1

        # ── Auto-fit, freeze, filter ──────────────────────────────────
        styles.auto_fit_columns(ws, min_width=12)
        styles.freeze_and_filter(ws, detail_header_row)

        # ── Screenshot-style console visual (one-week findings) ────────
        row += 1
        self._embed_findings_screenshot(ws, row, findings, date_range, styles)

        return ws

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _write_collection_error(self, ws, row, total_cols, error):
        """Render a distinct red banner when Inspector collection failed."""
        ws.merge_cells(f"A{row}:{get_column_letter(total_cols)}{row}")
        msg = (
            "⚠ Inspector findings could NOT be collected (collection failed / "
            "access denied). This is NOT a clean week - the result below is "
            "empty because the API call did not succeed. Verify the service is "
            "enabled and the IAM role has inspector2:ListFindings."
        )
        cell = ws.cell(row=row, column=1, value=msg)
        cell.fill = ReportStyles.FILL_RED
        cell.font = Font(name="Calibri", size=10, bold=True, color=ReportStyles.RED_FG)
        cell.alignment = ReportStyles.ALIGN_LEFT
        ws.row_dimensions[row].height = 42
        row += 1
        if error:
            ws.merge_cells(f"A{row}:{get_column_letter(total_cols)}{row}")
            det = ws.cell(row=row, column=1, value=f"Details: {error}")
            det.font = ReportStyles.FONT_SMALL
            det.alignment = ReportStyles.ALIGN_LEFT
            row += 1
        row += 1
        return row

    def _embed_findings_screenshot(self, ws, row, findings, date_range, styles):
        """
        Render and embed a console-screenshot-style PNG of the Inspector
        one-week findings (High / Medium / Low severity).

        Any failure (missing Pillow, drawing error, embed error) is logged
        and swallowed so it never aborts the sheet.
        """
        try:
            total_cols = len(self.DETAIL_HEADERS)
            styles.apply_section_header(
                ws, row, "Inspector Console - Findings (Screenshot)", total_cols
            )
            anchor_row = row + 1

            headers = ["Severity", "Title", "Resource", "CVE", "First Observed"]
            img_rows = []
            for finding in findings:
                img_rows.append([
                    (finding.get("severity", "-") or "-").upper(),
                    finding.get("title", "-"),
                    finding.get("resource_id", "-") or "-",
                    finding.get("cve", "-"),
                    finding.get("first_observed", "-"),
                ])

            title = "Amazon Inspector - Findings (One Week)"
            if date_range and date_range != "N/A":
                title = f"{title}  |  {date_range}"

            png = screenshots.render_findings_table(
                title, headers, img_rows, severity_col=0
            )
            if png is None:
                return

            from openpyxl.drawing.image import Image as XLImage

            img = XLImage(png)
            ws.add_image(img, f"A{anchor_row}")
        except Exception as exc:
            logger.warning(
                "Could not embed Inspector findings screenshot: %s", exc
            )

    @staticmethod
    def _apply_severity_fill(severity, *cells):
        """Apply a RED/YELLOW/GREEN fill to *cells* based on severity."""
        severity = (severity or "").upper()
        if severity in ("CRITICAL", "HIGH"):
            fill = ReportStyles.FILL_RED
            fg = ReportStyles.RED_FG
        elif severity == "MEDIUM":
            fill = ReportStyles.FILL_YELLOW
            fg = ReportStyles.YELLOW_FG
        else:
            fill = ReportStyles.FILL_GREEN
            fg = ReportStyles.GREEN_FG

        for cell in cells:
            cell.fill = fill
            cell.font = Font(name="Calibri", size=10, bold=True, color=fg)
