"""
WAF sheet generator for the AWS Weekly BAU Report.

Displays Web Application Firewall metrics including total, blocked,
allowed, CAPTCHA, and challenged request counts with highlighting.
"""

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from report.styles import ReportStyles


class WAFSheet:
    """Generates the WAF worksheet."""

    SHEET_NAME = "WAF"

    HEADERS = [
        "WAF Name",
        "Total Request",
        "Blocked Request",
        "Allowed Request",
    ]

    def generate(self, wb, waf_data, styles=None):
        """
        Create the WAF sheet.

        Parameters
        ----------
        wb : openpyxl.Workbook
        waf_data : list[dict]
            WAF data from the WAF collector.
        styles : ReportStyles class (optional)

        Returns
        -------
        openpyxl.worksheet.worksheet.Worksheet
        """
        styles = styles or ReportStyles
        ws = wb.create_sheet(self.SHEET_NAME)
        ws.sheet_properties.tabColor = ReportStyles.DARK_BLUE

        waf_data = waf_data or []
        total_cols = len(self.HEADERS)

        row = 1

        # ── Title bar ────────────────────────────────────────────────
        styles.apply_title_bar(ws, row, "AWS WAF Summary", total_cols)
        row += 2

        # ── Section header ────────────────────────────────────────────
        styles.apply_section_header(ws, row, "Web ACL Request Metrics", total_cols)
        row += 1

        # ── Header row ────────────────────────────────────────────────
        header_row = row
        styles.apply_header_row(ws, row, self.HEADERS)
        row += 1

        # ── Data rows ────────────────────────────────────────────────
        if not waf_data:
            ws.merge_cells(f"A{row}:{get_column_letter(total_cols)}{row}")
            cell = ws.cell(row=row, column=1, value="No WAF data available.")
            cell.font = ReportStyles.FONT_DATA
            cell.alignment = ReportStyles.ALIGN_CENTER
            row += 1
        else:
            for i, waf in enumerate(waf_data):
                even = i % 2 == 0
                fill = ReportStyles.FILL_LIGHT_GRAY if even else ReportStyles.FILL_WHITE

                row_data = [
                    waf.get("name", "-"),
                    styles.format_count(waf.get("total_requests")),
                    styles.format_count(waf.get("blocked_requests")),
                    styles.format_count(waf.get("allowed_requests")),
                ]
                styles.apply_data_row(ws, row, row_data, even=even)

                # Highlight blocked-requests cell if non-zero
                blocked = waf.get("blocked_requests", 0)
                if blocked and float(blocked) > 0:
                    blocked_cell = ws.cell(row=row, column=3)
                    blocked_cell.fill = PatternFill(
                        start_color=ReportStyles.YELLOW_BG,
                        end_color=ReportStyles.YELLOW_BG,
                        fill_type="solid",
                    )
                    blocked_cell.font = Font(
                        name="Calibri", size=10, bold=True,
                        color=ReportStyles.YELLOW_FG,
                    )

                row += 1

            # ── Totals row ────────────────────────────────────────────
            row += 1
            styles.apply_data_cell(
                ws, row, 1, "TOTAL", even=False,
                align=ReportStyles.ALIGN_RIGHT,
            )
            ws.cell(row=row, column=1).font = ReportStyles.FONT_DATA_BOLD

            sum_fields = [
                "total_requests", "blocked_requests", "allowed_requests",
            ]
            for col_idx, field in enumerate(sum_fields, start=2):
                total_val = sum(w.get(field, 0) or 0 for w in waf_data)
                cell = styles.apply_data_cell(
                    ws, row, col_idx, styles.format_count(total_val), even=False,
                )
                cell.font = ReportStyles.FONT_DATA_BOLD

        # ── Auto-fit, freeze, filter ──────────────────────────────────
        styles.auto_fit_columns(ws, min_width=14)
        styles.freeze_and_filter(ws, header_row)

        return ws
