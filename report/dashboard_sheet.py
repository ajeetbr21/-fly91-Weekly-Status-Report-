"""
Dashboard sheet generator for the AWS Weekly BAU Report.

Places 8 professional charts in a 2×4 grid layout:
  1. Weekly Cost Trend (Line)
  2. Service-wise Cost Distribution (Pie)
  3. EC2 CPU Usage (Bar)
  4. Memory Usage (Bar)
  5. Disk Usage (Bar)
  6. Network Usage (Grouped Bar)
  7. S3 Storage Distribution (Pie)
  8. ALB Request Count (Bar)

Source data is written to a hidden area below the chart grid, and chart
references point to that data.
"""

import logging

from openpyxl.chart import Reference
from openpyxl.utils import get_column_letter

from report.styles import ReportStyles
from report.charts import create_bar_chart, create_pie_chart, create_line_chart

logger = logging.getLogger(__name__)


class DashboardSheet:
    """Generates the Dashboard worksheet with charts."""

    SHEET_NAME = "Dashboard"

    # Chart grid layout (anchor cells) – 2 charts per row
    # Each chart ≈ 15 cols wide × 20 rows tall
    CHART_POSITIONS = [
        "A2",   # Row 1, Col 1  – Weekly Cost Trend
        "I2",   # Row 1, Col 2  – Service Cost Pie
        "A22",  # Row 2, Col 1  – CPU Usage
        "I22",  # Row 2, Col 2  – Memory Usage
        "A42",  # Row 3, Col 1  – Disk Usage
        "I42",  # Row 3, Col 2  – Network Usage
        "A62",  # Row 4, Col 1  – S3 Storage Pie
        "I62",  # Row 4, Col 2  – ALB Requests
    ]

    # Hidden data area starts well below the charts
    DATA_START_ROW = 85

    def generate(self, wb, all_data, styles=None):
        """
        Create the Dashboard sheet with 8 charts.

        Parameters
        ----------
        wb : openpyxl.Workbook
        all_data : dict
            Complete collected_data dictionary.
        styles : ReportStyles class (optional)

        Returns
        -------
        openpyxl.worksheet.worksheet.Worksheet
        """
        styles = styles or ReportStyles
        ws = wb.create_sheet(self.SHEET_NAME)
        ws.sheet_properties.tabColor = ReportStyles.DARK_BLUE

        # Title bar across the top
        styles.apply_title_bar(ws, 1, "Dashboard – Visual Analytics", 16)

        # Security-findings summary line (Inspector + GuardDuty) so the
        # at-a-glance Dashboard references the headline security posture.
        try:
            self._write_security_summary(ws, all_data, styles)
        except Exception as exc:
            logger.warning("Failed to write Dashboard security summary: %s", exc)

        # Track current data-write row
        self._data_row = self.DATA_START_ROW

        # ── Build each chart ──────────────────────────────────────────
        try:
            chart1 = self._chart_cost_trend(ws, all_data)
            if chart1:
                ws.add_chart(chart1, self.CHART_POSITIONS[0])
        except Exception as exc:
            logger.warning("Failed to create Cost Trend chart: %s", exc)

        try:
            chart2 = self._chart_service_cost_pie(ws, all_data)
            if chart2:
                ws.add_chart(chart2, self.CHART_POSITIONS[1])
        except Exception as exc:
            logger.warning("Failed to create Service Cost Pie chart: %s", exc)

        try:
            chart3 = self._chart_cpu_usage(ws, all_data)
            if chart3:
                ws.add_chart(chart3, self.CHART_POSITIONS[2])
        except Exception as exc:
            logger.warning("Failed to create CPU Usage chart: %s", exc)

        try:
            chart4 = self._chart_memory_usage(ws, all_data)
            if chart4:
                ws.add_chart(chart4, self.CHART_POSITIONS[3])
        except Exception as exc:
            logger.warning("Failed to create Memory Usage chart: %s", exc)

        try:
            chart5 = self._chart_disk_usage(ws, all_data)
            if chart5:
                ws.add_chart(chart5, self.CHART_POSITIONS[4])
        except Exception as exc:
            logger.warning("Failed to create Disk Usage chart: %s", exc)

        try:
            chart6 = self._chart_network_usage(ws, all_data)
            if chart6:
                ws.add_chart(chart6, self.CHART_POSITIONS[5])
        except Exception as exc:
            logger.warning("Failed to create Network Usage chart: %s", exc)

        try:
            chart7 = self._chart_s3_pie(ws, all_data)
            if chart7:
                ws.add_chart(chart7, self.CHART_POSITIONS[6])
        except Exception as exc:
            logger.warning("Failed to create S3 Pie chart: %s", exc)

        try:
            chart8 = self._chart_alb_requests(ws, all_data)
            if chart8:
                ws.add_chart(chart8, self.CHART_POSITIONS[7])
        except Exception as exc:
            logger.warning("Failed to create ALB Requests chart: %s", exc)

        ws.sheet_view.showGridLines = False
        return ws

    # ------------------------------------------------------------------
    # Security summary
    # ------------------------------------------------------------------

    # Row for the security-findings summary line — below the chart grid
    # (charts occupy rows 2..~82) and above the hidden data area (row 85).
    SECURITY_SUMMARY_ROW = 83

    def _write_security_summary(self, ws, all_data, styles):
        """Write a one-line Inspector + GuardDuty security summary."""
        inspector = all_data.get("inspector") or {}
        guardduty = all_data.get("guardduty") or {}

        insp_counts = inspector.get("severity_counts", {}) or {}
        gd_counts = guardduty.get("severity_counts", {}) or {}

        insp_total = int(inspector.get("total_findings", 0) or 0)
        gd_total = int(guardduty.get("total_findings", 0) or 0)
        insp_crit = int(insp_counts.get("CRITICAL", 0) or 0)
        insp_high = int(insp_counts.get("HIGH", 0) or 0)
        gd_high = int(gd_counts.get("HIGH", 0) or 0)

        insp_status = inspector.get("collection_status", "ok")
        gd_status = guardduty.get("collection_status", "ok")

        row = self.SECURITY_SUMMARY_ROW
        styles.apply_section_header(ws, row, "Security Findings (Inspector + GuardDuty)", 16)
        row += 1

        if "error" in (insp_status, gd_status):
            text = (
                "⚠ Security findings collection failed / access denied — "
                "counts below may be incomplete."
            )
        else:
            text = (
                f"Amazon Inspector: {insp_total} findings "
                f"(Critical {insp_crit}, High {insp_high})   |   "
                f"Amazon GuardDuty: {gd_total} findings (High {gd_high}). "
                "See the Inspector and GuardDuty sheets for detail."
            )

        from openpyxl.utils import get_column_letter as _gcl
        ws.merge_cells(f"A{row}:{_gcl(16)}{row}")
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = ReportStyles.FONT_DATA_BOLD
        cell.alignment = ReportStyles.ALIGN_LEFT
        ws.row_dimensions[row].height = 24

    # ------------------------------------------------------------------
    # Data-writing helpers
    # ------------------------------------------------------------------

    def _write_data_block(self, ws, headers, rows):
        """
        Write a block of data into the hidden area and return
        (start_row, end_row, start_col=1).

        Returns (header_row, data_end_row, col_start) where header_row
        is the row containing *headers*.
        """
        start = self._data_row
        # Headers
        for c, h in enumerate(headers, start=1):
            ws.cell(row=start, column=c, value=h)
        # Data
        for r_idx, row_data in enumerate(rows, start=1):
            for c_idx, val in enumerate(row_data, start=1):
                ws.cell(row=start + r_idx, column=c_idx, value=val)
        end = start + len(rows)
        self._data_row = end + 2  # leave a gap
        return start, end, 1

    # ------------------------------------------------------------------
    # Chart builders
    # ------------------------------------------------------------------

    def _chart_cost_trend(self, ws, all_data):
        """1. Weekly Cost Trend (Line)."""
        daily = all_data.get("cost", {}).get("daily_costs", [])
        if not daily:
            return None

        headers = ["Date", "Daily Cost ($)"]
        rows = [[d.get("date", ""), d.get("cost", 0)] for d in daily]
        h_row, d_end, col = self._write_data_block(ws, headers, rows)

        cats = Reference(ws, min_col=1, min_row=h_row + 1, max_row=d_end)
        vals = Reference(ws, min_col=2, min_row=h_row, max_row=d_end)

        return create_line_chart(
            "Weekly Cost Trend",
            cats, [vals], ["Daily Cost ($)"],
            width=16, height=12,
        )

    def _chart_service_cost_pie(self, ws, all_data):
        """2. Service-wise Cost Distribution (Pie)."""
        services = all_data.get("cost", {}).get("top_services", [])
        if not services:
            return None

        services = sorted(services, key=lambda s: s.get("cost", 0), reverse=True)[:8]
        headers = ["Service", "Cost ($)"]
        rows = [[s.get("service", ""), s.get("cost", 0)] for s in services]
        h_row, d_end, col = self._write_data_block(ws, headers, rows)

        cats = Reference(ws, min_col=1, min_row=h_row + 1, max_row=d_end)
        vals = Reference(ws, min_col=2, min_row=h_row, max_row=d_end)

        return create_pie_chart(
            "Service-wise Cost Distribution",
            vals, cats,
            width=16, height=12,
        )

    def _chart_cpu_usage(self, ws, all_data):
        """3. EC2 CPU Usage (Bar)."""
        ec2 = all_data.get("ec2", [])
        instances = [i for i in ec2 if i.get("cpu_avg") is not None]
        if not instances:
            return None

        headers = ["Instance", "Avg CPU (%)"]
        rows = [
            [i.get("name", i.get("instance_id", "?")), i.get("cpu_avg", 0)]
            for i in instances
        ]
        h_row, d_end, col = self._write_data_block(ws, headers, rows)

        cats = Reference(ws, min_col=1, min_row=h_row + 1, max_row=d_end)
        vals = Reference(ws, min_col=2, min_row=h_row, max_row=d_end)

        return create_bar_chart(
            "EC2 Average CPU Usage (%)",
            cats, [vals], ["Avg CPU (%)"],
            width=16, height=12,
        )

    def _chart_memory_usage(self, ws, all_data):
        """4. Memory Usage (Bar)."""
        ec2 = all_data.get("ec2", [])
        instances = [i for i in ec2 if i.get("memory_avg") is not None]
        if not instances:
            return None

        headers = ["Instance", "Avg Memory (%)"]
        rows = [
            [i.get("name", i.get("instance_id", "?")), i.get("memory_avg", 0)]
            for i in instances
        ]
        h_row, d_end, col = self._write_data_block(ws, headers, rows)

        cats = Reference(ws, min_col=1, min_row=h_row + 1, max_row=d_end)
        vals = Reference(ws, min_col=2, min_row=h_row, max_row=d_end)

        return create_bar_chart(
            "EC2 Average Memory Usage (%)",
            cats, [vals], ["Avg Memory (%)"],
            width=16, height=12,
        )

    def _chart_disk_usage(self, ws, all_data):
        """5. Disk Usage (Bar) – parses primary disk % from string."""
        ec2 = all_data.get("ec2", [])
        instances = [i for i in ec2 if i.get("disk_utilization")]
        if not instances:
            return None

        headers = ["Instance", "Disk Usage (%)"]
        rows = []
        for inst in instances:
            disk_str = inst.get("disk_utilization", "")
            pct = self._parse_primary_disk_pct(disk_str)
            if pct is not None:
                rows.append([
                    inst.get("name", inst.get("instance_id", "?")),
                    pct,
                ])
        if not rows:
            return None

        h_row, d_end, col = self._write_data_block(ws, headers, rows)

        cats = Reference(ws, min_col=1, min_row=h_row + 1, max_row=d_end)
        vals = Reference(ws, min_col=2, min_row=h_row, max_row=d_end)

        return create_bar_chart(
            "EC2 Disk Usage (%)",
            cats, [vals], ["Disk Usage (%)"],
            width=16, height=12,
        )

    def _chart_network_usage(self, ws, all_data):
        """6. Network Usage – Grouped Bar (In / Out per instance)."""
        ec2 = all_data.get("ec2", [])
        instances = [
            i for i in ec2
            if i.get("network_in") is not None or i.get("network_out") is not None
        ]
        if not instances:
            return None

        headers = ["Instance", "Network In (MB)", "Network Out (MB)"]
        rows = []
        for inst in instances:
            net_in = (inst.get("network_in") or 0) / (1024 * 1024)
            net_out = (inst.get("network_out") or 0) / (1024 * 1024)
            rows.append([
                inst.get("name", inst.get("instance_id", "?")),
                round(net_in, 2),
                round(net_out, 2),
            ])
        h_row, d_end, col = self._write_data_block(ws, headers, rows)

        cats = Reference(ws, min_col=1, min_row=h_row + 1, max_row=d_end)
        vals_in = Reference(ws, min_col=2, min_row=h_row, max_row=d_end)
        vals_out = Reference(ws, min_col=3, min_row=h_row, max_row=d_end)

        return create_bar_chart(
            "EC2 Network Usage (MB)",
            cats, [vals_in, vals_out],
            ["Network In (MB)", "Network Out (MB)"],
            width=16, height=12, grouped=True,
        )

    def _chart_s3_pie(self, ws, all_data):
        """7. S3 Storage Distribution (Pie)."""
        s3 = all_data.get("s3", [])
        buckets = [b for b in s3 if b.get("total_size_bytes")]
        if not buckets:
            return None

        headers = ["Bucket", "Size (GB)"]
        rows = [
            [
                b.get("bucket_name", "?"),
                round((b.get("total_size_bytes", 0) or 0) / (1024 ** 3), 2),
            ]
            for b in buckets
        ]
        h_row, d_end, col = self._write_data_block(ws, headers, rows)

        cats = Reference(ws, min_col=1, min_row=h_row + 1, max_row=d_end)
        vals = Reference(ws, min_col=2, min_row=h_row, max_row=d_end)

        return create_pie_chart(
            "S3 Storage Distribution",
            vals, cats,
            width=16, height=12,
        )

    def _chart_alb_requests(self, ws, all_data):
        """8. ALB Request Count (Bar)."""
        elb = all_data.get("elb", [])
        albs = [a for a in elb if a.get("requests") is not None]
        if not albs:
            return None

        headers = ["ALB Name", "Requests"]
        rows = [[a.get("name", "?"), a.get("requests", 0)] for a in albs]
        h_row, d_end, col = self._write_data_block(ws, headers, rows)

        cats = Reference(ws, min_col=1, min_row=h_row + 1, max_row=d_end)
        vals = Reference(ws, min_col=2, min_row=h_row, max_row=d_end)

        return create_bar_chart(
            "ALB Request Count",
            cats, [vals], ["Requests"],
            width=16, height=12,
        )

    # ------------------------------------------------------------------
    # Utilities
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_primary_disk_pct(disk_str):
        """
        Extract the first numeric percentage from a disk string.

        Examples:
            "C: 49.4% D: 13%" → 49.4
            "40.3%"           → 40.3
            None / ""         → None
        """
        if not disk_str:
            return None
        import re
        match = re.search(r"([\d.]+)\s*%", str(disk_str))
        if match:
            try:
                return float(match.group(1))
            except ValueError:
                return None
        return None
