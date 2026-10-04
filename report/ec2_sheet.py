"""
EC2 sheet generator for the AWS Weekly BAU Report.

Produces two tables:
  1. Server Utilisation – CPU, memory, disk metrics per instance.
  2. Bandwidth & Network – network in/out and packet counts per instance.

Conditional formatting is applied to CPU, memory, and disk columns.
"""

from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from report.styles import ReportStyles


class EC2Sheet:
    """Generates the EC2 Instances worksheet."""

    SHEET_NAME = "EC2 Instances"

    # Table 1 headers
    UTIL_HEADERS = [
        "Name", "Instance ID", "Instance Type",
        "Min CPU (%)", "Max CPU (%)", "Avg CPU (%)",
        "Memory Utilization (%)",
        "Current Disk",
    ]

    # Table 2 headers
    NET_HEADERS = [
        "Name", "Instance ID",
        "Network In", "Network Out",
        "Network Packets In", "Network Packets Out",
    ]

    def generate(self, wb, ec2_data, styles=None):
        """
        Create the EC2 Instances sheet.

        Parameters
        ----------
        wb : openpyxl.Workbook
        ec2_data : list[dict]
            EC2 instance data from the EC2 collector.
        styles : ReportStyles class (optional)

        Returns
        -------
        openpyxl.worksheet.worksheet.Worksheet
        """
        styles = styles or ReportStyles
        ws = wb.create_sheet(self.SHEET_NAME)
        ws.sheet_properties.tabColor = ReportStyles.DARK_BLUE

        ec2_data = ec2_data or []
        total_cols = max(len(self.UTIL_HEADERS), len(self.NET_HEADERS))

        row = 1

        # ── Title bar ────────────────────────────────────────────────
        styles.apply_title_bar(ws, row, "EC2 Instance Details", total_cols)
        row += 2

        # ══════════════════════════════════════════════════════════════
        # TABLE 1 – Server Utilisation
        # ══════════════════════════════════════════════════════════════
        styles.apply_section_header(ws, row, "Server Utilisation", total_cols)
        row += 1

        header_row_1 = row
        styles.apply_header_row(ws, row, self.UTIL_HEADERS)
        row += 1

        if not ec2_data:
            ws.merge_cells(f"A{row}:{get_column_letter(total_cols)}{row}")
            cell = ws.cell(row=row, column=1, value="No EC2 instance data available.")
            cell.font = ReportStyles.FONT_DATA
            cell.alignment = ReportStyles.ALIGN_CENTER
            row += 2
        else:
            data_start_1 = row
            for i, inst in enumerate(ec2_data):
                even = i % 2 == 0
                row_data = [
                    inst.get("name", "-"),
                    inst.get("instance_id", "-"),
                    inst.get("instance_type", "-"),
                    styles.safe_float(inst.get("cpu_min")),
                    styles.safe_float(inst.get("cpu_max")),
                    styles.safe_float(inst.get("cpu_avg")),
                    styles.safe_float(inst.get("memory_max")),
                    inst.get("disk_utilization") or "-",
                ]
                styles.apply_data_row(ws, row, row_data, even=even)
                row += 1
            data_end_1 = row - 1

            # ── Conditional formatting for Table 1 ────────────────────
            # Columns D=Min CPU, E=Max CPU, F=Avg CPU → CPU thresholds
            for col_letter in ("D", "E", "F"):
                styles.apply_conditional_format(
                    ws, col_letter, data_start_1, data_end_1,
                    ReportStyles.CPU_WARNING, ReportStyles.CPU_CRITICAL,
                )
            # Column G=Memory Utilization → Memory thresholds
            styles.apply_conditional_format(
                ws, "G", data_start_1, data_end_1,
                ReportStyles.MEMORY_WARNING, ReportStyles.MEMORY_CRITICAL,
            )
            # Note: Disk is a string ("C: 49.4% D: 13%"), so conditional
            # formatting on numeric thresholds is not directly applicable.
            # We leave it as-is for now.

            row += 1

        # ══════════════════════════════════════════════════════════════
        # TABLE 2 – Bandwidth & Network
        # ══════════════════════════════════════════════════════════════
        styles.apply_section_header(ws, row, "Bandwidth & Network", total_cols)
        row += 1

        styles.apply_header_row(ws, row, self.NET_HEADERS)
        row += 1

        if not ec2_data:
            ws.merge_cells(f"A{row}:{get_column_letter(total_cols)}{row}")
            cell = ws.cell(row=row, column=1, value="No EC2 instance data available.")
            cell.font = ReportStyles.FONT_DATA
            cell.alignment = ReportStyles.ALIGN_CENTER
            row += 1
        else:
            for i, inst in enumerate(ec2_data):
                even = i % 2 == 0
                row_data = [
                    inst.get("name", "-"),
                    inst.get("instance_id", "-"),
                    styles.format_bytes(inst.get("network_in")),
                    styles.format_bytes(inst.get("network_out")),
                    styles.format_count(inst.get("network_packets_in")),
                    styles.format_count(inst.get("network_packets_out")),
                ]
                styles.apply_data_row(ws, row, row_data, even=even)
                row += 1

        # ── Auto-fit, freeze, filter ──────────────────────────────────
        styles.auto_fit_columns(ws, min_width=14)
        styles.freeze_and_filter(ws, header_row_1)

        return ws
