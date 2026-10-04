"""
RDS sheet generator for the FLY-91 Weekly Report.
"""

from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from report.styles import ReportStyles


class RDSSheet:
    """Generates the RDS Instances worksheet."""

    SHEET_NAME = "RDS Instances"

    HEADERS_1 = [
        "RDS Name",
        "Down Time",
        "Instance type",
        "Minimum Utilization",
        "Maximum Utilization",
        "Average Utilization",
        "Free Memory",
        "Free Storage",
    ]

    HEADERS_2 = [
        "RDS Name",
        "Network Transmit Throughput (Bytes per second)",
        "Network Receive Throughput (Bytes per second)",
        "Max Database Connection (Count)",
    ]

    def generate(self, wb, rds_data, styles=None):
        styles = styles or ReportStyles
        ws = wb.create_sheet(self.SHEET_NAME)
        ws.sheet_properties.tabColor = ReportStyles.DARK_BLUE

        rds_data = rds_data or []

        # Table 1: Instance & CPU & Storage Info
        total_cols_1 = len(self.HEADERS_1)
        widths_1 = [30, 15, 20, 20, 20, 20, 25, 25]
        for idx, w in enumerate(widths_1, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = w

        row = 1
        styles.apply_title_bar(ws, row, "RDS Instance Details", total_cols_1)
        row += 2

        if not rds_data:
            ws.merge_cells(f"A{row}:{get_column_letter(total_cols_1)}{row}")
            cell = ws.cell(row=row, column=1, value="No RDS instances available.")
            cell.font = Font(name="Calibri", size=12, bold=True, color=ReportStyles.LIGHT_BLUE)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[row].height = 30
            return ws

        styles.apply_header_row(ws, row, self.HEADERS_1)
        row += 1

        for i, db in enumerate(rds_data):
            even = i % 2 == 0
            
            # Format Free Memory
            free_mem = db.get("free_memory_bytes")
            tot_mem = db.get("total_memory_gb")
            if free_mem is not None and tot_mem is not None:
                free_mem_gb = free_mem / (1024 ** 3)
                mem_str = f"{free_mem_gb:.2f}GB out of {tot_mem} GB"
            else:
                mem_str = "-"
                
            # Format Free Storage
            free_stg = db.get("free_storage_bytes")
            tot_stg = db.get("allocated_storage_gb")
            if free_stg is not None and tot_stg is not None:
                free_stg_gb = free_stg / (1024 ** 3)
                stg_str = f"{free_stg_gb:.1f}GB out of {tot_stg} GB"
            else:
                stg_str = "-"

            row_data = [
                db.get("db_identifier", "-"),
                "No",  # Down Time
                db.get("instance_class", "-"),
                f"{db.get('cpu_min', '-')}%" if db.get('cpu_min') is not None else "-",
                f"{db.get('cpu_max', '-')}%" if db.get('cpu_max') is not None else "-",
                f"{db.get('cpu_avg', '-')}%" if db.get('cpu_avg') is not None else "-",
                mem_str,
                stg_str,
            ]
            styles.apply_data_row(ws, row, row_data, even=even)
            row += 1

        row += 3

        # Table 2: Bandwidth and Network status
        total_cols_2 = len(self.HEADERS_2)
        widths_2 = [30, 45, 45, 35]
        # We only override widths if they are wider than Table 1, or we just rely on Table 1's widths for cols A-D
        # For columns B, C, D which are now wider, we will adjust them
        ws.column_dimensions['B'].width = max(ws.column_dimensions['B'].width, 45)
        ws.column_dimensions['C'].width = max(ws.column_dimensions['C'].width, 45)
        ws.column_dimensions['D'].width = max(ws.column_dimensions['D'].width, 35)

        # Title for Table 2
        cell = ws.cell(row=row, column=1, value="Bandwidth and Network status")
        cell.font = Font(name="Calibri", size=14, bold=True)
        row += 2

        styles.apply_header_row(ws, row, self.HEADERS_2)
        row += 1

        for i, db in enumerate(rds_data):
            even = i % 2 == 0
            
            tx = db.get("network_tx_bytes_sec")
            rx = db.get("network_rx_bytes_sec")
            conn = db.get("db_connections")

            row_data = [
                db.get("db_identifier", "-"),
                styles.format_bytes(tx) if tx is not None else "-",
                styles.format_bytes(rx) if rx is not None else "-",
                f"{conn:,.0f}" if conn is not None else "-",
            ]
            styles.apply_data_row(ws, row, row_data, even=even)
            row += 1

        # Adjust auto-filter to only apply to the first table (or disable it, let's just let format_table do its thing)
        # We don't apply an auto_filter across both tables easily, so we skip it.
        return ws
