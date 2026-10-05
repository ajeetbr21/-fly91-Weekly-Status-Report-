"""
Report styles module for AWS Weekly BAU Report.

Provides the ReportStyles class with all named styles, color constants,
border definitions, font definitions, alignment definitions, number format
constants, and helper methods for professional Excel formatting.
"""

from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, NamedStyle, numbers
)
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter


class ReportStyles:
    """Centralised style definitions for the AWS Weekly BAU Report."""

    # ── Colour constants ──────────────────────────────────────────────
    DARK_BLUE = "002060"
    LIGHT_BLUE = "4472C4"
    WHITE = "FFFFFF"
    LIGHT_GRAY = "F2F2F2"
    BLACK = "000000"
    MEDIUM_GRAY = "D9D9D9"
    BORDER_GRAY = "BFBFBF"

    # Conditional-formatting palettes
    GREEN_BG = "C6EFCE"
    GREEN_FG = "006100"
    YELLOW_BG = "FFEB9C"
    YELLOW_FG = "9C6500"
    RED_BG = "FFC7CE"
    RED_FG = "9C0006"

    # Chart colour palette (professional, consistent)
    CHART_COLORS = [
        "4472C4", "ED7D31", "A5A5A5", "FFC000",
        "5B9BD5", "70AD47", "264478", "9B57A0",
        "636363", "255E91",
    ]

    # ── Threshold defaults ────────────────────────────────────────────
    CPU_WARNING = 70
    CPU_CRITICAL = 90
    MEMORY_WARNING = 75
    MEMORY_CRITICAL = 90
    DISK_WARNING = 75
    DISK_CRITICAL = 90
    COST_INCREASE_WARNING = 10
    COST_INCREASE_CRITICAL = 25

    # ── Number-format constants ───────────────────────────────────────
    FMT_CURRENCY = '$#,##0.00'
    FMT_PERCENT = '0.00"%"'
    FMT_PERCENT_NATIVE = '0.00%'
    FMT_NUMBER = '#,##0'
    FMT_NUMBER_2DP = '#,##0.00'
    FMT_DATE = 'YYYY-MM-DD'

    # ── Fonts ─────────────────────────────────────────────────────────
    FONT_TITLE = Font(name="Calibri", size=18, bold=True, color=WHITE)
    FONT_HEADER = Font(name="Calibri", size=11, bold=True, color=WHITE)
    FONT_SECTION = Font(name="Calibri", size=12, bold=True, color=WHITE)
    FONT_DATA = Font(name="Calibri", size=10, color=BLACK)
    FONT_DATA_BOLD = Font(name="Calibri", size=10, bold=True, color=BLACK)
    FONT_LABEL = Font(name="Calibri", size=10, bold=True, color="333333")
    FONT_VALUE = Font(name="Calibri", size=10, color="333333")
    FONT_SMALL = Font(name="Calibri", size=9, color="666666")
    FONT_SUBTITLE = Font(name="Calibri", size=14, bold=True, color=DARK_BLUE)
    FONT_COVER_TITLE = Font(name="Calibri", size=28, bold=True, color=WHITE)
    FONT_COVER_SUBTITLE = Font(name="Calibri", size=16, color=WHITE)
    FONT_COVER_DETAIL = Font(name="Calibri", size=12, color=DARK_BLUE)
    FONT_COVER_DETAIL_BOLD = Font(name="Calibri", size=12, bold=True, color=DARK_BLUE)

    # ── Fills ─────────────────────────────────────────────────────────
    FILL_DARK_BLUE = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
    FILL_LIGHT_BLUE = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
    FILL_WHITE = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
    FILL_LIGHT_GRAY = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
    FILL_GREEN = PatternFill(start_color=GREEN_BG, end_color=GREEN_BG, fill_type="solid")
    FILL_YELLOW = PatternFill(start_color=YELLOW_BG, end_color=YELLOW_BG, fill_type="solid")
    FILL_RED = PatternFill(start_color=RED_BG, end_color=RED_BG, fill_type="solid")

    # ── Alignments ────────────────────────────────────────────────────
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
    ALIGN_TITLE = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ── Borders ───────────────────────────────────────────────────────
    THIN_BORDER = Border(
        left=Side(style="thin", color=BORDER_GRAY),
        right=Side(style="thin", color=BORDER_GRAY),
        top=Side(style="thin", color=BORDER_GRAY),
        bottom=Side(style="thin", color=BORDER_GRAY),
    )
    NO_BORDER = Border()

    # ------------------------------------------------------------------
    # Named-style registration
    # ------------------------------------------------------------------
    _registered = False

    @classmethod
    def register_named_styles(cls, wb):
        """Register named styles on the workbook (idempotent)."""
        if cls._registered:
            return

        styles_map = {
            "title_style": {
                "font": cls.FONT_TITLE,
                "fill": cls.FILL_DARK_BLUE,
                "alignment": cls.ALIGN_TITLE,
            },
            "header_style": {
                "font": cls.FONT_HEADER,
                "fill": cls.FILL_DARK_BLUE,
                "alignment": cls.ALIGN_CENTER,
                "border": cls.THIN_BORDER,
            },
            "section_style": {
                "font": cls.FONT_SECTION,
                "fill": cls.FILL_LIGHT_BLUE,
                "alignment": cls.ALIGN_LEFT,
                "border": cls.THIN_BORDER,
            },
            "data_style": {
                "font": cls.FONT_DATA,
                "alignment": cls.ALIGN_LEFT,
                "border": cls.THIN_BORDER,
            },
            "data_center_style": {
                "font": cls.FONT_DATA,
                "alignment": cls.ALIGN_CENTER,
                "border": cls.THIN_BORDER,
            },
            "currency_style": {
                "font": cls.FONT_DATA,
                "alignment": cls.ALIGN_RIGHT,
                "border": cls.THIN_BORDER,
                "number_format": cls.FMT_CURRENCY,
            },
            "percent_style": {
                "font": cls.FONT_DATA,
                "alignment": cls.ALIGN_RIGHT,
                "border": cls.THIN_BORDER,
                "number_format": cls.FMT_PERCENT,
            },
            "number_style": {
                "font": cls.FONT_DATA,
                "alignment": cls.ALIGN_RIGHT,
                "border": cls.THIN_BORDER,
                "number_format": cls.FMT_NUMBER,
            },
        }

        for name, props in styles_map.items():
            if name not in wb.named_styles:
                ns = NamedStyle(name=name)
                ns.font = props.get("font", cls.FONT_DATA)
                ns.fill = props.get("fill", cls.FILL_WHITE)
                ns.alignment = props.get("alignment", cls.ALIGN_LEFT)
                ns.border = props.get("border", cls.NO_BORDER)
                if "number_format" in props:
                    ns.number_format = props["number_format"]
                wb.add_named_style(ns)

        cls._registered = True

    # ------------------------------------------------------------------
    # Helper methods
    # ------------------------------------------------------------------
    @classmethod
    def apply_title_bar(cls, ws, row, text, cols):
        """Merge cells across *cols* columns and apply the title style."""
        start_col = get_column_letter(1)
        end_col = get_column_letter(cols)
        ws.merge_cells(f"{start_col}{row}:{end_col}{row}")
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = cls.FONT_TITLE
        cell.fill = cls.FILL_DARK_BLUE
        cell.alignment = cls.ALIGN_TITLE
        ws.row_dimensions[row].height = 40

    @classmethod
    def apply_section_header(cls, ws, row, text, cols):
        """Apply a light-blue section header spanning *cols* columns."""
        start_col = get_column_letter(1)
        end_col = get_column_letter(cols)
        ws.merge_cells(f"{start_col}{row}:{end_col}{row}")
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = cls.FONT_SECTION
        cell.fill = cls.FILL_LIGHT_BLUE
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 28

    @classmethod
    def apply_header_row(cls, ws, row, headers, col_start=1):
        """Write header labels with dark-blue background and white bold text."""
        for idx, header in enumerate(headers, start=col_start):
            cell = ws.cell(row=row, column=idx, value=header)
            cell.font = cls.FONT_HEADER
            cell.fill = cls.FILL_DARK_BLUE
            cell.alignment = cls.ALIGN_CENTER
            cell.border = cls.THIN_BORDER
        ws.row_dimensions[row].height = 24

    @classmethod
    def apply_data_row(cls, ws, row, data, col_start=1, even=False):
        """Write a row of data values with alternating row colours."""
        fill = cls.FILL_LIGHT_GRAY if even else cls.FILL_WHITE
        for idx, value in enumerate(data, start=col_start):
            cell = ws.cell(row=row, column=idx, value=value)
            cell.font = cls.FONT_DATA
            cell.fill = fill
            cell.alignment = cls.ALIGN_CENTER
            cell.border = cls.THIN_BORDER
        ws.row_dimensions[row].height = 20

    @classmethod
    def apply_data_cell(cls, ws, row, col, value, fmt=None, even=False, align=None):
        """Write a single data cell with optional number format."""
        fill = cls.FILL_LIGHT_GRAY if even else cls.FILL_WHITE
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = cls.FONT_DATA
        cell.fill = fill
        cell.alignment = align or cls.ALIGN_CENTER
        cell.border = cls.THIN_BORDER
        if fmt:
            cell.number_format = fmt
        return cell

    @classmethod
    def apply_conditional_format(cls, ws, col_letter, start_row, end_row,
                                 warning_threshold, critical_threshold):
        """
        Apply three-tier conditional formatting (green / yellow / red)
        to a column range.

        Values are expected as plain numbers (e.g. 70 for 70 %).
        """
        cell_range = f"{col_letter}{start_row}:{col_letter}{end_row}"

        # Red – critical
        ws.conditional_formatting.add(
            cell_range,
            CellIsRule(
                operator="greaterThanOrEqual",
                formula=[str(critical_threshold)],
                fill=PatternFill(start_color=cls.RED_BG, end_color=cls.RED_BG, fill_type="solid"),
                font=Font(color=cls.RED_FG),
            ),
        )
        # Yellow – warning
        ws.conditional_formatting.add(
            cell_range,
            CellIsRule(
                operator="between",
                formula=[str(warning_threshold), str(critical_threshold - 0.01)],
                fill=PatternFill(start_color=cls.YELLOW_BG, end_color=cls.YELLOW_BG, fill_type="solid"),
                font=Font(color=cls.YELLOW_FG),
            ),
        )
        # Green – healthy
        ws.conditional_formatting.add(
            cell_range,
            CellIsRule(
                operator="lessThan",
                formula=[str(warning_threshold)],
                fill=PatternFill(start_color=cls.GREEN_BG, end_color=cls.GREEN_BG, fill_type="solid"),
                font=Font(color=cls.GREEN_FG),
            ),
        )

    @classmethod
    def auto_fit_columns(cls, ws, min_width=10, max_width=45, padding=3):
        """
        Auto-fit column widths based on cell content.

        Scans every cell in the worksheet and sets the column width to
        accommodate the longest value, bounded by *min_width* and *max_width*.
        """
        column_widths = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    col_letter = get_column_letter(cell.column)
                    length = len(str(cell.value)) + padding
                    current = column_widths.get(col_letter, min_width)
                    column_widths[col_letter] = min(max(current, length), max_width)

        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width

    @classmethod
    def freeze_and_filter(cls, ws, row):
        """Freeze panes at the given row and enable auto-filters."""
        ws.freeze_panes = f"A{row + 1}"
        max_col = ws.max_column or 1
        end_col = get_column_letter(max_col)
        ws.auto_filter.ref = f"A{row}:{end_col}{ws.max_row}"

    # ------------------------------------------------------------------
    # Utility formatters
    # ------------------------------------------------------------------
    @staticmethod
    def format_bytes(value):
        """Convert bytes to a human-readable string (KB / MB / GB / TB) using Base-1000 to match AWS CloudWatch."""
        if value is None:
            return "-"
        try:
            value = float(value)
        except (ValueError, TypeError):
            return "-"
        if value < 0:
            return "-"
        for unit in ("B", "KB", "MB", "GB", "TB"):
            if abs(value) < 1000:
                return f"{value:,.2f} {unit}"
            value /= 1000
        return f"{value:,.2f} PB"

    @staticmethod
    def format_count(value):
        """Format large counts with K / M suffixes."""
        if value is None:
            return "-"
        try:
            value = float(value)
        except (ValueError, TypeError):
            return "-"
        if value >= 1_000_000:
            return f"{value / 1_000_000:,.2f}M"
        if value >= 1_000:
            return f"{value / 1_000:,.2f}K"
        return f"{int(value):,}"

    @staticmethod
    def safe_value(value, default="-"):
        """Return *value* if truthy (or zero), else *default*."""
        if value is None:
            return default
        return value

    @staticmethod
    def safe_float(value, decimals=2, default="-"):
        """Safely format a float to the given decimal places."""
        if value is None:
            return default
        try:
            return round(float(value), decimals)
        except (ValueError, TypeError):
            return default

    @classmethod
    def memory_cell(cls, inst, decimals=2, default="-"):
        """Resolve the Memory Utilisation cell value for an EC2 instance.

        Mirrors the Word report's memory helper so the Excel sheet and the
        Word report render the same figure. Prefers an explicit
        ``memory_avg_display`` string (reference/mock figures such as
        "50.62"), then falls back to numeric ``memory_avg`` and
        ``memory_max`` (populated by the AWS collector via CWAgent).

        A numeric value (or a numeric-looking display string) is returned as a
        rounded float so the existing conditional formatting still applies;
        non-numeric values (e.g. "-") are returned as-is.
        """
        display = inst.get("memory_avg_display")
        if display is not None and display != "":
            as_float = cls.safe_float(display, decimals=decimals, default=None)
            return as_float if as_float is not None else display
        for key in ("memory_avg", "memory_max"):
            value = inst.get(key)
            if value is not None:
                as_float = cls.safe_float(value, decimals=decimals, default=None)
                if as_float is not None:
                    return as_float
        return default
