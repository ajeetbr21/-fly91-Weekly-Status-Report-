"""
Chart helper module for AWS Weekly BAU Report.

Provides factory functions for creating professional bar, pie, and line
charts with a consistent colour palette using openpyxl's charting API.
"""

from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.series import DataPoint, SeriesLabel
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DrawingFont

from report.styles import ReportStyles


def _chart_font(size=10, bold=False, color="333333"):
    """Return an openpyxl drawing CharacterProperties for chart text."""
    return CharacterProperties(
        sz=size * 100,
        b=bold,
        solidFill=color,
        latin=DrawingFont(typeface="Calibri"),
    )


def _apply_common_style(chart, title, width=15, height=10):
    """Apply shared visual defaults to any chart object."""
    chart.title = title
    chart.width = width
    chart.height = height
    chart.style = 10
    chart.legend.position = "b"

    # Title font
    if chart.title:
        chart.title.txPr = RichText(
            p=[Paragraph(
                pPr=ParagraphProperties(defRPr=_chart_font(size=12, bold=True, color="002060")),
                endParaRPr=_chart_font(size=12, bold=True, color="002060"),
            )]
        )


def _color_series(chart, color_list=None):
    """Apply the professional colour palette to each series."""
    colors = color_list or ReportStyles.CHART_COLORS
    for idx, series in enumerate(chart.series):
        color = colors[idx % len(colors)]
        series.graphicalProperties.solidFill = color


# ------------------------------------------------------------------
# Public factory functions
# ------------------------------------------------------------------

def create_bar_chart(title, x_data_ref, y_data_refs, labels,
                     width=15, height=10, grouped=False):
    """
    Create a professional bar chart.

    Parameters
    ----------
    title : str
        Chart title.
    x_data_ref : openpyxl.chart.Reference
        Categories (x-axis) reference.
    y_data_refs : list[openpyxl.chart.Reference]
        One or more data series references.
    labels : list[str]
        Series names corresponding to *y_data_refs*.
    width, height : int
        Chart dimensions in cm.
    grouped : bool
        If True, use grouped (side-by-side) bars; else stacked is *not*
        used — standard clustered bars.

    Returns
    -------
    openpyxl.chart.BarChart
    """
    chart = BarChart()
    _apply_common_style(chart, title, width, height)
    chart.type = "col"
    chart.grouping = "clustered"

    for ref, label in zip(y_data_refs, labels):
        chart.add_data(ref, titles_from_data=True)

    chart.set_categories(x_data_ref)

    # Rename series after adding data
    for idx, label in enumerate(labels):
        if idx < len(chart.series):
            chart.series[idx].title = SeriesLabel(v=label)

    _color_series(chart)

    # Axis fonts
    if chart.x_axis:
        chart.x_axis.txPr = RichText(
            p=[Paragraph(pPr=ParagraphProperties(defRPr=_chart_font(size=8)))]
        )
        chart.x_axis.delete = False
    if chart.y_axis:
        chart.y_axis.txPr = RichText(
            p=[Paragraph(pPr=ParagraphProperties(defRPr=_chart_font(size=8)))]
        )
        chart.y_axis.delete = False

    return chart


def create_pie_chart(title, data_ref, cat_ref, width=15, height=10):
    """
    Create a professional pie chart with data labels.

    Parameters
    ----------
    title : str
        Chart title.
    data_ref : openpyxl.chart.Reference
        Data values reference.
    cat_ref : openpyxl.chart.Reference
        Category labels reference.
    width, height : int
        Chart dimensions in cm.

    Returns
    -------
    openpyxl.chart.PieChart
    """
    chart = PieChart()
    _apply_common_style(chart, title, width, height)

    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cat_ref)

    # Apply colours to individual data points
    colors = ReportStyles.CHART_COLORS
    if chart.series:
        series = chart.series[0]
        for idx in range(data_ref.max_row - data_ref.min_row):
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = colors[idx % len(colors)]
            series.data_points.append(pt)

        # Data labels
        series.dLbls = DataLabelList()
        series.dLbls.showPercent = True
        series.dLbls.showCatName = True
        series.dLbls.showVal = False
        series.dLbls.showSerName = False

    return chart


def create_line_chart(title, x_data_ref, y_data_refs, labels,
                      width=15, height=10):
    """
    Create a professional line chart.

    Parameters
    ----------
    title : str
        Chart title.
    x_data_ref : openpyxl.chart.Reference
        Categories (x-axis) reference.
    y_data_refs : list[openpyxl.chart.Reference]
        One or more data series references.
    labels : list[str]
        Series names corresponding to *y_data_refs*.
    width, height : int
        Chart dimensions in cm.

    Returns
    -------
    openpyxl.chart.LineChart
    """
    chart = LineChart()
    _apply_common_style(chart, title, width, height)
    chart.grouping = "standard"

    for ref, label in zip(y_data_refs, labels):
        chart.add_data(ref, titles_from_data=True)

    chart.set_categories(x_data_ref)

    for idx, label in enumerate(labels):
        if idx < len(chart.series):
            chart.series[idx].title = SeriesLabel(v=label)

    _color_series(chart)

    # Smooth lines
    for s in chart.series:
        s.smooth = True

    # Axis fonts
    if chart.x_axis:
        chart.x_axis.txPr = RichText(
            p=[Paragraph(pPr=ParagraphProperties(defRPr=_chart_font(size=8)))]
        )
        chart.x_axis.delete = False
    if chart.y_axis:
        chart.y_axis.txPr = RichText(
            p=[Paragraph(pPr=ParagraphProperties(defRPr=_chart_font(size=8)))]
        )
        chart.y_axis.delete = False

    return chart
