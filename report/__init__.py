"""
Report package for the AWS Weekly BAU Report.

Exports the main :class:`ExcelGenerator` class for convenient access::

    from report import ExcelGenerator
"""

from report.excel_generator import ExcelGenerator

__all__ = ["ExcelGenerator"]
