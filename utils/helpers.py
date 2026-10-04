"""
Helper utilities for the AWS BAU Report application.

Provides formatting, conversion, and data manipulation functions.
"""

import math
from datetime import datetime, date, timedelta
from typing import Any, Optional, Union


# ---------------------------------------------------------------------------
# Byte / size formatting
# ---------------------------------------------------------------------------

def format_bytes(size_bytes: Optional[float], precision: int = 1) -> str:
    """
    Convert bytes to a human-readable string with appropriate unit.

    Args:
        size_bytes: Size in bytes (can be None)
        precision: Number of decimal places

    Returns:
        Formatted string like '1.5 GB', '256.3 MB', etc.
    """
    if size_bytes is None or size_bytes == 0:
        return "-"

    units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB']
    size_bytes = float(size_bytes)

    if size_bytes < 0:
        return "-"

    unit_index = 0
    while abs(size_bytes) >= 1024.0 and unit_index < len(units) - 1:
        size_bytes /= 1024.0
        unit_index += 1

    return f"{size_bytes:.{precision}f} {units[unit_index]}"


def format_bytes_to_unit(size_bytes: Optional[float], target_unit: str = 'GB',
                         precision: int = 1) -> str:
    """
    Convert bytes to a specific unit.

    Args:
        size_bytes: Size in bytes
        target_unit: Target unit (B, KB, MB, GB, TB)
        precision: Decimal places

    Returns:
        Formatted string with unit
    """
    if size_bytes is None:
        return "-"

    divisors = {
        'B': 1,
        'KB': 1024,
        'MB': 1024 ** 2,
        'GB': 1024 ** 3,
        'TB': 1024 ** 4,
        'PB': 1024 ** 5,
    }

    divisor = divisors.get(target_unit.upper(), 1)
    value = float(size_bytes) / divisor

    return f"{value:.{precision}f} {target_unit}"


# ---------------------------------------------------------------------------
# Number formatting
# ---------------------------------------------------------------------------

def format_count(value: Optional[float], precision: int = 1) -> str:
    """
    Format a count with K, M, B suffixes for readability.

    Args:
        value: Numeric value
        precision: Decimal places

    Returns:
        Formatted string like '1.5K', '2.3M', etc.
    """
    if value is None:
        return "-"

    value = float(value)

    if abs(value) >= 1_000_000_000:
        return f"{value / 1_000_000_000:.{precision}f}B"
    elif abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.{precision}f}M"
    elif abs(value) >= 1_000:
        return f"{value / 1_000:.{precision}f}k"
    else:
        return f"{value:.0f}"


def format_currency(value: Optional[float], precision: int = 2) -> str:
    """
    Format a value as USD currency.

    Args:
        value: Dollar amount
        precision: Decimal places

    Returns:
        Formatted string like '$1,234.56'
    """
    if value is None:
        return "-"
    return f"${value:,.{precision}f}"


def format_percentage(value: Optional[float], precision: int = 2) -> str:
    """
    Format a value as a percentage string.

    Args:
        value: Percentage value (e.g., 45.67)
        precision: Decimal places

    Returns:
        Formatted string like '45.67%'
    """
    if value is None:
        return "N/A"
    return f"{value:.{precision}f}%"


def format_duration(seconds: Optional[float], precision: int = 2) -> str:
    """
    Format seconds into a human-readable duration.

    Args:
        seconds: Duration in seconds
        precision: Decimal places

    Returns:
        Formatted string like '1.74s', '250ms'
    """
    if seconds is None:
        return "-"

    if seconds < 0.001:
        return f"{seconds * 1_000_000:.0f}µs"
    elif seconds < 1:
        return f"{seconds * 1000:.0f}ms"
    elif seconds < 60:
        return f"{seconds:.{precision}f}s"
    elif seconds < 3600:
        return f"{seconds / 60:.{precision}f}m"
    else:
        return f"{seconds / 3600:.{precision}f}h"


# ---------------------------------------------------------------------------
# Date utilities
# ---------------------------------------------------------------------------

def parse_date(date_str: str) -> date:
    """
    Parse a date string in YYYY-MM-DD format.

    Args:
        date_str: Date string

    Returns:
        datetime.date object

    Raises:
        ValueError: If the date string is invalid
    """
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        raise ValueError(
            f"Invalid date format '{date_str}'. Expected YYYY-MM-DD."
        )


def validate_date_range(start_date: date, end_date: date) -> None:
    """
    Validate that the date range is sensible.

    Args:
        start_date: Start of range
        end_date: End of range

    Raises:
        ValueError: If the range is invalid
    """
    if start_date > end_date:
        raise ValueError(
            f"Start date ({start_date}) cannot be after end date ({end_date})."
        )

    days_diff = (end_date - start_date).days
    if days_diff > 31:
        raise ValueError(
            f"Date range ({days_diff} days) exceeds maximum of 31 days."
        )

    if end_date > date.today():
        raise ValueError(
            f"End date ({end_date}) cannot be in the future."
        )


def format_date_range(start_date: date, end_date: date) -> str:
    """
    Format a date range for display (e.g., '16th June – 22nd June 2026').

    Args:
        start_date: Start of range
        end_date: End of range

    Returns:
        Formatted date range string
    """
    def ordinal(n: int) -> str:
        if 11 <= (n % 100) <= 13:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
        return f"{n}{suffix}"

    start_str = f"{ordinal(start_date.day)} {start_date.strftime('%B')}"
    end_str = f"{ordinal(end_date.day)} {end_date.strftime('%B %Y')}"

    if start_date.year != end_date.year:
        start_str += f" {start_date.year}"

    return f"{start_str} – {end_str}"


def get_previous_period(start_date: date, end_date: date):
    """
    Calculate the previous period of the same length.

    Args:
        start_date: Current period start
        end_date: Current period end

    Returns:
        Tuple of (prev_start, prev_end) dates
    """
    period_length = (end_date - start_date).days + 1
    prev_end = start_date - timedelta(days=1)
    prev_start = prev_end - timedelta(days=period_length - 1)
    return prev_start, prev_end


# ---------------------------------------------------------------------------
# Data access utilities
# ---------------------------------------------------------------------------

def safe_get(data: dict, *keys, default: Any = None) -> Any:
    """
    Safely navigate nested dictionaries.

    Args:
        data: Source dictionary
        *keys: Sequence of keys to traverse
        default: Value to return if path doesn't exist

    Returns:
        Value at the specified path, or default
    """
    current = data
    for key in keys:
        if isinstance(current, dict):
            current = current.get(key, default)
        elif isinstance(current, (list, tuple)) and isinstance(key, int):
            try:
                current = current[key]
            except (IndexError, TypeError):
                return default
        else:
            return default

        if current is default:
            return default

    return current


def calculate_percentage_change(old_value: float, new_value: float) -> Optional[float]:
    """
    Calculate the percentage change between two values.

    Args:
        old_value: Previous value
        new_value: Current value

    Returns:
        Percentage change (e.g., 15.5 for a 15.5% increase), or None if old_value is 0
    """
    if old_value == 0:
        return None if new_value == 0 else 100.0

    return ((new_value - old_value) / abs(old_value)) * 100.0


# ---------------------------------------------------------------------------
# EC2 tag extraction
# ---------------------------------------------------------------------------

def get_instance_name(instance: dict) -> str:
    """
    Extract the Name tag from an EC2 instance.

    Args:
        instance: EC2 instance dictionary from describe_instances()

    Returns:
        Instance name or 'N/A'
    """
    for tag in instance.get('Tags', []):
        if tag.get('Key') == 'Name':
            return tag.get('Value', 'N/A')
    return 'N/A'


# ---------------------------------------------------------------------------
# CloudWatch helpers
# ---------------------------------------------------------------------------

def aggregate_datapoints(datapoints: list, statistic: str = 'Average') -> Optional[float]:
    """
    Aggregate CloudWatch datapoints into a single value.

    Args:
        datapoints: List of CloudWatch datapoint dicts
        statistic: Which statistic to extract ('Average', 'Sum', 'Minimum', 'Maximum')

    Returns:
        Aggregated value or None if no datapoints
    """
    if not datapoints:
        return None

    values = [dp.get(statistic) for dp in datapoints if dp.get(statistic) is not None]

    if not values:
        return None

    if statistic == 'Sum':
        return sum(values)
    elif statistic == 'Minimum':
        return min(values)
    elif statistic == 'Maximum':
        return max(values)
    elif statistic == 'Average':
        return sum(values) / len(values)
    else:
        return sum(values) / len(values)


def get_report_period_seconds(start_date: date, end_date: date) -> int:
    """
    Calculate the total seconds in the reporting period.

    Args:
        start_date: Period start
        end_date: Period end

    Returns:
        Total seconds
    """
    days = (end_date - start_date).days + 1
    return days * 86400
