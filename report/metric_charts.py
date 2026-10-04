"""
CloudWatch-style metric chart helper for the AWS Weekly Status Report.

Renders AWS-CloudWatch-style line charts of per-EC2 metric time-series (CPU,
memory, disk, network) over the reporting week entirely in memory using
matplotlib (headless Agg backend). The resulting PNG can be embedded into the
Word (.docx) report via ``python-docx`` ``doc.add_picture(buf, width=...)``.

The module is intentionally dependency-light (only matplotlib) and degrades
gracefully, mirroring report/screenshots.py: if matplotlib is unavailable or
any drawing step fails, the public helper logs a warning and returns ``None``
so the caller can simply skip the embed instead of aborting the report.

All type hints are Python 3.9 compatible (typing.Optional / typing.List,
no PEP 604 "X | None" unions).
"""

from io import BytesIO
from typing import Any, Dict, List, Optional

from utils.logger import get_logger

logger = get_logger("bau_report")

# ── Optional matplotlib import (graceful fallback, headless Agg backend) ────
try:
    import matplotlib
    matplotlib.use("Agg")  # headless backend; no display required
    import matplotlib.pyplot as plt  # noqa: E402
    import matplotlib.dates as mdates  # noqa: E402
    _MPL_AVAILABLE = True
except Exception as exc:  # pragma: no cover - exercised only without matplotlib
    matplotlib = None  # type: ignore
    plt = None  # type: ignore
    mdates = None  # type: ignore
    _MPL_AVAILABLE = False
    logger.warning(
        "matplotlib is not available - CloudWatch metric charts will be "
        "skipped: %s", exc
    )


# ── AWS-CloudWatch-ish colour palette ──────────────────────────────────────
_LINE_COLORS = [
    "#1f77b4",  # blue
    "#ff7f0e",  # orange
    "#2ca02c",  # green
    "#9467bd",  # purple
]
_WARNING_COLOR = "#f0ad4e"   # amber
_CRITICAL_COLOR = "#d62728"  # red
_GRID_COLOR = "#d9d9d9"
_TITLE_COLOR = "#232f3e"     # AWS dark navy


def render_metric_chart(
    server_name: str,
    series_list: List[Dict[str, Any]],
    thresholds: Optional[Dict[str, Any]] = None,
) -> Optional[BytesIO]:
    """
    Render one or more metric time-series as a CloudWatch-style line chart.

    Parameters
    ----------
    server_name : str
        Server/instance label used in the chart title (e.g.
        ``"fly91-app-prod(i-0a1b2c3d4e5f60718)"``).
    series_list : list of dict
        Each series dict should contain:
            ``label``       - metric label shown in the legend (e.g.
                              ``"CPUUtilization"``).
            ``unit``        - y-axis unit string (e.g. ``"%"``).
            ``timestamps``  - list of ``datetime`` objects (x values).
            ``values``      - list of floats (y values), same length as
                              ``timestamps``.
            ``warning``     - optional numeric warning threshold for a dashed
                              amber line.
            ``critical``    - optional numeric critical threshold for a dashed
                              red line.
    thresholds : dict, optional
        Fallback config thresholds (``config["thresholds"]``) used to infer
        warning/critical lines by metric type when a series does not carry its
        own ``warning``/``critical`` values.

    Returns
    -------
    io.BytesIO | None
        A ``BytesIO`` positioned at 0 containing PNG bytes, or ``None`` if
        matplotlib is unavailable, there is no plottable data, or rendering
        failed (a warning is logged).
    """
    if not _MPL_AVAILABLE:
        logger.warning(
            "Skipping metric chart for '%s' - matplotlib is not installed.",
            server_name,
        )
        return None

    try:
        # Keep only series that actually have aligned, non-empty data.
        plottable: List[Dict[str, Any]] = []
        for series in series_list or []:
            timestamps = series.get("timestamps") or []
            values = series.get("values") or []
            if timestamps and values and len(timestamps) == len(values):
                plottable.append(series)

        if not plottable:
            logger.warning(
                "Skipping metric chart for '%s' - no plottable series.",
                server_name,
            )
            return None

        thresholds = thresholds or {}

        fig, ax = plt.subplots(figsize=(9.0, 3.6))
        fig.patch.set_facecolor("white")
        ax.set_facecolor("white")

        # Determine whether this chart mixes percent and non-percent (byte)
        # series. When it does, byte series go on a secondary y-axis so their
        # large magnitudes do not flatten the 0-100 percent lines.
        units: List[str] = []
        for series in plottable:
            unit = str(series.get("unit", "") or "")
            if unit and unit not in units:
                units.append(unit)

        has_percent = any(u == "%" for u in units)
        has_non_percent = any(u != "%" for u in units)
        mixed = has_percent and has_non_percent

        # Secondary axis only created when we actually mix percent + byte.
        ax2 = ax.twinx() if mixed else None
        if ax2 is not None:
            ax2.set_facecolor("none")

        warn_lines: Dict[float, str] = {}
        crit_lines: Dict[float, str] = {}

        for idx, series in enumerate(plottable):
            color = _LINE_COLORS[idx % len(_LINE_COLORS)]
            label = str(series.get("label", "metric"))
            unit = str(series.get("unit", "") or "")
            # In mixed mode, non-percent (byte) series plot on the secondary
            # axis; everything else stays on the primary axis.
            target = ax
            if mixed and unit != "%":
                target = ax2
            target.plot(
                series["timestamps"],
                series["values"],
                color=color,
                linewidth=1.6,
                label=label,
            )

            warning, critical = _resolve_thresholds(series, thresholds)
            if isinstance(warning, (int, float)):
                warn_lines[float(warning)] = color
            if isinstance(critical, (int, float)):
                crit_lines[float(critical)] = color

        # Dashed horizontal warning/critical threshold lines (percent metrics
        # only, so always on the primary axis).
        for value in sorted(warn_lines):
            ax.axhline(
                value, color=_WARNING_COLOR, linestyle="--", linewidth=1.1,
                alpha=0.9, label="Warning %g" % value,
            )
        for value in sorted(crit_lines):
            ax.axhline(
                value, color=_CRITICAL_COLOR, linestyle="--", linewidth=1.1,
                alpha=0.9, label="Critical %g" % value,
            )

        # Title styled like the CloudWatch metrics panel header.
        ax.set_title(
            "CloudWatch Metrics - %s" % server_name,
            fontsize=11, fontweight="bold", color=_TITLE_COLOR, loc="left",
        )

        if mixed:
            # Primary axis holds the percent series; secondary holds bytes.
            ax.set_ylabel("%", fontsize=9)
            ax.set_ylim(0, 100)
            non_percent_units = [u for u in units if u != "%"]
            ax2.set_ylabel(
                non_percent_units[0] if non_percent_units else "Value",
                fontsize=9,
            )
        else:
            y_label = units[0] if len(units) == 1 else "Value"
            ax.set_ylabel(y_label, fontsize=9)
            # Percent metrics read best pinned to a 0-100 range.
            if units and all(u == "%" for u in units):
                ax.set_ylim(0, 100)
        ax.set_xlabel("Time (UTC)", fontsize=9)

        # Date/time formatting across the reporting week.
        try:
            ax.xaxis.set_major_locator(mdates.AutoDateLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%H:%M"))
        except Exception:
            pass
        for tick in ax.get_xticklabels():
            tick.set_fontsize(7)

        ax.grid(True, color=_GRID_COLOR, linewidth=0.7, alpha=0.9)

        # Merge both axes' handles/labels into a single legend.
        handles, labels = ax.get_legend_handles_labels()
        if ax2 is not None:
            h2, l2 = ax2.get_legend_handles_labels()
            handles += h2
            labels += l2
        ax.legend(
            handles, labels,
            loc="upper left", fontsize=7, ncol=2, framealpha=0.9,
        )

        fig.tight_layout()

        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=110)
        plt.close(fig)
        buf.seek(0)
        return buf
    except Exception as exc:
        logger.warning(
            "Failed to render metric chart for '%s' - skipping embed: %s",
            server_name, exc,
        )
        # Best-effort cleanup so a half-built figure never leaks.
        try:
            plt.close("all")
        except Exception:
            pass
        return None


# ── Private helpers ─────────────────────────────────────────────────────────

def _resolve_thresholds(series: Dict[str, Any],
                        thresholds: Dict[str, Any]) -> tuple:
    """
    Determine (warning, critical) threshold values for a series.

    A series may carry explicit ``warning``/``critical`` keys. Otherwise we
    infer them from the config thresholds based on the metric label/unit
    (cpu / memory / disk percentages). Returns ``(None, None)`` when nothing
    applies (e.g. network byte counts).
    """
    warning = series.get("warning")
    critical = series.get("critical")
    if warning is not None or critical is not None:
        return warning, critical

    label = str(series.get("label", "")).lower()
    unit = str(series.get("unit", "") or "")
    if unit != "%":
        return None, None

    if "cpu" in label:
        return thresholds.get("cpu_warning"), thresholds.get("cpu_critical")
    if "mem" in label:
        return thresholds.get("memory_warning"), thresholds.get("memory_critical")
    if "disk" in label:
        return thresholds.get("disk_warning"), thresholds.get("disk_critical")
    return None, None
