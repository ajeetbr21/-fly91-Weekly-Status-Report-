"""
Screenshot-style visual helper for the AWS Weekly BAU Report.

Renders console-screenshot-style PNG tables (as if captured from the AWS
Inspector / GuardDuty console 'Findings' view) entirely in memory using
Pillow (PIL). The resulting PNG can be embedded into an Excel worksheet via
openpyxl.drawing.image.Image.

The module is intentionally dependency-light (only Pillow) and degrades
gracefully: if Pillow is unavailable or any drawing step fails, the public
helper logs a warning and returns ``None`` so the caller can simply skip
embedding instead of aborting the report.
"""

from io import BytesIO
from typing import List, Optional, Sequence

from utils.logger import get_logger

logger = get_logger("bau_report")

# ── Optional Pillow import (graceful fallback) ─────────────────────────────
try:
    from PIL import Image, ImageDraw, ImageFont  # type: ignore
    _PIL_AVAILABLE = True
except Exception as exc:  # pragma: no cover - exercised only without Pillow
    Image = ImageDraw = ImageFont = None  # type: ignore
    _PIL_AVAILABLE = False
    logger.warning(
        "Pillow (PIL) is not available - screenshot-style visuals will be "
        "skipped: %s", exc
    )


# ── Colour palette (mirrors report.styles, as RGB tuples) ──────────────────
_TITLE_BG = (0, 32, 96)          # DARK_BLUE
_HEADER_BG = (68, 114, 196)      # LIGHT_BLUE
_WHITE = (255, 255, 255)
_BLACK = (33, 37, 41)
_ROW_EVEN = (255, 255, 255)
_ROW_ODD = (242, 242, 242)       # LIGHT_GRAY
_BORDER = (191, 191, 191)        # BORDER_GRAY

# Severity chip colours (High=red, Medium=orange, Low=blue)
_SEVERITY_CHIP = {
    "CRITICAL": (158, 0, 6),
    "HIGH": (214, 40, 40),
    "MEDIUM": (237, 125, 49),
    "LOW": (68, 114, 196),
    "INFORMATIONAL": (108, 117, 125),
}
_SEVERITY_DEFAULT_CHIP = (108, 117, 125)

# ── Layout constants (pixels) ──────────────────────────────────────────────
_PADDING = 12
_TITLE_H = 40
_HEADER_H = 30
_ROW_H = 26
_CELL_PAD_X = 10
_MIN_COL_W = 70
_MAX_COL_W = 320
_CHIP_PAD_X = 8
_CHIP_PAD_Y = 4


def _load_font(size: int, bold: bool = False):
    """Load a TrueType font if one is available, else the PIL default font."""
    candidates = [
        "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
        else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, size)
        except Exception:
            continue
    try:
        return ImageFont.load_default()
    except Exception:
        return None


def _text_size(draw, text, font):
    """Return (width, height) of *text* for both new and old Pillow APIs."""
    text = str(text)
    try:
        left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
        return right - left, bottom - top
    except Exception:
        try:
            return draw.textlength(text, font=font), font.size  # type: ignore
        except Exception:
            # Rough fallback estimate
            return len(text) * 7, 12


# Maximum pixel dimension (largest side) for any image embedded into a
# report. Oversized assets are downscaled to this cap so the embedded pixel
# count stays well under Pillow's DecompressionBomb limit (~89.5M px) and the
# image is a sensible size for a document.
_MAX_EMBED_DIMENSION = 2000


def safe_image_for_embedding(
    path: str,
    max_dimension: int = _MAX_EMBED_DIMENSION,
) -> Optional[BytesIO]:
    """Load an on-disk image and return a safe-to-embed PNG ``BytesIO``.

    Guards against Pillow ``DecompressionBombWarning`` / ``DecompressionBombError``
    on oversized images (e.g. a user drops a huge screenshot into ``assets/``):
    the image is downscaled so its largest dimension is at most
    *max_dimension*, keeping the pixel count well under Pillow's safety limit.

    The function degrades gracefully, mirroring the rest of this module: if
    Pillow is unavailable or anything goes wrong it logs a warning and returns
    ``None`` so the caller can skip that single image instead of aborting the
    report.

    Parameters
    ----------
    path : str
        Path to the source image on disk.
    max_dimension : int, optional
        Maximum allowed size (in pixels) for the largest side of the returned
        image. Defaults to :data:`_MAX_EMBED_DIMENSION`.

    Returns
    -------
    io.BytesIO | None
        A ``BytesIO`` positioned at 0 containing PNG bytes, or ``None`` if the
        image could not be safely loaded.
    """
    if not _PIL_AVAILABLE:
        logger.warning(
            "Skipping image '%s' - Pillow is not installed.", path
        )
        return None

    # Temporarily raise Pillow's decompression-bomb guard only so we can OPEN
    # the file to inspect/downscale it. We never leave it disabled, and we
    # always downscale oversized images rather than embedding them raw.
    prev_limit = getattr(Image, "MAX_IMAGE_PIXELS", None)
    try:
        try:
            Image.MAX_IMAGE_PIXELS = None  # allow opening to downscale
            with Image.open(path) as im:
                im.load()
                src = im.convert("RGB") if im.mode not in ("RGB", "RGBA") else im.copy()
        finally:
            Image.MAX_IMAGE_PIXELS = prev_limit

        width, height = src.size
        largest = max(width, height)
        if largest > max_dimension and largest > 0:
            scale = float(max_dimension) / float(largest)
            new_size = (
                max(1, int(width * scale)),
                max(1, int(height * scale)),
            )
            resample = getattr(getattr(Image, "Resampling", Image), "LANCZOS",
                               getattr(Image, "LANCZOS", 1))
            src = src.resize(new_size, resample)

        buf = BytesIO()
        src.save(buf, format="PNG")
        buf.seek(0)
        return buf
    except Exception as exc:
        logger.warning(
            "Failed to load image '%s' for embedding - skipping: %s",
            path, exc
        )
        return None


def render_findings_table(
    title: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[object]],
    severity_col: Optional[int] = None,
) -> Optional[BytesIO]:
    """
    Render a console-screenshot-style findings table to an in-memory PNG.

    Parameters
    ----------
    title : str
        Text drawn in the dark-blue title bar.
    headers : sequence of str
        Column header labels.
    rows : sequence of sequences
        Row values (coerced to ``str`` for display).
    severity_col : int, optional
        Zero-based column index whose values should be drawn as colour-coded
        severity chips (High=red / Medium=orange / Low=blue). ``None`` disables
        chip rendering.

    Returns
    -------
    io.BytesIO | None
        A ``BytesIO`` positioned at 0 containing PNG bytes, or ``None`` if
        Pillow is unavailable or rendering failed (a warning is logged).
    """
    if not _PIL_AVAILABLE:
        logger.warning(
            "Skipping screenshot '%s' - Pillow is not installed.", title
        )
        return None

    try:
        headers = list(headers)
        rows = [list(r) for r in rows]
        ncols = max(len(headers), *(len(r) for r in rows)) if rows else len(headers)

        title_font = _load_font(18, bold=True)
        header_font = _load_font(13, bold=True)
        cell_font = _load_font(12)
        chip_font = _load_font(11, bold=True)

        # Measure column widths on a throwaway canvas.
        probe = ImageDraw.Draw(Image.new("RGB", (10, 10)))

        col_widths = [_MIN_COL_W] * ncols
        for c in range(ncols):
            header_text = headers[c] if c < len(headers) else ""
            w, _ = _text_size(probe, header_text, header_font)
            col_widths[c] = max(col_widths[c], w + 2 * _CELL_PAD_X)
            for r in rows:
                val = r[c] if c < len(r) else ""
                w, _ = _text_size(probe, val, cell_font)
                extra = 2 * _CHIP_PAD_X if (severity_col == c) else 0
                col_widths[c] = max(
                    col_widths[c], w + 2 * _CELL_PAD_X + extra
                )
            col_widths[c] = min(col_widths[c], _MAX_COL_W)

        table_w = sum(col_widths)
        img_w = table_w + 2 * _PADDING
        img_h = (
            _PADDING + _TITLE_H + _HEADER_H + len(rows) * _ROW_H + _PADDING
        )

        img = Image.new("RGB", (int(img_w), int(img_h)), _WHITE)
        draw = ImageDraw.Draw(img)

        x0 = _PADDING
        y = _PADDING

        # ── Title bar ──────────────────────────────────────────────
        draw.rectangle(
            [x0, y, x0 + table_w, y + _TITLE_H], fill=_TITLE_BG
        )
        _draw_text_vcenter(draw, x0 + _CELL_PAD_X, y, _TITLE_H, title,
                           title_font, _WHITE)
        y += _TITLE_H

        # ── Header row ─────────────────────────────────────────────
        cx = x0
        for c in range(ncols):
            draw.rectangle(
                [cx, y, cx + col_widths[c], y + _HEADER_H],
                fill=_HEADER_BG, outline=_BORDER,
            )
            label = headers[c] if c < len(headers) else ""
            _draw_text_vcenter(draw, cx + _CELL_PAD_X, y, _HEADER_H, label,
                               header_font, _WHITE)
            cx += col_widths[c]
        y += _HEADER_H

        # ── Data rows (zebra striping) ─────────────────────────────
        for ri, row in enumerate(rows):
            bg = _ROW_EVEN if ri % 2 == 0 else _ROW_ODD
            draw.rectangle([x0, y, x0 + table_w, y + _ROW_H], fill=bg)
            cx = x0
            for c in range(ncols):
                draw.rectangle(
                    [cx, y, cx + col_widths[c], y + _ROW_H],
                    outline=_BORDER,
                )
                val = str(row[c]) if c < len(row) else ""
                if severity_col == c and val.strip():
                    _draw_severity_chip(draw, cx, y, col_widths[c], val,
                                        chip_font)
                else:
                    val = _truncate(draw, val, cell_font,
                                    col_widths[c] - 2 * _CELL_PAD_X)
                    _draw_text_vcenter(draw, cx + _CELL_PAD_X, y, _ROW_H, val,
                                       cell_font, _BLACK)
                cx += col_widths[c]
            y += _ROW_H

        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf
    except Exception as exc:
        logger.warning(
            "Failed to render screenshot '%s' - skipping embed: %s",
            title, exc
        )
        return None


# ── Private drawing helpers ────────────────────────────────────────────────

def _draw_text_vcenter(draw, x, y_top, box_h, text, font, fill):
    """Draw *text* vertically centred within a box of height *box_h*."""
    _, th = _text_size(draw, text, font)
    ty = y_top + max(0, (box_h - th) // 2)
    draw.text((x, ty), str(text), font=font, fill=fill)


def _draw_severity_chip(draw, cell_x, cell_y, cell_w, label, font):
    """Draw a rounded colour chip for a severity value inside a cell."""
    sev = (label or "").strip().upper()
    colour = _SEVERITY_CHIP.get(sev, _SEVERITY_DEFAULT_CHIP)
    display = label if label else sev.title()
    tw, th = _text_size(draw, display, font)
    chip_w = tw + 2 * _CHIP_PAD_X
    chip_h = th + 2 * _CHIP_PAD_Y
    chip_x = cell_x + _CELL_PAD_X
    chip_y = cell_y + max(0, (_ROW_H - chip_h) // 2)
    box = [chip_x, chip_y, chip_x + chip_w, chip_y + chip_h]
    try:
        draw.rounded_rectangle(box, radius=6, fill=colour)
    except Exception:
        draw.rectangle(box, fill=colour)
    draw.text((chip_x + _CHIP_PAD_X, chip_y + _CHIP_PAD_Y), str(display),
              font=font, fill=_WHITE)


def _truncate(draw, text, font, max_w):
    """Truncate *text* with an ellipsis so it fits within *max_w* pixels."""
    text = str(text)
    w, _ = _text_size(draw, text, font)
    if w <= max_w or max_w <= 0:
        return text
    ellipsis = "..."
    while text and w > max_w:
        text = text[:-1]
        w, _ = _text_size(draw, text + ellipsis, font)
    return (text + ellipsis) if text else ellipsis
