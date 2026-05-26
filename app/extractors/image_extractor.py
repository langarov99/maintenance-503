import logging
import os
from pathlib import Path
import pytesseract
from pytesseract import Output
from PIL import Image, ImageEnhance, ImageFilter
try:
    import cv2
    import numpy as np
    _OPENCV_AVAILABLE = True
except ImportError:
    _OPENCV_AVAILABLE = False

logger = logging.getLogger(__name__)

# Maps app language codes to Tesseract language codes
LANGUAGE_MAP = {
    "bg": "bul",
    "en": "eng",
    "pl": "pol",
    "cs": "ces",
    "it": "ita",
    "de": "deu",
}

# Base dir = root of the project (two levels up from this file)
_BASE_DIR = Path(__file__).parent.parent.parent

# Portable install next to the project takes priority over system installs
_TESSERACT_PATHS = [
    str(_BASE_DIR / "Tesseract-OCR" / "tesseract.exe"),   # portable on flash drive
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


def _find_tesseract():
    import shutil
    try:
        if pytesseract.get_tesseract_version():
            return
    except Exception:
        pass
    for p in _TESSERACT_PATHS:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            logger.info("Tesseract found at: %s", p)
            return
    found = shutil.which("tesseract")
    if found:
        pytesseract.pytesseract.tesseract_cmd = found
        logger.info("Tesseract found in PATH: %s", found)


_find_tesseract()


def _preprocess(img: Image.Image) -> Image.Image:
    """Upscale + enhance contrast to improve Tesseract accuracy on scanned invoices.

    Target: ≥3200 px wide so Tesseract has enough pixels per character.
    At 300 DPI an A4 page is ~2480 px — still below the ideal Tesseract target,
    so we upscale it to ~3200 px (scale ×2 for low-res, ×1.3 for 300-DPI sources).
    """
    w, h = img.size
    if w < 3200:
        scale = max(2, 3200 // max(w, 1))
        img = img.resize((w * scale, h * scale), Image.LANCZOS)
    img = img.convert("L")
    img = ImageEnhance.Contrast(img).enhance(1.8)
    img = img.filter(ImageFilter.SHARPEN)
    return img


def _otsu_threshold(img: Image.Image) -> int:
    """Compute Otsu's optimal binarization threshold from image histogram."""
    hist = img.histogram()
    total = sum(hist)
    sum_total = sum(i * h for i, h in enumerate(hist))
    best_t, best_var, sum_b, weight_b = 0, 0.0, 0, 0
    for t in range(256):
        weight_b += hist[t]
        if weight_b == 0 or weight_b == total:
            continue
        weight_f = total - weight_b
        sum_b += t * hist[t]
        mean_b = sum_b / weight_b
        mean_f = (sum_total - sum_b) / weight_f
        var = weight_b * weight_f * (mean_b - mean_f) ** 2
        if var > best_var:
            best_var, best_t = var, t
    return best_t


def _preprocess_binary(img: Image.Image) -> Image.Image:
    """Hard Otsu binarization — better for dense table grids and faint ink."""
    w, h = img.size
    if w < 3200:
        scale = max(2, 3200 // max(w, 1))
        img = img.resize((w * scale, h * scale), Image.LANCZOS)
    img = img.convert("L")
    img = ImageEnhance.Contrast(img).enhance(2.5)
    t = _otsu_threshold(img)
    t = max(100, min(200, t))
    return img.point(lambda p: 255 if p > t else 0, 'L')


def _find_grid_lines(binary_arr, axis: int,
                     min_dark: float = 0.25,
                     min_gap: int = 15) -> list:
    """Return pixel positions of grid lines detected along the given axis.

    axis=0 → horizontal lines (high dark-pixel ratio across each row)
    axis=1 → vertical lines  (high dark-pixel ratio across each column)
    """
    import numpy as np
    ratios = binary_arr.mean(axis=1 - axis)
    is_line = ratios > min_dark

    padded = np.concatenate(([False], is_line, [False]))
    changes = np.diff(padded.astype(np.int8))
    starts = np.where(changes == 1)[0]
    ends   = np.where(changes == -1)[0]

    lines = []
    for s, e in zip(starts, ends):
        center = int((s + e) // 2)
        if not lines or center - lines[-1] >= min_gap:
            lines.append(center)
    return lines


def _filter_table_lines(lines: list, tolerance_pct: float = 0.35,
                        min_count: int = 5) -> list:
    """Keep only the longest run of evenly-spaced lines (= the table body).

    Tries every candidate gap value (20-150px range) and picks the one that
    produces the LONGEST consistent run of lines.  A 26-row product table
    will always produce a longer run than header/footer patterns regardless
    of which gap value is most frequent overall.
    """
    if len(lines) < min_count:
        return lines

    gaps = [lines[i + 1] - lines[i] for i in range(len(lines) - 1)]

    # Collect unique candidate gap buckets (5 px resolution) in table-row range
    from collections import Counter
    table_gaps = [g for g in gaps if 20 <= g <= 150]
    if not table_gaps:
        return lines

    candidates = sorted(set(round(g / 5) * 5 for g in table_gaps))

    best_result = lines
    best_run_len = 0
    best_gap = 0

    for bucket in candidates:
        tol = max(6, int(bucket * tolerance_pct))
        curr_start, curr_len = 0, 1
        run_start, run_len = 0, 1
        for i, g in enumerate(gaps):
            if abs(g - bucket) <= tol:
                curr_len += 1
                if curr_len > run_len:
                    run_start, run_len = curr_start, curr_len
            else:
                curr_start = i + 1
                curr_len = 1

        if run_len > best_run_len:
            best_run_len = run_len
            best_gap = bucket
            # +1: include closing line of the last row
            best_result = lines[run_start: run_start + run_len + 1]

    logger.info("Line filter: %d → %d lines (best gap %dpx, tol ±%dpx)",
                len(lines), len(best_result), best_gap,
                max(6, int(best_gap * tolerance_pct)))
    if len(best_result) < min_count:
        return lines
    return best_result


def _cluster_lines(positions: list, gap: int = 20) -> list:
    """Merge nearby pixel positions into a single representative position."""
    if not positions:
        return []
    positions = sorted(set(positions))
    clusters = [[positions[0]]]
    for p in positions[1:]:
        if p - clusters[-1][-1] <= gap:
            clusters[-1].append(p)
        else:
            clusters.append([p])
    return [int(sum(c) / len(c)) for c in clusters]


def _extract_table_cells_cv(img: Image.Image, lang: str) -> list:
    """OpenCV HoughLinesP-based grid detection with per-cell OCR crops.

    Uses cv2.HoughLinesP to find actual line segments, clusters them into
    a grid, then runs pytesseract on individual cell crops for high accuracy.
    """
    # ── Scale image ───────────────────────────────────────────────────────────
    w, h = img.size
    if w < 2000:
        scale = max(2, 2400 // max(w, 1))
        img = img.resize((w * scale, h * scale), Image.LANCZOS)
        w, h = img.size

    gray_arr = np.array(img.convert("L"))

    # ── Adaptive threshold → edges → Hough lines ─────────────────────────────
    # Use THRESH_BINARY_INV so dark lines become white (255) on black background
    blur = cv2.GaussianBlur(gray_arr, (3, 3), 0)
    _, binary = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    # Erode horizontally/vertically to isolate line segments
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 20))
    h_lines_img = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)
    v_lines_img = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel)

    # Detect horizontal line positions (y-coordinates)
    h_segs = cv2.HoughLinesP(h_lines_img, 1, 3.14159/180, threshold=80,
                              minLineLength=w // 4, maxLineGap=60)
    # Detect vertical line positions (x-coordinates)
    v_segs = cv2.HoughLinesP(v_lines_img, 1, 3.14159/180, threshold=60,
                              minLineLength=h // 8, maxLineGap=40)

    if h_segs is None or v_segs is None:
        logger.info("HoughLinesP: insufficient lines (h=%s v=%s) — falling back",
                    0 if h_segs is None else len(h_segs),
                    0 if v_segs is None else len(v_segs))
        return []

    raw_h = [int((seg[0][1] + seg[0][3]) / 2) for seg in h_segs]
    raw_v = [int((seg[0][0] + seg[0][2]) / 2) for seg in v_segs]

    hlines = _cluster_lines(raw_h, gap=15)
    vlines = _cluster_lines(raw_v, gap=30)

    logger.info("HoughLinesP: %d h-segs → %d hlines, %d v-segs → %d vlines",
                len(h_segs), len(hlines), len(v_segs), len(vlines))

    if len(hlines) < 5 or len(vlines) < 3:
        logger.info("Not enough grid lines (h=%d, v=%d) — falling back",
                    len(hlines), len(vlines))
        return []

    # ── Filter to main table body ─────────────────────────────────────────────
    hlines = _filter_table_lines(sorted(hlines), tolerance_pct=0.40, min_count=5)
    vlines = sorted(vlines)

    n_rows = len(hlines) - 1
    n_cols = len(vlines) - 1
    logger.info("Grid: %d rows × %d cols", n_rows, n_cols)

    if n_rows < 3 or n_cols < 2:
        return []

    # ── OCR per cell ──────────────────────────────────────────────────────────
    pil_gray = Image.fromarray(gray_arr)
    result = []
    pad = 3  # pixel padding inside each cell to avoid border noise

    for r in range(n_rows):
        row_cells = []
        y1 = hlines[r] + pad
        y2 = hlines[r + 1] - pad
        if y2 - y1 < 5:
            row_cells = [""] * n_cols
            result.append(row_cells)
            continue
        for c in range(n_cols):
            x1 = vlines[c] + pad
            x2 = vlines[c + 1] - pad
            if x2 - x1 < 5:
                row_cells.append("")
                continue
            cell_img = pil_gray.crop((x1, y1, x2, y2))
            # Upscale small cells for better OCR
            cw, ch = cell_img.size
            if cw < 100 or ch < 20:
                cell_img = cell_img.resize((max(cw, 100), max(ch, 30)), Image.LANCZOS)
            cell_img = ImageEnhance.Contrast(cell_img).enhance(2.0)
            text = pytesseract.image_to_string(
                cell_img, lang=lang,
                config="--psm 7 --oem 3 -c tessedit_char_blacklist=|"
            ).strip()
            row_cells.append(text)
        if any(row_cells):
            result.append(row_cells)

    logger.info("Cell grid built: %d rows × %d cols", len(result), n_cols)
    for ri, row in enumerate(result[:6]):
        logger.info("  row %d: %s", ri, " | ".join(f'[{c[:20]}]' for c in row))

    return result


def _extract_table_cells(img: Image.Image, lang: str) -> list:
    """Detect invoice table grid and extract cell text.

    Tries OpenCV HoughLinesP first (more reliable for scanned images),
    falls back to numpy dark-ratio method if OpenCV is unavailable.
    """
    if _OPENCV_AVAILABLE:
        try:
            result = _extract_table_cells_cv(img, lang)
            if result:
                return result
            logger.info("OpenCV grid extraction returned empty — trying numpy fallback")
        except Exception as e:
            logger.warning("OpenCV grid extraction failed: %s — trying numpy fallback", e)

    return _extract_table_cells_numpy(img, lang)


def _extract_table_cells_numpy(img: Image.Image, lang: str) -> list:
    """Numpy dark-ratio fallback for grid detection (used when OpenCV unavailable)."""
    try:
        import numpy as _np
    except ImportError:
        logger.warning("numpy not available — skipping grid table extraction")
        return []

    # ── 1. Scale & binarise ──────────────────────────────────────────────────
    w, h = img.size
    if w < 2000:
        scale = max(2, 2400 // max(w, 1))
        img = img.resize((w * scale, h * scale), Image.LANCZOS)

    gray = img.convert("L")
    hi = ImageEnhance.Contrast(gray).enhance(3.5)
    arr = _np.array(hi, dtype=_np.float32) / 255.0
    binary = (arr < 0.35).astype(_np.float32)

    hlines_all = _find_grid_lines(binary, axis=0, min_dark=0.20, min_gap=15)
    logger.info("Horizontal lines (raw): %d", len(hlines_all))
    hlines = _filter_table_lines(hlines_all, tolerance_pct=0.40, min_count=5)

    if len(hlines) < 5:
        logger.info("Not enough horizontal lines after filtering (%d) — falling back",
                    len(hlines))
        return []

    y_top    = max(0, hlines[0] - 5)
    y_bottom = min(binary.shape[0], hlines[-1] + 5)
    table_strip = binary[y_top:y_bottom, :]
    vlines = _find_grid_lines(table_strip, axis=1, min_dark=0.40, min_gap=150)
    logger.info("Vertical lines (table area): %d", len(vlines))

    if len(vlines) < 3:
        logger.info("Not enough vertical lines (%d) — falling back", len(vlines))
        return []

    ocr_img = ImageEnhance.Contrast(gray).enhance(2.0)
    data = pytesseract.image_to_data(
        ocr_img, lang=lang,
        config="--psm 6 --oem 3",
        output_type=Output.DICT,
    )

    n_rows = len(hlines) - 1
    n_cols = len(vlines) - 1
    grid = [[[] for _ in range(n_cols)] for _ in range(n_rows)]

    for i, word in enumerate(data["text"]):
        word = str(word).strip()
        if not word or int(data["conf"][i]) <= 0:
            continue
        wx = data["left"][i] + data["width"][i] // 2
        wy = data["top"][i] + data["height"][i] // 2
        row_idx = next((r for r in range(n_rows) if hlines[r] <= wy < hlines[r + 1]), None)
        col_idx = next((c for c in range(n_cols) if vlines[c] <= wx < vlines[c + 1]), None)
        if row_idx is not None and col_idx is not None:
            grid[row_idx][col_idx].append(word)

    result = []
    for row in grid:
        cells = [" ".join(words).strip() for words in row]
        if any(cells):
            result.append(cells)

    logger.info("Cell grid (numpy): %d rows × %d cols", len(result), n_cols)
    return result


class ImageExtractor:
    def __init__(self, languages: list[str] = None):
        lang_codes = [LANGUAGE_MAP.get(l, "eng") for l in (languages or ["bg", "en"])]
        self.lang_str = "+".join(lang_codes)
        logger.info("ImageExtractor ready (pytesseract, langs=%s)", self.lang_str)

    def extract(self, file_path: str) -> dict:
        img = Image.open(file_path).convert("RGB")
        return self.extract_from_pil(img)

    def extract_from_pil(self, img: Image.Image) -> dict:
        try:
            # ── Grid-based cell extraction (best for table invoices) ──────────
            cell_table = _extract_table_cells(img, self.lang_str)

            # ── Pass A: soft contrast + PSM 4 (text fallback) ────────────────
            img_a = _preprocess(img)
            text_a = pytesseract.image_to_string(img_a, lang=self.lang_str,
                                                 config="--psm 4 --oem 3")
            lines_a = [l for l in text_a.splitlines() if l.strip()]

            # ── Pass B: Otsu binary + PSM 6 (catches faint rows) ─────────────
            img_b = _preprocess_binary(img)
            text_b = pytesseract.image_to_string(img_b, lang=self.lang_str,
                                                 config="--psm 6 --oem 3")
            lines_b = [l for l in text_b.splitlines() if l.strip()]

            logger.info("OCR: grid=%d rows, textA=%d lines, textB=%d lines",
                        len(cell_table), len(lines_a), len(lines_b))
            return {
                "text":   "\n".join(lines_a),
                "text2":  "\n".join(lines_b),
                "tables": [cell_table] if cell_table else [],
                "source": "ocr_grid" if cell_table else "ocr_dual",
            }
        except Exception as e:
            logger.error("OCR failed: %s", e)
            return {"text": "", "text2": "", "tables": [], "source": "ocr_error"}
