import logging
import os
from pathlib import Path
import pytesseract
from pytesseract import Output
from PIL import Image, ImageEnhance, ImageFilter

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
    """Upscale + enhance contrast to improve Tesseract accuracy on scanned invoices."""
    w, h = img.size
    if w < 2000:
        scale = max(2, 2400 // max(w, 1))
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
    if w < 2000:
        scale = max(2, 2400 // max(w, 1))
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


def _extract_table_cells(img: Image.Image, lang: str) -> list:
    """Detect invoice table grid via line analysis and assign OCR words to cells.

    Steps
    -----
    1. High-contrast binary array  (numpy)
    2. Horizontal lines with HIGH threshold (≥50 % dark) — only thick table borders
    3. Filter to the longest evenly-spaced run  → table body rows only
    4. Vertical lines detected INSIDE the table area only
    5. ONE pytesseract.image_to_data() call → word bounding boxes
    6. Each word assigned to its cell by comparing word centre to grid boundaries
    7. Return list[list[str]] ready for _parse_mafra_table()
    """
    try:
        import numpy as np
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
    arr = np.array(hi, dtype=np.float32) / 255.0
    binary = (arr < 0.35).astype(np.float32)   # 1.0 = dark pixel

    # ── 2. Horizontal lines ───────────────────────────────────────────────────
    # 0.20 threshold catches thin inner row separators; _filter_table_lines
    # removes false positives from header/footer using most-common-gap logic.
    hlines_all = _find_grid_lines(binary, axis=0, min_dark=0.20, min_gap=15)
    logger.info("Horizontal lines (raw): %d", len(hlines_all))

    # ── 3. Keep only the evenly-spaced run = product table rows ──────────────
    hlines = _filter_table_lines(hlines_all, tolerance_pct=0.40, min_count=5)

    if len(hlines) < 5:
        logger.info("Not enough horizontal lines after filtering (%d) — falling back",
                    len(hlines))
        return []

    # ── 4. Vertical lines — restricted to table y-range, higher threshold ────
    # Threshold 0.40: horizontal grid lines create a ~0.30 dark-ratio "floor".
    # min_gap=150: invoice column separators are often double-ruled; pairs of
    # lines up to ~120 px apart must be merged into one logical separator.
    y_top    = max(0, hlines[0] - 5)
    y_bottom = min(binary.shape[0], hlines[-1] + 5)
    table_strip = binary[y_top:y_bottom, :]
    vlines = _find_grid_lines(table_strip, axis=1, min_dark=0.40, min_gap=150)
    logger.info("Vertical lines (table area): %d", len(vlines))

    if len(vlines) < 3:
        logger.info("Not enough vertical lines (%d) — falling back", len(vlines))
        return []

    # ── 5. Single OCR pass ────────────────────────────────────────────────────
    ocr_img = ImageEnhance.Contrast(gray).enhance(2.0)
    data = pytesseract.image_to_data(
        ocr_img, lang=lang,
        config="--psm 6 --oem 3",
        output_type=Output.DICT,
    )

    # ── 6. Assign words to cells ──────────────────────────────────────────────
    n_rows = len(hlines) - 1
    n_cols = len(vlines) - 1
    grid = [[[] for _ in range(n_cols)] for _ in range(n_rows)]

    for i, word in enumerate(data["text"]):
        word = str(word).strip()
        if not word or int(data["conf"][i]) <= 0:
            continue

        wx = data["left"][i] + data["width"][i] // 2
        wy = data["top"][i] + data["height"][i] // 2

        row_idx = None
        for r in range(n_rows):
            if hlines[r] <= wy < hlines[r + 1]:
                row_idx = r
                break

        col_idx = None
        for c in range(n_cols):
            if vlines[c] <= wx < vlines[c + 1]:
                col_idx = c
                break

        if row_idx is not None and col_idx is not None:
            grid[row_idx][col_idx].append(word)

    # ── 7. Flatten, drop empty rows ───────────────────────────────────────────
    result = []
    for row in grid:
        cells = [" ".join(words).strip() for words in row]
        if any(cells):
            result.append(cells)

    logger.info("Cell grid built: %d rows × %d cols", len(result), n_cols)
    for ri, row in enumerate(result[:6]):
        logger.info("  row %d: %s", ri, " | ".join(f'[{c[:20]}]' for c in row))

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
