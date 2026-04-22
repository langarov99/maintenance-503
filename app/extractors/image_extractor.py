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
    # Clamp: never go below 100 or above 200 to avoid all-black/all-white
    t = max(100, min(200, t))
    return img.point(lambda p: 255 if p > t else 0, 'L')


def _extract_structured_text(img: Image.Image, lang: str) -> str:
    """Use word bounding boxes to reconstruct properly aligned text lines.

    Tesseract's image_to_string loses table structure; image_to_data returns
    per-word (x, y) coordinates so we can group words by row ourselves.
    """
    data = pytesseract.image_to_data(
        img, lang=lang,
        config="--psm 6 --oem 3",
        output_type=Output.DICT,
    )

    words = []
    for i in range(len(data["text"])):
        text = str(data["text"][i]).strip()
        conf = int(data["conf"][i])
        if text and conf > 10:
            words.append({
                "text": text,
                "x": data["left"][i],
                "y": data["top"][i],
                "h": data["height"][i],
            })

    if not words:
        return ""

    # Estimate average character height for row-grouping tolerance
    avg_h = sum(w["h"] for w in words) / len(words)
    tolerance = max(8, int(avg_h * 0.6))

    # Sort words top-to-bottom, left-to-right
    words.sort(key=lambda w: (w["y"], w["x"]))

    # Group into rows: words whose y-tops differ by less than tolerance
    # share the same row
    rows: list[list[dict]] = []
    current: list[dict] = []
    row_y = words[0]["y"]

    for w in words:
        if w["y"] - row_y > tolerance:
            if current:
                rows.append(sorted(current, key=lambda c: c["x"]))
            current = [w]
            row_y = w["y"]
        else:
            current.append(w)
    if current:
        rows.append(sorted(current, key=lambda c: c["x"]))

    lines = [" ".join(w["text"] for w in row) for row in rows]
    logger.info("Structured OCR: %d raw words → %d lines", len(words), len(lines))
    return "\n".join(lines)


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
            # Pass A: soft contrast + PSM 4 (single column — code+price stay on same line)
            img_a = _preprocess(img)
            text_a = pytesseract.image_to_string(img_a, lang=self.lang_str,
                                                 config="--psm 4 --oem 3")
            lines_a = [l for l in text_a.splitlines() if l.strip()]

            # Pass B: hard Otsu binarization + PSM 6 (uniform block — catches faint rows)
            img_b = _preprocess_binary(img)
            text_b = pytesseract.image_to_string(img_b, lang=self.lang_str,
                                                 config="--psm 6 --oem 3")
            lines_b = [l for l in text_b.splitlines() if l.strip()]

            logger.info("OCR dual-pass: A=%d lines, B=%d lines", len(lines_a), len(lines_b))
            return {
                "text":  "\n".join(lines_a),
                "text2": "\n".join(lines_b),
                "tables": [],
                "source": "ocr_dual",
            }
        except Exception as e:
            logger.error("OCR failed: %s", e)
            return {"text": "", "text2": "", "tables": [], "source": "ocr_error"}
