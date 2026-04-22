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
            img = _preprocess(img)
            text = _extract_structured_text(img, self.lang_str)
            lines = [l for l in text.splitlines() if l.strip()]
            logger.info("OCR: %d lines extracted", len(lines))
            return {"text": "\n".join(lines), "tables": [], "source": "ocr"}
        except Exception as e:
            logger.error("OCR failed: %s", e)
            return {"text": "", "tables": [], "source": "ocr_error"}
