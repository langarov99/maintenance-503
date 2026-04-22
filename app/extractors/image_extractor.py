import logging
import os
from pathlib import Path
import pytesseract
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
    # 1. Already configured
    try:
        if pytesseract.get_tesseract_version():
            return
    except Exception:
        pass
    # 2. Try known Windows paths
    for p in _TESSERACT_PATHS:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            logger.info("Tesseract found at: %s", p)
            return
    # 3. Try PATH
    found = shutil.which("tesseract")
    if found:
        pytesseract.pytesseract.tesseract_cmd = found
        logger.info("Tesseract found in PATH: %s", found)


_find_tesseract()


def _preprocess(img: Image.Image) -> Image.Image:
    """Upscale + enhance contrast to improve Tesseract accuracy on scanned invoices."""
    # Upscale small images — Tesseract works best at ~300 DPI
    w, h = img.size
    if w < 2000:
        scale = max(2, 2400 // max(w, 1))
        img = img.resize((w * scale, h * scale), Image.LANCZOS)

    # Grayscale → easier for Tesseract
    img = img.convert("L")

    # Boost contrast so table lines don't bleed into text
    img = ImageEnhance.Contrast(img).enhance(1.8)

    # Light sharpening
    img = img.filter(ImageFilter.SHARPEN)

    return img


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
            # PSM 4 = single-column layout (better for invoices with wide table rows)
            text = pytesseract.image_to_string(img, lang=self.lang_str,
                                               config="--psm 4 --oem 3")
            lines = [l for l in text.splitlines() if l.strip()]
            logger.info("OCR: %d lines extracted", len(lines))
            return {"text": "\n".join(lines), "tables": [], "source": "ocr"}
        except Exception as e:
            logger.error("OCR failed: %s", e)
            return {"text": "", "tables": [], "source": "ocr_error"}
