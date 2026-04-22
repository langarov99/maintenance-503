import logging
import pytesseract
from PIL import Image

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

# Common Tesseract install paths on Windows
_TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Users\Public\Tesseract-OCR\tesseract.exe",
]


def _find_tesseract():
    import shutil, os
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
            text = pytesseract.image_to_string(img, lang=self.lang_str,
                                               config="--psm 6")
            lines = [l for l in text.splitlines() if l.strip()]
            logger.info("OCR: %d lines extracted", len(lines))
            return {"text": "\n".join(lines), "tables": [], "source": "ocr"}
        except Exception as e:
            logger.error("OCR failed: %s", e)
            return {"text": "", "tables": [], "source": "ocr_error"}
