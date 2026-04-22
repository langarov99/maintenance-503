import easyocr
import numpy as np
from PIL import Image
from pathlib import Path


# Maps short codes to EasyOCR language codes
LANGUAGE_MAP = {
    "bg": "bg",
    "en": "en",
    "pl": "pl",
    "cs": "cs",
    "it": "it",
    "de": "de",
}


class ImageExtractor:
    def __init__(self, languages: list[str] = None):
        lang_codes = [LANGUAGE_MAP.get(l, l) for l in (languages or ["bg", "en"])]
        # EasyOCR reader is heavy — initialize once
        self.reader = easyocr.Reader(lang_codes, gpu=True, model_storage_directory="models")

    def extract(self, file_path: str) -> dict:
        img = Image.open(file_path).convert("RGB")
        return self.extract_from_pil(img)

    def extract_from_pil(self, img: Image.Image) -> dict:
        results = self.reader.readtext(np.array(img), detail=1, paragraph=False)
        lines = [text for (_, text, conf) in results if conf > 0.3]
        return {"text": "\n".join(lines), "tables": [], "source": "ocr"}
