import pdfplumber
from PIL import Image
import io
from .image_extractor import ImageExtractor

try:
    import fitz  # PyMuPDF
    _FITZ_AVAILABLE = True
except Exception:
    _FITZ_AVAILABLE = False


class PDFExtractor:
    def __init__(self, ocr_languages: list[str] = None):
        self.ocr_languages = ocr_languages or ["bg", "en"]
        self._image_extractor = None

    @property
    def image_extractor(self):
        if self._image_extractor is None:
            self._image_extractor = ImageExtractor(self.ocr_languages)
        return self._image_extractor

    def extract(self, file_path: str) -> dict:
        text = self._extract_text(file_path)
        tables = self._extract_tables(file_path)

        if not text.strip() and not tables:
            text = self._extract_via_ocr(file_path)
            source = "ocr"
        else:
            source = "text"

        return {"text": text, "tables": tables, "source": source}

    def _extract_text(self, file_path: str) -> str:
        lines = []
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        lines.append(page_text)
        except Exception:
            pass
        return "\n".join(lines)

    def _extract_tables(self, file_path: str) -> list[list[list]]:
        tables = []
        try:
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    for table in page.extract_tables():
                        if table:
                            tables.append(table)
        except Exception:
            pass
        return tables

    def _extract_via_ocr(self, file_path: str) -> str:
        if _FITZ_AVAILABLE:
            return self._ocr_via_fitz(file_path)
        return self._ocr_via_pdfplumber(file_path)

    def _ocr_via_fitz(self, file_path: str) -> str:
        doc = fitz.open(file_path)
        all_text = []
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            result = self.image_extractor.extract_from_pil(img)
            all_text.append(result["text"])
        doc.close()
        return "\n".join(all_text)

    def _ocr_via_pdfplumber(self, file_path: str) -> str:
        # Fallback: convert pages to images via pdfplumber+Pillow
        try:
            import subprocess, tempfile, os
            all_text = []
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    img = page.to_image(resolution=200).original
                    result = self.image_extractor.extract_from_pil(img)
                    all_text.append(result["text"])
            return "\n".join(all_text)
        except Exception:
            return ""
