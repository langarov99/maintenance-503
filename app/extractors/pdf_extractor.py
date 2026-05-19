import logging
import re
import pdfplumber
from PIL import Image
import io
from .image_extractor import ImageExtractor

logger = logging.getLogger(__name__)

try:
    import fitz  # PyMuPDF (legacy import name)
    _FITZ_AVAILABLE = True
except Exception:
    try:
        import pymupdf as fitz  # PyMuPDF >= 1.24 new import name
        _FITZ_AVAILABLE = True
    except Exception:
        _FITZ_AVAILABLE = False


def _is_cid_garbage(text: str) -> bool:
    """Return True when text is mostly (cid:XX) sequences — undecodable custom font."""
    if not text or len(text) < 20:
        return False
    cid_hits = len(re.findall(r'\(cid:\d+\)', text))
    return cid_hits > 10 and cid_hits * 7 > len(text) * 0.3


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
        text   = self._extract_text(file_path)
        tables = self._extract_tables(file_path)

        if not text.strip() or _is_cid_garbage(text):
            # pdfplumber produced garbage — try PyMuPDF text first
            if _FITZ_AVAILABLE:
                fitz_text = self._extract_text_via_fitz(file_path)
                if fitz_text.strip() and not _is_cid_garbage(fitz_text):
                    logger.info("CID garbage detected — switched to PyMuPDF text")
                    text   = fitz_text
                    source = "text_fitz"
                    # Also re-try tables via fitz if pdfplumber found none
                    if not tables:
                        tables = self._extract_tables_via_fitz(file_path)
                else:
                    logger.info("CID garbage detected — falling back to OCR")
                    text   = self._extract_via_ocr(file_path)
                    source = "ocr"
            else:
                logger.info("CID garbage detected — falling back to OCR (no fitz)")
                text   = self._extract_via_ocr(file_path)
                source = "ocr"
        else:
            source = "text"

        lines = [l for l in text.splitlines() if l.strip()]
        logger.info("PDF extracted: source=%s  chars=%d  lines=%d",
                    source, len(text), len(lines))
        return {"text": text, "tables": tables, "source": source}

    def _extract_text(self, file_path: str) -> str:
        pages_text = []
        try:
            with pdfplumber.open(file_path) as pdf:
                total = len(pdf.pages)
                for page_num, page in enumerate(pdf.pages, start=1):
                    page_text = page.extract_text() or ""
                    pages_text.append(page_text)
                    logger.info("  Page %d/%d: %d chars", page_num, total, len(page_text))
        except Exception as e:
            logger.error("PDF text extraction error: %s", e)
        return "\n".join(pages_text)

    def _extract_text_via_fitz(self, file_path: str) -> str:
        """Extract text using PyMuPDF — handles custom/embedded fonts better."""
        doc = fitz.open(file_path)
        pages_text = []
        total = len(doc)
        for page_num, page in enumerate(doc, start=1):
            page_text = page.get_text("text") or ""
            pages_text.append(page_text)
            logger.info("  fitz Page %d/%d: %d chars", page_num, total, len(page_text))
        doc.close()
        return "\n".join(pages_text)

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

    def _extract_tables_via_fitz(self, file_path: str) -> list[list[list]]:
        """Extract tables via PyMuPDF (available in fitz >= 1.23)."""
        tables = []
        try:
            doc = fitz.open(file_path)
            for page in doc:
                for tab in page.find_tables():
                    rows = tab.extract()
                    if rows:
                        tables.append(rows)
            doc.close()
        except Exception as e:
            logger.debug("fitz table extraction failed: %s", e)
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
