import chardet
from pathlib import Path


class TextExtractor:
    def extract(self, file_path: str) -> dict:
        encoding = self._detect_encoding(file_path)
        try:
            with open(file_path, "r", encoding=encoding, errors="replace") as f:
                text = f.read()
        except Exception as e:
            return {"text": "", "tables": [], "source": "text", "error": str(e)}

        tables = self._try_parse_table(text)
        return {"text": text, "tables": tables, "source": "text"}

    def _detect_encoding(self, file_path: str) -> str:
        try:
            with open(file_path, "rb") as f:
                raw = f.read(20000)
            result = chardet.detect(raw)
            return result.get("encoding") or "utf-8"
        except Exception:
            return "utf-8"

    def _try_parse_table(self, text: str) -> list:
        lines = [l for l in text.splitlines() if l.strip()]
        if not lines:
            return []

        # Detect delimiter from first non-empty line
        sample = lines[0]
        for delimiter in ["\t", ";", "|", ","]:
            if delimiter in sample and sample.count(delimiter) >= 2:
                rows = [line.split(delimiter) for line in lines]
                return [rows]

        return []
