import pandas as pd
from pathlib import Path


class ExcelExtractor:
    def extract(self, file_path: str) -> dict:
        suffix = Path(file_path).suffix.lower()
        try:
            if suffix in (".xlsx", ".xlsm"):
                dfs = pd.read_excel(file_path, sheet_name=None, engine="openpyxl", header=None)
            elif suffix in (".xls",):
                dfs = pd.read_excel(file_path, sheet_name=None, engine="xlrd", header=None)
            elif suffix == ".csv":
                encoding = self._detect_encoding(file_path)
                df = pd.read_csv(file_path, header=None, encoding=encoding, sep=None, engine="python")
                dfs = {"Sheet1": df}
            else:
                return {"text": "", "tables": [], "source": "excel"}

            tables = []
            text_lines = []
            for sheet_name, df in dfs.items():
                df = df.fillna("").astype(str)
                rows = df.values.tolist()
                tables.append(rows)
                for row in rows:
                    text_lines.append("\t".join(str(c) for c in row))

            return {"text": "\n".join(text_lines), "tables": tables, "source": "excel"}
        except Exception as e:
            return {"text": "", "tables": [], "source": "excel", "error": str(e)}

    def _detect_encoding(self, file_path: str) -> str:
        try:
            import chardet
            with open(file_path, "rb") as f:
                result = chardet.detect(f.read(10000))
            return result.get("encoding") or "utf-8"
        except Exception:
            return "utf-8"
