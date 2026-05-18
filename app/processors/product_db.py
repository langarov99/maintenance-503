import logging
import re
import pandas as pd
from pathlib import Path
from dataclasses import dataclass
from typing import Optional

logger = logging.getLogger(__name__)


@dataclass
class ProductInfo:
    internal_code: Optional[str] = None      # Код (колона A от основния файл)
    supplier_article: Optional[str] = None   # Арт. номер при доставчик (колона B)
    description: Optional[str] = None        # Описание (колона C)
    unit_price: Optional[str] = None         # Ед. цена (колона D)
    price_with_vat: Optional[str] = None     # Цената вкл. ДДС (колона O)
    ean: Optional[str] = None                # EAN номер (колона G от EAN файл)
    main_barcode: Optional[str] = None       # Главен баркод (колона E от EAN файл)


class ProductDatabase:
    def __init__(self, data_dir: str):
        self.data_dir = Path(data_dir)
        # Index by supplier article number (OSRAM code from invoice)
        self._by_supplier_article: dict[str, ProductInfo] = {}
        # Index by internal code
        self._by_internal_code: dict[str, ProductInfo] = {}
        # Index by EAN
        self._by_ean: dict[str, ProductInfo] = {}
        self._loaded = False

    def load(self):
        def _find_file(exact: str, pattern: str) -> Optional[Path]:
            p = self.data_dir / exact
            if p.exists():
                return p
            matches = [f for f in sorted(self.data_dir.glob(pattern))
                       if f.name == exact or not f.name.startswith('~$')]
            if matches:
                logger.info("OSRAM: '%s' not found, using '%s'", exact, matches[0].name)
                return matches[0]
            logger.info("OSRAM: no file matching '%s' in %s", pattern, self.data_dir)
            return None

        main_file = _find_file("osram-products.xlsx", "osram*[!ean]*.xlsx")
        ean_file  = _find_file("osram-export-all-ean-code.xlsx", "osram*ean*.xlsx")

        if main_file:
            try:
                self._load_main(main_file)
                logger.info("Main product file loaded: %d records", len(self._by_internal_code))
            except Exception as e:
                logger.error("Failed to load main product file: %s", e)
        else:
            logger.info("Main product file not found, EAN-only mode")

        if ean_file:
            try:
                self._load_ean(ean_file)
                logger.info("EAN file loaded: %d EAN entries", len(self._by_ean))
            except Exception as e:
                logger.error("Failed to load EAN file: %s", e)
        else:
            logger.warning("EAN file not found")

        self._loaded = True
        logger.info("Product DB ready: %d by code, %d by EAN",
                    len(self._by_internal_code), len(self._by_ean))

    def _load_main(self, path: Path):
        df = pd.read_excel(path, engine="openpyxl", header=0, dtype=str)
        df = df.fillna("")

        # Column positions (0-based): A=0, B=1, C=2, D=3, O=14
        col_code        = df.columns[0]   # A — Код
        col_supplier    = df.columns[1]   # B — Арт. номер при доставчик
        col_description = df.columns[2]   # C — Описание
        col_price       = df.columns[3]   # D — Ед. цена
        col_price_vat   = df.columns[14]  # O — Цената вкл. ДДС

        for _, row in df.iterrows():
            info = ProductInfo(
                internal_code    = str(row[col_code]).strip(),
                supplier_article = str(row[col_supplier]).strip(),
                description      = str(row[col_description]).strip(),
                unit_price       = str(row[col_price]).strip(),
                price_with_vat   = str(row[col_price_vat]).strip(),
            )
            if info.internal_code:
                self._by_internal_code[info.internal_code.upper()] = info
            if info.supplier_article:
                self._by_supplier_article[info.supplier_article.upper()] = info

    @staticmethod
    def _clean_val(val: str) -> str:
        """Normalize cell value: handle scientific notation and trailing .0 from numeric strings."""
        v = val.strip()
        # Scientific notation: "4.062172416160e+12" → "4062172416160"
        if re.search(r'[eE][+\-]?\d+', v):
            try:
                v = str(int(float(v)))
            except (ValueError, OverflowError):
                pass
        # Trailing .0: "4062172416160.0" → "4062172416160"
        elif v.endswith(".0") and v[:-2].lstrip("-").isdigit():
            v = v[:-2]
        return v

    @staticmethod
    def _find_col(headers: list[str], keywords: list[str]) -> Optional[int]:
        """Find column index whose header contains any of the keywords."""
        for kw in keywords:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return None

    def _load_ean(self, path: Path):
        df = pd.read_excel(path, engine="openpyxl", header=0, dtype=str)
        df = df.fillna("")

        headers = [str(c).lower().strip() for c in df.columns]
        logger.info("EAN file columns: %s", headers)

        # Product code column: first column with keyword, else column A
        art_idx = self._find_col(headers, ["артикул", "арт.", "код", "article", "code", "item"]) or 0
        col_art = df.columns[art_idx]

        desc_idx = self._find_col(headers, ["описание", "description", "desc", "naziv"])
        col_desc = df.columns[desc_idx] if desc_idx is not None else None

        # EAN column: detect by header keyword, then by scanning values for
        # barcode-like content, then fall back to column C (index 2)
        ean_idx = self._find_col(headers, ["ean", "gtin", "баркод", "barcode"])
        if ean_idx is None:
            # Scan each column: pick the one with the most 8-14 digit values
            best_col, best_count = None, 0
            for ci, col in enumerate(df.columns):
                count = df[col].str.match(r'^\d{8,14}$', na=False).sum()
                if count > best_count:
                    best_count, best_col = count, ci
            ean_idx = best_col if best_col is not None else min(2, len(df.columns) - 1)

        col_ean = df.columns[ean_idx]
        logger.info("EAN file mapping — art col:%s  ean col:%s  (total cols: %d)",
                    col_art, col_ean, len(df.columns))

        for _, row in df.iterrows():
            art_code   = self._clean_val(str(row[col_art]))
            ean_number = self._clean_val(str(row[col_ean]))
            desc_text  = str(row[col_desc]).strip() if col_desc is not None else ""

            if not ean_number or not re.match(r'^\d{8,14}$', ean_number):
                continue

            info = self._by_internal_code.get(art_code.upper())
            if info and not info.ean:
                info.ean = ean_number

            self._by_ean[ean_number] = info or ProductInfo(
                internal_code=art_code,
                description=desc_text,
                ean=ean_number,
            )

    def lookup(self, code: str) -> Optional[ProductInfo]:
        if not code or not self._loaded:
            return None
        key = code.strip().upper()
        ean_key = self._clean_val(code.strip())
        return (
            self._by_supplier_article.get(key) or
            self._by_internal_code.get(key) or
            self._by_ean.get(ean_key)
        )

    @property
    def is_loaded(self) -> bool:
        return self._loaded and (bool(self._by_internal_code) or bool(self._by_ean))


# Singleton
_db: Optional[ProductDatabase] = None


def get_product_db(data_dir: str) -> ProductDatabase:
    global _db
    if _db is None:
        _db = ProductDatabase(data_dir)
        _db.load()
    return _db


# ---------------------------------------------------------------------------
# Rezaw-Plast product database
# ---------------------------------------------------------------------------

class RezawPlastDatabase:
    """
    File 1 — rezaw-plast-all-export-products.xlsx
        Код | Описание | Ед_ цена | Продуктова група | ...

    File 2 — rezaw-plast-all-export-ean.xlsx
        Артикул (код) | Артикул (Описание) | EAN номер
    """

    def __init__(self, data_dir: str):
        self.data_dir = Path(data_dir)
        self._by_code: dict[str, ProductInfo] = {}
        self._by_ean:  dict[str, ProductInfo] = {}
        self._loaded = False

    def load(self):
        all_xlsx = sorted(f.name for f in self.data_dir.glob("*.xlsx"))
        logger.info("Rezaw-Plast: data_dir=%s  xlsx files=%s", self.data_dir, all_xlsx)

        def _find_file(exact: str, pattern: str) -> Optional[Path]:
            p = self.data_dir / exact
            if p.exists():
                logger.info("Rezaw-Plast: found exact file '%s'", exact)
                return p
            matches = sorted(self.data_dir.glob(pattern))
            if matches:
                logger.info("Rezaw-Plast: '%s' not found, using '%s'", exact, matches[0].name)
                return matches[0]
            logger.info("Rezaw-Plast: no file matching '%s' in %s", pattern, self.data_dir)
            return None

        products_file = _find_file(
            "rezawplast-products.xlsx",
            "rezaw*plast*product*.xlsx",
        )
        ean_file = _find_file(
            "rezaw-plast-all-export-ean.xlsx",
            "rezaw*plast*ean*.xlsx",
        )

        if products_file:
            try:
                self._load_products(products_file)
                logger.info("Rezaw-Plast products loaded: %d records", len(self._by_code))
            except Exception as e:
                logger.error("Failed to load Rezaw-Plast products: %s", e)

        if ean_file:
            try:
                self._load_ean(ean_file)
                logger.info("Rezaw-Plast EAN loaded: %d entries", len(self._by_ean))
            except Exception as e:
                logger.error("Failed to load Rezaw-Plast EAN: %s", e)

        self._loaded = True
        logger.info("Rezaw-Plast DB ready: %d by code, %d by EAN",
                    len(self._by_code), len(self._by_ean))

    def _load_products(self, path: Path):
        df = pd.read_excel(path, engine="openpyxl", header=0, dtype=str)
        df = df.fillna("")
        headers = [str(c).lower().strip() for c in df.columns]

        code_idx  = ProductDatabase._find_col(headers, ["код", "code", "артикул"]) or 0
        desc_idx  = ProductDatabase._find_col(headers, ["описание", "description", "naziv"]) or 1
        price_idx = ProductDatabase._find_col(headers, ["ед_ цена", "ед.цена", "цена", "price"])
        if price_idx is None and len(df.columns) > 2:
            price_idx = 2

        col_code  = df.columns[code_idx]
        col_desc  = df.columns[desc_idx]
        col_price = df.columns[price_idx] if price_idx is not None else None

        logger.info("Rezaw-Plast products — code:%s  desc:%s  price:%s",
                    col_code, col_desc, col_price)

        sample_logged = False
        for _, row in df.iterrows():
            code = str(row[col_code]).strip()
            if not code or code.lower() in ("nan", ""):
                continue
            # Normalize numeric codes read as floats (e.g. "100112.0" → "100112")
            if re.match(r'^\d+\.0$', code):
                code = code[:-2]
            if not sample_logged:
                logger.info("Rezaw-Plast products sample code (raw→norm): %r", code)
                sample_logged = True
            info = ProductInfo(
                internal_code = code,
                description   = str(row[col_desc]).strip(),
                unit_price    = str(row[col_price]).strip() if col_price is not None else "",
            )
            self._by_code[code.upper()] = info

    def _load_ean(self, path: Path):
        df = pd.read_excel(path, engine="openpyxl", header=0, dtype=str)
        df = df.fillna("")
        headers = [str(c).lower().strip() for c in df.columns]

        code_idx = ProductDatabase._find_col(headers, ["артикул", "код", "code"]) or 0
        ean_idx  = ProductDatabase._find_col(headers, ["ean", "баркод", "barcode"])
        if ean_idx is None:
            best_col, best_count = None, 0
            for ci, col in enumerate(df.columns):
                count = df[col].str.match(r'^\d{8,14}$', na=False).sum()
                if count > best_count:
                    best_count, best_col = count, ci
            ean_idx = best_col if best_col is not None else min(2, len(df.columns) - 1)

        col_code = df.columns[code_idx]
        col_ean  = df.columns[ean_idx]
        logger.info("Rezaw-Plast EAN — code col:%s  ean col:%s", col_code, col_ean)

        for _, row in df.iterrows():
            code = ProductDatabase._clean_val(str(row[col_code]).strip())
            ean  = ProductDatabase._clean_val(str(row[col_ean]).strip())

            if not ean or not re.match(r'^\d{8,14}$', ean):
                continue

            info = self._by_code.get(code.upper())
            if info and not info.ean:
                info.ean = ean

            self._by_ean[ean] = info or ProductInfo(internal_code=code, ean=ean)

    def lookup(self, code: str) -> Optional[ProductInfo]:
        if not code or not self._loaded:
            return None
        key = code.strip().upper()
        # Normalize numeric codes that may arrive as "100112.0"
        if re.match(r'^\d+\.0$', key):
            key = key[:-2]
        ean_key = ProductDatabase._clean_val(code.strip())
        return self._by_code.get(key) or self._by_ean.get(ean_key)

    @property
    def is_loaded(self) -> bool:
        return self._loaded and (bool(self._by_code) or bool(self._by_ean))


# Singleton
_rp_db: Optional[RezawPlastDatabase] = None


def get_rezaw_plast_db(data_dir: str) -> RezawPlastDatabase:
    global _rp_db
    if _rp_db is None:
        _rp_db = RezawPlastDatabase(data_dir)
        _rp_db.load()
    return _rp_db


# ---------------------------------------------------------------------------
# Generic supplier name database (Code → Description)
# Used by Maxton, Avisa, Amio and any future supplier.
# File format: Excel with two columns — Code and Description.
# Column headers are auto-detected by keyword (BG/EN).
# ---------------------------------------------------------------------------

class SupplierNameDatabase:
    def __init__(self, data_dir: str, filename: str,
                 desc_keywords: list | None = None,
                 strip_prefix: str = ""):
        self.data_dir = Path(data_dir)
        self.filename = filename
        self._strip_prefix = strip_prefix.upper()
        # Caller can override to prefer a specific column (e.g. "eshop" for Rigum)
        self._desc_keywords = desc_keywords or [
            "описание", "description", "name", "naziv", "наименование"
        ]
        self._by_code: dict[str, ProductInfo] = {}
        self._loaded = False

    def load(self):
        path = self.data_dir / self.filename
        if not path.exists():
            logger.info("Supplier DB not found (optional): %s", path)
            self._loaded = True
            return
        try:
            df = pd.read_excel(path, engine="openpyxl", header=0, dtype=str)
            df = df.fillna("")
            headers = [str(c).lower().strip() for c in df.columns]

            code_idx = ProductDatabase._find_col(
                headers, ["код", "code", "артикул", "article", "ref", "item"]
            ) or 0
            desc_idx = ProductDatabase._find_col(
                headers, self._desc_keywords
            ) or 1

            col_code = df.columns[code_idx]
            col_desc = df.columns[desc_idx]
            logger.info("%s — code:%s  desc:%s", self.filename, col_code, col_desc)

            for _, row in df.iterrows():
                code = str(row[col_code]).strip()
                desc = str(row[col_desc]).strip()
                if not code or code.lower() in ("nan", ""):
                    continue
                self._by_code[code.upper()] = ProductInfo(
                    internal_code=code,
                    description=desc or None,
                )
            logger.info("%s loaded: %d records", self.filename, len(self._by_code))
        except Exception as e:
            logger.error("Failed to load %s: %s", self.filename, e)
        self._loaded = True

    def lookup(self, code: str) -> Optional[ProductInfo]:
        if not code or not self._loaded:
            return None
        key = code.strip().upper()
        if self._strip_prefix and key.startswith(self._strip_prefix):
            key = key[len(self._strip_prefix):]
        result = self._by_code.get(key)
        if result is None and re.match(r'^[A-Z]{2,6}-', key):
            # Fallback: catalog entry may lack the leading alpha prefix (e.g. 'AMIO-SED31269' → 'SED31269')
            fallback_key = key.split('-', 1)[1]
            result = self._by_code.get(fallback_key)
        return result

    def lookup_by_desc_words(self, words: list[str],
                              min_overlap: int = 3,
                              min_ratio: float = 0.40) -> Optional[tuple[str, "ProductInfo"]]:
        """Return (code, info) for the DB entry whose description best matches
        the given word list.  Uses word-overlap scoring; requires the best match
        to be at least 30 % better than the second-best to avoid ambiguity.
        """
        if not words or not self._loaded:
            return None
        query = {w.lower() for w in words if len(w) > 2}
        if not query:
            return None
        scores: list[tuple[float, str]] = []
        for code, info in self._by_code.items():
            if not info.description:
                continue
            db_words = {w.lower() for w in info.description.split() if len(w) > 2}
            if not db_words:
                continue
            overlap = len(query & db_words)
            if overlap < min_overlap:
                continue
            score = overlap / max(len(query), len(db_words))
            if score >= min_ratio:
                scores.append((score, code))
        if not scores:
            return None
        scores.sort(reverse=True)
        best_score, best_code = scores[0]
        # Require uniqueness: best must be ≥30 % better than second
        if len(scores) >= 2 and scores[1][0] >= best_score * 0.70:
            return None  # ambiguous — skip
        return best_code, self._by_code[best_code]

    @property
    def is_loaded(self) -> bool:
        return self._loaded and bool(self._by_code)


# Singletons — one per supplier
_mx_db:           Optional[SupplierNameDatabase] = None
_avisa_db:        Optional[SupplierNameDatabase] = None
_amio_db:         Optional[SupplierNameDatabase] = None
_mtech_db:        Optional[SupplierNameDatabase] = None
_mafra_db:        Optional[SupplierNameDatabase] = None
_amal_plast_db:   Optional[SupplierNameDatabase] = None
_car_passion_db:  Optional[SupplierNameDatabase] = None
_vinove_db:       Optional[SupplierNameDatabase] = None
_gumarny_zubri_db: Optional[SupplierNameDatabase] = None
_rigum_db:         Optional[SupplierNameDatabase] = None
_petex_db:         Optional[SupplierNameDatabase] = None
_gh_db:            Optional[SupplierNameDatabase] = None
_frogum_db:        Optional[SupplierNameDatabase] = None
_kegel_db:         Optional[SupplierNameDatabase] = None
_automania_db:     Optional[SupplierNameDatabase] = None


def get_maxton_db(data_dir: str) -> SupplierNameDatabase:
    global _mx_db
    if _mx_db is None:
        _mx_db = SupplierNameDatabase(data_dir, "maxton-products.xlsx")
        _mx_db.load()
    return _mx_db


def get_avisa_db(data_dir: str) -> SupplierNameDatabase:
    global _avisa_db
    if _avisa_db is None:
        _avisa_db = SupplierNameDatabase(data_dir, "avisa-products.xlsx")
        _avisa_db.load()
    return _avisa_db


def get_amio_db(data_dir: str) -> SupplierNameDatabase:
    global _amio_db
    if _amio_db is None:
        _amio_db = SupplierNameDatabase(data_dir, "amio-products.xlsx")
        _amio_db.load()
    return _amio_db


def get_mtech_db(data_dir: str) -> SupplierNameDatabase:
    global _mtech_db
    if _mtech_db is None:
        _mtech_db = SupplierNameDatabase(data_dir, "mtech-products.xlsx")
        _mtech_db.load()
    return _mtech_db


def get_mafra_db(data_dir: str) -> SupplierNameDatabase:
    global _mafra_db
    if _mafra_db is None:
        _mafra_db = SupplierNameDatabase(data_dir, "mafra-products.xlsx")
        _mafra_db.load()
    return _mafra_db


def get_amal_plast_db(data_dir: str) -> SupplierNameDatabase:
    global _amal_plast_db
    if _amal_plast_db is None:
        _amal_plast_db = SupplierNameDatabase(data_dir, "amal-plast.xlsx")
        _amal_plast_db.load()
    return _amal_plast_db


def get_car_passion_db(data_dir: str) -> SupplierNameDatabase:
    global _car_passion_db
    if _car_passion_db is None:
        _car_passion_db = SupplierNameDatabase(data_dir, "car-passion.xlsx")
        _car_passion_db.load()
    return _car_passion_db


def get_vinove_db(data_dir: str) -> SupplierNameDatabase:
    global _vinove_db
    if _vinove_db is None:
        _vinove_db = SupplierNameDatabase(data_dir, "vinove-products.xlsx")
        _vinove_db.load()
    return _vinove_db


def get_gumarny_zubri_db(data_dir: str) -> SupplierNameDatabase:
    global _gumarny_zubri_db
    if _gumarny_zubri_db is None:
        _gumarny_zubri_db = SupplierNameDatabase(data_dir, "gumarny-zubri-products.xlsx")
        _gumarny_zubri_db.load()
    return _gumarny_zubri_db


def get_frogum_db(data_dir: str) -> SupplierNameDatabase:
    global _frogum_db
    if _frogum_db is None:
        _frogum_db = SupplierNameDatabase(data_dir, "frogum-products.xlsx")
        _frogum_db.load()
    return _frogum_db


_gelly_plast_db: SupplierNameDatabase | None = None


def get_gelly_plast_db(data_dir: str) -> SupplierNameDatabase:
    global _gelly_plast_db
    if _gelly_plast_db is None:
        _gelly_plast_db = SupplierNameDatabase(data_dir, "gelly-plast-products.xlsx")
        _gelly_plast_db.load()
    return _gelly_plast_db


_farad_db: SupplierNameDatabase | None = None


def get_farad_db(data_dir: str) -> SupplierNameDatabase:
    global _farad_db
    if _farad_db is None:
        _farad_db = SupplierNameDatabase(data_dir, "farad-products.xlsx")
        _farad_db.load()
    return _farad_db


def get_geyer_hosaja_db(data_dir: str) -> SupplierNameDatabase:
    global _gh_db
    if _gh_db is None:
        _gh_db = SupplierNameDatabase(data_dir, "geyer-hosaja-products.xlsx")
        _gh_db.load()
    return _gh_db


def get_petex_db(data_dir: str) -> SupplierNameDatabase:
    global _petex_db
    if _petex_db is None:
        _petex_db = SupplierNameDatabase(data_dir, "petex-products.xlsx")
        _petex_db.load()
    return _petex_db


def get_rigum_db(data_dir: str) -> SupplierNameDatabase:
    global _rigum_db
    if _rigum_db is None:
        # Prefer "Описание eShop" (detailed) over a short "Описание" column
        _rigum_db = SupplierNameDatabase(
            data_dir, "rigum-products.xlsx",
            desc_keywords=["eshop", "описание", "description", "name"],
        )
        _rigum_db.load()
    return _rigum_db


def get_kegel_blazusiak_db(data_dir: str) -> SupplierNameDatabase:
    global _kegel_db
    if _kegel_db is None:
        _kegel_db = SupplierNameDatabase(data_dir, "kegel-products.xlsx")
        _kegel_db.load()
    return _kegel_db


def get_automania_db(data_dir: str) -> SupplierNameDatabase:
    global _automania_db
    if _automania_db is None:
        _automania_db = SupplierNameDatabase(data_dir, "automania-products.xlsx",
                                             strip_prefix="AVM-")
        _automania_db.load()
    return _automania_db


_hakr_db: Optional[SupplierNameDatabase] = None


def get_hakr_db(data_dir: str) -> SupplierNameDatabase:
    global _hakr_db
    if _hakr_db is None:
        _hakr_db = SupplierNameDatabase(data_dir, "hark-products.xlsx")
        _hakr_db.load()
    return _hakr_db


_tompar_db: Optional[SupplierNameDatabase] = None


def get_tompar_db(data_dir: str) -> SupplierNameDatabase:
    global _tompar_db
    if _tompar_db is None:
        _tompar_db = SupplierNameDatabase(data_dir, "tompar-products.xlsx")
        _tompar_db.load()
    return _tompar_db


_senax_db: Optional[SupplierNameDatabase] = None


def get_senax_db(data_dir: str) -> SupplierNameDatabase:
    global _senax_db
    if _senax_db is None:
        _senax_db = SupplierNameDatabase(data_dir, "sonax-products.xlsx")
        _senax_db.load()
    return _senax_db


_heko_db: Optional[SupplierNameDatabase] = None


def get_heko_db(data_dir: str) -> SupplierNameDatabase:
    global _heko_db
    if _heko_db is None:
        _heko_db = SupplierNameDatabase(data_dir, "heko-products.xlsx")
        _heko_db.load()
    return _heko_db


# ---------------------------------------------------------------------------
# Supplier code mapping  (supplier invoice code → our internal code)
# File format: Excel with two columns:
#   Column A — "Наш код"       (internal code, e.g. "1-HA1/E")
#   Column B — "Код доставчик" (code as it appears on the invoice)
# Column headers are auto-detected by keyword.
# ---------------------------------------------------------------------------

class SupplierCodeMapping:
    """Translates a supplier's invoice code to our internal product code."""

    def __init__(self, data_dir: str, filename: str):
        self.data_dir = Path(data_dir)
        self.filename = filename
        self._map: dict[str, str] = {}   # supplier_code.upper() → internal_code
        self._loaded = False

    def load(self):
        path = self.data_dir / self.filename
        if not path.exists():
            logger.info("Code mapping not found (optional): %s", path)
            self._loaded = True
            return
        try:
            df = pd.read_excel(path, engine="openpyxl", header=0, dtype=str)
            df = df.fillna("")
            headers = [str(c).lower().strip() for c in df.columns]

            our_idx = ProductDatabase._find_col(
                headers, ["наш", "internal", "наш код", "our", "код", "code"]
            ) or 0
            sup_idx = ProductDatabase._find_col(
                headers, ["доставчик", "supplier", "farad", "производител",
                          "артикул", "article", "invoice"]
            )
            if sup_idx is None:
                sup_idx = 1 if our_idx == 0 else 0

            col_our = df.columns[our_idx]
            col_sup = df.columns[sup_idx]
            logger.info("%s — our_code col:%s  supplier_code col:%s",
                        self.filename, col_our, col_sup)

            for _, row in df.iterrows():
                our_code = str(row[col_our]).strip()
                sup_code = ProductDatabase._clean_val(str(row[col_sup]).strip())
                if not our_code or not sup_code:
                    continue
                if our_code.lower() in ("nan", "") or sup_code.lower() in ("nan", ""):
                    continue
                # Strip trailing .0 from numeric codes read as floats
                if re.match(r'^\d+\.0$', sup_code):
                    sup_code = sup_code[:-2]
                key = re.sub(r'\s+', ' ', sup_code).upper()
                self._map[key] = our_code
                # Also index without "1-" prefix so both forms match
                if key.startswith("1-"):
                    self._map[key[2:]] = our_code

            logger.info("%s loaded: %d code mappings", self.filename, len(self._map))
        except Exception as e:
            logger.error("Failed to load code mapping %s: %s", self.filename, e)
        self._loaded = True

    def translate(self, supplier_code: str) -> Optional[str]:
        """Return internal code for the given supplier code, or None if not found.

        Tries in order:
        1. Full string                 ("1-Z4/E SC.NERA 1CH.")
        2. Full string trailing-stripped ("1-Z4/E SC.NERA 1CH")
        3. Without "1-" prefix         ("Z4/E SC.NERA 1CH.")
        4. Without "1-", stripped      ("Z4/E SC.NERA 1CH")
        5. Short code only             ("Z4/E") — first word
        6. Short code without "1-"     ("Z4/E" from "1-Z4/E")
        """
        if not supplier_code or not self._loaded:
            return None
        key = re.sub(r'\s+', ' ', supplier_code.strip()).upper()
        key_stripped = key.rstrip(".,")

        candidates = [key]
        if key_stripped != key:
            candidates.append(key_stripped)
        if key.startswith("1-"):
            candidates.append(key[2:])
            if key_stripped.startswith("1-"):
                candidates.append(key_stripped[2:])
        short = key.split()[0]
        if short != key:
            candidates.append(short)
            if short.startswith("1-"):
                candidates.append(short[2:])

        for candidate in candidates:
            result = self._map.get(candidate)
            if result is not None:
                return result
        return None

    @property
    def is_loaded(self) -> bool:
        return self._loaded and bool(self._map)


_bmw_db: Optional[SupplierNameDatabase] = None


def get_bmw_db(data_dir: str) -> SupplierNameDatabase:
    global _bmw_db
    if _bmw_db is None:
        _bmw_db = SupplierNameDatabase(data_dir, "bmw-products.xlsx")
        _bmw_db.load()
    return _bmw_db


_farad_code_map: Optional[SupplierCodeMapping] = None


def get_farad_code_map(data_dir: str) -> SupplierCodeMapping:
    global _farad_code_map
    if _farad_code_map is None:
        _farad_code_map = SupplierCodeMapping(data_dir, "farad-code-map.xlsx")
        _farad_code_map.load()
    return _farad_code_map


_gz_code_map: Optional[SupplierCodeMapping] = None


def get_gumarny_zubri_code_map(data_dir: str) -> SupplierCodeMapping:
    global _gz_code_map
    if _gz_code_map is None:
        _gz_code_map = SupplierCodeMapping(data_dir, "gumarny-zubri-code-map.xlsx")
        _gz_code_map.load()
    return _gz_code_map


# ---------------------------------------------------------------------------
# Wunder-Baum product catalog
# File: wunder-baum-products.xlsx
# Columns auto-detected: internal code, description, optionally EAN barcode.
# Lookup by: WB-XXXXXXXX code, raw article number, or EAN barcode.
# ---------------------------------------------------------------------------

class WunderBaumDatabase:
    def __init__(self, data_dir: str):
        self.data_dir = Path(data_dir)
        self._by_code: dict[str, ProductInfo] = {}  # code.upper() → info
        self._by_ean: dict[str, ProductInfo] = {}   # ean → info
        self._loaded = False

    def load(self):
        path = self.data_dir / "wunder-baum-products.xlsx"
        if not path.exists():
            logger.info("WunderBaum DB not found (optional): %s", path)
            self._loaded = True
            return
        try:
            df = pd.read_excel(path, engine="openpyxl", header=0, dtype=str)
            df = df.fillna("")
            headers = [str(c).lower().strip() for c in df.columns]

            code_idx = ProductDatabase._find_col(
                headers, ["код", "code", "артикул", "article", "ref", "item"]
            ) or 0
            desc_idx = ProductDatabase._find_col(
                headers, ["описание", "description", "name", "стока", "наименование"]
            ) or 1
            ean_idx = ProductDatabase._find_col(
                headers, ["ean", "баркод", "barcode", "gtin", "upc"]
            )

            col_code = df.columns[code_idx]
            col_desc = df.columns[desc_idx]
            col_ean  = df.columns[ean_idx] if ean_idx is not None else None
            logger.info("WunderBaum — code:%s  desc:%s  ean:%s", col_code, col_desc, col_ean)

            for _, row in df.iterrows():
                code = ProductDatabase._clean_val(str(row[col_code]).strip())
                desc = str(row[col_desc]).strip()
                ean  = ProductDatabase._clean_val(str(row[col_ean]).strip()) if col_ean is not None else ""

                if not code or code.lower() in ("nan", ""):
                    continue
                if re.match(r'^\d+\.0$', code):
                    code = code[:-2]
                # Strip any invisible characters from numeric codes (e.g. Text-formatted
                # cells in Excel may carry hidden markers); mirrors how the invoice parser
                # normalises codes with re.sub(r'[^0-9]', '', ean_raw).
                digits_only = re.sub(r'[^0-9]', '', code)
                if re.match(r'^\d{4,14}$', digits_only) and digits_only != code:
                    code = digits_only

                info = ProductInfo(
                    internal_code=code,
                    description=desc or None,
                    ean=ean or None,
                )
                self._by_code[code.upper()] = info

                if ean and re.match(r'^\d{8,14}$', ean):
                    self._by_ean[ean] = info

            logger.info("WunderBaum DB loaded: %d code entries, %d EAN entries",
                        len(self._by_code), len(self._by_ean))
        except Exception as e:
            logger.error("Failed to load WunderBaum DB: %s", e)
        self._loaded = True

    def lookup(self, code: str) -> Optional[ProductInfo]:
        if not code or not self._loaded:
            return None
        key = code.strip().upper()
        info = self._by_code.get(key)
        if info:
            return info
        # EAN lookup (13-digit numeric code from invoice Код column)
        ean_clean = re.sub(r'[^0-9]', '', code)
        if re.match(r'^\d{8,14}$', ean_clean):
            return self._by_ean.get(ean_clean)
        return None

    @property
    def is_loaded(self) -> bool:
        return self._loaded and (bool(self._by_code) or bool(self._by_ean))


_wb_db: Optional[WunderBaumDatabase] = None


def get_wunder_baum_db(data_dir: str) -> WunderBaumDatabase:
    global _wb_db
    if _wb_db is None:
        _wb_db = WunderBaumDatabase(data_dir)
        _wb_db.load()
    return _wb_db


_wb_code_map: Optional[SupplierCodeMapping] = None

_bardahl_db: Optional[SupplierNameDatabase] = None


def get_bardahl_db(data_dir: str) -> SupplierNameDatabase:
    global _bardahl_db
    if _bardahl_db is None:
        _bardahl_db = SupplierNameDatabase(data_dir, "bardahl-products.xlsx")
        _bardahl_db.load()
    return _bardahl_db


def get_wunder_baum_code_map(data_dir: str) -> SupplierCodeMapping:
    global _wb_code_map
    if _wb_code_map is None:
        _wb_code_map = SupplierCodeMapping(data_dir, "wunder-baum-code-map.xlsx")
        _wb_code_map.load()
    return _wb_code_map


_areon_db: Optional[SupplierNameDatabase] = None


def get_areon_db(data_dir: str) -> SupplierNameDatabase:
    global _areon_db
    if _areon_db is None:
        _areon_db = SupplierNameDatabase(data_dir, "areon-products.xlsx")
        _areon_db.load()
    return _areon_db


_areon_code_map: Optional[SupplierCodeMapping] = None


def get_areon_code_map(data_dir: str) -> SupplierCodeMapping:
    global _areon_code_map
    if _areon_code_map is None:
        _areon_code_map = SupplierCodeMapping(data_dir, "areon-code-map.xlsx")
        _areon_code_map.load()
    return _areon_code_map


_slime_db: Optional[SupplierNameDatabase] = None


def get_slime_db(data_dir: str) -> SupplierNameDatabase:
    global _slime_db
    if _slime_db is None:
        _slime_db = SupplierNameDatabase(data_dir, "slime-products.xlsx")
        _slime_db.load()
    return _slime_db


_xado_db: Optional[SupplierNameDatabase] = None


def get_xado_db(data_dir: str) -> SupplierNameDatabase:
    global _xado_db
    if _xado_db is None:
        _xado_db = SupplierNameDatabase(data_dir, "xado-products.xlsx")
        _xado_db.load()
    return _xado_db
