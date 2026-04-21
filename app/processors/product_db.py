import logging
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
        main_file = self.data_dir / "osram-export-all.xlsx"
        ean_file  = self.data_dir / "osram-export-all-ean-code.xlsx"

        if main_file.exists():
            try:
                self._load_main(main_file)
                logger.info("Main product file loaded: %d records", len(self._by_internal_code))
            except Exception as e:
                logger.error("Failed to load main product file: %s", e)
        else:
            logger.info("Main product file not found, EAN-only mode: %s", main_file)

        if ean_file.exists():
            try:
                self._load_ean(ean_file)
                logger.info("EAN file loaded: %d EAN entries", len(self._by_ean))
            except Exception as e:
                logger.error("Failed to load EAN file: %s", e)
        else:
            logger.warning("EAN file not found: %s", ean_file)

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

    def _load_ean(self, path: Path):
        df = pd.read_excel(path, engine="openpyxl", header=0, dtype=str)
        df = df.fillna("")

        # A=0 Артикул(код), E=4 Главен баркод, G=6 EAN номер
        col_art      = df.columns[0]   # A — Артикул (код)
        col_main_bc  = df.columns[4]   # E — Главен баркод
        col_ean      = df.columns[6]   # G — EAN номер

        for _, row in df.iterrows():
            art_code     = str(row[col_art]).strip().upper()
            main_barcode = str(row[col_main_bc]).strip()
            ean_number   = str(row[col_ean]).strip()

            info = self._by_internal_code.get(art_code)
            if info:
                # Prefer first non-empty EAN
                if not info.ean and ean_number:
                    info.ean = ean_number
                if not info.main_barcode and main_barcode:
                    info.main_barcode = main_barcode

            # Index by EAN for reverse lookup
            if ean_number:
                self._by_ean[ean_number] = info or ProductInfo(
                    internal_code=art_code, ean=ean_number, main_barcode=main_barcode
                )

    def lookup(self, code: str) -> Optional[ProductInfo]:
        """Lookup by supplier article number, internal code, or EAN."""
        if not code or not self._loaded:
            return None
        key = code.strip().upper()
        return (
            self._by_supplier_article.get(key) or
            self._by_internal_code.get(key) or
            self._by_ean.get(key)
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
