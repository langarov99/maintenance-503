import re
import json
import logging
from dataclasses import dataclass, field, asdict
from typing import Optional

logger = logging.getLogger(__name__)

# Minimum number of filled fields to consider regex extraction successful
REGEX_CONFIDENCE_THRESHOLD = 3


@dataclass
class ProductRecord:
    product_code: Optional[str] = None       # Колона 1
    quantity: Optional[str] = None           # Колона 2
    price: Optional[str] = None             # Колона 3
    product_name: Optional[str] = None      # Колона 4
    ean: Optional[str] = None               # Колона 5
    weight_kg: Optional[str] = None         # Колона 6
    parts_in_set: Optional[str] = None      # Колона 7
    color: Optional[str] = None             # Колона 8
    extraction_method: str = "regex"

    def filled_count(self) -> int:
        fields = [self.product_code, self.quantity, self.price,
                  self.product_name, self.ean, self.weight_kg,
                  self.parts_in_set, self.color]
        return sum(1 for f in fields if f and str(f).strip())

    def to_dict(self) -> dict:
        return asdict(self)


# ---------------------------------------------------------------------------
# Regex patterns — multilingual (BG, EN, PL, CS, IT, DE)
# ---------------------------------------------------------------------------

EAN_PATTERN = re.compile(r'\b(\d{8}|\d{12}|\d{13}|\d{14}|\d{18}|\d{20})\b')

PRICE_PATTERN = re.compile(
    r'(?:(?:цена|price|preis|prezzo|cena|prix|brutto|netto|без\s*ддс|с\s*ддс|'
    r'net|gross|inkl|excl)[^\d]{0,15})?'
    r'(\d{1,6}[.,]\d{2})\s*'
    r'(лв\.?|BGN|EUR|€|PLN|zł|CZK|Kč|CHF|GBP|£|\$|USD)?',
    re.IGNORECASE
)

QUANTITY_PATTERN = re.compile(
    r'(?:qty|quantity|количество|menge|anzahl|quantità|cantidad|množství|ilość)'
    r'\s*[:\-=]?\s*(\d+(?:[.,]\d+)?)',
    re.IGNORECASE
)

WEIGHT_PATTERN = re.compile(
    r'(\d+(?:[.,]\d+)?)\s*'
    r'(kg|кг|kilogram|kilogramm|chilogrammo|кило|g\b|gram|грам)',
    re.IGNORECASE
)

PARTS_PATTERN = re.compile(
    r'(?:бр\.|pcs|pieces|stück|pezzi|sztuk|ks|části|комплект|set|parts?)'
    r'\s*[:\-=]?\s*(\d+)',
    re.IGNORECASE
)

COLOR_PATTERN = re.compile(
    r'(?:цвят|color|colour|farbe|colore|kolor|barva)\s*[:\-=]?\s*'
    r'([а-яА-Яa-zA-ZäöüÄÖÜßąćęłńóśźżčďěňřšůýžàèéìíòóùúâêîôûæœ][а-яА-Яa-zA-ZäöüÄÖÜßąćęłńóśźżčďěňřšůýžàèéìíòóùúâêîôûæœ\s]{1,30})',
    re.IGNORECASE
)

PRODUCT_CODE_PATTERN = re.compile(
    r'(?:код|code|art\.?(?:nr\.?|no\.?|#)?|artikel|codice|item\s*(?:no|nr|#)?)'
    r'\s*[:\-=]?\s*([A-Za-z0-9\-_/\.]{3,20})',
    re.IGNORECASE
)

PRODUCT_NAME_PATTERN = re.compile(
    r'(?:наименование|продукт|product|artikel|name|bezeichnung|nome|produkt|nazwa)'
    r'\s*[:\-=]?\s*(.{3,80})',
    re.IGNORECASE
)


# ---------------------------------------------------------------------------
# Table-based extraction (for structured Excel / PDF tables)
# ---------------------------------------------------------------------------

HEADER_ALIASES = {
    "product_code": ["код", "code", "art", "artikel", "codice", "item", "артикул", "арт", "nr", "no", "référence"],
    "quantity":     ["кол", "qty", "quantity", "menge", "anzahl", "quantità", "ilość", "množství", "доставено", "delivered", "geliefert", "consegnato", "поръчано", "ordered", "бр"],
    "price":        ["цена", "price", "preis", "prezzo", "cena", "prix", "единична цена", "unit price"],
    "product_name": ["наименование", "описание", "продукт", "name", "bezeichnung", "nome", "nazwa", "název", "description", "omschrijving", "клиентско", "artikel"],
    "ean":          ["ean", "баркод", "barcode", "gtin", "upc", "ean код"],
    "weight_kg":    ["кг", "kg", "weight", "gewicht", "peso", "waga", "hmotnost", "брутo", "нето", "brutto", "netto", "gross", "net", "тегло"],
    "parts_in_set": ["единични", "пълни", "pcs", "pieces", "stück", "set", "комплект", "sztuk", "ks", "бр"],
    "color":        ["цвят", "color", "colour", "farbe", "colore", "kolor", "barva"],
}


def _match_header(cell: str) -> Optional[str]:
    cell_lower = cell.lower().strip()
    for field_name, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            if alias in cell_lower:
                return field_name
    return None


def _clean_cell(value) -> str:
    if value is None:
        return ""
    # Take only first non-empty line from multi-line cells
    lines = [l.strip() for l in str(value).split("\n") if l.strip()]
    return lines[0] if lines else ""


def _extract_barcode_from_cell(value) -> str:
    """Extract long barcode from second line of multi-line cell (e.g. OSRAM article cells)."""
    if value is None:
        return ""
    lines = [l.strip() for l in str(value).split("\n") if l.strip()]
    for line in lines[1:]:
        m = re.search(r'\b(\d{8}|\d{12}|\d{13}|\d{14}|\d{18}|\d{20})\b', line)
        if m:
            return m.group(1)
    return ""


def extract_from_table(table: list[list]) -> list[ProductRecord]:
    if not table or len(table) < 2:
        return []

    # Scan up to first 8 rows for header — accumulate best mapping
    best_mapping = {}
    best_row = None

    for row_idx, row in enumerate(table[:8]):
        mapping = {}
        for col_idx, cell in enumerate(row):
            if not cell or not str(cell).strip():
                continue
            # Check each line in multi-line cell
            for line in str(cell).split("\n"):
                field_name = _match_header(line)
                if field_name and field_name not in mapping:
                    mapping[field_name] = col_idx
        if len(mapping) > len(best_mapping):
            best_mapping = mapping
            best_row = row_idx

    if not best_mapping or best_row is None:
        return []

    records = []
    for row in table[best_row + 1:]:
        if not any(str(c).strip() for c in row):
            continue
        rec = ProductRecord(extraction_method="table")
        for field_name, col_idx in best_mapping.items():
            if col_idx < len(row):
                raw = row[col_idx]
                value = _clean_cell(raw)
                if value and value.lower() not in ("none", "nan", ""):
                    setattr(rec, field_name, value)
                # Try to pick barcode from second line of product_code cell
                if field_name == "product_code" and not rec.ean:
                    barcode = _extract_barcode_from_cell(raw)
                    if barcode:
                        rec.ean = barcode
        # Skip rows that look like sub-headers or empty
        if rec.filled_count() >= 1 and rec.product_code and not re.match(r'^[\d\s]+$', rec.product_code or ""):
            records.append(rec)

    return records


# ---------------------------------------------------------------------------
# Regex-based extraction (for raw text)
# ---------------------------------------------------------------------------

def extract_via_regex(text: str) -> ProductRecord:
    rec = ProductRecord(extraction_method="regex")

    m = EAN_PATTERN.search(text)
    if m:
        rec.ean = m.group(1)

    m = PRICE_PATTERN.search(text)
    if m:
        rec.price = m.group(1) + (" " + m.group(2) if m.group(2) else "")

    m = QUANTITY_PATTERN.search(text)
    if m:
        rec.quantity = m.group(1)

    m = WEIGHT_PATTERN.search(text)
    if m:
        rec.weight_kg = m.group(1) + " " + m.group(2)

    m = PARTS_PATTERN.search(text)
    if m:
        rec.parts_in_set = m.group(1)

    m = COLOR_PATTERN.search(text)
    if m:
        rec.color = m.group(1).strip()

    m = PRODUCT_CODE_PATTERN.search(text)
    if m:
        rec.product_code = m.group(1).strip()

    m = PRODUCT_NAME_PATTERN.search(text)
    if m:
        rec.product_name = m.group(1).strip()[:80]

    return rec


# ---------------------------------------------------------------------------
# LLM-based extraction (fallback via llama.cpp)
# ---------------------------------------------------------------------------

LLM_PROMPT_TEMPLATE = """You are a data extraction assistant. Extract product information from the following text and return ONLY valid JSON with these keys (use null if not found):
- product_code
- quantity
- price
- product_name
- ean
- weight_kg
- parts_in_set
- color

Text:
{text}

JSON:"""


def extract_via_llm(text: str, llm) -> ProductRecord:
    prompt = LLM_PROMPT_TEMPLATE.format(text=text[:3000])
    try:
        output = llm(prompt, max_tokens=512, stop=["```", "\n\n\n"], temperature=0.1)
        raw = output["choices"][0]["text"].strip()

        # Extract JSON block
        json_match = re.search(r'\{.*\}', raw, re.DOTALL)
        if not json_match:
            return ProductRecord(extraction_method="llm_failed")

        data = json.loads(json_match.group())
        rec = ProductRecord(extraction_method="llm")
        for key in ["product_code", "quantity", "price", "product_name",
                    "ean", "weight_kg", "parts_in_set", "color"]:
            val = data.get(key)
            if val and str(val).lower() not in ("null", "none", ""):
                setattr(rec, key, str(val))
        return rec
    except Exception as e:
        logger.warning("LLM extraction failed: %s", e)
        return ProductRecord(extraction_method="llm_failed")


# ---------------------------------------------------------------------------
# OSRAM-specific extractor
# ---------------------------------------------------------------------------

# Supplier article numbers: AM510460055, AM4317600EC, AA577421804
# Requires AM/AA/4M/ST prefix to avoid matching document IDs like BG203318362
OSRAM_ARTICLE_RE = re.compile(r'\b((?:AM|AA|4M|ST)\d{6,10}[A-Z0-9]{0,4})\b')
OSRAM_QTY_RE     = re.compile(r'\b(\d+)\s*(?:PCE|pce|STK|stk)\b')

# Country-of-origin lines: "Китай 1,124/ 1,173/ 0,002" — not product names
_ORIGIN_RE = re.compile(
    r'^(Китай|Германия|Словакия|Тайван|Унгария|Полша|Чехия|Австрия|'
    r'China|Germany|Slovakia|Taiwan|Hungary|Poland|Italy|Czech|Austria)',
    re.IGNORECASE
)


def _is_osram_document(text: str) -> bool:
    return bool(re.search(r'OSRAM\s+GMBH|ams-osram|OSRAM\s+GmbH', text, re.IGNORECASE))


def extract_osram_products(text: str) -> list[ProductRecord]:
    """Extract product rows from OSRAM delivery note using article number anchoring."""
    records = []
    lines = text.splitlines()

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = OSRAM_ARTICLE_RE.search(line)
        if not m:
            i += 1
            continue

        article = m.group(1)
        context_lines = lines[i:i + 8]
        context = "\n".join(context_lines)

        rec = ProductRecord(extraction_method="osram")
        rec.product_code = article

        # Quantity — "X PCE" anywhere in context
        qty_m = OSRAM_QTY_RE.search(context)
        if qty_m:
            rec.quantity = qty_m.group(1) + " PCE"

        # Only process lines that look like real product rows (have a quantity)
        if rec.quantity is None:
            i += 1
            continue

        # EAN — standalone 8/12/13/14-digit number on its own line
        for ctx_line in context_lines[1:]:
            ean_m = re.match(r'^(\d{8}|\d{12}|\d{13}|\d{14})$', ctx_line.strip())
            if ean_m:
                rec.ean = ean_m.group(1)
                break

        # Product name — skip origin/price lines and barcode lines
        for j in range(1, 5):
            if i + j >= len(lines):
                break
            candidate = lines[i + j].strip()
            if not candidate or len(candidate) < 5:
                continue
            if re.match(r'^\d+$', candidate):          # pure digits = barcode
                continue
            if OSRAM_ARTICLE_RE.match(candidate):      # another article = next product
                break
            if _ORIGIN_RE.match(candidate):            # country+price line
                continue
            if re.match(r'^\d+\s*PCE', candidate, re.IGNORECASE):  # qty line
                continue
            # Looks like a real description
            rec.product_name = candidate[:80]
            break

        # Weight — last decimal in context, excluding date-like and 4-digit numbers
        weights = re.findall(r'\b(\d{1,3}[.,]\d{1,3})\b', context)
        weights = [w for w in weights if float(w.replace(",", ".")) < 500]
        if weights:
            rec.weight_kg = weights[-1] + " kg"

        records.append(rec)
        i += 1

    # Deduplicate by product_code, keep first occurrence
    seen: set[str] = set()
    unique = []
    for r in records:
        if r.product_code not in seen:
            seen.add(r.product_code)
            unique.append(r)

    return unique


# ---------------------------------------------------------------------------
# Auto-switch orchestrator
# ---------------------------------------------------------------------------

class FieldMapper:
    def __init__(self, llm=None):
        self.llm = llm  # Optional llama-cpp-python Llama instance

    def map(self, extracted: dict) -> list[ProductRecord]:
        tables = extracted.get("tables", [])
        text = extracted.get("text", "")

        # Step 0 — OSRAM supplier detection (before generic table/regex)
        if text and _is_osram_document(text):
            records = extract_osram_products(text)
            if records:
                logger.info("Extraction method: OSRAM-specific (%d records)", len(records))
                return records

        # Step 1 — try table extraction (highest confidence)
        records = []
        for table in tables:
            records.extend(extract_from_table(table))

        if records:
            logger.info("Extraction method: table (%d records)", len(records))
            return records

        # Step 2 — try regex on raw text
        if text.strip():
            rec = extract_via_regex(text)
            if rec.filled_count() >= REGEX_CONFIDENCE_THRESHOLD:
                logger.info("Extraction method: regex (confidence %d/8)", rec.filled_count())
                return [rec]

            logger.info("Regex confidence too low (%d/8) — switching to LLM", rec.filled_count())

            # Step 3 — fallback to LLM
            if self.llm is not None:
                llm_rec = extract_via_llm(text, self.llm)
                if llm_rec.filled_count() >= 1:
                    logger.info("Extraction method: LLM (%d fields found)", llm_rec.filled_count())
                    return [llm_rec]

            # Return partial regex result if LLM unavailable
            if rec.filled_count() >= 1:
                return [rec]

        return []
