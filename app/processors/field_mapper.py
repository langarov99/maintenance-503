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
    price: Optional[str] = None             # Колона 3 — единична цена
    total_price: Optional[str] = None       # Колона 4 — обща сума (qty × unit price)
    product_name: Optional[str] = None      # Колона 5
    ean: Optional[str] = None               # Колона 6
    weight_kg: Optional[str] = None         # Колона 7
    parts_in_set: Optional[str] = None      # Колона 8
    color: Optional[str] = None             # Колона 9
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

# OSRAM supplier article (IC код): AM460790055, AM4317600EC, AA577421804
# Used only for product database lookup — NOT shown as product code
OSRAM_ARTICLE_RE = re.compile(r'\b((?:AM|AA|4M|ST)\d{6,10}[A-Z0-9]{0,4})\b')

# Position line anchor: 000020, 001110, 001120 etc. OR 80-002 style (delivery sub-line)
_POS_RE = re.compile(r'^(0{2,5}\d{1,4}|\d{2,3}-\d{3})\b')

# Weight triplet: "1,200/ 1,232/ 0,009"
# Invoice columns: Нето (kg) / Брутo (kg) / Обем (cbm)  — take group 1 and 2 (kg only)
_WEIGHT_TRIPLET_RE = re.compile(
    r'(\d{1,4}[,.]\d{1,4})\s*/\s*(\d{1,4}[,.]\d{1,4})\s*/\s*(\d{1,4}[,.]\d{1,4})'
)

# Quantity with Bulgarian or EN unit
_QTY_RE = re.compile(r'\b(\d+)\s*(?:Брой|бр\.?|PCE|STK)\b', re.IGNORECASE)

# Unit price: "10,77/ 1 PCE"
_UNIT_PRICE_RE = re.compile(r'([\d,.]+)\s*/\s*1\s*PCE', re.IGNORECASE)

# Tokens that mark the start of technical specs on a position line
_OSRAM_SPEC_RE = re.compile(
    r'^\d+[.,]\d*[WwVvKk]'   # 1,8W  36V  2700K
    r'|^\d+[WwVvKk]$'         # 4W  12V
    r'|^PG\d'                  # PG20-1
    r'|^G\d+[.\-/]?\d*$'      # G4  G13
    r'|^E\d+$'                 # E14  E27
    r'|^\d{4}K$',              # 2700K
    re.IGNORECASE
)


def _extract_osram_product_code(pos_line: str) -> Optional[str]:
    """Extract catalog code (e.g. LEDPWL ACC 103 30X1) from a position line.

    Line format: 000020  5  [Cyrillic category]  CODE specs...  OSRAM  price
    """
    tokens = pos_line.strip().split()
    i = 0
    # Skip position number (000NNN)
    while i < len(tokens) and _POS_RE.match(tokens[i]):
        i += 1
    # Skip quantity (pure integer)
    if i < len(tokens) and re.match(r'^\d+$', tokens[i]):
        i += 1
    # Skip Cyrillic category words
    while i < len(tokens) and re.search(r'[а-яА-Я]', tokens[i]):
        i += 1
    # Collect code tokens until specs / supplier name / price
    code_tokens = []
    while i < len(tokens):
        tok = tokens[i]
        if _OSRAM_SPEC_RE.match(tok):
            break
        if tok.upper().rstrip('.,') in {'OSRAM', 'LEDVANCE', 'PHILIPS'}:
            break
        if re.match(r'^\d{1,6}[.,]\d{2}$', tok):
            break
        code_tokens.append(tok)
        i += 1
    result = ' '.join(code_tokens).strip()
    return result if len(result) >= 3 else None


def _is_osram_document(text: str) -> bool:
    return bool(re.search(r'OSRAM\s+GMBH|ams-osram|OSRAM\s+GmbH', text, re.IGNORECASE))


def extract_osram_products(text: str) -> list[ProductRecord]:
    """
    Extract product rows from OSRAM delivery note.

    Primary method: anchor on OSRAM article number (AM460790055 etc.)
    The article line also contains EAN, weight triplet, and country.
    Quantity and price are found in the surrounding context lines.
    """
    lines = text.splitlines()
    return _parse_osram_by_article(lines)


def _parse_osram_blocks(lines: list[str], block_starts: list[int]) -> list[ProductRecord]:
    records = []
    for bi, start in enumerate(block_starts):
        end = block_starts[bi + 1] if bi + 1 < len(block_starts) else len(lines)
        block = lines[start:end]
        block_text = "\n".join(block)

        # AM article (IC код) — used only for DB lookup, not shown as product code
        art_m = OSRAM_ARTICLE_RE.search(block_text)
        if not art_m:
            continue
        osram_article = art_m.group(1)

        rec = ProductRecord(extraction_method="osram")

        # ── product_code: filled by DB lookup (internal code like LEDIL432)
        # Leave None here — main.py enrichment will set it from info.internal_code
        rec.product_code = None
        rec._osram_article = osram_article  # type: ignore[attr-defined]

        # ── product_name: description from position line, Cyrillic prefix stripped
        # Line format: "000020  5  [Кирилски категория] OSRAM_CODE specs  OSRAM  53,85"
        first_line_tokens = block[0].strip().split()
        desc_tokens = []
        skip_leading = True
        _SUPPLIERS = {"OSRAM", "BOSRAM", "2BOSRAM", "LEDVANCE", "PHILIPS"}
        for tok in first_line_tokens:
            if skip_leading and re.match(r'^\d+$', tok):
                continue                                  # skip pos number and qty
            skip_leading = False
            if re.match(r'^\d{1,6}[.,]\d{2}$', tok):    # trailing total price
                break
            if tok.upper().rstrip(".,") in _SUPPLIERS:   # supplier name
                continue
            if not re.search(r'[а-яА-Я]', tok):           # keep non-Cyrillic tokens
                desc_tokens.append(tok)
        candidate = " ".join(desc_tokens).strip()
        if len(candidate) >= 3:
            rec.product_name = candidate[:80]

        # ── EAN: 13 or 14 consecutive digits
        ean_m = re.search(r'(?<!\d)(\d{13}|\d{14})(?!\d)', block_text)
        if ean_m:
            rec.ean = ean_m.group(1)

        # ── Количество: first "N Брой/PCE"
        qty_m = _QTY_RE.search(block_text)
        if qty_m:
            rec.quantity = qty_m.group(1) + " PCE"

        # ── Единична цена: "10,77/ 1 PCE"
        up_m = _UNIT_PRICE_RE.search(block_text)
        if up_m:
            rec.price = up_m.group(1).replace(",", ".") + " EUR"

        # ── Обща сума: rightmost decimal on position line (e.g. 53,85)
        total_m = re.search(r'\b(\d{1,6}[.,]\d{2})\s*$', block[0].strip())
        if total_m:
            rec.total_price = total_m.group(1) + " EUR"

        # ── Тегло: Нето / Бруто kg (1st and 2nd values of triplet — 3rd is cbm volume)
        wt_m = _WEIGHT_TRIPLET_RE.search(block_text)
        if wt_m:
            net   = wt_m.group(1).replace(",", ".")
            gross = wt_m.group(2).replace(",", ".")
            rec.weight_kg = f"{net} / {gross} kg"

        records.append(rec)

    seen: set[str] = set()
    return [r for r in records if r.product_code not in seen and not seen.add(r.product_code)]


def _parse_osram_by_article(lines: list[str]) -> list[ProductRecord]:
    """Primary extractor: anchor on OSRAM article number (AM/AA prefix codes)."""
    am_hits = sum(1 for l in lines if OSRAM_ARTICLE_RE.search(l.strip()))
    logger.info("OSRAM extraction: %d total lines, %d AM article hits", len(lines), am_hits)
    records = []
    for i, line in enumerate(lines):
        m = OSRAM_ARTICLE_RE.search(line.strip())
        if not m:
            continue

        before_lines = lines[max(0, i - 15):i]
        after_lines  = lines[i:min(len(lines), i + 10)]
        ctx_lines    = before_lines + after_lines
        ctx          = "\n".join(ctx_lines)
        after_ctx    = "\n".join(after_lines)

        rec = ProductRecord(extraction_method="osram")
        osram_article = m.group(1)
        rec.product_code = None           # set from position line below
        rec._osram_article = osram_article  # type: ignore[attr-defined]

        # EAN is on the same line as the article code in OSRAM invoices.
        # Search only ±2 lines to avoid picking up EAN from the previous block.
        ean_narrow = "\n".join(lines[max(0, i - 2):min(len(lines), i + 3)])
        ean_m = re.search(r'(?<!\d)(\d{13}|\d{14})(?!\d)', ean_narrow)
        if ean_m:
            rec.ean = ean_m.group(1)

        # ── Product code: "LEDPWL ACC 103 30X1" style — from position line
        for bl in reversed(before_lines):
            if _POS_RE.match(bl.strip()):
                pc = _extract_osram_product_code(bl)
                if pc:
                    rec.product_code = pc
                break
        if not rec.product_code:
            rec.product_code = osram_article  # fallback

        # ── Quantity: 2nd token on position line (000NNN  QTY  description...)
        # Iterate backwards through lines before the article to find nearest pos line
        for bl in reversed(before_lines):
            if _POS_RE.match(bl.strip()):
                tokens = bl.strip().split()
                # tokens[0]=position, tokens[1]=quantity (pure digits)
                if len(tokens) >= 2 and re.match(r'^\d+$', tokens[1]):
                    rec.quantity = tokens[1] + " PCE"
                break
        # Fallback: search for explicit "N Брой" in before context
        if not rec.quantity:
            for qty_m in re.finditer(r'\b(\d+)\s*(?:Брой|бр\.?)\b', "\n".join(before_lines), re.IGNORECASE):
                rec.quantity = qty_m.group(1) + " PCE"
                break

        # ── Unit price: "10,77/ 1 PCE" — look only AFTER the article line
        # (avoids recycling-fee lines that appear before the article)
        up_m = _UNIT_PRICE_RE.search(after_ctx)
        if up_m:
            rec.price = up_m.group(1).replace(",", ".") + " EUR"

        # ── Total price: rightmost decimal on the nearest position line
        for bl in reversed(before_lines):
            if _POS_RE.match(bl.strip()):
                pm = re.search(r'\b(\d{1,6}[.,]\d{2})\s*$', bl.strip())
                if pm:
                    rec.total_price = pm.group(1) + " EUR"
                break

        # ── Weight: net / gross kg (1st and 2nd triplet values; 3rd is cbm)
        wt_m = _WEIGHT_TRIPLET_RE.search(ctx)
        if wt_m:
            net   = wt_m.group(1).replace(",", ".")
            gross = wt_m.group(2).replace(",", ".")
            rec.weight_kg = f"{net} / {gross} kg"

        if rec.quantity:
            records.append(rec)
        else:
            logger.warning("OSRAM: dropped %s — no quantity found. Context: %s",
                           osram_article, " | ".join(before_lines[-4:]))

    seen: set[str] = set()
    deduped = [r for r in records if r.product_code not in seen and not seen.add(r.product_code)]
    if len(deduped) < len(records):
        logger.info("OSRAM dedup: %d → %d (removed %d duplicates)",
                    len(records), len(deduped), len(records) - len(deduped))
    return deduped


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
