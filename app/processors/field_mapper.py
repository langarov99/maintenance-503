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

# Exact strings that indicate a repeated column-header row masquerading as a data row
_HEADER_CODE_WORDS = frozenset([
    "article number", "article no.", "article no", "article",
    "code", "item code", "item number", "item no.", "item",
    "artikel", "artikelnummer", "codice",
    "nr.", "nr", "no.", "number",
    "référence", "ref",
    "код", "артикул", "арт",
])

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
    "product_code": ["код", "code", "art", "artikel", "codice", "item", "артикул", "арт", "nr", "no", "référence", "article", "number"],
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

    # Scan up to first 100 rows for header (files may have preamble rows)
    best_mapping = {}
    best_row = None

    for row_idx, row in enumerate(table[:100]):
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
        # Skip rows that look like empty or pure-whitespace product codes
        if rec.filled_count() >= 1 and rec.product_code and rec.product_code.strip():
            # Skip repeated column-header rows (e.g. "article number" appearing as data)
            if rec.product_code.lower().strip() in _HEADER_CODE_WORDS:
                continue
            # Extract leading number from descriptive parts_in_set values
            # e.g. "3-pcs (1 and 2 row of seats)" → "3"
            if rec.parts_in_set:
                pcs_m = re.match(r'^(\d+)', rec.parts_in_set.strip())
                if pcs_m:
                    rec.parts_in_set = pcs_m.group(1)
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

# Position line anchor: 000020, 001110 etc. OR "80-001" sub-line style
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
                if len(tokens) >= 2 and re.match(r'^\d{1,5}$', tokens[1]):
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
# Rezaw-Plast specific extractor
# ---------------------------------------------------------------------------

def _is_rezaw_plast_document(text: str) -> bool:
    return bool(re.search(r'rezaw.?plast', text, re.IGNORECASE))


def _parse_rezaw_plast_table(table: list[list]) -> list[ProductRecord]:
    """Parse one PDF table from a Rezaw-Plast price list.

    Expected columns: Description | Years of production | Article number | Net Price | EAN CODE
    Blue brand-header rows have no article number — tracked as context only.
    """
    if not table or len(table) < 2:
        return []

    # Locate header row by looking for 'article' and 'ean' keywords.
    # Search the full table — Excel files may have many preamble rows before headers.
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if "article" in joined and ("ean" in joined or "code" in joined):
            header_idx = i
            break
    if header_idx is None:
        logger.info("Rezaw-Plast: no header row found in table (%d rows)", len(table))
        return []
    logger.info("Rezaw-Plast: header row found at index %d", header_idx)

    headers = [str(c or "").lower().strip() for c in table[header_idx]]

    def find(kws):
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return None

    desc_idx  = 0                                          # always first column
    art_idx   = find(["article", "арт", "number", "номер", "kod", "code"])
    price_idx = find(["price", "цена", "preis", "net"])
    ean_idx   = find(["ean", "баркод", "barcode", "gtin"])
    years_idx = find(["year", "production", "год"])
    pcs_idx   = find(["pcs", "pieces", "set", "parts", "бр", "стелки", "количество"])

    if art_idx is None:
        return []

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(idx):
            if idx is None or idx >= len(row):
                return ""
            return str(row[idx] or "").strip()

        article_raw = cell(art_idx)
        # Some cells contain two codes: "210601 / 211201" — take the first
        art_m = re.match(r'(\d{3,8})', article_raw)
        # Skip brand/section header rows (no numeric article number)
        if not art_m:
            continue
        article = art_m.group(1)

        rec = ProductRecord(extraction_method="table")
        rec.product_code = article

        # Product name: description + years for full context
        desc  = cell(desc_idx)
        years = cell(years_idx) if years_idx is not None else ""
        name_parts = [p for p in [desc, years] if p and p.strip() and p.strip() != article]
        if name_parts:
            rec.product_name = " | ".join(name_parts)[:120]

        # Price
        price_raw = cell(price_idx) if price_idx is not None else ""
        if price_raw and re.match(r'^\d+[.,]\d+$', price_raw):
            rec.price = price_raw.replace(",", ".") + " EUR"

        # EAN
        ean_raw = cell(ean_idx) if ean_idx is not None else ""
        if re.match(r'^\d{8,14}$', ean_raw):
            rec.ean = ean_raw

        # Pieces in set — extract leading number from e.g. "3-pcs (1 and 2 row of seats)"
        pcs_raw = cell(pcs_idx) if pcs_idx is not None else ""
        if pcs_raw and pcs_raw.strip():
            pcs_m = re.match(r'^(\d+)', pcs_raw.strip())
            if pcs_m:
                rec.parts_in_set = pcs_m.group(1)

        records.append(rec)

    return records


def extract_rezaw_plast_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Rezaw-Plast: %d table(s) received", len(tables))
    records = []
    for table in tables:
        records.extend(_parse_rezaw_plast_table(table))
    logger.info("Rezaw-Plast extraction: %d records from %d tables", len(records), len(tables))
    return records


# ---------------------------------------------------------------------------
# Avisa-specific extractor
# ---------------------------------------------------------------------------

def _is_avisa_document(text: str) -> bool:
    return bool(re.search(r'avisa', text, re.IGNORECASE))


def _parse_avisa_table(table: list[list]) -> list[ProductRecord]:
    """Parse one PDF table from an Avisa invoice.

    Columns: No. | Code | Description | Barcode | Quantity | Unit |
             Net price (regular) | Discount [%] | Net price (discounted) |
             Tax rate | Net value | Tax value | Gross value

    Export: Code, Description, Barcode, Quantity, second Net price (after discount).
    """
    if not table or len(table) < 2:
        return []

    # Find header row — must contain "code" and "description"
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if "code" in joined and "description" in joined:
            header_idx = i
            break
    if header_idx is None:
        return []

    headers = [str(c or "").lower().strip() for c in table[header_idx]]
    logger.info("Avisa: header row at index %d: %s", header_idx, headers)

    def find(kws):
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return None

    code_idx    = find(["code"])
    desc_idx    = find(["description", "desc"])
    barcode_idx = find(["barcode", "ean", "bar"])
    qty_idx     = find(["quantit", "qty"])
    discount_idx = find(["discount", "rabat", "%"])

    # Second "net price" = the one after the Discount column
    price_idx = None
    net_hits = []
    for i, h in enumerate(headers):
        if "net price" in h or h.strip() in ("net price", "netprice"):
            net_hits.append(i)
    if len(net_hits) >= 2:
        price_idx = net_hits[1]
    elif discount_idx is not None:
        for i in range(discount_idx + 1, len(headers)):
            if "net" in headers[i]:
                price_idx = i
                break

    if code_idx is None:
        return []

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(idx):
            if idx is None or idx >= len(row):
                return ""
            return str(row[idx] or "").strip()

        code = cell(code_idx)
        if not code or code.lower() in ("no.", "no", "#", ""):
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code

        desc = cell(desc_idx) if desc_idx is not None else ""
        if desc:
            rec.product_name = desc[:120]

        barcode_raw = cell(barcode_idx) if barcode_idx is not None else ""
        if re.match(r'^\d{8,14}$', barcode_raw):
            rec.ean = barcode_raw

        qty_raw = cell(qty_idx) if qty_idx is not None else ""
        if qty_raw:
            try:
                qty_f = float(qty_raw.replace(",", "."))
                qty_str = str(int(qty_f)) if qty_f == int(qty_f) else str(qty_f)
            except ValueError:
                qty_str = qty_raw
            rec.quantity = qty_str + " szt."

        price_raw = cell(price_idx) if price_idx is not None else ""
        if price_raw and re.match(r'^\d+[.,]\d+$', price_raw):
            rec.price = price_raw.replace(",", ".") + " EUR"

        records.append(rec)

    return records


def extract_avisa_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Avisa: %d table(s) received", len(tables))
    records = []
    for table in tables:
        records.extend(_parse_avisa_table(table))
    logger.info("Avisa extraction: %d records from %d tables", len(records), len(tables))
    return records


# ---------------------------------------------------------------------------
# Amio-specific extractor
# ---------------------------------------------------------------------------

def _is_amio_document(text: str) -> bool:
    return bool(re.search(r'\bamio\b', text, re.IGNORECASE))


def _parse_amio_table(table: list[list]) -> list[ProductRecord]:
    """Parse one PDF table from an Amio invoice.

    Columns: Lp. | Kod produktu/Product number | Nazwa towaru/Product name |
             EAN | Ilość/Qty | J.m/Unit | VAT/Tax | Cena/Price EUR | Wartość/Value EUR

    Export: product code, product name, EAN, quantity (number only).
    """
    if not table or len(table) < 2:
        return []

    # Find header row containing product code and EAN columns
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if ("kod" in joined or "product number" in joined) and "ean" in joined:
            header_idx = i
            break
    if header_idx is None:
        return []

    headers = [str(c or "").lower().strip() for c in table[header_idx]]
    logger.info("Amio: header row at index %d", header_idx)

    def find(kws):
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return None

    code_idx  = find(["kod produktu", "product number", "kod"])
    desc_idx  = find(["nazwa towaru", "product name", "nazwa"])
    ean_idx   = find(["ean"])
    qty_idx   = find(["ilość", "qty", "quantity"])
    # Last "value" column — Wartość/Value EUR (last column in table)
    value_idx = find(["wartosc", "wartość", "value eur", "value"])
    if value_idx is None and len(headers) > 0:
        value_idx = len(headers) - 1  # fallback: last column

    if code_idx is None or ean_idx is None:
        return []

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(idx):
            if idx is None or idx >= len(row):
                return ""
            return str(row[idx] or "").strip()

        code = cell(code_idx)
        # Skip non-product rows: empty codes, row numbers (1–999), summary lines
        if not code or (code.isdigit() and len(code) <= 3):
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code

        desc = cell(desc_idx) if desc_idx is not None else ""
        if desc:
            rec.product_name = desc[:120]

        barcode_raw = cell(ean_idx)
        if re.match(r'^\d{8,14}$', barcode_raw):
            rec.ean = barcode_raw

        qty_raw = cell(qty_idx) if qty_idx is not None else ""
        if qty_raw:
            try:
                qty_f = float(qty_raw.replace(",", "."))
                qty_str = str(int(qty_f)) if qty_f == int(qty_f) else str(qty_f)
            except ValueError:
                qty_str = qty_raw
            rec.quantity = qty_str

        value_raw = cell(value_idx) if value_idx is not None else ""
        if value_raw and re.match(r'^\d+[.,]\d+$', value_raw):
            rec.price = value_raw.replace(",", ".") + " EUR"

        records.append(rec)

    return records


def extract_amio_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Amio: %d table(s) received", len(tables))
    records = []
    seen_codes: set[str] = set()
    for table in tables:
        for rec in _parse_amio_table(table):
            if rec.product_code not in seen_codes:
                seen_codes.add(rec.product_code)
                records.append(rec)
    logger.info("Amio extraction: %d records from %d tables", len(records), len(tables))
    return records


# ---------------------------------------------------------------------------
# Maxton Design-specific extractor
# ---------------------------------------------------------------------------

def _is_maxton_document(text: str) -> bool:
    return bool(re.search(r'maxton', text, re.IGNORECASE))


_MAXTON_CODE_RE = re.compile(r'^([A-Z]{2}-[A-Z0-9][A-Z0-9\-]+)\s+(.*)', re.DOTALL)
# No-anchor variant for text scanning; includes '+' for compound codes like FD1G+FD1RG
_MAXTON_CODE_TEXT_RE = re.compile(r'(?<!\w)([A-Z]{2}-[A-Z0-9][A-Z0-9\-\+]{3,})')


def _strip_diacritics(s: str) -> str:
    for src, dst in [('ą','a'),('ć','c'),('ę','e'),('ł','l'),('ń','n'),
                     ('ó','o'),('ś','s'),('ź','z'),('ż','z')]:
        s = s.replace(src, dst).replace(src.upper(), dst.upper())
    return s


def _parse_maxton_table(table: list[list]) -> list[ProductRecord]:
    """Parse one PDF table from a Maxton Design invoice.

    Columns: Lp. | Nazwa towaru/usługi | Ilość | J.m. | VAT |
             Cena netto EUR | Wartość netto EUR

    Product code is the first token of the description (e.g. BM-3-20-MPACK-FD5G).
    Export: product_code, product_name, quantity, Wartość netto EUR as price.
    """
    if not table or len(table) < 2:
        return []

    def norm(s):
        return _strip_diacritics(str(s or "").lower())

    # Find header row: contains "lp" AND ("nazwa" OR "netto")
    # Also accept row where first cell is "Lp." alone
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(norm(c) for c in row)
        first  = norm(row[0]) if row else ""
        if (first.strip().rstrip('.') == "lp" and "netto" in joined):
            header_idx = i
            break
        if "nazwa" in joined and ("ilosc" in joined or "qty" in joined or "netto" in joined):
            header_idx = i
            break
    if header_idx is None:
        logger.warning("Maxton: no header row found in table (%d rows)", len(table))
        return []

    headers = [norm(c) for c in table[header_idx]]
    logger.info("Maxton: header row at index %d: %s", header_idx, headers)

    def find(kws):
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return None

    desc_idx  = find(["nazwa towaru", "product name", "nazwa", "product"])
    qty_idx   = find(["ilosc", "qty", "quantity"])
    value_idx = find(["wartosc netto", "value eur", "wartosc", "value"])
    if value_idx is None:
        value_idx = len(headers) - 1  # fallback: last column

    if desc_idx is None:
        # Fallback: scan rows directly for Maxton product code pattern
        desc_idx = 0

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(idx):
            if idx is None or idx >= len(row):
                return ""
            return str(row[idx] or "").strip()

        raw_desc = cell(desc_idx)
        if not raw_desc:
            continue

        # Try to match Maxton product code pattern directly
        m = _MAXTON_CODE_RE.match(raw_desc)
        if m:
            code = m.group(1)
            name = m.group(2).strip()
        else:
            parts = raw_desc.split(None, 1)
            code  = parts[0]
            name  = parts[1].strip() if len(parts) > 1 else ""

        # Skip row numbers, header text, short/invalid codes
        if not code or (code.isdigit() and len(code) <= 3):
            continue
        if code.lower() in ("lp.", "lp", "nazwa", "no.", "no"):
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        if name:
            rec.product_name = name[:120]

        qty_raw = cell(qty_idx) if qty_idx is not None else ""
        if qty_raw:
            try:
                qty_f = float(qty_raw.replace(",", "."))
                n = int(qty_f) if qty_f == int(qty_f) else qty_f
                rec.quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
            except ValueError:
                rec.quantity = qty_raw

        value_raw = cell(value_idx)
        if value_raw and re.match(r'^\d+[.,]\d+$', value_raw):
            rec.price = value_raw.replace(",", ".") + " EUR"

        records.append(rec)

    return records


def _parse_maxton_from_text(text: str) -> list[ProductRecord]:
    """Text-based fallback when pdfplumber finds no usable tables."""
    records = []
    seen: set[str] = set()
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _MAXTON_CODE_TEXT_RE.search(line)
        if not m:
            i += 1
            continue

        code = m.group(1)
        if code in seen:
            i += 1
            continue

        # Description: text after the code; strip trailing 'N szt|kpl 0% price price'
        after = line[m.end():].strip()
        name = re.sub(r'\s+\d+\s+(?:szt|kpl)\b.*$', '', after, flags=re.IGNORECASE).strip()
        if not name:
            name = after

        # Append wrapped continuation lines (e.g. "V8 PACK", "AMG")
        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if (next_line
                    and not _MAXTON_CODE_TEXT_RE.search(next_line)
                    and not re.match(r'^\d+\s+[A-Z]{2}-', next_line)
                    and not re.search(r'\b\d+\s+(?:szt|kpl)\b', next_line, re.IGNORECASE)):
                name = (name + " " + next_line).strip()

        # Skip rows where description is a column header
        if re.search(r'\b(?:nazwa|ilosc|quantity|netto|brutto|lp\.?)\b', name, re.IGNORECASE):
            i += 1
            continue

        # Quantity with Bulgarian unit label
        qty_m = re.search(r'\b(\d{1,4})\s*(?:szt|kpl)\b', line, re.IGNORECASE)
        if qty_m:
            n = int(qty_m.group(1))
            quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
        else:
            quantity = None

        # Two prices on each line: Cena netto EUR (unit) and Wartość netto EUR (total)
        prices = re.findall(r'\b(\d{1,6}[.,]\d{2})\b', line)
        if len(prices) >= 2:
            price       = prices[-2].replace(',', '.') + ' EUR'
            total_price = prices[-1].replace(',', '.') + ' EUR'
        elif len(prices) == 1:
            price       = prices[-1].replace(',', '.') + ' EUR'
            total_price = None
        else:
            price       = None
            total_price = None

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name[:120] if name else None
        rec.quantity     = quantity
        rec.price        = price
        rec.total_price  = total_price

        seen.add(code)
        records.append(rec)
        i += 1

    # Handle shipping/freight row — no standard XX-XXXX code, price merged in PDF text
    # e.g. "41 265,000 % 265,001SHIPPING EXPORT+WDT Wysyłka / Shipping DB Schenker"
    for raw_line in lines:
        line = raw_line.strip()
        if not re.search(r'\b(?:shipping|wysyłka|wyslka|freight)\b', line, re.IGNORECASE):
            continue
        # Lenient price regex — no word boundaries to handle merged digits
        prices = list(dict.fromkeys(re.findall(r'(\d{1,6}[.,]\d{2})', line)))
        if not prices:
            continue
        price_val = prices[-1].replace(',', '.') + ' EUR'
        rec = ProductRecord(extraction_method="table")
        rec.product_code  = "SHIPPING"
        rec.product_name  = "Wysyłka / Shipping"
        rec.quantity      = "1"
        rec.price         = price_val
        rec.total_price   = price_val
        records.append(rec)
        logger.info("Maxton: shipping row added — %s", price_val)
        break

    logger.info("Maxton text extraction: %d records", len(records))
    return records


def extract_maxton_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Maxton Design: %d table(s) received", len(tables))
    records = []
    seen_codes: set[str] = set()
    for table in tables:
        for rec in _parse_maxton_table(table):
            if rec.product_code not in seen_codes:
                seen_codes.add(rec.product_code)
                records.append(rec)

    # Fallback to text-based extraction if tables yielded nothing
    if not records and text:
        logger.info("Maxton: no table records — trying text extraction")
        records = _parse_maxton_from_text(text)

    logger.info("Maxton Design extraction: %d records total", len(records))
    return records


# ---------------------------------------------------------------------------
# M-Tech Poland-specific extractor
# ---------------------------------------------------------------------------

def _clean_num(raw: str) -> str:
    """Strip non-numeric chars, round to max 3 decimal places, ensure 2 minimum."""
    p = re.sub(r'[^\d.,]', '', str(raw)).replace(',', '.')
    if not p:
        return ""
    try:
        val = round(float(p), 3)
        # Format with 3 decimals then strip trailing zeros, keep minimum 2
        s = f"{val:.3f}".rstrip('0')
        if '.' not in s:
            s += '.00'
        elif s.endswith('.'):
            s += '00'
        elif len(s.split('.')[1]) < 2:
            s += '0'
        return s
    except ValueError:
        return p


def _is_mtech_document(text: str) -> bool:
    return bool(re.search(r'm[-\s]?tech', text, re.IGNORECASE))


def extract_mtech_products(tables: list, text: str = "") -> list[ProductRecord]:
    """
    M-Tech Excel invoice: 2 rows per product.
      Row 1 (bold):   No. | product_code | CN Code | Country | EAN | ... | Weight
      Row 2 (italic):     | description  |         |         |     | Qty | unit | VAT | Unit price | Total net value
    Row 1 is identified by a sequential integer in the first (No.) column.
    """
    records = []
    logger.info("M-Tech: %d table(s) received", len(tables))

    for table in tables:
        if not table or len(table) < 3:
            continue

        def norm(s):
            return str(s or "").lower().strip()

        # Find header row: must contain "item description" or "description" + "ean"
        header_idx = None
        for i, row in enumerate(table):
            joined = " ".join(norm(c) for c in row)
            if ("item description" in joined or
                    ("description" in joined and ("ean" in joined or "qty" in joined))):
                header_idx = i
                break
        if header_idx is None:
            logger.warning("M-Tech: no header row found in table (%d rows)", len(table))
            continue

        headers = [norm(c) for c in table[header_idx]]
        logger.info("M-Tech header: %s", headers)

        def find_col(kws):
            for kw in kws:
                for idx, h in enumerate(headers):
                    if kw in h:
                        return idx
            return None

        no_idx     = find_col(["no.", "no ", "lp.", "pos"]) or 0
        desc_idx = find_col(["item description", "description", "item"]) or 1

        def cell(row, idx):
            if idx is None or idx >= len(row):
                return ""
            # Strip regular and non-breaking spaces
            return re.sub(r'[\xa0\s]+', ' ', str(row[idx] or "")).strip()

        # Calibrate actual column positions from the first product's rows.
        # Header indices are unreliable due to merged cells in the Excel.
        ean_pos = weight_pos = qty_pos = price_pos = total_pos = None

        data_rows = table[header_idx + 1:]
        for dr in data_rows:
            code = cell(dr, desc_idx)
            if not (code and re.match(r'^[A-Z][A-Z0-9\-/]{1,15}$', code)):
                continue
            # code_row found — desc_row is 2 rows later
            if data_rows.index(dr) + 2 >= len(data_rows):
                break
            desc_row_sample = data_rows[data_rows.index(dr) + 2]
            for ci, val in enumerate(dr):
                v = cell(dr, ci)
                if re.match(r'^\d{8,14}$', v):
                    ean_pos = ci          # EAN in code row
                elif re.match(r'^\d+\.\d+$', v) and 0.05 < float(v) < 100:
                    weight_pos = ci       # Weight in code row (small decimal)
            for ci, val in enumerate(desc_row_sample):
                v = cell(desc_row_sample, ci)
                if ci == ean_pos and re.match(r'^\d+$', v):
                    qty_pos = ci          # Qty in desc row at same col as EAN
                elif re.match(r'^\d+[.,]\d{2}$', v) and price_pos is None and ci > (ean_pos or 0):
                    price_pos = ci        # First clean decimal after qty → unit price
                elif re.search(r'\d+[.,]\d{2}.*EUR', v) and ci > (price_pos or 0):
                    total_pos = ci        # "53,15 EUR" pattern → total
            logger.info("M-Tech calibrated: ean=%s weight=%s qty=%s price=%s total=%s",
                        ean_pos, weight_pos, qty_pos, price_pos, total_pos)
            break

        i = 0
        while i < len(data_rows):
            row1 = data_rows[i]
            code = cell(row1, desc_idx)

            # Product code row: col 1 has a short uppercase code (e.g. CP14W, CP5S)
            # The No. number appears on the NEXT row, description on the row after that.
            if not (code and re.match(r'^[A-Z][A-Z0-9\-/]{1,15}$', code)):
                i += 1
                continue

            ean    = cell(row1, ean_pos) if ean_pos is not None else ""
            weight = cell(row1, weight_pos) if weight_pos is not None else ""

            # Structure per product: [code_row] [number_row] [description_row]
            desc_row  = data_rows[i + 2] if i + 2 < len(data_rows) else []
            desc      = cell(desc_row, desc_idx)
            qty_raw   = cell(desc_row, qty_pos) if qty_pos is not None else ""
            price_raw = cell(desc_row, price_pos) if price_pos is not None else ""
            total_raw = cell(desc_row, total_pos) if total_pos is not None else ""

            rec = ProductRecord(extraction_method="table")
            rec.product_code = code
            rec.product_name = desc or None

            ean_clean = re.sub(r'[^\d]', '', ean)
            if re.match(r'^\d{8,14}$', ean_clean):
                rec.ean = ean_clean

            if qty_raw:
                try:
                    n = int(float(qty_raw.replace(",", ".")))
                    rec.quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
                except ValueError:
                    rec.quantity = qty_raw

            if price_raw:
                p = _clean_num(price_raw)
                if p:
                    rec.price = p + " EUR"

            if total_raw:
                t = _clean_num(total_raw)
                if t:
                    rec.total_price = t + " EUR"

            if weight:
                rec.weight_kg = weight

            records.append(rec)
            i += 3  # code_row + number_row + desc_row

    logger.info("M-Tech extraction: %d records", len(records))
    return records


# ---------------------------------------------------------------------------
# Auto-switch orchestrator
# ---------------------------------------------------------------------------

class FieldMapper:
    def __init__(self, llm=None):
        self.llm = llm  # Optional llama-cpp-python Llama instance

    def map(self, extracted: dict, supplier: str = "auto") -> list[ProductRecord]:
        tables = extracted.get("tables", [])
        text = extracted.get("text", "")

        logger.info("Extraction requested: supplier=%s", supplier)

        # Step 0 — OSRAM (explicit selection or auto-detection)
        if supplier == "osram" or (supplier == "auto" and text and _is_osram_document(text)):
            if text:
                records = extract_osram_products(text)
                if records:
                    logger.info("Extraction method: OSRAM-specific (%d records)", len(records))
                    return records

        # Step 0b — Rezaw-Plast (explicit selection or auto-detection)
        if supplier == "rezaw_plast" or (supplier == "auto" and _is_rezaw_plast_document(text)):
            records = extract_rezaw_plast_products(tables, text)
            if records:
                logger.info("Extraction method: Rezaw-Plast (%d records)", len(records))
                return records

        # Step 0c — Avisa (explicit selection or auto-detection)
        if supplier == "avisa" or (supplier == "auto" and _is_avisa_document(text)):
            records = extract_avisa_products(tables, text)
            if records:
                logger.info("Extraction method: Avisa (%d records)", len(records))
                return records

        # Step 0d — Amio (explicit selection or auto-detection)
        if supplier == "amio" or (supplier == "auto" and _is_amio_document(text)):
            records = extract_amio_products(tables, text)
            if records:
                logger.info("Extraction method: Amio (%d records)", len(records))
                return records

        # Step 0e — Maxton Design (explicit selection or auto-detection)
        if supplier == "maxton_design" or (supplier == "auto" and _is_maxton_document(text)):
            records = extract_maxton_products(tables, text)
            if records:
                logger.info("Extraction method: Maxton Design (%d records)", len(records))
                return records

        # Step 0f — M-Tech Poland (explicit selection or auto-detection)
        if supplier == "mtech" or (supplier == "auto" and _is_mtech_document(text)):
            records = extract_mtech_products(tables, text)
            if records:
                logger.info("Extraction method: M-Tech (%d records)", len(records))
                return records

        # For explicitly selected non-OSRAM supplier skip straight to table/regex
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
