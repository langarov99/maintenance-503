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
    is_new_product: bool = True             # Да = не е намерен в каталога на доставчика
    extraction_method: str = "regex"
    merged_count: int = 1                  # how many invoice rows this record represents

    def filled_count(self) -> int:
        fields = [self.product_code, self.quantity, self.price,
                  self.product_name, self.ean, self.weight_kg,
                  self.parts_in_set, self.color]
        return sum(1 for f in fields if f and str(f).strip())

    def to_dict(self) -> dict:
        d = asdict(self)
        d.pop('merged_count', None)
        return d


def _smart_merge_or_add(records: list, seen_map: dict, rec: "ProductRecord") -> None:
    """Generic smart duplicate handler for all suppliers.

    Same code + same unit price → merge (sum qty and total).
    Same code + different unit price → keep as a separate row.
    seen_map must be dict[code, list[int]] (code → indices in records).
    """
    code = rec.product_code
    if code in seen_map:
        for idx in seen_map[code]:
            existing = records[idx]
            if existing.price == rec.price:
                existing.merged_count += 1
                if rec.quantity and existing.quantity:
                    try:
                        em = re.match(r'(\d+)', existing.quantity.strip())
                        nm = re.match(r'(\d+)', str(rec.quantity).strip())
                        if em and nm:
                            total_n = int(em.group(1)) + int(nm.group(1))
                            existing.quantity = re.sub(r'^\d+', str(total_n), existing.quantity, count=1)
                    except Exception:
                        pass
                if rec.total_price and existing.total_price:
                    try:
                        ep = float(existing.total_price.replace(',', '.').split()[0])
                        np_ = float(rec.total_price.replace(',', '.').split()[0])
                        suffix = ' EUR' if 'EUR' in (existing.total_price or '') else ''
                        existing.total_price = f"{ep + np_:.2f}{suffix}"
                    except Exception:
                        pass
                return
        seen_map[code].append(len(records))
    else:
        seen_map[code] = [len(records)]
    records.append(rec)


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
    "kod",
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
    "product_code": ["арт. № дост", "арт.№дост", "арт № дост", "код", "code", "art", "artikel",
                     "codice", "item", "артикул", "арт", "nr", "no", "référence", "article",
                     "number", "kod"],
    "quantity":     ["к-во", "кол-во", "кол", "qty", "quantity", "menge", "anzahl", "quantità",
                     "ilość", "množství", "доставено", "delivered", "geliefert", "consegnato",
                     "поръчано", "ordered", "бр"],
    "price":        ["ед.цена", "ед. цена", "единична цена", "unit price", "цена", "price",
                     "preis", "prezzo", "cena", "prix"],
    "total_price":  ["общо", "total", "wartość", "total value", "net value", "gross value",
                     "gesamtwert", "total price", "total amount", "valore totale"],
    "product_name": ["наименование", "описание", "продукт", "name", "bezeichnung", "nome",
                     "nazwa", "název", "description", "omschrijving", "клиентско", "artikel"],
    "ean":          ["ean", "баркод", "barcode", "gtin", "upc", "ean код"],
    "weight_kg":    ["кг", "kg", "weight", "gewicht", "peso", "waga", "hmotnost", "брутo",
                     "нето", "brutto", "netto", "gross", "net", "тегло"],
    "parts_in_set": ["единични", "пълни", "pcs", "pieces", "stück", "set", "комплект",
                     "sztuk", "ks"],
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

# Position line anchor: 000020, 001110 etc. OR "80-001" sub-line style.
# Negative lookahead (?!\.\d) excludes sub-number lines like "000045.001".
_POS_RE = re.compile(r'^(0{2,5}\d{1,4}(?!\.\d)|\d{2,3}-\d{3})\b')

# Numeric-prefix OSRAM product codes embedded at the start of description lines,
# concatenated with wattage/spec without a separator: "64210DWNBSP-1HB16W12V..."
# The (?=\d) lookahead stops the match just before the wattage digits begin.
_OSRAM_NUM_CODE_RE = re.compile(r'^(\d{3,5}[A-Z]+-\d+[A-Z]+)(?=\d)')

# Weight triplet: "1,200/ 1,232/ 0,009"
# Invoice columns: Нето (kg) / Брутo (kg) / Обем (cbm)  — take group 1 and 2 (kg only)
_WEIGHT_TRIPLET_RE = re.compile(
    r'(\d{1,4}[,.]\d{1,4})\s*/\s*(\d{1,4}[,.]\d{1,4})\s*/\s*(\d{1,4}[,.]\d{1,4})'
)

# Quantity with Bulgarian or EN unit
_QTY_RE = re.compile(r'\b(\d+)\s*(?:Брой|бр\.?|PCE|STK)\b', re.IGNORECASE)

# Unit price: "10,77/ 1 PCE"
_UNIT_PRICE_RE = re.compile(r'([\d,.]+)\s*/\s*1\s*PCE', re.IGNORECASE)

# Code suffixes indicating 2-per-blister packaging (e.g. 2721-2BL, 62150CBB-2HB).
# When no explicit "(N Blister)" text is found, divide piece qty by 2 to get BLI count.
_BLISTER2_SUFFIX_RE = re.compile(r'-(?:02BL|2BL|2HB)$', re.IGNORECASE)

# Tokens that mark the start of technical specs on a position line
_OSRAM_SPEC_RE = re.compile(
    r'^\d+[.,]\d*[WwVvKk]'   # 1,8W  36V  2700K
    r'|^\d+[WwVvKk]$'         # 4W  12V
    r'|^\d+/\d+[WwVvKkAa]'   # 35/35W  12/24V
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


def _extract_osram_num_code(line: str) -> Optional[str]:
    """Extract numeric-prefix code from a description line where code and spec are
    concatenated without a separator, e.g. '64210DWNBSP-1HB16W12VPX26D4X100TRG2OSRAM'
    → '64210DWNBSP-1HB'. The wattage digits (16W, 13W …) mark the boundary."""
    m = _OSRAM_NUM_CODE_RE.match(line.strip())
    return m.group(1) if m else None


def _product_code_from_pos_and_next(lines: list[str], pos_idx: int) -> Optional[str]:
    """Try to extract product code from position line and the 1-3 lines that follow it."""
    pc = _extract_osram_product_code(lines[pos_idx])
    for la in range(1, 4):
        if pc:
            break
        next_idx = pos_idx + la
        if next_idx >= len(lines):
            break
        nxt = lines[next_idx].strip()
        if not nxt or _POS_RE.match(nxt):  # blank or next product's position line
            break
        pc = _extract_osram_product_code(lines[next_idx]) or _extract_osram_num_code(nxt)
    return pc


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

        # ── Обща сума: rightmost European-format decimal on position line
        total_m = re.search(r'\b(\d{1,3}(?:\.\d{3})*,\d{2})\s*$', block[0].strip())
        if total_m:
            raw_total = float(total_m.group(1).replace('.', '').replace(',', '.'))
            rec.total_price = f"{raw_total:.2f} EUR"

        # ── Единична цена primary: sum all "X/ 1 PCE" lines within this block only
        blk_price_primary: Optional[float] = None
        up_matches = _UNIT_PRICE_RE.findall(block_text)  # block is already one product
        if up_matches:
            try:
                blk_price_primary = round(sum(float(p.replace(",", ".")) for p in up_matches), 2)
            except ValueError:
                try:
                    blk_price_primary = round(float(up_matches[0].replace(",", ".")), 2)
                except ValueError:
                    pass

        # ── Единична цена secondary: total / qty cross-check
        blk_price_secondary: Optional[float] = None
        if rec.total_price and rec.quantity:
            try:
                t_val = float(re.search(r'[\d.]+', rec.total_price).group())
                q_m = re.match(r'(\d+)', str(rec.quantity))
                if q_m and int(q_m.group(1)) > 0:
                    blk_price_secondary = round(t_val / int(q_m.group(1)), 2)
            except (ValueError, AttributeError, ZeroDivisionError):
                pass

        if blk_price_primary is not None and blk_price_secondary is not None:
            if abs(blk_price_primary - blk_price_secondary) <= 0.02:
                rec.price = f"{blk_price_primary:.2f} EUR"
            else:
                logger.warning(
                    "OSRAM block %s: price mismatch — Σ(/1PCE)=%.2f vs total/qty=%.2f"
                    " — using total/qty.",
                    osram_article, blk_price_primary, blk_price_secondary,
                )
                rec.price = f"{blk_price_secondary:.2f} EUR"
        elif blk_price_primary is not None:
            rec.price = f"{blk_price_primary:.2f} EUR"
        elif blk_price_secondary is not None:
            rec.price = f"{blk_price_secondary:.2f} EUR"

        # ── Тегло: Нето / Бруто kg (1st and 2nd values of triplet — 3rd is cbm volume)
        wt_m = _WEIGHT_TRIPLET_RE.search(block_text)
        if wt_m:
            net   = wt_m.group(1).replace(",", ".")
            gross = wt_m.group(2).replace(",", ".")
            rec.weight_kg = f"{net} / {gross} kg"

        records.append(rec)

    seen: dict[str, list[int]] = {}
    deduped: list[ProductRecord] = []
    for r in records:
        _smart_merge_or_add(deduped, seen, r)
    return deduped


_POS_NUM_RE = re.compile(r'^(0{2,5}\d{1,4})(?!\.\d)')


def _parse_osram_by_article(lines: list[str]) -> list[ProductRecord]:
    """Primary extractor: anchor on OSRAM article number (AM/AA prefix codes)."""
    am_hits = sum(1 for l in lines if OSRAM_ARTICLE_RE.search(l.strip()))
    logger.info("OSRAM extraction: %d total lines, %d AM article hits", len(lines), am_hits)
    records = []

    # Collect every position-line number found in the document for gap detection.
    all_pos_nums: dict[str, tuple[int, str]] = {}   # pos_num → (line_index, line_text)
    for _li, _l in enumerate(lines):
        _pm = _POS_NUM_RE.match(_l.strip())
        if _pm:
            all_pos_nums[_pm.group(1)] = (_li, _l.strip())

    # Forward pre-scan: build phase1_pos_nums and am_line_to_pos BEFORE Phase-1
    # loop so that _osram_pos can be set reliably from the pre-scan map rather
    # than an unreliable backward scan.
    _pos_by_line = sorted(
        (line_idx, pn) for pn, (line_idx, _) in all_pos_nums.items()
    )
    phase1_pos_nums: set[str] = set()
    am_line_to_pos: dict[int, str] = {}   # AM article line index → position number
    for _k, (_li, _pn) in enumerate(_pos_by_line):
        _next_li = _pos_by_line[_k + 1][0] if _k + 1 < len(_pos_by_line) else len(lines)
        for _j in range(_li + 1, _next_li):
            if OSRAM_ARTICLE_RE.search(lines[_j].strip()):
                phase1_pos_nums.add(_pn)
                am_line_to_pos[_j] = _pn
                break

    # Track the last seen position line so cross-page products (where the position
    # line is on page N but the article line is on page N+1 separated by headers)
    # can still recover quantity, product_code, and total_price.
    last_pos: dict = {"quantity": None, "product_code": None, "total_price": None,
                      "pos_num": None}

    for i, line in enumerate(lines):
        # Update last-seen position line as we scan forward.
        # Also checks the 1-3 lines after the position line for the product code,
        # which is needed when the code is on a description line (e.g. for LED
        # downlights where the code is concatenated with the spec string).
        _s = line.strip()
        if _POS_RE.match(_s):
            _t = _s.split()
            _qty = _t[1] if len(_t) >= 2 and re.match(r'^\d{1,5}$', _t[1]) else None
            _pc = _product_code_from_pos_and_next(lines, i)
            _pm = re.search(r'\b(\d{1,3}(?:\.\d{3})*,\d{2})\s*$', _s)
            _pn = _POS_NUM_RE.match(_s)
            last_pos = {
                "quantity": _qty,
                "product_code": _pc,
                "total_price": (
                    f"{float(_pm.group(1).replace('.', '').replace(',', '.')):.2f} EUR"
                    if _pm else None
                ),
                "pos_num": _pn.group(1) if _pn else None,
            }

        m = OSRAM_ARTICLE_RE.search(_s)
        if not m:
            continue

        before_lines = lines[max(0, i - 25):i]
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

        # ── Direct position line via pre-scan map (reliable even for cross-page products
        # where the position line may be >25 lines before the AM article code).
        # am_line_to_pos stores only the FIRST AM code per position; subsequent AM codes
        # within the same position block fall through to the backward-scan path below.
        _direct_pos_num = am_line_to_pos.get(i)
        _direct_pos_li: Optional[int] = None
        _direct_pos_lt: Optional[str] = None
        if _direct_pos_num and _direct_pos_num in all_pos_nums:
            _direct_pos_li, _direct_pos_lt = all_pos_nums[_direct_pos_num]

        # ── Product code: direct position line lookup first, backward scan as fallback
        if _direct_pos_li is not None:
            pc = _product_code_from_pos_and_next(lines, _direct_pos_li)
            if pc:
                rec.product_code = pc
        if not rec.product_code:
            for idx_r, bl in enumerate(reversed(before_lines)):
                if _POS_RE.match(bl.strip()):
                    pos_idx_in_before = len(before_lines) - 1 - idx_r
                    global_pos_idx = max(0, i - 25) + pos_idx_in_before
                    pc = _product_code_from_pos_and_next(lines, global_pos_idx)
                    if pc:
                        rec.product_code = pc
                    break
        if not rec.product_code:
            if last_pos["product_code"]:
                rec.product_code = last_pos["product_code"]
                logger.info("OSRAM %s: cross-page fallback product_code=%r",
                            osram_article, rec.product_code)
            else:
                rec.product_code = osram_article  # final fallback

        # ── Quantity: 2nd token on position line; direct lookup first, backward scan fallback
        if _direct_pos_lt:
            _dpt = _direct_pos_lt.split()
            _dti = 0
            while _dti < len(_dpt) and _POS_NUM_RE.match(_dpt[_dti]):
                _dti += 1
            if _dti < len(_dpt) and re.match(r'^\d{1,5}$', _dpt[_dti]):
                rec.quantity = _dpt[_dti] + " PCE"
        if not rec.quantity:
            for bl in reversed(before_lines):
                if _POS_RE.match(bl.strip()):
                    tokens = bl.strip().split()
                    if len(tokens) >= 2 and re.match(r'^\d{1,5}$', tokens[1]):
                        rec.quantity = tokens[1] + " PCE"
                    break
        if not rec.quantity:
            for qty_m in re.finditer(r'\b(\d+)\s*(?:Брой|бр\.?)\b', ctx, re.IGNORECASE):
                rec.quantity = qty_m.group(1) + " PCE"
                break
        if not rec.quantity:
            paren_m = re.search(r'^\s*\d+\s+\((\d+)\)\b', "\n".join(after_lines), re.MULTILINE)
            if paren_m:
                rec.quantity = paren_m.group(1) + " PCE"
                logger.info("OSRAM %s: paren-quantity fallback quantity=%s",
                            osram_article, paren_m.group(1))
        if not rec.quantity and last_pos["quantity"]:
            rec.quantity = last_pos["quantity"] + " PCE"
            logger.info("OSRAM %s: cross-page fallback quantity=%s",
                        osram_article, last_pos["quantity"])

        # ── Record position number (from pre-scan map; fallback to last_pos)
        rec._osram_pos = _direct_pos_num or last_pos.get("pos_num")  # type: ignore[attr-defined]

        # ── Total price: direct position line first (most reliable), then backward scan.
        # The backward-scan-only approach fails when the position line is on the previous
        # page (>25 lines before the AM code) and an intermediate Phase-2 position line
        # is within the 25-line window — that causes the wrong total to be captured.
        if _direct_pos_lt:
            pm = re.search(r'\b(\d{1,3}(?:\.\d{3})*,\d{2})\s*$', _direct_pos_lt)
            if pm:
                rec.total_price = f"{float(pm.group(1).replace('.', '').replace(',', '.')):.2f} EUR"
        if not rec.total_price:
            for bl in reversed(before_lines):
                if _POS_RE.match(bl.strip()):
                    pm = re.search(r'\b(\d{1,3}(?:\.\d{3})*,\d{2})\s*$', bl.strip())
                    if pm:
                        raw_total = float(pm.group(1).replace('.', '').replace(',', '.'))
                        rec.total_price = f"{raw_total:.2f} EUR"
                    break
        if not rec.total_price and last_pos["total_price"]:
            rec.total_price = last_pos["total_price"]
            logger.info("OSRAM %s: cross-page fallback total_price=%r",
                        osram_article, rec.total_price)

        # ── Unit price primary: sum ALL "X/ 1 PCE" lines (net price + any fees).
        # Narrow the search window to stop at the next position line so that
        # the next product's pricing lines are not included in the sum.
        price_win_end = len(after_lines)
        for _j in range(1, len(after_lines)):
            if _POS_RE.match(after_lines[_j].strip()):
                price_win_end = _j
                break
        price_ctx = "\n".join(after_lines[:price_win_end])

        # ── Blister packaging override: when the block contains "(N Blister)",
        # the sellable unit is the blister, not the individual piece.
        # Uses price_ctx (bounded to this product's window) to avoid picking up
        # "(N Blister)" from the next product's lines.
        # Setting qty=N here causes price_secondary (total/qty) below to yield the
        # correct BLI unit price; price_primary (/1 PCE) will mismatch and be discarded.
        _bli_m = re.search(r'\((\d+)\s+Blister\)', price_ctx, re.IGNORECASE)
        if _bli_m:
            rec.quantity = _bli_m.group(1) + " BLI"
            logger.info("OSRAM %s: blister packaging — quantity overridden to %s",
                        osram_article, rec.quantity)
        elif rec.product_code and _BLISTER2_SUFFIX_RE.search(rec.product_code):
            _qty_m2 = re.match(r'(\d+)', str(rec.quantity or ''))
            if _qty_m2:
                _bli_n = int(_qty_m2.group(1)) // 2
                if _bli_n > 0:
                    rec.quantity = str(_bli_n) + " BLI"
                    logger.info("OSRAM %s: suffix-based blister → quantity halved to %d BLI",
                                osram_article, _bli_n)
        price_primary: Optional[float] = None
        up_matches = _UNIT_PRICE_RE.findall(price_ctx)
        if up_matches:
            try:
                price_primary = round(sum(float(p.replace(",", ".")) for p in up_matches), 2)
            except ValueError:
                try:
                    price_primary = round(float(up_matches[0].replace(",", ".")), 2)
                except ValueError:
                    pass

        # ── Unit price secondary: total_price / quantity (cross-check).
        price_secondary: Optional[float] = None
        if rec.total_price and rec.quantity:
            try:
                t_val = float(re.search(r'[\d.]+', rec.total_price).group())
                q_m = re.match(r'(\d+)', str(rec.quantity))
                if q_m and int(q_m.group(1)) > 0:
                    price_secondary = round(t_val / int(q_m.group(1)), 2)
            except (ValueError, AttributeError, ZeroDivisionError):
                pass

        # Choose price: agree within 0.02 EUR → use primary (Σ /1 PCE values);
        # mismatch → use secondary (position-line total is authoritative) and warn.
        if price_primary is not None and price_secondary is not None:
            if abs(price_primary - price_secondary) <= 0.02:
                rec.price = f"{price_primary:.2f} EUR"
            else:
                logger.warning(
                    "OSRAM %s: unit price mismatch — Σ(/1PCE)=%.2f vs total/qty=%.2f"
                    " — using total/qty (position-line total is authoritative).",
                    osram_article, price_primary, price_secondary,
                )
                rec.price = f"{price_secondary:.2f} EUR"
        elif price_primary is not None:
            rec.price = f"{price_primary:.2f} EUR"
        elif price_secondary is not None:
            rec.price = f"{price_secondary:.2f} EUR"

        # ── Weight: net / gross kg (1st and 2nd triplet values; 3rd is cbm)
        wt_m = _WEIGHT_TRIPLET_RE.search(ctx)
        if wt_m:
            net   = wt_m.group(1).replace(",", ".")
            gross = wt_m.group(2).replace(",", ".")
            rec.weight_kg = f"{net} / {gross} kg"

        # ── total_price fallback: qty × price when position-line total was not found
        if not rec.total_price and rec.price and rec.quantity:
            try:
                p_val = float(re.search(r'[\d.]+', rec.price).group())
                q_m2  = re.match(r'(\d+)', str(rec.quantity))
                if q_m2 and int(q_m2.group(1)) > 0:
                    computed = round(p_val * int(q_m2.group(1)), 2)
                    rec.total_price = f"{computed:.2f} EUR"
                    logger.warning("OSRAM %s: total_price missing — computed from price×qty=%s",
                                   osram_article, rec.total_price)
            except (ValueError, AttributeError):
                pass
        if not rec.total_price:
            logger.warning("OSRAM %s: total_price is None after all extraction attempts", osram_article)

        if rec.quantity:
            records.append(rec)
        else:
            logger.warning("OSRAM: dropped %s — no quantity found. Context: %s",
                           osram_article, " | ".join(before_lines[-4:]))

    # ── Position-sequence gap detection + Phase-2 extraction
    # Products like traditional halogen/auxiliary bulbs (2721, 64210, 7528…)
    # have NO AM/AA article code in the invoice — their catalog code sits
    # directly on the position line.  Collect these and extract them now.
    #
    # Gap detection: forward-scan approach.  For each position line scan
    # forward up to the next position line; if an AM/AA article code appears
    # in that window the position is a Phase-1 product, otherwise Phase-2.
    # This is more reliable than the _osram_pos backward-assignment approach,
    # which fails for cross-page products (article on next page after a page
    # break) and can cause double-counting when a Phase-1 backward scan
    # mistakenly picks up a Phase-2 position line's total.
    unmatched = {k: v for k, v in all_pos_nums.items() if k not in phase1_pos_nums}
    logger.info("OSRAM gap detection: %d Phase-1 positions, %d Phase-2 positions",
                len(phase1_pos_nums), len(unmatched))

    phase2_added = 0
    for pos_num, (pos_line_idx, pos_line) in sorted(unmatched.items()):
        tokens = pos_line.split()
        # Skip position number token(s)
        ti = 0
        while ti < len(tokens) and _POS_NUM_RE.match(tokens[ti]):
            ti += 1
        # Quantity: next pure-integer token
        qty: Optional[str] = None
        if ti < len(tokens) and re.match(r'^\d{1,5}$', tokens[ti]):
            qty = tokens[ti]
            ti += 1
        # Product code via standard extractor; strip concatenated specs from
        # numeric-prefix codes like "7528ULT-2BL21/5W..." → "7528ULT-2BL"
        pc = _extract_osram_product_code(pos_line)
        if pc:
            cleaned = _extract_osram_num_code(pc)
            if cleaned:
                pc = cleaned
        # Total price at end of line
        pm2 = re.search(r'\b(\d{1,3}(?:\.\d{3})*,\d{2})\s*$', pos_line.strip())
        total_price: Optional[str] = None
        if pm2:
            total_price = f"{float(pm2.group(1).replace('.', '').replace(',', '.')):.2f} EUR"

        if not pc or not qty:
            logger.warning("OSRAM pos-only %s: cannot extract code/qty — skipping: %s",
                           pos_num, pos_line[:80])
            continue

        rec2 = ProductRecord(extraction_method="osram")
        rec2.product_code = pc
        rec2.quantity = qty + " PCE"
        rec2.total_price = total_price
        rec2._osram_article = pc   # type: ignore[attr-defined]
        rec2._osram_pos = pos_num  # type: ignore[attr-defined]

        # Scan up to 30 lines for EAN and blister info. The wider window handles
        # cross-page products where the sub-line data appears after a page header
        # (~15-20 lines). Stop conditions prevent picking up the next product's data.
        ean_m2 = None
        bli_qty_int: Optional[int] = None
        for _scan_ln in lines[pos_line_idx + 1 : pos_line_idx + 30]:
            _sl = _scan_ln.strip()
            if not _sl:
                continue
            if _POS_RE.match(_sl):          # next product's position line
                break
            if OSRAM_ARTICLE_RE.search(_sl):  # AM code → next Phase-1 product
                break
            if ean_m2 is None:
                _ean_hit = re.search(r'(?<!\d)(\d{13}|\d{14})(?!\d)', _sl)
                if _ean_hit:
                    ean_m2 = _ean_hit
            if bli_qty_int is None:
                _bli_m2 = re.search(r'\((\d+)\s+Blister\)', _sl, re.IGNORECASE)
                if _bli_m2:
                    bli_qty_int = int(_bli_m2.group(1))
                    rec2.quantity = str(bli_qty_int) + " BLI"
                    logger.info("OSRAM pos-only %s: blister packaging — quantity overridden to %d BLI",
                                pos_num, bli_qty_int)
            if ean_m2 is not None and bli_qty_int is not None:
                break  # found both — no need to scan further
        if ean_m2:
            rec2.ean = ean_m2.group(1)
            logger.info("OSRAM pos-only %s: found EAN %s in sub-lines", pos_num, ean_m2.group(1))

        if bli_qty_int is None and pc and _BLISTER2_SUFFIX_RE.search(pc):
            _bli_n2 = int(qty) // 2
            if _bli_n2 > 0:
                bli_qty_int = _bli_n2
                rec2.quantity = str(_bli_n2) + " BLI"
                logger.info("OSRAM pos-only %s: suffix-based blister → quantity halved to %d BLI",
                            pos_num, _bli_n2)

        qty_int = bli_qty_int if bli_qty_int else int(qty)
        if total_price and qty_int > 0:
            try:
                t2 = float(re.search(r'[\d.]+', total_price).group())
                rec2.price = f"{round(t2 / qty_int, 2):.2f} EUR"
            except (ValueError, AttributeError, ZeroDivisionError):
                pass
        records.append(rec2)
        phase2_added += 1
        logger.info("OSRAM pos-only %s: code=%r qty=%s total=%r price=%r ean=%r",
                    pos_num, pc, qty, total_price, rec2.price, getattr(rec2, 'ean', None))

    if phase2_added:
        logger.info("OSRAM Phase-2 (position-only): added %d records", phase2_added)
    elif not unmatched:
        logger.info("OSRAM position-gap: all %d positions matched (no gaps)", len(all_pos_nums))

    # Sort by invoice position number so export order matches the invoice.
    records.sort(key=lambda r: getattr(r, '_osram_pos', None) or '999999')

    # Position-pair dedup: one OSRAM position block can contain multiple AM sub-codes
    # (e.g. left indicator + right indicator + wiring harness).  All sub-codes extract
    # the SAME product_code from the shared position line, so _smart_merge_or_add would
    # sum their quantities (4+4+4=12) and totals (211×3=633) — wrong.
    # Solution: keep only the FIRST record per (product_code, _osram_pos) pair.
    # Records without a position number pass through unchanged (safety net).
    _seen_pos_pairs: set[tuple] = set()
    _pos_deduped: list[ProductRecord] = []
    for r in records:
        _pos = getattr(r, '_osram_pos', None)
        _code = r.product_code
        if _pos and _code:
            _pair = (_code, _pos)
            if _pair in _seen_pos_pairs:
                logger.info("OSRAM: dropping sub-article duplicate code=%r pos=%s", _code, _pos)
                continue
            _seen_pos_pairs.add(_pair)
        _pos_deduped.append(r)
    if len(_pos_deduped) < len(records):
        logger.info("OSRAM position-pair dedup: %d → %d (removed %d sub-article duplicates)",
                    len(records), len(_pos_deduped), len(records) - len(_pos_deduped))

    seen: dict[str, list[int]] = {}
    deduped: list[ProductRecord] = []
    for r in _pos_deduped:
        _smart_merge_or_add(deduped, seen, r)
    if len(deduped) < len(_pos_deduped):
        logger.info("OSRAM dedup: %d → %d (removed/merged %d duplicates)",
                    len(_pos_deduped), len(deduped), len(_pos_deduped) - len(deduped))
    return deduped


# ---------------------------------------------------------------------------
# Rezaw-Plast specific extractor
# ---------------------------------------------------------------------------

def _is_rezaw_plast_document(text: str) -> bool:
    return bool(re.search(r'rezaw.?plast', text, re.IGNORECASE))


def _parse_rezaw_plast_table(table: list[list]) -> list[ProductRecord]:
    """Parse one PDF table from a Rezaw-Plast invoice or price list.

    Supports two formats:
    - Price list: Description | Years | Article number | Net Price | EAN CODE
    - Invoice:    Lp | Indeks/Article | Product name | unit | Quantity | Unit price | Discount | Net price | Net value | VAT
    """
    if not table or len(table) < 2:
        return []

    # Locate header row — accept 'article' paired with 'ean'/'code' (price list)
    # or 'quantity'/'ilość' (invoice format without EAN).
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if "article" in joined and (
            "ean" in joined or "code" in joined
            or "quantity" in joined or "ilość" in joined
        ):
            header_idx = i
            break
    if header_idx is None:
        logger.info("Rezaw-Plast: no header row found in table (%d rows), first 3 rows: %s",
                    len(table), [[str(c or "")[:30] for c in r] for r in table[:3]])
        return []
    logger.info("Rezaw-Plast: header row found at index %d", header_idx)

    headers = [str(c or "").lower().strip() for c in table[header_idx]]

    def find(kws):
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return None

    desc_idx  = find(["product", "nazwa", "description", "desc", "opis"])
    if desc_idx is None:
        desc_idx = 0
    art_idx   = find(["article", "арт", "number", "номер", "indeks", "kod", "code"])
    qty_idx   = find(["quantity", "ilość"])
    # Prefer discounted net price over unit price (invoice has both columns)
    price_idx = find(["net price", "unit price", "price", "цена", "preis"])
    ean_idx   = find(["ean", "баркод", "barcode", "gtin"])
    total_idx = find(["net value", "wartość", "total", "нв"])
    years_idx = find(["year", "production", "год"])
    pcs_idx   = find(["pcs", "pieces", "set", "parts", "бр", "стелки", "количество"])

    if art_idx is None:
        return []

    def parse_money(raw: str) -> str:
        """Parse European-format price strings like '13,90 EUR' or '13.90'."""
        if not raw or not raw.strip():
            return ""
        m = re.match(r'^([\d\s.,]+)\s*([A-Za-z]+)?$', raw.strip())
        if not m:
            return ""
        num_str = m.group(1).strip()
        currency = (m.group(2) or "EUR").upper()
        # European format: comma = decimal separator, dot = thousands
        if ',' in num_str:
            num_str = num_str.replace(' ', '').replace('.', '').replace(',', '.')
        else:
            num_str = num_str.replace(' ', '').replace(',', '')
        try:
            float(num_str)
        except ValueError:
            return ""
        return f"{num_str} {currency}"

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(idx):
            if idx is None or idx >= len(row):
                return ""
            return str(row[idx] or "").strip()

        article_raw = cell(art_idx)
        # Handle "two codes in one cell": "210601 / 211201" → take the first.
        # Split only on space-slash-space to avoid breaking codes like "232110/B".
        if ' / ' in article_raw:
            article_raw = article_raw.split(' / ')[0].strip()
        # Match full article code: digits with optional letter suffix (e.g. 200103A)
        # and/or slash-suffix (e.g. 232110/B).
        art_m = re.match(r'(\d{3,8}[A-Za-z]*(?:/[A-Za-z0-9]+)?)', article_raw)
        if not art_m:
            # Non-numeric code (e.g. "palet" surcharge row) — keep as-is unless
            # it looks like a repeated header cell.
            if not article_raw or article_raw.lower() in {
                'article', 'indeks', 'indeks/article', 'kod', 'number', 'code'
            }:
                continue
            article = article_raw
        else:
            article = art_m.group(1)

        rec = ProductRecord(extraction_method="table")
        rec.product_code = article

        # Product name: description + years for full context
        desc  = cell(desc_idx)
        years = cell(years_idx) if years_idx is not None else ""
        name_parts = [p for p in [desc, years] if p and p.strip() and p.strip() != article]
        if name_parts:
            rec.product_name = " | ".join(name_parts)[:120]

        # Quantity (invoice format: plain integer like "2")
        if qty_idx is not None:
            qty_raw = cell(qty_idx).strip()
            if qty_raw and re.match(r'^\d+$', qty_raw):
                rec.quantity = qty_raw + " pcs"

        # Price
        if price_idx is not None:
            rec.price = parse_money(cell(price_idx))

        # Total
        if total_idx is not None:
            rec.total_price = parse_money(cell(total_idx))

        # EAN
        if ean_idx is not None:
            ean_raw = cell(ean_idx)
            if re.match(r'^\d{8,14}$', ean_raw):
                rec.ean = ean_raw

        # Pieces in set — extract leading number from e.g. "3-pcs (1 and 2 row of seats)"
        if pcs_idx is not None:
            pcs_raw = cell(pcs_idx)
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


def _parse_amio_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for Amio invoices when pdfplumber finds no tables.

    Two-step approach:
      1. Locate every row start: "N.  <4-6 digit code>" in the flattened text.
      2. For each segment between consecutive row starts, anchor on the
         13-digit EAN to split name / numeric fields.
    """
    records: list[ProductRecord] = []
    seen_rows: set[str] = set()       # dedup by row number
    seen_merge: dict[str, list[int]] = {}  # for _smart_merge_or_add

    # Flatten to one string — multi-line product names become contiguous.
    full = " ".join(ln.strip() for ln in text.splitlines() if ln.strip())

    # Step 1 — find all row starts.
    # Product codes may be purely numeric, alphanumeric, or contain hyphens (e.g. "04335", "SED31269", "04335-B").
    row_re = re.compile(r'\b(\d{1,3})\.\s+([A-Za-z0-9][A-Za-z0-9\-]{1,14})\s+')
    row_starts = list(row_re.finditer(full))
    detected_nums = [m.group(1) for m in row_starts]
    logger.info("Amio: detected row numbers (%d): %s", len(row_starts), detected_nums)
    if not row_starts:
        return records

    seen_codes: dict[str, list[int]] = {}  # for duplicate-code detection

    for i, rm in enumerate(row_starts):
        code = rm.group(2)
        seg_start = rm.end()
        seg_end   = row_starts[i + 1].start() if i + 1 < len(row_starts) else len(full)
        segment   = full[seg_start:seg_end]
        logger.debug("Amio: match %d → row_num=%s code=%s matched=%r seg[:80]=%r",
                     i, rm.group(1), code, rm.group(0), segment[:80])

        # Step 2 — anchor on 13-digit EAN.
        ean_m = re.search(r'\b(\d{13})\b', segment)
        if not ean_m:
            logger.warning("Amio: row %s (code=%s) has no 13-digit EAN — skipped. segment=%r",
                           rm.group(1), code, segment[:120])
            continue

        name  = re.sub(r'\s+', ' ', segment[:ean_m.start()]).strip()
        after = segment[ean_m.end():].strip()
        ean   = ean_m.group(1)

        # Parse after-EAN fields step by step — more robust than one big regex.
        # Layout: CN(8dig) QTY UOM ...price(may split)... VAT% total COO
        qty = price_str = total_str = None
        cn_m = re.match(r'(\d{8})\s+(.*)', after, re.DOTALL)
        if cn_m:
            rest = cn_m.group(2).strip()
            # Anchor on VAT% to separate price from total
            vat_m = re.search(r'\b\d+%\s+([\d,]+)', rest)
            if vat_m:
                total_str = vat_m.group(1).replace(',', '.')
                before_vat = rest[:vat_m.start()].strip()
                # QTY + UOM: search anywhere in before_vat; price part is optional
                # (for some rows pdfplumber places the price cell text after VAT%).
                bv_m = re.search(r'\b(\d+)\s+([a-zA-Z]{2,6})\b(?:\s+(.*))?',
                                 before_vat, re.DOTALL)
                if not bv_m:
                    logger.warning("Amio: no QTY/UOM in before_vat=%r", before_vat[:120])
                if bv_m:
                    qty = bv_m.group(1)
                    price_seg = (bv_m.group(3) or "").strip()
                    if price_seg:
                        pm = re.match(r'([\d,]+(?:\s+\d+)?)', price_seg)
                        if pm:
                            raw = pm.group(1).replace(' ', '').replace(',', '.')
                            try:
                                price_str = f"{round(float(raw), 2):.2f}"
                            except ValueError:
                                pass
                    # Price not in before_vat — derive from total / qty
                    if price_str is None and total_str and qty:
                        try:
                            price_str = f"{float(total_str) / int(qty):.2f}"
                        except (ValueError, ZeroDivisionError):
                            pass

            # Fallback name: pdfplumber places long (wrapped) product names AFTER
            # the numeric columns in the flattened text stream.
            # Layout tail: VAT% total COO <product_name>
            if not name and vat_m:
                after_total = rest[vat_m.end():].strip()
                coo_m = re.match(r'^[A-Z]{2}\s+(.*)', after_total, re.DOTALL)
                fallback = (coo_m.group(1) if coo_m else after_total).strip()
                if len(fallback) > 3:
                    name = re.sub(r'\s+', ' ', fallback)[:120]

        row_num = rm.group(1)
        if row_num in seen_rows:
            continue
        seen_rows.add(row_num)

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name[:120] if name else None
        rec.ean          = ean

        if qty:
            rec.quantity = qty
        if price_str:
            rec.price = price_str + ' EUR'
        if total_str:
            rec.total_price = total_str + ' EUR'

        if code in seen_codes:
            logger.warning(
                "Amio text: duplicate code %s — row_num=%s (prev at rows %s). "
                "qty=%s total=%s | match=%r seg[:80]=%r",
                code, rm.group(1), seen_codes[code],
                rec.quantity, rec.total_price,
                rm.group(0), segment[:80],
            )
        seen_codes.setdefault(code, []).append(rm.group(1))

        _smart_merge_or_add(records, seen_merge, rec)
        logger.info("Amio text: code=%s qty=%s total=%s",
                    code, rec.quantity, rec.total_price)

    return records


def extract_amio_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Amio: %d table(s) received", len(tables))
    records = []
    seen_codes: dict[str, list[int]] = {}
    for table in tables:
        for rec in _parse_amio_table(table):
            _smart_merge_or_add(records, seen_codes, rec)
    logger.info("Amio extraction: %d records from %d tables", len(records), len(tables))

    if not records and text:
        logger.info("Amio: no table records — trying text extraction")
        records = _parse_amio_from_text(text)
        logger.info("Amio text extraction: %d records", len(records))

    # Add AMIO- prefix to all product codes (both table and text paths).
    for rec in records:
        if rec.product_code and not rec.product_code.startswith("AMIO-"):
            rec.product_code = "AMIO-" + rec.product_code

    return records


# ---------------------------------------------------------------------------
# Maxton Design-specific extractor
# ---------------------------------------------------------------------------

def _is_maxton_document(text: str) -> bool:
    return bool(re.search(r'maxton', text, re.IGNORECASE))


_MAXTON_CODE_RE = re.compile(r'^([A-Z]{2}[A-Z0-9]*-[A-Z0-9][A-Z0-9\-\_]+)\s+(.*)', re.DOTALL)
# No-anchor variant for text scanning; includes '+' for compound codes like FD1G+FD1RG
# and '_' for suffix variants like RS1GO_O / RS1GO__O (PDF uses underscores, catalog uses hyphens)
_MAXTON_CODE_TEXT_RE = re.compile(r'(?<!\w)([A-Z]{2}[A-Z0-9]*-[A-Z0-9][A-Z0-9\-\+\_]{3,})')


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

    def find_last(kws):
        """Return the LAST column index matching any keyword (for rightmost EUR col)."""
        result = None
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    result = i
        return result

    desc_idx  = find(["nazwa towaru", "product name", "nazwa", "product"])
    qty_idx   = find(["ilosc", "qty", "quantity"])
    # Proforma uses '[eur]' headers; invoice uses 'wartosc netto' / 'value'
    # Two [eur] cols: first = unit price, last = total value
    price_idx = find(["cena netto", "unit price", "cena", "[eur]", "eur"])
    value_idx = find_last(["wartosc netto", "value eur", "wartosc", "value", "[eur]", "eur"])
    # If both found the same column (only one [eur] col), clear price_idx
    if price_idx is not None and price_idx == value_idx:
        price_idx = None
    if value_idx is None:
        # Last resort: rightmost non-empty column in header
        for i in range(len(headers) - 1, -1, -1):
            if headers[i].strip():
                value_idx = i
                break

    if desc_idx is None:
        # Fallback: scan rows directly for Maxton product code pattern
        desc_idx = 0

    logger.info("Maxton cols — desc=%s qty=%s price=%s total=%s",
                desc_idx, qty_idx, price_idx, value_idx)

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(idx):
            if idx is None or idx >= len(row):
                return ""
            return str(row[idx] or "").replace('\xa0', ' ').strip()

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

        # Normalize: ERP/catalog uses hyphens; proforma may use underscores (incl. double __)
        code = re.sub(r'_+', '-', code)

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        if name:
            rec.product_name = name[:120]

        # Quantity: try detected index and ±1 neighbors (header/data column shift)
        qty_raw = ""
        for qi in ([qty_idx] if qty_idx is not None else []) + [10, 9, 11, 8]:
            v = cell(qi)
            if v and re.match(r'^\d+([.,]\d+)?$', v):
                qty_raw = v
                break
        if qty_raw:
            try:
                qty_f = float(qty_raw.replace(",", "."))
                n = int(qty_f) if qty_f == int(qty_f) else qty_f
                rec.quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
            except ValueError:
                rec.quantity = qty_raw

        def _eur_val(raw):
            """Parse a numeric string to '1234.56 EUR', rounding float precision."""
            raw = raw.replace(" ", "").replace("\xa0", "")
            if re.match(r'^\d+[.,]\d+$', raw) or re.match(r'^\d+$', raw):
                try:
                    return f"{round(float(raw.replace(',', '.')), 2):.2f} EUR"
                except ValueError:
                    pass
            return None

        # Monetary values: scan row right-to-left to find total then price
        # (avoids hardcoded index offsets for different proforma/invoice layouts)
        _nums = []
        for ci in range(len(row) - 1, -1, -1):
            v = str(row[ci] or "").replace('\xa0', ' ').strip()
            if v and re.match(r'^\d[\d.,]*$', v):
                try:
                    fv = float(v.replace(',', '.'))
                    if fv > 0:
                        _nums.append(v)
                except ValueError:
                    pass

        if len(_nums) >= 2:
            rec.total_price = _eur_val(_nums[0])   # rightmost  = total
            rec.price       = _eur_val(_nums[1])   # next right = unit price
        elif len(_nums) == 1:
            rec.total_price = _eur_val(_nums[0])

        records.append(rec)

    return records


_MAXTON_SHIPPING_RE = re.compile(
    r'(?<!\w)(?:shipping|wysyłka|wysylka|wyslka|freight)(?!\w)', re.IGNORECASE
)


def _split_merged_maxton_lines(lines: list[str]) -> list[str]:
    """Split lines where PDF extraction merged two product rows together.

    Splits at:
    - Subsequent Maxton codes when the preceding segment already has qty data
    - Shipping/freight keywords that appear after qty data
    """
    out: list[str] = []
    for line in lines:
        # Collect candidate split positions: embedded Maxton codes + shipping keywords
        split_points = [0]

        code_matches = list(_MAXTON_CODE_TEXT_RE.finditer(line))
        for match in code_matches[1:]:
            segment_before = line[split_points[-1]:match.start()]
            if re.search(r'\b\d{1,4}\s*(?:szt|kpl)', segment_before, re.IGNORECASE):
                split_points.append(match.start())

        ship_match = _MAXTON_SHIPPING_RE.search(line)
        if ship_match:
            segment_before = line[split_points[-1]:ship_match.start()]
            if re.search(r'\b\d{1,4}\s*(?:szt|kpl)', segment_before, re.IGNORECASE):
                split_points.append(ship_match.start())

        split_points.append(len(line))
        split_points = sorted(set(split_points))

        if len(split_points) == 2:
            out.append(line)
        else:
            for k in range(len(split_points) - 1):
                seg = line[split_points[k]:split_points[k + 1]].strip()
                if seg:
                    out.append(seg)
    return out


def _parse_maxton_from_text(text: str) -> list[ProductRecord]:
    """Text-based fallback when pdfplumber finds no usable tables."""
    records: list[ProductRecord] = []
    seen: dict[str, int] = {}  # kept for shipping-row dedup only (not product merging)
    lines = _split_merged_maxton_lines(text.splitlines())
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _MAXTON_CODE_TEXT_RE.search(line)
        if not m:
            i += 1
            continue

        # Only treat as a real product line when the code appears at the very
        # start (or after a row-number prefix like "4.").  Codes that appear
        # mid-sentence in a description (e.g. "GT-LINE" in
        # "RENAULT MEGANE 4 GT-LINE") must be skipped.
        line_prefix = line[:m.start()].strip()
        if line_prefix and not re.match(r'^\d{1,2}\.?\s*$', line_prefix):
            i += 1
            continue

        code = re.sub(r'_+', '-', m.group(1))
        # Real Maxton codes always contain at least one digit (e.g. ME-S-222-CAP2G).
        # Brand names like "MERCEDES-BENZ" match the pattern but have no digits — skip them.
        if not re.search(r'\d', code):
            i += 1
            continue

        # Description: text after the code; strip leading ';' (new invoice format
        # puts code and name in the same cell separated by ';').
        after = line[m.end():].strip().lstrip(';').strip()

        # Collect up to 5 non-product lines after the code line; skip blank lines.
        # A line is a "new product" only when its code appears at the start —
        # codes embedded mid-sentence in a description are consumed as extra_lines
        # so they don't pollute `seen` and don't cause the real product to be skipped.
        extra_lines: list[str] = []
        j = i + 1
        while j < len(lines) and j <= i + 5:
            nl = lines[j].strip()
            if not nl:
                j += 1  # blank line — keep scanning, don't break
                continue
            # Stop at shipping/freight rows so their prices don't bleed into
            # the preceding product's price extraction.
            if re.search(r'\b(?:shipping|wysyłka|wyslka|freight)\b', nl, re.IGNORECASE):
                break
            mm = _MAXTON_CODE_TEXT_RE.search(nl)
            if mm:
                nl_prefix = nl[:mm.start()].strip()
                if not nl_prefix or re.match(r'^\d{1,2}\.?\s*$', nl_prefix):
                    break  # real product line — stop scanning
            extra_lines.append(nl)
            j += 1

        # Build description from the 'after' part; stop before any data line
        name_parts = [after] if after else []
        for nl in extra_lines:
            if re.search(r'\b\d+\s*(?:szt|kpl)\.?\b', nl, re.IGNORECASE):
                break
            if re.search(r'\b\d{1,6}[.,]\d{2}\b', nl):
                break
            name_parts.append(nl)

        name = " ".join(name_parts).strip()
        # Strip trailing quantity/price block (old single-line format)
        name = re.sub(r'\s+\d+\s+(?:szt|kpl)\.?\b.*$', '', name, flags=re.IGNORECASE).strip()
        # Strip trailing customs/PKWiU code (e.g. " 29.32") from new format
        name = re.sub(r'\s+\d{2,3}\.\d{2}\s*$', '', name).strip()

        # Skip rows where description is a column header
        if re.search(r'\b(?:nazwa|ilosc|quantity|netto|brutto|lp\.?)\b', name, re.IGNORECASE):
            i += 1
            continue

        # Search current line + continuation lines for numeric data
        search_text = " ".join([line] + extra_lines)

        # Quantity: find "N szt" or "N kpl" (kpl. with trailing dot also accepted)
        qty_m = re.search(r'\b(\d{1,4})\s*(?:szt|kpl)\.?\b', search_text, re.IGNORECASE)
        if qty_m:
            n = int(qty_m.group(1))
            quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
            # Prices appear AFTER the unit — PKWiU customs code (e.g. 29.32) is before
            after_unit = search_text[qty_m.end():]
            prices = re.findall(r'\b(\d{1,6}[.,]\d{2})\b', after_unit)
            # Anchor total on first price after the VAT% marker (e.g. "0%") so that
            # stray trailing numbers (discounts, dates) don't replace the real total.
            vat_m = re.search(r'\b\d+\s*%', after_unit)
            after_vat_prices = (
                re.findall(r'\b(\d{1,6}[.,]\d{2})\b', after_unit[vat_m.end():])
                if vat_m else []
            )
        else:
            quantity = None
            prices = re.findall(r'\b(\d{1,6}[.,]\d{2})\b', search_text)
            after_vat_prices = []

        if len(prices) >= 2:
            price       = prices[0].replace(',', '.') + ' EUR'
            total_price = (after_vat_prices[0] if after_vat_prices else prices[-1]).replace(',', '.') + ' EUR'
        elif len(prices) == 1:
            price       = prices[0].replace(',', '.') + ' EUR'
            total_price = None
        else:
            price       = None
            total_price = None

        logger.info("Maxton code=%s qty=%s price=%s total=%s | search: %s",
                    code, quantity, price, total_price, search_text[:160])

        if code in seen:
            existing = records[seen[code]]
            existing.merged_count += 1
            if quantity and existing.quantity:
                try:
                    en = int(re.search(r'\d+', existing.quantity).group())
                    nn = int(re.search(r'\d+', quantity).group())
                    mn = en + nn
                    existing.quantity = f"{mn} {'Брой' if mn == 1 else 'Броя'}"
                except Exception:
                    pass
            if total_price and existing.total_price:
                try:
                    ep = float(existing.total_price.replace(' EUR', '').replace(',', '.'))
                    np = float(total_price.replace(' EUR', '').replace(',', '.'))
                    existing.total_price = f"{ep + np:.2f} EUR"
                except Exception:
                    pass
        else:
            rec = ProductRecord(extraction_method="table")
            rec.product_code = code
            rec.product_name = name[:120] if name else None
            rec.quantity     = quantity
            rec.price        = price
            rec.total_price  = total_price
            seen[code]       = len(records)
            records.append(rec)

        i = j  # skip past already-consumed extra lines

    # Handle shipping/freight row — no standard XX-XXXX code.
    # In newer invoice formats the price appears on the NEXT line(s), not on the
    # same line as "Wysyłka / Shipping", so we also check 1-2 adjacent lines.
    for idx, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not re.search(r'\b(?:shipping|wysyłka|wyslka|freight)\b', line, re.IGNORECASE):
            continue
        # Collect the keyword line plus up to 2 following lines
        search_lines = [line]
        for offset in (1, 2):
            if idx + offset < len(lines):
                search_lines.append(lines[idx + offset].strip())
        combined = " ".join(search_lines)
        # Lenient price regex — no word boundaries to handle merged digits
        prices = list(dict.fromkeys(re.findall(r'(\d{1,6}[.,]\d{2})', combined)))
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
    records: list[ProductRecord] = []
    seen_codes: dict[str, int] = {}  # code → index in records (for merging)

    def _merge_or_add(rec: ProductRecord):
        code = rec.product_code
        if code in seen_codes:
            existing = records[seen_codes[code]]
            existing.merged_count += 1
            # Sum quantities
            if rec.quantity and existing.quantity:
                try:
                    en = int(re.search(r'\d+', existing.quantity).group())
                    nn = int(re.search(r'\d+', rec.quantity).group())
                    mn = en + nn
                    existing.quantity = f"{mn} {'Брой' if mn == 1 else 'Броя'}"
                except Exception:
                    pass
            # Sum totals
            if rec.total_price and existing.total_price:
                try:
                    ep = float(existing.total_price.replace(' EUR', '').replace(',', '.'))
                    np = float(rec.total_price.replace(' EUR', '').replace(',', '.'))
                    existing.total_price = f"{ep + np:.2f} EUR"
                except Exception:
                    pass
        else:
            seen_codes[code] = len(records)
            records.append(rec)

    for table in tables:
        for rec in _parse_maxton_table(table):
            _merge_or_add(rec)

    if not records and text:
        logger.info("Maxton: no table records — trying text extraction")
        records = _parse_maxton_from_text(text)

    logger.info("Maxton Design extraction: %d records total", len(records))
    return records


# ---------------------------------------------------------------------------
# M-Tech Poland-specific extractor
# ---------------------------------------------------------------------------

_MTECH_CODE_RE = re.compile(
    r'^([A-Z][A-Z0-9\-/]{1,20})\s+'   # product code
    r'\d{6,10}\s+'                      # CN code
    r'[A-Z]{2}\s+'                      # country of origin
    r'(\d{8,14})\s+'                    # EAN
    r'([\d,]+)\s*$',                    # weight
    re.IGNORECASE,
)
_MTECH_DESC_RE = re.compile(
    r'^(.+?)\s+'
    r'(\d+(?:[,.]\d+)?)\s+'
    r'(szt\.?|kpl\.?|pcs\.?|set|pce)\s+'
    r'\d+\s*%\s+'
    r'([\d,]+)\s+'
    r'([\d,]+)\s+EUR\s*$',
    re.IGNORECASE,
)


def _parse_mtech_from_text(text: str) -> list[ProductRecord]:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    records = []
    i = 0
    while i < len(lines):
        m_code = _MTECH_CODE_RE.match(lines[i])
        if m_code and i + 2 < len(lines):
            code   = m_code.group(1).upper()
            ean    = m_code.group(2)
            weight = m_code.group(3).replace(',', '.')
            # lines[i+1] = sequential number (skip); lines[i+2] = description
            m_desc = _MTECH_DESC_RE.match(lines[i + 2])
            if m_desc:
                rec = ProductRecord(extraction_method="regex")
                rec.product_code  = code
                rec.ean           = ean
                rec.weight_kg     = weight
                rec.product_name  = m_desc.group(1).strip()
                qty_num = m_desc.group(2).replace(',', '.')
                unit    = m_desc.group(3)
                try:
                    q = float(qty_num)
                    rec.quantity = f"{int(q) if q == int(q) else q} {unit}"
                except ValueError:
                    rec.quantity = f"{qty_num} {unit}"
                rec.price       = m_desc.group(4).replace(',', '.') + " EUR"
                rec.total_price = m_desc.group(5).replace(',', '.') + " EUR"
                records.append(rec)
                i += 3
                continue
        i += 1
    return records


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
            for ri, rw in enumerate(table[:4]):
                logger.warning("M-Tech:   row[%d]: %s", ri, [str(c or "")[:30] for c in rw])
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

        # Prefer explicit header positions (new invoice format has named columns).
        # Fall back to calibration from first product row for older formats.
        ean_pos    = find_col(["ean"])
        weight_pos = find_col(["weight"])
        qty_pos    = find_col(["qty", "quantity"])
        price_pos  = find_col(["subtotal price", "unit price", "price"])
        total_pos  = find_col(["subtotal\nvalue", "subtotal value", "net value", "total value", "total"])

        data_rows = table[header_idx + 1:]
        for dr in data_rows:
            code = cell(dr, desc_idx)
            if not (code and re.match(r'^[A-Z][A-Z0-9\-/]{1,15}$', code)):
                continue
            # code_row found — desc_row is 2 rows later
            if data_rows.index(dr) + 2 >= len(data_rows):
                break
            desc_row_sample = data_rows[data_rows.index(dr) + 2]
            # Calibrate only positions not already found in the header
            for ci, val in enumerate(dr):
                v = cell(dr, ci)
                if ean_pos is None and re.match(r'^\d{8,14}$', v):
                    ean_pos = ci
                elif weight_pos is None and re.match(r'^\d+\.\d+$', v) and 0.05 < float(v) < 100:
                    weight_pos = ci
            for ci, val in enumerate(desc_row_sample):
                v = cell(desc_row_sample, ci)
                if qty_pos is None and ci == ean_pos and re.match(r'^\d+$', v):
                    qty_pos = ci
                elif price_pos is None and re.match(r'^\d+[.,]\d{2}$', v) and ci > (ean_pos or 0):
                    price_pos = ci
                elif total_pos is None and re.search(r'\d+[.,]\d{2}.*EUR', v) and ci > (price_pos or 0):
                    total_pos = ci
            logger.info("M-Tech calibrated: ean=%s weight=%s qty=%s price=%s total=%s",
                        ean_pos, weight_pos, qty_pos, price_pos, total_pos)
            break

        i = 0
        while i < len(data_rows):
            row1 = data_rows[i]
            code = cell(row1, desc_idx)

            # Product code row: col 1 has a short uppercase code (e.g. CP14W, CP5S)
            # The No. number appears on the NEXT row, description on the row after that.
            if not (code and re.match(r'^[A-Z][A-Z0-9\-/ ]{1,20}$', code)):
                if code:
                    logger.warning("M-Tech: skipped row — code %r doesn't match | row snippet: %s",
                                   code, [str(c or "")[:20] for c in row1[:8]])
                i += 1
                continue

            ean    = cell(row1, ean_pos) if ean_pos is not None else ""
            weight = cell(row1, weight_pos) if weight_pos is not None else ""

            # Structure per product: [code_row] [number_row] [description_row]
            desc_row  = data_rows[i + 2] if i + 2 < len(data_rows) else []
            desc      = cell(desc_row, desc_idx)

            def _from_either(pos):
                """Return value from row1 or desc_row — whichever is non-empty."""
                if pos is None:
                    return ""
                v1 = cell(row1, pos)
                v2 = cell(desc_row, pos) if desc_row else ""
                return v1 if v1 else v2

            qty_raw   = _from_either(qty_pos)
            price_raw = _from_either(price_pos)
            total_raw = _from_either(total_pos)

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

            # Derive unit price from total / qty when price column is absent
            if not rec.price and rec.total_price and rec.quantity:
                try:
                    _qm = re.match(r'(\d+)', str(rec.quantity))
                    _n = int(_qm.group(1)) if _qm else 0
                    if _n > 0:
                        _t = float(re.search(r'[\d.]+', rec.total_price).group())
                        rec.price = f"{round(_t / _n, 2):.2f} EUR"
                except Exception:
                    pass

            if weight:
                rec.weight_kg = weight

            records.append(rec)
            i += 3  # code_row + number_row + desc_row

    if not records and text:
        logger.info("M-Tech: table extraction failed, trying text parser")
        records = _parse_mtech_from_text(text)
        if records:
            logger.info("M-Tech text parser: %d records", len(records))
    logger.info("M-Tech extraction: %d records", len(records))
    return records


# ---------------------------------------------------------------------------
# Ma*Fra (Aviатранс / pochisti.bg) extractor
# ---------------------------------------------------------------------------
#
# Invoice structure (image/scan → OCR text):
#   No | Код | Наименование | Кол-во | Марка | Ед. цена | Вал. | ТО | Общо в НВ
#
# Product codes: 1-2 capital letters + 2-5 digits  (A0521, H0050, MF66, P0494)
# Unit: always "бр" — translate to Брой/Броя
# Currency: EUR

_MAFRA_CODE_RE = re.compile(r'^([A-Z]{1,2}\d{2,5})$')

# Cyrillic look-alike → Latin substitution table for OCR-damaged product codes
_OCR_FIX = str.maketrans({
    'А': 'A', 'В': 'B', 'С': 'C', 'Е': 'E', 'Н': 'H',
    'К': 'K', 'М': 'M', 'О': '0', 'о': '0', 'Р': 'P',
    'Т': 'T', 'Х': 'X', 'Ф': 'F', '#': 'H',
})

# Applied only to the digit portion of a code (after the 1-2 letter prefix)
_DIGIT_FIX = str.maketrans({'O': '0', 'I': '1', 'L': '1', 'V': '0', 'Z': '2', 'S': '5', 'G': '6', 'B': '8', '+': '4'})


def _fix_mafra_code(s: str) -> str:
    s = s.upper().translate(_OCR_FIX)
    # Apply digit-lookalike fixes from position 1 onward.
    # Single-letter prefix (A, H, P, M...) is the common case; applying digit
    # fixes from char 1 handles "AO193" → "A0193" without breaking "MF66"
    # because 'F' is not in _DIGIT_FIX.
    if len(s) >= 2:
        s = s[0] + s[1:].translate(_DIGIT_FIX)
    return s


def _is_mafra_code(token: str) -> bool:
    return bool(re.match(r'^[A-Z]{1,2}\d{2,5}$', _fix_mafra_code(token)))


def _is_mafra_document(text: str) -> bool:
    return bool(re.search(r'pochisti|авиатранс|mafra', text, re.IGNORECASE))


def _parse_mafra_table(table: list[list]) -> list[ProductRecord]:
    """Parse a Ma*Fra table (from PDF or grid-based OCR cell extraction).

    Header keywords are matched loosely to survive OCR garbling:
    "Коло" matches "Код", "Няименование" contains "аим", etc.
    For code cells, _fix_mafra_code() is applied to handle OCR errors.
    """
    if not table or len(table) < 2:
        return []

    # Find header row — accept OCR variants of the Bulgarian keywords
    # "коло" = garbled "Код",  "яим"/"аим" = inside garbled "Наименование"
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        has_code = any(kw in joined for kw in ("код", "коло", "code", "кол\"", "no."))
        has_desc = any(kw in joined for kw in ("наим", "яим", "аим", "описание",
                                                "description", "стока", "наименован"))
        if has_code and has_desc:
            header_idx = i
            break
    if header_idx is None:
        # Fallback: row with most non-empty cells that contains some table keywords
        for i, row in enumerate(table):
            joined = " ".join(str(c or "").lower() for c in row)
            score = sum(kw in joined for kw in ("кол", "цена", "общо", "мярка", "вал"))
            if score >= 2:
                header_idx = i
                break
    if header_idx is None:
        logger.info("Ma*Fra table: header row not found")
        return []

    headers = [str(c or "").lower().strip().replace('\n', '') for c in table[header_idx]]
    logger.info("Ma*Fra table: header at row %d: %s", header_idx, headers)

    def find(kws):
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return None

    # "коло" = OCR garbling of "Код"
    code_idx  = find(["код", "коло", "code"])
    desc_idx  = find(["наим", "яим", "аим", "описание", "description", "стока"])
    qty_idx   = find(["кол-во", "кол во", "кол-", "qty", "количество"])
    price_idx = find(["ед. цена", "ед.цена", "ед цена", "unit price", "цена"])
    total_idx = find(["общо", "обшо", "total", "нв", "o6mo", "обмо"])
    val_idx   = find(["вал.", "вал", "валута", "currency"])

    if code_idx is None:
        logger.info("Ma*Fra table: code column not found in headers")
        return []

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(idx):
            if idx is None or idx >= len(row):
                return ""
            return str(row[idx] or "").strip()

        raw_code = cell(code_idx)
        if not raw_code:
            continue
        # Normalize: strip newlines (code may wrap across lines in narrow column).
        # Do NOT apply _fix_mafra_code here — OCR digit-substitutions (V→0, G→6, B→8)
        # destroy valid letter codes like AVMFGLOVEGREEN08 in text-based PDFs.
        code = raw_code.replace('\n', '').replace('\r', '').strip().upper()
        # Accept any alphanumeric code ≥ 3 chars (table column already identifies it).
        if not re.match(r'^[A-Z0-9]{3,}$', code):
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code

        if desc_idx is not None:
            desc_val = cell(desc_idx)[:150]
            rec.product_name = desc_val or None

        if qty_idx is not None:
            qty_raw = cell(qty_idx)
            try:
                n = int(float(qty_raw.replace(",", ".")))
                rec.quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
            except ValueError:
                rec.quantity = qty_raw or None

        currency = cell(val_idx).upper() if val_idx is not None else "EUR"
        if not currency:
            currency = "EUR"

        if price_idx is not None:
            price_raw = cell(price_idx)
            p = _clean_num(price_raw)
            if p:
                rec.price = p + " " + currency

        if total_idx is not None:
            total_raw = cell(total_idx)
            t = _clean_num(total_raw)
            if t:
                rec.total_price = t + " " + currency

        records.append(rec)

    return records


def _parse_mafra_from_text(text: str) -> list[ProductRecord]:
    """Parse Ma*Fra invoice from OCR text (image/scanned PDF).

    Tesseract output for scanned invoices is noisy: table borders become |,
    Cyrillic chars replace Latin ones, # is misread as H, etc.
    Strategy: clean each line, look for a product code anywhere in the first
    5 tokens (after OCR correction), then extract total from the last decimal.
    """
    records = []
    lines = text.splitlines()

    logger.info("Ma*Fra OCR text (%d lines):\n%s", len(lines),
                "\n".join(f"  {i:3d}: {l}" for i, l in enumerate(lines[:60])))

    seen_codes: set[str] = set()

    for raw in lines:
        # Remove table border characters that OCR picks up
        line = re.sub(r'[|\[\](){}]', ' ', raw)
        line = re.sub(r'\s+', ' ', line).strip()
        if len(line) < 4:
            continue

        tokens = line.split()

        # Pass 1: find a Ma*Fra code with letter prefix (with OCR correction)
        code = None
        code_pos = -1
        for ti, tok in enumerate(tokens[:10]):
            fc = _fix_mafra_code(tok)
            if _MAFRA_CODE_RE.match(fc):
                code = fc
                code_pos = ti
                break

        # Pass 2: pure-digit token (3-5 digits) → try H prefix (most common in Ma*Fra)
        # Require the token to consist ONLY of digits to avoid false positives
        # like "70x90cm" → "7090".
        if not code:
            for ti, tok in enumerate(tokens[:10]):
                if re.match(r'^\d{3,5}$', tok):
                    candidate = 'H' + tok
                    if _MAFRA_CODE_RE.match(candidate):
                        code = candidate
                        code_pos = ti
                        break

        # Pass 3: row-number merged with code, e.g. "140378" → last 4 digits "0378" → H0378
        if not code:
            for ti, tok in enumerate(tokens[:4]):
                if re.match(r'^\d{5,7}$', tok):
                    suffix = tok[-4:]
                    candidate = 'H' + suffix
                    if _MAFRA_CODE_RE.match(candidate):
                        code = candidate
                        code_pos = ti
                        break

        # Pass 4: strip non-alphanumeric noise characters from token and retry
        # Catches cases like "P0+94" ('+' misread from '4'), "A.0521", "H|0050"
        if not code:
            for ti, tok in enumerate(tokens[:10]):
                cleaned = re.sub(r'[^A-Za-z0-9АаВвСсЕеНнКкМмОоРрТтХхФф#]', '', tok)
                if cleaned and cleaned != tok and len(cleaned) >= 3:
                    fc = _fix_mafra_code(cleaned)
                    if _MAFRA_CODE_RE.match(fc):
                        code = fc
                        code_pos = ti
                        break

        if not code or code in seen_codes:
            continue

        # Extract all X.XX or X,XX numbers on the line
        # Allow leading non-digit chars (OCR sometimes adds '(' before prices)
        decimals = re.findall(r'(?<!\d)\d{1,6}[.,]\d{2}(?!\d)', line)
        if not decimals:
            continue

        seen_codes.add(code)

        total_raw = decimals[-1]
        # Unit price = last decimal before EUR keyword, or second-to-last
        price_raw = None
        eur_m = re.search(r'(\d+[.,]\d{2})\s*(?:EUR|BUR|eur)', line, re.IGNORECASE)
        if eur_m:
            price_raw = eur_m.group(1)
        elif len(decimals) >= 2:
            price_raw = decimals[-2]

        # Description: word tokens after code, before first decimal
        first_dec_pos = line.find(decimals[0]) if decimals else len(line)
        code_end = line.upper().find(code)
        if code_end >= 0:
            code_end += len(code)
        desc_raw = line[code_end:first_dec_pos] if code_end >= 0 else ""
        # Strip non-word noise and keep only alphabetic tokens
        desc_parts = [t for t in desc_raw.split()
                      if re.match(r'^[A-Za-zА-Яа-яЁё]', t) and len(t) > 1
                      and t.upper() not in ('EUR', 'BUR', 'ML', 'PZ')]
        desc = " ".join(desc_parts[:12]).strip()[:120] or None

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = desc

        t = _clean_num(total_raw)
        if t:
            rec.total_price = t + " EUR"

        if price_raw and price_raw != total_raw:
            p = _clean_num(price_raw)
            if p:
                rec.price = p + " EUR"

        # Quantity: a standalone integer after the code position (e.g. "H0050 ... 12 ...")
        # Look for the first 1-3 digit standalone integer after code_pos
        qty_tokens = tokens[code_pos + 1:code_pos + 8] if code_pos >= 0 else tokens[1:8]
        for qt in qty_tokens:
            qt_clean = re.sub(r'[^\d]', '', qt)
            if qt_clean and re.match(r'^\d{1,3}$', qt_clean) and not re.search(r'[.,]\d{2}', qt):
                n = int(qt_clean)
                if 1 <= n <= 9999:
                    rec.quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
                    break

        records.append(rec)

    return records


def extract_mafra_products(tables: list, text: str = "", text2: str = "") -> list[ProductRecord]:
    logger.info("Ma*Fra: %d table(s) received", len(tables))
    for ti, table in enumerate(tables):
        logger.info("Ma*Fra table[%d] (%d rows): first row=%r", ti, len(table),
                    [str(c or "")[:40] for c in (table[0] if table else [])])

    # Try table extraction first (real PDF tables)
    records = []
    for table in tables:
        records.extend(_parse_mafra_table(table))

    if records:
        logger.info("Ma*Fra: %d records from table(s)", len(records))
        return records

    # Fallback: parse OCR text — merge results from both preprocessing passes
    logger.info("Ma*Fra: no table records — trying text extraction")
    records = _parse_mafra_from_text(text)

    if text2:
        records2 = _parse_mafra_from_text(text2)
        seen = {r.product_code for r in records}
        added = 0
        for rec in records2:
            if rec.product_code not in seen:
                records.append(rec)
                seen.add(rec.product_code)
                added += 1
        if added:
            logger.info("Ma*Fra dual-pass: added %d new records from pass B", added)

    logger.info("Ma*Fra text extraction: %d records total", len(records))
    return records


# ---------------------------------------------------------------------------
# Car Passion specific extractor
# ---------------------------------------------------------------------------
#
# Invoice layout (Polish):
#   Lp. | Kod towaru/usługi | Nazwa towaru/usługi | Ilość | J.m. | VAT |
#   Cena netto EUR | Wartość netto EUR
#
# Product codes: mostly numeric (4-7 digits) or alphanumeric e.g. "PURE S",
# "10017 M-L", "10017 S-M".  Unit is always "szt".

def _is_car_passion_document(text: str) -> bool:
    return bool(re.search(r'car.?passion', text, re.IGNORECASE))


def _parse_car_passion_table(table: list[list]) -> list[ProductRecord]:
    if not table or len(table) < 2:
        return []

    def norm(s):
        return str(s or "").lower().strip()

    def cell(row, idx):
        if idx is None or idx < 0 or idx >= len(row):
            return ""
        val = str(row[idx] or "").strip()
        for line in val.split("\n"):
            if line.strip():
                return line.strip()
        return val

    # Find header row: must have "kod" (code) AND "nazwa" (name)
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(norm(c) for c in row)
        if ("kod" in joined or "code" in joined) and ("nazwa" in joined or "name" in joined):
            header_idx = i
            break
    if header_idx is None:
        logger.info("Car Passion: no header row found in %d-row table", len(table))
        return []

    headers = [norm(c) for c in table[header_idx]]
    logger.info("Car Passion: header at row %d: %s", header_idx, headers)

    def find(kws):
        for kw in kws:
            for i, h in enumerate(headers):
                if kw in h:
                    return i
        return None

    code_idx  = find(["kod towaru", "kod"])
    desc_idx  = find(["nazwa towaru", "nazwa", "name"])
    qty_idx   = find(["ilość", "ilosc", "qty", "quantity"])
    unit_idx  = find(["j.m", "jm", "unit", "jednostk"])
    price_idx = find(["cena netto", "net price", "cena"])
    total_idx = find(["wartość netto", "wartosc netto", "wartość", "wartosc", "net value"])

    if code_idx is None:
        logger.info("Car Passion: code column not found in %s", headers)
        return []

    records = []
    _SKIP = {"lp.", "lp", "kod", "code", "nan", ""}
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        code = cell(row, code_idx)
        if not code or code.lower() in _SKIP:
            continue
        # Skip pure row-number cells (1, 2, 3 ...)
        if code.isdigit() and len(code) <= 2:
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code

        if desc_idx is not None:
            desc = cell(row, desc_idx)
            if desc and len(desc) > 2:
                rec.product_name = desc[:120]

        unit_raw = cell(row, unit_idx) if unit_idx is not None else ""
        qty_raw  = cell(row, qty_idx)  if qty_idx  is not None else ""
        if qty_raw:
            try:
                n = int(float(qty_raw.replace(",", ".")))
                qty_str = str(n)
            except (ValueError, TypeError):
                qty_str = qty_raw
            if unit_raw:
                rec.quantity = qty_str + " " + unit_raw.upper()
            else:
                rec.quantity = qty_str

        price_raw = cell(row, price_idx) if price_idx is not None else ""
        if price_raw:
            p = re.sub(r'[^\d.,]', '', price_raw).replace(',', '.')
            if p:
                rec.price = p + " EUR"

        total_raw = cell(row, total_idx) if total_idx is not None else ""
        if total_raw:
            t = re.sub(r'[^\d.,]', '', total_raw).replace(',', '.')
            if t:
                rec.total_price = t + " EUR"

        records.append(rec)

    logger.info("Car Passion: %d records extracted", len(records))
    return records


_CP_DEC_RE  = re.compile(r'\b(\d{1,6}[.,]\d{2})\b')
_CP_SZT_RE  = re.compile(r'\bszt\.?\b', re.IGNORECASE)
_CP_LP_RE   = re.compile(r'^\d{1,2}\s+')
# Alphanumeric suffix patterns: M-L, S-M, XL, XXL, PURE S ...
_CP_SFXRE   = re.compile(r'^([A-Z]{1,3}-[A-Z]{1,3}|[A-Z]{2,3}|S)$')
_CP_SKIP    = frozenset(["lp.", "lp", "kod", "nazwa", "ilość", "j.m.", "j.m", "vat",
                          "cena", "wartość", "wartosc", "towaru", "usługi"])


def _parse_car_passion_from_text(text: str) -> list[ProductRecord]:
    """Text-based fallback for Car Passion invoices.

    pdfplumber often fails to detect the product table (returns only a 3-row
    header table).  This parser scans raw PDF text for lines that contain the
    unit word 'szt' and at least two decimal prices, then reconstructs:
      code, description, quantity, unit-price, total-price

    Expected line layout (after row-number stripping):
      {code} [{suffix}] {description words}  {qty}  szt  {VAT%}  {price}  {total}
    """
    records = []
    seen_codes: dict[str, list[int]] = {}
    lines = text.splitlines()

    # Log first 30 lines for debugging
    logger.info("Car Passion text fallback – first 30 lines:\n%s",
                "\n".join(f"  {i:3d}: {l}" for i, l in enumerate(lines[:30])))

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1

        # Must contain unit word "szt"
        if not _CP_SZT_RE.search(line):
            continue

        # Must have at least 2 decimal prices after szt
        szt_m = _CP_SZT_RE.search(line)
        post_szt = line[szt_m.end():]
        decimals = _CP_DEC_RE.findall(post_szt)
        if len(decimals) < 1:
            continue

        # Strip optional leading row number (1-2 digits)
        cleaned = _CP_LP_RE.sub('', line).strip()
        tokens = cleaned.split()
        if not tokens:
            continue

        # Find position of "szt" in token list
        szt_idx = next((j for j, t in enumerate(tokens) if _CP_SZT_RE.fullmatch(t)), None)
        if szt_idx is None or szt_idx < 2:
            continue

        # Product code = first token; check for two-word codes like "PURE S", "10017 M-L"
        code = tokens[0]
        if code.lower() in _CP_SKIP:
            continue
        # Skip pure single/double-digit row-number tokens
        if code.isdigit() and len(code) <= 2:
            continue

        code_end = 1
        if szt_idx >= 3 and _CP_SFXRE.match(tokens[1]):
            code = tokens[0] + " " + tokens[1]
            code_end = 2

        # Quantity = token immediately before szt
        qty_str = tokens[szt_idx - 1]
        try:
            qty_val = int(float(qty_str.replace(',', '.')))
        except (ValueError, TypeError):
            qty_val = None

        # Description = tokens between code and qty
        desc_tokens = tokens[code_end: szt_idx - 1]
        desc = " ".join(desc_tokens).strip()

        # If description is missing or too short, look at the PREVIOUS line
        if len(desc) < 4 and i >= 2:
            prev = lines[i - 2].strip()
            if prev and not _CP_SZT_RE.search(prev):
                desc = (prev + " " + desc).strip()

        rec = ProductRecord(extraction_method="text")
        rec.product_code = code
        if desc and len(desc) > 2:
            rec.product_name = desc[:120]
        if qty_val is not None:
            rec.quantity = str(qty_val) + " SZT"

        if decimals:
            p = _clean_num(decimals[0])
            if p:
                rec.price = p + " EUR"
        if len(decimals) >= 2:
            t = _clean_num(decimals[-1])
            if t:
                rec.total_price = t + " EUR"

        _smart_merge_or_add(records, seen_codes, rec)
        logger.info("Car Passion text: code=%s qty=%s price=%s total=%s",
                    code, rec.quantity, rec.price, rec.total_price)

    return records


def extract_car_passion_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Car Passion: %d table(s) received", len(tables))
    records = []
    seen_codes: dict[str, list[int]] = {}
    for table in tables:
        for rec in _parse_car_passion_table(table):
            _smart_merge_or_add(records, seen_codes, rec)

    if not records and text:
        logger.info("Car Passion: table extraction yielded 0 records, trying text fallback")
        records = _parse_car_passion_from_text(text)
        logger.info("Car Passion text fallback: %d records", len(records))
    else:
        logger.info("Car Passion extraction: %d records total", len(records))
    return records


# ---------------------------------------------------------------------------
# Vinove-specific extractor
# ---------------------------------------------------------------------------
#
# Actual header (bilingual PL/EN):
#   No. | Items (Nazwa towaru/uslugi) | Quantity (Ilosc) | Unit (Jm) |
#   Price per unit (Cena) | Amount net (Wart. netto) | VAT% | VAT | Amount (Wart. brutto)
#
# Product code is the first whitespace-separated token of the Items cell.
# Codes follow the pattern: 1-3 letters + 0-4 digits + dash + 2+ digits  (e.g. V22-07)

def _is_vinove_document(text: str) -> bool:
    return bool(re.search(r'vinove', text, re.IGNORECASE))


# Cell-level: code is the FIRST token of description cell  (V22-07 Description...)
_VINOVE_CODE_RE = re.compile(
    r'^([A-Z]{1,3}\d{0,4}-[A-Z0-9]{2,}(?:-[A-Z0-9]+)*)\s+(.*)', re.DOTALL)
# Scan whole text line for a Vinove code
_VINOVE_CODE_TEXT_RE = re.compile(r'(?<!\w)([A-Z]{1,3}\d{1,4}-\d{2,}(?:-\d+)*)')
# Quick cell match (startswith) used for continuation-table column scan
_VINOVE_CODE_CELL_RE = re.compile(r'^[A-Z]{1,3}\d{0,4}-\d{2,}')


def _parse_vinove_table(table: list[list]) -> list[ProductRecord]:
    if not table or len(table) < 2:
        return []

    def norm(s):
        return _strip_diacritics(str(s or "").lower())

    def cell(row, idx):
        if idx is None or idx < 0 or idx >= len(row):
            return ""
        return str(row[idx] or "").strip()

    # ── Phase 1: locate header row ────────────────────────────────────────────
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(norm(c) for c in row)
        first  = norm(row[0]) if row else ""
        if first.strip().rstrip('.') == "lp" and "netto" in joined:
            header_idx = i
            break
        if "nazwa" in joined and ("ilosc" in joined or "qty" in joined or "netto" in joined):
            header_idx = i
            break
        # Bilingual headers: "items" + "quantity" / "ilosc"
        if "items" in joined and ("quantity" in joined or "ilosc" in joined):
            header_idx = i
            break

    if header_idx is not None:
        headers = [norm(c) for c in table[header_idx]]
        logger.info("Vinove: header row at index %d: %s", header_idx, headers)

        def find(kws):
            for kw in kws:
                for i, h in enumerate(headers):
                    if kw in h:
                        return i
            return None

        desc_idx  = find(["items", "nazwa towaru", "product name", "nazwa", "product"])
        qty_idx   = find(["ilosc", "qty", "quantity"])
        unit_idx  = find(["unit\n", "jm", "jednostk"])
        # Unit price: "price per unit (cena)" — must be found BEFORE "amount net"
        price_idx = find(["price per unit", "cena netto", "unit price", "cena"])
        # Net total: "amount net (wart. netto)"
        value_idx = find(["amount net", "wart. netto", "netto eur", "wartosc netto",
                          "wartosc", "net value", "value eur"])
        if value_idx is None:
            # Last-resort: last column (may be gross total — not ideal but better than nothing)
            value_idx = len(headers) - 1
        if desc_idx is None:
            desc_idx = 0
        data_start = header_idx + 1

    else:
        # ── Phase 2: continuation table (page 2+) — no header ────────────────
        logger.info("Vinove: no header in %d-row table — scanning for code column", len(table))
        ncols = max((len(r) for r in table if r), default=0)
        if not ncols:
            return []
        counts = [0] * ncols
        for row in table[:min(20, len(table))]:
            for ci in range(min(len(row), ncols)):
                if row[ci] and _VINOVE_CODE_CELL_RE.match(str(row[ci]).strip()):
                    counts[ci] += 1
        best = max(range(ncols), key=lambda i: counts[i])
        if counts[best] == 0:
            logger.info("Vinove: no code pattern in continuation table — skipping")
            return []
        desc_idx  = best
        qty_idx   = best + 1 if best + 1 < ncols else None
        unit_idx  = best + 2 if best + 2 < ncols else None
        price_idx = best + 3 if best + 3 < ncols else None
        value_idx = best + 4 if best + 4 < ncols else None
        logger.info("Vinove: continuation columns — desc=%d qty=%s unit=%s price=%s value=%s",
                    desc_idx, qty_idx, unit_idx, price_idx, value_idx)
        data_start = 0

    records = []
    _SKIP = {"lp.", "lp", "no.", "no", "items", "nazwa", ""}
    for row in table[data_start:]:
        if not any(str(c or "").strip() for c in row):
            continue

        raw_desc = cell(row, desc_idx)
        if not raw_desc:
            continue

        m = _VINOVE_CODE_RE.match(raw_desc)
        if m:
            code = m.group(1)
            name = m.group(2).strip()
        else:
            parts = raw_desc.split(None, 1)
            code  = parts[0]
            name  = parts[1].strip() if len(parts) > 1 else ""

        if not code or code.lower() in _SKIP:
            continue
        if code.isdigit() and len(code) <= 3:
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        if name:
            rec.product_name = name[:120]

        qty_raw  = cell(row, qty_idx)  if qty_idx  is not None else ""
        unit_raw = cell(row, unit_idx) if unit_idx is not None else ""
        if qty_raw:
            try:
                qty_f = float(qty_raw.replace(",", "."))
                n = int(qty_f) if qty_f == int(qty_f) else qty_f
                rec.quantity = str(n) + (" " + unit_raw.upper() if unit_raw else "")
            except ValueError:
                rec.quantity = qty_raw

        price_raw = cell(row, price_idx) if price_idx is not None else ""
        if price_raw and re.match(r'^\d+[.,]\d+$', price_raw):
            rec.price = price_raw.replace(",", ".") + " EUR"

        value_raw = cell(row, value_idx) if value_idx is not None else ""
        if value_raw and re.match(r'^\d+[.,]\d+$', value_raw):
            rec.total_price = value_raw.replace(",", ".") + " EUR"

        records.append(rec)

    logger.info("Vinove table: %d records extracted", len(records))
    return records


def _parse_vinove_from_text(text: str) -> list[ProductRecord]:
    """Text supplement: scan raw PDF text for Vinove codes missed by table parser."""
    records = []
    seen: set[str] = set()
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _VINOVE_CODE_TEXT_RE.search(line)
        if not m:
            i += 1
            continue

        code = m.group(1)
        if code in seen:
            i += 1
            continue

        after = line[m.end():].strip()
        name = re.sub(r'\s+\d+\s+(?:szt|kpl|pcs|unit)\b.*$', '', after,
                      flags=re.IGNORECASE).strip()
        if not name:
            name = after

        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if (next_line
                    and not _VINOVE_CODE_TEXT_RE.search(next_line)
                    and not re.search(r'\b\d+\s+(?:szt|kpl|unit|pcs)\b', next_line,
                                      re.IGNORECASE)):
                name = (name + " " + next_line).strip()

        if re.search(r'\b(?:nazwa|ilosc|quantity|netto|brutto|lp\.?|items)\b',
                     name, re.IGNORECASE):
            i += 1
            continue

        qty_m = re.search(r'\b(\d{1,4})\s*(?:szt|kpl|pcs|unit)\b', line, re.IGNORECASE)
        if qty_m:
            n = int(qty_m.group(1))
            quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
        else:
            quantity = None

        prices = re.findall(r'\b(\d{1,6}[.,]\d{2})\b', line)
        if len(prices) >= 2:
            price       = prices[-2].replace(',', '.') + ' EUR'
            total_price = prices[-1].replace(',', '.') + ' EUR'
        elif len(prices) == 1:
            price       = prices[-1].replace(',', '.') + ' EUR'
            total_price = None
        else:
            price = total_price = None

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name[:120] if name else None
        rec.quantity     = quantity
        rec.price        = price
        rec.total_price  = total_price

        seen.add(code)
        records.append(rec)
        i += 1

    logger.info("Vinove text extraction: %d records", len(records))
    return records


def extract_vinove_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Vinove: %d table(s) received", len(tables))
    records = []
    seen_codes: dict[str, list[int]] = {}
    for table in tables:
        for rec in _parse_vinove_table(table):
            _smart_merge_or_add(records, seen_codes, rec)

    # Always supplement with text to catch continuation-page rows that table
    # missed (e.g. page 2+ tables where pdfplumber finds no header row).
    if text:
        text_records = _parse_vinove_from_text(text)
        added_before = len(records)
        for rec in text_records:
            _smart_merge_or_add(records, seen_codes, rec)
        added = len(records) - added_before
        if added:
            logger.info("Vinove: text supplement added %d records", added)

    logger.info("Vinove extraction: %d records total", len(records))
    return records


# ---------------------------------------------------------------------------
# Amal-Plast specific extractor
# ---------------------------------------------------------------------------

_AMAL_PLAST_AP_RE = re.compile(r'^AP\d+', re.IGNORECASE)


def _is_amal_plast_document(text: str) -> bool:
    return bool(re.search(r'amal.?plast', text, re.IGNORECASE))


def _amal_find_ap_column(table: list[list], start_row: int) -> int:
    """Return the column index that contains the most AP-prefix codes."""
    ncols = max((len(r) for r in table if r), default=0)
    if not ncols:
        return -1
    counts = [0] * ncols
    for row in table[start_row: start_row + 15]:
        for i, cell in enumerate(row[:ncols]):
            if cell and _AMAL_PLAST_AP_RE.match(str(cell).strip()):
                counts[i] += 1
    best = max(range(ncols), key=lambda i: counts[i])
    return best if counts[best] > 0 else -1


def _parse_amal_plast_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from an Amal-Plast invoice.

    Column layout (0-indexed):
      0:Lp  1:Towar/Usługa  2:Kod  3:Jednostka  4:Ilość
      5:Cena netto  6:VAT  7:Cena brutto  8:Wartość netto  9:Wartość brutto

    Works whether or not pdfplumber includes the header row.
    """
    if not table or len(table) < 2:
        return []

    _UNIT_WORDS = {"kpl", "szt", "pcs", "szt.", "j.m", "unit", "jednostka", "set"}

    def norm(s):
        return str(s or "").lower().strip()

    def cell_text(row, idx):
        if idx is None or idx < 0 or idx >= len(row):
            return ""
        val = str(row[idx] or "").strip()
        for line in val.split("\n"):
            if line.strip():
                return line.strip()
        return val

    # ── Phase 1: try to find header row ──────────────────────────────────────
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(norm(c) for c in row)
        if any(kw in joined for kw in ["code", "kod"]) and \
           any(kw in joined for kw in ["qty", "ilosc", "ilo"]):
            header_idx = i
            break

    if header_idx is not None:
        headers = [norm(c) for c in table[header_idx]]
        logger.info("Amal-Plast: header at row %d: %s", header_idx, headers)

        def find(kws):
            for kw in kws:
                for i, h in enumerate(headers):
                    if kw in h:
                        return i
            return None

        code_idx  = find(["code", "kod"])
        qty_idx   = find(["ilo", "qty", "quantity"])
        unit_idx  = find(["jednostk", "unit", "j.m"])
        price_idx = find(["cena netto", "net price", "netto"])
        total_idx = find(["warto", "net value", "wartosc"])
        desc_idx  = find(["product", "towar", "service", "nazwa"])
        data_start = header_idx + 1

    else:
        # ── Phase 2: no header — detect columns from data ────────────────────
        logger.info("Amal-Plast: no header in %d-row table, scanning for AP codes", len(table))
        code_idx = _amal_find_ap_column(table, 0)
        if code_idx < 0:
            logger.info("Amal-Plast: no AP codes found, skipping table")
            return []
        logger.info("Amal-Plast: AP code column detected at index %d", code_idx)
        # Relative layout: code  unit  qty  net_price  VAT  gross_price  net_total  gross_total
        unit_idx  = code_idx + 1
        qty_idx   = code_idx + 2
        price_idx = code_idx + 3
        # skip VAT column (+4)
        total_idx = code_idx + 6   # Wartość netto
        desc_idx  = code_idx - 1 if code_idx > 0 else None
        data_start = 0

    # Validate / auto-correct code_idx by sampling
    def _sample_ap(col):
        for row in table[data_start: data_start + 5]:
            v = cell_text(row, col)
            if v and v.lower() not in ("kod", "code", "nan", "") \
                    and v.lower() not in _UNIT_WORDS:
                return v
        return ""

    sample = _sample_ap(code_idx)
    if sample and not _AMAL_PLAST_AP_RE.match(sample):
        for candidate in [code_idx - 1, code_idx + 1, code_idx + 2]:
            s = _sample_ap(candidate)
            if s and _AMAL_PLAST_AP_RE.match(s):
                logger.info("Amal-Plast: code_idx corrected %d→%d (%r→%r)",
                            code_idx, candidate, sample, s)
                code_idx = candidate
                break

    records = []
    for row in table[data_start:]:
        if not any(str(c or "").strip() for c in row):
            continue

        code = cell_text(row, code_idx)
        if not code or code.lower() in ("kod", "code", "nan", "") \
                or code.lower() in _UNIT_WORDS:
            continue
        if code.isdigit() and len(code) <= 3:
            continue
        if not _AMAL_PLAST_AP_RE.match(code):
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code

        if desc_idx is not None:
            desc = cell_text(row, desc_idx)
            if desc and len(desc) > 3 and not re.match(r'^\d+$', desc):
                rec.product_name = desc[:120]

        unit_raw = cell_text(row, unit_idx) if unit_idx is not None else ""
        qty_raw  = cell_text(row, qty_idx)  if qty_idx  is not None else ""
        if qty_raw:
            try:
                n = int(float(qty_raw.replace(",", ".")))
                qty_str = str(n)
            except (ValueError, TypeError):
                qty_str = qty_raw
            # Append unit (Jednostka) so it appears in the QUANTITY column
            if unit_raw and unit_raw.lower() in _UNIT_WORDS:
                rec.quantity = qty_str + " " + unit_raw.upper()
            else:
                rec.quantity = qty_str

        price_raw = cell_text(row, price_idx) if price_idx is not None else ""
        if price_raw:
            p = re.sub(r'[^\d.,]', '', price_raw).replace(',', '.')
            if p:
                rec.price = p + " EUR"

        total_raw = cell_text(row, total_idx) if total_idx is not None else ""
        if total_raw:
            t = re.sub(r'[^\d.,]', '', total_raw).replace(',', '.')
            if t:
                rec.total_price = t + " EUR"

        records.append(rec)

    logger.info("Amal-Plast table: %d records extracted", len(records))
    return records


_AMAL_DEC_RE   = re.compile(r'\b(\d{1,6}[.,]\d{2})\b')
_AMAL_UNIT_RE  = re.compile(r'\b(kpl|szt\.?|pcs|set)\b', re.IGNORECASE)
_AMAL_ROW_RE   = re.compile(r'^\d+\s+')


def _parse_amal_plast_from_text(text: str) -> list[ProductRecord]:
    """Text-based fallback: scan raw PDF text lines for AP codes.

    Expected line layout (pdfplumber text output):
      {lp}  {description}  {AP_code}  {unit}  {qty}  {cena_netto}  {VAT}
      {cena_brutto}  {wartość_netto}  {wartość_brutto}
    """
    records = []
    seen_codes: dict[str, list[int]] = {}

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        ap_matches = list(re.finditer(r'\b(AP\d{4,})\b', line, re.IGNORECASE))
        if not ap_matches:
            continue

        # Use last AP code occurrence (code column, not description column)
        m = ap_matches[-1]
        code = m.group(1).upper()

        rec = ProductRecord(extraction_method="text")
        rec.product_code = code

        # Description: text before the first AP occurrence minus leading row number
        pre = line[:ap_matches[0].start()].strip()
        pre = _AMAL_ROW_RE.sub('', pre).strip()
        if pre and len(pre) > 3:
            rec.product_name = pre[:120]

        # Analyse text AFTER the last AP code
        post = line[m.end():].strip()

        unit_m = _AMAL_UNIT_RE.search(post)
        unit_str = ""
        if unit_m:
            unit_str = unit_m.group(1).upper()
            after_unit = post[unit_m.end():].strip()
            qty_m = re.match(r'(\d+)\b', after_unit)
            if qty_m:
                rec.quantity = qty_m.group(1) + " " + unit_str

        # Decimal numbers after AP code: cena_netto  (VAT)  cena_brutto
        # wartość_netto  wartość_brutto
        decimals = _AMAL_DEC_RE.findall(post)
        if decimals:
            p = _clean_num(decimals[0])
            if p:
                rec.price = p + " EUR"
        if len(decimals) >= 3:
            # wartość netto is 3rd decimal (VAT is not decimal, so offset is 0)
            t = _clean_num(decimals[2])
            if t:
                rec.total_price = t + " EUR"
        elif len(decimals) == 2:
            t = _clean_num(decimals[1])
            if t:
                rec.total_price = t + " EUR"

        _smart_merge_or_add(records, seen_codes, rec)
        logger.info("Amal-Plast text: %s qty=%s price=%s total=%s",
                    code, rec.quantity, rec.price, rec.total_price)

    return records


def extract_amal_plast_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Amal-Plast: %d table(s) received", len(tables))
    records = []
    for table in tables:
        records.extend(_parse_amal_plast_table(table))

    if not records and text:
        logger.info("Amal-Plast: table extraction yielded 0 records, trying text fallback")
        records = _parse_amal_plast_from_text(text)
        logger.info("Amal-Plast text fallback: %d records", len(records))
    else:
        logger.info("Amal-Plast extraction: %d records total", len(records))
    return records


# ---------------------------------------------------------------------------
# Gumarny Zubri-specific extractor
# ---------------------------------------------------------------------------
#
# Invoice layout (English):
#   Item code | Name | Tax No. | Quantity | Price | Unit | VAT % | Rabat % | Total amount
#
# Product codes: 4-8 digits optionally followed by 0-4 letters  (222349, 216527BAL)
# Section-header rows have short letter-only codes (BO, SP) — no digits → skipped.
# Number format: Czech comma-decimal  (5,000 = 5 units;  16,60 = 16.60 EUR)

_GZ_CODE_RE = re.compile(r'^[A-Z]{0,2}\d{4,8}[A-Z]{0,4}$')
# Short integer that acts as tax-category separator between description and qty
_GZ_TAX_SEP_RE = re.compile(r'^\d{1,3}$')


def _is_gumarny_zubri_document(text: str) -> bool:
    # Match both ASCII "Gumarny Zubri" and Czech "Gumárny Zubří" (á=á, ř=ř, í=í)
    return bool(re.search(r'gum[aá]rny|zubr[řr][ií]', text, re.IGNORECASE))


def _gz_num(raw: str) -> str:
    """Czech comma-decimal → dot-decimal string, integer-stripped if no fraction."""
    if not raw or not raw.strip():
        return ""
    v = raw.strip().replace(',', '.')
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else f"{f:.2f}"
    except ValueError:
        return raw.strip()


def _gz_eff_price(list_price_raw: str, rabat_raw: str) -> str:
    """Effective unit price after Czech rabat (discount) percentage."""
    try:
        lp = float(list_price_raw.replace(',', '.'))
        rabat = float(rabat_raw.replace(',', '.'))
        return f"{lp * (1 - rabat / 100):.2f}"
    except ValueError:
        return _gz_num(list_price_raw)


def _gz_merge_into(existing: "ProductRecord", newer: "ProductRecord") -> None:
    """Merge a duplicate GZ record into an existing one by summing qty and totals."""
    # Sum quantity
    if newer.quantity and existing.quantity:
        try:
            q = float(existing.quantity) + float(newer.quantity)
            existing.quantity = str(int(q)) if q == int(q) else f"{q:.3f}"
        except ValueError:
            pass
    elif newer.quantity and not existing.quantity:
        existing.quantity = newer.quantity

    # Sum total_price
    def _strip_eur(s):
        return s.replace(" EUR", "").strip() if s else ""

    if newer.total_price and existing.total_price:
        try:
            t = float(_strip_eur(existing.total_price)) + float(_strip_eur(newer.total_price))
            existing.total_price = f"{t:.2f} EUR"
        except ValueError:
            pass
    elif newer.total_price and not existing.total_price:
        existing.total_price = newer.total_price

    # Fill missing fields from newer record
    if not existing.price and newer.price:
        existing.price = newer.price
    if not existing.product_name and newer.product_name:
        existing.product_name = newer.product_name


def _parse_gumarny_zubri_table(table: list[list]) -> list[ProductRecord]:
    if not table or len(table) < 2:
        return []

    def norm(s):
        # Collapse all whitespace (including \n from multi-line header cells)
        return re.sub(r'\s+', ' ', str(s or "").lower()).strip()

    def cell(row, idx):
        if idx is None or idx < 0 or idx >= len(row):
            return ""
        # Normalize multi-line cell content to single line
        return re.sub(r'\s+', ' ', str(row[idx] or "")).strip()

    # ── Phase 1: locate header row ────────────────────────────────────────────
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(norm(c) for c in row)
        if "item code" in joined and ("quantity" in joined or "total" in joined):
            header_idx = i
            break
        # Fallback: "item" + "name" + "quantity" without requiring "item code"
        if "item" in joined and "name" in joined and "quantity" in joined:
            header_idx = i
            break

    if header_idx is not None:
        headers = [norm(c) for c in table[header_idx]]
        ncols_header = len(headers)
        logger.info("Gumarny Zubri: header at row %d: %s", header_idx, headers)

        def find(kws):
            for kw in kws:
                for i, h in enumerate(headers):
                    if kw in h:
                        return i
            return None

        code_idx  = find(["item code", "item"])
        desc_idx  = find(["name", "description"])
        qty_idx   = find(["quantity", "qty"])
        # "unit" must not accidentally match "quantity" — check explicitly
        unit_idx  = find(["unit no", "unit"])
        price_idx = find(["price"])
        rabat_idx = find(["rabat"])
        # "total amount" before plain "total" to avoid matching quantity/total confusion
        total_idx = find(["total amount", "total"])

        # If "unit" hit "price" or "total" column (shouldn't happen but guard it)
        if unit_idx is not None and unit_idx == price_idx:
            unit_idx = None
        if unit_idx is not None and unit_idx == total_idx:
            unit_idx = None

        # Detect degenerate merged header: all key columns collapse to same index
        # (pdfplumber merged all header cells into one) → fall through to Phase 2
        key_indices = [i for i in [code_idx, desc_idx, qty_idx, price_idx, total_idx]
                       if i is not None]
        header_is_degenerate = ncols_header <= 2 or (
            len(key_indices) >= 3 and len(set(key_indices)) == 1
        )

        if header_is_degenerate:
            logger.info("Gumarny Zubri: header degenerate (merged cells) — falling to Phase 2")
            header_idx = None  # force Phase 2 below
        else:
            logger.info("Gumarny Zubri: cols — code=%s desc=%s qty=%s unit=%s price=%s rabat=%s total=%s",
                        code_idx, desc_idx, qty_idx, unit_idx, price_idx, rabat_idx, total_idx)
            data_start = header_idx + 1

    if header_idx is None:
        # ── Phase 2: continuation table (page 2+) — scan for code column ─────
        logger.info("Gumarny Zubri: no header in %d-row table — scanning for code column",
                    len(table))
        ncols = max((len(r) for r in table if r), default=0)
        if not ncols:
            return []
        counts = [0] * ncols
        for row in table[:min(20, len(table))]:
            for ci in range(min(len(row), ncols)):
                if row[ci] and _GZ_CODE_RE.match(str(row[ci]).strip()):
                    counts[ci] += 1
        best = max(range(ncols), key=lambda i: counts[i])
        if counts[best] == 0:
            logger.info("Gumarny Zubri: no code pattern in continuation table — skipping")
            return []
        # Column layout relative to code column:
        # code | desc | tax_no | qty | price | unit | vat% | rabat% | total
        # Column layout: code | desc | tax_no | qty | price | unit | vat% | rabat% | total
        code_idx  = best
        desc_idx  = best + 1 if best + 1 < ncols else None
        qty_idx   = best + 3 if best + 3 < ncols else None
        price_idx = best + 4 if best + 4 < ncols else None
        unit_idx  = best + 5 if best + 5 < ncols else None
        rabat_idx = best + 7 if best + 7 < ncols else None
        total_idx = best + 8 if best + 8 < ncols else None
        # If no numeric columns are reachable (table too narrow) the data is
        # useless — text extraction will handle these rows.
        if qty_idx is None and price_idx is None and total_idx is None:
            logger.info("Gumarny Zubri: continuation table too narrow — skipping, text will cover")
            return []
        logger.info("Gumarny Zubri: continuation cols — code=%d desc=%s qty=%s "
                    "price=%s unit=%s rabat=%s total=%s",
                    code_idx, desc_idx, qty_idx, price_idx, unit_idx, rabat_idx, total_idx)
        data_start = 0

    if code_idx is None:
        return []

    records = []
    for row in table[data_start:]:
        if not any(str(c or "").strip() for c in row):
            continue
        code = cell(row, code_idx)
        if not code:
            continue
        if not re.search(r'\d', code):   # skip section headers like "BO"
            continue
        if not _GZ_CODE_RE.match(code):
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code

        if desc_idx is not None:
            desc = cell(row, desc_idx)
            if desc and len(desc) > 2:
                rec.product_name = desc[:120]

        qty_raw  = cell(row, qty_idx)  if qty_idx  is not None else ""
        unit_raw = cell(row, unit_idx) if unit_idx is not None else ""
        if qty_raw:
            qty_str = _gz_num(qty_raw)
            if qty_str:
                rec.quantity = qty_str + (" " + unit_raw if unit_raw else "")

        price_raw = cell(row, price_idx) if price_idx is not None else ""
        rabat_raw = cell(row, rabat_idx) if rabat_idx is not None else ""
        if price_raw:
            p = _gz_eff_price(price_raw, rabat_raw) if rabat_raw else _gz_num(price_raw)
            if p:
                rec.price = p + " EUR"

        total_raw = cell(row, total_idx) if total_idx is not None else ""
        if total_raw:
            t = _gz_num(total_raw)
            if t:
                rec.total_price = t + " EUR"

        records.append(rec)

    logger.info("Gumarny Zubri table: %d records", len(records))
    return records


def _parse_gumarny_zubri_from_text(text: str) -> list[ProductRecord]:
    """Parse Gumarny Zubri invoice lines from raw text.

    Line format: {code} {desc...} {TAX_NO} {qty} {price} {unit} {vat%} {rabat%} {total}
    TAX_NO is a short integer (e.g. 47, 21, 0) — the CZ tax-category separator.
    It is identified as any 1–3 digit token immediately followed by a Czech-format
    quantity (e.g. 5,000).  This avoids hard-coding "47" and captures items with
    other tax codes (like P214423FL which may use 21 or 0).

    Duplicate codes are merged: quantities and total prices are summed.
    """
    records: list[ProductRecord] = []
    seen_codes: dict[str, int] = {}  # code → index in records

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        tokens = line.split()
        if len(tokens) < 4:
            continue
        code = tokens[0]
        if not _GZ_CODE_RE.match(code):
            continue

        # Locate tax-category separator: a short integer immediately followed by
        # a Czech-format decimal (quantity field).
        sep_idx = None
        for i in range(1, len(tokens) - 1):
            if _GZ_TAX_SEP_RE.match(tokens[i]) and re.match(r'^\d+[,.]\d+$', tokens[i + 1]):
                sep_idx = i
                break
        if sep_idx is None:
            continue

        # Description: tokens between code and separator
        desc = " ".join(tokens[1:sep_idx]).strip()[:120] or None

        # After separator: qty price unit vat% rabat% total
        after = tokens[sep_idx + 1:]

        # Quantity: first token after separator in Czech 3-decimal format (5,000 = 5)
        qty_str = ""
        if after and re.match(r'^\d+[,.]\d+$', after[0]):
            qty_str = _gz_num(after[0])

        # 2-decimal numbers from the *after* section only (avoids false matches in desc)
        # Layout: price | unit | vat% | rabat% | total
        # e.g.   18,90   Set    0,00   40,00    56,70  → after_dec = ['18,90','0,00','40,00','56,70']
        after_dec = re.findall(r'\b\d{1,6}[,.]\d{2}\b', ' '.join(after))

        # Effective unit price: list_price × (1 − rabat/100)
        price_eff = ""
        if after_dec:
            rabat_raw = after_dec[-2] if len(after_dec) >= 3 else ""
            price_eff = _gz_eff_price(after_dec[0], rabat_raw) if rabat_raw else _gz_num(after_dec[0])

        total_eff = _gz_num(after_dec[-1]) if after_dec else ""

        # Duplicate: merge into existing record
        if code in seen_codes:
            existing = records[seen_codes[code]]
            newer = ProductRecord(extraction_method="text")
            newer.product_code = code
            newer.product_name = desc
            newer.quantity = qty_str if qty_str else None
            newer.price = (price_eff + " EUR") if price_eff else None
            newer.total_price = (total_eff + " EUR") if total_eff else None
            _gz_merge_into(existing, newer)
            continue

        seen_codes[code] = len(records)
        rec = ProductRecord(extraction_method="text")
        rec.product_code = code
        rec.product_name = desc

        if qty_str:
            rec.quantity = qty_str
        if price_eff:
            rec.price = price_eff + " EUR"
        if total_eff:
            rec.total_price = total_eff + " EUR"

        records.append(rec)

    logger.info("Gumarny Zubri text: %d records", len(records))
    return records


def extract_gumarny_zubri_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Gumarny Zubri: %d table(s) received", len(tables))
    records: list[ProductRecord] = []
    seen_codes: dict[str, int] = {}  # code → index in records

    for table in tables:
        for rec in _parse_gumarny_zubri_table(table):
            code = rec.product_code
            if code not in seen_codes:
                seen_codes[code] = len(records)
                records.append(rec)
            else:
                _gz_merge_into(records[seen_codes[code]], rec)

    # Always supplement with text — fills in records missed by table extraction
    # and patches missing numeric fields on records that table found but couldn't
    # parse (e.g. degenerate 2-column tables with merged headers).
    if text:
        text_records = _parse_gumarny_zubri_from_text(text)
        added = 0
        for rec in text_records:
            code = rec.product_code
            if code not in seen_codes:
                seen_codes[code] = len(records)
                records.append(rec)
                added += 1
            else:
                existing = records[seen_codes[code]]
                if not existing.price and rec.price:
                    existing.price = rec.price
                if not existing.total_price and rec.total_price:
                    existing.total_price = rec.total_price
                if not existing.quantity and rec.quantity:
                    existing.quantity = rec.quantity
                if not existing.product_name and rec.product_name:
                    existing.product_name = rec.product_name
        if added:
            logger.info("Gumarny Zubri text: added %d records not found in tables", added)

    logger.info("Gumarny Zubri extraction: %d records total", len(records))
    return records


# ---------------------------------------------------------------------------
# RIGUM
# ---------------------------------------------------------------------------

# Group 1 = 6-digit product code; group 2 = rest of line (description on invoice)
_RIGUM_CODE_RE = re.compile(r'(?:^|\s)(\d{6}):[A-Z0-9]+(?:-[A-Z0-9]+)?\s*(.*)', re.UNICODE)


def _is_rigum_document(text: str) -> bool:
    return bool(re.search(r'rigum', text, re.IGNORECASE))


def _parse_rigum_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Rigum/POHODA invoice.

    Expected columns: Označení dodávky | Množství | J.cena | Sleva | € Celkem
    """
    if not table or len(table) < 2:
        return []

    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if "označení" in joined or "množství" in joined or "j.cena" in joined:
            header_idx = i
            break
    if header_idx is None:
        return []

    headers = [str(c or "").lower().strip() for c in table[header_idx]]

    def find(kws):
        for kw in kws:
            for idx, h in enumerate(headers):
                if kw in h:
                    return idx
        return None

    desc_idx  = find(["označení", "dodávky"]) or 0
    qty_idx   = find(["množství", "množstvi"])
    price_idx = find(["j.cena", "jednotk"])
    total_idx = find(["celkem", "total"])

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(ci):
            if ci is None or ci >= len(row):
                return ""
            return str(row[ci] or "").strip()

        desc_raw = cell(desc_idx)
        m = _RIGUM_CODE_RE.search(desc_raw)
        if not m:
            continue

        code = m.group(1)
        name = m.group(2).strip()

        quantity = total_price = price = None

        # Try separate columns first
        qty_raw = cell(qty_idx)
        if qty_raw:
            qm = re.match(r'^(\d+)\s*(sada|ks)\.?', qty_raw, re.IGNORECASE)
            if qm:
                quantity = _rigum_qty_label(int(qm.group(1)), qm.group(2))

        price_raw = cell(price_idx)
        if price_raw and re.search(r'\d', price_raw):
            price = price_raw.replace(',', '.') + ' EUR'

        total_raw = cell(total_idx)
        if total_raw and re.search(r'\d', total_raw):
            total_price = total_raw.replace(',', '.') + ' EUR'

        # When all numeric data is merged into the description cell, parse it out
        if quantity is None and price is None:
            qty_m = re.search(r'\b(\d{1,3})\s+(sada|ks)\.?\b', name, re.IGNORECASE)
            if qty_m:
                quantity   = _rigum_qty_label(int(qty_m.group(1)), qty_m.group(2))
                after_unit = name[qty_m.end():]
                prices     = re.findall(r'\b(\d{1,6}[.,]\d{2})\b', after_unit)
                name       = name[:qty_m.start()].strip()
                if len(prices) >= 2:
                    price       = prices[0].replace(',', '.') + ' EUR'
                    total_price = prices[-1].replace(',', '.') + ' EUR'
                elif len(prices) == 1:
                    price = prices[0].replace(',', '.') + ' EUR'

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        records.append(rec)

    return records


def _rigum_qty_label(n: int, unit: str) -> str:
    """Convert Czech unit to Bulgarian label."""
    if unit.lower() == "sada":
        return f"{n} {'Комплект' if n == 1 else 'Комплекта'}"
    return f"{n} {'Брой' if n == 1 else 'Броя'}"


def _rigum_merge_or_add(records: list, seen_map: dict, rec: ProductRecord) -> None:
    """Append rec to records, merging with an existing entry if same code+price.

    Same unit price → sum quantities and totals (one export row).
    Different unit price → keep as a separate row.
    seen_map: dict[code, list[int]] mapping product_code → indices in records.
    """
    code = rec.product_code
    if code in seen_map:
        for idx in seen_map[code]:
            existing = records[idx]
            if existing.price == rec.price:
                existing.merged_count += 1
                if rec.quantity and existing.quantity:
                    try:
                        em = re.search(r'(\d+)\s+(Брой|Броя|Комплект|Комплекта)', existing.quantity)
                        nm = re.search(r'(\d+)\s+(Брой|Броя|Комплект|Комплекта)', rec.quantity)
                        if em and nm:
                            total_n = int(em.group(1)) + int(nm.group(1))
                            unit_w = em.group(2)
                            if unit_w in ('Брой', 'Броя'):
                                existing.quantity = f"{total_n} {'Брой' if total_n == 1 else 'Броя'}"
                            else:
                                existing.quantity = f"{total_n} {'Комплект' if total_n == 1 else 'Комплекта'}"
                    except Exception:
                        pass
                if rec.total_price and existing.total_price:
                    try:
                        ep = float(existing.total_price.replace(' EUR', '').replace(',', '.'))
                        np_ = float(rec.total_price.replace(' EUR', '').replace(',', '.'))
                        existing.total_price = f"{ep + np_:.2f} EUR"
                    except Exception:
                        pass
                return
        seen_map[code].append(len(records))
    else:
        seen_map[code] = [len(records)]
    records.append(rec)


def _parse_rigum_from_text(text: str) -> list[ProductRecord]:
    """Text fallback when pdfplumber finds no usable tables."""
    records: list[ProductRecord] = []
    seen: dict[str, list[int]] = {}
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _RIGUM_CODE_RE.search(line)
        if not m:
            i += 1
            continue

        code  = m.group(1)
        after = m.group(2).strip()

        # Collect continuation lines until next product code
        extra_lines: list[str] = []
        j = i + 1
        while j < len(lines) and j <= i + 4:
            nl = lines[j].strip()
            if not nl:
                j += 1
                continue
            if _RIGUM_CODE_RE.search(nl):
                break
            extra_lines.append(nl)
            j += 1

        # Build description — stop at quantity/price data
        name_parts = [after] if after else []
        for nl in extra_lines:
            if re.search(r'\b\d+\s+(?:sada|ks)\b', nl, re.IGNORECASE):
                break
            if re.search(r'\b\d{1,6}[.,]\d{2}\b', nl):
                break
            name_parts.append(nl)
        name = " ".join(name_parts).strip()

        search_text = " ".join([line] + extra_lines)

        # Quantity: "5 sada" / "12 ks"
        qty_m = re.search(r'\b(\d{1,3})\s+(sada|ks)\.?\b', search_text, re.IGNORECASE)
        if qty_m:
            n    = int(qty_m.group(1))
            unit = qty_m.group(2)
            quantity    = _rigum_qty_label(n, unit)
            after_unit  = search_text[qty_m.end():]
            prices      = re.findall(r'\b(\d{1,6}[.,]\d{2})\b', after_unit)
        else:
            quantity = None
            prices   = re.findall(r'\b(\d{1,6}[.,]\d{2})\b', search_text)

        # Use prices[1] as total — not prices[-1] which may pick up
        # cumulative page totals that appear after the product line.
        if len(prices) >= 2:
            price       = prices[0].replace(',', '.') + ' EUR'
            total_price = prices[1].replace(',', '.') + ' EUR'
        elif len(prices) == 1:
            price       = prices[0].replace(',', '.') + ' EUR'
            total_price = None
        else:
            price = total_price = None

        logger.info("Rigum code=%s qty=%s price=%s total=%s | %s",
                    code, quantity, price, total_price, search_text[:120])

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        _rigum_merge_or_add(records, seen, rec)
        i = j

    logger.info("Rigum text extraction: %d records", len(records))
    return records


def extract_rigum_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Rigum: %d table(s) received", len(tables))
    logger.info("Rigum text first 600 chars:\n%s", text[:600])
    for ti, table in enumerate(tables[:3]):
        for ri, row in enumerate(table[:3]):
            logger.info("Rigum table[%d] row[%d]: %s", ti, ri, [str(c or "")[:50] for c in row])

    table_records: list[ProductRecord] = []
    seen: dict[str, list[int]] = {}
    for table in tables:
        for rec in _parse_rigum_table(table):
            _rigum_merge_or_add(table_records, seen, rec)

    text_records = _parse_rigum_from_text(text) if text else []

    # Always prefer whichever source yields more records
    records = text_records if len(text_records) > len(table_records) else table_records
    logger.info("Rigum: %d records (table=%d text=%d)",
                len(records), len(table_records), len(text_records))
    return records


# ---------------------------------------------------------------------------
# FROGUM
# ---------------------------------------------------------------------------

_FROGUM_LINE_RE = re.compile(r'^\s*\d{1,3}\s+([A-Z0-9][A-Z0-9\-\.]{3,11})\s+(.*)', re.UNICODE)


def _is_frogum_document(text: str) -> bool:
    return bool(re.search(r'frogum', text, re.IGNORECASE))


def _frogum_qty_label(n: int, unit: str) -> str:
    if re.search(r'\b(set|kpl)\b', unit, re.IGNORECASE):
        return f"{n} {'Комплект' if n == 1 else 'Комплекта'}"
    return f"{n} {'Брой' if n == 1 else 'Броя'}"


def _parse_frogum_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Frogum invoice.

    Columns: Item | Reference | Description | PKWiU | Unit | Q-TY |
             VAT[%] | Net price | VAT amount | Gross value | Net value
    """
    if not table or len(table) < 2:
        return []

    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(re.sub(r'\s+', ' ', str(c or "")).lower() for c in row)
        if "reference" in joined and ("q-ty" in joined or "net price" in joined):
            header_idx = i
            break
    if header_idx is None:
        return []

    headers = [re.sub(r'\s+', ' ', str(c or "")).lower().strip() for c in table[header_idx]]

    def find(kws):
        for kw in kws:
            for idx, h in enumerate(headers):
                if kw in h:
                    return idx
        return None

    ref_idx        = find(["reference"])
    desc_idx       = find(["description"])
    unit_idx       = find(["unit"])
    qty_idx        = find(["q-ty", "qty", "quantity", "ilość", "ilosc"])
    list_price_idx = find(["net price"])
    disc_idx       = find(["discount %", "discount"])
    price_idx      = find(["discounted net price"])
    total_idx      = find(["net value", "gross value"])

    logger.info("Frogum table cols → ref=%s qty=%s list_price=%s disc=%s price=%s total=%s | headers: %s",
                ref_idx, qty_idx, list_price_idx, disc_idx, price_idx, total_idx, headers)

    if ref_idx is None:
        return []

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(ci):
            if ci is None or ci >= len(row):
                return ""
            return str(row[ci] or "").strip()

        code = cell(ref_idx)
        # Normalize multi-line code cells (PDF wraps long codes) — take first line
        code = code.split('\n')[0].strip()
        if not re.match(r'^[A-Z0-9][A-Z0-9\-\.]{3,11}$', code):
            if code:
                logger.warning("Frogum: skipped row — code %r doesn't match pattern | row: %s",
                               code, [str(c or "")[:30] for c in row])
            continue

        name     = cell(desc_idx) if desc_idx is not None else ""
        unit_raw = cell(unit_idx) or ""

        qty_raw = cell(qty_idx) or ""
        quantity = None
        if qty_raw:
            _qm = re.match(r'^(\d+)', qty_raw)
            if _qm:
                n = int(_qm.group(1))
                quantity = _frogum_qty_label(n, unit_raw)

        price_raw = cell(price_idx) or ""
        if not price_raw and list_price_idx is not None and disc_idx is not None:
            lp_raw = cell(list_price_idx)
            dc_raw = cell(disc_idx)
            if lp_raw and dc_raw and re.search(r'\d', lp_raw) and re.search(r'\d', dc_raw):
                try:
                    lp = float(lp_raw.replace(',', '.'))
                    dc = float(dc_raw.replace(',', '.'))
                    price_raw = f"{lp * (1 - dc / 100):.2f}"
                except Exception:
                    pass
        price = (price_raw.replace(',', '.') + ' EUR') if price_raw and re.search(r'\d', price_raw) else None

        total_raw = cell(total_idx) or ""
        total_price = (total_raw.replace(',', '.') + ' EUR') if total_raw and re.search(r'\d', total_raw) else None

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        records.append(rec)

    return records


def _parse_frogum_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for Frogum invoices."""
    records: list[ProductRecord] = []
    seen: set[str] = set()
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _FROGUM_LINE_RE.match(line)
        if not m:
            i += 1
            continue

        code  = m.group(1)
        after = m.group(2).strip()

        if code in seen:
            logger.info("Frogum text: skipping duplicate code %s", code)
            i += 1
            continue

        # Collect continuation lines until next product
        extra_lines: list[str] = []
        j = i + 1
        while j < len(lines) and j <= i + 3:
            nl = lines[j].strip()
            if not nl:
                j += 1
                continue
            if _FROGUM_LINE_RE.match(nl):
                break
            extra_lines.append(nl)
            j += 1

        search_text = " ".join([line] + extra_lines)

        # Unit marker — English (set/pcs) or Polish (szt./kpl.)
        unit_m = re.search(r'\b(set|pcs?|szt\.?|kpl\.?)\b', search_text, re.IGNORECASE)
        quantity = price = total_price = None

        if unit_m:
            unit_str   = unit_m.group(1)
            after_unit = search_text[unit_m.end():]

            # Q-TY: first positive integer after unit that is not:
            #  - a VAT/discount rate followed by "%" ("0%", "25%")
            #  - the integer part of a decimal number ("10,25")
            #  - a bare "0" (VAT rate shown without % symbol in new format)
            for _qm in re.finditer(r'\b(\d+)\b', after_unit):
                _val = int(_qm.group(1))
                if _val == 0:
                    continue  # bare "0" = VAT rate without % symbol
                _sfx = after_unit[_qm.end():_qm.end() + 2].lstrip(' ')
                if _sfx.startswith('%') or _sfx.startswith(',') or _sfx.startswith('.'):
                    continue
                quantity = _frogum_qty_label(_val, unit_str)
                break

            # Collect monetary decimals; skip anything followed by % (VAT/disc rate).
            # Strategy: the VAT amount is always 0,00 on Frogum export invoices.
            # Old format 4 values: [net_price, 0,00, gross, net]  → 3 non-zero
            # New format 5 values: [list_price, net_price, 0,00, gross, net] → 4 non-zero
            # Using non-zero count lets us detect the discount column automatically.
            _nonzero: list[str] = []
            for _dm in re.finditer(r'\b(\d{1,6}[.,]\d{2})\b', after_unit):
                _sfx = after_unit[_dm.end():_dm.end() + 2].lstrip()
                if _sfx.startswith('%'):
                    continue
                if float(_dm.group(1).replace(',', '.')) > 0:
                    _nonzero.append(_dm.group(1))
            if len(_nonzero) >= 4:
                # New format with discount: skip first (list price), use second as unit price
                price       = _nonzero[1].replace(',', '.') + ' EUR'
                total_price = _nonzero[-1].replace(',', '.') + ' EUR'
            elif len(_nonzero) >= 2:
                price       = _nonzero[0].replace(',', '.') + ' EUR'
                total_price = _nonzero[-1].replace(',', '.') + ' EUR'
            elif len(_nonzero) == 1:
                price = _nonzero[0].replace(',', '.') + ' EUR'

        # Build description from the 'after' part (stop before unit/numeric data)
        name_parts = []
        for part in [after] + extra_lines:
            if re.search(r'\b(?:set|pcs?|szt\.?|kpl\.?)\b', part, re.IGNORECASE):
                break
            name_parts.append(part)
        name = " ".join(name_parts).strip()

        logger.info("Frogum code=%s qty=%s price=%s total=%s", code, quantity, price, total_price)

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        seen.add(code)
        records.append(rec)
        i = j

    logger.info("Frogum text extraction: %d records", len(records))
    return records


def extract_frogum_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Frogum: %d table(s) received", len(tables))

    table_records: list[ProductRecord] = []
    seen: dict[str, list[int]] = {}
    raw_table_count = 0
    for table in tables:
        recs = _parse_frogum_table(table)
        raw_table_count += len(recs)
        for rec in recs:
            _smart_merge_or_add(table_records, seen, rec)
    merged_count = raw_table_count - len(table_records)
    if merged_count:
        merged_codes = [c for c, idxs in seen.items() if table_records[idxs[0]].merged_count > 1]
        logger.info("Frogum table: %d raw rows → %d after merge (%d merged). Merged codes: %s",
                    raw_table_count, len(table_records), merged_count, merged_codes)
    else:
        logger.info("Frogum table: %d rows (no merges)", raw_table_count)

    text_records = _parse_frogum_from_text(text) if text else []

    # Use text for full code coverage; fill missing prices/qty from table
    if text_records and table_records:
        table_by_code = {r.product_code: r for r in table_records}
        text_codes = {r.product_code for r in text_records}
        only_in_table = [c for c in table_by_code if c not in text_codes]
        only_in_text  = [r.product_code for r in text_records if r.product_code not in table_by_code]
        if only_in_table:
            logger.warning("Frogum: codes in table but NOT in text (will be missing): %s", only_in_table)
        if only_in_text:
            logger.info("Frogum: codes in text but not in table: %s", only_in_text)
        for rec in text_records:
            t = table_by_code.get(rec.product_code)
            if t:
                # Always prefer table for all numeric fields — table uses
                # explicit column headers and is immune to text-layer noise
                # (discount %, VAT rate, carton qty appearing as product qty).
                if t.quantity:
                    rec.quantity = t.quantity
                if t.price:
                    rec.price = t.price
                if t.total_price:
                    rec.total_price = t.total_price
        # Append any table-only codes that text layer missed
        for code in only_in_table:
            records_copy = table_by_code[code]
            logger.info("Frogum: adding table-only record %s", code)
            text_records.append(records_copy)
        records = text_records
    else:
        records = table_records if table_records else text_records

    # Post-process: derive unit price from total / qty.
    # This corrects any column-mapping errors (e.g. discount % read as price)
    # and is always mathematically correct when both values are available.
    for _rec in records:
        if _rec.total_price and _rec.quantity:
            _qm = re.match(r'(\d+)', str(_rec.quantity))
            if _qm:
                _n = int(_qm.group(1))
                if _n > 0:
                    try:
                        _t = float(re.search(r'[\d.]+', _rec.total_price).group())
                        _sfx = ' EUR' if 'EUR' in (_rec.total_price or '') else ''
                        _rec.price = f"{round(_t / _n, 2):.2f}{_sfx}"
                    except Exception:
                        pass

    logger.info("Frogum: %d records (table=%d text=%d)",
                len(records), len(table_records), len(text_records))
    return records


# ---------------------------------------------------------------------------
# GELLY PLAST
# ---------------------------------------------------------------------------

_GELLY_PLAST_CODE_RE = re.compile(r'^GP\d{4,6}(?:[+\-][A-Z0-9]+)?$', re.IGNORECASE)


def _is_gelly_plast_document(text: str) -> bool:
    return bool(re.search(r'gelly.?plast|georgiades.*ltd|wind\s+deflectors', text, re.IGNORECASE))


def _gelly_plast_qty_label(n: int) -> str:
    return f"{n} {'Комплект' if n == 1 else 'Комплекта'}"


def _parse_gelly_plast_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Gelly Plast invoice.

    Columns: Code | Description | Quantity | Price per set € | Price €
    """
    if not table or len(table) < 2:
        return []

    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(re.sub(r'\s+', ' ', str(c or "")).lower() for c in row)
        if "quantity" in joined and "price" in joined:
            header_idx = i
            break

    if header_idx is not None:
        headers = [re.sub(r'\s+', ' ', str(c or "")).lower().strip() for c in table[header_idx]]

        def find(kws):
            for kw in kws:
                for idx, h in enumerate(headers):
                    if kw in h:
                        return idx
            return None

        def find_all(kw):
            return [idx for idx, h in enumerate(headers) if kw in h]

        qty_idx    = find(["quantity"])
        price_idx  = find(["per set"])
        price_cols = find_all("price")
        total_idx  = next((i for i in price_cols if i != price_idx), None)
        data_start = header_idx + 1
    else:
        # Continuation page — no header row; use fixed column layout
        # Col 0: code, Col 1: description, Col 2: qty, Col 3: price per set, Col 4: total
        qty_idx   = 2
        price_idx = 3
        total_idx = 4
        data_start = 0

    records = []
    for row in table[data_start:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(ci):
            if ci is None or ci >= len(row):
                return ""
            return re.sub(r'\s+', ' ', str(row[ci] or "")).strip()

        code = cell(0)
        if not _GELLY_PLAST_CODE_RE.match(code):
            continue

        name = cell(1) if len(row) > 1 else ""

        qty_raw = cell(qty_idx) or ""
        quantity = None
        if qty_raw and re.match(r'^\d+$', qty_raw):
            quantity = _gelly_plast_qty_label(int(qty_raw))

        price_raw = cell(price_idx) or ""
        price = (price_raw.replace(',', '.') + ' EUR') \
            if price_raw and re.match(r'^\d+(?:[.,]\d+)?$', price_raw) else None

        total_raw = cell(total_idx) or ""
        total_price = (total_raw.replace(',', '.') + ' EUR') \
            if total_raw and re.match(r'^\d+(?:[.,]\d+)?$', total_raw) else None

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code.upper()
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        records.append(rec)

    return records


def _parse_gelly_plast_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for Gelly Plast invoices."""
    records: list[ProductRecord] = []
    seen: set[str] = set()
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = re.match(r'^\s*(GP\d{4,6}(?:[+\-][A-Z0-9]+)?)\s+(.*)', line, re.IGNORECASE)
        if not m:
            i += 1
            continue

        code  = m.group(1).upper()
        after = m.group(2).strip()

        if code in seen:
            i += 1
            continue

        extra_lines: list[str] = []
        j = i + 1
        while j < len(lines) and j <= i + 5:
            nl = lines[j].strip()
            if not nl:
                j += 1
                continue
            if re.match(r'^\s*GP\d{4,6}(?:[+\-][A-Z0-9]+)?\s+', nl, re.IGNORECASE):
                break
            extra_lines.append(nl)
            j += 1

        # If 'after' (part on the same line as code) contains no Cyrillic,
        # the numeric data (qty, price, total) is in 'after'; description is
        # in continuation lines.  Otherwise description and numbers are mixed
        # in the full search_text — fall back to post-Cyrillic heuristic.
        after_has_cyrillic = bool(re.search(r'[Ѐ-ӿ]', after))

        if not after_has_cyrillic:
            numeric_part = after
            cyrillic_positions = []
            for part in extra_lines:
                cp = [k for k, c in enumerate(part) if 'Ѐ' <= c <= 'ӿ']
                if cp:
                    cyrillic_positions = cp  # last-found wins
            name = " ".join(extra_lines).strip()
        else:
            search_text = " ".join([after] + extra_lines)
            cyrillic_positions = [k for k, c in enumerate(search_text) if 'Ѐ' <= c <= 'ӿ']
            numeric_part = search_text[cyrillic_positions[-1] + 1:] if cyrillic_positions else search_text
            name = (search_text[:cyrillic_positions[-1] + 1] if cyrillic_positions else after).strip()

        nums = re.findall(r'\b(\d+(?:[.,]\d+)?)\b', numeric_part)
        quantity = price = total_price = None
        if len(nums) >= 3:
            try:
                n = int(float(nums[-3].replace(',', '.')))
                quantity    = _gelly_plast_qty_label(n)
                price       = nums[-2].replace(',', '.') + ' EUR'
                total_price = nums[-1].replace(',', '.') + ' EUR'
            except (ValueError, IndexError):
                pass
        elif len(nums) == 2:
            price       = nums[0].replace(',', '.') + ' EUR'
            total_price = nums[1].replace(',', '.') + ' EUR'

        logger.info("GellyPlast code=%s qty=%s price=%s total=%s", code, quantity, price, total_price)

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        seen.add(code)
        records.append(rec)
        i = j

    logger.info("GellyPlast text extraction: %d records", len(records))
    return records


def extract_gelly_plast_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("GellyPlast: %d table(s) received", len(tables))

    table_records: list[ProductRecord] = []
    seen: dict[str, list[int]] = {}
    for table in tables:
        for rec in _parse_gelly_plast_table(table):
            _smart_merge_or_add(table_records, seen, rec)

    text_records = _parse_gelly_plast_from_text(text) if text else []

    if text_records and table_records:
        table_by_code = {r.product_code: r for r in table_records}
        for rec in text_records:
            t = table_by_code.get(rec.product_code)
            if t:
                # Table column extraction is authoritative for prices/qty
                if t.price:
                    rec.price = t.price
                if t.total_price:
                    rec.total_price = t.total_price
                if t.quantity:
                    rec.quantity = t.quantity
        records = text_records
    else:
        records = table_records if table_records else text_records

    logger.info("GellyPlast: %d records (table=%d text=%d)",
                len(records), len(table_records), len(text_records))
    return records


# ---------------------------------------------------------------------------
# FARAD (Evolution SRL)
# ---------------------------------------------------------------------------

def _is_farad_document(text: str) -> bool:
    return bool(re.search(r'evolution\s+srl|articolo.*imponibile|imponibile.*articolo',
                          text, re.IGNORECASE))


def _farad_num(s: str, decimals: int = 2) -> str | None:
    """Italian decimal format: 1.202,04 → 1202.04"""
    s = (s or "").strip()
    if not s:
        return None
    converted = s.replace('.', '').replace(',', '.')
    try:
        return f"{float(converted):.{decimals}f}"
    except ValueError:
        return None


def _farad_qty_label(n: int, um: str) -> str:
    if re.match(r'^C\.?$', um.strip(), re.IGNORECASE):
        return f"{n} {'Комплект' if n == 1 else 'Комплекта'}"
    return f"{n} {'Брой' if n == 1 else 'Броя'}"


def _parse_farad_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Farad/Evolution SRL invoice.

    Columns: Item/Articolo | Description | Um | Qty | Unit Price | Disc. | Net Price | Amount
    """
    if not table or len(table) < 2:
        return []

    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(re.sub(r'\s+', ' ', str(c or "")).lower() for c in row)
        if ("articolo" in joined or "item" in joined) and \
                ("qty" in joined or "imponibile" in joined or "netto" in joined):
            header_idx = i
            break

    if header_idx is not None:
        headers = [re.sub(r'\s+', ' ', str(c or "")).lower().strip()
                   for c in table[header_idx]]

        def find(kws):
            for kw in kws:
                for idx, h in enumerate(headers):
                    if kw in h:
                        return idx
            return None

        code_idx  = find(["articolo", "item"])
        desc_idx  = find(["description", "descrizione"])
        um_idx    = find(["um"])
        qty_idx   = find(["qty", "quantit"])
        price_idx = find(["netto", "net price", "prezzo netto", "p.netto",
                          "unit price", "prezzo unit", "prezzo"])
        total_idx = find(["imponibile", "amount"])
        data_start = header_idx + 1
    else:
        # Fixed column layout (continuation pages)
        code_idx, desc_idx, um_idx = 0, 1, 2
        qty_idx, price_idx, total_idx = 3, 6, 7
        data_start = 0

    records = []
    for row in table[data_start:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(ci):
            if ci is None or ci >= len(row):
                return ""
            return re.sub(r'\s+', ' ', str(row[ci] or "")).strip()

        # Code cell may contain both code and description separated by \n
        raw_code_cell = str(row[code_idx] or "") if code_idx is not None and code_idx < len(row) else ""
        code_parts = [p.strip() for p in raw_code_cell.split('\n') if p.strip()]
        code = code_parts[0] if code_parts else ""

        if not code or not re.match(r'^1-|\d{4,6}/', code, re.IGNORECASE):
            continue

        # Description: prefer dedicated column; fall back to embedded part.
        # Guard: when pdfplumber merges columns, desc_idx may point at the Um
        # column whose content is a unit marker like "NR" or "C." — skip it.
        desc_raw = str(row[desc_idx] or "").strip() if desc_idx is not None and desc_idx < len(row) else ""
        if desc_raw and not re.match(r'^(NR|C\.?)\s*$', desc_raw, re.IGNORECASE):
            name = re.sub(r'\s+', ' ', desc_raw).strip()
        elif len(code_parts) > 1:
            name = ' '.join(code_parts[1:])
        else:
            name = ""

        um    = cell(um_idx)

        quantity = None
        qty_raw = cell(qty_idx)
        qty_val = _farad_num(qty_raw)
        if qty_val:
            try:
                n = int(float(qty_val))
                quantity = _farad_qty_label(n, um)
            except ValueError:
                pass

        p = _farad_num(cell(price_idx), decimals=3)
        price = (p + " EUR") if p is not None else None

        t = _farad_num(cell(total_idx))
        total_price = (t + ' EUR') if t is not None else None

        # Fallback: derive unit price from total ÷ qty when price column absent
        if price is None and total_price and quantity:
            try:
                qty_n = int(quantity.split()[0])
                total_f = float(total_price.replace(' EUR', ''))
                if qty_n > 0:
                    price = f"{total_f / qty_n:.3f} EUR"
            except (ValueError, ZeroDivisionError):
                pass

        logger.info("Farad code=%s qty=%s price=%s total=%s", code, quantity, price, total_price)

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        records.append(rec)

    return records


def _parse_farad_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for Farad/Evolution SRL invoices."""
    records: list[ProductRecord] = []
    seen: set[str] = set()
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        # Line starts with 1- code (or bare numeric code like 90241/8 110), then UM, then numbers
        m = re.match(r'^((?:1-)?\d[\w/]*(?:\s+\S+){0,5}?)\s+(NR|C\.?)\s+([\d,]+)', line, re.IGNORECASE)
        if not m:
            i += 1
            continue

        code    = m.group(1).strip()
        um      = m.group(2)
        qty_raw = m.group(3)

        if code in seen:
            i += 1
            continue

        qty_val = _farad_num(qty_raw)
        quantity = None
        if qty_val:
            try:
                n = int(float(qty_val))
                quantity = _farad_qty_label(n, um)
            except ValueError:
                pass

        # Numbers after qty: unit_price disc net_price amount
        # Discount is like "60+10%" — skip non-pure-decimal tokens
        rest = line[m.end():].strip()
        decimal_nums = [_farad_num(tok, decimals=3) for tok in re.split(r'\s+', rest)
                        if re.match(r'^[\d.,]+$', tok)]
        decimal_nums = [v for v in decimal_nums if v is not None]

        price = total_price = None
        if len(decimal_nums) >= 2:
            price       = decimal_nums[-2] + ' EUR'
            total_price = _farad_num(decimal_nums[-1]) + ' EUR'
        elif len(decimal_nums) == 1:
            total_price = _farad_num(decimal_nums[0]) + ' EUR'

        # Description: collect continuation lines before next 1- code
        name_parts: list[str] = []
        j = i + 1
        while j < len(lines) and j <= i + 2:
            nl = lines[j].strip()
            if not nl or re.match(r'^1-', nl, re.IGNORECASE):
                break
            name_parts.append(nl)
            j += 1

        name = " ".join(name_parts).strip()

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        seen.add(code)
        records.append(rec)
        i += 1

    logger.info("Farad text extraction: %d records", len(records))
    return records


def extract_farad_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Farad: %d table(s) received", len(tables))

    table_records: list[ProductRecord] = []
    free_records: list[ProductRecord] = []
    seen: set[str] = set()
    seen_free: set[str] = set()
    for table in tables:
        for rec in _parse_farad_table(table):
            try:
                is_free = rec.price is not None and float(rec.price.replace(" EUR", "")) == 0.0
            except (ValueError, AttributeError):
                is_free = False
            if is_free:
                if rec.product_code not in seen_free:
                    seen_free.add(rec.product_code)
                    free_records.append(rec)
            else:
                if rec.product_code not in seen:
                    seen.add(rec.product_code)
                    table_records.append(rec)

    text_records = _parse_farad_from_text(text) if text else []

    # Table records are the primary source — the table parser correctly separates
    # code from description using dedicated columns. Text parsing merges them and
    # produces mismatched strings that break deduplication.
    if table_records:
        records = list(table_records)
        if text_records:
            # Only add text records for codes genuinely absent from table
            table_first_words = {r.product_code.split()[0] for r in table_records}
            for rec in text_records:
                first = rec.product_code.split()[0] if rec.product_code else ""
                if first not in table_first_words:
                    records.append(rec)
    else:
        records = list(text_records)

    records.extend(free_records)
    logger.info("Farad: %d records (table=%d text=%d free=%d)",
                len(records), len(table_records), len(text_records), len(free_records))
    return records


# ---------------------------------------------------------------------------
# GEYER & HOSAJA
# ---------------------------------------------------------------------------

# Short product code on line 2 of description cell: e.g. "802/4C Car mat ..."
_GH_CODE_RE = re.compile(r'^(\d{3}/\d[A-Z])\s+(.*)', re.UNICODE)


def _is_geyer_hosaja_document(text: str) -> bool:
    return bool(re.search(r'geyer.?hosaja|geter.?hosaja', text, re.IGNORECASE))


def _gh_qty_label(n: float, unit: str) -> str:
    n_int = int(n) if n == int(n) else n
    if re.search(r'\bset\b', unit, re.IGNORECASE):
        return f"{n_int} {'Комплект' if n_int == 1 else 'Комплекта'}"
    return f"{n_int} {'Брой' if n_int == 1 else 'Броя'}"


def _parse_geyer_hosaja_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Geyer & Hosaja invoice.

    Description cell has 3 lines:
      1. internal catalog code (01 1 99 036 21 00 1 1)  → ignored
      2. short code + name  (802/4C Car mat OPEL CORSA D black)
      3. order reference    (ORDER 30.03)                → ignored

    Columns: Lp. | Description | Unit | Quantity | Price | Discount |
             Price after discount | Net value | VAT rate | VAT value | Gross
    """
    if not table or len(table) < 2:
        return []

    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if ("nazwa" in joined or "description" in joined) and \
           ("quantity" in joined or "ilość" in joined):
            header_idx = i
            break
    if header_idx is None:
        return []

    headers = [str(c or "").lower().strip() for c in table[header_idx]]

    def find(kws):
        for kw in kws:
            for idx, h in enumerate(headers):
                if kw in h:
                    return idx
        return None

    desc_idx  = find(["nazwa", "description", "goods"]) or 0
    unit_idx  = find(["j.miary", "unit"])
    qty_idx   = find(["ilość", "quantity"])
    price_idx = find(["po upuście", "upuście", "after discount"])
    total_idx = find(["netto", "net value"])

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(ci):
            if ci is None or ci >= len(row):
                return ""
            return str(row[ci] or "").strip()

        # Description cell may have embedded newlines
        desc_raw = cell(desc_idx)
        code = name = None
        for dline in desc_raw.replace('\r', '\n').split('\n'):
            m = _GH_CODE_RE.match(dline.strip())
            if m:
                code = m.group(1)
                name = m.group(2).strip()
                break
        if not code:
            continue

        # Unit + quantity
        unit_raw = cell(unit_idx) or ""
        qty_raw  = cell(qty_idx)  or ""
        quantity = None
        if qty_raw:
            try:
                qty_val  = float(qty_raw.replace(',', '.'))
                quantity = _gh_qty_label(qty_val, unit_raw)
            except ValueError:
                pass

        # Unit price = Price after discount
        price_raw = cell(price_idx) or ""
        price = (price_raw.replace(',', '.') + ' EUR') if price_raw and re.search(r'\d', price_raw) else None

        # Total = Net value
        total_raw = cell(total_idx) or ""
        total_price = (total_raw.replace(',', '.') + ' EUR') if total_raw and re.search(r'\d', total_raw) else None

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        records.append(rec)

    return records


def _parse_geyer_hosaja_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for Geyer & Hosaja invoices.

    Anchors on lines matching the short code pattern (802/4C …).
    Searches that line and its ±3 neighbours for: unit, qty, price after
    discount (4th decimal after unit), net value (5th decimal after unit).
    """
    records: list[ProductRecord] = []
    seen: set[str] = set()
    lines = text.splitlines()

    for i, raw in enumerate(lines):
        line = raw.strip()
        m = _GH_CODE_RE.match(line)
        if not m:
            continue

        code = m.group(1)
        name = m.group(2).strip()
        # Strip trailing order reference appended on same line
        name = re.sub(r'\s+ORDER\s+.*$', '', name, flags=re.IGNORECASE).strip()

        if code in seen:
            continue

        # Gather context: 2 lines before + current + 3 lines after
        ctx_lines = []
        for k in range(max(0, i - 2), min(len(lines), i + 4)):
            ctx_lines.append(lines[k].strip())
        search_text = " ".join(ctx_lines)

        # Locate unit marker (set / pc / pcs)
        unit_m = re.search(r'\b(set|pc\.?|pcs\.?)\b', search_text, re.IGNORECASE)
        quantity = price = total_price = None

        if unit_m:
            unit_str   = unit_m.group(1)
            after_unit = search_text[unit_m.end():]
            # Decimal numbers after unit: qty, price, discount%, price_after, net, ...
            nums = re.findall(r'\b(\d+[.,]\d+)\b', after_unit)
            # Filter out VAT-like "0.00" duplicates and % values captured as decimals
            if len(nums) >= 5:
                try:
                    qty_val  = float(nums[0].replace(',', '.'))
                    quantity = _gh_qty_label(qty_val, unit_str)
                    price    = nums[3].replace(',', '.') + ' EUR'   # price after discount
                    total_price = nums[4].replace(',', '.') + ' EUR'  # net value
                except (ValueError, IndexError):
                    pass
            elif len(nums) >= 2:
                try:
                    qty_val  = float(nums[0].replace(',', '.'))
                    quantity = _gh_qty_label(qty_val, unit_str)
                    total_price = nums[-1].replace(',', '.') + ' EUR'
                except ValueError:
                    pass

        logger.info("G&H code=%s qty=%s price=%s total=%s", code, quantity, price, total_price)

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        seen.add(code)
        records.append(rec)

    logger.info("Geyer & Hosaja text extraction: %d records", len(records))
    return records


def extract_geyer_hosaja_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Geyer & Hosaja: %d table(s) received", len(tables))

    table_records: list[ProductRecord] = []
    seen: dict[str, list[int]] = {}
    for table in tables:
        for rec in _parse_geyer_hosaja_table(table):
            _smart_merge_or_add(table_records, seen, rec)

    text_records = _parse_geyer_hosaja_from_text(text) if text else []

    records = text_records if len(text_records) > len(table_records) else table_records
    logger.info("Geyer & Hosaja: %d records (table=%d text=%d)",
                len(records), len(table_records), len(text_records))
    return records


# ---------------------------------------------------------------------------
# PETEX
# ---------------------------------------------------------------------------

# Line starts with: {1-3 digit pos}  {7-9 digit article}  {rest}
_PETEX_LINE_RE = re.compile(r'^\s*\d{1,3}\s+(\d{7,13})\s+(.*)', re.UNICODE)
# Non-EAN charge lines: pos CODE desc... e.g. "3 EWP Einwegpaletten ..."
# Max 3-char code to exclude Petex annotation rows like "Text"
_PETEX_NONEAN_RE = re.compile(r'^\s*\d{1,3}\s+([A-Z][A-Z0-9]{1,2})\s+(.*)', re.IGNORECASE)


def _is_petex_document(text: str) -> bool:
    return bool(re.search(r'petex|artikelnummer|artikelbezeichnung', text, re.IGNORECASE))


def _petex_qty_label(n: float, meh: str) -> str:
    n_int = int(n) if n == int(n) else n
    if re.search(r'\bsa\.?', meh, re.IGNORECASE):
        return f"{n_int} {'Комплект' if n_int == 1 else 'Комплекта'}"
    return f"{n_int} {'Брой' if n_int == 1 else 'Броя'}"


def _petex_net_price(preis: str, rabatt_pct: str) -> str:
    """Return unit price after discount, rounded to 2 decimal places."""
    try:
        p = float(preis.replace(',', '.'))
        r = float(rabatt_pct.replace(',', '.')) / 100
        return f"{round(p * (1 - r), 2):.2f}"
    except ValueError:
        return preis.replace(',', '.')


def _petex_eur_num(s: str) -> str | None:
    """Parse a German number (European format: dot=thousands, comma=decimal) to '1234.56'."""
    s = (s or "").strip()
    if not s or not re.search(r'\d', s):
        return None
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return None


def _parse_petex_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Petex invoice.

    Columns: Pos. | Artikelnummer | Artikelbezeichnung | Menge | MEH | Preis | Rabatt | Betrag EUR
    """
    if not table or len(table) < 2:
        return []

    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if "artikelnummer" in joined or "artikelbezeichnung" in joined:
            header_idx = i
            break
    if header_idx is None:
        return []

    headers = [str(c or "").lower().strip() for c in table[header_idx]]

    def find(kws):
        for kw in kws:
            for idx, h in enumerate(headers):
                if kw in h:
                    return idx
        return None

    art_idx   = find(["artikelnummer", "artikel"])
    desc_idx  = find(["artikelbezeichnung", "bezeichnung"])
    qty_idx   = find(["menge"])
    meh_idx   = find(["meh"])
    price_idx = find(["preis"])
    rabatt_idx= find(["rabatt"])
    total_idx = find(["betrag"])

    if art_idx is None:
        return []

    records = []
    for row in table[header_idx + 1:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(ci):
            if ci is None or ci >= len(row):
                return ""
            return str(row[ci] or "").strip()

        code_raw = cell(art_idx)
        if not re.match(r'^\d{7,9}$', code_raw):
            continue

        name = cell(desc_idx) if desc_idx is not None else ""

        # Quantity + unit
        qty_raw = cell(qty_idx)
        meh_raw = cell(meh_idx)
        quantity = None
        qty_val  = None
        if qty_raw:
            qm = re.match(r'^(\d+)[,.](\d+)$', qty_raw)
            if qm:
                qty_val  = float(qty_raw.replace(',', '.'))
                quantity = _petex_qty_label(qty_val, meh_raw)

        # Net unit price = Preis × (1 - Rabatt%)
        preis_raw  = cell(price_idx)
        rabatt_raw = cell(rabatt_idx)
        price = None
        if preis_raw and re.search(r'\d', preis_raw):
            pct_m = re.search(r'(\d+[,.]\d+)', rabatt_raw) if rabatt_raw else None
            if pct_m:
                price = _petex_net_price(preis_raw, pct_m.group(1)) + ' EUR'
            else:
                price = preis_raw.replace(',', '.') + ' EUR'

        # Total (Betrag EUR)
        total_raw = cell(total_idx)
        total_price = (total_raw.replace(',', '.') + ' EUR') if total_raw and re.search(r'\d', total_raw) else None

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code_raw
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        records.append(rec)

    return records


def _parse_petex_product_table(table: list[list]) -> ProductRecord | None:
    """Parse a single-product Petex table (new format: one table per product, no header).

    Row 0: Pos | EAN-13 | Description | Menge | MEH | Preis | Rabatt | Betrag EUR
    Row 1+: '' | internal_code | description continuation | ...
    """
    if not table:
        return None
    row0 = table[0]
    if len(row0) < 7:
        return None

    def c(row, i):
        return str(row[i] or "").strip() if i < len(row) else ""

    pos = c(row0, 0)
    if not re.match(r'^\d{1,3}$', pos):
        return None
    ean_raw = c(row0, 1)
    is_ean = bool(re.match(r'^\d{7,13}$', ean_raw))
    is_nonean = bool(re.match(r'^[A-Z][A-Z0-9]{1,2}$', ean_raw, re.IGNORECASE))
    if not is_ean and not is_nonean:
        return None

    desc_parts = [c(row0, 2)]
    qty_raw   = c(row0, 3)
    unit_raw  = c(row0, 4)
    price_raw = c(row0, 5)
    rabatt_raw= c(row0, 6) if len(row0) > 6 else ""
    total_raw = c(row0, 7) if len(row0) > 7 else ""

    # Internal Petex code is in col 1 of the first continuation row (e.g. "13010")
    # Non-EAN items (EWP, HP, ...) don't have an internal code — use the code as-is.
    internal_code = None
    if is_ean:
        for extra in table[1:]:
            candidate = c(extra, 1)
            if candidate and re.match(r'^\d{4,9}$', candidate):
                internal_code = candidate
            part = c(extra, 2)
            if part:
                desc_parts.append(part)

    product_code = internal_code or ean_raw
    desc = " ".join(p for p in desc_parts if p).strip()

    qty_val = None
    try:
        qty_val = float(qty_raw.replace(',', '.'))
    except ValueError:
        pass
    quantity = _petex_qty_label(qty_val, unit_raw) if qty_val is not None else None

    price = None
    if price_raw and re.search(r'\d', price_raw):
        pct_m = re.search(r'(\d+[,.]\d+)', rabatt_raw) if rabatt_raw else None
        if pct_m:
            price = _petex_net_price(price_raw, pct_m.group(1)) + ' EUR'
        else:
            n = _petex_eur_num(price_raw)
            price = (n + ' EUR') if n else None

    n = _petex_eur_num(total_raw)
    total_price = (n + ' EUR') if n else None

    rec = ProductRecord(extraction_method="table")
    rec.product_code = product_code
    if re.match(r'^\d{13}$', ean_raw):
        rec.ean = ean_raw
    rec.product_name = desc[:120] if desc else None
    rec.quantity     = quantity
    rec.price        = price
    rec.total_price  = total_price
    logger.info("Petex product table: code=%s qty=%s price=%s total=%s | %s",
                product_code, quantity, price, total_price, desc[:60])
    return rec


def _parse_petex_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for Petex invoices."""
    records: list[ProductRecord] = []
    seen: set[str] = set()
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _PETEX_LINE_RE.match(line)
        is_nonean = False
        if not m:
            m = _PETEX_NONEAN_RE.match(line)
            if not m:
                i += 1
                continue
            is_nonean = True

        ean_raw = m.group(1)
        after = m.group(2).strip()

        if ean_raw in seen:
            i += 1
            continue

        # Collect continuation lines until next product
        extra_lines: list[str] = []
        j = i + 1
        while j < len(lines) and j <= i + 5:
            nl = lines[j].strip()
            if not nl:
                j += 1
                continue
            if _PETEX_LINE_RE.match(nl) or _PETEX_NONEAN_RE.match(nl):
                break
            extra_lines.append(nl)
            j += 1

        # Internal Petex code is the first 4-9 digit token on any continuation line
        # (non-EAN items like EWP/HP don't have an internal code)
        internal_code = None
        if not is_nonean:
            for nl in extra_lines:
                ic_m = re.match(r'^(\d{4,9})(?:\s|$)', nl)
                if ic_m:
                    internal_code = ic_m.group(1)
                    break
        product_code = internal_code or ean_raw

        search_text = line  # numeric extraction uses only the main product line

        # Build description — stop before quantity/numeric data block
        name_parts = [after] if after else []
        for nl in extra_lines:
            if re.search(r'\b\d+[,.]\d{2}\s*(?:sa\.|st\.?|stk)', nl, re.IGNORECASE):
                break
            if re.search(r'\b\d+[,.]\d+\s*%', nl):
                break
            name_parts.append(nl)
        name = " ".join(name_parts).strip()
        # Trim trailing numeric block that pdfplumber left in description
        name = re.sub(r'\s+\d+[,.]\d{2}\s+(?:sa\.|st\.?|stk).*$', '', name,
                      flags=re.IGNORECASE).strip()

        # Quantity: "5,00 sa." or "5,00 St."
        # No trailing \b — the dot in "sa." is non-word so \b would fail there.
        qty_m = re.search(r'\b(\d+[,.]\d{2})\s*(sa\.|st\.|stk\.?)(?=\s|$)',
                          search_text, re.IGNORECASE)
        quantity = None
        qty_val  = None
        if qty_m:
            qty_val  = float(qty_m.group(1).replace(',', '.'))
            quantity = _petex_qty_label(qty_val, qty_m.group(2))
            after_qty = search_text[qty_m.end():]
        else:
            after_qty = search_text

        # Preis + Rabatt% + Betrag after the unit marker
        # Use European number pattern: optional thousands dots + comma decimal (e.g. "1.294,56")
        _EUR_NUM_RE = r'\d{1,3}(?:[.]\d{3})*,\d{2}'
        price_block = re.findall(rf'\b({_EUR_NUM_RE})\b', after_qty)
        rabatt_m    = re.search(r'\b(\d+[,.]\d+)\s*%', after_qty)

        price = total_price = None
        if price_block:
            preis_str = price_block[0]
            betrag_str = price_block[-1] if len(price_block) >= 2 else None
            if rabatt_m:
                price = _petex_net_price(preis_str, rabatt_m.group(1)) + ' EUR'
            else:
                n = _petex_eur_num(preis_str)
                price = (n + ' EUR') if n else None
            if betrag_str and betrag_str != preis_str:
                n = _petex_eur_num(betrag_str)
                total_price = (n + ' EUR') if n else None

        logger.info("Petex code=%s qty=%s price=%s total=%s | %s",
                    product_code, quantity, price, total_price, search_text[:120])

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = product_code
        if re.match(r'^\d{13}$', ean_raw):
            rec.ean = ean_raw
        rec.product_name  = name[:120] if name else None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        seen.add(ean_raw)
        records.append(rec)
        i = j

    logger.info("Petex text extraction: %d records", len(records))
    return records


def extract_petex_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Petex: %d table(s) received", len(tables))

    table_records: list[ProductRecord] = []
    seen: dict[str, list[int]] = {}

    # Try classic parser first (old format: one big table with header + all products)
    for table in tables:
        for rec in _parse_petex_table(table):
            _smart_merge_or_add(table_records, seen, rec)

    # New format: one table per product (each table has pos|EAN|desc|qty|unit|price|rabatt|total)
    if not table_records:
        for table in tables:
            rec = _parse_petex_product_table(table)
            if rec:
                _smart_merge_or_add(table_records, seen, rec)

    n_table = len(table_records)
    text_records = _parse_petex_from_text(text) if text else []

    # Text parser reads lines sequentially → correct invoice order.
    # Table parser has reliable numeric data (clean cell values).
    # Strategy: text defines order; patch with table numerics when complete.
    table_by_code = {r.product_code: r for r in table_records}
    ordered: list[ProductRecord] = []
    seen_codes: set[str] = set()

    for r in text_records:
        t = table_by_code.get(r.product_code)
        if t:
            if t.quantity and t.total_price:
                r.quantity    = t.quantity
                r.price       = t.price
                r.total_price = t.total_price
            if t.ean and not r.ean:
                r.ean = t.ean
        ordered.append(r)
        seen_codes.add(r.product_code)

    # Append any table records the text parser didn't find
    for r in table_records:
        if r.product_code not in seen_codes:
            ordered.append(r)

    records = ordered
    logger.info("Petex: %d records (table_parser=%d text=%d)",
                len(records), n_table, len(text_records))
    return records


# ---------------------------------------------------------------------------
# Hakr (ASN HAKR Brno s.r.o.)
# ---------------------------------------------------------------------------

def _is_hakr_document(text: str) -> bool:
    return bool(re.search(r'hakr\s*brno|hakrbrno\.cz|ASN\s+HAKR', text, re.IGNORECASE))


def _parse_hakr_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Hakr invoice.

    Handles multi-line cells: pdfplumber sometimes splits a wrapped row into two
    physical rows — the code appears in the first sub-row (with empty numeric cols)
    and the qty/price/total appear in the continuation sub-row.

    Column layout: Description | Q'ty | Unit price | Discount | Price | %VAT | VAT | Total
    """
    if not table:
        return []

    ncols = max(len(r) for r in table) if table else 0
    logger.debug("Hakr: table %d rows × %d cols; first row: %r",
                 len(table), ncols, [str(c or "")[:40] for c in table[0]] if table else [])

    records: list[ProductRecord] = []
    pending_rec: Optional[ProductRecord] = None  # code found; waiting for numeric data
    pending_base: int = 0                         # desc_col of the pending record

    for row in table:
        if not any(str(c or "").strip() for c in row):
            continue

        # Find cell matching CODE:Name (allow optional space before colon, e.g. "HV2401 :")
        desc = ""
        desc_col: Optional[int] = None
        for ci, c in enumerate(row):
            s = re.sub(r'\s+', ' ', str(c or "")).strip()
            if re.match(r'^[A-Za-z]{1,8}\d{2,10}\s*:', s):
                desc = s
                desc_col = ci
                break

        def _cell(offset: int, base: Optional[int] = None) -> str:
            b = base if base is not None else (desc_col or 0)
            idx = b + offset
            if idx < 0 or idx >= len(row):
                return ""
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip()

        if desc and desc_col is not None:
            # Flush any pending code that had no data row following it
            if pending_rec is not None:
                records.append(pending_rec)
                logger.info("Hakr table: code=%s qty=%s total=%s",
                            pending_rec.product_code, pending_rec.quantity, pending_rec.total_price)
                pending_rec = None

            m = re.match(r'^([A-Za-z]{1,8}\d{2,10})\s*:', desc)
            if not m or not re.match(r'^[A-Za-z]{1,8}\d{2,10}$', m.group(1)):
                continue
            code_raw = m.group(1)
            name = desc[m.end():].strip()[:120] or None

            qty_m = re.match(r'(\d+)', _cell(1))
            qty = qty_m.group(1) if qty_m else None
            price_str = _kegel_num(_cell(2))
            total_str = _kegel_num(_cell(7)) or _kegel_num(_cell(6))

            rec = ProductRecord(extraction_method="table")
            rec.product_code = code_raw
            rec.product_name = name
            if qty:
                rec.quantity = qty
            if price_str:
                rec.price = price_str + ' EUR'
            if total_str:
                rec.total_price = total_str + ' EUR'

            if qty or total_str:
                records.append(rec)
                logger.info("Hakr table: code=%s qty=%s total=%s", code_raw, qty, total_str)
            else:
                # Numeric data not yet seen — may arrive in the next (continuation) row
                pending_rec = rec
                pending_base = desc_col

        elif pending_rec is not None:
            # Continuation row — pick up qty/price/total at the same column offsets
            qty_m = re.match(r'(\d+)', _cell(1, pending_base))
            qty = qty_m.group(1) if qty_m else None
            price_str = _kegel_num(_cell(2, pending_base))
            total_str = _kegel_num(_cell(7, pending_base)) or _kegel_num(_cell(6, pending_base))

            if qty:
                pending_rec.quantity = qty
            if price_str:
                pending_rec.price = price_str + ' EUR'
            if total_str:
                pending_rec.total_price = total_str + ' EUR'

            records.append(pending_rec)
            logger.info("Hakr table: code=%s qty=%s total=%s (from continuation row)",
                        pending_rec.product_code, qty, total_str)
            pending_rec = None

    if pending_rec is not None:
        records.append(pending_rec)
        logger.info("Hakr table: code=%s qty=%s total=%s",
                    pending_rec.product_code, pending_rec.quantity, pending_rec.total_price)

    return records


def _parse_hakr_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for Hakr invoices.

    Uses a single full-row regex: everything before the colon is the product code,
    everything after is the name until the Npcs data block.  This handles any code
    format (HV codes, mKayak L, adaptér, etc.) without enumerating patterns.

    Thousands commas in amounts (e.g. 1,370.00) are stripped before float conversion.
    """
    lines = [ln for ln in text.splitlines() if ln.strip()]
    records: list[ProductRecord] = []
    seen: set[str] = set()

    # Full product row in one pass:
    #   group 1 = code  (lazy: shortest text before the colon)
    #   group 2 = name  (lazy: shortest text before qty-pcs block)
    #   group 3 = qty
    #   group 4 = unit price
    #   group 5 = line price (= unit × qty, ignored)
    #   group 6 = total
    full_row_re = re.compile(
        r'^(.+?)\s*:(.+?)\s+(\d+)\s*pcs\s+([\d.,]+)\s+([\d.,]+)\s+\d+%\s+[\d.,]+\s+([\d.,]+)',
        re.IGNORECASE
    )

    def _to_float(s: str) -> Optional[float]:
        try:
            # Strip thousands commas: "1,370.00" → "1370.00" (decimal stays as-is)
            return float(s.replace(',', ''))
        except (ValueError, AttributeError):
            return None

    for i, line in enumerate(lines):
        m = full_row_re.match(line.strip())
        if not m:
            continue

        code = m.group(1).strip().upper()
        # Normalize variant suffix: "HV1122 - E" → "HV1122E"
        code = re.sub(r'\s+-\s+', '', code)
        # Explicit remaps: invoice code → catalog code
        code = {
            'MKAYAK L': 'HVMKAYAK E',
            'MKAYAK M': 'MKAYAK',
            'ADAPTÉR': '7-13',
        }.get(code, code)
        name = m.group(2).strip()[:120] or None
        qty = m.group(3)
        up = _to_float(m.group(4))
        tot = _to_float(m.group(6))

        if code in seen:
            continue
        seen.add(code)

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name
        if qty:
            rec.quantity = qty
        if up is not None:
            rec.price = f"{up:.2f} EUR"
        if tot is not None:
            rec.total_price = f"{tot:.2f} EUR"
        records.append(rec)
        logger.info("Hakr text: code=%s qty=%s price=%s total=%s", code, qty, up, tot)

    return records


def extract_hakr_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Hakr: %d table(s) received", len(tables))
    raw: list[ProductRecord] = []
    for table in tables:
        raw.extend(_parse_hakr_table(table))

    seen: dict[str, list[int]] = {}
    records: list[ProductRecord] = []
    for rec in raw:
        _smart_merge_or_add(records, seen, rec)

    # Only trust table results when at least some records have qty AND total
    if records and any(r.quantity and r.total_price for r in records):
        logger.info("Hakr: %d records from tables", len(records))
        return records

    if records:
        logger.info("Hakr: %d incomplete table records — falling back to text", len(records))

    if text:
        text_records = _parse_hakr_from_text(text)
        if text_records:
            logger.info("Hakr text extraction: %d records", len(text_records))
            return text_records

    if records:
        logger.info("Hakr: returning %d incomplete table records (no text match)", len(records))
    return records


# ---------------------------------------------------------------------------
# AutoMania
# ---------------------------------------------------------------------------

def _is_automania_document(text: str) -> bool:
    return bool(re.search(r'автомания|automania', text, re.IGNORECASE))


def _parse_automania_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from an AutoMania invoice.

    Columns: № | Наименование | Код | Мярка | Колич. | Ед.цена EUR | Стойност EUR
    """
    if not table or len(table) < 2:
        return []

    records = []
    for row in table:
        if not any(str(c or "").strip() for c in row):
            continue
        if len(row) < 5:
            continue

        def cell(idx):
            if idx >= len(row):
                return ""
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip()

        # Skip header rows
        code_raw = cell(2)
        if not code_raw or re.search(r'код|code|наименование', code_raw, re.IGNORECASE):
            continue
        # Skip rows where code is just a number that looks like a row number or price
        if not re.match(r'^(?:\d{4,6}|[A-Za-z]{2,4}\s*\d{2,4})$', code_raw):
            continue

        name      = cell(1) or None
        qty       = cell(4) or None
        price_str = _kegel_num(cell(5))
        total_str = _kegel_num(cell(6))

        rec = ProductRecord(extraction_method="table")
        rec.product_code = f"AVM-{code_raw}"
        rec.product_name = name[:120] if name else None
        if qty:
            rec.quantity = qty
        if price_str:
            rec.price = price_str + ' EUR'
        if total_str:
            rec.total_price = total_str + ' EUR'

        records.append(rec)
        logger.info("AutoMania table: code=%s qty=%s total=%s", rec.product_code, qty, total_str)

    return records


def _parse_automania_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for AutoMania invoices."""
    records = []
    seen_codes: set[str] = set()

    # Match row pattern: row_num  product_name  code  unit  qty  price  total
    row_re = re.compile(
        r'\b(\d{1,3})\s+'                          # row number
        r'(.+?)\s+'                                 # product name (non-greedy)
        r'(\d{4,6}|[A-Z]{2,4}\s*\d{2,4})\s+'      # code
        r'(?:БР|PCS|SET|бр)[.\s]+'                 # unit
        r'(\d+)\s+'                                 # qty
        r'([\d.]+)\s+'                              # unit price
        r'([\d.]+)',                                # total
        re.IGNORECASE
    )
    for m in row_re.finditer(text):
        code_raw = re.sub(r'\s+', ' ', m.group(3)).strip()
        code = f"AVM-{code_raw}"
        if code in seen_codes:
            continue
        seen_codes.add(code)

        name = re.sub(r'\s+', ' ', m.group(2)).strip()[:120]
        qty  = m.group(4)
        try:
            price_str = f"{float(m.group(5)):.2f}"
        except ValueError:
            price_str = None
        try:
            total_str = f"{float(m.group(6)):.2f}"
        except ValueError:
            total_str = None

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name or None
        if qty:
            rec.quantity = qty
        if price_str:
            rec.price = price_str + ' EUR'
        if total_str:
            rec.total_price = total_str + ' EUR'

        records.append(rec)
        logger.info("AutoMania text: code=%s qty=%s total=%s", code, qty, total_str)

    return records


def extract_automania_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("AutoMania: %d table(s) received", len(tables))
    raw: list[ProductRecord] = []
    for table in tables:
        raw.extend(_parse_automania_table(table))

    # Deduplicate by code — PDF contains both original and copy of the invoice
    seen: set[str] = set()
    records = []
    for rec in raw:
        if rec.product_code not in seen:
            seen.add(rec.product_code)
            records.append(rec)

    if records:
        logger.info("AutoMania: %d records from tables (after dedup)", len(records))
        return records

    if text:
        records = _parse_automania_from_text(text)
        logger.info("AutoMania text extraction: %d records", len(records))
    return records


# ---------------------------------------------------------------------------
# Kegel-Błażusiak
# ---------------------------------------------------------------------------

_KEGEL_CODE_RE = re.compile(r'(\d[\-–—]\d{4}[\-–—]\d{3}[\-–—]\d{4})')


def _is_kegel_blazusiak_document(text: str) -> bool:
    return bool(re.search(r'kegel[-\s]?b[łl]azusiak|/KBT/', text, re.IGNORECASE))


def _kegel_num(s: str) -> str | None:
    s = (s or "").strip().replace(',', '.')
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return None


def _parse_kegel_blazusiak_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Kegel-Błażusiak invoice.

    Columns: No. | Item/Artikel (PL name + code + EN name) | Custom code (ignored) |
             Unit | Quantity | Unit price EUR | Amount EUR
    """
    if not table or len(table) < 2:
        return []

    records = []
    for row in table:
        if not any(str(c or "").strip() for c in row):
            continue

        # Find the item cell that contains the product code (5-XXXX-XXX-XXXX)
        item_text = ""
        for c in row:
            s = str(c or "")
            if _KEGEL_CODE_RE.search(s):
                item_text = s
                break
        if not item_text:
            continue

        code_m = _KEGEL_CODE_RE.search(item_text)
        if not code_m:
            continue
        code = code_m.group(1)

        # Name = lines above the code line inside the item cell
        lines = [ln.strip() for ln in item_text.splitlines() if ln.strip()]
        code_line_idx = next((i for i, ln in enumerate(lines) if _KEGEL_CODE_RE.search(ln)), None)
        if code_line_idx is not None and code_line_idx > 0:
            name = " ".join(lines[:code_line_idx])[:120]
        else:
            raw = item_text[:code_m.start()].strip()
            name = re.sub(r'\s+', ' ', raw)[:120] or None

        # Scan from right to collect total, price, qty.
        # Works for both the 7-col PDF layout and the wide (30-col) Excel export.
        numerics: list[str] = []
        for c in reversed(row):
            v = re.sub(r'\s+', ' ', str(c or '')).strip()
            if v and _kegel_num(v) is not None:
                numerics.append(v)
                if len(numerics) == 3:
                    break

        total_str = numerics[0] if len(numerics) >= 1 else None
        price_str = numerics[1] if len(numerics) >= 2 else None
        qty       = numerics[2] if len(numerics) >= 3 else None

        # Sanity check: qty > 9999 is almost certainly a misread customs/tariff code
        # (e.g. 630790980). Discard it and recover from total / price instead.
        if qty:
            try:
                _qv = float(qty.replace(',', '.'))
                if _qv > 9999 or _qv != int(_qv):
                    logger.warning("Kegel: implausible qty=%s for code=%s — discarding", qty, code)
                    qty = None
            except ValueError:
                qty = None

        # qty was discarded — attempt recovery in order of reliability:
        _qty_recovered = False

        # 1. price > total is physically impossible (unit_price × qty = total).
        #    When this occurs, pdfplumber put the qty value in the price column.
        #    Swap: real qty = price_str (as integer), real price = total / qty.
        if not _qty_recovered and price_str and total_str:
            try:
                _pv = float(price_str.replace(',', '.'))
                _tv = float(total_str.replace(',', '.'))
                if _pv > _tv and 0 < _pv <= 9999 and _pv == int(_pv):
                    _swap_qty = int(_pv)
                    qty = str(_swap_qty)
                    # total_str is actually the unit price; compute the real total
                    price_str = total_str
                    total_str = f"{_swap_qty * _tv:.2f}".replace('.', ',')
                    _qty_recovered = True
                    logger.info("Kegel: price>total swap → qty=%s price=%s total=%s for code=%s",
                                qty, price_str, total_str, code)
            except (ValueError, ZeroDivisionError):
                pass

        # 2. price == total: pdfplumber merged columns; scan item cell for
        #    a data line of the form: [customs_code] [unit] QTY PRICE TOTAL
        if not _qty_recovered and total_str and price_str and \
                price_str.replace(',', '.') == total_str.replace(',', '.'):
            _total_esc = re.escape(total_str)
            for _ln in lines:
                if _KEGEL_CODE_RE.search(_ln):
                    continue
                _m = re.search(r'\b(\d{1,4})\s+([\d]+[,.][\d]+)\s+' + _total_esc + r'(?:\s|$)', _ln)
                if _m:
                    try:
                        _qv2 = int(_m.group(1))
                        _pv2 = float(_m.group(2).replace(',', '.'))
                        _tv2 = float(total_str.replace(',', '.'))
                        if _qv2 > 0 and abs(_qv2 * _pv2 - _tv2) <= 0.02 * _tv2 + 0.01:
                            qty = str(_qv2)
                            price_str = _m.group(2)
                            _qty_recovered = True
                            logger.info("Kegel: extracted qty=%s price=%s from item cell for code=%s",
                                        qty, price_str, code)
                            break
                    except (ValueError, ZeroDivisionError):
                        pass

        # 3. Fallback: compute qty = round(total / price)
        if not _qty_recovered and total_str and price_str:
            try:
                _tv = float(total_str.replace(',', '.'))
                _pv = float(price_str.replace(',', '.'))
                if _pv > 0:
                    _calc = round(_tv / _pv)
                    if _calc > 0:
                        qty = str(_calc)
                        logger.info("Kegel: recovered qty=%s from total/price for code=%s", qty, code)
            except (ValueError, ZeroDivisionError):
                pass

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = (re.sub(r'\s+', ' ', name).strip() or None) if name else None
        if qty:
            rec.quantity = qty
        if price_str:
            rec.price = price_str + ' EUR'
        if total_str:
            rec.total_price = total_str + ' EUR'

        records.append(rec)
        logger.info("Kegel table: code=%s qty=%s total=%s", code, qty, total_str)

    return records


def _parse_kegel_blazusiak_from_text(text: str) -> list[ProductRecord]:
    """Text fallback for Kegel-Błażusiak when pdfplumber finds no usable tables."""
    records = []
    seen_codes: set[str] = set()

    logger.info("Kegel text fallback — first 600 chars:\n%s", repr(text[:600]))
    code_matches = list(_KEGEL_CODE_RE.finditer(text))
    if not code_matches:
        logger.warning("Kegel: no code matches found (pattern=%s)", _KEGEL_CODE_RE.pattern)
        return records

    logger.info("Kegel text: found %d code occurrences", len(code_matches))

    for i, cm in enumerate(code_matches):
        code = cm.group(1)
        if code in seen_codes:
            continue
        seen_codes.add(code)

        seg_end = code_matches[i + 1].start() if i + 1 < len(code_matches) else len(text)
        after   = text[cm.end():seg_end].strip()

        # Polish name is on the line immediately before the code
        before_lines = [ln.strip() for ln in text[:cm.start()].splitlines() if ln.strip()]
        name = before_lines[-1][:120] if before_lines else None

        qty = price_str = total_str = None
        decimals = [float(n.replace(',', '.')) for n in re.findall(r'\d+[.,]\d+', after)]
        integers = [int(n) for n in re.findall(r'\b(\d{1,4})\b', after)]

        if integers:
            qty = str(integers[0])
        if len(decimals) >= 2:
            price_str = f"{decimals[0]:.2f}"
            total_str = f"{decimals[1]:.2f}"
        elif len(decimals) == 1:
            total_str = f"{decimals[0]:.2f}"
            if qty:
                try:
                    price_str = f"{decimals[0] / int(qty):.2f}"
                except (ValueError, ZeroDivisionError):
                    pass

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name
        if qty:
            rec.quantity = qty
        if price_str:
            rec.price = price_str + ' EUR'
        if total_str:
            rec.total_price = total_str + ' EUR'

        records.append(rec)
        logger.info("Kegel text: code=%s qty=%s total=%s", code, qty, total_str)

    return records


def extract_kegel_blazusiak_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Kegel: %d table(s) received", len(tables))
    records = []
    for table in tables:
        records.extend(_parse_kegel_blazusiak_table(table))
    if records:
        logger.info("Kegel: %d records from tables", len(records))
        return records

    if text:
        records = _parse_kegel_blazusiak_from_text(text)
        logger.info("Kegel text extraction: %d records", len(records))
    return records


# ---------------------------------------------------------------------------
# ToM-PaR
# ---------------------------------------------------------------------------

_TOMPAR_CODE_RE = re.compile(r'\b(TP\d{6})\b', re.IGNORECASE)


def _is_tompar_document(text: str) -> bool:
    return bool(re.search(r'tom[-\s]?par|NIP[:\s]+1182250133', text, re.IGNORECASE))


def extract_tompar_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("ToMPaR: %d table(s) received", len(tables))
    records = _parse_tompar_from_text(text)
    logger.info("ToMPaR text extraction: %d records", len(records))
    return records


def _parse_tompar_from_text(text: str) -> list[ProductRecord]:
    """
    ToM-PaR invoice layout: each product has a data line where the Lp number is
    glued directly to the name (e.g. '1Szczotka Do Mycia Nice 18 szt. 3,11 ...').
    The TP###### code appears on a separate line 1-4 lines below.
    """
    records: list[ProductRecord] = []
    seen: dict[str, list[int]] = {}
    lines = text.splitlines()

    # Data line: Lp+Name QTY szt. price_no_disc rabat unit_price net St% tax brutto
    # Columns: (Lp)(Name) QTY szt. price1 0,00 price2 net 0 0,00 brutto
    data_line_re = re.compile(
        r'^(\d+)(.+?)\s+(\d+)\s+szt\.\s+'
        r'([\d,]+)\s+[\d,]+\s+[\d,]+\s+[\d,]+\s+\d+\s+[\d,]+\s+([\d,]+)',
        re.IGNORECASE
    )

    def _to_float(s: str) -> float | None:
        try:
            return float(s.replace(',', '.'))
        except ValueError:
            return None

    for i, line in enumerate(lines):
        m = _TOMPAR_CODE_RE.search(line)
        if not m:
            continue
        code = m.group(1).upper()

        # Find closest preceding data line (up to 10 lines back)
        data_match = None
        data_idx = i
        for j in range(i - 1, max(i - 10, -1), -1):
            dm = data_line_re.match(lines[j])
            if dm:
                data_match = dm
                data_idx = j
                break

        qty = None
        unit_price = None
        total = None
        name_parts = []

        if data_match:
            qty = data_match.group(3)
            unit_price = _to_float(data_match.group(4))  # cena netto bez rabatu
            total = _to_float(data_match.group(5))        # wartość brutto
            name_parts.append(data_match.group(2).strip())
            # Name continuation lines between data line and code line
            for j in range(data_idx + 1, i):
                part = lines[j].strip()
                if part and not _TOMPAR_CODE_RE.search(part):
                    name_parts.append(part)

        name = " ".join(name_parts).strip()[:150] or None

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name
        if qty:
            rec.quantity = qty
        if unit_price is not None:
            rec.price = f"{unit_price:.2f} EUR"
        if total is not None:
            rec.total_price = f"{total:.2f} EUR"

        _smart_merge_or_add(records, seen, rec)
        logger.info("ToMPaR: code=%s qty=%s price=%s total=%s name=%r",
                    code, qty, unit_price, total, name)

    return records


# ---------------------------------------------------------------------------
# Sonax / Сенакс ООД
# ---------------------------------------------------------------------------

# 8-digit Sonax article code with optional suffix: 04172000.01 or 03234000-544
_SENAX_CODE_RE = re.compile(r'\b(\d{8}(?:[.\-]\d+)?)\b')


def _is_senax_document(text: str) -> bool:
    return bool(re.search(r'сенакс|senax|СЕНАКС|BG831566723', text, re.IGNORECASE))


def _senax_num(s: str) -> str | None:
    s = (s or "").strip().replace(' ', '').replace(',', '.')
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return None


def _parse_senax_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Senax/Sonax invoice.

    Expected columns: No. | Наименование (code + dash + name) | Кол. | Мярка | Ед. цена | Стойност
    Returns [] when pdfplumber has merged multiple rows into one cell.
    """
    if not table or len(table) < 2:
        return []

    records = []
    for row in table:
        if not any(str(c or "").strip() for c in row):
            continue

        # Find cell containing 8-digit code
        desc_text = ""
        desc_idx = -1
        for i, c in enumerate(row):
            s = str(c or "")
            if _SENAX_CODE_RE.search(s):
                desc_text = s
                desc_idx = i
                break
        if not desc_text:
            continue

        # If the cell contains more than one code, pdfplumber merged multiple rows — skip.
        # Return None (not []) so the caller knows to prefer text fallback.
        if len(_SENAX_CODE_RE.findall(desc_text)) > 1:
            logger.info("Senax: merged table cell detected — will use text fallback")
            return None

        m = _SENAX_CODE_RE.search(desc_text)
        if not m:
            continue

        code = m.group(1)
        after_code = desc_text[m.end():].strip()
        name = re.sub(r'^[\s\-–—]+', '', after_code).strip() or None
        if name:
            name = re.sub(r'\s+', ' ', name)[:150]

        def cell(idx):
            if idx < 0 or idx >= len(row):
                return ""
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip()

        qty       = cell(desc_idx + 1) or None
        price_str = _senax_num(cell(desc_idx + 3))
        total_str = _senax_num(cell(desc_idx + 4))

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name
        if qty:
            rec.quantity = qty
        if price_str:
            rec.price = price_str + " лв."
        if total_str:
            rec.total_price = total_str + " лв."

        records.append(rec)
        logger.info("Senax table: code=%s name=%r qty=%s total=%s", code, name, qty, total_str)

    return records


def _parse_senax_from_text(text: str) -> list[ProductRecord]:
    """Chunk-based text parser: splits text at each 8-digit code boundary.

    For each code, takes the text up to the next code and looks for:
    qty + unit (бр) + price + total at the end of the chunk.
    """
    records: list[ProductRecord] = []
    seen: dict[str, list[int]] = {}

    logger.info("Senax text fallback — first 500 chars:\n%s", text[:500])

    # When the document contains both a warehouse receipt (складова разписка)
    # and an invoice (фактура), restrict parsing to the invoice section only
    # to avoid duplicating quantities.
    _invoice_m = re.search(
        r'(?:данъчна\s+фактура|фактура\s*№|фактура\s*no|invoice)',
        text, re.IGNORECASE
    )
    if _invoice_m:
        text = text[_invoice_m.start():]
        logger.info("Senax: invoice section starts at char %d — warehouse receipt skipped",
                    _invoice_m.start())
    else:
        logger.info("Senax: no invoice section marker found — parsing full text")

    # All positions of 8-digit codes in the full text
    code_positions = [(m.start(), m.group(1)) for m in _SENAX_CODE_RE.finditer(text)]
    if not code_positions:
        logger.warning("Senax: no 8-digit codes found in text")
        return records

    qty_price_re = re.compile(
        r'(\d+(?:[,\.]\d+)?)\s+(?:бр\.?|pcs\.?|szt\.?|set)((?:\s+[\d,\.]+){2,4})',
        re.IGNORECASE,
    )

    for idx, (pos, code) in enumerate(code_positions):

        # Stop when we see a code we've already parsed — signals the start of a
        # duplicate section (e.g. warehouse receipt appended after the invoice).
        if code in seen:
            logger.info("Senax: repeated code %s at pos %d — stopping (duplicate section)", code, pos)
            break

        # Chunk: from this code to the next code occurrence (or end of text)
        end_pos = code_positions[idx + 1][0] if idx + 1 < len(code_positions) else len(text)
        chunk = text[pos:end_pos]

        # Strip the code (incl. optional suffix like .01 or -544) and the " - " separator
        after_code = re.sub(r'^\d{8}(?:[.\-]\d+)?\s*[-–—]\s*', '', chunk).strip()

        pm = qty_price_re.search(after_code)
        if code == '03325000':
            logger.info("DEBUG 03325000: after_code=%r pm=%s", after_code, bool(pm))
        if pm:
            name_raw = after_code[:pm.start()].strip()
            qty = pm.group(1)
            # Extract numbers from the tail, handling space-as-thousands-separator:
            # "1 010,88" must be treated as one number (1010.88), not split into two.
            # Pattern: 1-3 digits, optionally followed by groups of (space + 3 digits),
            # optionally followed by decimal part.
            _num_re = re.compile(r'\d{1,3}(?:\s\d{3})*(?:[,\.]\d+)?')
            nums = [_senax_num(t) for t in _num_re.findall(pm.group(2))]
            nums = [n for n in nums if n]
            if code == '03325000':
                logger.info("DEBUG 03325000: group2=%r findall=%r nums=%r",
                            pm.group(2), _num_re.findall(pm.group(2)), nums)
            # The Sonax invoice always ends with (final_price, total) — last two items.
            # Handles layouts: [price, total], [price, price, total],
            # [unit, disc%, price, total], etc.
            if len(nums) >= 2:
                price_str = nums[-2]   # Ед. цена с ТО
                total_str = nums[-1]   # Стойност
            else:
                price_str = total_str = None
        else:
            name_raw  = after_code
            qty = price_str = total_str = None

        name = re.sub(r'\s+', ' ', name_raw).strip()[:150] or None

        rec = ProductRecord(extraction_method="text")
        rec.product_code = code
        rec.product_name = name
        if qty:
            rec.quantity = qty
        if price_str:
            rec.price = price_str + " лв."
        if total_str:
            rec.total_price = total_str + " лв."
        _smart_merge_or_add(records, seen, rec)
        logger.info("Senax text: code=%s name=%r qty=%s price=%s total=%s",
                    code, name, qty, price_str, total_str)

    return records


_SENAX_CODE_REMAP = {
    '02063000': '02063000-544',
}


def extract_senax_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Senax: %d table(s) received", len(tables))
    records = []
    had_merged = False
    for table in tables:
        result = _parse_senax_table(table)
        if result is None:
            had_merged = True  # merged cells detected — text fallback preferred
        else:
            records.extend(result)

    if records and not had_merged:
        logger.info("Senax: %d records from tables", len(records))
    elif text:
        text_records = _parse_senax_from_text(text)
        logger.info("Senax text extraction: %d records", len(text_records))
        if text_records:
            records = text_records
        else:
            logger.info("Senax: %d records from tables (text fallback empty)", len(records))

    for rec in records:
        if rec.product_code in _SENAX_CODE_REMAP:
            rec.product_code = _SENAX_CODE_REMAP[rec.product_code]

    return records


# ---------------------------------------------------------------------------
# Heko / Team Heko
# ---------------------------------------------------------------------------

def _is_heko_document(text: str) -> bool:
    return bool(re.search(r'\bHEKO\b|СПЕСИФИКАЦИЯ\s+КЪМ\s+ПОРЪЧКА|team[\s\-]?heko', text, re.IGNORECASE))


def _heko_num(s: str) -> str | None:
    s = (s or "").strip().replace(' ', '').replace(',', '.')
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return None


def _parse_heko_table(table: list[list]) -> list[ProductRecord]:
    """Parse Heko order specification Excel table.

    Columns (0-based): A=name, B=code, C=price_vat, D=price_net, E=qty, F=total_net
    Header rows and empty rows are skipped.
    Code gets HK- prefix.
    """
    if not table:
        return []

    # Diagnostic: log first 3 rows to understand structure
    for i, row in enumerate(table[:3]):
        logger.info("Heko table row[%d]: %s", i, [str(c or "")[:40] for c in row])

    records = []
    for row in table:
        if len(row) < 5:
            continue

        def cell(i):
            return re.sub(r'\s+', ' ', str(row[i] or "")).strip() if i < len(row) else ""

        code_raw = cell(1)
        # Excel may read numeric codes as floats: 10109.0 → 10109
        if re.match(r'^\d+\.0$', code_raw):
            code_raw = code_raw[:-2]
        # Skip header / empty rows — code must be numeric
        if not re.match(r'^\d+$', code_raw):
            continue

        name      = cell(0) or None
        code      = "HK-" + code_raw
        price_str = _heko_num(cell(3))   # Цена без ДДС
        qty_str   = cell(4) or None
        total_str = _heko_num(cell(5))   # Ст-ст без ДДС

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name
        if qty_str:
            rec.quantity = qty_str
        if price_str:
            rec.price = price_str + " лв."
        if total_str:
            rec.total_price = total_str + " лв."

        records.append(rec)
        logger.info("Heko table: code=%s name=%r qty=%s total=%s", code, name, qty_str, total_str)

    return records


def extract_heko_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Heko: %d table(s) received", len(tables))
    records = []
    for table in tables:
        records.extend(_parse_heko_table(table))
    if records:
        logger.info("Heko: %d records from tables", len(records))
    return records


# ---------------------------------------------------------------------------
# BMW Group / BMW Service
# Columns: Номенкл.Номер | Описание | Мярка | Кол. | Ед.Цена | Отст | Ст-ст BGN | Ст-ст EUR
# ---------------------------------------------------------------------------

_BMW_CODE_RE = re.compile(r'^\d{11}$')


def _is_bmw_document(text: str) -> bool:
    return bool(re.search(r'BMW\s+Service|BMW\s+Дилър|BMW\s+Dealer', text, re.IGNORECASE)) or \
           bool(re.search(r'BMW', text, re.IGNORECASE) and
                re.search(r'Номенкл\.?Номер|Номенкл\b', text, re.IGNORECASE))


def _parse_bmw_table(table: list[list]) -> list[ProductRecord]:
    if not table or len(table) < 2:
        return []

    # Find header row
    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(re.sub(r'\s+', ' ', str(c or "")).lower() for c in row)
        if "номенкл" in joined and ("кол" in joined or "цена" in joined):
            header_idx = i
            break

    if header_idx is not None:
        headers = [re.sub(r'\s+', ' ', str(c or "")).lower().strip()
                   for c in table[header_idx]]

        def find(kws):
            for kw in kws:
                for idx, h in enumerate(headers):
                    if kw in h:
                        return idx
            return None

        code_idx  = find(["номенкл"])          or 0
        name_idx  = find(["описание", "desc"])  or 1
        unit_idx  = find(["мярка", "мярк"])
        qty_idx   = find(["кол"])
        price_idx = find(["ед.цена", "ед цена", "unit price", "цена"])
        disc_idx  = find(["отст", "discount", "отстъпка", "rabat"])
        eur_idx   = find(["ст-ст eur", "eur"])
        bgn_idx   = find(["ст-ст bgn", "bgn"])
        data_start = header_idx + 1
    else:
        # Continuation page — fixed layout
        code_idx  = 0
        name_idx  = 1
        unit_idx  = 2
        qty_idx   = 3
        price_idx = 4
        disc_idx  = 5
        eur_idx   = 7
        bgn_idx   = 6
        data_start = 0

    records = []
    for row in table[data_start:]:
        if not any(str(c or "").strip() for c in row):
            continue

        def cell(ci):
            if ci is None or ci >= len(row):
                return ""
            return re.sub(r'\s+', ' ', str(row[ci] or "")).strip()

        code = cell(code_idx)
        if not _BMW_CODE_RE.match(code):
            continue

        name     = cell(name_idx)
        unit_raw = cell(unit_idx) if unit_idx is not None else "бр."
        unit     = unit_raw if unit_raw else "бр."

        qty_raw = cell(qty_idx) if qty_idx is not None else ""
        quantity = None
        if qty_raw and re.match(r'^\d+$', qty_raw):
            quantity = f"{qty_raw} {unit}"

        def _num(ci):
            raw = cell(ci).replace(',', '.').replace('\xa0', '').replace(' ', '')
            return raw if raw and re.match(r'^\d+(?:\.\d+)?$', raw) else None

        # Calculate net unit price after discount
        gross_val = _num(price_idx) if price_idx is not None else None
        disc_val  = _num(disc_idx)  if disc_idx  is not None else None
        if gross_val and disc_val:
            try:
                net = float(gross_val) * (1 - float(disc_val) / 100)
                price = f"{net:.2f} EUR"
            except (ValueError, ZeroDivisionError):
                price = f"{gross_val} EUR"
        elif gross_val:
            price = f"{gross_val} EUR"
        else:
            price = None

        total_val = _num(eur_idx) if eur_idx is not None else None
        if total_val is None:
            total_val = _num(bgn_idx) if bgn_idx is not None else None
            total_currency = "BGN"
        else:
            total_currency = "EUR"
        total_price = f"{total_val} {total_currency}" if total_val else None

        rec = ProductRecord(extraction_method="table")
        rec.product_code  = code
        rec.product_name  = name or None
        rec.quantity      = quantity
        rec.price         = price
        rec.total_price   = total_price
        records.append(rec)

    return records


def extract_bmw_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("BMW: %d table(s) received", len(tables))
    records = []
    for table in tables:
        records.extend(_parse_bmw_table(table))
    if records:
        logger.info("BMW: %d records from tables", len(records))
    return records


# ---------------------------------------------------------------------------
# Areon (distributor: Ареон България ЕООД)
# ---------------------------------------------------------------------------
# Xado (Tuning Oils Club / ТУНИНГ ОЙЛС КЛУБ ЕООД)
# Bulgarian invoice: # | Вид | Основание и предмет на сделката | мярка | колич. | ед.цена | EUR | BGN
# No product codes in invoice — descriptions matched via xado-code-map.xlsx
# ---------------------------------------------------------------------------

def _is_xado_document(text: str) -> bool:
    return bool(re.search(r'тунинг\s+ойлс\s+клуб|tuning\s+oils\s+club|xado', text, re.IGNORECASE))


def _xado_num(s: str) -> str | None:
    s = re.sub(r'\s+', '', (s or "").replace('\xa0', ''))
    if not s:
        return None
    s = s.replace(',', '.')
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return None


# Matches text rows: rownum M description бр qty unit_price total_eur
_XADO_ROW_RE = re.compile(
    r'(\d{1,3})\s+'        # row number
    r'M\s+'                # Вид = M
    r'(.+?)\s+'            # description (lazy)
    r'бр\s+'               # мярка
    r'(\d+)\s+'            # quantity
    r'([\d,.]+)\s+'        # unit price
    r'([\d,.]+)',          # total EUR
    re.IGNORECASE | re.DOTALL,
)


def _parse_xado_table(table: list[list]) -> list[ProductRecord]:
    if not table:
        return []

    header_idx = desc_idx = qty_idx = price_idx = total_idx = None

    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if re.search(r'основание|предмет на сделката', joined):
            header_idx = i
            headers = [re.sub(r'\s+', ' ', str(c or "")).lower().strip() for c in row]

            def _find(kws):
                for kw in kws:
                    for j, h in enumerate(headers):
                        if kw in h:
                            return j
                return None

            desc_idx  = _find(["основание", "предмет"])
            qty_idx   = _find(["колич"])
            # Search for unit price: avoid matching "ед. мярка" (unit of measure)
            price_idx = _find(["ед.цена", "единична цена", "ед. цена"])
            if price_idx is None:
                price_idx = next((j for j, h in enumerate(headers) if "цена" in h and "мярка" not in h), None)
            total_idx = _find(["eur"])
            # Fallback: pdfplumber may label the column "СТОЙНОСТ" without "EUR"
            if total_idx is None and price_idx is not None:
                for j in range(price_idx + 1, len(headers)):
                    if "стойност" in headers[j]:
                        total_idx = j
                        break
            logger.info("Xado table headers: %s", headers)
            break

    if header_idx is None:
        return []

    logger.info("Xado table header at row %d: desc=%s qty=%s price=%s total=%s",
                header_idx, desc_idx, qty_idx, price_idx, total_idx)

    records = []
    for row in table[header_idx + 1:]:
        def cell(idx):
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip() if idx is not None and idx < len(row) else ""

        desc = cell(desc_idx)
        # Skip summary/footer rows; do NOT skip "total" — Xado product names contain it (e.g. "Total Flush")
        if not desc or re.search(r'общо|словом|данък|ддс|vat|дан\.?\s*основа|сума за плащане|дс:', desc.lower()):
            continue

        qty_raw   = cell(qty_idx)
        qty_clean = re.sub(r'\s*бр\.?\s*$', '', qty_raw, flags=re.IGNORECASE).strip()
        qty_clean = re.sub(r'^[^\d]*', '', qty_clean).strip()  # strip leading non-digits (e.g. "И 2" → "2")
        pv = _xado_num(cell(price_idx))
        tv = _xado_num(cell(total_idx))

        quantity = None
        qty_n = None
        if qty_clean:
            try:
                qty_n = int(float(qty_clean.replace(',', '.')))
                quantity = f"{qty_n} {'Брой' if qty_n == 1 else 'Броя'}"
            except ValueError:
                quantity = qty_clean

        # Derive missing price or total from the other when qty is known
        if qty_n and qty_n > 0:
            if pv is None and tv is not None:
                try:
                    pv = f"{float(tv) / qty_n:.2f}"
                except (ValueError, ZeroDivisionError):
                    pass
            elif tv is None and pv is not None:
                try:
                    tv = f"{float(pv) * qty_n:.2f}"
                except ValueError:
                    pass

        rec = ProductRecord(extraction_method="table")
        rec.product_code = re.sub(r'\s+', ' ', desc).strip()
        rec.quantity     = quantity
        rec.price        = (pv + " EUR") if pv else None
        rec.total_price  = (tv + " EUR") if tv else None
        records.append(rec)
        logger.info("Xado table: desc=%s qty=%s price=%s total=%s", desc[:50], quantity, pv, tv)

    return records


def _parse_xado_text(text: str) -> list[ProductRecord]:
    records = []
    logger.info("Xado text fallback — first 400 chars:\n%s", repr(text[:400]))
    for m in _XADO_ROW_RE.finditer(text):
        _pos, desc, qty_raw, price_raw, total_raw = m.groups()
        desc = re.sub(r'\s+', ' ', desc).strip()
        if re.search(r'общо|total|словом|ддс|дан\.?\s*основа', desc.lower()):
            continue
        pv = _xado_num(price_raw)
        tv = _xado_num(total_raw)
        try:
            n = int(float(qty_raw))
            quantity = f"{n} {'Брой' if n == 1 else 'Броя'}"
        except ValueError:
            quantity = qty_raw
        rec = ProductRecord(extraction_method="text")
        rec.product_code = desc
        rec.quantity     = quantity
        rec.price        = (pv + " EUR") if pv else None
        rec.total_price  = (tv + " EUR") if tv else None
        records.append(rec)
        logger.info("Xado text: desc=%s qty=%s price=%s total=%s", desc[:50], quantity, pv, tv)
    logger.info("Xado text extraction: %d records", len(records))
    return records


def extract_xado_products(tables: list, text: str = "") -> list[ProductRecord]:
    for table in tables:
        recs = _parse_xado_table(table)
        if recs:
            logger.info("Xado table extraction: %d records", len(recs))
            return recs
    return _parse_xado_text(text)


# ---------------------------------------------------------------------------
# Areon (Ареон България)
# ---------------------------------------------------------------------------

def _is_areon_document(text: str) -> bool:
    return bool(re.search(
        r'ареон\s+българия|аромати\s+българия|areon\s+car\s+perfume'
        r'|ареон\s+(?:еоод|оод|ад|bulgaria)'
        r'|\bareon\s+(?:bulgaria|eood|ood)\b'
        r'|\bareon\b',
        text, re.IGNORECASE))


def _areon_num(s: str) -> str | None:
    m = re.search(r'\d+[.,]\d+|\d+', (s or "").replace('\xa0', '').replace(' ', ''))
    if m:
        try:
            return f"{float(m.group().replace(',', '.')):.2f}"
        except ValueError:
            return None
    return None


# Matches a product row in extracted text:
# <pos> <DESCRIPTION> <qty> БР <price> EUR <per> БР <total> EUR [optional suffix]
_AREON_ROW_RE = re.compile(
    # qty has NO space before БР ("10БР"); per DOES have space ("1 БР") — use this to split
    r'(\d+)\s+'           # position number
    r'([А-ЯA-Z][^\n]*?)'  # description (lazy, no newlines)
    r'\s+(\d+)БР\s+'      # qty immediately followed by БР
    r'([\d,.]+)\s+EUR\s+' # unit price
    r'\d+\s+БР\s+'        # per (ignored)
    r'([\d,.]+)\s+EUR'    # total
    r'([ \t][^\n]*)?',    # optional description suffix on same line (e.g. "ЧЕР.ВАНИЛИЯ")
    re.IGNORECASE,
)


def _parse_areon_table(table: list[list]) -> list[ProductRecord]:
    if not table:
        return []

    header_idx = poz_idx = desc_idx = qty_idx = price_idx = total_idx = None

    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if re.search(r'поз|poz', joined) and re.search(r'описание|description', joined):
            header_idx = i
            headers = [str(c or "").lower().strip() for c in row]

            def _find(kws):
                for kw in kws:
                    for j, h in enumerate(headers):
                        if kw in h:
                            return j
                return None

            poz_idx   = _find(["поз", "poz", "no.", "№"])
            desc_idx  = _find(["описание", "description", "артикул"])
            qty_idx   = _find(["кол", "qty"])
            price_idx = _find(["цена", "price"])
            total_idx = _find(["стойност", "total", "amount"])
            break

    if header_idx is None:
        return []

    logger.info("Areon table header at row %d: poz=%s desc=%s qty=%s price=%s total=%s",
                header_idx, poz_idx, desc_idx, qty_idx, price_idx, total_idx)

    records = []
    for row in table[header_idx + 1:]:
        def cell(idx):
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip() if idx is not None and idx < len(row) else ""

        desc = cell(desc_idx)
        if not desc or re.search(r'общо|total|словом|данък|ддс|vat|получател', desc.lower()):
            continue

        qty_raw   = cell(qty_idx)
        qty_clean = re.sub(r'\s*[A-ZА-Яa-zа-я]+\.?\s*$', '', qty_raw).strip()

        rec = ProductRecord(extraction_method="table")
        rec.product_code = re.sub(r'\s+', ' ', desc).upper()
        if qty_clean:
            rec.quantity = qty_clean
        pv = _areon_num(cell(price_idx))
        if pv:
            rec.price = pv + " EUR"
        tv = _areon_num(cell(total_idx))
        if tv:
            rec.total_price = tv + " EUR"

        records.append(rec)
        logger.info("Areon table: desc=%s qty=%s price=%s total=%s", desc[:50], qty_clean, pv, tv)

    return records


def _parse_areon_text(text: str) -> list[ProductRecord]:
    """Text-based fallback when pdfplumber cannot split Areon table columns."""
    records = []
    logger.info("Areon text fallback — first 500 chars:\n%s", repr(text[:500]))

    matches = list(_AREON_ROW_RE.finditer(text))
    for i, m in enumerate(matches):
        poz, desc, qty_raw, price_raw, total_raw, desc_suffix = m.groups()
        desc = re.sub(r'\s+', ' ', desc).strip().upper()

        # Same-line suffix (captured by regex group 6)
        if desc_suffix:
            suffix_clean = re.sub(r'\s+', ' ', desc_suffix).strip().upper()
            if suffix_clean:
                desc = desc + " " + suffix_clean

        # Next-line suffix: text between this match's end and the next match's start.
        # PDF often wraps the description continuation on the line after the price.
        next_start = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        between = text[m.end():next_start]
        between_clean = re.sub(r'\s+', ' ', between).strip()
        # Append only if it looks like a description fragment (Cyrillic present, no digits)
        if between_clean and re.search(r'[А-ЯЁа-яё]', between_clean) and not re.search(r'\d', between_clean):
            desc = desc + " " + between_clean.upper()

        if re.search(r'общо|total|словом|ддс|vat', desc.lower()):
            continue

        rec = ProductRecord(extraction_method="text")
        rec.product_code = desc
        qty_clean = re.sub(r'[^0-9.,]', '', qty_raw).strip()
        if qty_clean:
            rec.quantity = qty_clean
        pv = _areon_num(price_raw)
        if pv:
            rec.price = pv + " EUR"
        tv = _areon_num(total_raw)
        if tv:
            rec.total_price = tv + " EUR"

        records.append(rec)
        logger.info("Areon text: poz=%s desc=%s qty=%s price=%s total=%s",
                    poz, desc[:50], qty_clean, pv, tv)

    return records


def extract_areon_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Areon: %d table(s) received", len(tables))
    raw: list[ProductRecord] = []
    for table in tables:
        raw.extend(_parse_areon_table(table))

    # If table gave only merged-cell junk (product_code > 100 chars), discard and use text
    if raw and all(len(rec.product_code) > 100 for rec in raw):
        logger.warning("Areon: table results look like merged cells, switching to text fallback")
        raw = []

    # Aromati Bulgaria subsidiary — different table headers; try as fallback
    if not raw:
        for table in tables:
            raw.extend(_parse_aromati_table(table))
        if raw and all(len(rec.product_code) > 100 for rec in raw):
            raw = []

    if not raw:
        logger.info("Areon: table extraction yielded nothing, trying text fallback")
        raw = _parse_areon_text(text)
        # Aromati subsidiary — slightly different text layout; try as fallback
        if not raw:
            raw = _parse_aromati_text(text)

    seen: dict[str, ProductRecord] = {}
    for rec in raw:
        code = rec.product_code
        if code not in seen or rec.filled_count() > seen[code].filled_count():
            seen[code] = rec

    records = list(seen.values())
    if records:
        logger.info("Areon: %d records (after dedup)", len(records))
    return records


# ---------------------------------------------------------------------------
# Aromati Bulgaria (Аромати България ЕООД) — subsidiary of Areon
# Invoice format similar to Areon but with different column headers and spacing.
# ---------------------------------------------------------------------------

def _is_aromati_document(text: str) -> bool:
    return bool(re.search(r'аромати\s+българия', text, re.IGNORECASE))


# Matches product rows — same EUR structure as Areon but allows optional space before БР
_AROMATI_ROW_RE = re.compile(
    r'(\d+)\s+'             # position number
    r'([А-ЯA-Z][^\n\r]*?)'  # description (lazy, no CR/LF)
    r'\s+(\d+)\s*БР\.?\s+'  # qty + БР (space optional, dot optional)
    r'([\d,.]+)\s+EUR\s+'   # unit price
    r'\d+\s*БР\.?\s+'       # per (ignored)
    r'([\d,.]+)\s+EUR'      # total
    r'([ \t][^\n\r]*)?',    # optional same-line suffix
    re.IGNORECASE,
)


def _parse_aromati_table(table: list[list]) -> list[ProductRecord]:
    if not table:
        return []

    header_idx = poz_idx = desc_idx = qty_idx = price_idx = total_idx = None

    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        logger.info("Aromati table row %d: %s", i, joined[:150])
        if re.search(r'поз|poz|№|no\.', joined) and re.search(
                r'описание|description|артикул|наименование', joined):
            header_idx = i
            headers = [str(c or "").lower().strip() for c in row]

            def _find(kws):
                for kw in kws:
                    for j, h in enumerate(headers):
                        if kw in h:
                            return j
                return None

            poz_idx   = _find(["поз", "poz", "no.", "№"])
            desc_idx  = _find(["описание", "description", "артикул", "наименование"])
            qty_idx   = _find(["кол", "qty", "количество"])
            price_idx = _find(["цена", "price"])
            total_idx = _find(["стойност", "total", "amount", "сума", "общо"])
            break

    if header_idx is None:
        logger.warning("Aromati table: no header row — rows: %s",
                       [" ".join(str(c or "")[:20] for c in r) for r in table[:5]])
        return []

    logger.info("Aromati table header at row %d: poz=%s desc=%s qty=%s price=%s total=%s",
                header_idx, poz_idx, desc_idx, qty_idx, price_idx, total_idx)

    records = []
    for row in table[header_idx + 1:]:
        def cell(idx):
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip() if idx is not None and idx < len(row) else ""

        desc = cell(desc_idx)
        if not desc or re.search(r'словом|данък|ддс|vat|получател', desc.lower()):
            continue

        qty_raw   = cell(qty_idx)
        qty_clean = re.sub(r'\s*[A-ZА-Яa-zа-я]+\.?\s*$', '', qty_raw).strip()
        qty_n = None
        if qty_clean:
            try:
                qty_n = int(float(qty_clean.replace(',', '.')))
            except ValueError:
                pass

        rec = ProductRecord(extraction_method="table")
        rec.product_code = re.sub(r'\s+', ' ', desc).upper()
        if qty_n is not None:
            rec.quantity = str(qty_n)
        pv = _areon_num(cell(price_idx))
        tv = _areon_num(cell(total_idx))

        # Derive missing total or price from the other when qty is known
        if qty_n and qty_n > 0:
            if pv and not tv:
                try:
                    tv = f"{float(pv) * qty_n:.2f}"
                except ValueError:
                    pass
            elif tv and not pv:
                try:
                    pv = f"{float(tv) / qty_n:.2f}"
                except (ValueError, ZeroDivisionError):
                    pass

        if pv:
            rec.price = pv + " EUR"
        if tv:
            rec.total_price = tv + " EUR"

        records.append(rec)
        logger.info("Aromati table: desc=%s qty=%s price=%s total=%s", desc[:50], qty_n, pv, tv)

    return records


def _parse_aromati_text(text: str) -> list[ProductRecord]:
    records = []
    logger.info("Aromati text fallback — first 1500 chars:\n%s", repr(text[:1500]))

    matches = list(_AROMATI_ROW_RE.finditer(text))
    logger.info("Aromati text: regex found %d match(es)", len(matches))
    for i, m in enumerate(matches):
        poz, desc, qty_raw, price_raw, total_raw, desc_suffix = m.groups()
        desc = re.sub(r'\s+', ' ', desc).strip().upper()

        if desc_suffix:
            suffix_clean = re.sub(r'\s+', ' ', desc_suffix).strip().upper()
            if suffix_clean:
                desc = desc + " " + suffix_clean

        # Next-line suffix: same logic as Areon
        next_start = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        between = text[m.end():next_start]
        between_clean = re.sub(r'\s+', ' ', between).strip()
        if between_clean and re.search(r'[А-ЯЁа-яё]', between_clean) and not re.search(r'\d', between_clean):
            desc = desc + " " + between_clean.upper()

        if re.search(r'общо|total|словом|ддс|vat', desc.lower()):
            continue

        rec = ProductRecord(extraction_method="text")
        rec.product_code = desc
        qty_clean = re.sub(r'[^0-9.,]', '', qty_raw).strip()
        if qty_clean:
            rec.quantity = qty_clean
        pv = _areon_num(price_raw)
        if pv:
            rec.price = pv + " EUR"
        tv = _areon_num(total_raw)
        if tv:
            rec.total_price = tv + " EUR"

        records.append(rec)
        logger.info("Aromati text: poz=%s desc=%s qty=%s price=%s total=%s",
                    poz, desc[:50], qty_clean, pv, tv)

    return records


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Bardahl (distributor: ДНВ Проспийт ЕООД / ProSpeed)
# Bulgarian invoice columns: No, Код (BAR-XXXX), Наименование, К-во, Мярка, Ед. цена, ТО%, Стойност
# ---------------------------------------------------------------------------

def _is_bardahl_document(text: str) -> bool:
    return bool(re.search(r'bardahl', text, re.IGNORECASE))


_BARDAHL_CODE_RE = re.compile(r'^BAR-\d{3,5}$', re.IGNORECASE)


def _bardahl_num(s: str) -> str | None:
    s = (s or "").strip().replace(' ', '').replace(',', '.')
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return None


def _parse_bardahl_table(table: list[list]) -> list[ProductRecord]:
    if not table:
        return []

    header_idx = None
    kod_idx = name_idx = qty_idx = price_idx = total_idx = None

    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if re.search(r'\bкод\b', joined) and re.search(r'\bнаименование\b|\bстока\b', joined):
            header_idx = i
            headers = [str(c or "").lower().strip() for c in row]

            def _find(kws):
                for kw in kws:
                    for j, h in enumerate(headers):
                        if kw in h:
                            return j
                return None

            kod_idx   = _find(["код"])
            name_idx  = _find(["наименование", "стока"])
            qty_idx   = _find(["к-во", "кол", "qty"])
            price_idx = _find(["ед. цена", "ед.цена", "цена"])
            disc_idx  = _find(["отстъпка", "отст", "disc"])
            net_idx   = _find(["нето", "с отст", "нет цена", "крайна"])
            total_idx = _find(["стойност", "total"])
            break

    if header_idx is None:
        logger.warning("Bardahl: no header row found in table (%d rows)", len(table))
        return []

    logger.info("Bardahl headers: %s", list(enumerate(headers)))
    logger.info("Bardahl header at row %d: kod=%s name=%s qty=%s price=%s disc=%s net=%s total=%s",
                header_idx, kod_idx, name_idx, qty_idx, price_idx, disc_idx, net_idx, total_idx)

    records = []
    for row in table[header_idx + 1:]:
        def cell(idx):
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip() if idx is not None and idx < len(row) else ""

        code = cell(kod_idx).upper()
        if not _BARDAHL_CODE_RE.match(code):
            continue

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = cell(name_idx) or None

        qty_raw = cell(qty_idx)
        if qty_raw:
            rec.quantity = qty_raw

        list_price = _bardahl_num(cell(price_idx))
        net_price  = _bardahl_num(cell(net_idx))  if net_idx  is not None else None
        disc_raw   = _bardahl_num(cell(disc_idx)) if disc_idx is not None else None
        total_val  = _bardahl_num(cell(total_idx))

        # Determine the actual (after-discount) unit price
        if net_price:
            final_price = net_price
        elif list_price and disc_raw:
            try:
                final_price = f"{float(list_price) * (1 - float(disc_raw) / 100):.2f}"
            except (ValueError, ZeroDivisionError):
                final_price = list_price
        elif total_val and qty_raw:
            # Fallback: derive from total ÷ qty (works even if discount column name is unknown)
            try:
                qty_f = float(qty_raw.replace(',', '.').replace('\xa0', '').replace(' ', ''))
                computed = float(total_val) / qty_f if qty_f else None
                if computed and (list_price is None or abs(computed - float(list_price)) > 0.005):
                    final_price = f"{computed:.2f}"
                else:
                    final_price = list_price
            except (ValueError, ZeroDivisionError):
                final_price = list_price
        else:
            final_price = list_price

        if final_price:
            rec.price = final_price + " лв."
        if total_val:
            rec.total_price = total_val + " лв."

        records.append(rec)
        logger.info("Bardahl: code=%s qty=%s list=%s disc=%s net=%s total=%s | %s",
                    code, qty_raw, list_price, disc_raw, final_price, total_val,
                    (rec.product_name or "")[:50])

    return records


def extract_bardahl_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Bardahl: %d table(s) received", len(tables))
    raw: list[ProductRecord] = []
    for table in tables:
        raw.extend(_parse_bardahl_table(table))

    # PDF may contain original + copy of the invoice; keep the record with the
    # most data (price > no-price) for each code.
    seen: dict[str, ProductRecord] = {}
    for rec in raw:
        code = rec.product_code
        if code not in seen:
            seen[code] = rec
        elif rec.filled_count() > seen[code].filled_count():
            seen[code] = rec

    records = list(seen.values())
    if records:
        logger.info("Bardahl: %d records from tables (after dedup)", len(records))
    return records


# ---------------------------------------------------------------------------
# Wunder-Baum (distributor: Ауто Ойлс-ЕООД)
# Bulgarian invoice columns: №, Код (EAN-13), Стока, Мярка, Кол., Цена, ДДС %, Стойност
# ---------------------------------------------------------------------------

def _is_wunder_baum_document(text: str) -> bool:
    return bool(re.search(r'wunder.?baum', text, re.IGNORECASE))


def _wb_num(s: str) -> str | None:
    s = (s or "").strip().replace(' ', '').replace(',', '.')
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return None


def _wb_pua_decode(s: str) -> str:
    """Shift PUA characters (U+F000–U+F0FF) back to ASCII."""
    return ''.join(
        chr(ord(c) - 0xF000) if 0xF000 <= ord(c) <= 0xF0FF else c
        for c in (s or "")
    )


def _wb_is_pua(table: list[list]) -> bool:
    """Return True if the table cells are PUA-encoded (custom embedded font)."""
    for row in table[:3]:
        for cell in row:
            if any(0xF000 <= ord(c) <= 0xF0FF for c in str(cell or "")):
                return True
    return False


def _parse_wunder_baum_table(table: list[list]) -> list[ProductRecord]:
    if not table:
        return []

    for i, row in enumerate(table[:5]):
        logger.info("WB table row[%d]: %s", i, [str(c or "")[:40] for c in row])

    # PUA-encoded font: decode all cells and use fixed column positions
    # Standard Wunder-Baum layout (8 cols): №|Код|Стока|Мярка|Кол.|Цена|ДДС%|Стойност
    if _wb_is_pua(table):
        logger.info("WB: PUA-encoded font detected — applying decode + fixed column mapping")
        table = [[_wb_pua_decode(str(c or "")) for c in row] for row in table]
        ncols = len(table[0]) if table else 0
        if ncols != 8:
            logger.warning("WB PUA: unexpected column count %d", ncols)
            return []
        header_idx = 0
        kod_idx, name_idx, qty_idx, price_idx, total_idx = 1, 2, 4, 5, 7
    else:
        # Find header row by keyword (normal font)
        header_idx = None
        kod_idx = name_idx = qty_idx = price_idx = total_idx = None

        for i, row in enumerate(table):
            joined = " ".join(str(c or "").lower() for c in row)
            if re.search(r'\bкод\b', joined) and re.search(r'\bстока\b|\bкол\b', joined):
                header_idx = i
                headers = [str(c or "").lower().strip() for c in row]

                def _find(kws):
                    for kw in kws:
                        for j, h in enumerate(headers):
                            if kw in h:
                                return j
                    return None

                kod_idx   = _find(["код"])
                name_idx  = _find(["стока", "описание", "наименование"])
                qty_idx   = _find(["кол", "qty", "количество"])
                price_idx = _find(["цена", "price"])
                total_idx = _find(["стойност", "ст-ст", "total", "сумма"])
                break

        if header_idx is None:
            logger.warning("WB: no header row found in table (%d rows)", len(table))
            return []

    logger.info("WB header at row %d: kod=%s name=%s qty=%s price=%s total=%s",
                header_idx, kod_idx, name_idx, qty_idx, price_idx, total_idx)

    records = []
    for row in table[header_idx + 1:]:
        def cell(idx):
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip() if idx is not None and idx < len(row) else ""

        # Код: EAN баркод (13 цифри) или вътрешен код (≥4 цифри, напр. 8000, 8028)
        ean_raw = re.sub(r'\s+', '', cell(kod_idx)) if kod_idx is not None else ""
        code = re.sub(r'[^0-9]', '', ean_raw)
        if not re.match(r'^\d{4,14}$', code):
            continue  # header or totals row

        name_raw = cell(name_idx) if name_idx is not None else ""
        name = name_raw

        qty_raw   = cell(qty_idx)
        price_val = _wb_num(cell(price_idx))
        total_val = _wb_num(cell(total_idx))

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = name or None
        if qty_raw:
            rec.quantity = qty_raw
        if price_val:
            rec.price = price_val + " лв."
        if total_val:
            rec.total_price = total_val + " лв."

        records.append(rec)
        logger.info("WB: code=%s qty=%s price=%s total=%s | %s",
                    code, qty_raw, price_val, total_val, (name or "")[:50])

    return records


_WB_EAN_RE = re.compile(r'\b(76\d{11})\b')


def _parse_wunder_baum_from_text(text: str, seen_eans: set[str]) -> list[ProductRecord]:
    """Scan raw text for Wunder-Baum EAN codes not captured by the table parser."""
    records = []
    lines = text.splitlines()
    for i, line in enumerate(lines):
        m = _WB_EAN_RE.search(line)
        if not m:
            continue
        ean = m.group(1)
        if ean in seen_eans:
            continue

        # Collect this line plus next 2 lines; stop at the next EAN occurrence
        block_lines = [line]
        for nl in lines[i + 1:i + 3]:
            if _WB_EAN_RE.search(nl):
                break
            block_lines.append(nl)
        block = " ".join(block_lines)

        # Description: text after the EAN on the same line, up to first digit group
        after_ean = line[m.end():].strip()
        name = re.split(r'\s+\d+[,.]?\d*\s', after_ean)[0].strip() or None

        # Find integers (qty) and decimals (price, total) in the block after EAN
        after_block = block[block.index(ean) + len(ean):]
        decimals = re.findall(r'\b(\d{1,6}[.,]\d{2})\b', after_block)
        integers = re.findall(r'\b(\d{1,4})\b', after_block)

        qty_val   = integers[0] if integers else None
        price_val = _wb_num(decimals[0]) if len(decimals) >= 1 else None
        total_val = _wb_num(decimals[-1]) if len(decimals) >= 2 else None

        rec = ProductRecord(extraction_method="table")
        rec.product_code = ean
        rec.product_name = name
        if qty_val:
            rec.quantity = qty_val
        if price_val:
            rec.price = price_val + " лв."
        if total_val:
            rec.total_price = total_val + " лв."

        seen_eans.add(ean)
        records.append(rec)
        logger.info("WB text fallback: code=%s qty=%s price=%s total=%s | %s",
                    ean, qty_val, price_val, total_val, (name or "")[:50])

    return records


def extract_wunder_baum_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Wunder-Baum: %d table(s) received", len(tables))
    records = []
    seen_eans: set[str] = set()
    for table in tables:
        recs = _parse_wunder_baum_table(table)
        for r in recs:
            if r.ean:
                seen_eans.add(r.ean)
        records.extend(recs)
    if records:
        logger.info("Wunder-Baum: %d records from tables", len(records))

    # Text supplement — catch rows the table parser missed
    if text:
        extra = _parse_wunder_baum_from_text(text, seen_eans)
        if extra:
            logger.info("Wunder-Baum: %d additional records from text", len(extra))
            records.extend(extra)

    return records


# ---------------------------------------------------------------------------
# Slime (ITW Global Tire Repair Europe GmbH)
# English invoice columns: Customer Part | Item | Quantity | Price | Total Amount
# Description appears in the row immediately after each item-code row.
# Code transformation: <item> → SLIME-<item>
# ---------------------------------------------------------------------------

def _is_slime_document(text: str) -> bool:
    return bool(re.search(r'ITW\s+Global\s+Tire|itwgtr\.de|GLOBAL\s+TIRE\s+REPAIR', text, re.IGNORECASE))


# Matches Slime item codes: numeric (10026), numeric-dash (10193-51),
# alphanumeric (SDS-500/06-IN, CRK0305-IN). No spaces, at least 5 chars.
_SLIME_CODE_RE = re.compile(r'^[A-Z0-9][A-Z0-9\-/\.]{4,}$', re.IGNORECASE)

# Matches a product line in Slime invoice text:
# <item_code>  <qty> EA  <unit_price>  <total>
_SLIME_ROW_TEXT_RE = re.compile(
    r'([A-Z0-9][A-Z0-9\-/\.]{4,})'  # item code (no spaces)
    r'\s+(\d+)\s+EA'                  # quantity followed by EA
    r'\s+([\d,\.]+)'                  # unit price (e.g. 4,1500)
    r'\s+([\d,\.]+)',                 # total (e.g. 249,00)
    re.IGNORECASE,
)


def _slime_num(s: str) -> str | None:
    s = (s or "").strip().replace(' ', '')
    if not s:
        return None
    # European format: 1.247,76 → strip dots (thousands sep), replace comma with decimal
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return f"{float(s):.2f}"
    except ValueError:
        return None


def _parse_slime_table(table: list[list]) -> list[ProductRecord]:
    if not table:
        return []

    header_idx = item_idx = qty_idx = price_idx = total_idx = None

    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if "item" in joined and ("quantity" in joined or "price" in joined):
            header_idx = i
            headers = [str(c or "").lower().strip() for c in row]

            def _find(kws):
                for kw in kws:
                    for j, h in enumerate(headers):
                        if kw in h:
                            return j
                return None

            item_idx  = _find(["item"])
            qty_idx   = _find(["quantity", "qty"])
            price_idx = _find(["price"])
            total_idx = _find(["total"])
            break

    if header_idx is None:
        logger.warning("Slime: no header row found in table (%d rows)", len(table))
        return []

    logger.info("Slime header at row %d: item=%s qty=%s price=%s total=%s",
                header_idx, item_idx, qty_idx, price_idx, total_idx)

    def _cell(row, idx):
        if idx is None or idx >= len(row):
            return ""
        return re.sub(r'\s+', ' ', str(row[idx] or "")).strip()

    records = []
    rows = table[header_idx + 1:]
    i = 0
    while i < len(rows):
        row = rows[i]
        item_val = _cell(row, item_idx)

        if not item_val or not _SLIME_CODE_RE.match(item_val):
            i += 1
            continue

        code = "SLIME-" + item_val.upper()
        qty_raw = _cell(row, qty_idx)
        qty_clean = re.sub(r'[^0-9.]', '', qty_raw.split()[0]).strip() if qty_raw else ""
        pv = _slime_num(_cell(row, price_idx))
        tv = _slime_num(_cell(row, total_idx))

        # Description is in column 0 of the next row when that row has no item code
        desc = None
        if i + 1 < len(rows):
            next_row = rows[i + 1]
            next_item = _cell(next_row, item_idx)
            if not next_item or not _SLIME_CODE_RE.match(next_item):
                desc = _cell(next_row, 0) or None
                i += 2
            else:
                i += 1
        else:
            i += 1

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        if desc:
            rec.product_name = desc
        if qty_clean:
            rec.quantity = qty_clean
        if pv:
            rec.price = pv + " EUR"
        if tv:
            rec.total_price = tv + " EUR"

        records.append(rec)
        logger.info("Slime: code=%s qty=%s price=%s total=%s | %s",
                    code, qty_clean, pv, tv, (desc or "")[:50])

    return records


def _parse_slime_text(text: str) -> list[ProductRecord]:
    """Text-based fallback when pdfplumber does not extract the Slime product table."""
    records = []
    logger.info("Slime text fallback — first 600 chars:\n%s", repr(text[:600]))
    lines = text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = _SLIME_ROW_TEXT_RE.search(line)
        if m:
            item_code, qty_raw, price_raw, total_raw = m.groups()
            if not _SLIME_CODE_RE.match(item_code):
                i += 1
                continue
            code = "SLIME-" + item_code.upper()
            pv = _slime_num(price_raw)
            tv = _slime_num(total_raw)

            # Description: next non-empty line that is NOT itself a product row.
            # Only skip if the line contains qty+EA+price (product row); a line that
            # merely starts with a code-like word can still be a valid description.
            desc = None
            j = i + 1
            while j < len(lines):
                next_line = lines[j].strip()
                if not next_line:
                    j += 1
                    continue
                if not _SLIME_ROW_TEXT_RE.search(next_line):
                    desc = next_line
                    j += 1
                break
            i = j

            rec = ProductRecord(extraction_method="text")
            rec.product_code = code
            if desc:
                rec.product_name = desc
            if qty_raw:
                rec.quantity = qty_raw.strip()
            if pv:
                rec.price = pv + " EUR"
            if tv:
                rec.total_price = tv + " EUR"

            records.append(rec)
            logger.info("Slime text: code=%s qty=%s price=%s total=%s | %s",
                        code, qty_raw, pv, tv, (desc or "")[:50])
        else:
            i += 1

    return records


def extract_slime_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Slime: %d table(s) received", len(tables))
    raw: list[ProductRecord] = []
    for table in tables:
        raw.extend(_parse_slime_table(table))

    if not raw and text:
        logger.info("Slime: table extraction yielded nothing, trying text fallback")
        raw = _parse_slime_text(text)

    seen: dict[str, ProductRecord] = {}
    for rec in raw:
        code = rec.product_code
        if code not in seen or rec.filled_count() > seen[code].filled_count():
            seen[code] = rec

    records = list(seen.values())
    if records:
        logger.info("Slime: %d records (after dedup)", len(records))
    return records


# ---------------------------------------------------------------------------
# Bomar (БОМАР БЪЛГАРИЯ ООД — Bulgarian importer of Turtle Wax / car care)
# Invoice format: Bulgarian accounting system table
# Columns: №, Баркод, Код, Стока, Мярка, Кол-во, Кр.цена, Сума
# ---------------------------------------------------------------------------

def _is_bomar_document(text: str) -> bool:
    return bool(re.search(r'бомар\s+българия|bomar\s+bulgar', text, re.IGNORECASE))


def _bomar_num(s: str) -> str | None:
    """Parse a numeric cell from a Bomar invoice (may contain € sign)."""
    s = (s or "").strip().replace('€', '').replace('\xa0', '').replace(' ', '').replace(',', '.')
    m = re.search(r'\d+\.\d+|\d+', s)
    if m:
        try:
            return f"{float(m.group()):.2f}"
        except ValueError:
            return None
    return None


def _parse_bomar_table(table: list[list]) -> list[ProductRecord]:
    """Parse one pdfplumber table from a Бомар България invoice.

    Expected columns: №, Баркод, Код, Стока, Мярка, Кол-во, Кр.цена, Сума
    """
    if not table or len(table) < 2:
        return []

    header_idx = None
    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if "баркод" in joined or ("стока" in joined and "код" in joined):
            header_idx = i
            break
    if header_idx is None:
        return []

    headers = [str(c or "").lower().strip() for c in table[header_idx]]
    logger.debug("Bomar table headers: %s", headers)

    def find(kws):
        for kw in kws:
            for idx, h in enumerate(headers):
                if kw in h:
                    return idx
        return None

    ean_idx   = find(["баркод", "barcode", "ean"])
    # Use exact match for "код" to avoid matching "баркод" (which contains "код")
    code_idx  = next((idx for idx, h in enumerate(headers) if h == "код"), None)
    if code_idx is None:
        code_idx = find(["код"])
    desc_idx  = find(["стока", "описание", "наименование"])
    qty_idx   = find(["кол-во", "кол.", "qty", "количество"])
    price_idx = find(["кр.цена", "кр. цена", "цена"])
    total_idx = find(["сума", "total", "стойност"])

    logger.info("Bomar table header at row %d: ean=%s code=%s desc=%s qty=%s price=%s total=%s",
                header_idx, ean_idx, code_idx, desc_idx, qty_idx, price_idx, total_idx)

    records = []
    for row in table[header_idx + 1:]:
        def cell(ci):
            if ci is None or ci >= len(row):
                return ""
            return re.sub(r'\s+', ' ', str(row[ci] or "")).strip()

        code = cell(code_idx)
        if not code or not re.match(r'^[A-Z]{2}\d+', code):
            continue  # skip totals, headers, empty rows

        desc = cell(desc_idx)
        if not desc:
            continue

        ean_raw = cell(ean_idx)
        ean = ean_raw if re.match(r'^\d{8,14}$', ean_raw) else None

        qty_raw = cell(qty_idx)
        qty_m = re.search(r'\d+', qty_raw)
        qty = qty_m.group() if qty_m else None

        price = _bomar_num(cell(price_idx))
        total = _bomar_num(cell(total_idx))

        rec = ProductRecord(extraction_method="table")
        rec.product_code = code
        rec.product_name = desc
        if ean:
            rec.ean = ean
        if qty:
            rec.quantity = qty
        if price:
            rec.price = price + " EUR"
        if total:
            rec.total_price = total + " EUR"

        records.append(rec)
        logger.info("Bomar table: code=%s ean=%s desc=%s qty=%s price=%s total=%s",
                    code, ean, desc[:50], qty, price, total)

    return records


# Text-based fallback for Bomar (when pdfplumber cannot split table columns).
# Row format: <pos> [<ean>] <FGcode> <description...> БРОЙ <qty> <unit_price> <total>€
_BOMAR_ROW_RE = re.compile(
    r'^\s*(\d+)\s+'            # position number
    r'(?:(\d{8,14})\s+)?'      # optional EAN barcode
    r'([A-Z]{2}\d+)\s+'        # product code (e.g. FG7638)
    r'(.+?)\s+'                # description (lazy)
    r'(?:БРОЙ|брой|бр\.?)\s+'  # unit of measure
    r'(\d+)\s+'                # quantity
    r'([\d.,]+)\s+'            # unit price
    r'([\d.,]+)€?',            # total amount
    re.MULTILINE | re.IGNORECASE,
)


def _parse_bomar_text(text: str) -> list[ProductRecord]:
    """Text-based fallback when pdfplumber cannot split Bomar table columns."""
    logger.info("Bomar text fallback — first 500 chars:\n%s", repr(text[:500]))
    records = []
    for m in _BOMAR_ROW_RE.finditer(text):
        pos, ean_raw, code, desc, qty_raw, price_raw, total_raw = m.groups()
        desc = re.sub(r'\s+', ' ', desc).strip()
        ean = ean_raw if ean_raw and re.match(r'^\d{8,14}$', ean_raw) else None
        price = _bomar_num(price_raw)
        total = _bomar_num(total_raw)

        rec = ProductRecord(extraction_method="text")
        rec.product_code = code
        rec.product_name = desc
        if ean:
            rec.ean = ean
        rec.quantity = qty_raw.strip()
        if price:
            rec.price = price + " EUR"
        if total:
            rec.total_price = total + " EUR"

        records.append(rec)
        logger.info("Bomar text: code=%s ean=%s desc=%s qty=%s price=%s total=%s",
                    code, ean, desc[:50], qty_raw, price, total)

    return records


def extract_bomar_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Bomar: %d table(s) received", len(tables))
    records: list[ProductRecord] = []
    for table in tables:
        records.extend(_parse_bomar_table(table))

    if not records:
        logger.info("Bomar: table extraction yielded nothing, trying text fallback")
        records = _parse_bomar_text(text)

    logger.info("Bomar extraction: %d records", len(records))
    return records


# ---------------------------------------------------------------------------
# Rati (RATI KFT — Hungarian car accessories supplier)
# Invoice format: multi-row per product in "Part No" column:
#   Line 0: product code (e.g. V01945B)
#   Line 1: EAN in parentheses (e.g. (5998167719451))
#   Line 2: description (bilingual, split by " / " — take first part)
#   Line 3: customs code (all digits, 8 chars — ignored)
# Quantity column: "N pcs / db", Unit Price and Total Net Amount are separate cols.
# ---------------------------------------------------------------------------

def _is_rati_document(text: str) -> bool:
    return bool(re.search(r'rati\s+kft|www\.rati\.hu|contact@rati\.hu', text, re.IGNORECASE))


_RATI_CODE_RE = re.compile(r'^[A-Z]\d{3,7}[A-Z]?\d*$')
_RATI_EAN_RE  = re.compile(r'^\((\d{8,14})\)$')


def _rati_num(s: str) -> str | None:
    s = (s or "").strip().replace(' ', '')
    m = re.search(r'\d+[.,]\d+|\d+', s)
    if m:
        try:
            return f"{float(m.group().replace(',', '.')):.2f}"
        except ValueError:
            return None
    return None


def _parse_rati_cell(code_cell: str) -> tuple[str | None, str | None, str | None]:
    """Parse the multi-line Part No cell into (product_code, ean, description)."""
    lines = [l.strip() for l in re.split(r'\n|\r', code_cell) if l.strip()]
    if not lines:
        return None, None, None

    product_code = lines[0]
    if not _RATI_CODE_RE.match(product_code):
        return None, None, None

    ean = None
    desc_lines = []
    for line in lines[1:]:
        if ean is None:
            m_ean = _RATI_EAN_RE.match(line)
            if m_ean:
                ean = m_ean.group(1)
                continue
        if re.match(r'^\d+$', line):  # skip all-digit customs codes
            continue
        desc_lines.append(line)

    desc_raw = " ".join(desc_lines).strip()
    if " / " in desc_raw:
        desc_raw = desc_raw.split(" / ")[0].strip()

    return product_code, ean, desc_raw or None


def _parse_rati_table(table: list[list]) -> list[ProductRecord]:
    if not table:
        return []

    header_idx = code_idx = qty_idx = price_idx = total_idx = None

    for i, row in enumerate(table):
        joined = " ".join(str(c or "").lower() for c in row)
        if re.search(r'part\s*no|cikkszám', joined):
            header_idx = i
            headers = [re.sub(r'\s+', ' ', str(c or "")).lower().strip() for c in row]

            def _find(kws):
                for kw in kws:
                    for j, h in enumerate(headers):
                        if kw in h:
                            return j
                return None

            code_idx  = _find(["part no", "cikkszám"])
            qty_idx   = _find(["quantity", "mennyiség", "uom"])
            price_idx = _find(["unit price", "egységár"])
            total_idx = _find(["total net", "net amount", "érték (áfa nélkül)"])
            if total_idx is None:
                total_idx = _find(["érték"])  # first "érték" col = net (before gross)
            break

    if header_idx is None:
        logger.info("Rati table: no header found — rows: %s",
                    [" ".join(str(c or "")[:20] for c in r) for r in table[:3]])
        return []

    logger.info("Rati table header at row %d: code=%s qty=%s price=%s total=%s",
                header_idx, code_idx, qty_idx, price_idx, total_idx)

    records = []
    for row in table[header_idx + 1:]:
        def cell(idx):
            return re.sub(r'\s+', ' ', str(row[idx] or "")).strip() \
                if idx is not None and idx < len(row) else ""

        code_cell = cell(code_idx)
        if not code_cell:
            continue

        product_code, ean, desc = _parse_rati_cell(code_cell)
        if not product_code:
            continue

        qty_raw = cell(qty_idx)
        qty_m   = re.search(r'(\d+)', qty_raw)
        qty     = qty_m.group(1) if qty_m else None

        pv = _rati_num(cell(price_idx))
        tv = _rati_num(cell(total_idx))

        rec = ProductRecord(extraction_method="table")
        rec.product_code = product_code
        rec.ean          = ean
        if desc:
            rec.product_name = desc
        if qty:
            rec.quantity = qty
        if pv:
            rec.price = pv + " EUR"
        if tv:
            rec.total_price = tv + " EUR"

        records.append(rec)
        logger.info("Rati table: code=%s ean=%s desc=%s qty=%s price=%s total=%s",
                    product_code, ean, (desc or "")[:40], qty, pv, tv)

    return records


def _parse_rati_text(text: str) -> list[ProductRecord]:
    """Text fallback: product code line followed by EAN and description lines.

    Handles two OCR layouts:
    A) Code + qty + prices all on the same line (clean PyMuPDF / page 3 OCR):
       "V01945B  2 pcs / db  52.40  104.80"
    B) Code on one line, qty+prices on a nearby line (multi-column OCR pages 1-2):
       "V01945B"  ...  "2 pcs / db  52.40  104.80"
    """
    records = []
    # Normalize common OCR misreadings before parsing
    text = (text
            .replace('¥', 'V').replace('Ÿ', 'V').replace('У', 'V')
            .replace('рcs', 'pcs')   # Cyrillic р misread as Latin p before "cs"
            .replace('рс', 'pc')     # Cyrillic рс misread as pc
            )

    # Fix Cyrillic OCR artifacts in product code tokens on qty lines.
    # Tesseract reads Latin M/O/C/A and digit 3 as Cyrillic М/О/С/А/З in codes.
    _CYR_CODE_MAP = str.maketrans('МОСЗА', 'M0C3A')
    _fixed = []
    for _ln in text.splitlines():
        if re.search(r'p[ce]s\s*/\s*db', _ln, re.IGNORECASE):
            _toks = _ln.split()
            for _ti, _tok in enumerate(_toks):
                if re.match(r'^[А-ЯA-Z][А-ЯA-Z0-9]{3,8}$', _tok):
                    _toks[_ti] = _tok.translate(_CYR_CODE_MAP)
                    break
            _ln = ' '.join(_toks)
        _fixed.append(_ln)
    text = '\n'.join(_fixed)

    logger.info("Rati text fallback — first 2000 chars:\n%s", repr(text[:2000]))

    _RATI_CODE_ONLY_RE = re.compile(r'\b([A-Z]\d{3,7}[A-Z]?\d*|[A-Z]{3,}\d+[A-Z]{2,})\b')
    _RATI_QTY_PRICE_RE = re.compile(
        r'(\d+)\s*p[ce]s\s*/\s*db\s+([\d.]+)\s+([\d.]+)', re.IGNORECASE)

    lines = text.splitlines()
    used = set()  # line indices already consumed

    for i, line in enumerate(lines):
        if i in used:
            continue
        line_s = line.strip()

        # Does this line contain a product code?
        code_m = _RATI_CODE_ONLY_RE.search(line_s)
        if not code_m:
            continue
        code = code_m.group(1)

        # Rati codes always start with V (e.g. V01884A, V01946C) or are all-letter
        # freight codes (e.g. FUVB2BEU).  OCR running on the custom-font PDF often
        # misreads the V glyph as Cyrillic М; _CYR_CODE_MAP then converts М→M,
        # producing wrong codes like M01884A.  Correct M<digits>... back to V<digits>...
        # since Rati never uses M-prefix product codes.
        if code.startswith('M') and len(code) > 3 and code[1:2].isdigit():
            corrected = 'V' + code[1:]
            logger.debug("Rati: M→V prefix correction: %s → %s", code, corrected)
            code = corrected

        # Look for qty+prices on this line OR within the next 6 lines
        qty_s = pv_s = tv_s = None
        qty_line_idx = None
        window = [line_s] + [lines[k].strip() for k in range(i + 1, min(i + 7, len(lines)))]
        for wi, wl in enumerate(window):
            qm = _RATI_QTY_PRICE_RE.search(wl)
            if qm:
                qty_s, pv_s, tv_s = qm.groups()
                qty_line_idx = i + wi
                break

        if qty_s is None:
            continue  # no qty found near this code — not a product row

        # Mark code line and qty line as used
        used.add(i)
        if qty_line_idx != i:
            used.add(qty_line_idx)

        # Collect EAN and description from lines after the code line
        ean = None
        desc_lines = []
        scan_end = max(qty_line_idx, i) + 5
        for k in range(i + 1, min(scan_end + 1, len(lines))):
            kl = lines[k].strip()
            if not kl:
                continue
            # Stop at next product's qty line — don't consume it
            if k != qty_line_idx and _RATI_QTY_PRICE_RE.search(kl):
                break
            # Skip current product's qty line (already marked as used)
            if k == qty_line_idx:
                continue
            used.add(k)
            ean_m = _RATI_EAN_RE.match(kl)
            if ean_m:
                ean = ean_m.group(1)
            elif not re.match(r'^\d+$', kl) and not re.match(r'^[\d\s.]+$', kl):
                desc_lines.append(kl)

        desc_raw = " ".join(desc_lines).strip()
        if " / " in desc_raw:
            desc_raw = desc_raw.split(" / ")[0].strip()

        pv = _rati_num(pv_s)
        tv = _rati_num(tv_s)

        rec = ProductRecord(extraction_method="text")
        rec.product_code = code
        rec.ean          = ean
        if desc_raw:
            rec.product_name = desc_raw
        rec.quantity     = qty_s
        if pv:
            rec.price = pv + " EUR"
        if tv:
            rec.total_price = tv + " EUR"
        records.append(rec)
        logger.info("Rati text: code=%s ean=%s desc=%s qty=%s price=%s total=%s",
                    code, ean, desc_raw[:40] if desc_raw else "", qty_s, pv, tv)

    logger.info("Rati text extraction: %d records", len(records))
    # Log unmatched lines with qty pattern to help diagnose missed products
    for j, ln in enumerate(lines):
        if j not in used and _RATI_QTY_PRICE_RE.search(ln.strip()):
            logger.info("Rati unmatched qty line %d: %s", j, repr(ln.strip()))
    return records


def extract_rati_products(tables: list, text: str = "") -> list[ProductRecord]:
    logger.info("Rati: %d table(s) received", len(tables))
    for table in tables:
        recs = _parse_rati_table(table)
        if recs:
            logger.info("Rati table extraction: %d records", len(recs))
            return recs
    logger.info("Rati: table yielded nothing, trying text fallback")
    return _parse_rati_text(text)


# ---------------------------------------------------------------------------
# Generic Bulgarian invoice text parser
# Handles the standard Bulgarian ERP invoice format (Microinvest / Акаунт Е):
#   № | [Арт.№ Клиент] | Арт.№ Дост. | Наименование | МЕ | К-во | Ед.цена | Общо
# ---------------------------------------------------------------------------

_BG_UNITS = r'(?:брой|бр\.?|кг|л(?:итър)?|м(?:етър)?|к(?:омплект)?|оп\.?|пакет|пак\.?|set|pcs|pc)'

# Row pattern: leading row-number, optional client-art, supplier-art (4-10 digits),
# description (anything), unit, qty, unit-price, total
_BG_INV_ROW_RE = re.compile(
    r'^\s*\d+\s+'                             # row number
    r'(?:\S+\s+)?'                             # optional client article (skip)
    r'(\d{4,10})\s+'                           # supplier article → product code
    r'(.+?)\s+'                                # description (non-greedy)
    r'(' + _BG_UNITS + r')\s+'                # unit of measure
    r'(\d+(?:[.,]\d+)?)\s+'                   # quantity
    r'([\d.,]+)\s+'                            # unit price
    r'([\d.,]+)\s*$',                          # total
    re.IGNORECASE,
)

# Simpler fallback: row-number + code + qty + price + total (no unit / description)
_BG_INV_SIMPLE_RE = re.compile(
    r'^\s*\d+\s+'
    r'(\d{4,10})\s+'
    r'(.+?)\s+'
    r'(\d+(?:[.,]\d+)?)\s+'
    r'([\d.,]+)\s+'
    r'([\d.,]+)\s*$',
)


def _bg_num(s: str) -> str:
    """Normalise Bulgarian number: '1.234,56' → '1234.56'."""
    s = s.strip()
    if '.' in s and ',' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    return s


def _parse_bulgarian_invoice_text(text: str) -> list[ProductRecord]:
    """Extract product rows from a standard Bulgarian ERP invoice (text mode)."""
    records = []
    for line in text.splitlines():
        m = _BG_INV_ROW_RE.match(line)
        if m:
            code, name, unit, qty, price, total = m.groups()
            rec = ProductRecord(extraction_method="regex")
            rec.product_code = code.strip()
            rec.product_name = name.strip()[:120]
            rec.quantity     = _bg_num(qty) + " " + unit.strip()
            rec.price        = _bg_num(price) + " BGN"
            rec.total_price  = _bg_num(total) + " BGN"
            records.append(rec)
            logger.debug("BG invoice row: code=%s qty=%s price=%s total=%s",
                         code, qty, price, total)
            continue

        m = _BG_INV_SIMPLE_RE.match(line)
        if m:
            code, name, qty, price, total = m.groups()
            rec = ProductRecord(extraction_method="regex")
            rec.product_code = code.strip()
            rec.product_name = name.strip()[:120]
            rec.quantity     = _bg_num(qty)
            rec.price        = _bg_num(price) + " BGN"
            rec.total_price  = _bg_num(total) + " BGN"
            records.append(rec)
            logger.debug("BG invoice simple row: code=%s qty=%s price=%s total=%s",
                         code, qty, price, total)

    return records


class FieldMapper:
    def __init__(self, llm=None):
        self.llm = llm  # Optional llama-cpp-python Llama instance
        # Set to True after map() whenever a dedicated supplier extractor was attempted.
        # False only when the generic table/regex/BG-invoice fallback ran.
        self.last_specific_tried: bool = False

    def map(self, extracted: dict, supplier: str = "auto") -> list[ProductRecord]:
        tables = extracted.get("tables", [])
        text = extracted.get("text", "")
        self.last_specific_tried = False  # reset for this call

        logger.info("Extraction requested: supplier=%s", supplier)

        # Track whether a dedicated supplier extractor was attempted.
        # When True we never fall back to generic table/regex extraction — that
        # would produce garbage from preamble / address tables.
        _specific_tried = False

        # Step 0 — OSRAM (explicit selection or auto-detection)
        if supplier == "osram" or (supplier == "auto" and text and _is_osram_document(text)):
            _specific_tried = self.last_specific_tried = True
            if text:
                records = extract_osram_products(text)
                if records:
                    logger.info("Extraction method: OSRAM-specific (%d records)", len(records))
                    return records

        # Step 0b — Rezaw-Plast (explicit selection or auto-detection)
        if supplier == "rezaw_plast" or (supplier == "auto" and _is_rezaw_plast_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_rezaw_plast_products(tables, text)
            if records:
                logger.info("Extraction method: Rezaw-Plast (%d records)", len(records))
                return records

        # Step 0c — Avisa (explicit selection or auto-detection)
        if supplier == "avisa" or (supplier == "auto" and _is_avisa_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_avisa_products(tables, text)
            if records:
                logger.info("Extraction method: Avisa (%d records)", len(records))
                return records

        # Step 0d — Amio (explicit selection or auto-detection)
        if supplier == "amio" or (supplier == "auto" and _is_amio_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_amio_products(tables, text)
            if records:
                logger.info("Extraction method: Amio (%d records)", len(records))
                return records

        # Step 0e — Maxton Design (explicit selection or auto-detection)
        if supplier == "maxton_design" or (supplier == "auto" and _is_maxton_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_maxton_products(tables, text)
            if records:
                logger.info("Extraction method: Maxton Design (%d records)", len(records))
                return records

        # Step 0f — M-Tech Poland (explicit selection or auto-detection)
        if supplier == "mtech" or (supplier == "auto" and _is_mtech_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_mtech_products(tables, text)
            if records:
                logger.info("Extraction method: M-Tech (%d records)", len(records))
                return records

        # Step 0g — Ma*Fra / Авиатранс (explicit selection or auto-detection)
        text2 = extracted.get("text2", "")
        if supplier == "mafra" or (supplier == "auto" and (_is_mafra_document(text) or _is_mafra_document(text2))):
            _specific_tried = self.last_specific_tried = True
            records = extract_mafra_products(tables, text, text2)
            if records:
                logger.info("Extraction method: Ma*Fra (%d records)", len(records))
                return records

        # Step 0h — Amal-Plast (explicit selection or auto-detection)
        if supplier == "amal_plast" or (supplier == "auto" and _is_amal_plast_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_amal_plast_products(tables, text)
            if records:
                logger.info("Extraction method: Amal-Plast (%d records)", len(records))
                return records

        # Step 0i — Car Passion (explicit selection or auto-detection)
        if supplier == "car_passion" or (supplier == "auto" and _is_car_passion_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_car_passion_products(tables, text)
            if records:
                logger.info("Extraction method: Car Passion (%d records)", len(records))
                return records

        # Step 0j — Vinove (explicit selection or auto-detection)
        if supplier == "vinove" or (supplier == "auto" and _is_vinove_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_vinove_products(tables, text)
            if records:
                logger.info("Extraction method: Vinove (%d records)", len(records))
                return records

        # Step 0k — Gumarny Zubri (explicit selection or auto-detection)
        if supplier == "gumarny_zubri" or (supplier == "auto" and _is_gumarny_zubri_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_gumarny_zubri_products(tables, text)
            if records:
                logger.info("Extraction method: Gumarny Zubri (%d records)", len(records))
                return records

        # Step 0l — Rigum (explicit selection or auto-detection)
        if supplier == "rigum" or (supplier == "auto" and _is_rigum_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_rigum_products(tables, text)
            if records:
                logger.info("Extraction method: Rigum (%d records)", len(records))
                return records

        # Step 0m — Frogum (explicit selection or auto-detection)
        if supplier == "frogum" or (supplier == "auto" and _is_frogum_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_frogum_products(tables, text)
            if records:
                logger.info("Extraction method: Frogum (%d records)", len(records))
                return records

        # Step 0n — Petex (explicit selection or auto-detection)
        if supplier == "petex" or (supplier == "auto" and _is_petex_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_petex_products(tables, text)
            if records:
                logger.info("Extraction method: Petex (%d records)", len(records))
                return records

        # Step 0o — Gelly Plast (explicit selection or auto-detection)
        if supplier == "gelly_plast" or (supplier == "auto" and _is_gelly_plast_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_gelly_plast_products(tables, text)
            if records:
                logger.info("Extraction method: Gelly Plast (%d records)", len(records))
                return records

        # Step 0p — Farad / Evolution SRL (explicit selection or auto-detection)
        if supplier == "farad" or (supplier == "auto" and _is_farad_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_farad_products(tables, text)
            if records:
                logger.info("Extraction method: Farad (%d records)", len(records))
                return records

        # Step 0n — Geyer & Hosaja (explicit selection or auto-detection)
        if supplier == "geter_hosaja" or (supplier == "auto" and _is_geyer_hosaja_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_geyer_hosaja_products(tables, text)
            if records:
                logger.info("Extraction method: Geyer & Hosaja (%d records)", len(records))
                return records

        # Step 0r — Hakr / ASN HAKR Brno (explicit selection or auto-detection)
        if supplier == "hakr" or (supplier == "auto" and _is_hakr_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_hakr_products(tables, text)
            if records:
                logger.info("Extraction method: Hakr (%d records)", len(records))
                return records

        # Step 0s — AutoMania (explicit selection or auto-detection)
        if supplier == "automania" or (supplier == "auto" and _is_automania_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_automania_products(tables, text)
            if records:
                logger.info("Extraction method: AutoMania (%d records)", len(records))
                return records

        # Step 0q — Kegel-Błażusiak (explicit selection or auto-detection)
        if supplier == "kegel_blazusiak" or (supplier == "auto" and _is_kegel_blazusiak_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_kegel_blazusiak_products(tables, text)
            if records:
                logger.info("Extraction method: Kegel-Blazusiak (%d records)", len(records))
                return records

        # Step 0t — ToM-PaR (explicit selection or auto-detection)
        if supplier == "tompar" or (supplier == "auto" and _is_tompar_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_tompar_products(tables, text)
            if records:
                logger.info("Extraction method: ToM-PaR (%d records)", len(records))
                return records

        # Step 0u — Sonax / Сенакс ООД (explicit selection or auto-detection)
        if supplier == "sonax" or (supplier == "auto" and _is_senax_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_senax_products(tables, text)
            if records:
                logger.info("Extraction method: Sonax (%d records)", len(records))
                return records

        # Step 0v — Heko / Team Heko (explicit selection or auto-detection)
        if supplier == "team_heko" or (supplier == "auto" and _is_heko_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_heko_products(tables, text)
            if records:
                logger.info("Extraction method: Heko (%d records)", len(records))
                return records

        # Step 0w — BMW Group / BMW Service (explicit selection or auto-detection)
        if supplier == "bmw" or (supplier == "auto" and _is_bmw_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_bmw_products(tables, text)
            if records:
                logger.info("Extraction method: BMW (%d records)", len(records))
                return records

        # Step 0y — Bardahl / ProSpeed (explicit selection or auto-detection)
        if supplier == "bardahl" or (supplier == "auto" and _is_bardahl_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_bardahl_products(tables, text)
            if records:
                logger.info("Extraction method: Bardahl (%d records)", len(records))
                return records

        # Step 0x — Wunder-Baum (explicit selection or auto-detection)
        if supplier == "wunder_baum" or (supplier == "auto" and _is_wunder_baum_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_wunder_baum_products(tables, text)
            if records:
                logger.info("Extraction method: Wunder-Baum (%d records)", len(records))
                return records

        # Step 0z — Areon (explicit selection or auto-detection)
        if supplier == "areon" or (supplier == "auto" and _is_areon_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_areon_products(tables, text)
            if records:
                logger.info("Extraction method: Areon (%d records)", len(records))
                return records

        # Step 0z2 — Xado / Tuning Oils Club (explicit selection or auto-detection)
        if supplier == "xado" or (supplier == "auto" and _is_xado_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_xado_products(tables, text)
            if records:
                logger.info("Extraction method: Xado (%d records)", len(records))
                return records

        # Step 0z3 — Slime / ITW Global Tire Repair (explicit selection or auto-detection)
        if supplier == "slime" or (supplier == "auto" and _is_slime_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_slime_products(tables, text)
            if records:
                logger.info("Extraction method: Slime (%d records)", len(records))
                return records

        # Step 0z4 — Rati KFT (explicit selection or auto-detection)
        if supplier == "rati" or (supplier == "auto" and _is_rati_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_rati_products(tables, text)
            if records:
                logger.info("Extraction method: Rati (%d records)", len(records))
                return records

        if supplier == "bomar" or (supplier == "auto" and _is_bomar_document(text)):
            _specific_tried = self.last_specific_tried = True
            records = extract_bomar_products(tables, text)
            if records:
                logger.info("Extraction method: Bomar (%d records)", len(records))
                return records

        # If a dedicated extractor was attempted but returned 0, do NOT fall back
        # to generic table/text extraction — it would pick up preamble/address rows.
        self.last_specific_tried = _specific_tried
        if _specific_tried:
            logger.info("Specific extractor attempted but returned 0 records — skipping generic fallback")
            return []

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

            # Step 4 — generic Bulgarian invoice line parser
            # Handles standard ERP format: № | [client art] | supplier art | name | unit | qty | price | total
            bg_records = _parse_bulgarian_invoice_text(text)
            if bg_records:
                logger.info("Extraction method: generic BG invoice (%d records)", len(bg_records))
                return bg_records

            # Return partial regex result if nothing better found
            if rec.filled_count() >= 1:
                return [rec]

        return []
