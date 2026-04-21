"""
Run this script ONCE from F:\\maintenance-503\\ to apply all OSRAM extraction fixes.
Usage:  python patch_field_mapper.py
"""
import re
from pathlib import Path

TARGET = Path(__file__).parent / "app" / "processors" / "field_mapper.py"

src = TARGET.read_text(encoding="utf-8")

# ── PATCH 1: add spec-detection regex + product-code helper after _UNIT_PRICE_RE ──
OLD1 = """# Unit price: "10,77/ 1 PCE"
_UNIT_PRICE_RE = re.compile(r'([\\d,.]+)\\s*/\\s*1\\s*PCE', re.IGNORECASE)


def _is_osram_document(text: str) -> bool:"""

NEW1 = """# Unit price: "10,77/ 1 PCE"
_UNIT_PRICE_RE = re.compile(r'([\\d,.]+)\\s*/\\s*1\\s*PCE', re.IGNORECASE)

# Tokens that mark the start of technical specs on a position line
_OSRAM_SPEC_RE = re.compile(
    r'^\\d+[.,]\\d*[WwVvKk]'   # 1,8W  36V  2700K
    r'|^\\d+[WwVvKk]$'         # 4W  12V
    r'|^PG\\d'                  # PG20-1
    r'|^G\\d+[.\\-/]?\\d*$'      # G4  G13
    r'|^E\\d+$'                 # E14  E27
    r'|^\\d{4}K$',              # 2700K
    re.IGNORECASE
)


def _extract_osram_product_code(pos_line: str):
    tokens = pos_line.strip().split()
    i = 0
    while i < len(tokens) and _POS_RE.match(tokens[i]):
        i += 1
    if i < len(tokens) and re.match(r'^\\d+$', tokens[i]):
        i += 1
    while i < len(tokens) and re.search(r'[а-яА-Я]', tokens[i]):
        i += 1
    code_tokens = []
    while i < len(tokens):
        tok = tokens[i]
        if _OSRAM_SPEC_RE.match(tok):
            break
        if tok.upper().rstrip('.,') in {'OSRAM', 'LEDVANCE', 'PHILIPS'}:
            break
        if re.match(r'^\\d{1,6}[.,]\\d{2}$', tok):
            break
        code_tokens.append(tok)
        i += 1
    result = ' '.join(code_tokens).strip()
    return result if len(result) >= 3 else None


def _is_osram_document(text: str) -> bool:"""

# ── PATCH 2: fix quantity, price, total, and product_code in _parse_osram_by_article ──
OLD2 = """        rec = ProductRecord(extraction_method="osram")
        osram_article = m.group(1)
        rec.product_code = osram_article
        rec._osram_article = osram_article  # type: ignore[attr-defined]

        # EAN: 13 or 14 digit number"""

NEW2 = """        rec = ProductRecord(extraction_method="osram")
        osram_article = m.group(1)
        rec.product_code = None
        rec._osram_article = osram_article  # type: ignore[attr-defined]

        # EAN: 13 or 14 digit number"""

# ── PATCH 3: add product_code extraction from position line after EAN block ──
OLD3 = """        if ean_m:
            rec.ean = ean_m.group(1)

        # ── Quantity: 2nd token on position line"""

NEW3 = """        if ean_m:
            rec.ean = ean_m.group(1)

        # ── Product code: "LEDPWL ACC 103 30X1" from position line
        for bl in reversed(before_lines):
            if _POS_RE.match(bl.strip()):
                pc = _extract_osram_product_code(bl)
                if pc:
                    rec.product_code = pc
                break
        if not rec.product_code:
            rec.product_code = osram_article  # fallback

        # ── Quantity: 2nd token on position line"""

# ── PATCH 4: fix quantity logic (2nd token of pos line, not "N Брой" search) ──
OLD4 = """        # ── Quantity: 2nd token on position line (000NNN  QTY  description...)
        # Iterate backwards through lines before the article to find nearest pos line
        for bl in reversed(before_lines):
            if _POS_RE.match(bl.strip()):
                tokens = bl.strip().split()
                # tokens[0]=position, tokens[1]=quantity (pure digits)
                if len(tokens) >= 2 and re.match(r'^\\d+$', tokens[1]):
                    rec.quantity = tokens[1] + " PCE"
                break
        # Fallback: search for explicit "N Брой" in before context
        if not rec.quantity:
            for qty_m in re.finditer(r'\\b(\\d+)\\s*(?:Брой|бр\\.?)\\b', "\\n".join(before_lines), re.IGNORECASE):
                rec.quantity = qty_m.group(1) + " PCE"
                break"""

# Check if patch 4 is already applied (new version) or still has the old version
OLD4_OLD = """        # Quantity: "N Брой" or "N PCE" — skip "1 PCE" from price lines
        for qty_m in re.finditer(r'\\b(\\d+)\\s*(?:Брой|бр\\.?|PCE|STK)\\b', ctx, re.IGNORECASE):
            qty_val = int(qty_m.group(1))
            # "1 PCE" in "10,77/ 1 PCE" is the price denominator, not quantity
            before = ctx[max(0, qty_m.start() - 10):qty_m.start()]
            if re.search(r'[\\d,./]\\s*$', before) and qty_val == 1:
                continue
            rec.quantity = str(qty_val) + " PCE"
            break"""

NEW4_OLD = """        for bl in reversed(before_lines):
            if _POS_RE.match(bl.strip()):
                tokens = bl.strip().split()
                if len(tokens) >= 2 and re.match(r'^\\d+$', tokens[1]):
                    rec.quantity = tokens[1] + " PCE"
                break
        if not rec.quantity:
            for qty_m in re.finditer(r'\\b(\\d+)\\s*(?:Брой|бр\\.?)\\b', "\\n".join(before_lines), re.IGNORECASE):
                rec.quantity = qty_m.group(1) + " PCE"
                break"""

# ── PATCH 5: fix unit price to search only after article line ──
OLD5 = """        up_m = _UNIT_PRICE_RE.search(ctx)"""
NEW5 = """        up_m = _UNIT_PRICE_RE.search(after_ctx)"""

# ── Apply patches ──
applied = []
failed  = []

def apply(name, old, new):
    global src
    if old in src:
        src = src.replace(old, new, 1)
        applied.append(name)
    else:
        failed.append(name)

apply("PATCH 1 — spec regex + product code helper", OLD1, NEW1)
apply("PATCH 2 — product_code = None", OLD2, NEW2)
apply("PATCH 3 — product_code from pos line", OLD3, NEW3)

# Patch 4: try new version first, then old
if OLD4 in src:
    applied.append("PATCH 4 — quantity (already fixed)")
elif OLD4_OLD in src:
    src = src.replace(OLD4_OLD, NEW4_OLD, 1)
    applied.append("PATCH 4 — quantity fix applied")
else:
    failed.append("PATCH 4 — quantity (not matched)")

if OLD5 in src:
    src = src.replace(OLD5, NEW5, 1)
    applied.append("PATCH 5 — unit price after-only")
else:
    if "after_ctx" in src and "_UNIT_PRICE_RE.search(after_ctx)" in src:
        applied.append("PATCH 5 — unit price (already fixed)")
    else:
        failed.append("PATCH 5 — unit price (not matched)")

TARGET.write_text(src, encoding="utf-8")

print("=" * 50)
print("Applied:", len(applied))
for p in applied:
    print("  OK ", p)
if failed:
    print("Skipped (may already be applied):", len(failed))
    for p in failed:
        print("  -- ", p)
print("=" * 50)
print("Done! Restart the bot.")
