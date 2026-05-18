import asyncio
import logging
import os
import shutil
import signal
import threading
import time
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Query
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request
import uvicorn

from .extractors.pdf_extractor import PDFExtractor
from .extractors.excel_extractor import ExcelExtractor
from .extractors.text_extractor import TextExtractor
from .extractors.image_extractor import ImageExtractor
from .processors.field_mapper import FieldMapper, ProductRecord, _is_osram_document, _is_rigum_document, _is_bmw_document, _is_gumarny_zubri_document, _is_wunder_baum_document, _is_slime_document, _is_xado_document
from .processors.product_db import get_product_db, get_rezaw_plast_db, get_maxton_db, get_avisa_db, get_amio_db, get_mtech_db, get_mafra_db, get_amal_plast_db, get_car_passion_db, get_vinove_db, get_gumarny_zubri_db, get_rigum_db, get_petex_db, get_geyer_hosaja_db, get_frogum_db, get_gelly_plast_db, get_farad_db, get_farad_code_map, get_kegel_blazusiak_db, get_automania_db, get_hakr_db, get_tompar_db, get_senax_db, get_heko_db, get_bmw_db, get_wunder_baum_db, get_wunder_baum_code_map, get_bardahl_db, get_areon_db, get_areon_code_map, get_slime_db, get_xado_db
from .output.excel_writer import write_excel

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
MODEL_PATH = BASE_DIR / "models" / "phi-3-mini.gguf"

DATA_DIR = BASE_DIR / "data" / "suppliers"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)

SUPPORTED_TYPES = {
    ".pdf":  "pdf",
    ".xlsx": "excel",
    ".xls":  "excel",
    ".xlsm": "excel",
    ".csv":  "excel",
    ".txt":  "text",
    ".png":  "image",
    ".jpg":  "image",
    ".jpeg": "image",
    ".tiff": "image",
    ".tif":  "image",
    ".bmp":  "image",
    ".webp": "image",
}

SUPPORTED_LANGUAGES = ["bg", "en", "pl", "cs", "it", "de"]

app = FastAPI(title="Data Extraction Bot", version="1.0.0")

templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))
static_path = Path(__file__).parent / "static"
if static_path.exists():
    app.mount("/static", StaticFiles(directory=str(static_path)), name="static")

# ── Heartbeat-based auto-shutdown ────────────────────────────────────────────
# JS sends POST /heartbeat every 10s while the page is open.
# If no heartbeat arrives for 30s, the server shuts down automatically.
_last_heartbeat: float = time.time()
_HEARTBEAT_TIMEOUT = 120  # seconds


def _heartbeat_monitor():
    time.sleep(15)  # grace period at startup
    while True:
        time.sleep(5)
        if time.time() - _last_heartbeat > _HEARTBEAT_TIMEOUT:
            logger.info("No heartbeat for %ds — shutting down.", _HEARTBEAT_TIMEOUT)
            os.kill(os.getpid(), signal.SIGTERM)
            break


@app.post("/heartbeat")
async def heartbeat():
    global _last_heartbeat
    _last_heartbeat = time.time()
    return {"ok": True}


@app.post("/shutdown")
async def shutdown():
    """Explicit shutdown — called when browser tab closes."""
    def _kill():
        time.sleep(0.5)
        os.kill(os.getpid(), signal.SIGTERM)
    threading.Thread(target=_kill, daemon=True).start()
    return JSONResponse({"ok": True})


# Global state — initialized lazily to avoid heavy startup cost
_extractors: dict = {}
_mapper: FieldMapper | None = None
_records_cache: dict[str, list] = {}  # filename → records for custom export


def get_extractor(file_type: str, languages: list[str]):
    key = (file_type, tuple(sorted(languages)))
    if key not in _extractors:
        if file_type == "pdf":
            _extractors[key] = PDFExtractor(ocr_languages=languages)
        elif file_type == "excel":
            _extractors[key] = ExcelExtractor()
        elif file_type == "text":
            _extractors[key] = TextExtractor()
        elif file_type == "image":
            _extractors[key] = ImageExtractor(languages=languages)
    return _extractors[key]


def get_mapper() -> FieldMapper:
    global _mapper
    if _mapper is None:
        llm = None
        if MODEL_PATH.exists():
            try:
                from llama_cpp import Llama
                llm = Llama(model_path=str(MODEL_PATH), n_ctx=4096, n_threads=4, verbose=False)
                logger.info("LLM loaded: %s", MODEL_PATH.name)
            except Exception as e:
                logger.warning("LLM not available (%s) — regex-only mode", e)
        else:
            logger.info("No LLM model found at %s — regex-only mode", MODEL_PATH)
        _mapper = FieldMapper(llm=llm)
    return _mapper


@app.on_event("startup")
async def _start_heartbeat_monitor():
    t = threading.Thread(target=_heartbeat_monitor, daemon=True)
    t.start()


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/extract")
async def extract(
    file: UploadFile = File(...),
    languages: str = Form(default="bg,en"),
    supplier: str = Form(default="auto"),
):
    suffix = Path(file.filename).suffix.lower()
    file_type = SUPPORTED_TYPES.get(suffix)
    if not file_type:
        raise HTTPException(400, f"Неподдържан формат: {suffix}. Поддържани: {list(SUPPORTED_TYPES.keys())}")

    lang_list = [l.strip() for l in languages.split(",") if l.strip() in SUPPORTED_LANGUAGES]
    if not lang_list:
        lang_list = ["bg", "en"]

    # Save upload
    tmp_path = UPLOAD_DIR / file.filename
    with open(tmp_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        extractor = get_extractor(file_type, lang_list)
        loop = asyncio.get_event_loop()
        with ThreadPoolExecutor(max_workers=1) as pool:
            extracted = await loop.run_in_executor(pool, extractor.extract, str(tmp_path))
        mapper = get_mapper()
        records = mapper.map(extracted, supplier=supplier)
        _doc_text = extracted.get("text", "")

        if not records:
            return {
                "success": False,
                "message": "Не бяха открити данни в документа.",
                "records": [],
                "output_file": None,
            }

        # Enrich records from product database (OSRAM)
        db = get_product_db(str(DATA_DIR))
        if (supplier == "osram" or (supplier == "auto" and _is_osram_document(_doc_text))) and db.is_loaded:
            enriched = 0
            for rec in records:
                am_code = getattr(rec, "_osram_article", None)
                # 1. Try supplier article (AM code)
                info = db.lookup(am_code)
                # 2. Try extracted product code when different from AM code
                if not info and rec.product_code and rec.product_code != am_code:
                    info = db.lookup(rec.product_code)
                # 3. Always also try EAN — prefer result that has an internal code
                if rec.ean:
                    ean_info = db.lookup(rec.ean)
                    if ean_info and ean_info.internal_code and (
                        not info or not info.internal_code
                    ):
                        info = ean_info
                if not info:
                    logger.info("OSRAM unmatched: code=%r  am=%r  ean=%r",
                                rec.product_code, am_code, rec.ean)
                    continue
                rec.is_new_product = False
                # Always use catalog internal code — overrides AM fallback code
                if info.internal_code and rec.extraction_method == "osram":
                    rec.product_code = info.internal_code
                if not rec.product_name and info.description:
                    rec.product_name = info.description
                if not rec.ean and (info.ean or info.main_barcode):
                    rec.ean = info.ean or info.main_barcode
                enriched += 1
            if enriched:
                logger.info("Enriched %d records from product DB", enriched)

        # Enrich records from Rezaw-Plast product database
        if supplier == "rezaw_plast":
            rp_db = get_rezaw_plast_db(str(DATA_DIR))
            if rp_db.is_loaded:
                enriched_rp = 0
                for rec in records:
                    info = rp_db.lookup(rec.product_code)
                    if not info and rec.ean:
                        info = rp_db.lookup(rec.ean)
                    if not info:
                        continue
                    rec.is_new_product = False
                    if info.description:
                        rec.product_name = info.description
                    if not rec.ean and info.ean:
                        rec.ean = info.ean
                    enriched_rp += 1
                if enriched_rp:
                    logger.info("Rezaw-Plast: enriched %d records from DB", enriched_rp)

        # Amal-Plast: invoice uses AP codes, internal system uses SL codes
        # Convert before DB lookup: AP1101 → SL1101
        if supplier == "amal_plast":
            for rec in records:
                if rec.product_code and rec.product_code.upper().startswith("AP"):
                    suffix = rec.product_code[2:]
                    if suffix.isdigit():
                        rec.product_code = "SL" + suffix

        # Car Passion: add CP- prefix so codes match the reference DB
        # Invoice: 20108 → DB: CP-20108
        if supplier == "car_passion":
            for rec in records:
                if rec.product_code and not rec.product_code.upper().startswith("CP-"):
                    rec.product_code = "CP-" + rec.product_code

        # Gumarny Zubri: resolve code against catalog — try with/without P prefix.
        # Invoice may have P217134 or 217134; catalog may store either form.
        # Use whichever variant is found in the catalog; fall back to GZ-<original>.
        if supplier == "gumarny_zubri" or (supplier == "auto" and _is_gumarny_zubri_document(_doc_text)):
            gz_name_db = get_gumarny_zubri_db(str(DATA_DIR))
            for rec in records:
                if not rec.product_code:
                    continue
                raw = rec.product_code
                # Build variants: with P and without P
                if raw.upper().startswith("P") and len(raw) > 1:
                    variants = [raw, raw[1:]]           # P217134, 217134
                else:
                    variants = [raw, "P" + raw]         # 217134, P217134
                # Try each variant with GZ- prefix in the catalog
                resolved = None
                if gz_name_db.is_loaded:
                    for v in variants:
                        if gz_name_db.lookup("GZ-" + v):
                            resolved = "GZ-" + v
                            break
                rec.product_code = resolved or ("GZ-" + raw)

        # Farad-specific code resolution.
        # Priority: code map (original invoice string) → catalog → raw invoice code.
        if supplier == "farad":
            import re as _re
            _farad_code_re = _re.compile(r'^1-([A-Z0-9]+(?:/[A-Z0-9]+)?)', _re.IGNORECASE)
            farad_db = get_farad_db(str(DATA_DIR))
            farad_map = get_farad_code_map(str(DATA_DIR))
            cat_keys = (sorted(farad_db._by_code.keys(), key=len, reverse=True)
                        if farad_db.is_loaded else [])
            mapped_n = 0
            unmapped: list[str] = []

            for rec in records:
                if not rec.product_code:
                    continue
                original = rec.product_code  # e.g. "1-Z4/E SC.NERA 1CH."

                m = _farad_code_re.match(original)
                if not m:
                    continue

                # ── Step 1: code map with the original invoice string ────────
                if farad_map.is_loaded:
                    mapped = farad_map.translate(original)
                    if mapped:
                        rec.product_code = mapped
                        mapped_n += 1
                        # Enrich name from catalog if possible
                        if farad_db.is_loaded and not rec.product_name:
                            info = farad_db.lookup(mapped)
                            if info and info.description:
                                rec.product_name = info.description
                        continue  # mapping succeeded — skip catalog lookup

                # ── Step 2: catalog matching (fallback) ──────────────────────
                catalog_str = original[2:]  # strip "1-"
                matched_key = None
                matched_remainder = ""
                for ck in cat_keys:
                    for candidate in (original, catalog_str):
                        if not candidate.upper().startswith(ck.upper()):
                            continue
                        after = candidate[len(ck):]
                        if after == "" or after[0] in (" ", "/"):
                            matched_key = ck
                            matched_remainder = after.strip()
                            break
                    if matched_key:
                        break

                if matched_key:
                    if matched_remainder:
                        rec.product_name = matched_remainder

                    # Variant refinement: catalog may have both "Z4/E" and "Z4/E BLK".
                    # If invoice contains all words of a longer variant key, use it.
                    # e.g. "Z4/E STARL. BLK 1CH" → Z4/E BLK (BLK present) wins over Z4/E.
                    if matched_remainder:
                        inv_words = set(_re.split(r'[\s.,]+', catalog_str.upper()))
                        inv_words.discard('')
                        for ck in cat_keys:
                            if not ck.upper().startswith(matched_key.upper() + " "):
                                continue
                            ck_words = set(_re.split(r'[\s.,]+', ck.upper()))
                            ck_words.discard('')
                            if ck_words and ck_words.issubset(inv_words) and len(ck) > len(matched_key):
                                matched_key = ck
                                matched_remainder = ""
                                break

                    if farad_db.is_loaded:
                        info = farad_db.lookup(matched_key)
                        base_code = (info.internal_code
                                     if info and info.internal_code
                                     else matched_key)
                        if info and info.description:
                            rec.product_name = info.description
                    else:
                        base_code = matched_key
                    rec.product_code = ("1-" + base_code
                                        if base_code and base_code[0].isdigit()
                                        else base_code)
                else:
                    # ── Step 3: no match — clean raw invoice code ────────────
                    short_key = m.group(1)
                    if short_key and short_key[0].isdigit():
                        rec.product_code = original
                    else:
                        rec.product_code = catalog_str.rstrip(".,")
                    if not rec.product_name:
                        remainder = original[len(m.group(0)):].strip()
                        if remainder:
                            rec.product_name = remainder
                    if farad_db.is_loaded:
                        info = farad_db.lookup(short_key)
                        if info and info.description:
                            rec.product_name = info.description

                unmapped.append(rec.product_code)

            logger.info("Farad code map: %d mapped, %d unmapped", mapped_n, len(unmapped))
            if unmapped:
                logger.info("Farad unmapped codes: %s", ", ".join(unmapped))

        # Areon: translate invoice description → internal code via code map.
        # The code-map file has columns AREON (descriptions) and AUTOPRO (our codes).
        # SupplierCodeMapping auto-detects them reversed, so _map = {our_code: description}.
        # We build a reverse no-spaces lookup to handle the PDF's merged-word descriptions.
        if supplier == "areon":
            import re as _re
            # Normalize Latin look-alike characters → Cyrillic so that
            # e.g. Latin K-E-N (typed in Excel) matches Cyrillic К-Е-Н (from PDF).
            _LAT_TO_CYR = {
                ord('A'): 'А', ord('B'): 'В', ord('C'): 'С', ord('E'): 'Е',
                ord('H'): 'Н', ord('K'): 'К', ord('M'): 'М', ord('O'): 'О',
                ord('P'): 'Р', ord('T'): 'Т', ord('X'): 'Х',
                ord('a'): 'а', ord('c'): 'с', ord('e'): 'е', ord('o'): 'о',
                ord('p'): 'р', ord('x'): 'х',
            }

            def _areon_norm(s: str) -> str:
                return _re.sub(r'\s+', '', s.upper().translate(_LAT_TO_CYR))

            # Read the Excel directly so that duplicate codes (same code, multiple
            # descriptions) are all preserved — _map deduplicates by code and loses entries.
            import pandas as _pd
            _areon_map_path = DATA_DIR / "areon-code-map.xlsx"
            _desc_to_code: dict[str, str] = {}
            if _areon_map_path.exists():
                try:
                    _df = _pd.read_excel(_areon_map_path, engine="openpyxl", header=0, dtype=str)
                    _df = _df.fillna("")
                    _hdrs = [str(c).lower().strip() for c in _df.columns]
                    # AREON col = invoice descriptions, AUTOPRO col = our internal codes
                    _areon_col   = next((c for c in _df.columns if 'areon'   in str(c).lower()), _df.columns[0])
                    _autopro_col = next((c for c in _df.columns if 'autopro' in str(c).lower()), _df.columns[1])
                    for _, _row in _df.iterrows():
                        _desc = str(_row[_areon_col]).strip()
                        _code = str(_row[_autopro_col]).strip()
                        if _desc and _code and _desc.lower() != 'nan' and _code.lower() != 'nan':
                            _desc_to_code[_areon_norm(_desc)] = _code
                    logger.info("Areon code map (direct): %d description→code pairs", len(_desc_to_code))
                except Exception as _e:
                    logger.warning("Areon direct map read failed (%s), falling back to _map", _e)
                    areon_map = get_areon_code_map(str(DATA_DIR))
                    if areon_map.is_loaded:
                        for _code, _desc in areon_map._map.items():
                            _desc_to_code[_areon_norm(str(_desc))] = _code
            if not _desc_to_code:
                areon_map = get_areon_code_map(str(DATA_DIR))
                if areon_map.is_loaded:
                    for _code, _desc in areon_map._map.items():
                        _desc_to_code[_areon_norm(str(_desc))] = _code

            if _desc_to_code:
                mapped_n, unmapped = 0, []
                for rec in records:
                    if not rec.product_code:
                        continue
                    nospace = _areon_norm(rec.product_code)
                    if nospace not in _desc_to_code:
                        # Forward: description is a prefix of a code-map key (description truncated)
                        prefix_hits = [k for k in _desc_to_code if k.startswith(nospace) and len(k) - len(nospace) <= 12]
                        if len(prefix_hits) == 1:
                            nospace = prefix_hits[0]
                        elif not prefix_hits:
                            # Reverse: a code-map key is a prefix of the description (extra suffix appended)
                            rev_hits = [k for k in _desc_to_code if nospace.startswith(k)]
                            if rev_hits:
                                nospace = max(rev_hits, key=len)
                    if nospace in _desc_to_code:
                        rec.product_code = _desc_to_code[nospace]
                        mapped_n += 1
                    else:
                        unmapped.append(rec.product_code[:50])
                logger.info("Areon code map: %d mapped, %d unmapped", mapped_n, len(unmapped))
                if unmapped:
                    logger.info("Areon unmapped: %s", ", ".join(unmapped[:10]))

        # Slime: enrich product name from supplier DB; found → not new
        if supplier == "slime" or (supplier == "auto" and _is_slime_document(_doc_text)):
            slime_db = get_slime_db(str(DATA_DIR))
            if slime_db.is_loaded:
                enriched_n = 0
                for rec in records:
                    if not rec.product_code:
                        continue
                    info = slime_db.lookup(rec.product_code)
                    if info:
                        rec.is_new_product = False
                        if info.description:
                            rec.product_name = info.description
                            enriched_n += 1
                if enriched_n:
                    logger.info("Slime: name DB enriched %d records", enriched_n)

        # Xado: translate invoice description → internal code via code map, then enrich name
        if supplier == "xado" or (supplier == "auto" and _is_xado_document(_doc_text)):
            import re as _re
            import pandas as _pd
            _xado_map_path = DATA_DIR / "xado-code-map.xlsx"
            _xado_desc_to_code: dict[str, str] = {}
            if _xado_map_path.exists():
                try:
                    _df = _pd.read_excel(_xado_map_path, engine="openpyxl", header=0, dtype=str)
                    _df = _df.fillna("")
                    # НАШ КОД = col 0, ПРОИЗВОДИТЕЛ КОД (invoice desc) = col 1
                    _our_col  = _df.columns[0]
                    _desc_col = _df.columns[1]
                    for _, _row in _df.iterrows():
                        _desc = str(_row[_desc_col]).strip()
                        _code = str(_row[_our_col]).strip()
                        if _desc and _code and _desc.lower() != 'nan' and _code.lower() != 'nan':
                            _xado_desc_to_code[_re.sub(r'\s+', '', _desc.upper())] = _code
                    logger.info("Xado code map: %d description→code pairs", len(_xado_desc_to_code))
                except Exception as _e:
                    logger.warning("Xado code map read failed: %s", _e)

            if _xado_desc_to_code:
                mapped_n, unmapped = 0, []
                for rec in records:
                    if not rec.product_code:
                        continue
                    nospace = _re.sub(r'\s+', '', rec.product_code.upper())
                    if nospace not in _xado_desc_to_code:
                        prefix_hits = [k for k in _xado_desc_to_code if k.startswith(nospace) and len(k) - len(nospace) <= 12]
                        if len(prefix_hits) == 1:
                            nospace = prefix_hits[0]
                        elif not prefix_hits:
                            rev_hits = [k for k in _xado_desc_to_code if nospace.startswith(k)]
                            if rev_hits:
                                nospace = max(rev_hits, key=len)
                    if nospace in _xado_desc_to_code:
                        rec.product_code = _xado_desc_to_code[nospace]
                        mapped_n += 1
                    else:
                        unmapped.append(rec.product_code[:50])
                logger.info("Xado code map: %d mapped, %d unmapped", mapped_n, len(unmapped))
                if unmapped:
                    logger.info("Xado unmapped: %s", ", ".join(unmapped[:10]))

            xado_db = get_xado_db(str(DATA_DIR))
            if xado_db.is_loaded:
                enriched_n = 0
                for rec in records:
                    if not rec.product_code:
                        continue
                    info = xado_db.lookup(rec.product_code)
                    if info:
                        rec.is_new_product = False
                        if info.description:
                            rec.product_name = info.description
                            enriched_n += 1
                if enriched_n:
                    logger.info("Xado: name DB enriched %d records", enriched_n)

        # Wunder-Baum: translate supplier code → internal code via code map
        _wb_reverse: dict[str, str] = {}  # our_code.upper() → original supplier code
        if supplier == "wunder_baum" or (supplier == "auto" and _is_wunder_baum_document(_doc_text)):
            wb_map = get_wunder_baum_code_map(str(DATA_DIR))
            if wb_map.is_loaded:
                mapped_n = 0
                unmapped = []
                for rec in records:
                    if not rec.product_code:
                        continue
                    original_code = rec.product_code
                    mapped = wb_map.translate(original_code)
                    if mapped:
                        rec.product_code = mapped
                        rec.is_new_product = False
                        _wb_reverse[mapped.upper()] = original_code
                        mapped_n += 1
                    else:
                        unmapped.append(original_code)
                logger.info("Wunder-Baum code map: %d mapped, %d unmapped", mapped_n, len(unmapped))
                if unmapped:
                    logger.info("Wunder-Baum unmapped codes: %s", ", ".join(unmapped))

        # Enrich name from supplier-specific DB
        _name_db_map = {
            "maxton_design": get_maxton_db,
            "avisa":         get_avisa_db,
            "amio":          get_amio_db,
            "mtech":         get_mtech_db,
            "mafra":         get_mafra_db,
            "amal_plast":    get_amal_plast_db,
            "car_passion":   get_car_passion_db,
            "vinove":        get_vinove_db,
            "gumarny_zubri": get_gumarny_zubri_db,
            "rigum":         get_rigum_db,
            "petex":         get_petex_db,
            "geter_hosaja":  get_geyer_hosaja_db,
            "frogum":        get_frogum_db,
            "gelly_plast":     get_gelly_plast_db,
            "farad":           get_farad_db,
            "kegel_blazusiak": get_kegel_blazusiak_db,
            "automania":       get_automania_db,
            "xado":            get_xado_db,
            "hakr":            get_hakr_db,
            "tompar":          get_tompar_db,
            "sonax":           get_senax_db,
            "team_heko":       get_heko_db,
            "bmw":             get_bmw_db,
            "wunder_baum":     get_wunder_baum_db,
            "bardahl":         get_bardahl_db,
            "areon":           get_areon_db,
        }
        _name_supplier = supplier
        if supplier == "auto" and _is_rigum_document(_doc_text):
            _name_supplier = "rigum"
        if supplier == "auto" and _is_bmw_document(_doc_text):
            _name_supplier = "bmw"
        if supplier == "auto" and _is_wunder_baum_document(_doc_text):
            _name_supplier = "wunder_baum"
        if _name_supplier in _name_db_map:
            name_db = _name_db_map[_name_supplier](str(DATA_DIR))
            if name_db.is_loaded:
                enriched_n = 0
                not_found = []
                for rec in records:
                    info = name_db.lookup(rec.product_code)
                    # Wunder-Baum: if direct lookup fails, retry with original supplier code
                    # (code map translated supplier code → internal code, but catalog uses supplier code as Код)
                    if info is None and _name_supplier == "wunder_baum" and _wb_reverse:
                        orig_sup_code = _wb_reverse.get((rec.product_code or "").upper())
                        if orig_sup_code:
                            info = name_db.lookup(orig_sup_code)
                            if info:
                                logger.debug("WB reverse lookup: %s → %s (via orig code %s)",
                                             rec.product_code, info.internal_code, orig_sup_code)
                    if info:
                        rec.is_new_product = False
                        if info.description:
                            rec.product_name = info.description
                            enriched_n += 1
                    elif _name_supplier == "wunder_baum" and rec.is_new_product:
                        not_found.append(rec.product_code or "(empty)")
                if enriched_n:
                    logger.info("%s: name DB enriched %d records", _name_supplier, enriched_n)
                if not_found and _name_supplier == "wunder_baum":
                    logger.info("wunder_baum: NOT found in DB (%d): %s",
                                len(not_found), ", ".join(not_found))

                # Ma*Fra description-based recovery: scan OCR lines and match
                # products whose code is unreadable but description is legible.
                if supplier == "mafra":
                    import re as _re
                    # Cyrillic-lookalike → Latin (OCR confuses Cyrillic with Latin)
                    _NORM = str.maketrans({
                        'А': 'A', 'В': 'B', 'С': 'C', 'Е': 'E', 'Н': 'H',
                        'К': 'K', 'М': 'M', 'О': 'O', 'Р': 'P', 'Т': 'T',
                        'Х': 'X', 'а': 'a', 'е': 'e', 'о': 'o',
                    })
                    # Common words that appear in many products → exclude from matching
                    _STOP = {
                        'dual', 'spray', 'pcs', 'pro', 'and', 'the', 'for',
                        'with', 'type', 'auto', 'car', 'ml', 'pz', 'trigger',
                        'remover', 'cleaner', 'foam', 'shampoo',
                    }
                    seen_codes = {r.product_code for r in records if r.product_code}
                    added = 0
                    for ocr_text in (extracted.get("text", ""),
                                     extracted.get("text2", "")):
                        for raw_line in ocr_text.splitlines():
                            line = _re.sub(r'[|\[\](){}]', ' ', raw_line)
                            normalized = line.translate(_NORM)
                            # Keep Latin-only words ≥3 chars, excluding stop words
                            words = [w for w in normalized.split()
                                     if _re.match(r'^[A-Za-z]{3,}$', w)
                                     and w.lower() not in _STOP]
                            if len(words) < 2:
                                continue
                            match = name_db.lookup_by_desc_words(words)
                            if not match:
                                continue
                            code, info = match
                            if code in seen_codes:
                                continue
                            seen_codes.add(code)
                            from .processors.field_mapper import ProductRecord, _clean_num
                            rec = ProductRecord(extraction_method="table")
                            rec.product_code = info.internal_code or code
                            rec.product_name = info.description
                            decimals = _re.findall(r'\d{1,6}[.,]\d{2}', line)
                            if decimals:
                                t = _clean_num(decimals[-1])
                                if t:
                                    rec.total_price = t + " EUR"
                                if len(decimals) >= 2:
                                    p = _clean_num(decimals[-2])
                                    if p and p != t:
                                        rec.price = p + " EUR"
                            records.append(rec)
                            added += 1
                            logger.info("mafra desc-match: %s → %s", words[:5], code)
                    if added:
                        logger.info("mafra: desc-match recovery added %d records", added)

        # Post-enrichment dedup: two AM codes can resolve to the same internal
        # product code after DB lookup — keep the record with the most fields.
        if any(getattr(r, "extraction_method", "") == "osram" for r in records):
            pre = len(records)
            seen: dict[str, ProductRecord] = {}
            deduped: list[ProductRecord] = []
            for rec in records:
                key = rec.product_code
                if not key:
                    deduped.append(rec)
                elif key not in seen:
                    seen[key] = rec
                    deduped.append(rec)
                elif rec.filled_count() > seen[key].filled_count():
                    deduped[deduped.index(seen[key])] = rec
                    seen[key] = rec
            records = deduped
            if len(records) < pre:
                logger.info("Post-enrichment dedup: %d → %d records", pre, len(records))

        out_path = write_excel(records, str(OUTPUT_DIR), file.filename)
        _records_cache[Path(out_path).name] = records

        text_lines = [l for l in extracted.get("text", "").splitlines() if l.strip()]
        return {
            "success": True,
            "message": f"Успешно извлечени {len(records)} записа.",
            "records": [r.to_dict() for r in records],
            "output_file": Path(out_path).name,
            "extraction_source": extracted.get("source"),
            "text_lines": len(text_lines),
            "supplier": supplier,
        }
    finally:
        tmp_path.unlink(missing_ok=True)


@app.get("/download/{filename}")
async def download(filename: str):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(404, "Файлът не е намерен.")
    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/download-custom/{filename}")
async def download_custom(filename: str, fields: str = Query(...)):
    records = _records_cache.get(filename)
    if not records:
        raise HTTPException(404, "Записите не са намерени. Моля, направете ново извличане.")
    include_fields = set(f.strip() for f in fields.split(",") if f.strip())
    loop = asyncio.get_event_loop()
    with ThreadPoolExecutor(max_workers=1) as pool:
        out_path = await loop.run_in_executor(
            pool, lambda: write_excel(records, str(OUTPUT_DIR), filename, include_fields)
        )
    out_file = Path(out_path)
    return FileResponse(
        path=str(out_file),
        filename=out_file.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/health")
async def health():
    mapper = get_mapper()
    db    = get_product_db(str(DATA_DIR))
    rp_db = get_rezaw_plast_db(str(DATA_DIR))
    return {
        "status": "ok",
        "llm_available": mapper.llm is not None,
        "osram": {
            "db_loaded":  db.is_loaded,
            "by_code":    len(db._by_internal_code),
            "by_ean":     len(db._by_ean),
        },
        "rezaw_plast": {
            "db_loaded":  rp_db.is_loaded,
            "by_code":    len(rp_db._by_code),
            "by_ean":     len(rp_db._by_ean),
        },
        "maxton": {
            "db_loaded": get_maxton_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_maxton_db(str(DATA_DIR))._by_code),
        },
        "avisa": {
            "db_loaded": get_avisa_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_avisa_db(str(DATA_DIR))._by_code),
        },
        "amio": {
            "db_loaded": get_amio_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_amio_db(str(DATA_DIR))._by_code),
        },
        "mtech": {
            "db_loaded": get_mtech_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_mtech_db(str(DATA_DIR))._by_code),
        },
        "mafra": {
            "db_loaded": get_mafra_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_mafra_db(str(DATA_DIR))._by_code),
        },
        "amal_plast": {
            "db_loaded": get_amal_plast_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_amal_plast_db(str(DATA_DIR))._by_code),
        },
        "car_passion": {
            "db_loaded": get_car_passion_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_car_passion_db(str(DATA_DIR))._by_code),
        },
        "vinove": {
            "db_loaded": get_vinove_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_vinove_db(str(DATA_DIR))._by_code),
        },
        "gumarny_zubri": {
            "db_loaded": get_gumarny_zubri_db(str(DATA_DIR)).is_loaded,
            "by_code":   len(get_gumarny_zubri_db(str(DATA_DIR))._by_code),
        },
    }


if __name__ == "__main__":
    uvicorn.run("app.main:app", host="127.0.0.1", port=5000, reload=False)
