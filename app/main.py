import logging
import os
import shutil
import signal
import threading
import time
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request
import uvicorn

from .extractors.pdf_extractor import PDFExtractor
from .extractors.excel_extractor import ExcelExtractor
from .extractors.text_extractor import TextExtractor
from .extractors.image_extractor import ImageExtractor
from .processors.field_mapper import FieldMapper, ProductRecord
from .processors.product_db import get_product_db, get_rezaw_plast_db
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
# If no heartbeat arrives for 20s, the server shuts down automatically.
_last_heartbeat: float = time.time()
_HEARTBEAT_TIMEOUT = 20  # seconds


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
        extracted = extractor.extract(str(tmp_path))
        mapper = get_mapper()
        records = mapper.map(extracted, supplier=supplier)

        if not records:
            return {
                "success": False,
                "message": "Не бяха открити данни в документа.",
                "records": [],
                "output_file": None,
            }

        # Enrich records from product database (OSRAM)
        db = get_product_db(str(DATA_DIR))
        if supplier in ("auto", "osram") and db.is_loaded:
            enriched = 0
            for rec in records:
                # 1. Try supplier article (AM code) — most specific
                info = db.lookup(getattr(rec, "_osram_article", None))
                # 2. Fallback: lookup by EAN extracted from the document
                if not info and rec.ean:
                    info = db.lookup(rec.ean)
                if not info:
                    continue
                # Product code from DB overrides extracted text for OSRAM records
                if info.internal_code and rec.extraction_method == "osram":
                    rec.product_code = info.internal_code
                if not rec.product_name and info.description:
                    rec.product_name = info.description
                if not rec.ean and (info.ean or info.main_barcode):
                    rec.ean = info.ean or info.main_barcode
                if not rec.price and info.unit_price:
                    rec.price = info.unit_price
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
                    if not rec.product_name and info.description:
                        rec.product_name = info.description
                    if not rec.ean and info.ean:
                        rec.ean = info.ean
                    if not rec.price and info.unit_price:
                        rec.price = info.unit_price
                    enriched_rp += 1
                if enriched_rp:
                    logger.info("Rezaw-Plast: enriched %d records from DB", enriched_rp)

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
    }


if __name__ == "__main__":
    uvicorn.run("app.main:app", host="127.0.0.1", port=5000, reload=False)
