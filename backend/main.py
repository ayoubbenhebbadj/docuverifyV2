"""
DocuVerify AI - import/export trade compliance verification backend.

Run (PowerShell):
    $env:OPENROUTER_API_KEY = "your-key"
    python -m uvicorn main:app --reload --port 8000

Env:
    OPENROUTER_API_KEY   required
    OPENROUTER_MODEL     default google/gemini-2.5-flash
    MAX_OUTPUT_TOKENS    default 8000
    REASONING_TOKENS     default 1024 (0 = don't send a reasoning limit)
"""
import base64
import asyncio
import contextvars
import secrets
import time
from collections import defaultdict, deque
import io
import json
import logging
import os
import re
import unicodedata
from datetime import date
from pathlib import Path
from typing import List

import httpx
import pymupdf as fitz
from docx import Document
from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool
from typing import Optional

try:  # optional: load OPENROUTER_API_KEY from a .env file next to main.py
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env")
except ImportError:
    pass

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("docuverify")

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = os.getenv("OPENROUTER_MODEL", "google/gemini-2.5-flash")
MAX_FILES = 8
MAX_FILE_BYTES = int(os.getenv("MAX_FILE_MB", "20")) * 1024 * 1024
MAX_TOTAL_PAGES = int(os.getenv("MAX_TOTAL_PAGES", "40"))   # page-image cap per request
PDF_ZOOM = float(os.getenv("PDF_ZOOM", "2.0"))        # 2.0 = about 144 DPI
REQUEST_TIMEOUT = 240.0
MAX_OUTPUT_TOKENS = int(os.getenv("MAX_OUTPUT_TOKENS", "8000"))
REASONING_TOKENS = int(os.getenv("REASONING_TOKENS", "1024"))
MAX_ATTEMPTS = int(os.getenv("MAX_ATTEMPTS", "3"))

REQUIRED_DOCS = [
    "Commercial Invoice",
    "Packing List",
    "Bill of Lading (B/L)",
    "Certificate of Origin",
    "Certificate of Quality",
    "Certificate of Conformity",
    "Phytosanitary Certificate",
    "Fumigation Certificate",
]
CROSS_FIELDS = [
    "Exporter", "Importer", "Invoice Number & Date", "B/L Number",
    "Number of Packages", "Weights (Gross/Net)", "Ports (Loading/Discharge)",
    "Container Numbers",
]

SYSTEM_PROMPT = f"""You are a senior customs and trade-compliance auditor. You receive a batch of
international trade documents as page images (and some extracted Word text). Documents may mix
English and Arabic, and contain stamps, seals, handwriting, tables and signatures. Inspect them visually.

Audit the batch for these 8 mandatory documents, in this exact order:
{json.dumps(REQUIRED_DOCS)}

Then cross-validate these fields across all documents that mention them, in this exact order:
{json.dumps(CROSS_FIELDS)}

Rules:
- A document is "Present" only if you can actually see it in the batch. Never assume. Otherwise "Absent".
- Identify each document by the title printed on it, in any language. Arabic equivalents:
  Commercial Invoice = فاتورة تجارية; Packing List = بيان عبوة / قائمة التعبئة;
  Bill of Lading = بوليصة شحن; Certificate of Origin = شهادة منشأ;
  Certificate of Quality = شهادة جودة / Quality certificate;
  Certificate of Conformity = شهادة مطابقة; Phytosanitary Certificate = شهادة صحة نباتية;
  Fumigation Certificate = شهادة تبخير.
- Each page is one document. Use the exact page numbers from the "--- File: ... | Page N of M ---" labels.
  Put them in "page_or_file" (e.g. "file.pdf p.3"); use "-" if Absent.
- Every page image carries a small red 'PAGE n' tag in its top-left corner. Use that number.
- Some pages also come with a text layer. Use it to read titles and numbers, but judge stamps and
  handwriting from the image only.
- In "notes", state the printed title, then stamps and signatures. Handwritten signatures and ink stamps are
  easy to invent: only claim one if you clearly see ink on that page. Typed or Word-generated pages normally
  have none, so write 'no signature or stamp visible'. Never borrow details from another page.
- Cross-validation status: "Match" (identical everywhere), "Partial" (minor differences such as spelling,
  typos, formatting, or a field found in only some documents), "Mismatch" (material conflict).
- In "extracted_values", give the value per document, e.g. Invoice: 1,250 kg | B/L: 1,205 kg.
- Date logic. Find the shipment date: the B/L date, or if no B/L is in the batch, the shipment / on-board / loading
  date printed on any other document (for example the shipment date in box 6 of a Certificate of Origin, or a date
  next to the B/L number in a conformity letter). Compare every certificate's issue date with it. Flag any
  certificate issued AFTER the shipment date and state both dates. A certificate dated after the invoice is
  normal and must not be flagged. Also flag a certificate or inspection date earlier than the invoice date by
  more than a few days, stating both dates.
- The invoice date is never the shipment date. An invoice date quoted inside another document (for example
  box 10 of a Certificate of Origin) is the invoice date. A certificate's own serial or certificate number
  (for example the number printed at the top of a Certificate of Origin) is not an invoice number.
- Retroactive issue. Always flag a stamp or text meaning 'issued retroactively' or 'issued retrospectively'
  (Arabic: استخرجت بأثر رجعي) and name the document.
- Treatment consistency. If a Phytosanitary Certificate has blank or *** treatment fields while another document
  reports a fumigation or treatment, flag it and give the treatment date.
- Arabic tables read right to left: the FIRST column is the rightmost one. In packing lists,
  الوزن القائم = gross weight, الوزن الصافي = net weight, الأعداد = number of packages,
  الشركة المصدرة = exporting company (rightmost box), الشركة المستوردة = importing company.
  Gross weight is never lower than net weight; if your reading says otherwise, re-read the columns.
- Cross-checking. Use every document that shows a field, including the Certificate of Origin, the Phytosanitary
  Certificate and the Certificate of Conformity. Put a short document name before each value.
  If a field appears in only one document, set status to "Partial" and begin details with 'Only found in one document'.
- Do not repeat missing documents in action_items; the server adds those.
- Flag in action_items: incoterm vs freight-term contradictions, container-number typos, and unsigned,
  undated or unstamped certificates (say which ones). Keep each action item under 200 characters and
  state the dates or values involved.
- overall_status: "Approved" only if all 8 are Present and no Mismatch exists; "Rejected" if
  critical documents are missing or there are material mismatches; otherwise "Needs Review".
- Do not invent data. If a value is unreadable, say so.

JSON formatting rules (critical):
- Standard JSON: every key and every string value is wrapped in DOUBLE quotes.
- Inside the text of a value, never use a double quote character; write quoted words with single quotes, for example "notes": "Printed title 'Commercial Invoice'. No signature or stamp visible."
- No line breaks inside strings, no trailing commas, no comments.
- Keep every notes, details and extracted_values string under 200 characters.

Respond with ONE JSON object only, no prose, no markdown, exactly this shape:
{{
  "overall_status": "Approved" | "Needs Review" | "Rejected",
  "summary": "1-2 sentence executive summary",
  "required_documents": [{{"name": str, "status": "Present" | "Absent", "page_or_file": str, "notes": str}}],
  "cross_validation": [{{"field": str, "status": "Match" | "Mismatch" | "Partial", "extracted_values": str, "details": str}}],
  "action_items": [str]
}}"""

# --------------------------------------------------------------------------- deployment settings
APP_ENV = os.getenv("APP_ENV", "development")          # set to "production" when deployed
APP_USER = os.getenv("APP_USER", "docuverify")
APP_PASSWORD = os.getenv("APP_PASSWORD", "")           # required in production
RATE_LIMIT_PER_HOUR = int(os.getenv("RATE_LIMIT_PER_HOUR", "20"))  # audits per IP per hour, 0 = off
MAX_CONCURRENT_AUDITS = int(os.getenv("MAX_CONCURRENT_AUDITS", "2"))
TRUST_PROXY = os.getenv("TRUST_PROXY", "0") == "1"     # 1 when behind Caddy / Render / a load balancer
ALLOWED_ORIGINS = [o.strip() for o in os.getenv(
    "ALLOWED_ORIGINS", "http://localhost:8000,http://127.0.0.1:8000").split(",") if o.strip()]
ALLOWED_ORIGIN_REGEX = os.getenv("ALLOWED_ORIGIN_REGEX", "")   # e.g. https://docuverify-.*\\.vercel\\.app
LOG_CONTENT = os.getenv("LOG_CONTENT", "0" if APP_ENV == "production" else "1") == "1"

if APP_ENV == "production" and not APP_PASSWORD:
    logging.getLogger("docuverify").warning(
        "APP_PASSWORD is empty: anyone who knows the API address can run audits. Rate limits still apply.")


def _c(value):
    """Document text must not end up in production logs unless LOG_CONTENT=1."""
    return value if LOG_CONTENT else "[hidden]"


# --------------------------------------------------------------------------- languages
LANG = contextvars.ContextVar("lang", default="en")

TX = {
    "en": {
        "no_files": "No files uploaded.",
        "max_files": "Maximum {n} files per batch.",
        "empty": "'{name}' is empty.",
        "too_big": "'{name}' exceeds {mb} MB.",
        "unsupported": "Unsupported file type: '{name}'. Use PDF, PNG, JPG or DOCX.",
        "bad_pdf": "'{name}' is not a readable PDF: {err}",
        "pw": "'{name}' is password protected.",
        "bad_docx": "'{name}' is not a readable .docx: {err}",
        "no_key": "Server is missing the OPENROUTER_API_KEY environment variable.",
        "timeout": "The AI model took too long to respond. Try fewer or smaller files.",
        "unreachable": "Could not reach OpenRouter: {err}",
        "or_status": "OpenRouter returned {status}: {body}",
        "or_unexpected": "OpenRouter returned an unexpected response.",
        "or_rate": "The free AI model is rate-limited right now (free models allow only a limited number of requests per minute and per day). Wait a minute and try again.",
        "or_credit": "OpenRouter refused the request for credit reasons (error 402). Check your OpenRouter balance; even free models need a non-negative balance.",
        "r_length": "the response hit the token limit",
        "r_empty": "the model returned an empty answer",
        "r_json": "the model returned invalid JSON",
        "r_unknown": "unknown error",
        "failed": "The AI audit failed after several attempts: {reason}. Try again, raise MAX_OUTPUT_TOKENS, or use a different model.",
        "sum_found": "{n} of {total} mandatory documents found.",
        "sum_missing": "Missing: {items}.",
        "sum_mismatch": "Mismatch in: {items}.",
        "sum_minor": "Minor differences in: {items}.",
        "sum_ok": "All cross-checked fields agree.",
        "obtain": "Obtain and upload the missing {name}.",
        "retro_text": "Text on {where} indicates the document was issued retroactively; check the issue date and the reason.",
        "retro_scan": "{who} bears an 'issued retroactively' stamp. Ask why and check the dates.",
        "late": "{who} was issued on {issue}, after the shipment date {ship}.",
        "scan_note": "Found on a scanned page in a second check; confirm visually.",
        "gross_below": "Gross below net in '{seg}': check the column reading. ",
        "rate": "Too many audits from your network. Please wait a while and try again.",
        "busy": "The server is busy with other audits. Try again in a minute.",
    },
    "fr": {
        "no_files": "Aucun fichier téléversé.",
        "max_files": "Maximum {n} fichiers par lot.",
        "empty": "« {name} » est vide.",
        "too_big": "« {name} » dépasse {mb} Mo.",
        "unsupported": "Type de fichier non pris en charge : « {name} ». Utilisez PDF, PNG, JPG ou DOCX.",
        "bad_pdf": "« {name} » n'est pas un PDF lisible : {err}",
        "pw": "« {name} » est protégé par un mot de passe.",
        "bad_docx": "« {name} » n'est pas un fichier .docx lisible : {err}",
        "no_key": "La variable d'environnement OPENROUTER_API_KEY est absente sur le serveur.",
        "timeout": "Le modèle d'IA a mis trop de temps à répondre. Essayez avec moins de fichiers ou des fichiers plus légers.",
        "unreachable": "Impossible de joindre OpenRouter : {err}",
        "or_status": "OpenRouter a renvoyé {status} : {body}",
        "or_unexpected": "OpenRouter a renvoyé une réponse inattendue.",
        "or_rate": "Le modèle d'IA gratuit est limité pour le moment (les modèles gratuits n'autorisent qu'un nombre restreint de requêtes par minute et par jour). Attendez une minute et réessayez.",
        "or_credit": "OpenRouter a refusé la requête pour une raison de crédit (erreur 402). Vérifiez votre solde OpenRouter ; même les modèles gratuits exigent un solde non négatif.",
        "r_length": "la réponse a atteint la limite de jetons",
        "r_empty": "le modèle a renvoyé une réponse vide",
        "r_json": "le modèle a renvoyé un JSON invalide",
        "r_unknown": "erreur inconnue",
        "failed": "L'audit par IA a échoué après plusieurs tentatives : {reason}. Réessayez, augmentez MAX_OUTPUT_TOKENS ou utilisez un autre modèle.",
        "sum_found": "{n} documents obligatoires sur {total} trouvés.",
        "sum_missing": "Manquant(s) : {items}.",
        "sum_mismatch": "Écart sur : {items}.",
        "sum_minor": "Différences mineures sur : {items}.",
        "sum_ok": "Tous les champs contrôlés concordent.",
        "obtain": "Obtenir et téléverser le document manquant : {name}.",
        "retro_text": "Le texte de {where} indique que le document a été délivré avec effet rétroactif ; vérifier la date de délivrance et le motif.",
        "retro_scan": "{who} porte un cachet « délivré avec effet rétroactif ». Demander le motif et vérifier les dates.",
        "late": "{who} a été délivré le {issue}, après la date d'expédition {ship}.",
        "scan_note": "Trouvé sur une page scannée lors d'un second contrôle ; à confirmer visuellement.",
        "gross_below": "Poids brut inférieur au net dans « {seg} » : vérifier la lecture des colonnes. ",
        "rate": "Trop d'audits depuis votre réseau. Veuillez patienter un moment puis réessayer.",
        "busy": "Le serveur est occupé par d'autres audits. Réessayez dans une minute.",
    },
}

DOC_FR = {
    "Commercial Invoice": "Facture commerciale",
    "Packing List": "Liste de colisage",
    "Bill of Lading (B/L)": "Connaissement (B/L)",
    "Certificate of Origin": "Certificat d'origine",
    "Certificate of Quality": "Certificat de qualité",
    "Certificate of Conformity": "Certificat de conformité",
    "Phytosanitary Certificate": "Certificat phytosanitaire",
    "Fumigation Certificate": "Certificat de fumigation",
}
FIELD_FR = {
    "Exporter": "Exportateur",
    "Importer": "Importateur",
    "Invoice Number & Date": "N° et date de facture",
    "B/L Number": "N° de connaissement (B/L)",
    "Number of Packages": "Nombre de colis",
    "Weights (Gross/Net)": "Poids (brut/net)",
    "Ports (Loading/Discharge)": "Ports (chargement/déchargement)",
    "Container Numbers": "N° de conteneurs",
}
LANG_NOTE = {
    "en": "\n\nLanguage: write every free-text value (summary, notes, details, action_items) in English.",
    "fr": ("\n\nLanguage: write every free-text value (summary, notes, details, action_items) in French, "
           "using precise customs and trade vocabulary. Keep the JSON keys, the status values, the document "
           "names and the field names exactly as specified above, in English. Keep names, numbers, dates and "
           "quoted printed text exactly as they appear on the documents."),
}


def tr(key: str, **kw) -> str:
    table = TX.get(LANG.get(), TX["en"])
    return table.get(key, TX["en"][key]).format(**kw)


def dn(name: str) -> str:
    return DOC_FR.get(name, name) if LANG.get() == "fr" else name


def fn(field: str) -> str:
    return FIELD_FR.get(field, field) if LANG.get() == "fr" else field


def system_for_lang() -> str:
    return SYSTEM_PROMPT + LANG_NOTE.get(LANG.get(), LANG_NOTE["en"])


app = FastAPI(title="DocuVerify AI", version="1.8.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_origin_regex=ALLOWED_ORIGIN_REGEX or None,
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


@app.middleware("http")
async def security_headers(request: Request, call_next):
    resp = await call_next(request)
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "DENY")
    resp.headers.setdefault("Referrer-Policy", "no-referrer")
    if request.url.path.startswith("/api/"):
        resp.headers["Cache-Control"] = "no-store"
    return resp


BASE_DIR = Path(__file__).parent
# For local development the API can also serve the page from ../frontend. On Render only the backend exists.
FRONTEND_DIR = next((d for d in (BASE_DIR.parent / "frontend", BASE_DIR / "frontend", BASE_DIR)
                     if (d / "index.html").is_file()), None)
if FRONTEND_DIR and (FRONTEND_DIR / "static").is_dir():
    app.mount("/static", StaticFiles(directory=FRONTEND_DIR / "static"), name="static")

_basic = HTTPBasic(auto_error=False)


def require_auth(creds: Optional[HTTPBasicCredentials] = Depends(_basic)) -> None:
    """Shared login via the browser's built-in sign-in box. Disabled when APP_PASSWORD is empty (local dev)."""
    if not APP_PASSWORD:
        return
    ok = (creds is not None
          and secrets.compare_digest(creds.username.encode(), APP_USER.encode())
          and secrets.compare_digest(creds.password.encode(), APP_PASSWORD.encode()))
    if not ok:
        raise HTTPException(401, "Authentication required",
                            headers={"WWW-Authenticate": 'Basic realm="DocuVerify AI"'})


_hits = defaultdict(deque)
AUDIT_SLOTS = asyncio.Semaphore(max(MAX_CONCURRENT_AUDITS, 1))


def client_ip(request: Request) -> str:
    if TRUST_PROXY:
        fwd = request.headers.get("x-forwarded-for")
        if fwd:
            return fwd.split(",")[0].strip()
    return request.client.host if request.client else "unknown"


def check_rate(ip: str) -> None:
    if RATE_LIMIT_PER_HOUR <= 0:
        return
    now, q = time.time(), _hits[ip]
    while q and now - q[0] > 3600:
        q.popleft()
    if len(q) >= RATE_LIMIT_PER_HOUR:
        raise HTTPException(429, tr("rate"))
    q.append(now)



# --------------------------------------------------------------------------- parsing
HINT_CHARS = 2500


def _data_url(mime: str, raw: bytes) -> str:
    return f"data:{mime};base64,{base64.b64encode(raw).decode()}"


def parse_pdf(raw: bytes, filename: str, budget: int):
    """Returns (content_parts, page_records). Each page is sent as an image with a burned-in
    'PAGE n' tag, plus its text layer (when the PDF has one) as a hint."""
    parts, pages = [], []
    try:
        doc = fitz.open(stream=raw, filetype="pdf")
    except Exception as exc:
        raise HTTPException(400, tr("bad_pdf", name=filename, err=exc))
    with doc:
        if doc.needs_pass:
            raise HTTPException(400, tr("pw", name=filename))
        total = len(doc)
        for i, page in enumerate(doc):
            if i >= budget:
                parts.append({"type": "text",
                              "text": f"[{filename}: {total - i} further page(s) omitted, page limit reached]"})
                break
            text = page.get_text("text") or ""
            try:  # tag lives only in the in-memory render; the upload is never modified
                page.insert_text((6, 12), f"PAGE {i + 1}", fontsize=9, color=(1, 0, 0))
            except Exception:
                pass
            pix = page.get_pixmap(matrix=fitz.Matrix(PDF_ZOOM, PDF_ZOOM), alpha=False)
            parts.append({"type": "text", "text": f"--- File: {filename} | Page {i + 1} of {total} ---"})
            parts.append({"type": "image_url",
                          "image_url": {"url": _data_url("image/png", pix.tobytes("png"))}})
            if text.strip():
                parts.append({"type": "text",
                              "text": f"[Text layer of page {i + 1}; may be incomplete or garbled for Arabic]\n"
                                      f"{text.strip()[:HINT_CHARS]}"})
            log.info("%s p.%d/%d text_chars=%d head=%r", filename, i + 1, total, len(text.strip()),
                     _c(text.strip()[:70]))
            pages.append({"file": filename, "page": i + 1, "text": text,
                          "image": "" if text.strip() else parts[-1]["image_url"]["url"]})
    return parts, pages


def parse_docx(raw: bytes, filename: str):
    try:
        doc = Document(io.BytesIO(raw))
    except Exception as exc:
        raise HTTPException(400, tr("bad_docx", name=filename, err=exc))
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for t_idx, table in enumerate(doc.tables, 1):
        lines.append(f"[Table {t_idx}]")
        for row in table.rows:
            lines.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
    text = "\n".join(lines) or "[No extractable text in this Word document]"
    parts = [{"type": "text", "text": f"--- File: {filename} (Word document, extracted text) ---\n{text[:60000]}"}]
    return parts, [{"file": filename, "page": 1, "text": text}]


def parse_image(raw: bytes, filename: str, mime: str):
    parts = [
        {"type": "text", "text": f"--- File: {filename} (image) ---"},
        {"type": "image_url", "image_url": {"url": _data_url(mime, raw)}},
    ]
    return parts, [{"file": filename, "page": 1, "text": "",
                    "image": parts[1]["image_url"]["url"]}]


def build_content(files: list):
    """files: list of (filename, bytes). Returns (multimodal content array, page records)."""
    content = [{"type": "text", "text":
                f"Audit this batch of {len(files)} file(s). Files: " + ", ".join(n for n, _ in files)}]
    pages_all = []
    budget = MAX_TOTAL_PAGES
    for name, raw in files:
        ext = Path(name).suffix.lower()
        if ext == ".pdf":
            parts, pages = parse_pdf(raw, name, max(budget, 0))
            budget -= sum(1 for p in parts if p["type"] == "image_url")
        elif ext == ".docx":
            parts, pages = parse_docx(raw, name)
        elif ext in (".png", ".jpg", ".jpeg"):
            parts, pages = parse_image(raw, name, "image/png" if ext == ".png" else "image/jpeg")
            budget -= 1
        else:
            raise HTTPException(400, tr("unsupported", name=name))
        content.extend(parts)
        pages_all.extend(pages)
    return content, pages_all


# --------------------------------------------------------------------------- AI call
_SQ_VALUE = re.compile(r"(?<=[:\[,])(\s*)'((?:[^'\\\n]|\\.)*)'(?=\s*[,}\]])")


def _repair(text: str) -> str:
    """Fix common model slips: single-quoted values and trailing commas."""
    text = _SQ_VALUE.sub(lambda m: m.group(1) + json.dumps(m.group(2).replace("\\'", "'"), ensure_ascii=False), text)
    return re.sub(r",\s*([}\]])", r"\1", text)


def extract_json(text: str) -> dict:
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I)
    candidates = [text]
    m = re.search(r"\{.*\}", text, re.S)
    if m and m.group(0) != text:
        candidates.append(m.group(0))
    err = None
    for cand in candidates:
        for variant in (cand, _repair(cand)):
            try:
                return json.loads(variant)
            except json.JSONDecodeError as exc:
                err = exc
    raise err


REASONING_EFFORT = os.getenv("REASONING_EFFORT", "low")  # for :free models; "off" disables


def build_payload(content: list, max_tokens: int, use_reasoning: bool = True,
                  system: str = SYSTEM_PROMPT) -> dict:
    payload = {
        "model": MODEL,
        "temperature": 0,
        "max_tokens": max_tokens,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": content},
        ],
    }
    if not MODEL.endswith(":free"):
        payload["response_format"] = {"type": "json_object"}
        if use_reasoning and REASONING_TOKENS > 0:
            payload["reasoning"] = {"max_tokens": REASONING_TOKENS}
    elif use_reasoning and REASONING_EFFORT != "off":
        payload["reasoning"] = {"effort": REASONING_EFFORT}
    return payload


async def call_openrouter(content: list, system: str = SYSTEM_PROMPT, start_tokens: int = 0) -> dict:
    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        raise HTTPException(500, tr("no_key"))
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": os.getenv("APP_URL", "http://localhost:8000"),
        "X-Title": "DocuVerify AI",
    }

    max_tokens = start_tokens or MAX_OUTPUT_TOKENS
    use_reasoning = True
    last_problem = "r_unknown"
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            async with httpx.AsyncClient(timeout=REQUEST_TIMEOUT) as client:
                resp = await client.post(OPENROUTER_URL, json=build_payload(content, max_tokens, use_reasoning, system), headers=headers)
        except httpx.TimeoutException:
            raise HTTPException(504, tr("timeout"))
        except httpx.HTTPError as exc:
            raise HTTPException(502, tr("unreachable", err=exc))

        if resp.status_code == 400 and use_reasoning and "reasoning" in resp.text.lower():
            log.warning("Model rejected the reasoning setting; retrying without it.")
            use_reasoning = False
            continue
        if resp.status_code in (402, 429):
            log.error("OpenRouter %s: %s", resp.status_code, resp.text[:300])
            raise HTTPException(429 if resp.status_code == 429 else 502,
                                tr("or_rate" if resp.status_code == 429 else "or_credit"))
        if resp.status_code != 200:
            log.error("OpenRouter %s: %s", resp.status_code, resp.text[:500])
            raise HTTPException(502, tr("or_status", status=resp.status_code, body=resp.text[:300]))

        try:
            data = resp.json()
            choice = data["choices"][0]
        except Exception:
            log.error("Unexpected OpenRouter body: %s", resp.text[:500])
            raise HTTPException(502, tr("or_unexpected"))

        raw_text = (choice.get("message") or {}).get("content") or ""
        finish = choice.get("finish_reason")
        log.info("attempt=%d model=%s finish_reason=%s usage=%s chars=%d",
                 attempt, MODEL, finish, data.get("usage"), len(raw_text))

        if finish == "length":
            last_problem = "r_length"
            log.warning("Truncated output: %r", _c(raw_text[:300]))
            max_tokens = min(max_tokens * 2, 16000)
            continue
        if not raw_text.strip():
            last_problem = "r_empty"
            log.warning("Empty content. Full choice: %s", _c(json.dumps(choice)[:500]))
            continue
        try:
            return extract_json(raw_text)
        except json.JSONDecodeError as exc:
            last_problem = "r_json"
            log.warning("Invalid JSON (%s). Raw start: %r", exc, _c(raw_text[:500]))
            max_tokens = min(int(max_tokens * 1.5), 16000)
            continue

    raise HTTPException(502, tr("failed", reason=tr(last_problem)))


def normalize(result: dict) -> dict:
    """Guarantee schema shape, the 8 documents in order, and valid enum values."""
    status = result.get("overall_status")
    if status not in ("Approved", "Needs Review", "Rejected"):
        status = "Needs Review"

    by_name = {}
    for d in result.get("required_documents") or []:
        if isinstance(d, dict):
            by_name[str(d.get("name", "")).strip().lower()] = d

    docs = []
    for name in REQUIRED_DOCS:
        d = by_name.get(name.lower()) or next(
            (v for k, v in by_name.items() if k and (k in name.lower() or name.lower() in k)), None)
        d = d or {}
        present = d.get("status") == "Present"
        docs.append({
            "name": name,
            "status": "Present" if present else "Absent",
            "page_or_file": str(d.get("page_or_file") or "-"),
            "notes": str(d.get("notes") or ""),
        })

    cross = []
    for c in result.get("cross_validation") or []:
        if not isinstance(c, dict):
            continue
        st = c.get("status") if c.get("status") in ("Match", "Mismatch", "Partial") else "Partial"
        cross.append({
            "field": str(c.get("field", "")),
            "status": st,
            "extracted_values": str(c.get("extracted_values", "")),
            "details": str(c.get("details", "")),
        })

    actions = [str(a) for a in (result.get("action_items") or []) if str(a).strip()]

    # Server-side safety net: never "Approved" when something is missing or conflicting.
    if status == "Approved" and (any(d["status"] == "Absent" for d in docs)
                                 or any(c["status"] == "Mismatch" for c in cross)):
        status = "Needs Review"

    return {
        "overall_status": status,
        "summary": str(result.get("summary") or ""),
        "required_documents": docs,
        "cross_validation": cross,
        "action_items": actions,
    }


# --------------------------------------------------------------------------- cross-checks
_AR_MAP = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي", "ة": "ه"})


def norm(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "").lower().replace("\u0640", "")
    text = re.sub("[\u064b-\u065f]", "", text)
    return text.translate(_AR_MAP)


# Order matters: specific certificates first, generic invoice / B/L last.
DETECT = {
    "Fumigation Certificate": r"fumigation certificate|شهادة تبخير",
    "Phytosanitary Certificate": r"phytosanitary|صحة نباتية",
    "Certificate of Quality": r"quality certificate|certificate of quality|شهادة جودة",
    "Certificate of Conformity": r"certificate of conformity|conformity certificate|شهادة مطابقة",
    "Certificate of Origin": r"certificate of origin|شهادة منشا",
    "Packing List": r"packing list|بيان عبوة|قائمة التعبئة",
    "Commercial Invoice": r"commercial invoice|فاتورة تجارية|invoice no\s*:\s*\d",
    "Bill of Lading (B/L)": r"bill of lading|بوليصة شحن",
}


def detect_docs(pages: list) -> dict:
    """Title-based detection on pages that have a text layer. Only ever used to CONFIRM a
    document (scanned pages have no text and are left to the vision model)."""
    found = {}
    for p in pages:
        head = norm(p["text"])[:400]
        if not head.strip():
            continue
        for name, pat in DETECT.items():
            if re.search(norm(pat), head) or re.search(norm(pat), head[::-1]):
                found.setdefault(name, []).append(f"{p['file']} p.{p['page']}")
                break
    log.info("Text-layer detection: %s", found)
    return found


RETRO = re.compile(norm(r"retroactiv|retrospectiv|بأثر رجعي"))


def retro_flags(pages: list) -> list:
    """Backstop for the model: flag 'issued retroactively' wording found in a page's text layer."""
    out = []
    for p in pages:
        t = norm(p["text"])
        if RETRO.search(t):
            out.append(tr("retro_text", where=f"{p['file']} p.{p['page']}"))
    return out


_WEIGHTS = re.compile(r"g\.?\s?w\.?\s*[:=]?\s*([\d][\d,\.]*).{0,25}?n\.?\s?w\.?\s*[:=]?\s*([\d][\d,\.]*)", re.I)


def _num(text: str):
    try:
        return float(text.replace(",", "").rstrip("."))
    except ValueError:
        return None


def check_weights(result: dict) -> None:
    """Gross weight can never be below net weight. A reading like that means a swapped column."""
    for c in result["cross_validation"]:
        if "weight" not in c["field"].lower():
            continue
        for seg in c["extracted_values"].split("|"):
            m = _WEIGHTS.search(seg)
            if m:
                g, n = _num(m.group(1)), _num(m.group(2))
                if g is not None and n is not None and g < n:
                    if c["status"] == "Match":
                        c["status"] = "Partial"
                    c["details"] = (tr("gross_below", seg=seg.strip()) + c["details"])[:400]
                    break


def reconcile(result: dict, found: dict) -> dict:
    """Merge text-layer detections, then derive verdict and summary from the checklist so the
    banner, the checklist and the summary can never contradict each other."""
    check_weights(result)
    for d in result["required_documents"]:
        if d["name"] in found:
            d["status"] = "Present"
            d["page_or_file"] = ", ".join(found[d["name"]])
        elif d["status"] == "Absent":
            d["page_or_file"] = "-"

    absent = [d["name"] for d in result["required_documents"] if d["status"] == "Absent"]
    mismatches = [c["field"] for c in result["cross_validation"] if c["status"] == "Mismatch"]
    partials = [c["field"] for c in result["cross_validation"] if c["status"] == "Partial"]
    present_n = len(result["required_documents"]) - len(absent)

    if absent or mismatches:
        result["overall_status"] = "Rejected"
    elif partials or result["action_items"]:
        result["overall_status"] = "Needs Review"
    else:
        result["overall_status"] = "Approved"

    parts = [tr("sum_found", n=present_n, total=len(result["required_documents"]))]
    if absent:
        parts.append(tr("sum_missing", items=", ".join(dn(x) for x in absent)))
    if mismatches:
        parts.append(tr("sum_mismatch", items=", ".join(fn(x) for x in mismatches)))
    if partials:
        parts.append(tr("sum_minor", items=", ".join(fn(x) for x in partials)))
    if not (absent or mismatches or partials):
        parts.append(tr("sum_ok"))
    result["summary"] = " ".join(parts)

    for name in reversed(absent):
        keys = {name.split(" (")[0].lower(), DOC_FR.get(name, name).split(" (")[0].lower()}
        if not any(k in a.lower() for a in result["action_items"] for k in keys):
            result["action_items"].insert(0, tr("obtain", name=dn(name)))
    return result


CLASSIFY_PROMPT = f"""You inspect scanned trade-document pages. Pages are labelled '--- File: NAME | Page N ---'.
For every page return these fields:
- document: one of {json.dumps(REQUIRED_DOCS)}, or "Other". Judge by the printed title in any language.
  Arabic titles: شهادة منشأ = Certificate of Origin; بوليصة شحن = Bill of Lading; شهادة مطابقة = Certificate of Conformity;
  شهادة جودة = Certificate of Quality; بيان عبوة = Packing List; فاتورة تجارية = Commercial Invoice;
  شهادة صحة نباتية = Phytosanitary Certificate; شهادة تبخير = Fumigation Certificate.
- issue_date: the date this document was issued or certified, as YYYY-MM-DD, or "" if unreadable.
  This is NOT the shipment date, inspection date or invoice date.
- shipment_date: a shipment, B/L, on-board or loading date written on this page, as YYYY-MM-DD, or "".
- retroactive: true only if a stamp or text says the document was issued retroactively or retrospectively
  (Arabic: استخرجت بأثر رجعي), otherwise false.
- title_seen: the title text you read, short.
Read day and month carefully. Never guess a date; use "" when unsure.
Reply with ONE JSON object using double quotes and no double quotes inside values:
{{"pages": [{{"file": "NAME", "page": 1, "document": "Other", "issue_date": "", "shipment_date": "", "retroactive": false, "title_seen": ""}}]}}"""


def _d(value):
    try:
        return date.fromisoformat(str(value).strip())
    except Exception:
        return None


def _mentioned(actions: list, *groups) -> bool:
    """True if one action item contains, for every group, at least one of its alternatives."""
    def has(a, g):
        g = (g,) if isinstance(g, str) else g
        return any(t.lower() in a.lower() for t in g)
    return any(all(has(a, g) for g in groups) for a in actions)


async def inspect_scanned(result: dict, pages: list, found: dict) -> None:
    """Scanned pages have no text layer. Ask the model narrow questions about only those pages
    (which document, issue date, shipment date, retroactive stamp) and do the date comparison here
    in code. Never fatal: on any failure the main audit stands."""
    scanned = [p for p in pages if not p["text"].strip() and p.get("image")][:8]
    if not scanned:
        return
    content = [{"type": "text", "text": "Inspect these pages."}]
    for p in scanned:
        content.append({"type": "text", "text": f"--- File: {p['file']} | Page {p['page']} ---"})
        content.append({"type": "image_url", "image_url": {"url": p["image"]}})
    try:
        out = await call_openrouter(content, CLASSIFY_PROMPT)
    except Exception as exc:
        log.warning("Scanned-page inspection skipped: %s", getattr(exc, "detail", exc))
        return
    log.info("Scanned-page inspection: %s", _c(out))

    items = [i for i in (out.get("pages") or []) if isinstance(i, dict)]
    by_name = {d["name"]: d for d in result["required_documents"]}
    known = {f"{p['file']} p.{p['page']}" for p in scanned}

    def ref_of(item) -> str:
        ref = f"{item.get('file') or scanned[0]['file']} p.{item.get('page')}"
        if ref not in known:
            ref = next((k for k in known if k.endswith(f" p.{item.get('page')}")), ref)
        return ref

    # 1. documents the first pass called Absent but that appear on a scanned page
    for item in items:
        name = item.get("document")
        if name in by_name and by_name[name]["status"] == "Absent" and name not in found:
            found.setdefault(name, []).append(ref_of(item))
            by_name[name]["notes"] = tr("scan_note")

    # 2. shipment date from any page, then compare issue dates in code
    ships = [d for d in (_d(i.get("shipment_date")) for i in items) if d]
    ship = max(ships) if ships else None
    actions = result["action_items"]
    if ship:
        variants = {ship.isoformat(), ship.strftime("%d/%m/%Y"), ship.strftime("%Y/%m/%d"),
                    ship.strftime("%d.%m.%Y"), ship.strftime("%d-%m-%Y")}
        kept = []
        for a in actions:
            if re.search(r"after (the )?shipment date|après la date d.expédition", a, re.I) and not any(v in a for v in variants):
                log.info("Dropped action item using a different shipment date: %s", _c(a))
                continue
            kept.append(a)
        actions[:] = kept
    for item in items:
        name = item.get("document") if item.get("document") in by_name else None
        where = ref_of(item)
        who = f"{dn(name)} ({where})" if name else where
        keyg = (name.split(" (")[0], dn(name).split(" (")[0]) if name else None
        if item.get("retroactive") is True:
            if not _mentioned(actions, ("retro", "rétro"), *([keyg] if keyg else [])):
                actions.append(tr("retro_scan", who=who))
        issue = _d(item.get("issue_date"))
        if issue and ship and issue > ship:
            if not _mentioned(actions, *([keyg] if keyg else []), ("after", "après")):
                actions.append(tr("late", who=who, issue=issue.isoformat(), ship=ship.isoformat()))


# --------------------------------------------------------------------------- routes
@app.post("/api/verify")
async def verify(request: Request, files: List[UploadFile] = File(...), lang: str = Form("en"),
                 _auth: None = Depends(require_auth)):
    lang = lang if lang in TX else "en"
    LANG.set(lang)
    check_rate(client_ip(request))
    if not files:
        raise HTTPException(400, tr("no_files"))
    if len(files) > MAX_FILES:
        raise HTTPException(400, tr("max_files", n=MAX_FILES))
    if AUDIT_SLOTS.locked():
        raise HTTPException(429, tr("busy"))

    loaded = []
    for f in files:
        raw = await f.read()
        if not raw:
            raise HTTPException(400, tr("empty", name=f.filename))
        if len(raw) > MAX_FILE_BYTES:
            raise HTTPException(413, tr("too_big", name=f.filename, mb=MAX_FILE_BYTES // (1024 * 1024)))
        loaded.append((f.filename or "unnamed", raw))

    async with AUDIT_SLOTS:
        content, pages = await run_in_threadpool(build_content, loaded)
        result = normalize(await call_openrouter(content, system_for_lang()))
        found = detect_docs(pages)
        await inspect_scanned(result, pages, found)
        for line in retro_flags(pages):
            if not any(re.search(r"r[eé]tro", a, re.I) for a in result["action_items"]):
                result["action_items"].append(line)
        out = reconcile(result, found)
        out["lang"] = lang
        return out


@app.get("/api/health")
async def health():
    return {"status": "ok", "version": app.version}


@app.get("/api/status")
async def status(_auth: None = Depends(require_auth)):
    return {"status": "ok", "version": app.version, "model": MODEL,
            "key_configured": bool(os.getenv("OPENROUTER_API_KEY")), "env": APP_ENV,
            "auth": bool(APP_PASSWORD)}


@app.get("/")
async def index(_auth: None = Depends(require_auth)):
    if FRONTEND_DIR:
        return FileResponse(FRONTEND_DIR / "index.html")
    return {"service": "DocuVerify AI API", "version": app.version, "docs": "POST /api/verify, GET /api/health"}


@app.get("/config.js")
async def config_js():
    if FRONTEND_DIR and (FRONTEND_DIR / "config.js").is_file():
        return FileResponse(FRONTEND_DIR / "config.js", media_type="text/javascript")
    raise HTTPException(404, "Not found")
