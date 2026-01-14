from __future__ import annotations
"""
app.py — Version Render-ready (Flask + Playwright ASYNC)

Objectifs:
- Scraper Kayak en headless sur Render (Linux)
- Filtrer pubs (eDreams / Anantara / “Annonce” etc.)
- Filtrer définitivement certaines compagnies (Air India, British Airways) + codes parasites (ex: MUC si tu veux)
- Générer des résultats + export PDF
- Compatible Render:
  - Binding sur 0.0.0.0:$PORT (via gunicorn sur Render)
  - Profil Playwright sur /tmp (écriture autorisée)
  - Chromium Playwright (pas channel="chrome" sur Linux)

IMPORTANT:
- Ton template index.html doit poster en POST sur /run
- Sur Render, utilise gunicorn: `gunicorn -b 0.0.0.0:$PORT app:app`
"""

import re
import io
import os
import time
import asyncio
import datetime as dt
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Any, Tuple
from urllib.parse import quote
from html import escape

from flask import Flask, render_template, request, redirect, url_for, send_file

from playwright.async_api import async_playwright, Page as APage

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm


# -----------------------------
# APP
# -----------------------------
app = Flask(__name__)

# -----------------------------
# ENV / RENDER
# -----------------------------
IS_RENDER = bool(os.environ.get("RENDER")) or (os.environ.get("RENDER") == "true") or bool(os.environ.get("PORT"))
PORT = int(os.environ.get("PORT", "5000"))

# Profil Playwright (Windows local / Render Linux)
PW_PROFILE_DIR = os.environ.get("PW_PROFILE_DIR") or (
    r"C:\Users\mrabd\AppData\Local\flight-alert-playwright-profile"
    if os.name == "nt" else "/tmp/flight-alert-playwright-profile"
)

# -----------------------------
# DEFAULT CONFIG
# -----------------------------
DEFAULT_ORIGIN = "CDG"
DEFAULT_DEST = "MAA"
DEFAULT_PAX = 5

DEFAULT_DEPART_START = "2026-07-05"
DEFAULT_DEPART_END   = "2026-07-10"
DEFAULT_RETURN_START = "2026-08-24"
DEFAULT_RETURN_END   = "2026-08-30"

DEFAULT_EXCLUDE_AIRLINES = ["AI", "BA", "LH"]  # codes IATA à exclure via filtre Kayak

# Kayak fs base (sans airlines)
KAYAK_FS_BASE = "layoverdur=-560;stops=-2;cfc=1"

# Scraping behavior
MAX_TABS = 8
MIN_CARDS_PER_PAGE = 15
MAX_SCROLL_ROUNDS = 20
SCROLL_STEP = 1400
WAIT_AFTER_GOTO_MS = 3500
WAIT_AFTER_SCROLL_MS = 1200


# -----------------------------
# GLOBAL STATE
# -----------------------------
STATE: Dict[str, Any] = {
    "pw": None,          # Playwright object
    "context": None,     # BrowserContext
    "initialized": False
}

LAST_RESULTS: Dict[str, Any] = {
    "run_at": None,
    "status": None,
    "offers": [],
    "rejected": [],
    "errors": [],
    "diag": {"events": []},
    "cfg": {},
}


# -----------------------------
# MODEL
# -----------------------------
@dataclass
class Offer:
    site: str
    depart_date: str
    return_date: str
    companies: Optional[str]
    price_per_person_text: Optional[str]
    total_price_text: Optional[str]
    duration_text: Optional[str]
    stops_text: Optional[str]
    duration_min: Optional[int]
    stops: Optional[int]
    url: str
    reason: Optional[str] = None


# -----------------------------
# HELPERS
# -----------------------------
def date_range(start: dt.date, end: dt.date):
    d = start
    while d <= end:
        yield d
        d += dt.timedelta(days=1)

def normalize_spaces(txt: str) -> str:
    if txt is None:
        return ""
    txt = txt.replace("\u202f", " ").replace("\xa0", " ")
    txt = re.sub(r"[ \t]+", " ", txt)
    return txt.strip()

def parse_price_eur(txt: Optional[str]) -> Optional[int]:
    if not txt:
        return None
    t = normalize_spaces(txt)
    m = re.search(r"(\d[\d ]*)\s*€", t)
    if not m:
        return None
    raw = m.group(1).replace(" ", "")
    try:
        return int(raw)
    except Exception:
        return None


# -----------------------------
# AD / PUB FILTER
# -----------------------------
AD_MARKERS = [
    "annonce", "sponsorisé", "sponsored", "publicité",
    "edreams",
    "trouvez les meilleures offres sur edreams",
    "comparez plus de 600 compagnies aériennes",
    "anantara", "anantara hotels", "anantara hotels & resorts",
    "seasonal splendour",
    "hotel", "hotels", "hôtel", "hôtels", "resort", "resorts",
]

AD_REGEX = re.compile(r"(?i)\b(annonce|sponsorisé|sponsored|publicité)\b")

def is_ad_block(text: str) -> bool:
    t = normalize_spaces(text)
    if not t:
        return True
    tl = t.lower()

    if AD_REGEX.search(t):
        return True

    for m in AD_MARKERS:
        if m in tl:
            return True

    return False


# -----------------------------
# KAYAK PARSING (TEXT-BASED)
# -----------------------------
RE_PRICE_PER_PERSON = re.compile(r"(\d[\d \u202f\xa0]*)\s*€\s*/\s*personne", re.IGNORECASE)
RE_TOTAL_PRICE      = re.compile(r"(\d[\d \u202f\xa0]*)\s*€\s*au\s*total", re.IGNORECASE)
RE_STOPS            = re.compile(r"\b(\d+)\s*escale[s]?\b", re.IGNORECASE)
RE_DURATION         = re.compile(r"\b(\d+)\s*h(?:\s*(\d+)\s*min)?\b", re.IGNORECASE)

def extract_prices(text: str) -> Tuple[Optional[str], Optional[str]]:
    t = text or ""
    m1 = RE_PRICE_PER_PERSON.search(t)
    m2 = RE_TOTAL_PRICE.search(t)

    ppp = None
    tot = None
    if m1:
        ppp = normalize_spaces(m1.group(1)) + " €"
    if m2:
        tot = normalize_spaces(m2.group(1)) + " €"
    return ppp, tot

def extract_companies(text: str) -> Optional[str]:
    lines = [normalize_spaces(x) for x in (text or "").split("\n")]
    lines = [x for x in lines if x]

    bad_contains = [
        "enregistrer", "partager", "le meilleur choix", "le moins cher",
        "économique", "non remboursable"
    ]

    # 1) Ligne avec virgule (souvent “Compagnie1, Compagnie2”)
    for ln in lines:
        lnl = ln.lower()
        if any(b in lnl for b in bad_contains):
            continue
        if "€" in ln:
            continue
        if re.search(r"\b\d{1,2}:\d{2}\b", ln):
            continue
        if "," in ln and re.search(r"[A-Za-zÀ-ÿ]", ln):
            if is_ad_block(ln):
                continue
            return ln[:140]

    # 2) Fallback “air …”
    for ln in lines:
        lnl = ln.lower()
        if "€" in ln:
            continue
        if re.search(r"\b\d{1,2}:\d{2}\b", ln):
            continue
        if "air" in lnl and not is_ad_block(ln):
            return ln[:140]

    return None

def extract_stops_and_duration(text: str) -> Tuple[Optional[str], Optional[str], Optional[int], Optional[int]]:
    t = normalize_spaces(text)

    stops_text = None
    stops = None
    m = RE_STOPS.search(t)
    if m:
        stops = int(m.group(1))
        stops_text = f"{stops} escale" if stops == 1 else f"{stops} escales"
    else:
        if re.search(r"\bdirect\b|\bnonstop\b|\bsans escale\b", t, re.IGNORECASE):
            stops = 0
            stops_text = "Direct"

    durations = []
    for mh in RE_DURATION.finditer(t):
        h = int(mh.group(1))
        mn = int(mh.group(2)) if mh.group(2) else 0
        durations.append((h, mn))

    duration_text = None
    duration_min = None
    if durations:
        mins = [h * 60 + mn for (h, mn) in durations]
        duration_min = max(mins)  # conservateur
        max_idx = mins.index(duration_min)
        h, mn = durations[max_idx]
        duration_text = f"{h}h {mn}min" if mn else f"{h}h"

    return stops_text, duration_text, stops, duration_min


# -----------------------------
# KAYAK URL BUILD
# -----------------------------
def build_kayak_fs(base_fs: str, exclude_airlines: List[str]) -> str:
    ex = [a.strip().upper() for a in (exclude_airlines or []) if a.strip()]
    # garde tout (pas seulement AI/BA/LH) si tu veux:
    # ex = [a for a in ex if re.fullmatch(r"[A-Z0-9]{2,3}", a)]
    if ex:
        return f"airlines=-{','.join(ex)},flylocal;{base_fs}"
    return base_fs

def build_kayak_url(origin: str, dest: str, d1: str, d2: str, pax: int, base_fs: str, exclude_airlines: List[str]) -> str:
    fs = build_kayak_fs(base_fs, exclude_airlines)
    # IMPORTANT: on encode fs, en gardant quelques séparateurs lisibles
    fs_encoded = quote(fs, safe="-;,=")
    base = f"https://www.kayak.fr/flights/{origin}-{dest}/{d1}/{d2}/{pax}adults"
    return f"{base}?sort=price_a&fs={fs_encoded}"


# -----------------------------
# VALIDATION / FILTRES “NE JAMAIS AFFICHER”
# -----------------------------
BANNED_COMPANY_SUBSTR = [
    "air india",
    "british airways",
]
BANNED_COMPANY_CODES = [
    # si tu veux bannir des segments/codes visibles dans la zone “compagnies”
    # "MUC",
]

def validate_offer(o: Offer, max_stops: int, max_duration_min: int) -> Tuple[bool, str]:
    # Exclusion forte: compagnies interdites (même si Kayak “filtre” mal)
    if o.companies:
        comp_low = o.companies.lower()
        if any(b in comp_low for b in BANNED_COMPANY_SUBSTR):
            return False, "Compagnie exclue (Air India / British Airways)"
        if any(code.lower() in comp_low for code in BANNED_COMPANY_CODES):
            return False, "Code/segment exclu"

    if o.reason:
        return False, o.reason
    if not o.companies:
        return False, "companies=None"
    if parse_price_eur(o.price_per_person_text) is None:
        return False, "price_per_person=None"
    if parse_price_eur(o.total_price_text) is None:
        return False, "total_price=None"
    if o.stops is None:
        return False, "stops=None"
    if o.duration_min is None:
        return False, "duration_min=None"
    if o.stops > max_stops:
        return False, f"stops={o.stops} > max_stops={max_stops}"
    if o.duration_min > max_duration_min:
        return False, f"duration_min={o.duration_min} > max_duration_min={max_duration_min}"
    return True, "OK"


# -----------------------------
# PLAYWRIGHT (ASYNC) — Render-ready
# -----------------------------
async def ensure_browser_async(diag: Dict[str, Any]) -> None:
    if STATE.get("initialized") and STATE.get("context"):
        return

    os.makedirs(PW_PROFILE_DIR, exist_ok=True)

    diag["events"].append({
        "level": "INFO", "site": "kayak", "d1": "-", "d2": "-",
        "msg": f"Launching Playwright persistent context (profile={PW_PROFILE_DIR}) headless={True if IS_RENDER else False}"
    })

    pw = await async_playwright().start()

    # Render/Linux: PAS de channel="chrome"
    context = await pw.chromium.launch_persistent_context(
        user_data_dir=PW_PROFILE_DIR,
        headless=True if IS_RENDER else True,  # tu peux mettre False en local si tu veux voir
        viewport={"width": 1400, "height": 900},
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--no-first-run",
            "--no-default-browser-check",
        ],
    )

    STATE["pw"] = pw
    STATE["context"] = context
    STATE["initialized"] = True


async def close_browser_async() -> None:
    try:
        if STATE.get("context"):
            await STATE["context"].close()
    except Exception:
        pass
    try:
        if STATE.get("pw"):
            await STATE["pw"].stop()
    except Exception:
        pass
    STATE.update({"pw": None, "context": None, "initialized": False})


# -----------------------------
# WAIT READY / CARDS
# -----------------------------
async def detect_antibot(page: APage) -> bool:
    try:
        body = (await page.locator("body").inner_text(timeout=3000)).lower()
        return any(x in body for x in ["robot", "captcha", "unusual traffic", "verify you are a human", "inhabituel"])
    except Exception:
        return False

async def _wait_kayak_results_ready(page: APage, timeout_ms: int = 30_000) -> None:
    # stabilise un peu la page
    try:
        await page.wait_for_load_state("domcontentloaded", timeout=timeout_ms)
    except Exception:
        pass
    try:
        await page.wait_for_load_state("networkidle", timeout=timeout_ms)
    except Exception:
        pass

    start = time.time()
    last_counts: List[int] = []
    while (time.time() - start) * 1000 < timeout_ms:
        try:
            loc = page.locator("div[role='listitem'], [data-testid*='result'], div[class*='result']")
            c = await loc.count()
        except Exception:
            c = 0

        last_counts.append(c)
        if len(last_counts) > 6:
            last_counts.pop(0)

        # stable et suffisamment de résultats
        if len(last_counts) >= 3 and last_counts[-1] >= 8 and last_counts[-1] == last_counts[-2] == last_counts[-3]:
            return

        await page.wait_for_timeout(600)

async def get_candidate_cards(page: APage) -> List[str]:
    loc = page.locator("div[role='listitem'], [data-testid*='result'], div[class*='result']")
    try:
        count = await loc.count()
    except Exception:
        return []

    texts: List[str] = []
    for i in range(min(count, 120)):
        try:
            txt = (await loc.nth(i).inner_text(timeout=1500)).strip()
            if txt:
                texts.append(txt)
        except Exception:
            continue
    return texts


async def extract_min_cards(page: APage, url: str, d1: str, d2: str, diag: Dict[str, Any]) -> List[Offer]:
    if await detect_antibot(page):
        diag["events"].append({"level": "WARN", "site": "kayak", "d1": d1, "d2": d2, "msg": "Anti-bot detected"})
        return [Offer(
            site="kayak", depart_date=d1, return_date=d2,
            companies=None, price_per_person_text=None, total_price_text=None,
            duration_text=None, stops_text=None,
            duration_min=None, stops=None,
            url=url, reason="BLOQUÉ (anti-bot/captcha)"
        )]

    offers: List[Offer] = []

    for round_idx in range(MAX_SCROLL_ROUNDS):
        raw_texts = await get_candidate_cards(page)
        diag["events"].append({
            "level": "INFO", "site": "kayak", "d1": d1, "d2": d2,
            "msg": f"round={round_idx+1}/{MAX_SCROLL_ROUNDS} candidates={len(raw_texts)}"
        })

        parsed: List[Offer] = []
        for txt in raw_texts:
            if is_ad_block(txt):
                continue

            ppp, tot = extract_prices(txt)
            if parse_price_eur(ppp) is None or parse_price_eur(tot) is None:
                continue

            companies = extract_companies(txt)
            if not companies:
                continue

            stops_text, duration_text, stops, duration_min = extract_stops_and_duration(txt)

            parsed.append(Offer(
                site="kayak",
                depart_date=d1,
                return_date=d2,
                companies=companies,
                price_per_person_text=ppp,
                total_price_text=tot,
                duration_text=duration_text,
                stops_text=stops_text,
                duration_min=duration_min,
                stops=stops,
                url=url
            ))

        # de-dup
        seen = set()
        for o in parsed:
            key = (o.companies, o.price_per_person_text, o.total_price_text, o.duration_text, o.stops_text)
            if key in seen:
                continue
            seen.add(key)
            offers.append(o)

        if len(offers) >= MIN_CARDS_PER_PAGE:
            break

        try:
            await page.mouse.wheel(0, SCROLL_STEP)
        except Exception:
            pass
        await page.wait_for_timeout(WAIT_AFTER_SCROLL_MS)

    if not offers:
        return [Offer(
            site="kayak", depart_date=d1, return_date=d2,
            companies=None, price_per_person_text=None, total_price_text=None,
            duration_text=None, stops_text=None,
            duration_min=None, stops=None,
            url=url, reason="Aucune card vol valide trouvée"
        )]

    return offers[:MIN_CARDS_PER_PAGE]


# -----------------------------
# RUN CORE
# -----------------------------
async def _run_one_pair_on_page(page: APage, cfg: Dict[str, Any], d1s: str, d2s: str, diag: Dict[str, Any]):
    origin = cfg["origin"]
    dest = cfg["dest"]
    pax = int(cfg["pax"])
    max_stops = int(cfg["max_stops"])
    max_duration_min = int(cfg["max_duration_h"]) * 60
    exclude_airlines = cfg.get("exclude_airlines", [])

    url = build_kayak_url(
        origin=origin,
        dest=dest,
        d1=d1s,
        d2=d2s,
        pax=pax,
        base_fs=KAYAK_FS_BASE,
        exclude_airlines=exclude_airlines
    )

    diag["events"].append({"level": "INFO", "site": "kayak", "d1": d1s, "d2": d2s, "msg": f"GOTO url={url}"})

    await page.goto(url, wait_until="domcontentloaded", timeout=120_000)
    await _wait_kayak_results_ready(page, timeout_ms=30_000)
    await page.wait_for_timeout(WAIT_AFTER_GOTO_MS)

    raw_offers = await extract_min_cards(page, url, d1s, d2s, diag)

    valid_out: List[Offer] = []
    rejected_out: List[Offer] = []

    for o in raw_offers:
        ok, reason = validate_offer(o, max_stops, max_duration_min)
        if ok:
            valid_out.append(o)
        else:
            o.reason = reason
            rejected_out.append(o)

    return valid_out, rejected_out


async def run_kayak_pairs_async(cfg: Dict[str, Any]) -> Dict[str, Any]:
    diag = {"events": []}
    errors: List[str] = []
    valid_out: List[Offer] = []
    rejected_out: List[Offer] = []

    await ensure_browser_async(diag)
    context = STATE["context"]

    depart_start = dt.date.fromisoformat(cfg["depart_start"])
    depart_end = dt.date.fromisoformat(cfg["depart_end"])
    return_start = dt.date.fromisoformat(cfg["return_start"])
    return_end = dt.date.fromisoformat(cfg["return_end"])

    jobs = [(d1.isoformat(), d2.isoformat())
            for d1 in date_range(depart_start, depart_end)
            for d2 in date_range(return_start, return_end)]

    total_pairs = len(jobs)

    q: asyncio.Queue = asyncio.Queue()
    for j in jobs:
        await q.put(j)

    async def worker(worker_id: int):
        page = await context.new_page()
        try:
            while True:
                try:
                    d1s, d2s = q.get_nowait()
                except asyncio.QueueEmpty:
                    break

                try:
                    v, r = await _run_one_pair_on_page(page, cfg, d1s, d2s, diag)
                    valid_out.extend(v)
                    rejected_out.extend(r)
                except Exception as e:
                    msg = f"kayak {d1s}-{d2s}: {repr(e)}"
                    errors.append(msg)
                    diag["events"].append({"level": "ERROR", "site": "kayak", "d1": d1s, "d2": d2s, "msg": msg})
                finally:
                    q.task_done()
        finally:
            await page.close()

    n = min(MAX_TABS, total_pairs) if total_pairs > 0 else 1
    tasks = [asyncio.create_task(worker(i)) for i in range(n)]
    await asyncio.gather(*tasks)

    diag["events"].append({
        "level": "INFO", "site": "-", "d1": "-", "d2": "-",
        "msg": f"SUMMARY pairs_tested={total_pairs} valid_total={len(valid_out)} rejected_total={len(rejected_out)} errors_total={len(errors)}"
    })

    return {
        "run_at": dt.datetime.now().isoformat(timespec="seconds"),
        "status": "DONE",
        "offers": [asdict(o) for o in valid_out],
        "rejected": [asdict(o) for o in rejected_out],
        "errors": errors,
        "diag": diag,
        "cfg": cfg,
        "profile_dir": PW_PROFILE_DIR,
    }


def run_kayak_pairs(cfg: Dict[str, Any]) -> Dict[str, Any]:
    return asyncio.run(run_kayak_pairs_async(cfg))


# -----------------------------
# PDF EXPORT
# -----------------------------
def sort_offers_by_price_dicts(offers: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    def key(o: Dict[str, Any]) -> int:
        p = parse_price_eur(o.get("price_per_person_text") or o.get("price_text"))
        return p if p is not None else 10**9
    return sorted(offers, key=key)

def _trunc(s: str, n: int) -> str:
    s = (s or "").strip()
    return s if len(s) <= n else s[:n-1] + "…"

def _mk_pdf_bytes_from_dicts(offers: List[Dict[str, Any]]) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    left = 14 * mm
    right = 14 * mm
    top = 16 * mm
    bottom = 14 * mm
    table_w = w - left - right

    title = "Vols Kayak (triés par prix / personne)"
    c.setFont("Helvetica-Bold", 14)
    c.drawString(left, h - top, title)

    c.setFont("Helvetica", 9)
    c.drawString(left, h - top - 6 * mm, f"Export généré le {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}")

    y = h - top - 14 * mm
    header_h = 8 * mm
    row_h = 7 * mm

    col_w_mm = [20, 20, 58, 16, 16, 18, 16, 14]
    col_w = [x * mm for x in col_w_mm]

    total_cols = sum(col_w)
    if abs(total_cols - table_w) > 1:
        scale = table_w / total_cols
        col_w = [cw * scale for cw in col_w]

    headers = ["Départ", "Retour", "Compagnies", "€/pers", "Total", "Durée", "Escales", "Lien"]

    def _draw_header(y_top: float):
        c.setFillColorRGB(0.94, 0.94, 0.94)
        c.rect(left, y_top - header_h, table_w, header_h, fill=1, stroke=0)

        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 9)

        x = left
        for i, hd in enumerate(headers):
            c.drawString(x + 2.2 * mm, y_top - header_h + 2.6 * mm, hd)
            x += col_w[i]

        c.setLineWidth(0.4)
        c.setStrokeColorRGB(0.70, 0.70, 0.70)
        c.rect(left, y_top - header_h, table_w, header_h, fill=0, stroke=1)

        x = left
        for i in range(len(col_w) - 1):
            x += col_w[i]
            c.line(x, y_top - header_h, x, y_top)

        c.setStrokeColorRGB(0.65, 0.65, 0.65)
        c.line(left, y_top - header_h, left + table_w, y_top - header_h)

    def _new_page():
        c.showPage()
        c.setFont("Helvetica-Bold", 14)
        c.drawString(left, h - top, title)
        c.setFont("Helvetica", 9)
        c.drawString(left, h - top - 6 * mm, f"Export généré le {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}")
        return h - top - 14 * mm

    def _cell_text(s: str, max_chars: int) -> str:
        return _trunc((s or "").strip(), max_chars)

    _draw_header(y)
    y -= header_h

    c.setFont("Helvetica", 9)
    c.setLineWidth(0.3)

    for idx, o in enumerate(offers):
        if y - row_h < bottom:
            y = _new_page()
            _draw_header(y)
            y -= header_h

        if idx % 2 == 1:
            c.setFillColorRGB(0.98, 0.98, 0.98)
            c.rect(left, y - row_h, table_w, row_h, fill=1, stroke=0)

        c.setStrokeColorRGB(0.85, 0.85, 0.85)
        c.rect(left, y - row_h, table_w, row_h, fill=0, stroke=1)

        x = left
        for i in range(len(col_w) - 1):
            x += col_w[i]
            c.line(x, y - row_h, x, y)

        depart = _cell_text(o.get("depart_date", ""), 10)
        ret    = _cell_text(o.get("return_date", ""), 10)
        comp   = _cell_text(o.get("companies", ""), 45)
        ppp    = _cell_text(o.get("price_per_person_text", ""), 12)
        tot    = _cell_text(o.get("total_price_text", ""), 12)
        dur    = _cell_text(o.get("duration_text", ""), 12)
        stp    = _cell_text(o.get("stops_text", ""), 12)
        url    = (o.get("url", "") or "").strip()

        text_y = y - row_h + 2.4 * mm
        x = left

        c.setFillColorRGB(0, 0, 0)
        c.drawString(x + 2.2 * mm, text_y, depart); x += col_w[0]
        c.drawString(x + 2.2 * mm, text_y, ret);    x += col_w[1]
        c.drawString(x + 2.2 * mm, text_y, comp);   x += col_w[2]

        c.drawRightString(x + col_w[3] - 2.2 * mm, text_y, ppp); x += col_w[3]
        c.drawRightString(x + col_w[4] - 2.2 * mm, text_y, tot); x += col_w[4]

        c.drawString(x + 2.2 * mm, text_y, dur); x += col_w[5]
        c.drawString(x + 2.2 * mm, text_y, stp); x += col_w[6]

        link_label = "Ouvrir" if url else "—"
        if url:
            c.setFillColorRGB(0.10, 0.35, 0.75)
        else:
            c.setFillColorRGB(0.45, 0.45, 0.45)

        label_w = c.stringWidth(link_label, "Helvetica", 9)
        label_x = x + (col_w[7] - label_w) / 2
        c.drawString(label_x, text_y, link_label)

        if url:
            c.linkURL(url, (label_x, text_y - 1, label_x + label_w, text_y + 9), relative=0)

        c.setFillColorRGB(0, 0, 0)
        y -= row_h

    c.save()
    buf.seek(0)
    return buf.getvalue()


# -----------------------------
# ROUTES
# -----------------------------
@app.route("/", methods=["GET"])
def index():
    cfg = {
        "origin": DEFAULT_ORIGIN,
        "dest": DEFAULT_DEST,
        "pax": DEFAULT_PAX,
        "depart_start": DEFAULT_DEPART_START,
        "depart_end": DEFAULT_DEPART_END,
        "return_start": DEFAULT_RETURN_START,
        "return_end": DEFAULT_RETURN_END,
        "max_stops": 1,
        "max_duration_h": 22,
        "exclude_airlines": DEFAULT_EXCLUDE_AIRLINES,
        "profile_dir": PW_PROFILE_DIR,
    }
    return render_template("index.html", cfg=cfg, last=LAST_RESULTS)

@app.route("/run", methods=["POST"])
def run():
    global LAST_RESULTS

    exclude_airlines = request.form.getlist("exclude_airlines")

    cfg = {
        "origin": request.form.get("origin", DEFAULT_ORIGIN),
        "dest": request.form.get("dest", DEFAULT_DEST),
        "pax": request.form.get("pax", DEFAULT_PAX),
        "depart_start": request.form.get("depart_start", DEFAULT_DEPART_START),
        "depart_end": request.form.get("depart_end", DEFAULT_DEPART_END),
        "return_start": request.form.get("return_start", DEFAULT_RETURN_START),
        "return_end": request.form.get("return_end", DEFAULT_RETURN_END),
        "max_stops": request.form.get("max_stops", 1),
        "max_duration_h": request.form.get("max_duration_h", 22),
        "exclude_airlines": exclude_airlines,
        "profile_dir": PW_PROFILE_DIR,
    }

    LAST_RESULTS = {
        "run_at": dt.datetime.now().isoformat(timespec="seconds"),
        "status": "RUNNING",
        "offers": [],
        "rejected": [],
        "errors": [],
        "diag": {"events": []},
        "cfg": cfg,
    }

    LAST_RESULTS = run_kayak_pairs(cfg)
    return redirect(url_for("index"))

@app.route("/close", methods=["POST"])
def close():
    global LAST_RESULTS
    asyncio.run(close_browser_async())
    LAST_RESULTS["status"] = "CLOSED"
    LAST_RESULTS.setdefault("diag", {"events": []})
    LAST_RESULTS["diag"].setdefault("events", [])
    LAST_RESULTS["diag"]["events"].append({
        "level": "INFO", "site": "kayak", "d1": "-", "d2": "-",
        "msg": "Browser closed by user."
    })
    return redirect(url_for("index"))

@app.route("/export_pdf", methods=["POST"])
def export_pdf():
    # le front doit envoyer JSON: { offers: [...] }
    data = request.get_json(force=True, silent=False)
    offers = data.get("offers", [])
    offers_sorted = sort_offers_by_price_dicts(offers)
    pdf_bytes = _mk_pdf_bytes_from_dicts(offers_sorted)

    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name="vols_kayak_tri_prix.pdf"
    )


# -----------------------------
# LOCAL RUN (Render utilise gunicorn, tu ne supprimes pas ça)
# -----------------------------
if __name__ == "__main__":
    # En local seulement. Sur Render: gunicorn -b 0.0.0.0:$PORT app:app
    app.run(host="0.0.0.0", port=PORT, debug=not IS_RENDER)
