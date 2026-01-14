# app.py — Version COMPLETE (corrigée)
# - Scrape Kayak (texte cards) + anti pub
# - Paramètre Front: exclure compagnies (AI/BA/LH) => Kayak fs=airlines=-AI,BA,LH,flylocal;...
# - Export PDF (tableau) AVEC liens cliquables
# - Envoi Outlook automatique HTML (tableau propre + liens)
#   Subject: inclut la date du jour
#   Top 8: pas deux offres avec même compagnie+jour de départ
#
# Pré-requis:
#   python -m pip install --upgrade pip
#   pip install flask playwright reportlab pywin32
#   playwright install

import re
import io
import os
import time
import tempfile
import datetime as dt
from dataclasses import dataclass, asdict
from typing import Optional, List, Dict, Any, Tuple
from html import escape
from urllib.parse import quote

import pythoncom
import win32com.client  # pywin32 (Windows + Outlook)


import asyncio
from playwright.async_api import async_playwright, Page as APage

MAX_TABS = 8  # <= ton "MAX ONGLET"
HEADLESS = False  # tu peux passer True si tu veux accélérer


from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file
from playwright.sync_api import sync_playwright, Page

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

app = Flask(__name__)

# -----------------------------
# CONFIG
# -----------------------------
DEFAULT_ORIGIN = "CDG"
DEFAULT_DEST = "MAA"
DEFAULT_PAX = 5

DEFAULT_DEPART_START = "2026-07-05"
DEFAULT_DEPART_END   = "2026-07-10"
DEFAULT_RETURN_START = "2026-08-24"
DEFAULT_RETURN_END   = "2026-08-30"

# Exclusions compagnies (Front)
DEFAULT_EXCLUDE_AIRLINES = ["AI", "BA", "LH"]  # [] si vous voulez aucune exclusion par défaut

# Outlook
OUTLOOK_SENDER_SMTP = "aaa@aa"
OUTLOOK_TO = ["aa@aa"]

# Kayak base filters (sans airlines) , "605001@gmail.com"
KAYAK_FS_BASE = "layoverdur=-560;stops=-2;cfc=1"

PW_PROFILE_DIR = r"C:\Users\mrabd\AppData\Local\flight-alert-playwright-profile"

# Scraping behavior
MIN_CARDS_PER_PAGE = 15
MAX_SCROLL_ROUNDS = 20
SCROLL_STEP = 1400
WAIT_AFTER_GOTO_MS = 3500
WAIT_AFTER_SCROLL_MS = 1200

# -----------------------------
# GLOBAL STATE
# -----------------------------
STATE: Dict[str, Any] = {
    "playwright": None,
    "context": None,
    "page": None,
    "initialized": False,
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
# HELPERS
# -----------------------------

async def _wait_kayak_results_ready(page: APage, timeout_ms: int = 30_000):
    try:
        await page.wait_for_load_state("domcontentloaded", timeout=timeout_ms)
    except Exception:
        pass
    try:
        await page.wait_for_load_state("networkidle", timeout=timeout_ms)
    except Exception:
        pass

    start = time.time()
    last_counts = []
    while (time.time() - start) * 1000 < timeout_ms:
        try:
            loc = page.locator("div[role='listitem'], [data-testid*='result'], div[class*='result']")
            c = await loc.count()
        except Exception:
            c = 0

        last_counts.append(c)
        if len(last_counts) > 6:
            last_counts.pop(0)

        if len(last_counts) >= 3 and last_counts[-1] >= 8 and last_counts[-1] == last_counts[-2] == last_counts[-3]:
            return

        await page.wait_for_timeout(600)


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

async def detect_antibot(page: APage) -> bool:
    try:
        body = (await page.locator("body").inner_text(timeout=3000)).lower()
        return any(x in body for x in ["robot", "captcha", "unusual traffic", "verify you are a human", "inhabituel"])
    except Exception:
        return False


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
# KAYAK URL BUILD (CORRECT)
# -----------------------------
def build_kayak_fs(base_fs: str, exclude_airlines: Optional[List[str]] = None) -> str:
    """
    Construit la valeur fs EXACTEMENT au format Kayak.

    Exemple attendu si exclusions = ["AI","BA","LH"]:
      airlines=-AI,BA,LH,flylocal;layoverdur=-560;stops=-2;cfc=1

    Si aucune exclusion:
      layoverdur=-560;stops=-2;cfc=1
    """
    exclude_airlines = exclude_airlines or []
    ex = [a.strip().upper() for a in exclude_airlines if a and a.strip()]
    # whitelist (si vous voulez permettre plus tard d'autres codes, supprimez ce filtre)
    ex = [a for a in ex if a in ("AI", "BA", "LH")]

    if ex:
        airlines_part = "airlines=-" + ",".join(ex) + ",flylocal"
        return f"{airlines_part};{base_fs}"
    return base_fs

def build_kayak_url(origin: str, dest: str, d1: str, d2: str, pax: int,
                   base_fs: str, exclude_airlines: Optional[List[str]] = None) -> str:
    """
    Construit l’URL Kayak avec fs encodé comme l’exemple correct:
    airlines%3D-AI%2CBA%2CLH%2Cflylocal%3Blayoverdur%3D-560%3Bstops%3D-2%3Bcfc%3D1
    """
    fs = build_kayak_fs(base_fs, exclude_airlines)

    # On encode TOUT, sauf le '-' (sinon airlines=-... est cassé)
    fs_encoded = quote(fs, safe='-')

    base = f"https://www.kayak.fr/flights/{origin}-{dest}/{d1}/{d2}/{pax}adults"
    return f"{base}?sort=price_a&fs={fs_encoded}"

# -----------------------------
# OFFER MODEL
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

def validate_offer(o: Offer, max_stops: int, max_duration_min: int) -> Tuple[bool, str]:

    # 🔴 BLOCAGE compagnies interdites
    banned = ["air india", "british airways","MUC",", MUC",",MUC"," ,MUC",", MUC ",", MUC "]
    if o.companies:
        comp = o.companies.lower()
        if any(b in comp for b in banned):
            return False, "Compagnie exclue (Air India / British)"

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
# PLAYWRIGHT: SINGLE CHROME
# -----------------------------
async def ensure_browser_async(diag: Dict[str, Any]):
    if STATE.get("initialized") and STATE.get("context"):
        return

    diag["events"].append({
        "level": "INFO", "site": "kayak", "d1": "-", "d2": "-",
        "msg": "Launching Chrome (persistent profile) [ASYNC]..."
    })

    pw = await async_playwright().start()
    context = pw.chromium.launch(
    headless=True
)



    STATE["playwright"] = pw
    STATE["context"] = context
    STATE["initialized"] = True


def ensure_browser(diag: Dict[str, Any]):
    # wrapper sync -> async (on garde le même nom public si tu veux)
    asyncio.run(ensure_browser_async(diag))


async def close_browser_async():
    try:
        if STATE.get("context"):
            await STATE["context"].close()
    except Exception:
        pass
    try:
        if STATE.get("playwright"):
            await STATE["playwright"].stop()
    except Exception:
        pass
    STATE.update({
        "playwright": None,
        "context": None,
        "page": None,
        "initialized": False
    })


def close_browser():
    # wrapper sync -> async (même nom)
    asyncio.run(close_browser_async())


# -----------------------------
# KAYAK CARD PARSING (TEXT-BASED)
# -----------------------------
RE_PRICE_PER_PERSON = re.compile(r"(\d[\d \u202f\xa0]*)\s*€\s*/\s*personne", re.IGNORECASE)
RE_TOTAL_PRICE      = re.compile(r"(\d[\d \u202f\xa0]*)\s*€\s*au\s*total", re.IGNORECASE)
RE_STOPS            = re.compile(r"\b(\d+)\s*escale[s]?\b", re.IGNORECASE)
RE_DURATION         = re.compile(r"\b(\d+)\s*h(?:\s*(\d+)\s*min)?\b", re.IGNORECASE)

import re

# Marqueurs pubs / agrégateurs / hôtels (lowercase)
AD_MARKERS = [
    # pubs génériques
    "annonce", "sponsorisé", "sponsored", "publicité", "ad ",
    "voir l'offre", "voir loffre", "offres exclusives", "offers exclusives",
    "prix avantageux", "service haut de gamme", "seasonal splendour",

    # agrégateurs / OTA (exemples)
    "edreams",
    "trouvez les meilleures offres sur edreams",
    "comparez plus de 600 compagnies aériennes",

    # hôtels / resorts (exemples)
    "anantara", "anantara hotels", "anantara hotels & resorts",
    "resort", "resorts", "hotel", "hotels", "hôtel", "hôtels",
]

# Regex “Annonce” plus robuste (et séparateurs fréquents)
AD_REGEX = re.compile(
    r"(?i)\b(annonce|sponsorisé|sponsored|publicité)\b"
)

def is_ad_block(text: str) -> bool:
    t = normalize_spaces(text)
    if not t:
        return True

    tl = t.lower()

    # 1) Très discriminant: présence explicite "Annonce" / "Sponsored" / etc.
    if AD_REGEX.search(t):
        return True

    # 2) Marqueurs connus
    for m in AD_MARKERS:
        if m in tl:
            return True

    # 3) Cas “cards non-vol” typiques: beaucoup de marketing et pas de structure vol
    #    (garde-fou: si ça contient "edreams" ou "anantara" on a déjà filtré, ici c'est bonus)
    marketing_hits = 0
    for kw in ["offres", "exclusives", "comparez", "meilleures", "splendour"]:
        if kw in tl:
            marketing_hits += 1
    if marketing_hits >= 2 and ("€" in tl) and ("escale" in tl) and ("vol" in tl):
        # Si vous voulez être encore plus strict, laissez True.
        # Ici on le met en pub si la card ressemble à un encart marketing.
        return True

    return False


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

    bad_contains = ["enregistrer", "partager", "le meilleur choix", "le moins cher", "économique", "non remboursable"]
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
# RUN
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

    valid_out = []
    rejected_out = []

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

    # 1) Build jobs
    depart_start = dt.date.fromisoformat(cfg["depart_start"])
    depart_end = dt.date.fromisoformat(cfg["depart_end"])
    return_start = dt.date.fromisoformat(cfg["return_start"])
    return_end = dt.date.fromisoformat(cfg["return_end"])

    jobs = [(d1.isoformat(), d2.isoformat())
            for d1 in date_range(depart_start, depart_end)
            for d2 in date_range(return_start, return_end)]

    total_pairs = len(jobs)

    # 2) Queue
    q: asyncio.Queue = asyncio.Queue()
    for j in jobs:
        await q.put(j)

    # 3) Worker (1 onglet = 1 page)
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

    # 4) Run N tabs
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
    # même nom qu’avant, Flask continue d’appeler run_kayak_pairs()
    return asyncio.run(run_kayak_pairs_async(cfg))


# -----------------------------
# SORT / PDF / OUTLOOK
# -----------------------------
def sort_offers_by_price_dicts(offers: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    def key(o: Dict[str, Any]) -> int:
        p = parse_price_eur(o.get("price_per_person_text") or o.get("price_text"))
        return p if p is not None else 10**9
    return sorted(offers, key=key)

def _trunc(s: str, n: int) -> str:
    s = (s or "").strip()
    return s if len(s) <= n else s[:n-1] + "…"

# Top 8: pas 2 offres avec même "companies" + même depart_date
def _offer_key_same_company_same_day(o: Dict[str, Any]) -> Tuple[str, str]:
    comp = (o.get("companies") or "").strip().lower()
    dpt = (o.get("depart_date") or "").strip()
    return (comp, dpt)

def _pick_top_unique_company_same_day(offers_sorted: List[Dict[str, Any]], limit: int = 15) -> List[Dict[str, Any]]:
    picked: List[Dict[str, Any]] = []
    seen = set()
    for o in offers_sorted:
        k = _offer_key_same_company_same_day(o)
        if k in seen:
            continue
        seen.add(k)
        picked.append(o)
        if len(picked) >= limit:
            break
    return picked

def _mk_outlook_body_top8_html(offers_sorted: List[Dict[str, Any]]) -> str:
    top = _pick_top_unique_company_same_day(offers_sorted, limit=15)

    def e(s: Optional[str]) -> str:
        return escape((s or "").strip())

    # ✅ Cards only (OK desktop + mobile). No table, no @media needed.
    cards_html = []
    for o in top:
        depart = e(o.get("depart_date"))
        ret    = e(o.get("return_date"))
        comp   = e(o.get("companies"))
        ppp    = e(o.get("price_per_person_text"))
        tot    = e(o.get("total_price_text"))
        dur    = e(o.get("duration_text"))
        stp    = e(o.get("stops_text"))
        url    = (o.get("url") or "").strip()

        link_html = (
            f'<a href="{escape(url)}" target="_blank" '
            f'style="display:inline-block;'
            f'padding:10px 14px;'
            f'border:1px solid #dbe3f1;'
            f'border-radius:10px;'
            f'background:#ffffff;'
            f'color:#1a73e8;'
            f'font-family:Segoe UI, Arial;'
            f'font-size:12px;'
            f'font-weight:700;'
            f'text-decoration:none;'
            f'">Ouvrir</a>'
            if url else
            '<span style="color:#9aa0a6;">—</span>'
        )

        cards_html.append(f"""
          <div style="border:1px solid #e7ebf3;border-radius:14px;padding:14px;margin:0 0 12px 0;background:#ffffff;">
            <div style="font-family:Segoe UI, Arial;font-size:13px;font-weight:800;color:#111111;line-height:1.35;margin-bottom:8px;">
              {comp}
            </div>

            <div style="font-family:Segoe UI, Arial;font-size:12px;color:#374151;line-height:1.55;">
              <div style="margin-bottom:4px;">
                <span style="color:#6b7280;">Départ :</span> <b style="color:#111111;">{depart}</b>
              </div>
              <div style="margin-bottom:8px;">
                <span style="color:#6b7280;">Retour :</span> <b style="color:#111111;">{ret}</b>
              </div>

              <div style="padding:10px 10px;border:1px solid #edf0f5;border-radius:12px;background:#fafbfe;">
                <div style="margin-bottom:6px;">
                  <span style="color:#6b7280;">€/pers :</span>
                  <b style="color:#111111;">{ppp}</b>
                  <span style="color:#e5e7eb;padding:0 8px;">|</span>
                  <span style="color:#6b7280;">Total :</span>
                  <b style="color:#111111;">{tot}</b>
                </div>

                <div>
                  <span style="color:#6b7280;">Durée :</span>
                  <b style="color:#111111;">{dur}</b>
                  <span style="color:#e5e7eb;padding:0 8px;">|</span>
                  <span style="color:#6b7280;">Escales :</span>
                  <b style="color:#111111;">{stp}</b>
                </div>
              </div>

              <div style="margin-top:12px;">
                {link_html}
              </div>
            </div>
          </div>
        """)

    html = f"""
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>

      <body style="margin:0;padding:0;background:#f6f8fb;">
        <div style="padding:20px 18px;">

          <!-- ✅ Conteneur principal ARRONDI -->
          <div style="max-width:920px;margin:0 auto;background:#ffffff;border:1px solid #e7ebf3;border-radius:16px;overflow:hidden;">

            <!-- Header -->
            <div style="padding:18px 18px 12px 18px;border-bottom:1px solid #edf0f5;">
              <div style="font-family:Segoe UI, Arial;font-size:14px;color:#111111;line-height:1.45;">
                Bonjour  ET ,<br>
                Je suis <b>JARVIS</b>, l’assistant informatique d’Abdoullah.
              </div>

              <div style="height:10px;"></div>

              <div style="font-family:Segoe UI, Arial;font-size:13px;color:#333333;line-height:1.55;">
                <b>Critères de recherche appliqués :</b><br>
                Vol aller-retour au départ de <b>CDG (Paris)</b> vers <b>MAA (Chennai)</b> pour <b>5 passagers</b>,
                avec un départ entre le <b>05/07/2026</b> et le <b>10/07/2026</b>,
                un retour entre le <b>24/08/2026</b> et le <b>30/08/2026</b>,
                un maximum de <b>1 escale</b>,
                une durée maximale de <b>22 heures</b>,
                en excluant <b>Air India (AI)</b>, <b>British Airways (BA)</b> et <b>Lufthansa (LH)</b>.
              </div>

              <div style="height:12px;"></div>

              <div style="font-family:Segoe UI, Arial;font-size:13px;color:#333333;line-height:1.55;">
                Voici les <b>15 vols les moins chers</b> (triés par prix / personne).
                <div style="color:#6b7280;font-size:12px;margin-top:4px;line-height:1.4;">
                  Règle appliquée : pas deux offres avec la même <b>compagnie</b> et le même <b>jour de départ</b>.
                </div>
              </div>
            </div>

            <!-- ✅ Cards area -->
            <div style="padding:14px;background:#f6f8fb;">
              {''.join(cards_html) if cards_html else ''}

              <div style="height:8px;"></div>

              <div style="font-family:Segoe UI, Arial;font-size:12px;color:#333333;line-height:1.5;">
                Le tableau complet est en pièce jointe (PDF).
              </div>

              <div style="height:14px;"></div>

              <div style="font-family:Segoe UI, Arial;font-size:12px;color:#333333;">
                Cordialement,
              </div>

              <div style="height:8px;"></div>
            </div>

          </div>

        </div>
      </body>
    </html>
    """
    return html



def _mk_pdf_bytes_from_dicts(offers: List[Dict[str, Any]]) -> bytes:
    """
    PDF plus pro:
    - marges correctes
    - tableau avec grille fine
    - en-tête gris clair
    - alternance de lignes (gris très léger)
    - plus d’espace entre les lignes
    - lien "Ouvrir" simple (pas de gros bouton)
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # --- Layout (en points) ---
    left = 14 * mm
    right = 14 * mm
    top = 16 * mm
    bottom = 14 * mm
    table_w = w - left - right

    # --- Typo ---
    title = "Vols Kayak (triés par prix / personne)"
    c.setFont("Helvetica-Bold", 14)
    c.drawString(left, h - top, title)

    # Petite ligne de contexte (optionnel)
    c.setFont("Helvetica", 9)
    c.drawString(left, h - top - 6 * mm, f"Export généré le {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # --- Table geometry ---
    y = h - top - 14 * mm  # début du tableau sous le titre
    header_h = 8 * mm
    row_h = 7 * mm  # + d’espace entre les lignes

    # Colonnes (largeurs en mm -> converties en points)
    # Ajuste si tu veux plus de place sur "Compagnies"
    col_w_mm = [20, 20, 58, 16, 16, 18, 16, 14]  # total = 178mm
    col_w = [x * mm for x in col_w_mm]

    # Si total != table_w, on scale proportionnellement
    total_cols = sum(col_w)
    if abs(total_cols - table_w) > 1:
        scale = table_w / total_cols
        col_w = [cw * scale for cw in col_w]

    headers = ["Départ", "Retour", "Compagnies", "€/pers", "Total", "Durée", "Escales", "Lien"]

    def _draw_header(y_top: float):
        # fond header (gris clair)
        c.setFillColorRGB(0.94, 0.94, 0.94)
        c.rect(left, y_top - header_h, table_w, header_h, fill=1, stroke=0)

        # texte header
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 9)

        x = left
        for i, hd in enumerate(headers):
            # padding interne
            c.drawString(x + 2.2 * mm, y_top - header_h + 2.6 * mm, hd)
            x += col_w[i]

        # bordure + lignes verticales fines
        c.setLineWidth(0.4)
        c.setStrokeColorRGB(0.70, 0.70, 0.70)
        c.rect(left, y_top - header_h, table_w, header_h, fill=0, stroke=1)

        x = left
        for i in range(len(col_w) - 1):
            x += col_w[i]
            c.line(x, y_top - header_h, x, y_top)

        # ligne sous header
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

    # lignes (grille + alternance)
    c.setFont("Helvetica", 9)
    c.setLineWidth(0.3)
    c.setStrokeColorRGB(0.80, 0.80, 0.80)

    for idx, o in enumerate(offers):
        # saut de page si besoin
        if y - row_h < bottom:
            y = _new_page()
            _draw_header(y)
            y -= header_h

        # alternance légère
        if idx % 2 == 1:
            c.setFillColorRGB(0.98, 0.98, 0.98)
            c.rect(left, y - row_h, table_w, row_h, fill=1, stroke=0)

        # bordure ligne
        c.setStrokeColorRGB(0.85, 0.85, 0.85)
        c.rect(left, y - row_h, table_w, row_h, fill=0, stroke=1)

        # vertical lines
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

        # baseline texte
        text_y = y - row_h + 2.4 * mm
        x = left

        # Col 1: depart
        c.setFillColorRGB(0, 0, 0)
        c.drawString(x + 2.2 * mm, text_y, depart)
        x += col_w[0]

        # Col 2: retour
        c.drawString(x + 2.2 * mm, text_y, ret)
        x += col_w[1]

        # Col 3: compagnies
        c.drawString(x + 2.2 * mm, text_y, comp)
        x += col_w[2]

        # Col 4: €/pers (align right)
        c.drawRightString(x + col_w[3] - 2.2 * mm, text_y, ppp)
        x += col_w[3]

        # Col 5: total (align right)
        c.drawRightString(x + col_w[4] - 2.2 * mm, text_y, tot)
        x += col_w[4]

        # Col 6: durée
        c.drawString(x + 2.2 * mm, text_y, dur)
        x += col_w[5]

        # Col 7: escales
        c.drawString(x + 2.2 * mm, text_y, stp)
        x += col_w[6]

        # Col 8: lien (simple, discret)
        link_label = "Ouvrir" if url else "—"
        # bleu discret (pas un gros bouton)
        if url:
            c.setFillColorRGB(0.10, 0.35, 0.75)
        else:
            c.setFillColorRGB(0.45, 0.45, 0.45)

        # centre dans la dernière colonne
        link_x0 = x
        label_w = c.stringWidth(link_label, "Helvetica", 9)
        label_x = link_x0 + (col_w[7] - label_w) / 2
        c.setFont("Helvetica", 9)
        c.drawString(label_x, text_y, link_label)

        # zone cliquable (sur le texte seulement)
        if url:
            c.linkURL(
                url,
                (label_x, text_y - 1, label_x + label_w, text_y + 9),
                relative=0
            )

        # reset font/couleur
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica", 9)

        # next row
        y -= row_h

    c.save()
    buf.seek(0)
    return buf.getvalue()


# -----------------------------
# OUTLOOK SEND (HTML) + VERIFIED
# -----------------------------
olFolderOutbox = 4
olFolderSentMail = 5

def _get_outlook_account_by_smtp(outlook_app, smtp: str):
    ns = outlook_app.GetNamespace("MAPI")
    for acc in ns.Accounts:
        try:
            if str(acc.SmtpAddress).lower() == smtp.lower():
                return acc
        except Exception:
            continue
    return None

def _get_store_default_folder_for_account(account, folder_const):
    try:
        store = account.DeliveryStore
        return store.GetDefaultFolder(folder_const)
    except Exception:
        return None

def _outlook_try_find_by_subject(folder, subject: str, max_check: int = 40) -> bool:
    try:
        items = folder.Items
        items.Sort("[CreationTime]", True)  # desc
        m = min(max_check, items.Count)
        for i in range(1, m + 1):
            it = items.Item(i)
            try:
                if str(it.Subject) == subject:
                    return True
            except Exception:
                continue
    except Exception:
        return False
    return False

def send_outlook_mail_with_attachment_auto_verified(
    sender_smtp: str,
    to_list: List[str],
    subject: str,
    html_body: str,
    attachment_path: str,
    wait_seconds: int = 6
):
    pythoncom.CoInitialize()

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    account = _get_outlook_account_by_smtp(outlook, sender_smtp)
    if account is None:
        raise RuntimeError(f"Compte Outlook introuvable: {sender_smtp}. Vérifie qu'il est bien ajouté dans Outlook.")

    mail = outlook.CreateItem(0)
    mail.To = ";".join(to_list)
    mail.Subject = subject

    # HTML
    mail.HTMLBody = html_body

    # PJ
    mail.Attachments.Add(os.path.abspath(attachment_path))

    # Force account
    mail.SendUsingAccount = account

    sent_folder = _get_store_default_folder_for_account(account, olFolderSentMail)
    if sent_folder is not None:
        try:
            mail.SaveSentMessageFolder = sent_folder
        except Exception:
            pass

    try:
        mail.SentOnBehalfOfName = sender_smtp
    except Exception:
        pass

    #mail.Save()
    #mail.Send()
    mail.Display()   # Ouvre le mail dans Outlook (brouillon)

    

    try:
        ns.SendAndReceive(True)
    except Exception:
        pass

    time.sleep(max(1, wait_seconds))

    outbox = _get_store_default_folder_for_account(account, olFolderOutbox)
    if outbox is not None and _outlook_try_find_by_subject(outbox, subject, max_check=60):
        raise RuntimeError(
            "Outlook a créé le mail mais il est resté dans la Boîte d'envoi. "
            "Vérifie: mode 'Travail hors connexion', 'Envoyer plus tard', règles Outlook, blocage sécurité."
        )

    # Best-effort: si accessible, on check qu'il est dans Sent
    if sent_folder is not None:
        _outlook_try_find_by_subject(sent_folder, subject, max_check=80)

    return True

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
    close_browser()
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

@app.route("/export_pdf_and_outlook_send", methods=["POST"])
def export_pdf_and_outlook_send():
    data = request.get_json(force=True, silent=False)
    offers = data.get("offers", [])
    offers_sorted = sort_offers_by_price_dicts(offers)

    # PDF bytes
    pdf_bytes = _mk_pdf_bytes_from_dicts(offers_sorted)

    # write temp file
    tmp_dir = tempfile.gettempdir()
    pdf_path = os.path.join(tmp_dir, f"vols_kayak_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)

    today = dt.datetime.now().strftime("%Y-%m-%d")
    subject = f"Vols Kayak  - {today}"

    # IMPORTANT: HTML body (tableau propre)
    html_body = _mk_outlook_body_top8_html(offers_sorted)

    try:
        send_outlook_mail_with_attachment_auto_verified(
            sender_smtp=OUTLOOK_SENDER_SMTP,
            to_list=OUTLOOK_TO,
            subject=subject,
            html_body=html_body,
            attachment_path=pdf_path,
            wait_seconds=6
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

    return jsonify({"ok": True})

# -----------------------------
# MAIN
# -----------------------------
