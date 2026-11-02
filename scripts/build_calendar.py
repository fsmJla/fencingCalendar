#!/usr/bin/env python3
"""Génère la page HTML statique du calendrier escrime consolidé.

Sources :
  1. Grille FFE de la ligue (Excel « Calendrier FFE NA ») : circuit national.
  2. Calendrier des épreuves Nouvelle-Aquitaine (Excel) : épreuves régionales et stages.
  3. Agenda du site de la ligue (API REST « The Events Calendar ») : épreuves
     régionales publiées sur le site et absentes des Excel.
  4. data/evenements_manuels.json : épreuves départementales Gironde et corrections.

Chaque source téléchargée est mise en cache dans data/sources/ ; si un
téléchargement échoue, la dernière copie connue est utilisée, pour qu’une
panne du site de la ligue ne casse jamais la page.

Usage : python scripts/build_calendar.py [--offline] [--out dist/index.html]
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import re
import sys
import unicodedata
from collections import OrderedDict
from io import BytesIO
from pathlib import Path
from urllib.parse import urljoin

import openpyxl
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
CONFIG = ROOT / "config.json"
SOURCES = ROOT / "data" / "sources"
TEMPLATE = ROOT / "template" / "calendar.html"

UA = "Mozilla/5.0 (compatible; calendrier-escrime-pessac/1.0; +https://github.com/fsmJla/fencingCalendar)"

MONTHS_FR = {
    "janvier": 1, "janv": 1, "jan": 1,
    "fevrier": 2, "fevr": 2, "fev": 2, "fevier": 2,
    "mars": 3, "mar": 3,
    "avril": 4, "avr": 4,
    "mai": 5,
    "juin": 6,
    "juillet": 7, "juil": 7,
    "aout": 8,
    "septembre": 9, "sept": 9, "sep": 9,
    "octobre": 10, "oct": 10,
    "novembre": 11, "nov": 11,
    "decembre": 12, "dec": 12,
}
MN = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
      "septembre", "octobre", "novembre", "décembre"]

LOG: list[str] = []      # messages techniques (console)
NOTES: list[str] = []    # points à vérifier affichés sur la page


def log(msg: str) -> None:
    LOG.append(msg)
    print(msg, file=sys.stderr)


def norm(s) -> str:
    """minuscules, sans accents, espaces normalisés."""
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def clean(s) -> str:
    return re.sub(r"\s+", " ", "" if s is None else str(s)).strip()


def iso(d: dt.date) -> str:
    return d.isoformat()


# --------------------------------------------------------------------------
# Téléchargement
# --------------------------------------------------------------------------

def http_get(url: str, timeout: int = 30, tries: int = 3) -> bytes:
    """GET avec quelques nouvelles tentatives en cas de coupure réseau."""
    import time
    import requests  # import tardif : inutile en mode hors ligne
    last = None
    for i in range(tries):
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=timeout)
            r.raise_for_status()
            return r.content
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2 * (i + 1))
    raise last


def discover_excel_links(pages: list[str]) -> list[str]:
    links: list[str] = []
    for page in pages:
        try:
            body = http_get(page).decode("utf-8", "replace")
        except Exception as e:  # noqa: BLE001
            log(f"[sources] page inaccessible {page} : {e}")
            continue
        for href in re.findall(r'href=["\']([^"\']+\.xlsx?)(?:\?[^"\']*)?["\']', body, re.I):
            url = urljoin(page, html.unescape(href))
            if url not in links:
                links.append(url)
    log(f"[sources] {len(links)} fichier(s) Excel trouvés : {links}")
    return links


def classify_workbook(wb) -> str | None:
    """'ffe' pour la grille FFE, 'na' pour le calendrier régional."""
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 120), max_col=12, values_only=True):
            for v in row:
                n = norm(v)
                if n.startswith("intitule de l"):
                    return "na"
                if n == "vacances scolaires":
                    return "ffe"
    return None


def load_excels(cfg: dict, offline: bool) -> dict[str, tuple[openpyxl.Workbook, str]]:
    """Retourne {'ffe': (wb, nom), 'na': (wb, nom)}."""
    found: dict[str, tuple[openpyxl.Workbook, str]] = {}
    if not offline:
        urls = list(cfg.get("excel_urls", []))
        urls += [u for u in discover_excel_links(cfg.get("pages", [])) if u not in urls]
        for url in urls:
            try:
                data = http_get(url)
                wb = openpyxl.load_workbook(BytesIO(data))
            except Exception as e:  # noqa: BLE001
                log(f"[sources] échec {url} : {e}")
                continue
            kind = classify_workbook(wb)
            if not kind or kind in found:
                continue
            name = url.rsplit("/", 1)[-1]
            found[kind] = (wb, name)
            cache = SOURCES / f"{CACHE_NAMES[kind]}"
            cache.write_bytes(data)
            (SOURCES / f"{CACHE_NAMES[kind]}.url").write_text(url + "\n", encoding="utf-8")
            log(f"[sources] {kind} ← {url}")
    for kind, fname in CACHE_NAMES.items():
        if kind in found:
            continue
        path = SOURCES / fname
        if path.exists():
            wb = openpyxl.load_workbook(path)
            urlfile = SOURCES / f"{fname}.url"
            name = urlfile.read_text().strip().rsplit("/", 1)[-1] if urlfile.exists() else fname
            found[kind] = (wb, name)
            log(f"[sources] {kind} ← copie locale {path.name}")
        else:
            log(f"[sources] AUCUNE source pour {kind}")
    return found


CACHE_NAMES = {"ffe": "calendrier-ffe-na.xlsx", "na": "calendrier-epreuves-na.xlsx"}


# --------------------------------------------------------------------------
# Catégories
# --------------------------------------------------------------------------

CAT_SEQ = ["M9", "M11", "M13", "M15", "M17", "M20", "SEN", "VET"]
CAT_KEY = {"M9": "M9", "M11": "M11", "M13": "M13", "M15": "M13", "M17": "M17", "M20": "M20", "SEN": "SEN", "VET": "VET"}
CAT_HUMAN = {"M9": "M9", "M11": "M11", "M13": "M13", "M15": "M15", "M17": "M17", "M20": "M20", "SEN": "Seniors", "VET": "Vétérans"}


def cat_tokens(text: str) -> list[str]:
    n = norm(text)
    toks = []
    for m in re.finditer(r"m\s?(9|11|13|15|17|20|23)\b|senior|veteran|cadet|junior|minime|benjamin", n):
        t = m.group(0)
        if t.startswith("m"):
            num = m.group(1)
            toks.append("M20" if num == "23" else "M" + num)
        elif t.startswith("senior"):
            toks.append("SEN")
        elif t.startswith("veteran"):
            toks.append("VET")
        elif t == "cadet":
            toks.append("M17")
        elif t == "junior":
            toks.append("M20")
        elif t == "minime":
            toks.append("M15")
        elif t == "benjamin":
            toks.append("M13")
    return toks


def parse_categories(text: str) -> tuple[dict[str, str], str]:
    """« M15 A SENIORS » → ({M13,M17,M20,SEN}, 'M15 à Seniors')."""
    n = norm(text)
    toks = cat_tokens(text)
    if not toks:
        return {}, clean(text)
    sel: list[str] = []
    m = re.search(r"(m\s?\d{1,2}|senior|veteran)[^a]*?\s(a|au)\s+(m\s?\d{1,2}|senior|veteran)", n)
    if m and len(toks) >= 2:
        a, b = CAT_SEQ.index(toks[0]), CAT_SEQ.index(toks[1])
        sel = CAT_SEQ[min(a, b):max(a, b) + 1]
        human = f"{CAT_HUMAN[toks[0]]}" + (" (H2036)" if "2036" in n else "") + f" à {CAT_HUMAN[toks[1]]}"
        rest = toks[2:]
        sel += rest
        if rest:
            human += " + " + " + ".join(CAT_HUMAN[t] for t in rest)
    else:
        sel = toks
        seen = list(OrderedDict.fromkeys(toks))
        sep = " / " if "/" in n or " et " in n else " + "
        human = sep.join(CAT_HUMAN[t] for t in seen)
        if "open" in n:
            human += " open"
        if "1ere annee" in n or "1re annee" in n:
            human = clean(text).replace("1ère", "1re")
            human = re.sub(r"\bet\b", "et", human)
    cats = {CAT_KEY[t]: "DH" for t in sel}
    return cats, human


# --------------------------------------------------------------------------
# Dates
# --------------------------------------------------------------------------

def season_year_for(month: int, start_year: int) -> int:
    return start_year if month >= 8 else start_year + 1


def parse_fr_dates(value, start_year: int, ndays: int | None = None):
    """Texte libre (« 26/27 septembre », « 19 au 23 oct », « 31 et 1 novembre »,
    « 6/7 février »…) ou datetime → (début, fin) ou None."""
    if isinstance(value, dt.datetime):
        d = value.date()
        y = season_year_for(d.month, start_year)
        if d.year != y:
            d = d.replace(year=y)
        e = d + dt.timedelta(days=(ndays or 1) - 1)
        return d, e
    if isinstance(value, dt.date):
        return value, value + dt.timedelta(days=(ndays or 1) - 1)
    n = norm(value)
    if not n or "definir" in n:
        return None
    n = n.split(":", 1)[-1] if ":" in n else n
    months = [(m.start(), MONTHS_FR[m.group(1)]) for m in
              re.finditer(r"\b(" + "|".join(sorted(MONTHS_FR, key=len, reverse=True)) + r")\b\.?", n)]
    if not months:
        return None
    days = [(m.start(), int(m.group(1))) for m in re.finditer(r"\b(\d{1,2})(?:er)?\b", n)]
    days = [(p, d) for p, d in days if 1 <= d <= 31]
    if not days:
        return None
    # chaque jour prend le premier mois écrit après lui
    resolved = []
    for p, d in days:
        after = [mo for mp, mo in months if mp > p]
        resolved.append((d, after[0] if after else months[-1][1], bool(after) and after[0] != months[-1][1]))
    d1, m1, explicit1 = resolved[0]
    d2, m2, _ = resolved[-1]
    if len(resolved) > 1 and not explicit1 and d1 > d2 and m1 == m2:
        m1 = 12 if m2 == 1 else m2 - 1  # « 31 et 1 novembre »
    s = dt.date(season_year_for(m1, start_year), m1, d1)
    e = dt.date(season_year_for(m2, start_year), m2, d2)
    if len(resolved) == 1 and ndays:
        e = s + dt.timedelta(days=ndays - 1)
    if e < s:
        e = s
    return s, e


def parse_ndays(v) -> int | None:
    if isinstance(v, (int, float)):
        return int(v)
    m = re.search(r"(\d+)\s*j", norm(v))
    return int(m.group(1)) if m else None


# --------------------------------------------------------------------------
# Grille FFE (circuit national)
# --------------------------------------------------------------------------

FFE_ROWS = {"veterans": "VET", "seniors": "SEN", "m20": "M20", "m17": "M17", "m13/m15": "M13"}
WEAPON_WORDS = {"epee": "EPEE", "sabre": "SABRE", "fleuret": "FLEURET"}


def fill_key(cell):
    f = cell.fill
    if not f or f.fill_type is None:
        return None
    c = f.fgColor
    if c.type == "rgb":
        return ("rgb", c.rgb)
    return (c.type, getattr(c, c.type, None), round(c.tint or 0, 3))


def merged_value_map(ws):
    m = {}
    for rng in ws.merged_cells.ranges:
        v = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                m[(r, c)] = v
    return m


def ffe_title(text: str) -> str | None:
    t = clean(text)
    # « Monaco/ER 2 » : on retire la partie régionale
    parts = [p.strip() for p in re.split(r"/", t)]
    parts = [p for p in parts if not re.match(r"(?i)^ER\s*\d", p)] if "/" in t and re.search(r"(?i)\bER\s*\d", t) else [t]
    t = "/".join(parts).strip()
    if not t or re.match(r"(?i)^(ER\b|a definir|à définir)", norm(t)):
        return None
    return t


def parse_ffe(wb, start_year: int):
    events = []
    vac: dict[str, str] = {}
    weekends_seen = []
    region_cells: list[dict] = []
    for ws in wb.worksheets:
        mv = merged_value_map(ws)
        val = lambda r, c: mv.get((r, c), ws.cell(r, c).value)  # noqa: E731
        # ligne des mois et ligne des week-ends
        month_row = days_row = vac_row = None
        for r in range(1, min(ws.max_row, 10) + 1):
            texts = [norm(ws.cell(r, c).value) for c in range(3, ws.max_column + 1)]
            if month_row is None and any(re.match(r"(" + "|".join(MONTHS_FR) + r")\b.*20\d\d", t) for t in texts):
                month_row = r
            if days_row is None and sum(bool(re.fullmatch(r"\d{1,2}(-\d{1,2})+", t)) for t in texts) >= 3:
                days_row = r
        for r in range(1, ws.max_row + 1):
            if norm(ws.cell(r, 1).value).startswith("vacances scolaires"):
                vac_row = r
        if month_row is None or days_row is None:
            log(f"[ffe] feuille {ws.title} ignorée (en-têtes introuvables)")
            continue
        col_dates: dict[int, tuple[dt.date, dt.date]] = {}
        for c in range(3, ws.max_column + 1):
            mtxt = norm(val(month_row, c))
            dtxt = clean(ws.cell(days_row, c).value)
            m = re.match(r"(" + "|".join(sorted(MONTHS_FR, key=len, reverse=True)) + r")\w*\s+(20\d\d)", mtxt)
            if not m or not re.fullmatch(r"\d{1,2}(-\d{1,2})+", dtxt):
                continue
            month, year = MONTHS_FR[m.group(1)], int(m.group(2))
            ds = [int(x) for x in dtxt.split("-")]
            # « 31-1 » sous « Octobre » ; « 2-3 » sous « Janvier » etc.
            s = dt.date(year, month, ds[0])
            last = ds[-1]
            e_month, e_year = (month, year) if last >= ds[0] else ((1, year + 1) if month == 12 else (month + 1, year))
            e = dt.date(e_year, e_month, last)
            col_dates[c] = (s, e)
            weekends_seen.append(s)
            if vac_row:
                zones = clean(ws.cell(vac_row, c).value)
                if zones:
                    d = s
                    while d <= e:
                        vac[iso(d)] = zones
                        d += dt.timedelta(days=1)
        weapon = sex = None
        for r in range(days_row + 1, ws.max_row + 1):
            a = norm(val(r, 1))
            if a:
                weapon = next((k for w, k in WEAPON_WORDS.items() if a.startswith(w)), None)
                sex = "D" if "dame" in a else "H" if "homme" in a else "DH"
                if "laser" in a or a in ("sl", "art") or a.startswith("para"):
                    weapon = None
            b = norm(ws.cell(r, 2).value)
            cat = FFE_ROWS.get(b.replace(" ", ""))
            if not weapon or not cat:
                continue
            row_fill = fill_key(ws.cell(r, 2))
            for c, (s, e) in col_dates.items():
                cell = ws.cell(r, c)
                if cell.value is None:
                    continue
                # cellules d’une autre couleur que la ligne = épreuves régionales
                # (« ER 1 », « Bayonne », « Formation »…) : reprises du calendrier NA
                if fill_key(cell) != row_fill:
                    region_cells.append({"arm": weapon, "s": iso(s), "e": iso(e), "cat": cat,
                                         "text": clean(cell.value)})
                    # « Monaco/ER 2 » : la partie nationale est gardée
                    if not ("/" in str(cell.value) and re.search(r"(?i)\bER\s*\d", str(cell.value))):
                        continue
                title = ffe_title(cell.value)
                if not title:
                    continue
                events.append({"s": iso(s), "e": iso(e), "lv": "FR", "arm": weapon,
                               "cat": cat, "sex": sex, "title": title})
    # fusion : sexes puis catégories
    merged: "OrderedDict[tuple, dict]" = OrderedDict()
    for ev in events:
        k = (ev["s"], ev["e"], ev["arm"], ev["title"].lower(), ev["cat"])
        if k in merged:
            prev = merged[k]["sex"]
            merged[k]["sex"] = "".join(x for x in "DH" if x in prev + ev["sex"])
        else:
            merged[k] = dict(ev)
    out: "OrderedDict[tuple, dict]" = OrderedDict()
    for ev in merged.values():
        k = (ev["s"], ev["e"], ev["arm"], ev["title"].lower())
        o = out.setdefault(k, {"s": ev["s"], "e": ev["e"], "lv": "FR", "arms": [ev["arm"]], "cats": {},
                               "title": ev["title"], "catText": None, "detail": "", "warn": None,
                               "src": "Grille FFE"})
        o["cats"][ev["cat"]] = ev["sex"]
    # un même titre peut regrouper des sexes différents par catégorie : on scinde
    final = []
    for o in out.values():
        groups: dict[str, dict] = {}
        for cat, sx in o["cats"].items():
            groups.setdefault(sx, {})[cat] = sx
        for sx, cats in groups.items():
            final.append({**o, "cats": cats})
    log(f"[ffe] {len(final)} épreuves nationales")
    return final, vac, region_cells


# --------------------------------------------------------------------------
# Calendrier régional Nouvelle-Aquitaine
# --------------------------------------------------------------------------

def fmt_format(text: str) -> str:
    t = clean(text)
    if not t:
        return ""
    n = norm(t)
    rep = [
        (r"\bindiv(iduel)?\.?\b", "Individuel"),
        (r"\b(equipes?|eq)\b", "Équipe"),
    ]
    out = n
    for pat, r in rep:
        out = re.sub(pat, r, out)
    out = re.sub(r"\s*/\s*", " + ", out)
    out = re.sub(r"\s+et\s+", ", et ", out)
    out = out.replace("a l'epee", "à l’épée").replace("equipe", "Équipe")
    out = re.sub(r"\bm(\d\d)\b", lambda m: "M" + m.group(1), out)
    if re.fullmatch(r"(Individuel|Équipe)s?(\s*\+\s*(Individuel|Équipe)s?)*(\s+M\d\d)?", out):
        return out
    if "indiv" not in n and not re.search(r"\beq", n):
        return sentence_case(t)  # « Combat, Chorégraphie, Katas » : tel quel
    return out[0].upper() + out[1:] if out else ""


def sentence_case(t: str) -> str:
    t = clean(t)
    letters = [c for c in t if c.isalpha()]
    if letters and sum(c.isupper() for c in letters) / len(letters) > 0.7:
        t = t.lower()
        t = t[0].upper() + t[1:]
    return t


def na_title(raw: str, place: str | None, fmt: str) -> str:
    t = clean(raw)
    n = norm(t)
    if "zone" in n and "1/4" in n:
        title = "1/4 de finale H2036 (zone NA + Occitanie)"
    else:
        t = re.sub(r"(?i)[ée]preuve\s+(\d)\s+nouvelle[- ]aquitaine", r"Épreuve \1 NA", t)
        t = re.sub(r"(?i)[ée]preuve\s+r[ée]gi?o?nale\s+(\d)", r"Épreuve régionale \1", t)
        t = re.sub(r"(?i)championnats?\s+nouvelle[- ]aquitaine", lambda m: ("Championnats" if m.group(0).lower().startswith("championnats") else "Championnat") + " NA", t)
        t = re.sub(r"(?i)coupe nouvelle[- ]aquitaine", "Coupe NA", t)
        t = re.sub(r"(?i)^1/8 H2036$", "1/8 de finale H2036", t.strip())
        t = re.sub(r"(?i)\bopen de ([A-Z]+)\b", lambda m: "Open de " + m.group(1).capitalize(), t)
        t = clean(t)
        if re.match(r"(?i)^championnats? na$", t) and fmt == "Équipe":
            t += " par équipes"
        title = t[0].upper() + t[1:] if t else "Épreuve"
    if place and norm(place) not in norm(title):
        title += f" – {place}"
    return title


def pretty_place(p: str) -> str:
    p = clean(p)
    if not p:
        return ""
    if p.isupper() and len(p) >= 3 and p not in ("BBE", "BEC", "CAM"):
        p = p.title()
    p = re.sub(r"\s*\?$", " (à confirmer)", p)
    p = {"Cam": "Bordeaux CAM", "Angoûleme": "Angoulême", "Villeneuve sur Lot": "Villeneuve-sur-Lot", "Grad Talence": "Grad. Talence", "Bdx (à confirmer)": "Bordeaux (à confirmer)"}.get(p, p)
    return p


def parse_na(wb, start_year: int):
    events = []
    for ws in wb.worksheets:
        mv = merged_value_map(ws)
        val = lambda r, c: mv.get((r, c), ws.cell(r, c).value)  # noqa: E731
        header = None
        cols: dict[str, int] = {}
        section = None   # 'epreuves' | 'stages' | 'stages_anim' | 'stages_arb'
        stage_cols: dict[str, int] = {}
        weapon_label = None
        for r in range(1, ws.max_row + 1):
            rowtxt = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
            joined = " | ".join(rowtxt)
            if any(t.startswith("intitule de l") for t in rowtxt):
                header = r
                section = "epreuves"
                for c, t in enumerate(rowtxt, 1):
                    if t.startswith("intitule"): cols["title"] = c
                    elif t.startswith("categorie"): cols["cats"] = c
                    elif t.startswith("indiv"): cols["fmt"] = c
                    elif t.startswith("date"): cols["date"] = c
                    elif t.startswith("candidature"): cols["cand"] = c
                    elif t.startswith("proposition") or t.startswith("lieu"): cols["place"] = c
                # colonne sans en-tête juste après le tableau : remarques (« formation arbitrage épée »)
                last_named = max(c for c, t in enumerate(rowtxt, 1) if t)
                cols["note"] = last_named + 1
                continue
            if "calendrier des dates et lieux de stages" in joined:
                section = "stages_arb" if "arbitrage" in joined else "stages_anim" if "anim" in joined else "stages"
                stage_cols = {}
                continue
            if section and section.startswith("stages") and "dates" in rowtxt:
                for c, t in enumerate(rowtxt, 1):
                    if t == "dates": stage_cols["date"] = c
                    elif t.startswith("nombre"): stage_cols["ndays"] = c
                    elif t.startswith("categorie"): stage_cols["cats"] = c
                    elif t.startswith("lieu"): stage_cols["place"] = c
                if norm(val(r, 1)):
                    weapon_label = norm(val(r, 1))
                continue

            if section == "epreuves":
                a = norm(val(r, 1))
                if a:
                    weapon_label = a
                raw = ws.cell(r, cols["title"]).value
                if not clean(raw):
                    continue
                arm = WEAPON_WORDS.get(weapon_label)
                is_laser = "laser" in (weapon_label or "")
                if "artistique" in (weapon_label or ""):
                    continue  # escrime artistique : non reprise
                if is_laser:
                    arm = "LASER"
                    raw = re.sub(r"(?i)\s*sabre laser\s*", " ", str(raw)).strip()
                cat_raw = clean(val(r, cols.get("cats", 0)) if cols.get("cats") else "")
                arms = [arm] if arm else []
                if is_laser:
                    pass
                elif "3 armes" in norm(raw):
                    arms = ["EPEE", "SABRE", "FLEURET"]
                ncat = norm(cat_raw)
                if "epee" in ncat or "fleuret" in ncat or "sabre" in ncat:
                    arms = [k for w, k in WEAPON_WORDS.items() if w in ncat]
                    cat_raw = re.sub(r"(?i)\s*(ep[ée]e|fleuret|sabre)(/|\s|$)", " ", cat_raw).strip()
                if not arms:
                    continue
                dval = val(r, cols["date"])
                year_fixed = None
                if isinstance(dval, dt.datetime) and dval.year != season_year_for(dval.month, start_year):
                    year_fixed = f"{dval:%d/%m/%Y}"  # année erronée dans le fichier, corrigée plus bas
                fmt = fmt_format(val(r, cols["fmt"]) if cols.get("fmt") else "")
                dates = parse_fr_dates(dval, start_year)
                place = pretty_place(clean(ws.cell(r, cols["place"]).value) if cols.get("place") else "")
                cand = clean(val(r, cols["cand"])) if cols.get("cand") else ""
                if not place and cand and "(" not in cand and "/" not in cand and "attente" not in norm(cand):
                    place = pretty_place(cand)
                title = na_title(raw, place, fmt)
                if not dates:
                    armnames = {"EPEE": "Épée", "SABRE": "Sabre", "FLEURET": "Fleuret", "LASER": "Sabre laser"}
                    NOTES.append(f"{title} ({', '.join(armnames[a] for a in arms)}) : date « {clean(dval) or 'non renseignée'} » non exploitable, épreuve non placée.")
                    continue
                cats, cat_text = parse_categories(cat_raw)
                if is_laser and cats:
                    seq = [t for t in CAT_SEQ if t in cat_tokens(cat_raw)]
                    cat_text = f"{CAT_HUMAN[seq[0]]} à {CAT_HUMAN[seq[-1]]}"
                details = [fmt] if fmt else []
                if "attente" in norm(cand):
                    details.append(cand[0].lower() + cand[1:])
                note = clean(ws.cell(r, cols["note"]).value)
                if note:
                    details.append(note)
                if not place:
                    NOTES.append(f"{title} ({dates[0].day} {MN[dates[0].month - 1]}) : aucun lieu dans le calendrier NA.")
                warn = None
                if year_fixed:
                    warn = (f"{title} : le calendrier NA indique le {year_fixed}, "
                            f"année corrigée en {dates[0].year}")
                events.append({"s": iso(dates[0]), "e": iso(dates[1]), "lv": "NA", "arms": arms,
                               "cats": cats, "title": title, "catText": cat_text,
                               "detail": " · ".join(details), "warn": warn, "src": "Calendrier NA",
                               "_year_fixed": bool(year_fixed)})

            elif section and section.startswith("stages") and stage_cols:
                a = norm(val(r, 1))
                dval = ws.cell(r, stage_cols["date"]).value
                if section == "stages" and a:
                    weapon_label = a
                if dval is None or clean(dval) == "":
                    continue
                ndays = parse_ndays(ws.cell(r, stage_cols["ndays"]).value) if stage_cols.get("ndays") else None
                place = pretty_place(clean(val(r, stage_cols["place"])) if stage_cols.get("place") else "")
                cat_raw = clean(val(r, stage_cols["cats"])) if stage_cols.get("cats") else ""
                dates = parse_fr_dates(dval, start_year, ndays)
                label = clean(str(dval).split(":")[0]) if isinstance(dval, str) and ":" in dval else ""
                if section == "stages_anim":
                    arms, cats, cat_text = ["ANIM"], {}, "Animateur / Éducateur"
                    title = f"{label or 'Stage'} Animateur / Éducateur" + (f" – {place}" if place else "")
                elif section == "stages_arb":
                    arms = ["EPEE", "SABRE", "FLEURET"]
                    cats, cat_text = {k: "DH" for k in ("M13", "M17", "M20", "SEN", "VET")}, "Arbitrage"
                    title = f"{label or 'Stage'} arbitrage" + (f" – {place}" if place else "")
                else:
                    arm = WEAPON_WORDS.get(weapon_label)
                    if not arm:
                        continue
                    arms = [arm]
                    cats, cat_text = parse_categories(cat_raw)
                    wname = {"EPEE": "épée", "SABRE": "sabre", "FLEURET": "fleuret"}[arm]
                    title = f"Stage {wname}" + (f" – {place}" if place else "")
                if not dates:
                    NOTES.append(f"{title} : date « {clean(dval)} » non exploitable, stage non placé.")
                    continue
                detail = ""
                span = (dates[1] - dates[0]).days + 1
                if ndays and span != ndays and section == "stages":
                    detail = f"{ndays} jours selon le fichier"
                    NOTES.append(f"{title} : dates « {clean(dval)} » ({span} jours) alors que le fichier indique « {ndays} jours » ; affiché tel qu’écrit.")
                elif ndays and section == "stages":
                    detail = f"{ndays} jours"
                events.append({"s": iso(dates[0]), "e": iso(dates[1]), "lv": "ST", "arms": arms,
                               "cats": cats, "title": title, "catText": cat_text, "detail": detail,
                               "warn": None, "src": "Calendrier NA (stages)"})
    log(f"[na] {len(events)} épreuves régionales et stages")
    return events


GRID_NOT_PLACE = {"chpt", "indiv", "indiv.", "equipe", "formation", "certification", "zone", "bec", "m13", "m15",
                  "tal", "tal.", "grad", "1/8 h2036", "open dax m15"}


def grid_place(text: str) -> str:
    """« ER 3 - Anglet » → « Anglet » ; « ER 2 » → ''."""
    m = re.match(r"(?i)^ER\s*\d+\s*(?:M1[35])?\s*[-–]?\s*(.*)$", text)
    rest = clean(m.group(1)) if m else ""
    rest = re.sub(r"(?i)\b(M1[35]|tal\.?)\b", "", rest).strip(" -–")
    return "" if not rest or norm(rest) in GRID_NOT_PLACE or len(rest) < 3 else rest


def reconcile_with_grid(na_events: list[dict], cells: list[dict]) -> list[dict]:
    """Compare les « Épreuve n NA » aux cases « ER n » de la grille FFE :
    complète un lieu manquant et signale les dates qui divergent."""
    extra = []
    for ev in na_events:
        m = re.match(r"Épreuve (\d) NA", ev["title"])
        if ev["lv"] != "NA" or not m or len(ev["arms"]) != 1:
            continue
        n, arm = m.group(1), ev["arms"][0]
        vet_like = "VET" in ev["cats"]
        rows = {"VET", "M13"} if vet_like else {"SEN", "M20", "M17"}
        grid = [c for c in cells if c["arm"] == arm and c["cat"] in rows
                and re.search(rf"(?i)\bER\s*{n}\b", c["text"])]
        if not grid:
            continue
        same = [c for c in grid if c["s"] == ev["s"]]
        if same:
            if " – " not in ev["title"]:
                places = [grid_place(c["text"]) for c in same]
                # le lieu peut être écrit sur une autre ligne du même week-end
                places += [grid_place("ER 0 " + c["text"]) for c in cells
                           if c["arm"] == arm and c["s"] == ev["s"] and c["cat"] in rows
                           and not re.search(r"(?i)\bER\b", c["text"])]
                place = next((p for p in places if p), "")
                if place:
                    ev["title"] += f" – {place} (lieu d’après la grille FFE)"
                    NOTES[:] = [x for x in NOTES if not x.startswith(ev["title"].split(" – ")[0] + " (")]
                    NOTES.append(f"{ev['title'].split(' – ')[0]} ({'vétérans + M13' if vet_like else arm.lower()}) : aucun lieu dans le calendrier NA ; {place} vient de la grille FFE.")
            continue
        other = sorted({(c["s"], c["e"]) for c in grid})[0]
        d1, d2 = dt.date.fromisoformat(ev["s"]), dt.date.fromisoformat(other[0])
        if abs((d1 - d2).days) > 60:
            continue
        wname = {"EPEE": "Épée", "SABRE": "Sabre", "FLEURET": "Fleuret"}[arm]
        fmt = lambda a, b: f"{dt.date.fromisoformat(a).day}-{dt.date.fromisoformat(b).day} {MN[dt.date.fromisoformat(b).month - 1]}"  # noqa: E731
        warn = (f"{wname.lower()}, {ev['title']} : calendrier NA = {fmt(ev['s'], ev['e'])}, "
                f"grille FFE = {fmt(*other)}")
        ev["warn"] = warn
        extra.append({**ev, "s": other[0], "e": other[1], "title": ev["title"] + " (date selon la grille FFE)",
                      "detail": ev["detail"].split(" · ")[0]})
        NOTES.append(f"{wname}, {ev['title']} : le calendrier NA indique le {fmt(ev['s'], ev['e'])}, la grille FFE le {fmt(*other)}. Les deux dates sont affichées avec la mention « À confirmer ».")
    return na_events + extra


# --------------------------------------------------------------------------
# Agenda du site de la ligue (The Events Calendar)
# --------------------------------------------------------------------------

def fetch_tec_events(cfg: dict, season: tuple[dt.date, dt.date], offline: bool) -> list[dict]:
    api = cfg.get("events_api")
    cache = SOURCES / "agenda-ligue.json"
    raw: list[dict] = []
    if api and not offline:
        try:
            import requests
            url = f"{api}?per_page=50&start_date={season[0]}&end_date={season[1]}"
            while url:
                r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
                r.raise_for_status()
                j = r.json()
                raw += j.get("events", [])
                url = j.get("next_rest_url")
            cache.write_text(json.dumps(raw, ensure_ascii=False, indent=1), encoding="utf-8")
            log(f"[agenda] {len(raw)} événements lus sur le site de la ligue")
        except Exception as e:  # noqa: BLE001
            log(f"[agenda] échec de l’API ({e}) ; copie locale utilisée si présente")
            raw = []
    if not raw and cache.exists():
        raw = json.loads(cache.read_text(encoding="utf-8"))
    out = []
    for ev in raw:
        title = html.unescape(re.sub(r"<[^>]+>", "", ev.get("title", ""))).strip()
        cats = [norm(c.get("slug") or c.get("name")) for c in ev.get("categories", [])]
        if any("laser" in c for c in cats) or "laser" in norm(title):
            arms = ["LASER"]
        else:
            arms = [k for w, k in WEAPON_WORDS.items() if any(w in c for c in cats) or norm(title).startswith(w)]
        if "artistique" in norm(title):
            continue
        if not arms:
            continue
        try:
            s = dt.date.fromisoformat(ev["start_date"][:10])
            e = dt.date.fromisoformat(ev["end_date"][:10])
        except Exception:  # noqa: BLE001
            continue
        cats_d, cat_text = parse_categories(title)
        venue = ev.get("venue") or {}
        city = clean(venue.get("city") if isinstance(venue, dict) else "")
        if not city:  # « … à Dax », « … (Limoges) »
            m = re.search(r"(?:\bà\s+|\()([A-ZÉÈ][\w'’ -]+?)\)?\s*$", title)
            city = m.group(1).strip() if m else ""
        lv = "ST" if re.search(r"\bstages?\b", norm(title)) else "NA"
        out.append({"s": iso(s), "e": iso(e), "lv": lv, "arms": arms,
                    "cats": cats_d or {k: "DH" for k in ("M13", "M17", "M20", "SEN", "VET")},
                    "title": title, "catText": cat_text if cats_d else "Toutes catégories",
                    "detail": "" if norm(city) in norm(title) else city, "warn": None, "src": "Agenda du site de la ligue",
                    "url": ev.get("url")})
    return out


def overlaps(a: dict, b: dict, slack: int = 0) -> bool:
    a_s, a_e = dt.date.fromisoformat(a["s"]), dt.date.fromisoformat(a["e"])
    b_s, b_e = dt.date.fromisoformat(b["s"]), dt.date.fromisoformat(b["e"])
    return a_s <= b_e + dt.timedelta(days=slack) and b_s <= a_e + dt.timedelta(days=slack)


def fr_range(a: str, b: str) -> str:
    da, db = dt.date.fromisoformat(a), dt.date.fromisoformat(b)
    if da == db:
        return f"{da.day} {MN[da.month - 1]}"
    left = f"{da.day}" if da.month == db.month else f"{da.day} {MN[da.month - 1]}"
    return f"{left}-{db.day} {MN[db.month - 1]}"


def event_kind(title: str) -> str | None:
    """Nature d’un événement, pour ne rapprocher que des choses comparables."""
    n = norm(title)
    if "championnat" in n:
        return "championnat"
    if re.search(r"\bstages?\b", n):
        return "stage"
    m = re.search(r"(epreuve|regionale)\s+(?:regionale\s+)?(\d)", n)
    return f"epreuve{m.group(2)}" if m else None


def merge_agenda(na_events: list[dict], agenda: list[dict]) -> list[dict]:
    """Ajoute les événements de l’agenda absents du calendrier NA, et
    signale ceux dont la date diffère."""
    def same_as(ev):
        return [n for n in na_events if n["lv"] == ev["lv"] and set(n["arms"]) & set(ev["arms"]) and overlaps(n, ev)]

    confirmed: set[int] = set()   # événements NA retrouvés à la même date dans l’agenda
    for ev in agenda:
        for n in same_as(ev):
            confirmed.add(id(n))
            n.setdefault("url", ev.get("url"))
            if n.pop("_year_fixed", False):
                n["warn"] = None  # date confirmée par l’agenda du site
    added = []
    for ev in agenda:
        if same_as(ev):
            continue
        # même arme, même lieu, même nature, autre date → probable changement de date
        city = norm(ev["detail"])
        moved = [n for n in na_events if id(n) not in confirmed and n["lv"] == ev["lv"]
                 and set(n["arms"]) & set(ev["arms"])
                 and city and city in norm(n["title"]) and overlaps(n, ev, slack=45)
                 and event_kind(n["title"]) and event_kind(n["title"]) == event_kind(ev["title"])]
        warn = None
        if moved:
            m = moved[0]
            warn = (f"{m['title']} : calendrier NA = {fr_range(m['s'], m['e'])}, "
                    f"agenda du site = {fr_range(ev['s'], ev['e'])}")
            m["warn"] = warn
            NOTES.append(warn + ". Les deux dates sont affichées.")
        added.append({**ev, "warn": warn})
        log(f"[agenda]   + {ev['s']} {'/'.join(ev['arms'])} {ev['title']}")
    log(f"[agenda] {len(added)} événement(s) ajouté(s) depuis l’agenda du site")
    return added


# --------------------------------------------------------------------------
# Calendrier du comité départemental de Gironde (cdescrime33.fr)
# --------------------------------------------------------------------------

DATE_RE = re.compile(
    r"\b(\d{1,2})(?:er)?(?:\s*(?:ou|et|/|-|au)\s*(\d{1,2})(?:er)?)?\s+("
    + "|".join(sorted(MONTHS_FR, key=len, reverse=True)) + r")\.?(?:\s+(20\d\d))?", re.I)


def html_to_lines(body: str) -> list[str]:
    body = re.sub(r"(?is)<(script|style|noscript|svg|nav|header|footer)\b.*?</\1>", " ", body)
    body = re.sub(r"(?i)<br\s*/?>|</?(p|div|h[1-6]|li|ul|ol|tr|td|th|section|article|figure|figcaption|span class=\"elementor-heading-title)[^>]*>", "\n", body)
    body = re.sub(r"<[^>]+>", " ", body)
    lines = [clean(html.unescape(x)) for x in body.split("\n")]
    return [x for x in lines if x]


def strip_accents_keep(s: str) -> str:
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode()


def parse_gironde_lines(lines: list[str], start_year: int, url: str) -> list[dict]:
    events: "OrderedDict[tuple, dict]" = OrderedDict()
    weapon = None
    for i, line in enumerate(lines):
        n = norm(line)
        m = DATE_RE.search(strip_accents_keep(line).lower()) if len(line) < 160 else None
        if not m:
            # titre de section : « Rencontres départementales d'épée »
            if len(line) < 90 and not re.search(r"inscri|formulaire|feuille|google", n):
                found = [k for w, k in WEAPON_WORDS.items() if re.search(rf"\b{w}\b", n)]
                if len(found) == 1 and "laser" not in n:
                    weapon = found[0]
                elif "laser" in n:
                    weapon = "LASER"
            continue
        if not weapon:
            continue
        d1, d2, mon, year = m.group(1), m.group(2), MONTHS_FR[m.group(3).lower()], m.group(4)
        y = int(year) if year else season_year_for(mon, start_year)
        try:
            s = dt.date(y, mon, int(d1))
        except ValueError:
            continue
        e = s
        warn = None
        sep = norm(line[m.start():m.end()])
        if d2 and re.search(r"\bou\b", sep):
            warn = f"date à confirmer ({clean(line[m.start():m.end()])})"
        elif d2:
            try:
                e = dt.date(y, mon, int(d2))
            except ValueError:
                pass
        if "confirmer" in n and not warn:
            warn = "date à confirmer"
        # lieu : reste de la ligne, sinon la ligne suivante si elle ressemble à un lieu
        rest = clean(re.sub(r"(?i)\(?date à confirmer\)?", "", line[m.end():])).strip(" :–-—,.")
        if not rest and i + 1 < len(lines):
            nxt = lines[i + 1]
            nn = norm(nxt)
            if (len(nxt) < 50 and not DATE_RE.search(strip_accents_keep(nxt).lower())
                    and not re.search(r"rencontre|inscri|feuille|formulaire|confirmer|epee|fleuret|sabre|categorie", nn)):
                rest = nxt.strip(" :–-—,.")
        place = re.sub(r"\s*[—–-]\s*", " – ", rest) if rest else ""
        if "confirmer" in norm(place):
            place = ""
        wname = {"EPEE": "épée", "SABRE": "sabre", "FLEURET": "fleuret", "LASER": "sabre laser"}[weapon]
        key = (iso(s), weapon)
        if key in events:
            continue  # Elementor duplique parfois le contenu (mobile / bureau)
        events[key] = {
            "s": iso(s), "e": iso(e), "lv": "DEP", "arms": [weapon],
            "cats": {k: "DH" for k in ("M9", "M11", "M13", "M17", "M20", "SEN", "VET")},
            "title": f"Rencontre départementale {wname}" + (f" – {place}" if place else ""),
            "catText": "M9 à Vétérans", "detail": "", "warn": (f"Gironde, {wname} : {warn}" if warn else None),
            "src": "Comité départemental de Gironde", "url": url,
        }
    return list(events.values())


def fetch_gironde(cfg: dict, start_year: int, offline: bool) -> list[dict]:
    url = cfg.get("gironde_page")
    if not url:
        return []
    cache = SOURCES / "calendrier-gironde.html"
    body = None
    if not offline:
        try:
            body = http_get(url).decode("utf-8", "replace")
        except Exception as e:  # noqa: BLE001
            log(f"[gironde] page inaccessible ({e}) ; essai par l’API WordPress publique")
            slug = url.rstrip("/").rsplit("/", 1)[-1]
            api = urljoin(url, f"/wp-json/wp/v2/pages?slug={slug}&_fields=content")
            try:
                j = json.loads(http_get(api))
                body = j[0]["content"]["rendered"] if j else None
            except Exception as e2:  # noqa: BLE001
                log(f"[gironde] API inaccessible ({e2}) ; copie locale utilisée si présente")
        if body:
            cache.write_text(body, encoding="utf-8")
    if body is None and cache.exists():
        body = cache.read_text(encoding="utf-8")
    if body is None:
        log("[gironde] aucune source")
        return []
    evs = parse_gironde_lines(html_to_lines(body), start_year, url)
    for ev in evs:
        log(f"[gironde]   {ev['s']} {ev['arms'][0]} {ev['title']}")
    log(f"[gironde] {len(evs)} rencontres départementales")
    if not evs:
        NOTES.append("Calendrier de Gironde : aucune date reconnue sur cdescrime33.fr, la mise en page a peut-être changé.")
    return evs


# --------------------------------------------------------------------------
# Événements manuels (Gironde, corrections)
# --------------------------------------------------------------------------

def load_manual(path: Path) -> tuple[list[dict], list[dict]]:
    if not path.exists():
        return [], []
    j = json.loads(path.read_text(encoding="utf-8"))
    evs = []
    for ev in j.get("evenements", []):
        try:
            s = dt.date.fromisoformat(ev["debut"])
            e = dt.date.fromisoformat(ev.get("fin") or ev["debut"])
        except Exception as ex:  # noqa: BLE001
            log(f"[manuel] événement ignoré {ev} : {ex}")
            continue
        lv = {"gironde": "DEP", "departement": "DEP", "region": "NA", "france": "FR", "stage": "ST"}.get(
            norm(ev.get("niveau", "gironde")), "DEP")
        arms = [WEAPON_WORDS.get(norm(a), "ANIM" if "anim" in norm(a) else None) for a in ev.get("armes", [])]
        arms = [a for a in arms if a] or ["EPEE", "SABRE", "FLEURET"]
        cats_d, cat_text = parse_categories(" ".join(ev.get("categories", [])) or "M13 à Vétérans")
        sexe = {"dames": "D", "hommes": "H"}.get(norm(ev.get("sexe", "")), "DH")
        cats_d = {k: sexe for k in cats_d}
        title = ev["titre"] + (f" – {ev['lieu']}" if ev.get("lieu") and norm(ev["lieu"]) not in norm(ev["titre"]) else "")
        evs.append({"s": iso(s), "e": iso(e), "lv": lv, "arms": arms, "cats": cats_d,
                    "title": title, "catText": cat_text, "detail": ev.get("detail", ""),
                    "warn": ev.get("a_confirmer") or None, "src": ev.get("source", "Saisie manuelle"),
                    "url": ev.get("lien")})
    return evs, j.get("masquer", [])


def apply_hide(events: list[dict], rules: list[dict]) -> list[dict]:
    def match(ev, r):
        if r.get("debut") and ev["s"] != r["debut"]:
            return False
        if r.get("titre_contient") and norm(r["titre_contient"]) not in norm(ev["title"]):
            return False
        return bool(r.get("debut") or r.get("titre_contient"))
    return [ev for ev in events if not any(match(ev, r) for r in rules)]


# --------------------------------------------------------------------------
# Jours fériés
# --------------------------------------------------------------------------

def easter(y: int) -> dt.date:
    a = y % 19; b = y // 100; c = y % 100; d = b // 4; e = b % 4
    f = (b + 8) // 25; g = (b - f + 1) // 3; h = (19 * a + b - d - g + 15) % 30
    i = c // 4; k = c % 4; l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31; day = (h + l - 7 * m + 114) % 31 + 1
    return dt.date(y, month, day)


def holidays(start_year: int) -> dict[str, str]:
    out = {}
    for y in (start_year, start_year + 1):
        e = easter(y)
        for d, name in [
            (dt.date(y, 1, 1), "Jour de l’An"), (e + dt.timedelta(days=1), "Lundi de Pâques"),
            (dt.date(y, 5, 1), "Fête du Travail"), (dt.date(y, 5, 8), "Victoire 1945"),
            (e + dt.timedelta(days=39), "Ascension"), (e + dt.timedelta(days=50), "Lundi de Pentecôte"),
            (dt.date(y, 7, 14), "Fête nationale"), (dt.date(y, 8, 15), "Assomption"),
            (dt.date(y, 11, 1), "Toussaint"), (dt.date(y, 11, 11), "Armistice"), (dt.date(y, 12, 25), "Noël"),
        ]:
            out[iso(d)] = name
    return out


# --------------------------------------------------------------------------
# Assemblage
# --------------------------------------------------------------------------

def detect_season(cfg: dict, today: dt.date) -> int:
    if cfg.get("saison"):
        return int(str(cfg["saison"])[:4])
    return today.year if today.month >= 8 else today.year - 1


def build(offline: bool, out: Path) -> None:
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    SOURCES.mkdir(parents=True, exist_ok=True)
    today = dt.date.today()
    start_year = detect_season(cfg, today)
    season_label = f"{start_year}-{start_year + 1}"

    books = load_excels(cfg, offline)
    events: list[dict] = []
    vac: dict[str, str] = {}
    region_cells: list[dict] = []
    src_names = []
    if "ffe" in books:
        ffe, vac, region_cells = parse_ffe(books["ffe"][0], start_year)
        events += ffe
        src_names.append(books["ffe"][1])
    if "na" in books:
        na = parse_na(books["na"][0], start_year)
        events += reconcile_with_grid(na, region_cells)
        src_names.append(books["na"][1])
    season_bounds = (dt.date(start_year, 8, 1), dt.date(start_year + 1, 7, 31))
    agenda = fetch_tec_events(cfg, season_bounds, offline)
    events += merge_agenda(events, agenda)
    events += fetch_gironde(cfg, start_year, offline)
    manual, hide = load_manual(ROOT / "data" / "evenements_manuels.json")
    events += manual
    events = apply_hide(events, hide)
    events = [e for e in events if season_bounds[0] <= dt.date.fromisoformat(e["s"]) <= season_bounds[1]]
    if not events:
        raise SystemExit("Aucun événement : génération annulée pour ne pas publier une page vide.")

    events.sort(key=lambda e: (e["s"], e["lv"], e["title"]))
    for e in events:
        if e.pop("_year_fixed", False) and e["warn"]:
            NOTES.append(e["warn"] + " (non confirmé par l’agenda du site).")
    for i, e in enumerate(events):
        e["id"] = i
        e.setdefault("url", None)

    first = min(dt.date.fromisoformat(e["s"]) for e in events)
    last = max(dt.date.fromisoformat(e["s"]) for e in events)
    months = []
    y, m = first.year, first.month
    while (y, m) <= (last.year, last.month):
        months.append([y, m])
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)

    data = {"season": season_label, "months": months, "events": events, "vac": vac,
            "hol": holidays(start_year), "generated": today.isoformat()}

    period = f"de {MN[first.month - 1]} {first.year} à {MN[last.month - 1]} {last.year}"
    notes_html = render_notes(cfg, src_names)
    now = dt.datetime.now(dt.timezone.utc).astimezone()
    footer = (f"Mis à jour automatiquement le {now.day} {MN[now.month - 1]} {now.year}. "
              f"Sources : <a href=\"{html.escape(cfg['pages'][0])}\">site de la ligue Nouvelle-Aquitaine</a> "
              f"(grille FFE, calendrier régional, agenda) et <a href=\"{html.escape(cfg.get('gironde_page') or '')}\">comité de Gironde</a>. "
              f"Lieux NA : propositions du calendrier régional, susceptibles d’évoluer.")

    tpl = TEMPLATE.read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    page = (tpl.replace("__DATA__", payload).replace("__SEASON__", season_label)
            .replace("__PERIOD__", period).replace("__NOTES__", notes_html).replace("__FOOTER__", footer))
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(page, encoding="utf-8")
    (out.parent / "evenements.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    counts = {}
    for e in events:
        counts[e["lv"]] = counts.get(e["lv"], 0) + 1
    log(f"[ok] {len(events)} événements {counts} → {out}")


def render_notes(cfg: dict, src_names: list[str]) -> str:
    esc = html.escape
    li = lambda items: "\n".join(f"      <li>{x}</li>" for x in items)  # noqa: E731
    notes = list(OrderedDict.fromkeys(NOTES))
    parts = [
        "    <h3>Ce que contient ce calendrier</h3>\n    <ul>",
        li([
            "<b>Sources :</b> " + ", ".join(f"« {esc(n)} »" for n in src_names) +
            f", l’<a href=\"{esc(cfg['pages'][0])}\">agenda du site de la ligue</a> et le <a href=\"{esc(cfg.get('gironde_page') or '')}\">calendrier du comité de Gironde</a>. La page est régénérée chaque semaine.",
            "<b>Épreuves NA :</b> dates, catégories, format et lieu proposé viennent du calendrier régional. Les cases régionales de la grille FFE (« ER », « Formation »…) ne sont pas reprises pour éviter les doublons.",
            "<b>Sabre laser :</b> épreuves du calendrier NA et de l’agenda du site, avec leur propre bouton dans « Armes ».",
            "<b>Non repris :</b> escrime artistique, para-escrime et circuits internationaux (FIE, EFC, championnats du monde et d’Europe).",
            "<b>Filtres :</b> « M13 / M15 » et « M17 / M20 » regroupent deux catégories. Les événements réservés à un sexe sont suivis de (D) ou (H). Vos filtres sont mémorisés dans votre navigateur.",
        ]),
        "    </ul>",
        "    <h3>Abréviations de la grille FFE</h3>",
        "    <p>NA : Nouvelle-Aquitaine · EQ : équipes · 1/2 EQ : demi-finale par équipes · H2036 : filière jeunes (1/8 et 1/4 de finale) · FRA : championnat de France · FDJ : Fête des jeunes · CF : coupe de France. Les autres libellés sont repris tels quels.</p>",
    ]
    if notes:
        parts += ["    <h3>Points à vérifier</h3>\n    <ul>", li([esc(n) for n in notes]), "    </ul>"]
    return "\n".join(parts)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--offline", action="store_true", help="n’utilise que les copies locales de data/sources")
    ap.add_argument("--out", default=str(ROOT / "dist" / "index.html"))
    a = ap.parse_args()
    build(a.offline, Path(a.out))


if __name__ == "__main__":
    main()
