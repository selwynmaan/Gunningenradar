#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gunningenradar — HVS
====================
Haalt gegunde opdrachten (AGO) op via de openbare TenderNed publicatie-webservice,
zoekt per gunning uit welke partij de opdracht uitvoert, en schrijft een
zelfstandig HTML-dashboard weg.

Gebruik:
    python tenderned_gunningen.py                 # normale run, 90 dagen terug
    python tenderned_gunningen.py --dagen 365     # heel jaar, voor een doelgroeplijst
    python tenderned_gunningen.py --alles         # ook lage scores
    python tenderned_gunningen.py --demo          # testrun zonder internet
    python tenderned_gunningen.py --velden        # toon veldnamen van de API

Nodig:   pip install requests
Optie:   pip install pypdf     (extra route om winnaars uit de PDF te halen)

Heb je een TenderNed API-account (gratis aan te vragen via
functioneelbeheer@tenderned.nl), zet dan deze omgevingsvariabelen:
    TENDERNED_API_USERNAME
    TENDERNED_API_PASSWORD
Daarmee wordt de winnaar veel vaker gevonden.
"""

import argparse
import datetime as dt
import html
import io
import json
import os
import re
import sys
import xml.etree.ElementTree as ET

try:
    import requests
except ImportError:
    requests = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

# ---------------------------------------------------------------------------
# INSTELLINGEN
# ---------------------------------------------------------------------------

DAGEN_TERUG = 90          # gunningen komen minder vaak voorbij dan aankondigingen
MAX_PAGINAS = 120
PAGINA_GROOTTE = 100
MIN_SCORE = 3             # minimale relevantiescore
MAX_VERRIJKING = 150      # hoeveel gunningen per run maximaal uitgezocht worden

BASIS_URL = "https://www.tenderned.nl/papi/tenderned-rs-tns/v2/publicaties"
PUBLICATIE_URL = "https://www.tenderned.nl/aankondigingen/overzicht/{id}"

# Trefwoorden met gewicht
TREFWOORDEN = {
    # luchtverdeling en klimaat
    "inductie-unit": 6, "inductieunit": 6, "klimaatplafond": 6, "koelplafond": 6,
    "luchtbehandeling": 5, "luchtbehandelingskast": 5, "luchtverdeling": 5,
    "ventilatie": 4, "klimaatinstallatie": 5, "klimaatinstallaties": 5,
    "klimaatbeheersing": 4, "airconditioning": 4, "hvac": 5,
    "brandklep": 4, "brandkleppen": 4, "luchtroosters": 5, "diffusor": 5,
    "binnenklimaat": 4, "luchtkwaliteit": 3, "koeling": 3, "verwarming": 2,
    # uitvoering
    "werktuigbouwkundige installatie": 5, "werktuigbouwkundig": 5,
    "w-installatie": 5, "w-installaties": 5, "e- en w-installaties": 5,
    "installatietechniek": 4, "technische installaties": 4, "installatiewerk": 4,
    # advies (W-adviesbureaus)
    "installatieadvies": 6, "adviesdiensten installaties": 6,
    "raadgevend ingenieur": 5, "ingenieursdiensten": 4, "advieswerkzaamheden": 4,
    "technisch adviseur": 4, "adviseur installaties": 6, "bouwfysica": 4,
    "ontwerpwerkzaamheden installaties": 5, "engineering": 3,
    "bestek": 3, "voorlopig ontwerp": 3, "definitief ontwerp": 3,
    "haalbaarheidsstudie": 2, "directievoering": 3, "bouwmanagement": 3,
    # gebouwtypes binnen jouw scope
    "kantoor": 3, "kantoorgebouw": 4, "kantoorpand": 4,
    "onderwijs": 3, "school": 3, "scholen": 3, "schoolgebouw": 4,
    "universiteit": 3, "hogeschool": 3, "onderwijshuisvesting": 4,
    "gemeentehuis": 3, "stadhuis": 3, "rijksvastgoed": 4,
    "frisse scholen": 5, "renovatie": 2, "verduurzaming": 2, "nieuwbouw": 2,
}

CPV_PREFIXEN = {
    "45331": 6,   # installatie verwarming, ventilatie, airco
    "45300": 4,   # installatiewerkzaamheden in de bouw
    "45350": 4,   # werktuigbouwkundige installaties
    "39717": 5,   # ventilatoren en airconditioning
    "42500": 4,   # koel- en ventilatie-uitrusting
    "42512": 5,   # airconditioninginstallaties
    "71321": 6,   # ontwerp mechanische en elektrische installaties
    "71315": 5,   # bouwtechnische diensten
    "71300": 4,   # ingenieursdiensten
    "71310": 4,   # advies bouwkunde en techniek
    "71240": 3,   # architectuur, bouwkunde en planning
    "45214": 2,   # onderwijsgebouwen
    "45213": 2,   # kantoorgebouwen
    "45215": 2,   # gezondheidszorggebouwen
}

SEGMENTEN = {
    "Onderwijs": ["onderwijs", "school", "scholen", "universiteit", "hogeschool",
                  "campus", "mbo", "hbo", "schoolgebouw", "45214"],
    "Kantoor": ["kantoor", "kantoorgebouw", "kantoorpand", "45213"],
    "Overheid": ["gemeente", "provincie", "rijksvastgoed", "ministerie",
                 "stadhuis", "gemeentehuis", "rijkswaterstaat", "waterschap"],
    "Zorg": ["ziekenhuis", "zorg", "operatiekamer", "cleanroom", "laboratorium",
             "45215"],
}

# Soort partij, afgeleid uit de bedrijfsnaam
PARTIJSOORTEN = [
    ("Adviesbureau", ["advies", "adviseurs", "ingenieur", "engineering", "consult",
                      "raadgevend", "architect", "bouwfysica", "adviesgroep",
                      "royal haskoning", "arcadis", "sweco", "deerns", "nelissen",
                      "valstar", "huygen", "dwa", "techniplan", "peutz"]),
    ("Installateur", ["installatie", "installaties", "techniek", "technieken",
                      "technisch", "klimaat", "koeltechniek", "installatiebedrijf",
                      "unica", "kuijpers", "equans", "engie", "croonwolter",
                      "spie", "hellebrekers", "wolter", "breman", "linthorst"]),
    ("Aannemer", ["bouw", "aannem", "bouwbedrijf", "bouwgroep", "bam", "heijmans",
                  "vandervalk", "dura", "ballast", "van wijnen", "pellikaan"]),
]

TYPE_NAMEN = {
    "AGO": "Gegunde opdracht",
    "WNO": "Wijziging opdracht",
    "AAO": "Aankondiging opdracht",
    "VOP": "Vooraankondiging",
    "MAC": "Marktconsultatie",
    "REC": "Rectificatie",
}

MAP = os.path.dirname(os.path.abspath(__file__))
UITVOER = os.path.join(MAP, "gunningen-radar.html")
CSV_UIT = os.path.join(MAP, "gunningen.csv")
GEZIEN = os.path.join(MAP, "gezien-gunningen.json")
CACHE = os.path.join(MAP, "winnaars-cache.json")
RUW = os.path.join(MAP, "laatste-ruwe-publicatie.json")

SESSIE = None


# ---------------------------------------------------------------------------
# Ophalen van de lijst
# ---------------------------------------------------------------------------

def sessie():
    global SESSIE
    if SESSIE is None:
        SESSIE = requests.Session()
        SESSIE.headers.update({"Accept": "application/json, application/xml",
                               "User-Agent": "HVS-Gunningenradar/1.0"})
    return SESSIE


def haal_gunningen(dagen, toon_velden=False):
    if requests is None:
        stop("De module 'requests' ontbreekt. Installeer met: pip install requests")

    vandaag = dt.date.today()
    vanaf = vandaag - dt.timedelta(days=dagen)
    alles, pagina = [], 0

    while pagina < MAX_PAGINAS:
        params = {
            "page": pagina,
            "size": PAGINA_GROOTTE,
            "publicatieType": "AGO",
            "publicatieDatumVanaf": vanaf.isoformat(),
            "publicatieDatumTot": vandaag.isoformat(),
        }
        try:
            r = sessie().get(BASIS_URL, params=params, timeout=45)
            r.raise_for_status()
            data = r.json()
        except Exception as fout:
            if pagina == 0:
                stop("TenderNed is niet bereikbaar: %s" % fout)
            print("  pagina %d overgeslagen (%s)" % (pagina, fout))
            break

        rijen = data.get("contents") or data.get("content") or []
        if not rijen:
            break

        if pagina == 0:
            with open(RUW, "w", encoding="utf-8") as f:
                json.dump(rijen[0], f, ensure_ascii=False, indent=2)
            if toon_velden:
                print("Beschikbare velden in een publicatie:")
                for sleutel, waarde in rijen[0].items():
                    print("   %-38s %s" % (sleutel, str(waarde)[:70]))
                sys.exit(0)

        alles.extend(rijen)
        print("  pagina %d: %d publicaties (totaal %d)" % (pagina, len(rijen), len(alles)))

        if data.get("last") is True or len(rijen) < PAGINA_GROOTTE:
            break
        pagina += 1

    return alles


def stop(bericht):
    print("\n  " + bericht + "\n")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Normaliseren
# ---------------------------------------------------------------------------

def kies(rij, *namen):
    for naam in namen:
        if naam in rij and rij[naam] not in (None, "", [], {}):
            return rij[naam]
    return None


def tekst(waarde):
    if waarde is None:
        return ""
    if isinstance(waarde, dict):
        for sleutel in ("omschrijving", "naam", "waarde", "code", "label", "name"):
            if waarde.get(sleutel):
                return str(waarde[sleutel])
        return " ".join(str(v) for v in waarde.values() if isinstance(v, str))
    if isinstance(waarde, list):
        return ", ".join(tekst(v) for v in waarde)
    return str(waarde)


def datum(waarde):
    ruw = tekst(waarde)[:10]
    return ruw if re.match(r"^\d{4}-\d{2}-\d{2}$", ruw) else ""


def normaliseer(rij):
    pub_id = kies(rij, "publicatieId", "id", "publicatieID")
    cpv_ruw = kies(rij, "cpvCodes", "cpvCode", "hoofdCpvCode", "cpv") or []
    if not isinstance(cpv_ruw, list):
        cpv_ruw = [cpv_ruw]

    type_ruw = kies(rij, "publicatieType", "typePublicatie", "type")
    type_code = ""
    if isinstance(type_ruw, dict):
        type_code = tekst(type_ruw.get("code") or "")
    elif isinstance(type_ruw, str) and len(type_ruw) <= 4:
        type_code = type_ruw

    return {
        "id": tekst(pub_id),
        "titel": tekst(kies(rij, "aanbestedingNaam", "publicatieNaam", "opdrachtNaam",
                            "naam", "titel", "title")) or "Zonder titel",
        "dienst": tekst(kies(rij, "aanbestedendeDienstNaam", "aanbestedendeDienst",
                             "organisatieNaam", "opdrachtgever")),
        "omschrijving": tekst(kies(rij, "opdrachtBeschrijving", "korteBeschrijving",
                                   "beschrijving", "omschrijving"))[:600],
        "publicatie": datum(kies(rij, "publicatieDatum", "datumPublicatie", "publicatiedatum")),
        "type": TYPE_NAMEN.get(type_code.upper()) or tekst(type_ruw) or "Publicatie",
        "typecode": type_code.upper(),
        "procedure": tekst(kies(rij, "procedureNaam", "procedure", "procedureType")),
        "plaats": tekst(kies(rij, "plaats", "vestigingsplaats", "plaatsUitvoering")),
        "cpv": [tekst(c) for c in cpv_ruw],
        "link": PUBLICATIE_URL.format(id=tekst(pub_id)) if pub_id else "",
        "winnaars": [],
        "bron": "",
        "waarde": "",
    }


# ---------------------------------------------------------------------------
# Winnaar uitzoeken — drie routes, van goedkoop naar duur
# ---------------------------------------------------------------------------

WINNAAR_SLEUTELS = ("gegundeaan", "opdrachtnemer", "ondernemingnaam", "onderneming",
                    "winnaar", "contractant", "gegundeondernemingen", "inschrijver",
                    "leverancier", "contractor")


def zoek_diep(obj, sleutels, gevonden=None):
    """Doorzoekt geneste JSON op sleutels die op een winnaar wijzen."""
    if gevonden is None:
        gevonden = []
    if isinstance(obj, dict):
        for sleutel, waarde in obj.items():
            plat = sleutel.lower().replace("_", "")
            if any(s in plat for s in sleutels):
                waarde_tekst = tekst(waarde).strip()
                if 2 < len(waarde_tekst) < 160:
                    gevonden.append(waarde_tekst)
            zoek_diep(waarde, sleutels, gevonden)
    elif isinstance(obj, list):
        for item in obj:
            zoek_diep(item, sleutels, gevonden)
    return gevonden


def strip_ns(tag):
    return tag.split("}")[-1] if "}" in tag else tag


def uit_xml(xml_tekst, buyer):
    """Haalt winnaar en waarde uit TED SF-XML of eForms-XML."""
    try:
        wortel = ET.fromstring(xml_tekst.encode("utf-8") if isinstance(xml_tekst, str)
                               else xml_tekst)
    except ET.ParseError:
        return [], ""

    knopen = list(wortel.iter())
    namen, waarde = [], ""

    # Route A: klassiek TED-formulier (SF03)
    for knoop in knopen:
        if strip_ns(knoop.tag) == "ADDRESS_CONTRACTOR":
            for kind in knoop.iter():
                if strip_ns(kind.tag) == "OFFICIALNAME" and (kind.text or "").strip():
                    namen.append(kind.text.strip())
    for knoop in knopen:
        if strip_ns(knoop.tag) in ("VAL_TOTAL", "VAL_ESTIMATED_TOTAL") and knoop.text:
            waarde = waarde or knoop.text.strip()

    # Route B: eForms — organisaties opzoeken en de buyer eruit filteren
    if not namen:
        organisaties = []
        for knoop in knopen:
            if strip_ns(knoop.tag) in ("Organization", "Company"):
                for kind in knoop.iter():
                    if strip_ns(kind.tag) in ("CompanyName", "Name", "RegistrationName"):
                        waarde_tekst = (kind.text or "").strip()
                        if waarde_tekst:
                            organisaties.append(waarde_tekst)
                            break
        schoon_buyer = normaliseer_naam(buyer)
        for naam in organisaties:
            if normaliseer_naam(naam) != schoon_buyer and len(naam) > 2:
                namen.append(naam)
        if not waarde:
            for knoop in knopen:
                if strip_ns(knoop.tag) in ("PayableAmount", "TaxExclusiveAmount") and knoop.text:
                    waarde = knoop.text.strip()
                    break

    # dubbelen eruit, volgorde behouden
    uniek = []
    for naam in namen:
        if naam not in uniek:
            uniek.append(naam)
    return uniek[:4], waarde


PDF_MARKERS = ("naam en adres van de contractant", "contractant", "opdrachtnemer",
               "winnende inschrijver", "winnaar", "gegund aan", "de ondernemer")
PDF_NAAM = re.compile(r"(?:offici[eë]le\s+(?:benaming|naam)|naam\s+van\s+de\s+ondernemer|"
                      r"naam)\s*:?\s*(.{3,120})", re.IGNORECASE)


def uit_pdf(pdf_bytes, buyer):
    if PdfReader is None:
        return [], ""
    try:
        lezer = PdfReader(io.BytesIO(pdf_bytes))
        volledig = "\n".join((p.extract_text() or "") for p in lezer.pages[:12])
    except Exception:
        return [], ""

    laag = volledig.lower()
    start = -1
    for marker in PDF_MARKERS:
        positie = laag.find(marker)
        if positie != -1:
            start = positie
            break
    if start == -1:
        return [], ""

    staart = volledig[start:start + 2500]
    namen = []
    for treffer in PDF_NAAM.finditer(staart):
        kandidaat = treffer.group(1).strip().strip(":;-").strip()
        kandidaat = re.split(r"\s{3,}|\|", kandidaat)[0].strip()
        if (len(kandidaat) > 3 and normaliseer_naam(kandidaat) != normaliseer_naam(buyer)
                and kandidaat not in namen):
            namen.append(kandidaat)
        if len(namen) >= 3:
            break

    waarde = ""
    bedrag = re.search(r"(?:waarde|bedrag)[^\d]{0,40}([\d.,]{4,})", staart, re.IGNORECASE)
    if bedrag:
        waarde = bedrag.group(1)
    return namen, waarde


def normaliseer_naam(naam):
    schoon = re.sub(r"\b(b\.?v\.?|n\.?v\.?|v\.?o\.?f\.?|holding|nederland|group|groep)\b",
                    " ", (naam or "").lower())
    return re.sub(r"[^a-z0-9]+", "", schoon)


def verrijk(pub, cache, gebruikersnaam, wachtwoord):
    """Vult winnaars/waarde aan. Gebruikt de cache zodat elke run snel blijft."""
    if pub["id"] in cache:
        opgeslagen = cache[pub["id"]]
        pub["winnaars"] = opgeslagen.get("winnaars", [])
        pub["bron"] = opgeslagen.get("bron", "")
        pub["waarde"] = opgeslagen.get("waarde", "")
        return False

    basis = "%s/%s" % (BASIS_URL, pub["id"])
    namen, waarde, bron = [], "", ""

    # Route 1 — detail-JSON, openbaar
    try:
        r = sessie().get(basis, timeout=30)
        if r.ok:
            detail = r.json()
            kandidaten = [n for n in zoek_diep(detail, WINNAAR_SLEUTELS)
                          if normaliseer_naam(n) != normaliseer_naam(pub["dienst"])]
            if kandidaten:
                namen, bron = kandidaten[:3], "detail"
            if not pub["omschrijving"]:
                pub["omschrijving"] = tekst(kies(detail, "opdrachtBeschrijving",
                                                 "beschrijving"))[:600]
    except Exception:
        pass

    # Route 2 — eForms/TED XML
    if not namen:
        try:
            auth = (gebruikersnaam, wachtwoord) if gebruikersnaam else None
            r = sessie().get(basis + "/public-xml", auth=auth, timeout=40)
            if r.ok and r.text.strip().startswith("<"):
                namen, waarde = uit_xml(r.text, pub["dienst"])
                if namen:
                    bron = "XML"
        except Exception:
            pass

    # Route 3 — de openbare PDF van de aankondiging
    if not namen and PdfReader is not None:
        try:
            r = sessie().get(basis + "/pdf", timeout=60)
            if r.ok and r.content[:4] == b"%PDF":
                namen, pdf_waarde = uit_pdf(r.content, pub["dienst"])
                waarde = waarde or pdf_waarde
                if namen:
                    bron = "PDF"
        except Exception:
            pass

    pub["winnaars"] = namen
    pub["bron"] = bron
    pub["waarde"] = waarde
    cache[pub["id"]] = {"winnaars": namen, "bron": bron, "waarde": waarde}
    return True


def soort_partij(naam):
    laag = (naam or "").lower()
    for soort, sleutels in PARTIJSOORTEN:
        if any(s in laag for s in sleutels):
            return soort
    return "Overig"


# ---------------------------------------------------------------------------
# Scoren
# ---------------------------------------------------------------------------

def scoor(pub):
    hooiberg = " ".join([pub["titel"], pub["omschrijving"], pub["dienst"],
                         " ".join(pub["cpv"])]).lower()
    score, redenen = 0, []

    for woord, gewicht in TREFWOORDEN.items():
        if woord in hooiberg:
            score += gewicht
            redenen.append(woord)

    for code in pub["cpv"]:
        schoon = re.sub(r"\D", "", code)
        for prefix, gewicht in CPV_PREFIXEN.items():
            if schoon.startswith(prefix):
                score += gewicht
                redenen.append("cpv " + prefix)
                break

    segmenten = [s for s, sleutels in SEGMENTEN.items()
                 if any(k in hooiberg for k in sleutels)]

    pub["score"] = score
    pub["redenen"] = sorted(set(redenen))[:8]
    pub["segmenten"] = segmenten or ["Overig"]
    return pub


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------

SJABLOON = """<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Gunningenradar — HVS</title>
<style>
  :root {{
    --papier: #f1f3f2; --vlak: #ffffff; --inkt: #15201d; --grijs: #5d6b66;
    --lijn: #ccd6d1; --nieuw: #0e6a58; --rust: #3d5a80; --naam: #7a3b12;
  }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; background: var(--papier); color: var(--inkt);
    font: 15px/1.5 "Segoe UI", -apple-system, Roboto, Helvetica, Arial, sans-serif;
    -webkit-font-smoothing: antialiased; }}
  .omslag {{ max-width: 1080px; margin: 0 auto; padding: 32px 22px 80px; }}

  header {{ border-bottom: 2px solid var(--inkt); padding-bottom: 14px; }}
  h1 {{ font: 600 30px/1.15 Georgia, "Iowan Old Style", "Times New Roman", serif;
    margin: 0 0 6px; letter-spacing: -0.01em; }}
  .stand {{ color: var(--grijs); font-size: 14px; margin: 0; }}
  .stand b {{ color: var(--inkt); font-weight: 600; }}

  h2 {{ font: 600 17px/1.2 Georgia, serif; margin: 30px 0 10px; }}
  .ranglijst {{ background: var(--vlak); border: 1px solid var(--lijn);
    border-radius: 3px; padding: 4px 16px 10px; }}
  .ranglijst table {{ width: 100%; border-collapse: collapse; font-size: 13.5px; }}
  .ranglijst td {{ padding: 7px 0; border-bottom: 1px solid var(--lijn); }}
  .ranglijst tr:last-child td {{ border-bottom: 0; }}
  .ranglijst .partij {{ font-weight: 600; }}
  .ranglijst .soort {{ color: var(--grijs); }}
  .ranglijst .aantal {{ text-align: right; font-variant-numeric: tabular-nums;
    color: var(--grijs); white-space: nowrap; }}
  .staaf {{ display: inline-block; height: 7px; background: var(--rust);
    border-radius: 2px; vertical-align: middle; margin-right: 8px; }}

  .balk {{ display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin: 14px 0 6px; }}
  .knop {{ border: 1px solid var(--lijn); background: var(--vlak); color: var(--grijs);
    border-radius: 999px; padding: 5px 13px; font-size: 13px; cursor: pointer;
    font-family: inherit; }}
  .knop:hover {{ border-color: var(--grijs); color: var(--inkt); }}
  .knop[aria-pressed="true"] {{ background: var(--inkt); border-color: var(--inkt); color: #fff; }}
  .knop:focus-visible, #zoek:focus-visible {{ outline: 2px solid var(--rust); outline-offset: 2px; }}
  #zoek {{ flex: 1 1 180px; min-width: 160px; padding: 6px 12px; font-size: 13px;
    font-family: inherit; border: 1px solid var(--lijn); border-radius: 999px;
    background: var(--vlak); color: var(--inkt); }}
  .telling {{ color: var(--grijs); font-size: 13px; margin: 0 0 4px; }}

  .rij {{ background: var(--vlak); border: 1px solid var(--lijn); border-left-width: 3px;
    border-radius: 3px; padding: 14px 16px; margin-top: 10px; }}
  .rij[data-nieuw="1"] {{ border-left-color: var(--nieuw); }}
  .gegund {{ font: 600 17px/1.3 Georgia, serif; color: var(--naam); margin: 0 0 2px; }}
  .gegund .onbekend {{ color: var(--grijs); font-style: italic; font-weight: 400; font-size: 15px; }}
  .titel {{ font-size: 14.5px; margin: 0 0 3px; }}
  .titel a {{ color: inherit; text-decoration: none; border-bottom: 1px solid var(--lijn); }}
  .titel a:hover {{ border-color: var(--inkt); }}
  .dienst {{ color: var(--grijs); font-size: 13px; margin: 0 0 8px; }}
  .tekst {{ font-size: 13.5px; color: #3c4a45; margin: 0 0 9px; max-width: 72ch; }}

  .merken {{ display: flex; flex-wrap: wrap; gap: 6px; align-items: center; }}
  .merk {{ font-size: 11.5px; border: 1px solid var(--lijn); border-radius: 3px;
    padding: 2px 7px; color: var(--grijs); background: #fafbfa; }}
  .merk.nieuw {{ border-color: var(--nieuw); color: var(--nieuw); font-weight: 600; }}
  .merk.soort {{ border-color: var(--rust); color: var(--rust); }}
  .merk.score {{ margin-left: auto; font-variant-numeric: tabular-nums; }}

  .leeg {{ border: 1px dashed var(--lijn); border-radius: 3px; padding: 28px;
    text-align: center; color: var(--grijs); margin-top: 16px; background: var(--vlak); }}
  footer {{ margin-top: 34px; padding-top: 14px; border-top: 1px solid var(--lijn);
    color: var(--grijs); font-size: 12.5px; max-width: 72ch; }}
</style>
</head>
<body>
<div class="omslag">

<header>
  <h1>Gunningenradar</h1>
  <p class="stand">Bijgewerkt {bijgewerkt} · <b>{aantal}</b> relevante gunningen uit
     {totaal} van de afgelopen {dagen} dagen · bij <b>{metnaam}</b> is de uitvoerende
     partij gevonden · <b>{nieuw}</b> nieuw sinds de vorige run</p>
</header>

<h2>Wie wint er in jouw segmenten</h2>
<div class="ranglijst">
<table>
{ranglijst}
</table>
</div>

<h2>Gunningen</h2>
<div class="balk" role="group" aria-label="Filters">
  <input id="zoek" type="search" placeholder="Zoek op partij, opdrachtgever of project">
  <button class="knop" data-filter="nieuw" aria-pressed="false">Alleen nieuw</button>
  <button class="knop" data-filter="bekend" aria-pressed="false">Alleen met naam</button>
  {soortknoppen}
  {segmentknoppen}
</div>
<p class="telling" id="telling" aria-live="polite"></p>

<div id="lijst">
{rijen}
</div>
<div class="leeg" id="leeg" hidden>Niets binnen dit filter. Zet een filter uit of zoek breder.</div>

<footer>
  Bron: openbare publicatie-webservice van TenderNed, publicatietype “aankondiging
  van een gegunde opdracht”. De naam van de uitvoerende partij staat niet in de
  lijstdata zelf; die wordt per gunning uit het onderliggende XML- of PDF-document
  gehaald. Lukt dat niet, dan opent de link de publicatie op TenderNed. Onderhandse
  gunningen onder de drempel en private opdrachtgevers staan hier niet in.
  Zie ook gunningen.csv voor import in je CRM.
</footer>

</div>
<script>
const rijen = Array.from(document.querySelectorAll('.rij'));
const zoek = document.getElementById('zoek');
const telling = document.getElementById('telling');
const leeg = document.getElementById('leeg');
const actief = new Set();
const SOORTEN = ['Adviesbureau','Installateur','Aannemer','Overig'];

document.querySelectorAll('.knop').forEach(knop => {{
  knop.addEventListener('click', () => {{
    const naam = knop.dataset.filter;
    if (actief.has(naam)) {{ actief.delete(naam); knop.setAttribute('aria-pressed','false'); }}
    else {{ actief.add(naam); knop.setAttribute('aria-pressed','true'); }}
    ververs();
  }});
}});
zoek.addEventListener('input', ververs);

function ververs() {{
  const term = zoek.value.trim().toLowerCase();
  const gekozen = [...actief].filter(a => a !== 'nieuw' && a !== 'bekend');
  const soorten = gekozen.filter(s => SOORTEN.includes(s));
  const segmenten = gekozen.filter(s => !SOORTEN.includes(s));
  let zichtbaar = 0;
  rijen.forEach(rij => {{
    let toon = true;
    if (actief.has('nieuw') && rij.dataset.nieuw !== '1') toon = false;
    if (actief.has('bekend') && rij.dataset.bekend !== '1') toon = false;
    if (soorten.length && !soorten.some(s => rij.dataset.soorten.includes(s))) toon = false;
    if (segmenten.length && !segmenten.some(s => rij.dataset.segmenten.includes(s))) toon = false;
    if (term && !rij.dataset.zoek.includes(term)) toon = false;
    rij.hidden = !toon;
    if (toon) zichtbaar++;
  }});
  telling.textContent = zichtbaar + ' van ' + rijen.length + ' getoond';
  leeg.hidden = zichtbaar !== 0;
}}
ververs();
</script>
</body>
</html>
"""


def bouw_rij(pub, is_nieuw):
    soorten = [soort_partij(n) for n in pub["winnaars"]] or ["Overig"]

    if pub["winnaars"]:
        kop = " · ".join(html.escape(n) for n in pub["winnaars"])
    else:
        kop = '<span class="onbekend">Uitvoerende partij niet uit de publicatie te halen</span>'

    merken = []
    if is_nieuw:
        merken.append('<span class="merk nieuw">nieuw</span>')
    for soort in dict.fromkeys(soorten):
        merken.append('<span class="merk soort">%s</span>' % soort)
    for segment in pub["segmenten"]:
        merken.append('<span class="merk">%s</span>' % html.escape(segment))
    if pub["waarde"]:
        merken.append('<span class="merk">waarde %s</span>' % html.escape(pub["waarde"]))
    if pub["bron"]:
        merken.append('<span class="merk">naam uit %s</span>' % html.escape(pub["bron"]))
    for reden in pub["redenen"][:4]:
        merken.append('<span class="merk">%s</span>' % html.escape(reden))
    merken.append('<span class="merk score">score %d</span>' % pub["score"])

    titel = html.escape(pub["titel"])
    if pub["link"]:
        titel = '<a href="%s" target="_blank" rel="noopener">%s</a>' % (pub["link"], titel)

    zoekveld = html.escape(" ".join([pub["titel"], pub["dienst"], pub["omschrijving"],
                                     " ".join(pub["winnaars"])]).lower())

    return """<article class="rij" data-nieuw="{nieuw}" data-bekend="{bekend}"
         data-soorten="{soorten}" data-segmenten="{segmenten}" data-zoek="{zoek}">
  <p class="gegund">{kop}</p>
  <p class="titel">{titel}</p>
  <p class="dienst">Opdrachtgever: {dienst}{plaats} · gegund gepubliceerd {publicatie}{procedure}</p>
  {tekst}
  <div class="merken">{merken}</div>
</article>""".format(
        nieuw="1" if is_nieuw else "0",
        bekend="1" if pub["winnaars"] else "0",
        soorten=html.escape("|".join(dict.fromkeys(soorten))),
        segmenten=html.escape("|".join(pub["segmenten"])),
        zoek=zoekveld,
        kop=kop,
        titel=titel,
        dienst=html.escape(pub["dienst"] or "onbekend"),
        plaats=(" · " + html.escape(pub["plaats"])) if pub["plaats"] else "",
        publicatie=html.escape(pub["publicatie"] or "onbekend"),
        procedure=(" · " + html.escape(pub["procedure"])) if pub["procedure"] else "",
        tekst=('<p class="tekst">%s</p>' % html.escape(pub["omschrijving"][:300])
               if pub["omschrijving"] else ""),
        merken="".join(merken),
    )


def bouw_ranglijst(publicaties):
    tellers = {}
    for pub in publicaties:
        for naam in pub["winnaars"]:
            sleutel = normaliseer_naam(naam)
            if not sleutel:
                continue
            regel = tellers.setdefault(sleutel, {"naam": naam, "aantal": 0})
            regel["aantal"] += 1
    if not tellers:
        return ('<tr><td class="soort">Nog geen namen gevonden. '
                'Zie de toelichting onderaan.</td></tr>')

    top = sorted(tellers.values(), key=lambda r: -r["aantal"])[:12]
    hoogste = top[0]["aantal"] or 1
    regels = []
    for regel in top:
        breedte = max(6, int(120 * regel["aantal"] / hoogste))
        regels.append(
            '<tr><td class="partij">%s</td><td class="soort">%s</td>'
            '<td class="aantal"><span class="staaf" style="width:%dpx"></span>%d</td></tr>'
            % (html.escape(regel["naam"]), soort_partij(regel["naam"]), breedte, regel["aantal"]))
    return "\n".join(regels)


def schrijf_dashboard(publicaties, totaal, dagen, nieuwe_ids):
    publicaties.sort(key=lambda p: (p["publicatie"] or "", p["score"]), reverse=True)
    rijen = [bouw_rij(p, p["id"] in nieuwe_ids) for p in publicaties]

    soortknoppen = "".join(
        '<button class="knop" data-filter="%s" aria-pressed="false">%s</button>' % (s, s)
        for s in ("Adviesbureau", "Installateur", "Aannemer"))
    segmentknoppen = "".join(
        '<button class="knop" data-filter="%s" aria-pressed="false">%s</button>' % (s, s)
        for s in list(SEGMENTEN))

    pagina = SJABLOON.format(
        bijgewerkt=dt.datetime.now().strftime("%d-%m-%Y %H:%M"),
        aantal=len(publicaties),
        totaal=totaal,
        dagen=dagen,
        metnaam=sum(1 for p in publicaties if p["winnaars"]),
        nieuw=len(nieuwe_ids),
        ranglijst=bouw_ranglijst(publicaties),
        soortknoppen=soortknoppen,
        segmentknoppen=segmentknoppen,
        rijen="\n".join(rijen),
    )
    with open(UITVOER, "w", encoding="utf-8") as f:
        f.write(pagina)
    return UITVOER


def schrijf_csv(publicaties):
    import csv
    with open(CSV_UIT, "w", encoding="utf-8-sig", newline="") as f:
        schrijver = csv.writer(f, delimiter=";")
        schrijver.writerow(["Uitvoerende partij", "Soort", "Opdrachtgever", "Project",
                            "Segment", "Publicatiedatum", "Waarde", "Score", "Link"])
        for pub in publicaties:
            for naam in (pub["winnaars"] or [""]):
                schrijver.writerow([naam, soort_partij(naam) if naam else "",
                                    pub["dienst"], pub["titel"],
                                    "/".join(pub["segmenten"]), pub["publicatie"],
                                    pub["waarde"], pub["score"], pub["link"]])
    return CSV_UIT


# ---------------------------------------------------------------------------
# Geheugen
# ---------------------------------------------------------------------------

def lees_json(pad, standaard):
    if not os.path.exists(pad):
        return standaard
    try:
        with open(pad, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return standaard


def schrijf_json(pad, data):
    with open(pad, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


# ---------------------------------------------------------------------------

VOORBEELD = [
    {"publicatieId": 900101, "aanbestedingNaam": "Vervanging klimaatinstallatie kantoorgebouw Willemshof",
     "aanbestedendeDienstNaam": "Rijksvastgoedbedrijf", "plaats": "Den Haag",
     "opdrachtBeschrijving": "Vervanging van de luchtbehandeling en inductie-units in een bestaand kantoorgebouw.",
     "publicatieDatum": "2026-09-04", "publicatieType": {"code": "AGO"},
     "cpvCodes": ["45331000-6", "45213150-9"], "procedureNaam": "Europees openbaar",
     "_winnaars": ["Kuijpers Installaties B.V."], "_bron": "XML", "_waarde": "2.450.000"},
    {"publicatieId": 900102, "aanbestedingNaam": "Adviesdiensten installaties onderwijshuisvesting Campus Noord",
     "aanbestedendeDienstNaam": "Gemeente Zwolle", "plaats": "Zwolle",
     "opdrachtBeschrijving": "Advieswerkzaamheden W-installaties, van voorlopig ontwerp tot bestek, "
                             "voor een schoolgebouw conform Frisse Scholen klasse B.",
     "publicatieDatum": "2026-09-02", "publicatieType": {"code": "AGO"},
     "cpvCodes": ["71321000-4"], "procedureNaam": "Meervoudig onderhands",
     "_winnaars": ["Deerns Nederland B.V."], "_bron": "PDF", "_waarde": ""},
    {"publicatieId": 900103, "aanbestedingNaam": "Renovatie werktuigbouwkundige installaties gemeentehuis",
     "aanbestedendeDienstNaam": "Gemeente Apeldoorn", "plaats": "Apeldoorn",
     "opdrachtBeschrijving": "Renovatie van de W-installaties inclusief ventilatie en brandkleppen.",
     "publicatieDatum": "2026-08-28", "publicatieType": {"code": "AGO"},
     "cpvCodes": ["45350000-5"], "procedureNaam": "Nationaal openbaar",
     "_winnaars": [], "_bron": "", "_waarde": ""},
]


def main():
    p = argparse.ArgumentParser(description="Gunningenradar voor HVS")
    p.add_argument("--dagen", type=int, default=DAGEN_TERUG)
    p.add_argument("--alles", action="store_true", help="ook lage scores")
    p.add_argument("--demo", action="store_true", help="testrun met voorbeelddata")
    p.add_argument("--velden", action="store_true", help="toon veldnamen van de API")
    p.add_argument("--geen-verrijking", action="store_true",
                   help="sla het opzoeken van winnaars over")
    argumenten = p.parse_args()

    print("\nGunningenradar — %s" % dt.datetime.now().strftime("%d-%m-%Y %H:%M"))

    if argumenten.demo:
        ruwe = VOORBEELD
        print("  demo-modus: %d voorbeeldgunningen" % len(ruwe))
    else:
        print("  ophalen van gegunde opdrachten, %d dagen terug..." % argumenten.dagen)
        ruwe = haal_gunningen(argumenten.dagen, argumenten.velden)

    publicaties = []
    for rij in ruwe:
        pub = scoor(normaliseer(rij))
        if argumenten.demo:
            pub["winnaars"] = rij.get("_winnaars", [])
            pub["bron"] = rij.get("_bron", "")
            pub["waarde"] = rij.get("_waarde", "")
        publicaties.append(pub)

    # alleen gunningen, ook als de API het typefilter negeert
    if not argumenten.demo:
        publicaties = [p for p in publicaties if p["typecode"] in ("AGO", "WNO", "")]

    drempel = 0 if argumenten.alles else MIN_SCORE
    relevant = [p for p in publicaties if p["score"] >= drempel]
    print("  %d van %d gunningen relevant" % (len(relevant), len(publicaties)))

    if not argumenten.demo and not argumenten.geen_verrijking:
        cache = lees_json(CACHE, {})
        gebruikersnaam = os.environ.get("TENDERNED_API_USERNAME")
        wachtwoord = os.environ.get("TENDERNED_API_PASSWORD")
        if PdfReader is None:
            print("  let op: pypdf ontbreekt, PDF-route uitgeschakeld "
                  "(pip install pypdf)")
        opgezocht = 0
        for pub in relevant:
            if opgezocht >= MAX_VERRIJKING:
                break
            if verrijk(pub, cache, gebruikersnaam, wachtwoord):
                opgezocht += 1
                if opgezocht % 10 == 0:
                    print("  %d gunningen uitgezocht..." % opgezocht)
        schrijf_json(CACHE, cache)

    eerder = set(lees_json(GEZIEN, {}).get("ids", []))
    nieuwe = {p["id"] for p in relevant if p["id"] and p["id"] not in eerder}

    bestand = schrijf_dashboard(relevant, len(publicaties), argumenten.dagen, nieuwe)
    csv_bestand = schrijf_csv(relevant)
    schrijf_json(GEZIEN, {"bijgewerkt": dt.datetime.now().isoformat(timespec="seconds"),
                          "ids": sorted(eerder | {p["id"] for p in relevant if p["id"]})})

    metnaam = sum(1 for p in relevant if p["winnaars"])
    print("  uitvoerende partij gevonden bij %d van %d" % (metnaam, len(relevant)))
    print("  dashboard: %s" % bestand)
    print("  csv:       %s\n" % csv_bestand)


if __name__ == "__main__":
    main()
