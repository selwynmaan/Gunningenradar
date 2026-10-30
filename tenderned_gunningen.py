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
    python tenderned_gunningen.py --van 548 --tot 183 --xlsx exports/gunningen-historie.xlsx
                                                  # historie: 1,5 jaar tot een half jaar terug

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
import time
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
MAX_PAGINAS = 600
PAGINA_GROOTTE = 100
MIN_SCORE = 3             # minimale relevantiescore
MAX_VERRIJKING = 150      # hoeveel gunningen per run maximaal uitgezocht worden
MAX_POGINGEN = 3          # zo vaak opnieuw proberen als er geen winnaar gevonden is
PAUZE = 0.3               # seconden tussen detailverzoeken, om TenderNed niet te belasten

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


def haal_gunningen(van_dagen, tot_dagen=0, toon_velden=False):
    """Haalt gunningen op met publicatiedatum tussen van_dagen en tot_dagen geleden."""
    if requests is None:
        stop("De module 'requests' ontbreekt. Installeer met: pip install requests")

    vandaag = dt.date.today()
    vanaf = vandaag - dt.timedelta(days=van_dagen)
    tot = vandaag - dt.timedelta(days=tot_dagen)
    alles, pagina = [], 0

    while pagina < MAX_PAGINAS:
        params = {
            "page": pagina,
            "size": PAGINA_GROOTTE,
            "publicatieType": "AGO",
            "publicatieDatumVanaf": vanaf.isoformat(),
            "publicatieDatumTot": tot.isoformat(),
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

        # Datumfilter ook zelf toepassen: de API negeert soms de datumparameters.
        binnen = []
        oudste = None
        for rij in rijen:
            d = datum(kies(rij, "publicatieDatum", "datumPublicatie", "publicatiedatum"))
            if d:
                oudste = d if oudste is None or d < oudste else oudste
            if not d or vanaf.isoformat() <= d <= tot.isoformat():
                binnen.append(rij)
        alles.extend(binnen)
        print("  pagina %d: %d publicaties, %d binnen periode (totaal %d)"
              % (pagina, len(rijen), len(binnen), len(alles)))

        # Stoppen op de echte laatste pagina. NIET op len(rijen) < PAGINA_GROOTTE:
        # TenderNed geeft soms minder per pagina terug dan gevraagd.
        if data.get("last") is True:
            break
        totaal_paginas = data.get("totalPages")
        if isinstance(totaal_paginas, int) and pagina + 1 >= totaal_paginas:
            break
        # Lijst loopt van nieuw naar oud: verder bladeren heeft geen zin meer.
        if oudste and oudste < vanaf.isoformat():
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


def _kind(knoop, naam):
    for k in knoop:
        if strip_ns(k.tag) == naam:
            return k
    return None


def _id(knoop):
    for k in knoop.iter():
        if strip_ns(k.tag) == "ID" and (k.text or "").strip():
            return k.text.strip()
    return ""


def eforms_winnaars(wortel):
    org_naam = {}
    for knoop in wortel.iter():
        if strip_ns(knoop.tag) == "Organization":
            org_id, naam = "", ""
            for k in knoop.iter():
                t = strip_ns(k.tag)
                if t == "PartyIdentification" and not org_id:
                    org_id = _id(k)
                if t == "PartyName" and not naam:
                    for kk in k.iter():
                        if strip_ns(kk.tag) == "Name" and (kk.text or "").strip():
                            naam = kk.text.strip()
            if org_id and naam:
                org_naam[org_id] = naam
    partij_orgs = {}
    for knoop in wortel.iter():
        if strip_ns(knoop.tag) == "TenderingParty":
            tpa = ""
            for k in knoop:
                if strip_ns(k.tag) == "ID":
                    tpa = (k.text or "").strip()
            orgs = [_id(k) for k in knoop.iter() if strip_ns(k.tag) == "Tenderer"]
            if tpa:
                partij_orgs[tpa] = [o for o in orgs if o]
    inschrijving_partij = {}
    for knoop in wortel.iter():
        if strip_ns(knoop.tag) == "LotTender":
            ten, tpa = "", ""
            for k in knoop:
                tag = strip_ns(k.tag)
                if tag == "ID":
                    ten = (k.text or "").strip()
                elif tag == "TenderingParty":
                    tpa = _id(k)
            if ten and tpa:
                inschrijving_partij[ten] = tpa
    winnende_ten = []
    for knoop in wortel.iter():
        if strip_ns(knoop.tag) == "LotResult":
            code = ""
            for k in knoop.iter():
                if strip_ns(k.tag) == "TenderResultCode":
                    code = (k.text or "").strip()
            if code == "selec-w":
                for k in knoop:
                    if strip_ns(k.tag) == "LotTender":
                        winnende_ten.append(_id(k))
    namen = []
    for ten in winnende_ten:
        for org in partij_orgs.get(inschrijving_partij.get(ten, ""), []):
            naam = org_naam.get(org)
            if naam and naam not in namen:
                namen.append(naam)
    return namen


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

    # Route B: eForms — winnaars via LotResult (selec-w) -> LotTender -> TenderingParty -> ORG
    if not namen:
        namen = eforms_winnaars(wortel)

    # Route C: eForms — organisaties opzoeken en de buyer eruit filteren (grof)
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


def _schoon(regel):
    return re.sub(r"\s+", " ", (regel or "")).strip(" :;-")


def _veld(tekst_pdf, label_regex, maxlen=600):
    """Waarde na een label, tot het volgende label of de volgende kop."""
    m = re.search(label_regex + r"\s*:?\s*(.+?)(?=\n\S[^\n]{0,60}:\s|\n\d+(?:\.\d+)+\s"
                  r"|\n[IVX]+\.\d[\d.]*\)|\n(?:Identificatiecode|Interne|Eerdere|Type procedure"
                  r"|Aard van|Rechtsgrondslag|Belangrijkste)|\Z)",
                  tekst_pdf, re.IGNORECASE | re.DOTALL)
    return _schoon(m.group(1))[:maxlen] if m else ""


def lees_pdf_tekst(volledig, buyer=""):
    """Haalt winnaar(s), koper, titel, beschrijving en waarde uit de tekst van een
    TenderNed-publicatie. Werkt voor eForms (vanaf eind 2023) en de oude
    standaardformulieren."""
    info = {"winnaars": [], "koper": "", "titel": "", "beschrijving": "", "waarde": ""}
    if not volledig:
        return info
    t = volledig.replace("\r", "")
    namen = []

    # eForms, deel 6: "Winnaar:" gevolgd door "Officiële naam: X"
    for m in re.finditer(r"Winnaar\s*:\s*\n?\s*Offici[eë]le\s+naam\s*:?\s*([^\n]+)", t,
                         re.IGNORECASE):
        namen.append(_schoon(m.group(1)))

    # eForms, deel 8: organisaties met rol "Winnaar van deze percelen"
    if not namen:
        for blok in re.split(r"\n\s*8\.1\s+ORG-\d+", t):
            if re.search(r"Winnaar\s+van\s+de(?:ze)?", blok, re.IGNORECASE):
                m = re.search(r"Offici[eë]le\s+naam\s*:?\s*([^\n]+)", blok, re.IGNORECASE)
                if m:
                    namen.append(_schoon(m.group(1)))

    # Oude formulieren: "Naam en adres van de contractant" /
    # "Naam en adres van de ondernemer aan wie de opdracht is gegund"
    if not namen:
        for m in re.finditer(r"(naam\s+en\s+adres\s+van\s+de\s+(?:contractant|ondernemer"
                             r"\s+aan\s+wie\s+de\s+opdracht\s+is\s+gegund))", t, re.IGNORECASE):
            staart = t[m.end():m.end() + 600]
            n = re.search(r"Offici[eë]le\s+benaming\s*:?\s*\n?\s*([^\n]+)", staart, re.IGNORECASE)
            if n:
                namen.append(_schoon(n.group(1)))

    schoon_buyer = normaliseer_naam(buyer)
    for naam in namen:
        if len(naam) > 2 and naam not in info["winnaars"] and \
                (not schoon_buyer or normaliseer_naam(naam) != schoon_buyer):
            info["winnaars"].append(naam)
    info["winnaars"] = info["winnaars"][:6]

    # Koper / opdrachtgever
    m = re.search(r"1\.1\s+Koper.*?Offici[eë]le\s+naam\s*:?\s*([^\n]+)", t,
                  re.IGNORECASE | re.DOTALL)
    if not m:
        m = re.search(r"Offici[eë]le\s+benaming\s*:?\s*\n?\s*([^\n]+)", t, re.IGNORECASE)
    if m:
        info["koper"] = _schoon(m.group(1))

    # Titel en beschrijving
    m = re.search(r"2\.1\s+Procedure\s*\n(.*?)(?=\n2\.1\.1\s|\Z)", t, re.DOTALL)
    if m:
        blok = m.group(1)
        info["titel"] = _veld(blok, r"Titel", 250)
        info["beschrijving"] = _veld(blok, r"Beschrijving", 900)
    if not info["titel"]:
        info["titel"] = _veld(t, r"II\.1\.1\)\s*(?:Benaming|Aan de opdracht gegeven benaming)", 250)
    if not info["beschrijving"]:
        info["beschrijving"] = _veld(t, r"II\.1\.4\)\s*Korte\s+beschrijving", 900)

    # Waarde
    m = re.search(r"Waarde\s+van\s+alle.*?gegunde\s+opdrachten\s*:?\s*([\d .,]+\s*[A-Z]{3})",
                  t, re.IGNORECASE | re.DOTALL)
    if not m:
        m = re.search(r"Totale\s+waarde\s+van\s+de\s+opdracht[^:]*:\s*([\d .,]+)", t, re.IGNORECASE)
    if m:
        info["waarde"] = _schoon(m.group(1))
    return info


def uit_pdf(pdf_bytes, buyer):
    if PdfReader is None:
        return {"winnaars": []}
    try:
        lezer = PdfReader(io.BytesIO(pdf_bytes))
        volledig = "\n".join((p.extract_text() or "") for p in lezer.pages[:20])
    except Exception:
        return {"winnaars": []}
    return lees_pdf_tekst(volledig, buyer)


def normaliseer_naam(naam):
    schoon = re.sub(r"\b(b\.?v\.?|n\.?v\.?|v\.?o\.?f\.?|holding|nederland|group|groep)\b",
                    " ", (naam or "").lower())
    return re.sub(r"[^a-z0-9]+", "", schoon)


STATS = {}
DIAGNOSE = {}


def _monster(route, r):
    """Bewaart van elke route het eerste antwoord, voor exports/diagnose.json."""
    sleutel = route + "_antwoord"
    if sleutel in DIAGNOSE:
        return
    begin = r.content[:400]
    try:
        begin = begin.decode("utf-8", "replace")
    except Exception:
        begin = repr(begin)
    DIAGNOSE[sleutel] = {"status": r.status_code, "content_type": r.headers.get("Content-Type"),
                         "url": r.url, "begin": begin}


KOPER_SLEUTELS = ("aanbestedendedienst", "opdrachtgever", "koper", "buyer",
                  "contractingauthority", "organisatienaam", "officielenaam")


def zoek_koper(obj):
    """Zoekt de naam van de opdrachtgever in de detail-JSON."""
    if isinstance(obj, dict):
        for sleutel, waarde in obj.items():
            plat = sleutel.lower().replace("_", "")
            if any(k in plat for k in KOPER_SLEUTELS):
                if isinstance(waarde, str) and 2 < len(waarde) < 160:
                    return waarde.strip()
                if isinstance(waarde, dict):
                    for s2 in ("officieleNaam", "naam", "name", "organisatieNaam", "omschrijving"):
                        if isinstance(waarde.get(s2), str) and waarde[s2].strip():
                            return waarde[s2].strip()
        for waarde in obj.values():
            gevonden = zoek_koper(waarde)
            if gevonden:
                return gevonden
    elif isinstance(obj, list):
        for item in obj:
            gevonden = zoek_koper(item)
            if gevonden:
                return gevonden
    return ""


def _tel(sleutel):
    STATS[sleutel] = STATS.get(sleutel, 0) + 1


def _pas_toe(pub, opgeslagen):
    pub["winnaars"] = opgeslagen.get("winnaars", [])
    pub["bron"] = opgeslagen.get("bron", "")
    pub["waarde"] = opgeslagen.get("waarde", "") or pub.get("waarde", "")
    for veld, doel in (("omschrijving", "omschrijving"), ("koper", "dienst"), ("titel", "titel")):
        if opgeslagen.get(veld) and (not pub.get(doel) or doel == "omschrijving"
                                     or pub.get(doel) == "Zonder titel"):
            pub[doel] = opgeslagen[veld]


def verrijk(pub, cache, gebruikersnaam, wachtwoord):
    """Vult winnaars, waarde en omschrijving aan. Gevonden resultaten worden
    gecachet; lege resultaten worden maximaal MAX_POGINGEN keer opnieuw geprobeerd."""
    opgeslagen = cache.get(pub["id"])
    if opgeslagen and (opgeslagen.get("winnaars")
                       or opgeslagen.get("pogingen", 0) >= MAX_POGINGEN):
        _pas_toe(pub, opgeslagen)
        return False

    basis = "%s/%s" % (BASIS_URL, pub["id"])
    info, bron = {"winnaars": []}, ""

    # Route 1 — de openbare PDF van de publicatie (bewezen formaat, eForms en oud)
    if PdfReader is not None:
        try:
            r = sessie().get(basis + "/pdf", timeout=60,
                             headers={"Accept": "application/pdf, */*"})
            _tel("pdf http %s" % r.status_code)
            _monster("pdf", r)
            if r.ok and r.content[:4] == b"%PDF":
                info = uit_pdf(r.content, pub["dienst"])
                if info.get("winnaars"):
                    bron = "PDF"
                else:
                    _tel("pdf zonder winnaar")
        except Exception as fout:
            _tel("pdf fout %s" % type(fout).__name__)

    # Route 2 — eForms/TED XML
    if not info.get("winnaars"):
        try:
            auth = (gebruikersnaam, wachtwoord) if gebruikersnaam else None
            r = sessie().get(basis + "/public-xml", auth=auth, timeout=40,
                             headers={"Accept": "application/xml, text/xml, */*"})
            _tel("xml http %s" % r.status_code)
            _monster("xml", r)
            if r.ok and r.text.strip().startswith("<"):
                namen, waarde = uit_xml(r.text, pub["dienst"])
                if namen:
                    info["winnaars"], bron = namen, "XML"
                    info["waarde"] = info.get("waarde") or waarde
        except Exception as fout:
            _tel("xml fout %s" % type(fout).__name__)

    # Route 3 — detail-JSON
    if not info.get("winnaars") or not info.get("beschrijving") or not info.get("koper"):
        try:
            r = sessie().get(basis, timeout=30, headers={"Accept": "application/json"})
            _tel("json http %s" % r.status_code)
            if r.ok:
                detail = r.json()
                if "json_detail" not in DIAGNOSE:
                    DIAGNOSE["json_detail"] = detail
                if not info.get("koper"):
                    info["koper"] = zoek_koper(detail)
                if not info.get("titel"):
                    info["titel"] = tekst(kies(detail, "aanbestedingNaam", "publicatieNaam",
                                               "opdrachtNaam", "titel"))[:250]
                if not info.get("winnaars"):
                    kandidaten = [n for n in zoek_diep(detail, WINNAAR_SLEUTELS)
                                  if not n.isdigit()
                                  and normaliseer_naam(n) != normaliseer_naam(pub["dienst"])]
                    if kandidaten:
                        info["winnaars"], bron = kandidaten[:3], "detail"
                if not info.get("beschrijving"):
                    info["beschrijving"] = tekst(kies(detail, "opdrachtBeschrijving",
                                                      "beschrijving"))[:900]
        except Exception as fout:
            _tel("json fout %s" % type(fout).__name__)

    time.sleep(PAUZE)
    _tel("met winnaar" if info.get("winnaars") else "zonder winnaar")

    vorige = (opgeslagen or {}).get("pogingen", 0)
    nieuw = {"winnaars": info.get("winnaars", []), "bron": bron,
             "waarde": info.get("waarde", ""),
             "omschrijving": info.get("beschrijving", ""),
             "koper": info.get("koper", ""), "titel": info.get("titel", ""),
             "pogingen": vorige + 1}
    cache[pub["id"]] = nieuw
    _pas_toe(pub, nieuw)
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

KERN = ("installatie", "werktuigbouwk", "klimaat", "ventilatie", "luchtbehandel", "hvac",
        "inductie", "koelplafond", "brandklep", "verwarming", "koeling", "e&w", "e- en w",
        "w-install", "binnenklimaat", "luchtverdeling", "airco", "warmtepomp")
KERN_CPV = ("4533", "4535", "4530", "71321", "39717", "425")
RUIS = ("jeugdhulp", "meubilair", "leermiddelen", "voeding", "schoonmaak", "catering",
        "software", "laptop", "audiovisu", "av-middelen", "medische", "zonnepanel",
        "verzekering", "accountant", "uitzend", "vervoer", "taxi", "rioler")


def heeft_kern(pub):
    hooiberg = " ".join([pub["titel"], pub["omschrijving"]]).lower()
    if any(r in pub["titel"].lower() for r in RUIS):
        return False
    if any(k in hooiberg for k in KERN):
        return True
    return any(re.sub(r"\D", "", c).startswith(KERN_CPV) for c in pub["cpv"])


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


def schrijf_xlsx(publicaties, pad, van_dagen, tot_dagen):
    """Eén regel per gunning: project, inhoud, opdrachtgever, gegund aan."""
    map_ = os.path.dirname(pad)
    if map_:
        os.makedirs(map_, exist_ok=True)
    kop = ["Projectnaam", "Wat houdt het project in", "Opdrachtgever", "Gegund aan",
           "Soort partij", "Publicatiedatum", "Waarde", "Segment", "Score", "Link TenderNed"]
    rijen = []
    for pub in sorted(publicaties, key=lambda p: (p["publicatie"], p["score"]), reverse=True):
        winnaars = pub["winnaars"]
        rijen.append([pub["titel"], pub["omschrijving"], pub["dienst"],
                      "; ".join(winnaars) if winnaars else "niet gevonden, zie link",
                      " / ".join(sorted({soort_partij(n) for n in winnaars})) if winnaars else "",
                      pub["publicatie"], pub["waarde"], "/".join(pub["segmenten"]),
                      pub["score"], pub["link"]])
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        import csv
        pad = os.path.splitext(pad)[0] + ".csv"
        with open(pad, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(kop)
            w.writerows(rijen)
        return pad

    wb = Workbook()
    ws = wb.active
    ws.title = "Gunningen"
    ws.append(kop)
    for rij in rijen:
        ws.append(rij)
    vet = Font(name="Arial", bold=True, color="FFFFFF")
    for cel in ws[1]:
        cel.font = vet
        cel.fill = PatternFill("solid", fgColor="15201D")
        cel.alignment = Alignment(vertical="center", wrap_text=True)
    for rij in ws.iter_rows(min_row=2):
        for cel in rij:
            cel.font = Font(name="Arial", size=10)
            cel.alignment = Alignment(vertical="top", wrap_text=cel.column in (1, 2, 4))
        link = rij[9]
        if link.value:
            link.hyperlink = link.value
            link.font = Font(name="Arial", size=10, color="0563C1", underline="single")
    for kolom, breedte in zip("ABCDEFGHIJ", (42, 70, 30, 34, 16, 14, 16, 14, 8, 48)):
        ws.column_dimensions[kolom].width = breedte
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    info = wb.create_sheet("Toelichting")
    vandaag = dt.date.today()
    regels = [
        ["Gunningenradar — HVS"],
        ["Periode", "%s t/m %s" % ((vandaag - dt.timedelta(days=van_dagen)).strftime("%d-%m-%Y"),
                                    (vandaag - dt.timedelta(days=tot_dagen)).strftime("%d-%m-%Y"))],
        ["Aangemaakt", dt.datetime.now().strftime("%d-%m-%Y %H:%M")],
        ["Bron", "TenderNed publicatie-webservice (open data, CC-0): gegunde opdrachten (AGO)"],
        ["Gegund aan", "uit de publicatie-PDF (eForms deel 6 'Winnaar') of de eForms-XML"],
        ["Selectie", "relevantiescore op trefwoorden voor W-installaties, advies en gebouwtypes"],
        ["Aantal gunningen", len(rijen)],
        ["Met winnaar", sum(1 for p in publicaties if p["winnaars"])],
    ]
    for r in regels:
        info.append(r)
    info["A1"].font = Font(name="Arial", bold=True, size=12)
    info.column_dimensions["A"].width = 20
    info.column_dimensions["B"].width = 90
    wb.save(pad)
    return pad


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
    p.add_argument("--van", type=int, default=None,
                   help="begin van de periode, in dagen geleden (bijv. 548)")
    p.add_argument("--tot", type=int, default=0,
                   help="eind van de periode, in dagen geleden (bijv. 183)")
    p.add_argument("--xlsx", default="",
                   help="schrijf een spreadsheet naar dit pad (historie-modus)")
    p.add_argument("--max-verrijking", type=int, default=MAX_VERRIJKING)
    p.add_argument("--test-pdf", default="",
                   help="test de winnaar-herkenning op een lokaal PDF- of tekstbestand")
    argumenten = p.parse_args()

    print("\nGunningenradar — %s" % dt.datetime.now().strftime("%d-%m-%Y %H:%M"))

    if argumenten.test_pdf:
        with open(argumenten.test_pdf, "rb") as f:
            inhoud = f.read()
        if inhoud[:4] == b"%PDF":
            print(json.dumps(uit_pdf(inhoud, ""), ensure_ascii=False, indent=2))
        else:
            print(json.dumps(lees_pdf_tekst(inhoud.decode("utf-8"), ""),
                             ensure_ascii=False, indent=2))
        return

    van_dagen = argumenten.van if argumenten.van is not None else argumenten.dagen
    if van_dagen < argumenten.tot:
        # volgorde maakt niet uit: het grootste getal is altijd het begin van de periode
        van_dagen, argumenten.tot = argumenten.tot, van_dagen
    historie = bool(argumenten.xlsx)

    if argumenten.demo:
        ruwe = VOORBEELD
        print("  demo-modus: %d voorbeeldgunningen" % len(ruwe))
    else:
        print("  ophalen van gegunde opdrachten, van %d tot %d dagen terug..."
              % (van_dagen, argumenten.tot))
        ruwe = haal_gunningen(van_dagen, argumenten.tot, argumenten.velden)

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
            if opgezocht >= argumenten.max_verrijking:
                # boven de limiet: wel bekende gegevens uit de cache gebruiken
                if pub["id"] in cache:
                    _pas_toe(pub, cache[pub["id"]])
                continue
            if verrijk(pub, cache, gebruikersnaam, wachtwoord):
                opgezocht += 1
                if opgezocht % 10 == 0:
                    print("  %d gunningen uitgezocht..." % opgezocht)
                if opgezocht % 50 == 0:
                    schrijf_json(CACHE, cache)
        schrijf_json(CACHE, cache)
        if STATS:
            print("  diagnose: " + ", ".join("%s=%d" % kv for kv in sorted(STATS.items())))

    if historie:
        DIAGNOSE["stats"] = STATS
        DIAGNOSE["lijst_velden"] = sorted(ruwe[0].keys()) if ruwe and isinstance(ruwe[0], dict) else []
        DIAGNOSE["lijst_voorbeeld"] = ruwe[0] if ruwe else {}
        map_ = os.path.dirname(argumenten.xlsx) or "."
        os.makedirs(map_, exist_ok=True)
        with open(os.path.join(map_, "diagnose.json"), "w", encoding="utf-8") as f:
            json.dump(DIAGNOSE, f, ensure_ascii=False, indent=1, default=str)
        voor = len(relevant)
        relevant = [scoor(p) for p in relevant]
        if not argumenten.alles:
            relevant = [p for p in relevant if heeft_kern(p)]
        print("  %d van %d over na filter op installatie-inhoud" % (len(relevant), voor))
        pad = schrijf_xlsx(relevant, argumenten.xlsx, van_dagen, argumenten.tot)
        metnaam = sum(1 for p in relevant if p["winnaars"])
        print("  uitvoerende partij gevonden bij %d van %d" % (metnaam, len(relevant)))
        print("  spreadsheet: %s\n" % pad)
        return

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
