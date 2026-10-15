# -*- coding: utf-8 -*-
"""
TIMAR Tanger - Ventilation Automatique Groupage Alstom Fez
==========================================================

Application Web Streamlit (aucun droit administrateur requis) qui :
  1. lit un dossier PDF de remorque de groupage (factures, packing lists,
     CMR, manifeste), y compris les pages scannées (OCR) ;
  2. extrait chaque ligne de facture de manière STRICTEMENT individuelle ;
  3. rapproche chaque référence avec le fichier Excel de référence (NGP) ;
  4. génère un classeur Excel à 3 feuilles : TOUS_CLIENTS, CONTROLE, RECAP_GLOBAL.

Lancement :  streamlit run app.py
"""

from __future__ import annotations

import difflib
import io
import math
import re
import shutil
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st

import ui
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:  # PyPDF2 : extraction de secours si pdfplumber échoue sur une page
    from PyPDF2 import PdfReader
except Exception:  # pragma: no cover
    PdfReader = None


# =============================================================================
# CONSTANTES
# =============================================================================

APP_TITLE = "TIMAR Tanger - Ventilation Automatique Groupage Alstom Fez"
NON_TROUVE = "NON TROUVÉ"
REMARQUE_NOUVELLE_REF = "Nouveau Réf - NGP Facture appliqué"

COLUMNS = [
    "Client", "Job Ref manifest", "N° facture", "Date", "Item",
    "Réf article", "Désignation", "Origine",
    "Qté", "Prix unitaire", "Valeur HT",
    "NGP/code douanier (facture)", "NGP (REF_PAR_FRS)", "Onglet REF_PAR_FRS",
    "Contrôle NGP", "Remarque", "Page PDF",
]
NUMERIC_COLUMNS = ["Qté", "Prix unitaire", "Valeur HT"]
# Colonnes saisies / corrigeables par l'utilisateur (le matching est recalculé)
EDITABLE_COLUMNS = [c for c in COLUMNS if c not in ("NGP (REF_PAR_FRS)", "Contrôle NGP", "Onglet REF_PAR_FRS")]
MANIFEST_COLUMNS = ["Page PDF", "Shipment", "Expéditeur", "Poids (kg)", "Colis", "Goods Value", "Devise"]

CURRENCIES = ["EUR", "USD", "GBP", "CHF", "CNY", "JPY", "SEK", "DKK", "NOK",
              "PLN", "CZK", "HUF", "TRY", "CAD", "MAD", "INR"]
INCOTERMS = ["EXW", "FCA", "FAS", "FOB", "CFR", "CIF", "CPT", "CIP", "DAP",
             "DPU", "DAT", "DDP", "DDU"]
ISO_COUNTRIES = set("""AD AE AF AL AM AO AR AT AU AZ BA BD BE BG BH BR BY CA CH CL CN CO CR CY CZ DE
DK DZ EC EE EG ES FI FR GB GE GR HK HR HU ID IE IL IN IQ IR IS IT JO JP KE KR KW KZ LB LK LT LU LV
LY MA MC MD ME MK MT MX MY NG NL NO NZ OM PE PH PK PL PT QA RO RS RU SA SE SG SI SK SN TH TN TR TW
UA US UY UZ VN ZA""".split())
COUNTRY_NAMES = {
    "GERMANY": "DE", "ALLEMAGNE": "DE", "DEUTSCHLAND": "DE", "FRANCE": "FR",
    "ITALY": "IT", "ITALIE": "IT", "ITALIA": "IT", "SPAIN": "ES", "ESPAGNE": "ES",
    "ESPANA": "ES", "CHINA": "CN", "CHINE": "CN", "UNITED STATES": "US", "USA": "US",
    "ETATS-UNIS": "US", "UNITED KINGDOM": "GB", "ROYAUME-UNI": "GB", "UK": "GB",
    "POLAND": "PL", "POLOGNE": "PL", "CZECH REPUBLIC": "CZ", "CZECHIA": "CZ",
    "TURKEY": "TR", "TURQUIE": "TR", "JAPAN": "JP", "JAPON": "JP", "INDIA": "IN",
    "INDE": "IN", "SWITZERLAND": "CH", "SUISSE": "CH", "BELGIUM": "BE",
    "BELGIQUE": "BE", "NETHERLANDS": "NL", "PAYS-BAS": "NL", "AUSTRIA": "AT",
    "AUTRICHE": "AT", "SWEDEN": "SE", "SUEDE": "SE", "PORTUGAL": "PT",
    "MOROCCO": "MA", "MAROC": "MA", "ROMANIA": "RO", "ROUMANIE": "RO",
    "HUNGARY": "HU", "HONGRIE": "HU", "SLOVAKIA": "SK", "SLOVAQUIE": "SK",
    "KOREA": "KR", "COREE": "KR", "TAIWAN": "TW", "MEXICO": "MX", "MEXIQUE": "MX",
    "EU": "EU", "EUROPEAN UNION": "EU", "UNION EUROPEENNE": "EU",
}
UNITS = {"PCS", "PC", "PCE", "PCES", "EA", "EACH", "ST", "STK", "STCK", "UN", "U",
         "UNIT", "UNITS", "PZ", "PZA", "KG", "KGS", "G", "M", "MM", "M2", "M3",
         "L", "LTR", "SET", "SETS", "KIT", "BOX", "ROL", "ROLL", "PR", "PAIR",
         "PAA", "LOT", "PK", "PAC", "PACK", "MTR", "NR", "NOS"}

FEE_WORDS = re.compile(
    r"\b(freight|fret|transport|shipping|carriage|packing\s*cost|packaging|emballage|"
    r"frais|surcharge|insurance|assurance|handling|verpackung|fracht|porto|"
    r"palette\s*charge|pallet\s*charge|discount|remise|rabatt)\b", re.I)
TOTAL_WORDS = re.compile(
    r"\b(sub\s*-?\s*total|total|sous[- ]?total|summe|gesamt\w*|zwischensumme|"
    r"net\s*amount|amount\s*due|montant\s*(?:net|total|ht|ttc)|tva|vat|mwst|"
    r"carried\s*forward|report|übertrag|a\s*reporter)\b", re.I)

# Mapping des en-têtes de colonnes (FR / EN / DE / ES / IT) -> champ interne.
# L'ordre est important : le premier motif qui correspond l'emporte.
LINE_COLMAP = [
    ("ngp", r"\b(hs|h\.s\.|ngp|tarif\w*|tariff|customs|douan\w*|zoll\w*|taric|commodity|nomencl\w*|statist\w*)"),
    ("po", r"(\bp\.?\s?o\.?\b(?!\s*box)|purchase\s*order|order\s*(no|n°|nr|number|ref)|commande|bestell\w*|\bpo\s*(no|n°|number))"),
    ("alstom", r"(alstom|customer|client|your\s*(ref|part|p/?n|art)|votre|kunden|cust\.?\s*p)"),
    ("net", r"(net\w*\s*(weight|wt|kg|mass)|poids\s*net|netto\w*|n\.?\s?w\.?\b|peso\s*neto|peso\s*netto)"),
    ("gross", r"(gross\s*(weight|wt|kg|mass)|poids\s*brut|brutto\w*|g\.?\s?w\.?\b|peso\s*bruto|peso\s*lordo)"),
    ("price", r"(unit\s*price|prix\s*unit\w*|p\.\s?u\.?|^pu$|price|prix|preis|precio|prezzo|einzel\w*|rate)"),
    ("amount", r"(amount|montant|total|valeur|value|betrag|importe|importo|gesamt|line\s*total|ext\w*\s*price)"),
    ("qty", r"(qty|q\.?t[eé]|quantit\w*|menge|cantidad|qt[aà]|anzahl|^pcs$|^nb$)"),
    ("unit", r"^(unit|unit[eé]|uom|u\.?m\.?|me|einheit|unidad|um|u\.?)$"),
    ("origin", r"(origin\w*|origine|coo\b|ursprung\w*|c\.?o\.?o\.?|made\s*in|pays|herkunft)"),
    ("brand", r"(brand|marque|manufacturer|fabricant|hersteller|make\b|marca)"),
    ("packages", r"(colis|packages?|pkgs?|cartons?|packst\w*|bultos|nb\s*col)"),
    ("desc", r"(descr\w*|d[eé]sign\w*|bezeichnung|libell\w*|denominaz\w*|item\s*name|goods|text)"),
    ("ref", r"(part|article|art\.?\s*(no|nr|n°)?|r[eé]f\w*|material|item\s*(no|code|number|#)|code|artikel|product|sku|p/?n)"),
    ("pos", r"^(pos\w*|position|item|line|ligne|poste|no\.?|n°|#|nr\.?|lfd\.?|sr\.?\s*no\.?|s/n)$"),
]
MANIFEST_COLMAP = [
    ("shipment", r"(shipment|job\s*ref|bl\s*no|reference|ref\.?\s*exp|waybill|sxvf|\bawb\b|\bbno\b)"),
    ("invoice", r"(invoice|facture|inv\b|rechnung|factura)"),
    ("currency", r"(curr\w*|devise|cur\b|monnaie|w[aä]hrung)"),
    ("gross", r"(gross|brut|poids|weight|kg|gewicht)"),
    ("packages", r"(colis|pack\w*|pkgs?|pallets?|palettes?|cartons?|nb\b|qty|nombre)"),
    ("value", r"(value|valeur|montant|amount|goods|wert|prix|price)"),
    ("supplier", r"(shipper|supplier|fournisseur|exp[eé]diteur|exporter|vendor|lieferant|consignor|sender|soci[eé]t[eé]|company|name|nom)"),
]
REF_SHEET_COLMAP = [
    ("ngp", r"(ngp|\bhs\b|h\.s\.|nomencl\w*|tarif\w*|code\s*douan\w*|customs|\bsh\b|taric|commodity)"),
    ("alstom", r"(alstom|\bpn\b|p/n|part\s*number\s*alstom|client|customer)"),
    ("ref", r"(r[eé]f\w*|reference|article|part|material|code|artikel|item|sku)"),
]

HS_LABEL_RE = re.compile(
    r"(?:\bhs\b|h\.s\.|\bngp\b|tarif\w*|customs\s*(?:tariff|code)?|commodity\s*code|"
    r"code\s*douan\w*|zoll\w*|taric|nomenclature|statistical\s*n\w*)"
    r"\s*(?:code|no\.?|nr\.?|n°|number)?\s*[:#.]?\s*(\d[\d .]{5,14}\d)", re.I)
ORIGIN_LABEL_RE = re.compile(
    r"(?:country\s*of\s*origin|pays\s*d'?\s*origine|origin\w*|origine|ursprung\w*|\bcoo\b|made\s*in)"
    r"\s*[:#.]?\s*([A-Za-z][A-Za-z\- ]{1,25})", re.I)
ALSTOM_LABEL_RE = re.compile(
    r"(?:alstom\s*(?:p/?n|part\s*(?:no|number)?|r[eé]f\w*|code|art\w*)?|customer\s*(?:p/?n|part\s*(?:no\.?|number)?|r[eé]f\w*|item)|"
    r"your\s*(?:p/?n|part\s*(?:no\.?|number)?|r[eé]f\w*|item)|votre\s*r[eé]f\w*|r[eé]f\.?\s*client|kunden\w*)"
    r"\s*[:#.]?\s*([A-Z0-9][A-Z0-9\-_./]{2,})", re.I)
BRAND_LABEL_RE = re.compile(
    r"(?:brand|marque|manufacturer|fabricant|hersteller|marca)\s*[:#]\s*([A-Za-z0-9&\-. ]{2,30})", re.I)
LINE_NET_RE = re.compile(r"(?:net\s*weight|poids\s*net|nettogewicht|n\.w\.)\s*(?:\(kg\)|kgs?)?\s*[:=]?\s*([\d][\d.,]*)", re.I)
LINE_GROSS_RE = re.compile(r"(?:gross\s*weight|poids\s*brut|bruttogewicht|g\.w\.)\s*(?:\(kg\)|kgs?)?\s*[:=]?\s*([\d][\d.,]*)", re.I)

INV_NO_RE = re.compile(
    r"(?:commercial\s+invoice|invoice|facture(?:\s+commerciale)?|rechnung|factura|fattura|inv\.)"
    r"\s*(?:no\.?|n[°ºo]\.?|nr\.?|number|num[ée]ro|#|nummer)?\s*[:.#]?\s*"
    r"([A-Z0-9][A-Z0-9\-/_.]*\d[A-Z0-9\-/_]*)", re.I)
INV_LABEL_ONLY_RE = re.compile(
    r"(?:invoice|facture|rechnung|factura|fattura)\s*(?:no\.?|n[°º]\.?|nr\.?|number|num[ée]ro|#|nummer)", re.I)
BL_NO_RE = re.compile(
    r"(?:delivery\s*note|packing\s*list|bon\s*de\s*livraison|liste\s*de\s*colisage|lieferschein|"
    r"albar[aá]n|\bb\.?l\.?\b|\bd\.?n\.?\b|\bp\.?l\.?\b|shipment|exp[ée]dition|despatch\s*note|dispatch\s*note)"
    r"\s*(?:no\.?|n[°º]\.?|nr\.?|number|num[ée]ro|#|nummer)\s*[:.#]?\s*"
    r"([A-Z0-9][A-Z0-9\-/_.]*\d[A-Z0-9\-/_]*)", re.I)
BL_NO_RE2 = re.compile(
    r"(?:delivery\s*note|packing\s*list|bon\s*de\s*livraison|lieferschein|despatch\s*note|dispatch\s*note)"
    r"\s*[:.#]?\s*([A-Z0-9][A-Z0-9\-/_]*\d[A-Z0-9\-/_]*)", re.I)
MONTHS = r"(?:jan|feb|f[eé]v|mar|apr|avr|may|mai|jun|juin|jul|juil|aug|ao[uû]|sep|oct|nov|d[eé]c|dec)"
DATE_RE = re.compile(
    r"\b(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{4}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}\.?\s+" + MONTHS +
    r"[a-zéû]*\.?\s+\d{4}|" + MONTHS + r"[a-zéû]*\.?\s+\d{1,2},?\s+\d{4})\b", re.I)
TOTAL_NET_RE = re.compile(
    r"(?:total\s*)?(?:net\s*weight|poids\s*net(?:\s*total)?|nettogewicht|peso\s*neto|net\s*wt\.?|n\.\s?w\.)"
    r"\s*(?:\((?:kg|kgs)\)|kgs?|in\s*kg)?\s*[:=]?\s*(\d[\d.,\s]*\d|\d)\s*(?:kg|kgs)?", re.I)
TOTAL_GROSS_RE = re.compile(
    r"(?:total\s*)?(?:gross\s*weight|poids\s*brut(?:\s*total)?|bruttogewicht|peso\s*bruto|gross\s*wt\.?|g\.\s?w\.)"
    r"\s*(?:\((?:kg|kgs)\)|kgs?|in\s*kg)?\s*[:=]?\s*(\d[\d.,\s]*\d|\d)\s*(?:kg|kgs)?", re.I)
PACKAGES_RE = re.compile(
    r"(?:(?:total\s*)?(?:number\s*of\s*packages|no\.?\s*of\s*packages|packages|nombre\s*de\s*colis|nb\.?\s*colis|colis|"
    r"packst[üu]cke|bultos|pallets?|palettes?|cartons?)\s*[:=]?\s*(\d{1,4})\b)|"
    r"(?:\b(\d{1,4})\s*(?:colis|packages|pallets?|palettes?|cartons?|cases|caisses|crates)\b)", re.I)
SUPPLIER_LABEL_RE = re.compile(
    r"^\s*(?:seller|supplier|vendor|fournisseur|exp[ée]diteur|shipper|exporter|exportateur|lieferant|"
    r"verk[äa]ufer|absender|proveedor|fornitore|consignor|sender)\s*(?:/\s*\w+\s*)?[:\-]?\s*(.*)$", re.I)
LEGAL_RE = re.compile(
    r"\b(GmbH|AG|S\.?A\.?S\.?U?|S\.?A\.?|S\.?p\.?A\.?|S\.?r\.?l\.?|SARL|EURL|Ltd\.?|Limited|LLC|Inc\.?|Corp\.?|"
    r"B\.?V\.?|N\.?V\.?|S\.?L\.?U?|KG|Co\.?\s*KG|Oy|AB|A/S|ApS|Sp\.?\s*z\s*o\.?\s*o\.?|s\.?r\.?o\.?|a\.?s\.?|"
    r"PLC|SE|S\.?C\.?A\.?|Kft\.?|Industries|Group)\b")
CONSIGNEE_WORDS = re.compile(
    r"(alstom|timar|ship\s*to|bill\s*to|invoice\s*to|sold\s*to|deliver\w*\s*to|consignee|destinataire|"
    r"client|customer|buyer|acheteur|empf[äa]nger|livr[ée]\s*[àa])", re.I)


# =============================================================================
# OUTILS GÉNÉRIQUES
# =============================================================================

def strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def clean_cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    s = str(v).replace("\u00a0", " ").replace("\u202f", " ").strip()
    if s.lower() in ("nan", "none", "nat"):
        return ""
    return s


def detect_decimal(text: str) -> str:
    """Détecte le séparateur décimal dominant d'un document (',' européen ou '.' anglo-saxon)."""
    comma = len(re.findall(r"\d,\d{2}(?!\d)", text))
    dot = len(re.findall(r"\d\.\d{2}(?!\d)", text))
    return "," if comma > dot else "."


def parse_number(value, decimal: str = ".") -> float | None:
    """Convertit '1.234,56' / '1,234.56' / '1 234,56' / 12.5 en float."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if (isinstance(value, float) and math.isnan(value)) else float(value)
    t = clean_cell(value)
    if not t:
        return None
    t = re.sub(r"(?i)(EUR|USD|GBP|CHF|MAD|€|\$|£|kgs?|pcs|%)", "", t).strip()
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg, t = True, t[1:-1]
    if t.endswith("-"):
        neg, t = True, t[:-1]
    if t.startswith("-"):
        neg, t = True, t[1:]
    t = t.replace(" ", "").replace("'", "")
    if not t or not re.fullmatch(r"\d[\d.,]*", t) or t[-1] in ".,":
        return None
    has_dot, has_comma = "." in t, "," in t
    if has_dot and has_comma:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif has_comma or has_dot:
        sep = "," if has_comma else "."
        parts = t.split(sep)
        if len(parts) > 2:  # plusieurs séparateurs identiques = milliers
            if not all(len(p) == 3 for p in parts[1:]):
                return None
            t = t.replace(sep, "")
        elif sep == decimal:
            t = t.replace(",", ".")
        elif len(parts[1]) == 3 and parts[0] != "0":  # séparateur de milliers
            t = t.replace(sep, "")
        else:
            t = t.replace(",", ".")
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


def normalize_ngp(value) -> str:
    """Normalise un code NGP/HS : supprime espaces, points, tirets -> 8 ou 10 chiffres."""
    s = clean_cell(value)
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0+", s):  # nombre lu par Excel (84818099.0)
        s = s.split(".")[0]
    digits = re.sub(r"\D", "", s)
    if len(digits) < 6:
        return ""
    if len(digits) == 7:      # zéro de tête perdu par Excel (chapitres 01 à 09)
        digits = digits.zfill(8)
    elif len(digits) == 9:
        digits = digits.zfill(10)
    elif len(digits) > 10:
        digits = digits[:10]
    return digits


def norm_ref(value) -> str:
    s = clean_cell(value)
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    return re.sub(r"\s+", "", s).upper()


def compact_ref(value) -> str:
    return re.sub(r"[^A-Z0-9]", "", norm_ref(value))


def ocr_ref(value) -> str:
    """Clé tolérante aux confusions OCR courantes (O/0, I/L/1, S/5, B/8)."""
    return compact_ref(value).translate(str.maketrans("OILSB", "01158"))


def norm_company(name: str) -> str:
    s = strip_accents(clean_cell(name)).upper()
    s = LEGAL_RE.sub(" ", s) if s else s
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = re.sub(r"\b(GMBH|AG|SA|SAS|SPA|SRL|SARL|LTD|LIMITED|LLC|INC|CORP|BV|NV|SL|KG|CO|THE|ET|AND|CIE)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def company_similarity(a: str, b: str) -> float:
    na, nb = norm_company(a), norm_company(b)
    if not na or not nb:
        return 0.0
    if na == nb or na in nb or nb in na:
        return 1.0
    ta, tb = set(na.split()), set(nb.split())
    jacc = len(ta & tb) / max(1, len(ta | tb))
    return max(difflib.SequenceMatcher(None, na, nb).ratio(), jacc)


def normalize_date(raw: str) -> str:
    raw = raw.strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
                "%d.%m.%y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return raw


def normalize_origin(raw: str) -> str:
    s = strip_accents(clean_cell(raw)).upper().strip(" .:-")
    if not s:
        return ""
    if s in ISO_COUNTRIES or s == "EU":
        return s
    for name, code in COUNTRY_NAMES.items():
        if s.startswith(name):
            return code
    first = s.split()[0]
    if first in ISO_COUNTRIES:
        return first
    return s[:20]


def text_lines(text: str) -> list[str]:
    return [re.sub(r"[ \t]+", " ", ln).strip() for ln in (text or "").splitlines() if ln.strip()]


# =============================================================================
# LECTURE PDF + OCR
# =============================================================================

def tesseract_ready(custom_path: str = "") -> bool:
    try:
        import pytesseract
    except Exception:
        return False
    if custom_path:
        pytesseract.pytesseract.tesseract_cmd = custom_path
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def easyocr_ready() -> bool:
    try:
        import easyocr  # noqa: F401
        return True
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def get_easyocr_reader(langs: tuple):
    import easyocr
    return easyocr.Reader(list(langs), gpu=False)


def remove_table_lines(image, min_frac: float = 0.04):
    """Prétraitement OCR : niveaux de gris + effacement des filets de tableau (traits longs
    horizontaux / verticaux) qui perturbent la reconnaissance des caractères."""
    import numpy as np
    from PIL import Image
    gray = np.array(image.convert("L"))
    dark = gray < 160
    out = gray.copy()
    for axis in (1, 0):
        k = max(25, int(dark.shape[axis] * min_frac))
        arr = dark if axis == 1 else dark.T
        cs = np.cumsum(np.pad(arr.astype(np.int32), ((0, 0), (1, 0))), axis=1)
        full = (cs[:, k:] - cs[:, :-k]) == k          # fenêtre entièrement noire
        mark = np.zeros(arr.shape[1] + 1, dtype=np.int32)[None, :].repeat(arr.shape[0], 0)
        idx = np.nonzero(full)
        np.add.at(mark, (idx[0], idx[1]), 1)
        np.add.at(mark, (idx[0], idx[1] + k), -1)
        line_px = np.cumsum(mark, axis=1)[:, :-1] > 0
        if axis == 0:
            line_px = line_px.T
        out[line_px] = 255
    return Image.fromarray(out)


class OCREngine:
    """Encapsule pytesseract (si le binaire est présent) ou easyocr (100 % pip)."""

    def __init__(self, mode: str = "Automatique", tess_path: str = "", dpi: int = 300):
        self.dpi = dpi
        self.engine = None
        self.tess_path = tess_path
        if mode in ("Automatique", "Tesseract") and tesseract_ready(tess_path):
            self.engine = "tesseract"
        elif mode in ("Automatique", "EasyOCR") and easyocr_ready():
            self.engine = "easyocr"

    @property
    def available(self) -> bool:
        return self.engine is not None

    def image_to_words(self, image) -> list[dict]:
        """OCR d'une image -> liste de mots positionnés {x0, x1, top, bottom, text}."""
        words = []
        try:
            image = remove_table_lines(image)
        except Exception:
            pass
        if self.engine == "tesseract":
            import pytesseract
            data = None
            for lang in ("fra+eng+deu", "fra+eng", "eng"):
                try:
                    data = pytesseract.image_to_data(image, lang=lang, config="--psm 6",
                                                     output_type=pytesseract.Output.DICT)
                    break
                except Exception:
                    continue
            if data:
                for i, txt in enumerate(data["text"]):
                    txt = (txt or "").strip()
                    if txt and float(data["conf"][i]) > 0:
                        x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
                        words.append({"x0": x, "x1": x + w, "top": y, "bottom": y + h, "text": txt})
        elif self.engine == "easyocr":
            import numpy as np
            reader = get_easyocr_reader(("fr", "en", "de"))
            for box, txt, _conf in reader.readtext(np.array(image.convert("RGB")), detail=1, paragraph=False):
                xs, ys = [p[0] for p in box], [p[1] for p in box]
                for k, part in enumerate(str(txt).split()):
                    # découpage approximatif des boîtes multi-mots
                    n = len(str(txt).split())
                    w = (max(xs) - min(xs)) / max(n, 1)
                    words.append({"x0": min(xs) + k * w, "x1": min(xs) + (k + 1) * w,
                                  "top": min(ys), "bottom": max(ys), "text": part})
        out = []
        for w in words:
            if re.fullmatch(r"[|\[\]_—–=:;!]+", w["text"]):
                continue
            if len(re.findall(r"\d", w["text"])) >= 3:  # « DTRO00573478 » -> « DTR000573478 »
                w["text"] = re.sub(r"(?<=\d)[Oo]|[Oo](?=\d)", "0", w["text"])
                w["text"] = re.sub(r"(?<=\d)[Il|](?=\d)", "1", w["text"])
            out.append(w)
        return out


# =============================================================================
# RECONSTRUCTION DES LIGNES ET TABLEAUX À PARTIR DES POSITIONS DES MOTS
# =============================================================================

def group_word_lines(words: list[dict]) -> list[list[dict]]:
    if not words:
        return []
    heights = sorted(w["bottom"] - w["top"] for w in words)
    med_h = max(heights[len(heights) // 2], 1)
    lines, current, cur_y = [], [], None
    for w in sorted(words, key=lambda w: ((w["top"] + w["bottom"]) / 2, w["x0"])):
        y = (w["top"] + w["bottom"]) / 2
        if cur_y is None or abs(y - cur_y) <= 0.55 * med_h:
            current.append(w)
            cur_y = y if cur_y is None else (cur_y * (len(current) - 1) + y) / len(current)
        else:
            lines.append(sorted(current, key=lambda w: w["x0"]))
            current, cur_y = [w], y
    if current:
        lines.append(sorted(current, key=lambda w: w["x0"]))
    return lines


def words_to_text(words: list[dict]) -> str:
    return "\n".join(" ".join(w["text"] for w in line) for line in group_word_lines(words))


def _phrases(line: list[dict], gap: float) -> list[dict]:
    out = []
    for w in line:
        if out and w["x0"] - out[-1]["x1"] <= gap:
            out[-1]["x1"] = w["x1"]
            out[-1]["text"] += " " + w["text"]
        else:
            out.append(dict(w))
    return out


def words_to_tables(words: list[dict]) -> list[list[list[str]]]:
    """Reconstitue des tableaux (sans filets) : une ligne d'en-tête reconnue définit les colonnes,
    chaque mot des lignes suivantes est affecté à la colonne la plus proche."""
    lines = group_word_lines(words)
    if not lines:
        return []
    heights = sorted(w["bottom"] - w["top"] for w in words)
    gap = 0.8 * max(heights[len(heights) // 2], 1)
    header_idx = []
    for i, line in enumerate(lines):
        cells = [p["text"] for p in _phrases(line, gap)]
        ml, mm = map_header(cells, LINE_COLMAP), map_header(cells, MANIFEST_COLMAP)
        if (len(ml) >= 3 and ("qty" in ml or "amount" in ml or "net" in ml or "gross" in ml)) or \
                (len(mm) >= 3 and "supplier" in mm):
            header_idx.append(i)
    tables = []
    for n, h in enumerate(header_idx):
        header = _phrases(lines[h], gap)
        bounds = [(header[k]["x1"] + header[k + 1]["x0"]) / 2 for k in range(len(header) - 1)]
        hmap = map_header([p["text"] for p in header], LINE_COLMAP)
        numeric_cols = sorted(v for k, v in hmap.items() if k in ("qty", "price", "amount", "net", "gross", "packages"))
        stop = header_idx[n + 1] if n + 1 < len(header_idx) else len(lines)
        table = [[p["text"] for p in header]]
        for line in lines[h + 1:stop]:
            row = [""] * len(header)
            for w in line:
                cx = (w["x0"] + w["x1"]) / 2
                col = sum(1 for b in bounds if cx > b)
                row[col] = (row[col] + " " + w["text"]).strip()
            # mots de désignation qui débordent dans une colonne numérique : on les rend à la colonne de gauche
            for col in numeric_cols:
                toks = row[col].split()
                lead = []
                while len(toks) > 1 and not re.search(r"\d", toks[0]):
                    lead.append(toks.pop(0))
                if lead and col > 0:
                    row[col - 1] = (row[col - 1] + " " + " ".join(lead)).strip()
                    row[col] = " ".join(toks)
            table.append(row)
        tables.append(table)
    return tables


@dataclass
class PageData:
    file: str
    number: int
    text: str
    tables: list               # tableaux à filets (pdfplumber)
    ptables: list = field(default_factory=list)  # tableaux reconstitués par position des mots
    ocr: bool = False
    kind: str = "AUTRE"

    def table_sets(self) -> list:
        return [t for t in (self.tables, self.ptables) if t]


def read_pdf_pages(name: str, data: bytes, ocr: OCREngine, log: list, progress_cb=None) -> list[PageData]:
    pages: list[PageData] = []
    pypdf_reader = None
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            total = len(pdf.pages)
            for i, page in enumerate(pdf.pages, start=1):
                text, tables, ptables, used_ocr = "", [], [], False
                try:
                    text = page.extract_text(x_tolerance=1.5, y_tolerance=3) or ""
                except Exception:
                    text = ""
                if not text.strip() and PdfReader is not None:
                    try:
                        pypdf_reader = pypdf_reader or PdfReader(io.BytesIO(data))
                        text = pypdf_reader.pages[i - 1].extract_text() or ""
                    except Exception:
                        pass
                try:
                    tables = page.extract_tables() or []
                except Exception:
                    tables = []
                try:
                    ptables = words_to_tables(page.extract_words(x_tolerance=1.5, y_tolerance=3))
                except Exception:
                    ptables = []
                if len(re.sub(r"\s", "", text)) < 40:  # page scannée -> OCR
                    if ocr.available:
                        try:
                            img = page.to_image(resolution=ocr.dpi).original
                            words = ocr.image_to_words(img)
                            text, tables, ptables, used_ocr = words_to_text(words), [], words_to_tables(words), True
                        except Exception as exc:
                            log.append(f"⚠️ {name} p.{i} : échec OCR ({exc})")
                    else:
                        log.append(f"⚠️ {name} p.{i} : page scannée mais aucun moteur OCR disponible")
                pages.append(PageData(name, i, text, tables, ptables, used_ocr))
                if progress_cb:
                    progress_cb(i, total)
    except Exception as exc:
        log.append(f"❌ Impossible de lire {name} : {exc}")
    return pages


# =============================================================================
# CLASSIFICATION ET SEGMENTATION DES DOCUMENTS
# =============================================================================

KIND_PATTERNS = [
    ("MANIFESTE", re.compile(r"(manifest|manifeste|liste\s*de\s*chargement|loading\s*list|groupage\s*list|"
                             r"r[ée]capitulatif\s*groupage|borderau|bordereau\s*de\s*groupage|ladeliste)", re.I)),
    ("CMR", re.compile(r"(\bcmr\b|lettre\s*de\s*voiture|consignment\s*note|frachtbrief|carta\s*de\s*porte)", re.I)),
    ("PACKING", re.compile(r"(packing\s*list|liste\s*de\s*colisage|colisage|packliste|delivery\s*note|"
                           r"bon\s*de\s*livraison|lieferschein|albar[aá]n|despatch\s*note|dispatch\s*note|"
                           r"lista\s*de\s*empaque|packing\s*slip)", re.I)),
    ("FACTURE", re.compile(r"(invoice|facture|rechnung|factura|fattura|proforma)", re.I)),
]


def classify_page(page: PageData) -> str:
    lines = text_lines(page.text)
    top = "\n".join(lines[:12])
    # 1) titre du document dans le haut de page
    for kind, rx in KIND_PATTERNS:
        if kind == "FACTURE":
            # « Invoice No » sur un BL ne fait pas une facture : on cherche un titre
            if re.search(r"(^|\n)\s*(commercial\s+|proforma\s+|tax\s+)?(invoice|facture|rechnung|factura|fattura)\b"
                         r"(?!\s*(no|n°|nr|number|num))", top, re.I):
                return "FACTURE"
            continue
        if rx.search(top):
            # un BL qui comporte des prix et un titre « facture » reste une facture
            if kind == "PACKING" and re.search(r"(^|\n)\s*(commercial\s+)?(invoice|facture|rechnung)\b", top, re.I) \
                    and count_priced_lines(lines) >= 1:
                return "FACTURE"
            return kind
    # 2) contenu : présence de lignes valorisées + mot facture
    priced = count_priced_lines(lines)
    if priced and KIND_PATTERNS[3][1].search(page.text):
        return "FACTURE"
    for kind, rx in KIND_PATTERNS:
        if rx.search(page.text):
            return kind
    return "AUTRE"


def count_priced_lines(lines: list[str]) -> int:
    dec = detect_decimal("\n".join(lines))
    return sum(1 for ln in lines if parse_item_line(ln, dec) is not None)


def find_invoice_number(text: str) -> str:
    lines = text_lines(text)
    for i, ln in enumerate(lines):
        for m in INV_NO_RE.finditer(ln):
            val = m.group(1).strip(".-/")
            if DATE_RE.fullmatch(val) or len(val) < 3:
                continue
            return val
        # libellé seul ; valeur sur la ligne suivante (mise en page tabulaire)
        if INV_LABEL_ONLY_RE.search(ln) and i + 1 < len(lines):
            for tok in lines[i + 1].split():
                if re.search(r"\d", tok) and not DATE_RE.fullmatch(tok) and len(tok) >= 3:
                    return tok.strip(".,;")
    return ""


def find_bl_number(text: str) -> str:
    for rx in (BL_NO_RE, BL_NO_RE2):
        for m in rx.finditer(text):
            val = m.group(1).strip(".-/")
            if not DATE_RE.fullmatch(val) and len(val) >= 3:
                return val
    return ""


@dataclass
class Document:
    kind: str
    file: str
    pages: list = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.pages)

    @property
    def page_label(self) -> str:
        nums = [p.number for p in self.pages]
        return f"{nums[0]}" if len(nums) == 1 else f"{nums[0]}-{nums[-1]}"


def segment_documents(pages: list[PageData]) -> list[Document]:
    docs: list[Document] = []
    for page in pages:
        page.kind = classify_page(page)
        prev = docs[-1] if docs else None
        same_file = prev is not None and prev.file == page.file
        if page.kind == "AUTRE" and same_file:
            prev.pages.append(page)  # page de suite (page 2/3, conditions générales...)
            continue
        if same_file and prev.kind == page.kind:
            if page.kind == "FACTURE":
                n_prev, n_new = find_invoice_number(prev.text), find_invoice_number(page.text)
                if not n_new or n_new == n_prev or re.search(r"page\s*[2-9]|seite\s*[2-9]|\b[2-9]\s*/\s*\d\b",
                                                              page.text[:400], re.I) and not n_new:
                    prev.pages.append(page)
                    continue
            elif page.kind in ("PACKING", "CMR"):
                b_prev, b_new = find_bl_number(prev.text), find_bl_number(page.text)
                if not b_new or b_new == b_prev:
                    prev.pages.append(page)
                    continue
            else:
                prev.pages.append(page)
                continue
        docs.append(Document(page.kind, page.file, [page]))
    return docs


# =============================================================================
# EXTRACTION DES EN-TÊTES
# =============================================================================

def find_supplier(text: str) -> str:
    lines = text_lines(text)
    # 1) libellé explicite (Seller:, Fournisseur:, Shipper: ...)
    for i, ln in enumerate(lines[:40]):
        m = SUPPLIER_LABEL_RE.match(ln)
        if m:
            cand = m.group(1).strip(" :-")
            if not re.search(r"[A-Za-z]{3}", cand) and i + 1 < len(lines):
                cand = lines[i + 1]
            cand = re.split(r"\s{2,}|\s+(?:invoice|facture|date|n°|no\.)\b", cand, flags=re.I)[0]
            if cand and not CONSIGNEE_WORDS.search(cand):
                return cand.strip(" ,;:-")[:80]
    # 2) raison sociale avec forme juridique dans le haut du document
    for ln in lines[:25]:
        if CONSIGNEE_WORDS.search(ln) or KIND_PATTERNS[3][1].fullmatch(ln.strip()):
            continue
        m = LEGAL_RE.search(ln)
        if m and re.search(r"[A-Za-z]{3}", ln[:m.start()] or ""):
            return ln[:m.end()].strip(" ,;:-")[:80]
    # 3) première ligne « texte » qui n'est pas un titre
    for ln in lines[:8]:
        if re.search(r"[A-Za-z]{3}", ln) and not any(rx.search(ln) for _, rx in KIND_PATTERNS) \
                and not CONSIGNEE_WORDS.search(ln) and not re.search(r"page|date|tel|fax|www|@", ln, re.I):
            return ln[:80]
    return ""


def find_date(text: str) -> str:
    for ln in text_lines(text)[:60]:
        if re.search(r"(date|datum|fecha|data)\b", ln, re.I) and not re.search(r"(due|échéance|delivery|livraison|order)", ln, re.I):
            m = DATE_RE.search(ln)
            if m:
                return normalize_date(m.group(1))
    m = DATE_RE.search(text)
    return normalize_date(m.group(1)) if m else ""


def find_currency(text: str) -> str:
    counts = Counter()
    for cur in CURRENCIES:
        counts[cur] += len(re.findall(rf"\b{cur}\b", text))
    counts["EUR"] += text.count("€")
    counts["USD"] += len(re.findall(r"US\$|\$\s?\d", text))
    counts["GBP"] += text.count("£")
    cur, n = counts.most_common(1)[0]
    return cur if n else ""


def find_incoterm(text: str) -> str:
    m = re.search(r"\b(" + "|".join(INCOTERMS) + r")\b", text)
    return m.group(1) if m else ""


def find_weight(text: str, rx: re.Pattern, decimal: str) -> float | None:
    best = None
    for m in rx.finditer(text):
        v = parse_number(m.group(1).strip(), decimal)
        if v is None:
            v = parse_number(m.group(1).split()[0], decimal) if m.group(1).split() else None
        if v is None or v <= 0:
            continue
        is_total = "total" in m.group(0).lower()
        if best is None or is_total:
            best = v
    return best


def find_packages(text: str) -> int | None:
    best = None
    for m in PACKAGES_RE.finditer(text):
        val = m.group(1) or m.group(2)
        if val and 0 < int(val) < 10000:
            if best is None or "total" in m.group(0).lower():
                best = int(val)
    return best


# =============================================================================
# EXTRACTION DES LIGNES DE FACTURE
# =============================================================================

def map_header(cells: list, colmap: list) -> dict:
    """Associe les indices de colonnes d'un tableau aux champs internes."""
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(cells):
        label = strip_accents(clean_cell(cell)).lower().replace("\n", " ").strip()
        if not label or len(label) > 60:
            continue
        for key, pattern in colmap:
            if key in mapping:
                continue
            if re.search(pattern, label, re.I):
                mapping[key] = idx
                break
    return mapping


def extract_detail_fields(text: str, decimal: str) -> dict:
    """Champs complémentaires présents dans une désignation ou une ligne de détail."""
    out = {}
    m = HS_LABEL_RE.search(text)
    if m:
        ngp = normalize_ngp(m.group(1))
        if ngp:
            out["ngp"] = ngp
    m = ORIGIN_LABEL_RE.search(text)
    if m:
        out["origin"] = normalize_origin(m.group(1))
    m = ALSTOM_LABEL_RE.search(text)
    if m:
        out["alstom"] = m.group(1).strip(".,;")
    m = BRAND_LABEL_RE.search(text)
    if m:
        out["brand"] = m.group(1).strip()
    m = LINE_NET_RE.search(text)
    if m:
        out["net"] = parse_number(m.group(1), decimal)
    m = LINE_GROSS_RE.search(text)
    if m:
        out["gross"] = parse_number(m.group(1), decimal)
    return out


NGP_TOKEN_RE = re.compile(r"^(\d{4}[.\s]?\d{2}[.\s]?\d{2}(?:[.\s]?\d{2})?)$")


def _numeric_tokens(tokens: list[str], decimal: str):
    nums = []
    for i, tok in enumerate(tokens):
        raw = tok.strip(";:")
        if DATE_RE.fullmatch(raw) or re.fullmatch(r"\d{8,}", raw) or re.fullmatch(r"\d{4}\.\d{2}\.\d{2}(\.\d{2})?", raw):
            continue
        v = parse_number(raw, decimal)
        if v is not None and v >= 0:
            nums.append((i, v))
    return nums


def _find_triplet(nums):
    """Cherche (qté, prix, montant) tels que qté × prix ≈ montant."""
    candidates_k = nums[-4:][::-1]
    for k_idx, (k, a) in enumerate(candidates_k):
        if a <= 0:
            continue
        before_k = [n for n in nums if n[0] < k]
        for j, p in before_k[::-1]:
            if p <= 0:
                continue
            for i, q in [n for n in nums if n[0] < j][::-1]:
                if q <= 0:
                    continue
                for per in (1, 100, 1000):
                    calc = q * p / per
                    if abs(calc - a) <= max(0.011, 0.005 * a):
                        return i, j, k, q, p, a, per
    return None


def parse_item_line(line: str, decimal: str):
    """Analyse une ligne de texte « libre » ; retourne un dict si c'est une ligne article."""
    if TOTAL_WORDS.search(line) and not FEE_WORDS.search(line):
        return None
    tokens = line.split()
    if len(tokens) < 3:
        return None
    nums = _numeric_tokens(tokens, decimal)
    if len(nums) < 3:
        # séparateurs de milliers par espace : « 1 234,56 »
        merged = re.sub(r"(?<=\d) (?=\d{3}(?:[.,]\d+)?(?:\s|$))", "", line)
        if merged != line:
            return parse_item_line(merged, decimal)
        return None
    trip = _find_triplet(nums)
    if trip is None:
        merged = re.sub(r"(?<=\d) (?=\d{3}(?:[.,]\d+)(?:\s|$))", "", line)
        if merged != line:
            return parse_item_line(merged, decimal)
        return None
    i, j, k, q, p, a, per = trip
    before = tokens[:i]
    if not before:
        return None
    pos = ""
    if re.fullmatch(r"\d{1,5}[.)]?", before[0]) and len(before) > 1:
        pos = before[0].rstrip(".)")
        before = before[1:]
    ngp = ""
    rest = []
    for tok in before:
        if not ngp and NGP_TOKEN_RE.match(tok) and ("." in tok or len(tok) in (8, 10)) and rest:
            ngp = normalize_ngp(tok)
            continue
        rest.append(tok)
    ref = ""
    for idx, tok in enumerate(rest[:3]):
        if re.search(r"\d", tok) and len(tok) >= 3 and re.fullmatch(r"[A-Za-z0-9\-_./#]+", tok):
            ref = tok
            rest = rest[:idx] + rest[idx + 1:]
            break
    if not ref and rest and re.fullmatch(r"[A-Z0-9\-_./]{4,}", rest[0]):
        ref = rest.pop(0)
    desc = " ".join(rest).strip(" -|")
    if not desc and not ref:
        return None
    unit = ""
    middle = tokens[i + 1:j]
    for tok in middle:
        if tok.upper().strip(".") in UNITS:
            unit = tok.upper().strip(".")
    if not unit and before and before[-1].upper().strip(".") in UNITS:
        unit = before[-1].upper().strip(".")
        desc = desc.rsplit(" ", 1)[0] if desc.endswith(before[-1]) else desc
    after = tokens[k + 1:]
    origin = ""
    for tok in after + middle:
        t = tok.upper().strip(".,;")
        if not ngp and NGP_TOKEN_RE.match(tok):
            ngp = normalize_ngp(tok)
        elif t in ISO_COUNTRIES and not origin and t not in UNITS:
            origin = t
    return {
        "pos": pos, "ref": ref, "desc": desc, "qty": q, "unit": unit,
        "price": p, "amount": a, "ngp": ngp, "origin": origin,
        "price_per": per,
    }


def rows_from_tables(tables: list, decimal: str, carry_map: dict | None):
    """Extrait les lignes article des tableaux pdfplumber. Retourne (lignes, mapping)."""
    items = []
    mapping = carry_map
    for table in tables:
        if not table:
            continue
        start = 0
        local_map = None
        for r_idx, row in enumerate(table[:6]):
            m = map_header(row, LINE_COLMAP)
            if len(m) >= 3 and ("qty" in m or "amount" in m) and ("ref" in m or "desc" in m):
                local_map, start = m, r_idx + 1
                break
        if local_map is None:
            if mapping is None or len(table[0]) <= max(mapping.values()):
                continue
            local_map = mapping
        mapping = local_map
        for row in table[start:]:
            def cell(key):
                idx = local_map.get(key)
                return clean_cell(row[idx]) if idx is not None and idx < len(row) else ""
            joined = " ".join(clean_cell(c) for c in row)
            if not joined.strip():
                continue
            qty = parse_number(cell("qty").split("\n")[0].split(" ")[0], decimal) if cell("qty") else None
            amount = parse_number(cell("amount"), decimal)
            price = parse_number(cell("price"), decimal)
            # colonnes fusionnées par l'OCR / la mise en page : « 45,50 455,00 »
            for src in ("price", "amount"):
                parts = [parse_number(t, decimal) for t in cell(src).split()]
                parts = [v for v in parts if v is not None]
                if len(parts) >= 2 and (price is None or amount is None):
                    price, amount = (price if price is not None else parts[0]), parts[-1]
            ref = cell("ref").split("\n")[0]
            desc_full = cell("desc")
            desc = desc_full.split("\n")[0] if desc_full else ""
            if qty is None and amount is None:
                # ligne de suite de la désignation (même tableau) : rattachée à l'article précédent
                if items and desc_full and not ref and not TOTAL_WORDS.search(joined) \
                        and not any(cell(k) for k in ("qty", "price", "amount", "pos")):
                    items[-1]["desc_full"] = f"{items[-1].get('desc_full') or ''}\n{desc_full}".strip()
                continue
            if not (ref or desc):
                continue
            if TOTAL_WORDS.search(f"{ref} {desc}") and not FEE_WORDS.search(f"{ref} {desc}"):
                continue
            if amount is None and qty is not None and price is not None:
                amount = round(qty * price, 2)
            if price is None and qty and amount is not None:
                price = round(amount / qty, 4)
            details = extract_detail_fields(joined, decimal)
            unit = cell("unit")
            if not unit:
                um = re.search(r"\b(" + "|".join(sorted(UNITS, key=len, reverse=True)) + r")\b", cell("qty").upper())
                unit = um.group(1) if um else ""
            items.append({
                "pos": cell("pos"), "ref": ref, "desc": desc or desc_full,
                "desc_full": desc_full, "po": cell("po"), "unit_raw": cell("unit"),
                "alstom": cell("alstom") or details.get("alstom", ""),
                "brand": cell("brand") or details.get("brand", ""),
                "qty": qty, "unit": unit.upper(), "price": price, "amount": amount,
                "net": parse_number(cell("net"), decimal) if cell("net") else details.get("net"),
                "gross": parse_number(cell("gross"), decimal) if cell("gross") else details.get("gross"),
                "packages": parse_number(cell("packages"), decimal) if cell("packages") else None,
                "origin": normalize_origin(cell("origin")) if cell("origin") else details.get("origin", ""),
                "ngp": normalize_ngp(cell("ngp")) if cell("ngp") else details.get("ngp", ""),
            })
    return items, mapping


def rows_from_text(lines: list[str], decimal: str):
    """Analyse ligne à ligne (PDF sans tableau ou OCR). Retourne (lignes, NGP d'en-tête libres)."""
    items = []
    free_ngps = []
    last = None
    for ln in lines:
        item = parse_item_line(ln, decimal)
        if item:
            details = extract_detail_fields(ln, decimal)
            item.setdefault("alstom", details.get("alstom", ""))
            item.setdefault("brand", details.get("brand", ""))
            item["net"] = details.get("net")
            item["gross"] = details.get("gross")
            item["packages"] = None
            if not item["ngp"] and details.get("ngp"):
                item["ngp"] = details["ngp"]
            if not item["origin"] and details.get("origin"):
                item["origin"] = details["origin"]
            items.append(item)
            last = item
            continue
        details = extract_detail_fields(ln, decimal)
        if (last is not None and not details and not TOTAL_WORDS.search(ln) and len(ln) < 90
                and len(last.setdefault("desc_extra", [])) < 3
                and sum(c.isalpha() for c in ln) >= 0.6 * len(ln.replace(" ", ""))):
            last["desc_extra"].append(ln)  # suite de désignation sur la ligne suivante
            continue
        if last is not None and details and not TOTAL_WORDS.search(ln):
            # ligne de détail rattachée à l'article précédent
            for key, target in (("ngp", "ngp"), ("origin", "origin"), ("alstom", "alstom"),
                                ("brand", "brand"), ("net", "net"), ("gross", "gross")):
                if details.get(key) and not last.get(target):
                    last[target] = details[key]
            continue
        if details.get("ngp"):
            free_ngps.append(details["ngp"])
        if TOTAL_WORDS.search(ln):
            last = None
    return items, free_ngps


def find_header_ngp(text: str, items: list) -> str:
    line_ngps = {it.get("ngp") for it in items if it.get("ngp")}
    found = []
    for m in HS_LABEL_RE.finditer(text):
        n = normalize_ngp(m.group(1))
        if n:
            found.append(n)
    free = [n for n in found if n not in line_ngps]
    if free:
        return free[0]
    if found:
        return Counter(found).most_common(1)[0][0]
    return ""


@dataclass
class Invoice:
    doc: Document
    supplier: str = ""
    number: str = ""
    date: str = ""
    currency: str = ""
    incoterm: str = ""
    bl: str = ""
    header_ngp: str = ""
    origin: str = ""
    net_total: float | None = None
    gross_total: float | None = None
    packages: int | None = None
    items: list = field(default_factory=list)
    packing_lists: list = field(default_factory=list)


def parse_invoice(doc: Document) -> Invoice:
    text = doc.text
    dec = detect_decimal(text)
    inv = Invoice(doc=doc)
    inv.supplier = find_supplier(text)
    inv.number = find_invoice_number(text)
    inv.date = find_date(text)
    inv.currency = find_currency(text)
    inv.incoterm = find_incoterm(text)
    inv.bl = find_bl_number(text)
    inv.net_total = find_weight(text, TOTAL_NET_RE, dec)
    inv.gross_total = find_weight(text, TOTAL_GROSS_RE, dec)
    inv.packages = find_packages(text)
    m = re.search(r"(?:country\s*of\s*origin|pays\s*d'?\s*origine|origin\w*\s*of\s*goods|origine\s*des\s*marchandises|"
                  r"ursprungsland)\s*[:#.]?\s*([A-Za-z][A-Za-z\- ]{1,25})", text, re.I)
    if m:
        inv.origin = normalize_origin(m.group(1))

    items: list = []
    carry = None
    for page in doc.pages:
        # Candidats : tableaux à filets, tableaux reconstitués, analyse ligne à ligne.
        # On retient celui qui fournit le plus de lignes valorisées (montant présent).
        candidates = []
        for tset in page.table_sets():
            is_ruled = tset is page.tables
            found, new_carry = rows_from_tables(tset, dec, carry if is_ruled else None)
            if is_ruled:
                carry = new_carry
            candidates.append(found)
        candidates.append(rows_from_text(text_lines(page.text), dec)[0])

        def score(found):
            valued = sum(1 for it in found if it.get("amount") is not None)
            return valued, -sum(1 for it in found if it.get("amount") is None)
        page_items = max(candidates, key=score)  # à égalité : le premier (tableau) l'emporte
        items.extend(page_items)
    inv.items = items
    inv.header_ngp = find_header_ngp(text, items)
    return inv


# =============================================================================
# PACKING LISTS, CMR ET MANIFESTE
# =============================================================================

def row_cell(row: list, header: dict, key: str) -> str:
    idx = header.get(key)
    return clean_cell(row[idx]) if idx is not None and idx < len(row) else ""


def iter_table_rows(tables: list, colmap: list, accept):
    """Parcourt les lignes de données des tableaux dont l'en-tête satisfait `accept(mapping)`."""
    for table in tables:
        for r_idx, row in enumerate(table[:6]):
            m = map_header(row, colmap)
            if accept(m):
                for data_row in table[r_idx + 1:]:
                    if any(clean_cell(c) for c in data_row):
                        yield m, data_row
                break


@dataclass
class PackingList:
    doc: Document
    number: str = ""
    invoice_ref: str = ""
    supplier: str = ""
    net_total: float | None = None
    gross_total: float | None = None
    packages: int | None = None
    lines: list = field(default_factory=list)  # {ref, qty, net, gross, packages}


def parse_packing(doc: Document) -> PackingList:
    text = doc.text
    dec = detect_decimal(text)
    pl = PackingList(doc=doc)
    pl.number = find_bl_number(text)
    pl.invoice_ref = find_invoice_number(text)
    pl.supplier = find_supplier(text)
    pl.net_total = find_weight(text, TOTAL_NET_RE, dec)
    pl.gross_total = find_weight(text, TOTAL_GROSS_RE, dec)
    pl.packages = find_packages(text)
    for page in doc.pages:
        for tset in page.table_sets():  # tableaux à filets d'abord, puis reconstitués
            found = []
            for header, row in iter_table_rows(tset, LINE_COLMAP,
                                               lambda m: ("ref" in m or "alstom" in m) and ("net" in m or "gross" in m)):
                ref = row_cell(row, header, "ref").split("\n")[0] or row_cell(row, header, "alstom")
                if not ref or TOTAL_WORDS.search(ref):
                    continue
                found.append({
                    "ref": ref,
                    "qty": parse_number(row_cell(row, header, "qty"), dec),
                    "net": parse_number(row_cell(row, header, "net"), dec),
                    "gross": parse_number(row_cell(row, header, "gross"), dec),
                    "packages": parse_number(row_cell(row, header, "packages"), dec),
                })
            if found:
                pl.lines.extend(found)
                break
    return pl


def parse_manifest(doc: Document) -> list[dict]:
    dec = detect_decimal(doc.text)
    rows = []
    for page in doc.pages:
        for tset in page.table_sets():  # tableaux à filets d'abord, puis reconstitués
            found = []
            for header, row in iter_table_rows(tset, MANIFEST_COLMAP,
                                               lambda m: "supplier" in m and ("value" in m or "gross" in m)):
                def cell(key, row=row, header=header):
                    return row_cell(row, header, key)
                supplier = cell("supplier").replace("\n", " ")
                if not supplier or TOTAL_WORDS.search(supplier):
                    continue
                found.append({
                    "Page PDF": page.number,
                    "Shipment": cell("shipment").replace("\n", " "),
                    "Expéditeur": supplier,
                    "Poids (kg)": parse_number(cell("gross"), dec),
                    "Colis": parse_number(cell("packages"), dec),
                    "Goods Value": parse_number(cell("value"), dec),
                    "Devise": find_currency(cell("currency") + " " + cell("value")),
                })
            if found:
                rows.extend(found)
                break
    return rows


# =============================================================================
# CONSTRUCTION DES LIGNES + VENTILATION DES POIDS
# =============================================================================

def link_packing_lists(invoices: list[Invoice], packings: list[PackingList], docs: list[Document]):
    order = {id(d): i for i, d in enumerate(docs)}
    for pl in packings:
        target = None
        if pl.invoice_ref:
            target = next((inv for inv in invoices if inv.number and
                           compact_ref(inv.number) == compact_ref(pl.invoice_ref)), None)
        if target is None and pl.supplier:
            same = [inv for inv in invoices if company_similarity(inv.supplier, pl.supplier) >= 0.8]
            if len(same) == 1:
                target = same[0]
        if target is None and invoices:
            pos = order[id(pl.doc)]
            same_file = [inv for inv in invoices if inv.doc.file == pl.doc.file]
            pool = same_file or invoices
            before = [inv for inv in pool if order[id(inv.doc)] < pos]
            target = before[-1] if before else min(pool, key=lambda inv: abs(order[id(inv.doc)] - pos))
        if target is not None:
            target.packing_lists.append(pl)
            if not target.bl and pl.number:
                target.bl = pl.number


def allocate(total: float, weights: list[float], decimals: int = 3) -> list[float]:
    s = sum(weights)
    if not s:
        weights = [1.0] * len(weights)
        s = float(len(weights))
    parts = [round(total * w / s, decimals) for w in weights]
    if parts:
        parts[-1] = round(total - sum(parts[:-1]), decimals)
    return parts


def build_invoice_rows(inv: Invoice, basis: str) -> list[dict]:
    rows = []
    for n, it in enumerate(inv.items, start=1):
        desc = clean_cell(it.get("desc"))
        is_fee = bool(FEE_WORDS.search(f"{desc} {it.get('ref', '')}"))
        rows.append({
            "Client": inv.supplier,
            "Job Ref manifest": "",
            "N° facture": inv.number,
            "Date": inv.date,
            "Item": clean_cell(it.get("pos")) or str(n),
            "Réf article": clean_cell(it.get("ref")),
            "Désignation": desc,
            "Origine": clean_cell(it.get("origin")) or inv.origin,
            "Qté": it.get("qty"),
            "Prix unitaire": it.get("price"),
            "Valeur HT": it.get("amount"),
            "NGP/code douanier (facture)": it.get("ngp") or inv.header_ngp,
            "NGP (REF_PAR_FRS)": "",
            "Onglet REF_PAR_FRS": "",
            "Contrôle NGP": "",
            "Remarque": "",
            "Page PDF": inv.doc.pages[0].number if inv.doc.pages else 0,
            # internal fields (not in COLUMNS)
            "_réf_alstom": clean_cell(it.get("alstom")),
            "_devise": inv.currency,
            "_poids_net": it.get("net"),
            "_poids_brut": it.get("gross"),
            "_colis": it.get("packages"),
            "_is_fee": is_fee,
            "_line_ngp": bool(it.get("ngp")),
        })
    if not rows:
        return rows
    notes = []
    goods = [r for r in rows if not r["_is_fee"]] or rows

    # --- 1) Poids par ligne issus de la packing list (même référence, consommés dans l'ordre)
    pl_lines = [dict(l) for pl in inv.packing_lists for l in pl.lines]
    for r in goods:
        for key_line in ("Réf article", "_réf_alstom"):
            key = compact_ref(r[key_line])
            if not key:
                continue
            match = next((pl for pl in pl_lines if not pl.get("_used") and compact_ref(pl["ref"]) == key), None)
            if match:
                match["_used"] = True
                if r["_poids_net"] is None and match.get("net") is not None:
                    r["_poids_net"] = match["net"]
                if r["_poids_brut"] is None and match.get("gross") is not None:
                    r["_poids_brut"] = match["gross"]
                if r["_colis"] is None and match.get("packages") is not None:
                    r["_colis"] = match["packages"]
                r["Remarque"] = "Poids issus de la Packing List"
                break

    # --- 2) Totaux par facture (facture, sinon packing list)
    def first(*vals):
        return next((v for v in vals if v is not None), None)
    net_total = first(inv.net_total, *(pl.net_total for pl in inv.packing_lists))
    gross_total = first(inv.gross_total, *(pl.gross_total for pl in inv.packing_lists))
    pkg_total = first(inv.packages, *(pl.packages for pl in inv.packing_lists))

    for col, total, label in (("_poids_net", net_total, "net"), ("_poids_brut", gross_total, "brut")):
        missing = [r for r in goods if r[col] is None]
        if not missing or total is None:
            continue
        known = sum(r[col] for r in goods if r[col] is not None)
        remaining = total - known
        if remaining <= 0:
            continue
        if basis == "Quantité":
            w = [abs(r["Qté"] or 0) for r in missing]
            base = "quantité"
        else:
            w = [abs(r["Valeur HT"] or 0) for r in missing]
            base = "valeur HT"
        for r, v in zip(missing, allocate(remaining, w)):
            r[col] = v
            msg = f"Poids {label} ventilé au prorata {base}"
            r["Remarque"] = f"{r['Remarque']} | {msg}".strip(" |")
    # Poids brut absent mais poids net présent : aucune invention, on le signale
    if gross_total is None and any(r["_poids_brut"] is None for r in goods):
        notes.append("Poids brut non trouvé")
    if net_total is None and any(r["_poids_net"] is None for r in goods):
        notes.append("Poids net non trouvé")
    # --- 3) Colis : jamais dupliqués -> total facture porté sur la 1ère ligne
    if pkg_total is not None and all(r["_colis"] is None for r in rows):
        rows[0]["_colis"] = pkg_total
        rows[0]["Remarque"] = f"{rows[0]['Remarque']} | Colis global facture".strip(" |")
    if notes:
        for r in goods:
            if r["_poids_net"] is None or r["_poids_brut"] is None:
                r["Remarque"] = f"{r['Remarque']} | {' / '.join(notes)}".strip(" |")
    for r in rows:
        if not r["_line_ngp"] and r["NGP/code douanier (facture)"]:
            r["Remarque"] = f"{r['Remarque']} | NGP en-tête facture".strip(" |")
        r.pop("_line_ngp", None)
        r.pop("_is_fee", None)
    return rows


# =============================================================================
# FICHIER EXCEL DE RÉFÉRENCE ET MATCHING NGP
# =============================================================================

def load_reference(data: bytes, filename: str, log: list) -> dict:
    """Retourne {'exact': {ref: ngp}, 'compact': {ref: ngp}, 'ocr': {ref: ngp}, 'sheet_names': {ref: sheet}, 'count': n}."""
    exact, compact, ocrk, sheet_names = {}, {}, {}, {}
    try:
        if filename.lower().endswith(".csv"):
            sheets = {"CSV": pd.read_csv(io.BytesIO(data), header=None, dtype=str, sep=None, engine="python")}
        else:
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=str)
    except Exception as exc:
        log.append(f"❌ Lecture du fichier de référence impossible : {exc}")
        return {"exact": exact, "compact": compact, "ocr": ocrk, "count": 0}
    for sheet, df in sheets.items():
        if df.empty:
            continue
        header_idx, mapping, ref_cols = None, {}, []
        for r in range(min(25, len(df))):
            cells = list(df.iloc[r].values)
            m = map_header(cells, REF_SHEET_COLMAP)
            if "ngp" in m and ("ref" in m or "alstom" in m):
                header_idx, mapping = r, m
                # toutes les colonnes « référence » de la ligne d'en-tête
                for c_idx, c in enumerate(cells):
                    lab = strip_accents(clean_cell(c)).lower()
                    if c_idx != m["ngp"] and lab and re.search(REF_SHEET_COLMAP[1][1] + "|" + REF_SHEET_COLMAP[2][1], lab) \
                            and not re.search(r"(desig|descr|libell|ngp|\bhs\b)", lab):
                        ref_cols.append(c_idx)
                break
        if header_idx is None:
            # pas d'en-tête reconnu : colonne NGP = la plus riche en codes à 8-10 chiffres
            scores = {c: df[c].map(lambda v: len(normalize_ngp(v)) in (8, 10)).sum() for c in df.columns}
            ngp_col = max(scores, key=scores.get)
            if scores[ngp_col] == 0:
                log.append(f"⚠️ Feuille « {sheet} » ignorée (aucune colonne NGP détectée)")
                continue
            header_idx, mapping = -1, {"ngp": ngp_col}
            ref_cols = [c for c in df.columns if c != ngp_col][:1]
        body = df.iloc[header_idx + 1:]
        added = 0
        for _, row in body.iterrows():
            ngp = normalize_ngp(row.iloc[mapping["ngp"]])
            if not ngp:
                continue
            for c in ref_cols:
                key = norm_ref(row.iloc[c])
                if not key:
                    continue
                exact.setdefault(key, ngp)
                sheet_names.setdefault(key, sheet)
                ck = compact_ref(key)
                if ck:
                    compact.setdefault(ck, ngp)
                    ocrk.setdefault(ocr_ref(key), ngp)
                added += 1
        log.append(f"✅ Référentiel « {sheet} » : {added} références chargées")
    return {"exact": exact, "compact": compact, "ocr": ocrk, "sheet_names": sheet_names, "count": len(exact)}


def lookup_ngp(refdb: dict, *refs) -> tuple[str, str]:
    for label, ref in refs:
        key = norm_ref(ref)
        if key and key in refdb["exact"]:
            return refdb["exact"][key], label
    for label, ref in refs:
        key = compact_ref(ref)
        if key and key in refdb["compact"]:
            return refdb["compact"][key], label
    for label, ref in refs:
        key = ocr_ref(ref)
        if len(key) >= 5 and key in refdb.get("ocr", {}):
            return refdb["ocr"][key], f"{label} (tolérance OCR)"
    return "", ""


def control_ngp(ngp_ref: str, code_fact: str) -> str:
    """Réplique Python de la formule Excel « Contrôle NGP »."""
    if not ngp_ref or ngp_ref == NON_TROUVE or not code_fact:
        return "À vérifier"
    return "OK" if str(ngp_ref)[:8] == str(code_fact)[:8] else "ÉCART"


def apply_matching(df: pd.DataFrame, refdb: dict) -> pd.DataFrame:
    df = df.copy()
    ngp_ref_list, onglet_list, ctrl, remarks = [], [], [], []
    sheet_names = refdb.get("sheet_names", {})
    for _, r in df.iterrows():
        code_fact = normalize_ngp(r.get("NGP/code douanier (facture)", ""))
        rem = clean_cell(r.get("Remarque", ""))
        rem = " | ".join(p for p in rem.split(" | ") if p and p != REMARQUE_NOUVELLE_REF
                         and not p.startswith("Trouvé via"))
        ref_art = r.get("Réf article", "")
        ref_alstom = r.get("_réf_alstom", "")
        ngp, via = lookup_ngp(refdb, ("Réf article", ref_art), ("_réf_alstom", ref_alstom))
        if ngp:
            if via != "Réf article":
                rem = f"{rem} | Trouvé via {via}".strip(" |")
            # get sheet name for this reference
            found_key = norm_ref(ref_art) if via == "Réf article" else norm_ref(ref_alstom)
            onglet = sheet_names.get(found_key, "")
        else:
            ngp = NON_TROUVE
            onglet = ""
            rem = f"{rem} | {REMARQUE_NOUVELLE_REF}".strip(" |")
        ngp_ref_list.append(ngp)
        onglet_list.append(onglet)
        ctrl.append(control_ngp(ngp, code_fact))
        remarks.append(rem)
    df["NGP/code douanier (facture)"] = df["NGP/code douanier (facture)"].map(
        lambda v: normalize_ngp(v) if pd.notna(v) else "")
    df["NGP (REF_PAR_FRS)"] = ngp_ref_list
    df["Onglet REF_PAR_FRS"] = onglet_list
    df["Contrôle NGP"] = ctrl
    df["Remarque"] = remarks
    # ensure all COLUMNS exist
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df[COLUMNS]


# =============================================================================
# PIPELINE COMPLET
# =============================================================================

def run_pipeline(pdf_files: list, ref_file, ocr: OCREngine, basis: str, progress) -> dict:
    log: list[str] = []

    # ---------- Étape 1 : lecture OCR + extraction
    progress.progress(0.02, text="Étape 1/3 : Lecture OCR et extraction des données PDF...")
    all_pages: list[PageData] = []
    n_files = len(pdf_files)
    for f_idx, f in enumerate(pdf_files):
        data = f.getvalue()

        def cb(i, total, f_idx=f_idx, name=f.name):
            frac = (f_idx + i / max(total, 1)) / n_files
            progress.progress(0.02 + 0.55 * frac,
                              text=f"Étape 1/3 : Lecture OCR et extraction des données PDF... ({name} p.{i}/{total})")
        all_pages.extend(read_pdf_pages(f.name, data, ocr, log, cb))
    n_ocr = sum(p.ocr for p in all_pages)
    log.append(f"📄 {len(all_pages)} pages lues ({n_ocr} via OCR)")

    docs = segment_documents(all_pages)
    invoices = [parse_invoice(d) for d in docs if d.kind == "FACTURE"]
    packings = [parse_packing(d) for d in docs if d.kind == "PACKING"]
    manifest_rows = [r for d in docs if d.kind == "MANIFESTE" for r in parse_manifest(d)]
    link_packing_lists(invoices, packings, docs)

    # Fournisseur manquant : packing list liée, sinon fichier
    for inv in invoices:
        if not inv.supplier:
            inv.supplier = next((pl.supplier for pl in inv.packing_lists if pl.supplier), "") or "À compléter"
        if not inv.items:
            log.append(f"⚠️ Facture {inv.number or '?'} ({inv.supplier}) : aucune ligne article détectée "
                       f"({inv.doc.file} p.{inv.doc.page_label}) — saisie manuelle possible")

    docs_df = pd.DataFrame([{
        "Fichier": d.file, "Pages": d.page_label, "Type": d.kind,
        "N° détecté": find_invoice_number(d.text) if d.kind == "FACTURE" else find_bl_number(d.text),
        "Fournisseur": find_supplier(d.text) if d.kind in ("FACTURE", "PACKING") else "", "OCR": "Oui" if any(p.ocr for p in d.pages) else "Non",
    } for d in docs])

    # ---------- Étape 2 : réconciliation + matching NGP
    progress.progress(0.62, text="Étape 2/3 : Réconciliation et matching NGP avec l'Excel...")
    refdb = load_reference(ref_file.getvalue(), ref_file.name, log) if ref_file else {"exact": {}, "compact": {}, "ocr": {}, "sheet_names": {}, "count": 0}
    rows = [r for inv in invoices for r in build_invoice_rows(inv, basis)]
    # columns to initialise (COLUMNS minus matching-filled ones, plus internal keys)
    init_cols = [c for c in COLUMNS if c not in ("NGP (REF_PAR_FRS)", "Contrôle NGP", "Onglet REF_PAR_FRS")]
    internal_cols = ["_réf_alstom", "_devise", "_poids_net", "_poids_brut", "_colis"]
    lines_df = pd.DataFrame(rows, columns=init_cols + internal_cols) if rows else pd.DataFrame(columns=init_cols + internal_cols)
    for c in NUMERIC_COLUMNS:
        if c in lines_df.columns:
            lines_df[c] = pd.to_numeric(lines_df[c], errors="coerce")
    if "Page PDF" in lines_df.columns:
        lines_df["Page PDF"] = pd.to_numeric(lines_df["Page PDF"], errors="coerce").fillna(0).astype(int)
    for c in lines_df.columns:
        if c not in NUMERIC_COLUMNS and c not in ("Page PDF",) and not c.startswith("_"):
            lines_df[c] = lines_df[c].map(clean_cell)
    # fill Job Ref manifest: comma-join Shipment values from manifest for each supplier
    if manifest_rows:
        sup_to_shipments: dict = {}
        for mr in manifest_rows:
            sup = mr.get("Expéditeur", "")
            ship = mr.get("Shipment", "")
            if sup and ship:
                sup_to_shipments.setdefault(sup, [])
                if ship not in sup_to_shipments[sup]:
                    sup_to_shipments[sup].append(ship)
        if "Client" in lines_df.columns and "Job Ref manifest" in lines_df.columns:
            lines_df["Job Ref manifest"] = lines_df["Client"].map(
                lambda s: ", ".join(sup_to_shipments.get(s, [])))
    manifest_df = pd.DataFrame(manifest_rows, columns=MANIFEST_COLUMNS) if manifest_rows else pd.DataFrame(columns=MANIFEST_COLUMNS)
    for c in ("Poids (kg)", "Colis", "Goods Value"):
        if c in manifest_df.columns:
            manifest_df[c] = pd.to_numeric(manifest_df[c], errors="coerce")
    if manifest_df.empty:
        log.append("ℹ️ Aucun tableau de manifeste exploitable : saisissez les Goods Value dans l'onglet « Manifeste »")
    progress.progress(0.85, text="Étape 3/3 : Génération du classeur Excel à 5 feuilles...")
    return {"lines": lines_df, "manifest": manifest_df, "docs": docs_df, "refdb": refdb, "log": log}


# =============================================================================
# GÉNÉRATION DU CLASSEUR EXCEL
# =============================================================================

HEADER_FILL = PatternFill("solid", start_color="1F3864", end_color="1F3864")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name="Arial", size=10)
BOLD_FONT = Font(name="Arial", size=10, bold=True)
RED_FONT = Font(name="Arial", size=10, bold=True, color="C00000")
INPUT_FONT = Font(name="Arial", size=10, color="0000FF")
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
INPUT_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_OK = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
FILL_KO = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
FILL_WARN = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")


def style_header(ws, row: int, ncols: int, height: float = 30):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.border = HEADER_FILL, HEADER_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


def assign_suppliers_to_manifest(manifest: pd.DataFrame, lines: pd.DataFrame) -> pd.DataFrame:
    """Rattache chaque ligne du manifeste au fournisseur des factures (nom similaire)."""
    suppliers = [s for s in lines["Client"].dropna().unique() if s]
    out = manifest.copy()
    mapped = []
    for _, r in out.iterrows():
        sup = ""
        exp = clean_cell(r.get("Expéditeur", ""))
        if sup == "" and suppliers and exp:
            best = max(suppliers, key=lambda s: company_similarity(s, exp))
            if company_similarity(best, exp) >= 0.6:
                sup = best
        mapped.append(sup or exp)
    out["_supplier"] = mapped
    return out


def build_workbook(lines: pd.DataFrame, manifest: pd.DataFrame, docs: pd.DataFrame = None) -> bytes:
    wb = Workbook()
    YELLOW_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")

    # ------------------------------------------------------------------ TOUS_CLIENTS
    ws = wb.active
    ws.title = "TOUS_CLIENTS"
    ws.append(COLUMNS)
    # header styling: 1F3864 fill, white bold Arial 10, vertically centered, wrap_text, height 35
    for c_idx in range(1, len(COLUMNS) + 1):
        hcell = ws.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 35

    col_idx = {name: i + 1 for i, name in enumerate(COLUMNS)}
    col_ltr = {name: get_column_letter(i + 1) for i, name in enumerate(COLUMNS)}
    num_formats_tc = {
        "Qté": "#,##0.###",
        "Prix unitaire": "#,##0.00",
        "Valeur HT": "#,##0.00",
    }
    text_format_cols = {"N° facture", "Réf article", "NGP/code douanier (facture)"}

    n = len(lines)
    for i, (_, r) in enumerate(lines.iterrows(), start=2):
        for j, name in enumerate(COLUMNS, start=1):
            val = r.get(name)
            if name == "Contrôle NGP":
                m_col = col_ltr["NGP (REF_PAR_FRS)"]
                l_col = col_ltr["NGP/code douanier (facture)"]
                val = (f'=IF(OR({m_col}{i}="{NON_TROUVE}",{m_col}{i}="",{l_col}{i}=""),"À vérifier",'
                       f'IF(LEFT({m_col}{i},8)=LEFT({l_col}{i},8),"OK","ÉCART"))')
            elif name in NUMERIC_COLUMNS:
                val = None if (val is None or (isinstance(val, float) and math.isnan(val))) else float(val)
            elif name == "Page PDF":
                try:
                    val = int(val) if val is not None and not (isinstance(val, float) and math.isnan(val)) else None
                except (ValueError, TypeError):
                    val = None
            else:
                val = clean_cell(val) or None
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = BODY_FONT
            if name in num_formats_tc:
                cell.number_format = num_formats_tc[name]
            if name in text_format_cols:
                cell.number_format = "@"
            # red font for NON TROUVÉ in Réf article column
            if name == "Réf article" and val == NON_TROUVE:
                cell.font = RED_FONT
            # yellow fill for NGP (REF_PAR_FRS) and Onglet REF_PAR_FRS
            if name in ("NGP (REF_PAR_FRS)", "Onglet REF_PAR_FRS"):
                cell.fill = YELLOW_FILL

    last = max(n + 1, 2)
    tc_widths = {
        "Client": 30, "Job Ref manifest": 24, "N° facture": 14, "Date": 12,
        "Item": 8, "Réf article": 20, "Désignation": 42, "Origine": 14,
        "Qté": 10, "Prix unitaire": 13, "Valeur HT": 13,
        "NGP/code douanier (facture)": 18, "NGP (REF_PAR_FRS)": 14,
        "Onglet REF_PAR_FRS": 16, "Contrôle NGP": 14, "Remarque": 50, "Page PDF": 9,
    }
    for name, w in tc_widths.items():
        if name in col_ltr:
            ws.column_dimensions[col_ltr[name]].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    # Ranges used in CONTROLE / CONTROLE_FACTURES formulas (safe up to row 5000)
    MAX_ROW = 5000
    R_CLIENT = f"TOUS_CLIENTS!$A$2:$A${MAX_ROW}"
    R_FACTURE = f"TOUS_CLIENTS!$C$2:$C${MAX_ROW}"
    R_HT = f"TOUS_CLIENTS!$K$2:$K${MAX_ROW}"
    R_CTRL = f"TOUS_CLIENTS!$O$2:$O${MAX_ROW}"

    # ------------------------------------------------------------------ CONTROLE (8 cols)
    wc = wb.create_sheet("CONTROLE")
    ctrl_headers = ["Client", "Σ Valeur HT extraite", "Goods Value manifest",
                    "Écart", "Nb lignes extraites", "Nb À vérifier", "Nb ÉCART", "Statut"]
    wc.append(ctrl_headers)
    for c_idx in range(1, len(ctrl_headers) + 1):
        hcell = wc.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wc.row_dimensions[1].height = 24

    # collect unique suppliers using matched names (so manifest long names map to invoice short names)
    man_matched = assign_suppliers_to_manifest(manifest, lines) if not manifest.empty else pd.DataFrame(columns=["_supplier", "Goods Value"])
    man_sups = list(dict.fromkeys(man_matched["_supplier"].dropna().tolist())) if not man_matched.empty else []
    line_sups = list(lines["Client"].dropna().unique()) if not lines.empty else []
    all_sups = list(dict.fromkeys(man_sups + [s for s in line_sups if s not in man_sups]))

    for i, sup in enumerate(all_sups, start=2):
        msub = man_matched[man_matched["_supplier"] == sup] if not man_matched.empty else pd.DataFrame()
        gv = msub["Goods Value"].dropna().sum() if not msub.empty and "Goods Value" in msub.columns and msub["Goods Value"].notna().any() else None
        # determine statut
        nb_inv_found = int(lines[lines["Client"] == sup]["N° facture"].replace("", pd.NA).dropna().nunique()) if not lines.empty else 0
        if nb_inv_found > 0:
            statut_txt = "Toutes les factures extraites."
        else:
            statut_txt = "⚠ AUCUNE facture extraite pour l'instant"
        values = [
            sup,
            f"=SUMIFS({R_HT},{R_CLIENT},A{i})",
            gv,
            f"=ROUND(B{i}-C{i},2)",
            f"=COUNTIFS({R_CLIENT},A{i})",
            f'=COUNTIFS({R_CLIENT},A{i},{R_CTRL},"À vérifier")',
            f'=COUNTIFS({R_CLIENT},A{i},{R_CTRL},"ÉCART")',
            statut_txt,
        ]
        for j, v in enumerate(values, start=1):
            cell = wc.cell(row=i, column=j, value=v)
            cell.font = BODY_FONT
            if j == 3:  # Goods Value manifest — always yellow (user-editable)
                cell.fill = YELLOW_FILL
        for j, fmt in [(2, "#,##0.00"), (3, "#,##0.00"), (4, "#,##0.00"),
                       (5, "#,##0"), (6, "#,##0"), (7, "#,##0")]:
            wc.cell(row=i, column=j).number_format = fmt

    # TOTAL row
    tot_row = len(all_sups) + 2
    wc.cell(row=tot_row, column=1, value="TOTAL")
    for j in range(2, 8):
        L = get_column_letter(j)
        if j == 4:  # Écart = SUM
            wc.cell(row=tot_row, column=j, value=f"=SUM({L}2:{L}{tot_row - 1})")
        elif j == 3:  # Goods Value = SUM
            wc.cell(row=tot_row, column=j, value=f"=SUM({L}2:{L}{tot_row - 1})")
        else:
            wc.cell(row=tot_row, column=j, value=f"=SUM({L}2:{L}{tot_row - 1})")
    for j in range(1, len(ctrl_headers) + 1):
        c = wc.cell(row=tot_row, column=j)
        c.font = BOLD_FONT
    for j, fmt in [(2, "#,##0.00"), (3, "#,##0.00"), (4, "#,##0.00"),
                   (5, "#,##0"), (6, "#,##0"), (7, "#,##0")]:
        wc.cell(row=tot_row, column=j).number_format = fmt

    ctrl_widths = [38, 18, 14, 14, 18, 12, 10, 60]
    for j, w in enumerate(ctrl_widths, start=1):
        wc.column_dimensions[get_column_letter(j)].width = w
    wc.freeze_panes = "A2"
    wc.auto_filter.ref = f"A1:{get_column_letter(len(ctrl_headers))}1"

    # ------------------------------------------------------------------ CONTROLE_FACTURES (5 cols)
    wcf = wb.create_sheet("CONTROLE_FACTURES")
    cf_headers = ["Client", "N° facture", "Σ lignes extraites", "Total imprimé facture", "Écart"]
    wcf.append(cf_headers)
    for c_idx in range(1, len(cf_headers) + 1):
        hcell = wcf.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wcf.row_dimensions[1].height = 24

    # one row per unique invoice
    if not lines.empty and "N° facture" in lines.columns and "Client" in lines.columns:
        inv_pairs = lines[["Client", "N° facture"]].drop_duplicates().dropna(subset=["N° facture"])
        inv_pairs = inv_pairs[inv_pairs["N° facture"] != ""]
    else:
        inv_pairs = pd.DataFrame(columns=["Client", "N° facture"])

    for i, (_, row_inv) in enumerate(inv_pairs.iterrows(), start=2):
        sup_val = clean_cell(row_inv["Client"])
        inv_val = clean_cell(row_inv["N° facture"])
        values = [
            sup_val,
            inv_val,
            f"=SUMIFS({R_HT},{R_FACTURE},B{i})",
            None,  # printed total - empty by default
            f"=ROUND(C{i}-D{i},2)",
        ]
        for j, v in enumerate(values, start=1):
            cell = wcf.cell(row=i, column=j, value=v)
            cell.font = BODY_FONT
            if j == 2:
                cell.number_format = "@"
        for j, fmt in [(3, "#,##0.00"), (4, "#,##0.00"), (5, "#,##0.00")]:
            wcf.cell(row=i, column=j).number_format = fmt

    cf_widths = [34, 16, 18, 20, 12]
    for j, w in enumerate(cf_widths, start=1):
        wcf.column_dimensions[get_column_letter(j)].width = w
    wcf.freeze_panes = "A2"
    wcf.auto_filter.ref = f"A1:{get_column_letter(len(cf_headers))}1"

    # ------------------------------------------------------------------ MANIFEST (7 cols)
    wm = wb.create_sheet("MANIFEST")
    wm.append(MANIFEST_COLUMNS)
    for c_idx in range(1, len(MANIFEST_COLUMNS) + 1):
        hcell = wm.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wm.row_dimensions[1].height = 24

    man_num_cols = {"Poids (kg)", "Colis", "Goods Value"}
    if not manifest.empty:
        for _, mr in manifest.iterrows():
            row_vals = []
            for mc in MANIFEST_COLUMNS:
                v = mr.get(mc)
                if mc in man_num_cols:
                    try:
                        v = float(v) if v is not None and not (isinstance(v, float) and math.isnan(v)) else None
                    except (ValueError, TypeError):
                        v = None
                elif mc == "Page PDF":
                    try:
                        v = int(v) if v is not None and not (isinstance(v, float) and math.isnan(v)) else None
                    except (ValueError, TypeError):
                        v = None
                else:
                    v = clean_cell(v) or None
                row_vals.append(v)
            wm.append(row_vals)
    man_last = max(len(manifest) + 1, 2) if not manifest.empty else 2
    # format Goods Value column F
    gv_col = get_column_letter(MANIFEST_COLUMNS.index("Goods Value") + 1)
    for row_m in wm.iter_rows(min_row=2, max_row=man_last):
        for cell_m in row_m:
            cell_m.font = BODY_FONT
        row_m[MANIFEST_COLUMNS.index("Goods Value")].number_format = "#,##0.00"
    man_widths = [9, 20, 38, 12, 8, 14, 9]
    for j, w in enumerate(man_widths, start=1):
        wm.column_dimensions[get_column_letter(j)].width = w
    wm.freeze_panes = "A2"
    wm.auto_filter.ref = f"A1:{get_column_letter(len(MANIFEST_COLUMNS))}1"

    # ------------------------------------------------------------------ INVENTAIRE_PAGES (4 cols)
    wi = wb.create_sheet("INVENTAIRE_PAGES")
    ip_headers = ["Expéditeur / type", "Pages", "Type", "Statut"]
    wi.append(ip_headers)
    for c_idx in range(1, len(ip_headers) + 1):
        hcell = wi.cell(row=1, column=c_idx)
        hcell.fill = HEADER_FILL
        hcell.font = HEADER_FONT
        hcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wi.row_dimensions[1].height = 24

    if docs is not None and not docs.empty:
        # get set of suppliers that have extracted lines
        extracted_sups = set(lines["Client"].dropna().unique()) if not lines.empty else set()
        for _, doc_row in docs.iterrows():
            doc_type = str(doc_row.get("Type", ""))
            if doc_type not in ("FACTURE", "PACKING"):
                continue
            sup_name = str(doc_row.get("Fournisseur", "") or "")
            pages = str(doc_row.get("Pages", "") or "")
            ocr_val = str(doc_row.get("OCR", "") or "")
            type_label = "Scanné" if ocr_val == "Oui" else "Texte"
            # check if extracted
            found_sup = any(company_similarity(sup_name, s) >= 0.6 for s in extracted_sups) if sup_name else False
            statut = "Extrait" if found_sup else "À lire (probable facture manquante...)"
            wi.append([sup_name or doc_type, pages, type_label, statut])
        ip_last = max(len(docs) + 1, 2)
    else:
        ip_last = 2
    for row_i in wi.iter_rows(min_row=2, max_row=ip_last):
        for cell_i in row_i:
            cell_i.font = BODY_FONT
    ip_widths = [34, 26, 16, 70]
    for j, w in enumerate(ip_widths, start=1):
        wi.column_dimensions[get_column_letter(j)].width = w
    wi.freeze_panes = "A2"

    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# =============================================================================
# INTERFACE STREAMLIT
# =============================================================================

def compute_summary(lines: pd.DataFrame) -> dict:
    # group by internal _devise if present, otherwise no currency grouping
    if "_devise" in lines.columns and not lines.empty:
        by_cur = lines.groupby("_devise")["Valeur HT"].sum().to_dict()
    else:
        by_cur = {}
    return {
        "lignes": len(lines),
        "factures": len(lines[["Client", "N° facture"]].drop_duplicates()) if len(lines) else 0,
        "fournisseurs": lines["Client"].replace("", pd.NA).dropna().nunique() if len(lines) else 0,
        "non_trouves": int((lines["NGP (REF_PAR_FRS)"] == NON_TROUVE).sum()) if len(lines) else 0,
        "ecarts": int((lines["Contrôle NGP"] == "ÉCART").sum()) if len(lines) else 0,
        "a_verifier": int((lines["Contrôle NGP"] == "À vérifier").sum()) if len(lines) else 0,
        "ok": int((lines["Contrôle NGP"] == "OK").sum()) if len(lines) else 0,
        "par_devise": by_cur,
    }


def fmt(v: float, d: int = 2) -> str:
    return f"{v:,.{d}f}".replace(",", " ").replace(".", ",")


def highlight_rows(row):
    styles = [""] * len(row)
    cols = list(row.index)
    if "NGP (REF_PAR_FRS)" in cols and row.get("NGP (REF_PAR_FRS)") == NON_TROUVE:
        styles[cols.index("NGP (REF_PAR_FRS)")] = "color:#C00000;font-weight:bold"
    if "Réf article" in cols and row.get("Réf article") == NON_TROUVE:
        styles[cols.index("Réf article")] = "color:#C00000;font-weight:bold"
    ctrl = row.get("Contrôle NGP")
    color = {"OK": "#C6EFCE", "ÉCART": "#FFC7CE", "À vérifier": "#FFEB9C"}.get(ctrl)
    if color and "Contrôle NGP" in cols:
        styles[cols.index("Contrôle NGP")] = f"background-color:{color}"
    return styles


def main():
    st.set_page_config(page_title="TIMAR · Ventilation Groupage Alstom Fez", page_icon="🚛", layout="wide",
                       initial_sidebar_state="expanded")
    ui.inject_css()

    # ---------------- Paramètres (barre latérale)
    with st.sidebar:
        ui.sidebar_brand()
        st.markdown('<div class="side-cap">Module</div>', unsafe_allow_html=True)
        module = st.radio("Module", ["📥  IMPORT — Ventilation groupage", "📤  EXPORT — BL & Factures"],
                          key="module")
        is_export = module.startswith("📤")
        basis = "Valeur HT"
        if not is_export:
            st.subheader("⚖️ Ventilation des poids")
            basis = st.radio("Poids globaux répartis au prorata de :", ["Valeur HT", "Quantité"])
        st.subheader("🔍 OCR (pages scannées)")
        ocr_mode = st.selectbox("Moteur OCR", ["Automatique", "Tesseract", "EasyOCR", "Désactivé"])
        tess_path = st.text_input("Chemin tesseract.exe (portable, optionnel)",
                                  value=shutil.which("tesseract") or "",
                                  help="Ex. C:\\Users\\moi\\Tesseract-OCR\\tesseract.exe — aucune installation admin requise.")
        dpi = st.slider("Résolution OCR (DPI)", 150, 400, 300, 50)
        ocr = OCREngine(ocr_mode if ocr_mode != "Désactivé" else "none", tess_path, dpi)
        if ocr.available:
            ui.sidebar_pill(ui.GOOD, f"OCR actif : <b>&nbsp;{ocr.engine}</b>")
        elif ocr_mode == "Désactivé":
            ui.sidebar_pill(ui.INK_3, "OCR désactivé")
        else:
            ui.sidebar_pill(ui.WARNING, "Aucun moteur OCR : PDF texte uniquement")
            st.caption("Pour les scans : `pip install --user easyocr`, ou indiquez un Tesseract portable ci-dessus.")
        st.write("")
        if st.button("🔄 Nouveau dossier", use_container_width=True):
            keep = st.session_state.get("module")
            for k in [k for k in st.session_state if (k.startswith("exp") if is_export else not k.startswith("exp"))]:
                if k != "module":
                    del st.session_state[k]
            st.session_state["module"] = keep
            st.rerun()

    if is_export:
        import export_engine  # import local : évite l'import circulaire avec app.py
        export_engine.render_export(ocr)
        ui.footer()
        return

    # ---------------- En-tête 3D
    if "result" in st.session_state:
        s0 = compute_summary(apply_matching(st.session_state["result"]["lines"], st.session_state["result"]["refdb"]))
        hero_stats = {"lignes ventilées": s0["lignes"], "factures": s0["factures"],
                      "fournisseurs": s0["fournisseurs"], "NON TROUVÉ": s0["non_trouves"]}
    else:
        hero_stats = {"feuilles Excel": 3, "langues lues": "FR·EN·DE·ES·IT", "OCR": "intégré", "traitement": "100 % local"}
    ui.hero_3d(hero_stats)

    # ---------------- Imports
    c1, c2 = st.columns(2, gap="medium")
    with c1:
        with st.container(key="card_upl_pdf"):
            ui.card_title("📂", "Dossier PDF complet de la remorque",
                          "Importer le dossier PDF complet (Factures, Packing Lists, CMR, Manifeste)",
                          ["Factures", "Packing Lists", "CMR", "Manifeste", "Scans OCR"])
            pdf_files = st.file_uploader("Importer le dossier PDF complet (Factures, Packing Lists, CMR, Manifeste)",
                                         type=["pdf"], accept_multiple_files=True)
    with c2:
        with st.container(key="card_upl_ref"):
            ui.card_title("📗", "Référentiel articles & NGP",
                          "Importer le fichier Excel de référence (Références & NGP)",
                          ["Réf Article", "Réf Alstom (PN)", "Code NGP", ".xlsx / .xls / .csv"])
            ref_file = st.file_uploader("Importer le fichier Excel de référence (Références & NGP)",
                                        type=["xlsx", "xlsm", "xls", "csv"])

    done = "result" in st.session_state
    ui.steps_strip(2 if done else (1 if pdf_files else 0))

    launch = st.button("🚀  Lancer la ventilation du dossier", type="primary", use_container_width=True,
                       disabled=not pdf_files)
    if pdf_files and not ref_file:
        st.caption("⚠️ Sans fichier de référence, toutes les lignes seront « NON TROUVÉ ».")

    if launch:
        progress = st.progress(0.0, text="Initialisation...")
        try:
            result = run_pipeline(pdf_files, ref_file, ocr, basis, progress)
        except Exception as exc:  # garde-fou : l'application ne doit jamais « planter »
            progress.empty()
            st.error(f"Erreur pendant le traitement : {exc}")
            st.exception(exc)
            return
        st.session_state["result"] = result
        st.session_state["lines_edit"] = result["lines"]
        st.session_state["manifest_edit"] = result["manifest"]
        st.session_state["editor_version"] = st.session_state.get("editor_version", 0) + 1
        matched = apply_matching(result["lines"], result["refdb"])
        st.session_state["xlsx"] = build_workbook(matched, result["manifest"], result.get("docs"))
        progress.progress(1.0, text="✅ Traitement terminé — classeur Excel prêt.")
        st.rerun()  # rafraîchit l'en-tête et les étapes avec les résultats

    if "result" not in st.session_state:
        ui.footer()
        return
    result = st.session_state["result"]
    ver = st.session_state.get("editor_version", 0)

    tab_dash, tab_lines, tab_man, tab_docs = st.tabs(
        ["📊  Résumé exécutif", "📋  Lignes extraites (modifiables)", "🧾  Manifeste", "🗂️  Documents & journal"])

    with tab_lines:
        st.caption("Corrigez/complétez les données si besoin : le matching NGP, le contrôle et le classeur "
                   "Excel sont recalculés automatiquement. Ajout/suppression de lignes possible.")
        edited = st.data_editor(
            st.session_state["lines_edit"], num_rows="dynamic", use_container_width=True, height=460,
            key=f"lines_editor_{ver}",
            column_config={c: st.column_config.NumberColumn(c, format="%.2f")
                           for c in NUMERIC_COLUMNS})
    with tab_man:
        st.caption("Goods Value du manifeste par fournisseur (extraits automatiquement si le manifeste "
                   "contient un tableau ; sinon à saisir). Rattachement par nom.")
        man_edit = st.data_editor(
            st.session_state["manifest_edit"], num_rows="dynamic", use_container_width=True,
            key=f"manifest_editor_{ver}",
            column_config={"Goods Value": st.column_config.NumberColumn(format="%.2f"),
                           "Poids (kg)": st.column_config.NumberColumn(format="%.3f")})

    edited = edited.copy()
    for c in EDITABLE_COLUMNS:
        if c not in edited.columns:
            edited[c] = None
    for c in NUMERIC_COLUMNS:
        if c in edited.columns:
            edited[c] = pd.to_numeric(edited[c], errors="coerce")
    if "Page PDF" in edited.columns:
        edited["Page PDF"] = pd.to_numeric(edited["Page PDF"], errors="coerce").fillna(0).astype(int)
    for c in EDITABLE_COLUMNS:
        if c not in NUMERIC_COLUMNS and c != "Page PDF" and not c.startswith("_"):
            if c in edited.columns:
                edited[c] = edited[c].map(clean_cell)
    # preserve internal columns if present
    internal_cols_present = [c for c in ["_réf_alstom", "_devise", "_poids_net", "_poids_brut", "_colis"]
                             if c in edited.columns]
    editable_with_internal = [c for c in EDITABLE_COLUMNS if c in edited.columns] + internal_cols_present
    final = apply_matching(edited[editable_with_internal], result["refdb"])
    man_final = man_edit.copy()
    for c in ("Goods Value", "Poids (kg)", "Colis"):
        if c in man_final.columns:
            man_final[c] = pd.to_numeric(man_final[c], errors="coerce")
    for c in ("Expéditeur", "Shipment", "Devise"):
        if c in man_final.columns:
            man_final[c] = man_final[c].map(clean_cell)

    orig_lines = result["lines"]
    orig_editable = [c for c in EDITABLE_COLUMNS if c in orig_lines.columns]
    edit_editable = [c for c in EDITABLE_COLUMNS if c in edited.columns]
    try:
        lines_changed = not edited[edit_editable].reset_index(drop=True).equals(orig_lines[orig_editable].reset_index(drop=True))
    except Exception:
        lines_changed = True
    try:
        man_changed = not man_final.reset_index(drop=True).equals(result["manifest"].reset_index(drop=True))
    except Exception:
        man_changed = True
    changed = lines_changed or man_changed
    xlsx = build_workbook(final, man_final, result.get("docs")) if changed else st.session_state["xlsx"]

    with tab_dash:
        s = compute_summary(final)
        pct_ok = round(100 * s["ok"] / s["lignes"]) if s["lignes"] else 0
        ui.kpi_cards([
            {"icon": "🧾", "label": "Lignes extraites", "value": s["lignes"], "sub": "individualisées", "tone": "blue"},
            {"icon": "📑", "label": "Factures", "value": s["factures"], "sub": "dans le dossier", "tone": "cyan"},
            {"icon": "🏭", "label": "Fournisseurs", "value": s["fournisseurs"], "sub": "expéditeurs", "tone": "gold"},
            {"icon": "❓", "label": "NON TROUVÉ", "value": s["non_trouves"],
             "sub": "à créer dans le référentiel" if s["non_trouves"] else "aucun", "tone": "warn" if s["non_trouves"] else "good"},
            {"icon": "⚠️", "label": "Écarts NGP", "value": s["ecarts"],
             "sub": "NGP Excel ≠ facture" if s["ecarts"] else "aucun", "tone": "bad" if s["ecarts"] else "good"},
            {"icon": "✅", "label": "Contrôle OK", "value": f"{pct_ok} %", "sub": f"{s['ok']} / {s['lignes']} lignes",
             "tone": "good" if pct_ok == 100 else "blue"},
        ])
        cards2 = [{"icon": "💶", "label": f"Total Valeur HT {dev or '(devise ?)'}", "value": round(float(tot), 2),
                   "decimals": 2, "sub": dev or "", "tone": "blue"} for dev, tot in s["par_devise"].items()]
        if cards2:
            ui.kpi_cards(cards2[:6])

        if s["lignes"] and not s["non_trouves"] and not s["ecarts"]:
            ui.status_banner("ok", "🟢", "Statut général : CONFORME", "Dossier prêt pour la déclaration.")
        elif s["lignes"]:
            ui.status_banner("warn", "🟠", "Statut général : À VÉRIFIER",
                             f"{s['non_trouves']} NON TROUVÉ · {s['ecarts']} écart(s) NGP · "
                             f"{s['a_verifier']} ligne(s) à vérifier.")
        else:
            ui.status_banner("ko", "🔴", "Aucune ligne de facture extraite",
                             "Consultez l'onglet « Documents & journal » ou saisissez les lignes manuellement.")

        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button("📥  Télécharger le fichier Excel ventilé (.xlsx)", data=xlsx,
                           file_name=f"Ventilation_Groupage_Alstom_Fez_{stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", use_container_width=True)

        if len(final):
            # adapt dataframe for ui.py chart functions that use old column names
            _devise_col = final["_devise"] if "_devise" in final.columns else pd.Series([""] * len(final))
            final_for_ui = final.rename(columns={
                "Client": "Fournisseur / Expéditeur",
                "N° facture": "N° Facture",
                "NGP (REF_PAR_FRS)": "NGP (Excel Réf)",
            }).assign(**{"Devise": _devise_col})
            g1, g2 = st.columns([1.5, 1], gap="medium")
            with g1:
                with st.container(key="card_chart_val"):
                    ui.card_title("💶", "Valeur HT par fournisseur", "Survolez une barre pour le détail")
                    try:
                        ui.chart_value_by_supplier(final_for_ui)
                    except Exception:
                        st.write(final_for_ui.groupby("Fournisseur / Expéditeur")["Valeur HT"].sum().to_frame())
            with g2:
                with st.container(key="card_chart_ngp"):
                    ui.card_title("🎯", "Contrôle NGP", "Répartition des lignes par statut")
                    try:
                        ui.chart_ngp_status(final_for_ui)
                    except Exception:
                        st.write(final_for_ui["Contrôle NGP"].value_counts().to_frame())

            ui.section("Synthèse par fournisseur", "🏭")
            _devise_series = final["_devise"] if "_devise" in final.columns else pd.Series([""] * len(final), index=final.index)
            recap = (final.assign(
                NT=final["NGP (REF_PAR_FRS)"].eq(NON_TROUVE),
                EC=final["Contrôle NGP"].eq("ÉCART"),
                _dev=_devise_series,
            ).groupby(["Client", "_dev"], sort=False)
             .agg(Factures=("N° facture", "nunique"), Lignes=("N° facture", "size"),
                  Valeur_HT=("Valeur HT", "sum"),
                  NON_TROUVÉ=("NT", "sum"), ÉCARTS=("EC", "sum")).reset_index()
             .rename(columns={"Client": "Fournisseur / Expéditeur", "_dev": "Devise"}))
            st.dataframe(recap, use_container_width=True, hide_index=True,
                         column_config={"Valeur_HT": st.column_config.NumberColumn("Valeur HT", format="%.2f"),
                                        "Lignes": st.column_config.ProgressColumn(
                                            "Lignes", format="%d", min_value=0, max_value=int(recap["Lignes"].max()))})
            ui.section("Aperçu TOUS_CLIENTS", "📋")
            st.dataframe(final.style.apply(highlight_rows, axis=1), use_container_width=True, height=380,
                         hide_index=True)

    with tab_docs:
        ui.section("Documents détectés dans le dossier", "🗂️")
        st.dataframe(result["docs"], use_container_width=True, hide_index=True)
        st.caption(f"Référentiel chargé : {result['refdb']['count']} références.")
        ui.section("Journal de traitement", "📝")
        for msg in result["log"]:
            st.write(msg)

    ui.footer()


if __name__ == "__main__":
    main()
