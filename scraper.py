#!/usr/bin/env python3
"""Oreshak.bg -> Temu XLSX scraper.

The program deliberately edits the supplied Temu workbook through OOXML/ZIP
instead of rebuilding it. This preserves sheet names, hidden helper sheets,
drop-downs, formulas, conditional formatting and Temu metadata.
"""
from __future__ import annotations

import argparse
import csv
import itertools
import json
import logging
import math
import os
import posixpath
import random
import re
import shutil
import sys
import time
import unicodedata
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping, Sequence
from urllib.parse import parse_qsl, quote, unquote, urlencode, urljoin, urlparse, urlunparse
from zipfile import ZIP_DEFLATED, ZipFile

try:
    import cloudscraper
except ImportError:  # schema-only/local environments before pip install
    cloudscraper = None
import requests
from bs4 import BeautifulSoup, Tag
from lxml import etree

BASE_URL = "https://oreshak.bg/"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_XML = "http://www.w3.org/XML/1998/namespace"
NS = {"a": NS_MAIN, "r": NS_REL}
PROJECT_VERSION = "6.9"
MAX_TEMU_ROWS = 2000

SOURCE_DEFAULTS: dict[str, str] = {
    "aksesoari-za-lovni-trofei": "32650",
    "aksesoari-za-shah-i-tabla": "25613",
    "bitova-takan": "39650",
    "baklitsi-i-bureta": "10905",
    "darvorezbovani-pana-i-plastiki": "12151",
    "dyalani-unikati-ot-darvo": "12140",
    "kartini-ot-bulgaria": "12867",
    "komplekti-shah-i-tabla": "25615",
    "kutii-za-aksesoari": "12179",
    "kutii-za-shah-i-tabla": "25615",
    "kuhnenski-aksesoari-ot-darvo-oreshak": "9923",
    "nojove-ot-balgaria": "10059",
    "profesionalen-shahmat": "25613",
    "ruchno-izraboteni-chinii-ot-darvo": "10808",
    "kutii-za-vino-i-bijuta": "39880",
    "suveniri": "13020",
    "suveniri-ot-oreshak": "13020",
    "wooden-souvenirs-white-blank": "39981",
    "suveniri-ot-metal": "13020",
}

# Ordered from specific to general. The template supplied by the user contains
# only these Temu categories, so every mapping stays inside its valid universe.
KEYWORD_CATEGORY_RULES: list[tuple[re.Pattern[str], str, str]] = [
    (re.compile(r"\b(монет|coin|жетон)"), "54621", "collectible coin"),
    (re.compile(r"(ключодърж|талисман|чарм|медальон)"), "39350", "charm/keyring"),
    (re.compile(r"(висулк|пендант|pendant)"), "13027", "pendant ornament"),
    (re.compile(r"(?:кутия|ракла).*(?:бижут|часовник|пръстен|обеци)|(?:бижут|часовник|пръстен|обеци).*(?:кутия|ракла)"), "12735", "jewelry/watch box"),
    (re.compile(r"\bракла\b"), "12179", "decorative storage chest"),
    (re.compile(r"(кутия|кашон).*(подар|gift)"), "39880", "gift box"),
    (re.compile(r"(стойка|поставка|рафт).*(вино|бутил)"), "12686", "wine rack"),
    (re.compile(r"(кутия).*(вино|бутил)"), "39880", "wine gift box"),
    (re.compile(r"\bбуре|бъчв"), "10905", "barrel"),
    (re.compile(r"бъклиц|манерк|кег"), "10888", "flask/keg"),
    (re.compile(r"(трофе|глиги|сръндак|глиган|елен).*(дъск|поставка)|дъск.*трофе"), "32650", "trophy mount"),
    (re.compile(r"(?:торбичка|чанта|калъф|кейс).*(?:шах|табла|зар)|(?:шах|табла|зар).*(?:торбичка|чанта|калъф|кейс)"), "25613", "chess/backgammon storage accessory"),
    (re.compile(r"(?:професионал|състезател|стаунтон|лети).*(?:фигур)|(?:фигур).*(?:професионал|състезател|стаунтон)"), "25613", "professional chess pieces"),
    (re.compile(r"(шахматен часовник|фигури за шах|пулове|\bзар\b|зарове|зарчета|аксесоар.*шах|аксесоар.*табла)"), "25613", "game pieces/accessories"),
    (re.compile(r"(комплект|сет).*(шах|табла)|шах.*табла"), "25615", "board game"),
    (re.compile(r"професионален.*шах|шахмат.*професион"), "51777", "professional chess"),
    (re.compile(r"(солниц|соларник|salt cellar|salt box)"), "10703", "salt cellar/serveware accessory"),
    (re.compile(r"(черпак|черпало)"), "9998", "ladle"),
    (re.compile(r"(лъжиц|лъжич|шпатула)"), "9999", "cooking spoon"),
    (re.compile(r"(комплект|сет).*(прибор|кухненск|лъжиц|шпатула)"), "10006", "utensil set"),
    (re.compile(r"(?:дъска).*(?:рязане|сервира|мезе|сирена|кухнен|домаш|гурме)|(?:кухнен|домаш|гурме).*(?:дъска)|(?:тал[аъ]р)|serving board"), "54423", "serving board"),
    (re.compile(r"\bподнос|tray"), "10741", "serving tray"),
    (re.compile(r"\bплато|platter"), "10740", "platter"),
    (re.compile(r"чини(я|и).*(пирограф|фолклор|сувенир|закач)|(?:пирограф|фолклор|сувенир).*(чини(я|и))"), "10853", "novelty/decorative plate"),
    (re.compile(r"(десертн).*(чини)|чини.*десерт"), "10807", "dessert plate"),
    (re.compile(r"\bчини(я|и)|plate"), "10808", "dinner plate"),
    (re.compile(r"(поставка|органайзер).*(нож|прибор)"), "10328", "utensil rack"),
    (re.compile(r"(нож).*(хранене|масов)"), "10638", "dinner knife"),
    (re.compile(r"(нож).*(плод|универсал|ловен|турист|спортен|джоб)|\b(?:ловен|туристически|спортен) нож"), "10072", "utility knife"),
    (re.compile(r"(кухненски нож|готварски нож|домакински нож|домашен нож|нож.*(?:готвач|сирена)|касапски нож|сатър|chef)"), "10059", "chef knife"),
    (re.compile(r"(кутия).*(нож|аксесоар|пура|тютюн)"), "12179", "decorative box"),
    (re.compile(r"(отварачка).*(вино|тирбушон)"), "10875", "wine accessory"),
    (re.compile(r"(отварачка|gadget)"), "9923", "kitchen gadget"),
    (re.compile(r"(табел|надпис|плакет|plaque|дъска.*послание)"), "12193", "decorative plaque"),
    (re.compile(r"(доза|кутия).*(подправк)"), "10703", "spice container/serveware accessory"),
    (re.compile(r"(звънче|хлопка)"), "12141", "hanging bell ornament"),
    (re.compile(r"\bскрин\b"), "12179", "decorative storage box"),
    (re.compile(r"(керамичн).*(слон|цървул)|(?:слонче|цървулк).*(керами)"), "12140", "ceramic collectible figurine"),
    (re.compile(r"(пан[оo]|пластик|релеф|дърворезб)"), "12151", "wall sculpture"),
    (re.compile(r"(картина|живопис|painting)"), "12867", "painting"),
    (re.compile(r"(гравюра|дърворит|woodcut)"), "39291", "woodcut"),
    (re.compile(r"(смесена техника|mixed media)"), "39280", "mixed media"),
    (re.compile(r"(бяла заготовка|неоцветен|за оцветяване|blank)"), "39981", "wood art blank"),
    (re.compile(r"(колед|елха|орнамент|украса).*(фигур|статует)|figurine ornament"), "13031", "figurine ornament"),
    (re.compile(r"(висящ|за окачване|hanging ornament)"), "12141", "hanging ornament"),
    (re.compile(r"(фигур|статует|скулптур|figurine)"), "12140", "collectible figurine"),
    (re.compile(r"(тъкан|плат|fabric|текстил)"), "39650", "fabric"),
]

LOW_CONFIDENCE_PRODUCT_RULES: list[tuple[re.Pattern[str], str]] = [
    # Consumables and regulated goods are intentionally excluded because the
    # user-supplied Temu template has no matching food, alcohol, cosmetic or
    # fragrance categories.
    (re.compile(r"(?:гюлова|сливова|стара\s+сливова|кумова)\s+ракия|\bликьор\b|\bалкохол", re.I), "alcohol is outside the categories in the supplied Temu template"),
    (re.compile(r"\b(?:сладко|конфитюр|мармалад)\b", re.I), "food is outside the categories in the supplied Temu template"),
    (re.compile(r"(?:розова|лавандулова)\s+вода", re.I), "floral/cosmetic water is outside the categories in the supplied Temu template"),
    (re.compile(r"етеричн(?:о|и|а)\s+масл|(?:розово|лавандулово)\s+масло|масло\s+от\s+(?:роза|лавандула)", re.I), "essential oil is outside the categories in the supplied Temu template"),
    (re.compile(r"\bпарфюм\b|крем\s+за\s+ръце|\bкозмет", re.I), "perfume/cosmetic product is outside the categories in the supplied Temu template"),

    # Smoking, apparel and personal-accessory products must not inherit a box,
    # knife or figurine category simply because of their source section.
    (re.compile(r"хумидор|запалка|калъф.*тютюн|кутия.*пур", re.I), "no safe smoking-accessory category exists in the supplied Temu template"),
    (re.compile(r"\bколан\b|\bпафти?\b", re.I), "no belt/apparel-accessory category exists in the supplied Temu template"),
    (re.compile(r"^(?:кожена\s+)?кания\b|^калъф\s+за\s+нож\b|\bножница\b", re.I), "knife sheath is not a knife and no sheath category exists in the supplied Temu template"),
    (re.compile(r"\b(?:колие|гривна|обици)\b", re.I), "no jewelry category exists in the supplied Temu template"),

    # Functional home/kitchen objects for which the supplied template has no
    # defensible leaf category.
    (re.compile(r"стенен\s+часовник|часовников\s+механизъм", re.I), "no wall-clock category exists in the supplied Temu template"),
    (re.compile(r"стойка\s+за\s+ключове|\bзакачалка\b", re.I), "no key-holder/coat-hook category exists in the supplied Temu template"),
    (re.compile(r"\bточилка\b|\bхаван(?:че)?\b|\bхалба\b|\bюзче\b|\bшиш(?:ове)?\b|\bщипки\b", re.I), "no safe category for this kitchen utensil exists in the supplied Temu template"),
    (re.compile(r"дръжка.*дъск|дъск.*дръжка", re.I), "board handle/hardware is not a serving board and no hardware category exists in the supplied Temu template"),
    (re.compile(r"\bкупа\b|\bкупичка\b|\bгаванка\b", re.I), "no bowl category exists in the supplied Temu template"),
    (re.compile(r"запушалка.*(?:вино|бутилка)|(?:вино|бутилка).*запушалка", re.I), "single bottle stopper is not a Wine Accessory Set"),
    (re.compile(r"отварачка.*бира", re.I), "single beer opener has no safe category in the supplied Temu template"),
    (re.compile(r"бутилка\s+уникат|буркан.*ваза", re.I), "decorative bottle/vase has no safe category in the supplied Temu template"),
    (re.compile(r"поставка\s+за\s+химикалки|\bмоливник\b", re.I), "no pen-holder category exists in the supplied Temu template"),
    (re.compile(r"\b(?:меч|сабя|каракулак)\b", re.I), "sword/sabre product is not covered by the supplied kitchen-knife categories"),

    # Barrel taps/plugs are accessories, not barrels, and the supplied template has no safe leaf for them.
    (re.compile(r"(?:тапа|канелк).*(?:буре|бъчв)|(?:буре|бъчв).*(?:тапа|канелк)", re.I), "barrel tap/plug is not a barrel and no barrel-accessory category exists in the supplied Temu template"),

    # Previously identified unsupported products.
    (re.compile(r"бъклиц|манерк|hip\s*flask", re.I), "no flask/drinkware category exists in the supplied Temu template"),
    (re.compile(r"покривк|tablecloth|чорап|терлиц|пафт", re.I), "finished table linen/apparel is not raw Fabric and no suitable category exists in the supplied Temu template"),
    (re.compile(r"подков|horseshoe", re.I), "no safe horseshoe/decorative-hardware category exists in the supplied Temu template"),
    (re.compile(r"пепелник|ashtray", re.I), "no ashtray category exists in the supplied Temu template"),
    (re.compile(r"възглавничк|cushion|pillow", re.I), "no cushion category exists in the supplied Temu template"),
    (re.compile(r"фиолк.*есенц|есенц.*лавандул|fragrance vial", re.I), "no fragrance/essential-oil category exists in the supplied Temu template"),
    (re.compile(r"лъжиц.*(?:танц|хоро)|(?:танц|хоро).*лъжиц|лъжич.*сватб", re.I), "ceremonial/folk prop is not a cooking spoon"),
]

# These source sections contain many unrelated product types. A source-category
# default is therefore unsafe unless the title matched a specific rule above.
STRICT_EXPLICIT_MAPPING_SLUGS: set[str] = {
    "dyalani-unikati-ot-darvo",
    "kuhnenski-aksesoari-ot-darvo-oreshak",
    "nojove-ot-balgaria",
    "kutii-za-vino-i-bijuta",
    "suveniri",
    "suveniri-ot-oreshak",
    "suveniri-ot-metal",
}

MATERIAL_TRANSLATIONS: list[tuple[re.Pattern[str], list[str]]] = [
    (re.compile(r"\bкост\b|bone"), ["Bone"]),
    (re.compile(r"дамаск"), ["Damascus Steel", "Stainless Steel", "Steel"]),
    (re.compile(r"неръждаем"), ["Stainless Steel", "Steel"]),
    (re.compile(r"желяз|\biron\b"), ["Iron", "Cast Iron", "Metal", "Steel"]),
    (re.compile(r"високовъглерод"), ["High carbon steel", "Carbon Steel", "Steel"]),
    (re.compile(r"въглерод"), ["Carbon Steel", "High carbon steel", "Steel"]),
    (re.compile(r"стоман"), ["Steel", "Carbon Steel", "Stainless Steel"]),
    (re.compile(r"месинг"), ["Brass", "Copper Alloy", "Metal"]),
    (re.compile(r"медн|\bмед\b"), ["Copper", "Copper Alloy", "Metal"]),
    (re.compile(r"алумини"), ["Aluminum Alloy", "Aluminum", "Metal"]),
    (re.compile(r"метал"), ["Metal", "Iron", "Steel", "Zinc Alloy"]),
    (re.compile(r"кож|еленов рог"), ["Leather", "Genuine", "Buckskin", "Animal Skin, Fur, And Down"]),
    (re.compile(r"керами"), ["Ceramic", "Ceramics", "Clay"]),
    (re.compile(r"стък"), ["Glass"]),
    (re.compile(r"\bплат\b|\bтъкан\b|текстил|вата|плюш"), ["Fabric", "Textile", "Cotton", "Linen"]),
    (re.compile(r"харт|картон"), ["Paper", "Cardboard"]),
    (re.compile(r"смол|резин"), ["Resin"]),
    (re.compile(r"бамбук"), ["Bamboo", "Wood", "Log"]),
    (re.compile(r"дърв|бук|липа|орех|дъб|ясен|бор|махагон|акация"), ["Wood", "Log", "Solid Wood", "Natural Wood"]),
]

WOOD_SPECIES = {
    "бук": ["Beech", "European Beech"],
    "липа": ["Basswood", "Linden", "Lime Wood"],
    "орех": ["Walnut"],
    "дъб": ["Oak"],
    "ясен": ["Ash"],
    "бор": ["Pine"],
    "махагон": ["Mahogany"],
    "акация": ["Acacia"],
    "венге": ["Wenge"],
    "кестен": ["Chestnut"],
}

DEFAULT_PACKAGE: dict[str, tuple[float, float, float, float]] = {
    "32650": (600, 35, 25, 5), "25613": (300, 20, 15, 8), "25615": (1800, 40, 40, 8),
    "51777": (1800, 45, 45, 8), "39650": (500, 50, 35, 8), "10905": (1200, 30, 25, 25),
    "10888": (900, 30, 20, 20), "12151": (1200, 50, 35, 6), "12867": (1200, 60, 45, 6),
    "12140": (500, 25, 18, 15), "13020": (350, 20, 15, 12), "13031": (250, 18, 12, 10),
    "12141": (200, 18, 12, 8), "12193": (500, 35, 25, 5), "39981": (300, 25, 18, 8),
    "12179": (700, 30, 22, 12), "12735": (700, 30, 22, 12), "39880": (650, 38, 15, 12),
    "10059": (350, 40, 8, 5), "10072": (350, 40, 8, 5), "10638": (200, 30, 6, 4),
    "9998": (250, 35, 10, 8), "9999": (200, 35, 10, 8), "10006": (800, 40, 20, 12),
    "9923": (400, 30, 20, 12), "54423": (900, 45, 30, 5), "10740": (700, 40, 30, 6),
    "10741": (800, 45, 32, 8), "10807": (450, 25, 25, 6), "10808": (550, 30, 30, 7),
    "10853": (500, 28, 28, 7), "10328": (700, 35, 20, 18), "10875": (500, 30, 20, 10),
    "54621": (100, 10, 8, 3), "39350": (100, 12, 8, 3), "13027": (100, 12, 8, 3),
}

@dataclass
class ProductOption:
    name: str
    values: list[str]

@dataclass
class Product:
    url: str
    source_category_url: str
    source_category_name: str = ""
    title: str = ""
    code: str = ""
    description: str = ""
    bullet_points: list[str] = field(default_factory=list)
    attributes: dict[str, str] = field(default_factory=dict)
    tags: list[str] = field(default_factory=list)
    images: list[str] = field(default_factory=list)
    price_eur: Decimal | None = None
    list_price_eur: Decimal | None = None
    in_stock: bool = True
    stock_quantity: int | None = None
    options: list[ProductOption] = field(default_factory=list)
    category_id: str = ""
    mapping_reason: str = ""
    mapping_confidence: str = "medium"
    weight_g: float | None = None
    dimensions_cm: tuple[float, float, float] | None = None
    price_source: str = ""
    list_price_source: str = ""
    stock_source: str = ""
    weight_source: str = ""
    dimensions_source: str = ""
    warnings: list[str] = field(default_factory=list)

@dataclass
class Variant:
    product: Product
    option_values: dict[str, str]
    sku: str
    title: str
    image: str


def normalize_space(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def slug_key(url: str) -> str:
    return urlparse(url).path.rstrip("/").split("/")[-1].lower()


def canonical(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).casefold()
    return "".join(ch for ch in text if ch.isalnum())


def col_number(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def col_letter(number: int) -> str:
    out = ""
    while number:
        number, rem = divmod(number - 1, 26)
        out = chr(65 + rem) + out
    return out


def cell_col(address: str) -> str:
    match = re.match(r"([A-Z]+)", address)
    return match.group(1) if match else ""


def safe_decimal(value: str | float | int | Decimal | None) -> Decimal | None:
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    text = str(value).replace("\xa0", " ").strip().replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return Decimal(text).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None


def add_query(url: str, **params: Any) -> str:
    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query.update({k: str(v) for k, v in params.items()})
    return urlunparse(parsed._replace(query=urlencode(query)))


def dedupe(items: Iterable[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for item in items:
        item = normalize_space(item)
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result


class TemplateSchema:
    def __init__(self, template_path: Path):
        self.path = template_path
        self._zip = ZipFile(template_path)
        self.shared_strings = self._load_shared_strings()
        self.sheet_paths = self._load_sheet_paths()
        self.workbook_xml = etree.fromstring(self._zip.read("xl/workbook.xml"))
        self.defined_names = {
            node.get("name"): node.text
            for node in self.workbook_xml.xpath("//a:definedNames/a:definedName", namespaces=NS)
            if node.get("name") and node.text
        }
        self._sheet_cache: dict[str, dict[str, Any]] = {}
        self.headers, self.internal_keys = self._load_template_headers()
        self.category_names = self._load_categories()
        self.rules = self._load_rules()
        self.dropdown_map = self._load_dropdown_map()
        self.validation_formulas = self._load_validation_formulas()
        self._dropdown_cache: dict[str, list[str]] = {}

    def close(self) -> None:
        self._zip.close()

    def _load_shared_strings(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self._zip.namelist():
            return []
        root = etree.fromstring(self._zip.read("xl/sharedStrings.xml"))
        return ["".join(node.itertext()) for node in root.findall(f"{{{NS_MAIN}}}si")]

    def _load_sheet_paths(self) -> dict[str, str]:
        workbook = etree.fromstring(self._zip.read("xl/workbook.xml"))
        rels = etree.fromstring(self._zip.read("xl/_rels/workbook.xml.rels"))
        rel_map = {node.get("Id"): node.get("Target") for node in rels}
        paths: dict[str, str] = {}
        for sheet in workbook.xpath("//a:sheets/a:sheet", namespaces=NS):
            target = rel_map[sheet.get(f"{{{NS_REL}}}id")]
            paths[sheet.get("name")] = posixpath.normpath(posixpath.join("xl", target))
        return paths

    def cell_value(self, cell: etree._Element) -> Any:
        cell_type = cell.get("t")
        value = cell.find(f"{{{NS_MAIN}}}v")
        if value is None:
            inline = cell.find(f"{{{NS_MAIN}}}is")
            return "".join(inline.itertext()) if inline is not None else None
        if cell_type == "s":
            return self.shared_strings[int(value.text)]
        if cell_type == "b":
            return value.text == "1"
        return value.text

    def sheet_values(self, name: str) -> dict[str, Any]:
        if name in self._sheet_cache:
            return self._sheet_cache[name]
        root = etree.fromstring(self._zip.read(self.sheet_paths[name]))
        values = {
            cell.get("r"): self.cell_value(cell)
            for cell in root.findall(f".//{{{NS_MAIN}}}c")
        }
        self._sheet_cache[name] = values
        return values

    def _row_values(self, sheet: str, row_number: int) -> dict[str, Any]:
        root = etree.fromstring(self._zip.read(self.sheet_paths[sheet]))
        row = root.find(f".//{{{NS_MAIN}}}row[@r='{row_number}']")
        if row is None:
            return {}
        return {cell_col(c.get("r")): self.cell_value(c) for c in row.findall(f"{{{NS_MAIN}}}c")}

    def _load_template_headers(self) -> tuple[dict[str, str], dict[str, str]]:
        return self._row_values("Template", 2), self._row_values("Template", 4)

    def _load_categories(self) -> dict[str, str]:
        values = self.sheet_values("Category Name")
        result: dict[str, str] = {}
        for row in range(1, 1000):
            category_id = values.get(f"A{row}")
            name = values.get(f"B{row}")
            if category_id and name:
                result[str(category_id)] = str(name)
        return result

    def _load_rules(self) -> dict[str, dict[str, set[str]]]:
        root = etree.fromstring(self._zip.read(self.sheet_paths["GoodsLevelMode"]))
        result: dict[str, dict[str, set[str]]] = {}
        for row in root.findall(f".//{{{NS_MAIN}}}row"):
            cells = {cell_col(c.get("r")): self.cell_value(c) for c in row.findall(f"{{{NS_MAIN}}}c")}
            marker = str(cells.get("A") or "")
            match = re.match(r"(\d+)_(require|disabled|condition_require)$", marker)
            if not match:
                continue
            category_id, kind = match.groups()
            if category_id not in self.category_names:
                continue
            # Values in the cells are exactly require / disabled / condition_require.
            columns = {col for col, value in cells.items() if col != "A" and value == kind}
            result.setdefault(category_id, {})[kind] = columns
        return result

    def _load_dropdown_map(self) -> dict[str, str]:
        values = self.sheet_values("Dropdown Lists")
        result: dict[str, str] = {}
        for row in range(1, 5000):
            key, name = values.get(f"A{row}"), values.get(f"B{row}")
            if key and name:
                result[str(key)] = str(name)
        return result

    def _load_validation_formulas(self) -> dict[str, str]:
        root = etree.fromstring(self._zip.read(self.sheet_paths["Template"]))
        result: dict[str, str] = {}
        for validation in root.xpath("//a:dataValidations/a:dataValidation", namespaces=NS):
            sqref = validation.get("sqref") or ""
            col = cell_col(sqref)
            formula = validation.find("a:formula1", NS)
            if col and formula is not None and formula.text:
                result[col] = formula.text
        return result

    def values_from_defined_name(self, name: str) -> list[str]:
        reference = self.defined_names.get(name)
        if not reference:
            return []
        match = re.match(r"'([^']+)'!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$", reference)
        if not match:
            return []
        sheet, col1, row1, col2, row2 = match.groups()
        values = self.sheet_values(sheet)
        output: list[str] = []
        for row in range(int(row1), int(row2) + 1):
            for col in range(col_number(col1), col_number(col2) + 1):
                value = values.get(f"{col_letter(col)}{row}")
                if value not in (None, ""):
                    output.append(str(value))
        return output

    def dropdown_values(self, key: str) -> list[str]:
        if key not in self._dropdown_cache:
            self._dropdown_cache[key] = self.values_from_defined_name(self.dropdown_map.get(key, ""))
        return self._dropdown_cache[key]

    def dropdown_for(self, column: str, category_id: str, row: Mapping[str, Any]) -> list[str]:
        header = self.headers.get(column, "")
        if column == "E":
            return list(self.category_names)
        generic_keys = {
            "O": "t_1_Update or Add", "PF": f"t_4_{category_id}_Variation Theme",
            "QO": "t_6_Individually packed", "QQ": "t_6_Packaging unit",
            "QY": "t_7_Shipping Template", "QZ": "t_7_Handling Time",
            "RA": "t_7_Fulfillment Channel", "RB": "t_7_Item Tax Code",
            "RC": f"t_8_{category_id}_Country/Region of Origin",
            "TE": f"t_8_{category_id}_Manufacturer", "TF": f"t_8_{category_id}_EU Responsible person",
        }
        if column in generic_keys:
            return self.dropdown_values(generic_keys[column])
        if column == "RD":
            return self.dropdown_values(f"t_8_{row.get('RC', '')}_Province of Origin")
        if column in {"DS", "DT", "DU"}:
            return self.dropdown_values(f"t_3_{category_id}_121 - Material")
        if column in {"FO", "FP", "FQ"}:
            return self.dropdown_values(f"t_3_{category_id}_1920 - Major Material")
        if column in {"EX", "EY", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG"}:
            return self.dropdown_values(f"t_3_{category_id}_8319 - Food Contact Material")
        if column in {"FI", "FJ", "FK", "FL", "FM"}:
            return self.dropdown_values(f"t_3_{category_id}_{row.get('FH', '')}_8319 - Food Contact Material")
        if header:
            return self.dropdown_values(f"t_3_{category_id}_{header}") or self.dropdown_values(f"t_3_{header}")
        return []

    def required_columns(self, category_id: str) -> set[str]:
        return set(self.rules.get(category_id, {}).get("require", set()))

    def conditional_columns(self, category_id: str) -> set[str]:
        return set(self.rules.get(category_id, {}).get("condition_require", set()))


class OreshakClient:
    def __init__(self, config: Mapping[str, Any]):
        self.config = config
        self.delay = float(config.get("request_delay_seconds", 1.2))
        self.timeout = int(config.get("request_timeout_seconds", 35))
        self.max_retries = int(config.get("max_retries", 5))
        if cloudscraper is not None:
            self.scraper = cloudscraper.create_scraper(
                browser={"browser": "chrome", "platform": "windows", "mobile": False},
                delay=10,
            )
        else:
            logging.warning("cloudscraper is not installed; falling back to requests.Session")
            self.scraper = requests.Session()
        self.scraper.headers.update({
            "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.8",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Cache-Control": "no-cache",
        })
        self.last_request = 0.0

    def get(self, url: str) -> str:
        last_error: Exception | None = None
        for attempt in range(1, self.max_retries + 1):
            elapsed = time.monotonic() - self.last_request
            if elapsed < self.delay:
                time.sleep(self.delay - elapsed)
            try:
                response = self.scraper.get(url, timeout=self.timeout, allow_redirects=True)
                self.last_request = time.monotonic()
                text = response.text
                challenge = any(marker in text.lower() for marker in (
                    "please wait while your request is being verified", "cf-chl-", "just a moment", "loader"
                )) and len(text) < 15000
                if response.status_code == 200 and not challenge and len(text) > 500:
                    return text
                raise RuntimeError(f"HTTP {response.status_code}; challenge={challenge}; length={len(text)}")
            except Exception as exc:  # network retry boundary
                last_error = exc
                sleep_for = min(30, 2 ** attempt + random.random())
                logging.warning("Fetch attempt %s/%s failed for %s: %s", attempt, self.max_retries, url, exc)
                time.sleep(sleep_for)
        raise RuntimeError(f"Could not fetch {url}: {last_error}")

    def discover_product_links(self, category_url: str, max_pages: int = 100) -> list[str]:
        discovered: list[str] = []
        seen: set[str] = set()
        empty_pages = 0
        for page in range(1, max_pages + 1):
            page_url = add_query(category_url, limit=100, page=page)
            html = self.get(page_url)
            soup = BeautifulSoup(html, "lxml")
            links: list[str] = []
            selectors = [
                ".product-layout .product-thumb h4 a[href]", ".product-grid .product-thumb h4 a[href]",
                ".product-list .product-thumb h4 a[href]", ".product-thumb .caption h4 a[href]",
                "#content .product-thumb a[href]",
            ]
            for selector in selectors:
                for tag in soup.select(selector):
                    href = urljoin(category_url, tag.get("href", ""))
                    if self._looks_like_product(href, category_url):
                        links.append(href)
            # Fallback for themes with no standard OpenCart classes.
            if not links:
                for tag in soup.select("#content a[href]"):
                    href = urljoin(category_url, tag.get("href", ""))
                    text = normalize_space(tag.get_text(" ", strip=True))
                    if text and self._looks_like_product(href, category_url):
                        links.append(href)
            new_links = [link for link in dedupe(links) if link not in seen]
            for link in new_links:
                seen.add(link)
                discovered.append(link)
            logging.info("Category %s page %s: %s new product links", category_url, page, len(new_links))
            if not new_links:
                empty_pages += 1
            else:
                empty_pages = 0
            pagination_text = normalize_space(soup.get_text(" ", strip=True))
            pages_match = re.search(r"\((\d+)\s+Страници\)", pagination_text, re.I)
            if pages_match and page >= int(pages_match.group(1)):
                break
            if empty_pages >= 2:
                break
        return discovered

    @staticmethod
    def _looks_like_product(href: str, category_url: str) -> bool:
        parsed = urlparse(href)
        if parsed.netloc and parsed.netloc != urlparse(BASE_URL).netloc:
            return False
        lowered = href.lower()
        if any(part in lowered for part in ("/index.php?route=product/category", "information/", "account/", "checkout/")):
            return False
        if "product_id=" in lowered:
            return True
        category_path = urlparse(category_url).path.rstrip("/") + "/"
        path = parsed.path.rstrip("/")
        return path.startswith(category_path) and path != category_path.rstrip("/")

    def parse_product(self, url: str, category_url: str) -> Product:
        html = self.get(url)
        soup = BeautifulSoup(html, "lxml")
        product = Product(url=url, source_category_url=category_url)
        product.title = normalize_space(self._first_text(soup, ["#content h1", "h1", ".product-info h1"]))
        json_ld = self._json_ld_product(soup, page_url=url, title_hint=product.title)
        if not product.title:
            product.title = normalize_space(str(json_ld.get("name") or ""))
        if not product.title:
            raise ValueError("Product title not found")
        product.code = self._parse_code(soup, html, url)
        product.source_category_name = self._parse_source_category(soup)
        product.description = self._parse_description(soup, json_ld)
        product.attributes = self._parse_attributes(soup, product.description)
        product.bullet_points = self._make_bullets(product)
        product.tags = self._parse_tags(soup)
        product.images = self._parse_images(soup, json_ld, url)
        (
            product.price_eur,
            product.list_price_eur,
            product.price_source,
            product.list_price_source,
        ) = self._parse_prices(soup, json_ld, product.code, product.title)
        product.in_stock, product.stock_source, product.stock_quantity = self._parse_stock(soup, json_ld)
        product.options = self._parse_options(soup)
        source_text = product.description + " " + " ".join(product.attributes.values())
        product.weight_g = parse_weight(source_text)
        product.dimensions_cm = parse_dimensions(source_text)
        product.weight_source = "product description/attributes" if product.weight_g is not None else "missing"
        product.dimensions_source = "product description/attributes" if product.dimensions_cm is not None else "missing"
        if not product.code:
            product.code = self._generated_code(url, product.title)
            product.warnings.append("Product code was generated from URL/title")
        if not product.images:
            product.warnings.append("No product gallery image found")
        if product.price_eur is None:
            product.warnings.append("EUR price not found")
        measurement_mode = str(self.config.get("package_measurement_mode", "strict")).strip().casefold()
        strict_measurements = measurement_mode == "strict" or bool(
            self.config.get("omit_rows_with_fallback_measurements", True)
        )
        if product.weight_g is None:
            product.warnings.append(
                "Weight is not published; row will be omitted in strict mode"
                if strict_measurements
                else "Weight is not published; estimated package weight will use a category fallback and be marked REVIEW"
            )
        if product.dimensions_cm is None:
            product.warnings.append(
                "Complete three-dimensional size is not published; row will be omitted in strict mode"
                if strict_measurements
                else "Complete 3D size is not published; estimated package dimensions will use a category fallback and be marked REVIEW"
            )
        if product.list_price_eur is not None and product.price_eur is not None and product.list_price_eur <= product.price_eur:
            product.list_price_eur = None
            product.list_price_source = "discarded because it was not higher than base price"
            product.warnings.append("Published list price was not higher than base price; Temu list price set to N/A")
        return product

    @staticmethod
    def _json_ld_product(soup: BeautifulSoup, page_url: str = "", title_hint: str = "") -> dict[str, Any]:
        candidates: list[dict[str, Any]] = []
        for script in soup.select('script[type="application/ld+json"]'):
            try:
                data = json.loads(script.string or script.get_text())
            except Exception:
                continue
            queue = data if isinstance(data, list) else [data]
            for item in queue:
                if not isinstance(item, dict):
                    continue
                if item.get("@type") == "Product":
                    candidates.append(item)
                graph = item.get("@graph")
                if isinstance(graph, list):
                    candidates.extend(x for x in graph if isinstance(x, dict) and x.get("@type") == "Product")
        if not candidates:
            return {}

        wanted_title = canonical(title_hint)
        wanted_path = urlparse(page_url).path.rstrip("/")

        def score(item: Mapping[str, Any]) -> int:
            points = 0
            name = canonical(item.get("name"))
            if wanted_title and name == wanted_title:
                points += 20
            elif wanted_title and name and (wanted_title in name or name in wanted_title):
                points += 8
            urls: list[str] = []
            for key in ("url", "@id"):
                value = item.get(key)
                if isinstance(value, str):
                    urls.append(value)
            main = item.get("mainEntityOfPage")
            if isinstance(main, str):
                urls.append(main)
            elif isinstance(main, dict):
                for key in ("@id", "url"):
                    if isinstance(main.get(key), str):
                        urls.append(str(main[key]))
            if wanted_path and any(urlparse(value).path.rstrip("/") == wanted_path for value in urls):
                points += 15
            if item.get("offers"):
                points += 2
            return points

        return max(candidates, key=score)

    @staticmethod
    def _first_text(soup: BeautifulSoup, selectors: Sequence[str]) -> str:
        for selector in selectors:
            tag = soup.select_one(selector)
            if tag:
                return normalize_space(tag.get_text(" ", strip=True))
        return ""

    @staticmethod
    def _parse_code(soup: BeautifulSoup, html: str, url: str) -> str:
        text = normalize_space(soup.get_text(" ", strip=True))
        for pattern in (
            r"Код\s+на\s+продукта\s*:\s*([\w.\-/]+)", r"Product\s+Code\s*:\s*([\w.\-/]+)",
            r"Модел\s*:\s*([\w.\-/]+)", r"Model\s*:\s*([\w.\-/]+)",
        ):
            match = re.search(pattern, text, re.I)
            if match:
                return match.group(1).strip()
        match = re.search(r"product_id=(\d+)", url + " " + html)
        return match.group(1) if match else ""

    @staticmethod
    def _generated_code(url: str, title: str) -> str:
        path = slug_key(url)
        core = re.sub(r"[^a-z0-9]+", "-", unicodedata.normalize("NFKD", path).encode("ascii", "ignore").decode()).strip("-")
        if not core:
            core = str(abs(hash(title)) % 10_000_000)
        return f"OR-{core[:45]}".upper()

    @staticmethod
    def _parse_source_category(soup: BeautifulSoup) -> str:
        for selector in ("ul.breadcrumb li:nth-last-child(2) a", ".breadcrumb li:nth-last-child(2) a"):
            tag = soup.select_one(selector)
            if tag:
                return normalize_space(tag.get_text(" ", strip=True))
        text = normalize_space(soup.get_text(" ", strip=True))
        match = re.search(r"Категория\s+продукт\s*:\s*([^➔]+?)(?:Описание|$)", text, re.I)
        return normalize_space(match.group(1)) if match else ""

    @staticmethod
    def _parse_description(soup: BeautifulSoup, json_ld: Mapping[str, Any]) -> str:
        sections: list[str] = []
        for selector in ("#tab-description", ".tab-content #tab-description", ".product-description", "[itemprop='description']"):
            tag = soup.select_one(selector)
            if tag:
                sections.append(normalize_space(tag.get_text(" ", strip=True)))
                break
        if not sections:
            content = soup.select_one("#content")
            if content:
                text = normalize_space(content.get_text(" ", strip=True))
                start = re.search(r"Описание на продукта\s*:", text, re.I)
                end = re.search(r"Напишете отзив|Продукти на фокус", text, re.I)
                if start:
                    text = text[start.end(): end.start() if end else None]
                sections.append(text)
        if not any(sections) and json_ld.get("description"):
            sections.append(normalize_space(str(json_ld["description"])))
        description = normalize_space(" ".join(sections))
        description = re.sub(r"^.*?➔\s*Описание на продукта\s*:\s*", "", description, flags=re.I)
        description = re.split(r"➔\s*Подобни\s*:", description, maxsplit=1, flags=re.I)[0]
        return normalize_space(description)[:2000]

    @staticmethod
    def _parse_attributes(soup: BeautifulSoup, description: str) -> dict[str, str]:
        attrs: dict[str, str] = {}
        for row in soup.select("table tr"):
            cells = [normalize_space(cell.get_text(" ", strip=True)) for cell in row.select("th,td")]
            if len(cells) >= 2 and cells[0] and cells[1]:
                attrs[cells[0].rstrip(":")] = cells[1]
        for block in soup.select(".product-info .list-unstyled li, #content .attribute, .specification li"):
            text = normalize_space(block.get_text(" ", strip=True))
            if ":" in text:
                key, value = text.split(":", 1)
                if 1 <= len(key) <= 60 and value.strip():
                    attrs.setdefault(key.strip(), value.strip())

        characteristics = re.search(
            r"(?:➔\s*)?Характеристики\s*:\s*(.*?)(?=(?:➔\s*)?Предимства\s*:|(?:➔\s*)?Подобни\s*:|$)",
            description,
            re.I,
        )
        if characteristics:
            attrs.setdefault("Характеристики", normalize_space(characteristics.group(1)))

        field_patterns = {
            "Материал": r"Материал\s*:\s*(.+?)(?=\s*(?:,|;|➔|$))",
            "Размери": r"Размер(?:и)?(?:\s+на\s+[^:;,➔]{1,50})?\s*:\s*(\d+(?:[.,]\d+)?\s*[x/×]\s*\d+(?:[.,]\d+)?(?:\s*[x/×]\s*\d+(?:[.,]\d+)?)?\s*см)",
            "Дебелина": r"Дебелина(?:\s+на\s+[^:;,➔]{1,50})?\s*:\s*(\d+(?:[.,]\d+)?\s*см)",
            "Тегло": r"Тегло\s*:\s*(\d+(?:[.,]\d+)?\s*(?:кг|kg|гр|г|g))",
            "Дължина": r"Дължина(?:\s+на\s+[^:;,➔]{1,50})?\s*:\s*(\d+(?:[.,]\d+)?\s*см)",
            "Диаметър": r"Диаметър(?:\s+на\s+[^:;,➔]{1,50})?\s*:\s*(\d+(?:[.,]\d+)?\s*см)",
        }
        for heading, pattern in field_patterns.items():
            match = re.search(pattern, description, re.I)
            if match:
                attrs.setdefault(heading, normalize_space(match.group(1)))
        return attrs

    @staticmethod
    def _make_bullets(product: Product) -> list[str]:
        bullets: list[str] = []
        for key, value in product.attributes.items():
            if value and len(bullets) < 6:
                bullets.append(f"{key}: {value}"[:700])
        for sentence in re.split(r"(?<=[.!?])\s+", product.description):
            sentence = normalize_space(sentence)
            if 25 <= len(sentence) <= 300 and sentence not in bullets:
                bullets.append(sentence[:700])
            if len(bullets) >= 6:
                break
        if not bullets:
            bullets = ["Ръчно изработен продукт от България."]
        return bullets[:6]

    @staticmethod
    def _parse_tags(soup: BeautifulSoup) -> list[str]:
        tags: list[str] = []
        for tag in soup.select("a[href*='tag='], .tags a"):
            tags.append(normalize_space(tag.get_text(" ", strip=True)))
        return dedupe(tags)

    @staticmethod
    def _main_product_scope(soup: BeautifulSoup) -> Tag:
        direct = soup.select_one("#content .product-info, #content .product-page, #content .product-right")
        if isinstance(direct, Tag):
            return direct
        heading = soup.select_one("#content h1, h1")
        node: Tag | None = heading if isinstance(heading, Tag) else None
        best: Tag | None = node
        for _ in range(7):
            if node is None:
                break
            text = normalize_space(node.get_text(" ", strip=True))
            if node.select_one("#button-cart, button[id*='cart'], input[id*='cart']") or re.search(r"Код на продукта|Product Code", text, re.I):
                best = node
                if node.select_one("#button-cart, button[id*='cart']"):
                    break
            node = node.parent if isinstance(node.parent, Tag) else None
        return best or soup.select_one("#content") or soup

    @staticmethod
    def _normalize_image_url(url: str) -> str:
        parsed = urlparse(url)
        path = quote(unquote(parsed.path), safe="/@:+-._~!$&'()*+,;=")
        return urlunparse(parsed._replace(path=path, fragment=""))

    @staticmethod
    def _image_identity(url: str) -> str:
        parsed = urlparse(url)
        path = unquote(parsed.path).casefold()
        stem, dot, ext = path.rpartition(".")
        if not dot:
            stem, ext = path, ""
        stem = re.sub(r"-(?:\d{2,4}x\d{2,4})-(?:product_(?:popup|thumb)|popup|thumb)$", "", stem)
        stem = re.sub(r"-(?:product_(?:popup|thumb)|popup|thumb)$", "", stem)
        return stem + ("." + ext if ext else "")

    @staticmethod
    def _image_score(url: str) -> tuple[int, int]:
        lower = unquote(url).casefold()
        dims = re.search(r"-(\d{2,4})x(\d{2,4})-", lower)
        pixels = int(dims.group(1)) * int(dims.group(2)) if dims else 0
        quality = 3 if "product_popup" in lower else 1 if "product_thumb" in lower or "thumb" in lower else 2
        return quality, pixels

    @classmethod
    def _parse_images(cls, soup: BeautifulSoup, json_ld: Mapping[str, Any], base_url: str) -> list[str]:
        urls: list[str] = []
        json_images = json_ld.get("image")
        if isinstance(json_images, str):
            urls.append(urljoin(base_url, json_images))
        elif isinstance(json_images, list):
            urls.extend(urljoin(base_url, str(image)) for image in json_images)
        gallery_selectors = (
            "ul.thumbnails a.thumbnail[href]", ".thumbnails a[href]", ".image-additional a[href]",
            ".product-image a[href]", "#content .product-info a.thumbnail[href]",
        )
        for selector in gallery_selectors:
            for tag in soup.select(selector):
                href = tag.get("href") or tag.get("data-zoom-image") or tag.get("data-image")
                if href:
                    urls.append(urljoin(base_url, href))
                img = tag.find("img")
                if img:
                    src = img.get("data-zoom-image") or img.get("data-src") or img.get("src")
                    if src:
                        urls.append(urljoin(base_url, src))
        for tag in soup.select("#content .product-info img, #content .col-sm-4 img"):
            if tag.find_parent(id="tab-description"):
                continue
            src = tag.get("data-zoom-image") or tag.get("data-src") or tag.get("src")
            if src:
                urls.append(urljoin(base_url, src))

        # Current Oreshak pages use several gallery wrappers. Collect image links
        # broadly from the main content, but reject related/product-card images.
        excluded_ancestor = re.compile(r"related|featured|product-grid|product-thumb|carousel.*product", re.I)
        for tag in soup.select("#content a[href]"):
            if tag.find_parent(class_=excluded_ancestor):
                continue
            href = tag.get("href") or ""
            if re.search(r"/image/(?:cache/)?catalog/.*\.(?:jpe?g|png|webp)(?:\?|$)", href, re.I):
                urls.append(urljoin(base_url, href))
        for tag in soup.select("#content img"):
            if tag.find_parent(class_=excluded_ancestor) or tag.find_parent(id="tab-description"):
                continue
            src = tag.get("data-zoom-image") or tag.get("data-large") or tag.get("data-src") or tag.get("src")
            if src and "/image/" in src:
                urls.append(urljoin(base_url, src))

        best: dict[str, str] = {}
        order: list[str] = []
        for raw_url in urls:
            url = cls._normalize_image_url(raw_url)
            lower = url.casefold()
            if not lower.startswith(("http://", "https://")):
                continue
            if any(bad in lower for bad in ("logo", "no_image", "placeholder", "facebook", "loader", "icon")):
                continue
            if not re.search(r"\.(?:jpe?g|png|webp)(?:\?|$)", lower):
                continue
            identity = cls._image_identity(url)
            if identity not in best:
                best[identity] = url
                order.append(identity)
            elif cls._image_score(url) > cls._image_score(best[identity]):
                best[identity] = url
        return [best[key] for key in order]

    @classmethod
    def _parse_prices(
        cls,
        soup: BeautifulSoup,
        json_ld: Mapping[str, Any],
        product_code: str = "",
        title_hint: str = "",
    ) -> tuple[Decimal | None, Decimal | None, str, str]:
        """Extract the current and genuine pre-promotion prices from the main product panel.

        Oreshak renders a broad ``.product-right`` column that can also contain
        "Продукти на фокус".  Those cards may use the same ``h2.price`` markup as
        the real product.  Therefore DOM order and proximity to the cart button are
        not sufficient.  The parser first isolates the smallest panel containing
        the current product code (or title) together with the add-to-cart control,
        then ranks price candidates by their distance to those anchors.
        """
        broad_scope = cls._main_product_scope(soup)
        excluded_re = re.compile(
            r"(?:^|[-_\s])(?:related(?:-products?)?|featured(?:-products?)?|"
            r"product-grid|product-thumb|focus-products?|recommendations?|"
            r"similar-products?|also-(?:bought|like)|cross-sell|up-sell|"
            r"owl-carousel|swiper)(?:$|[-_\s])|carousel.*product|product.*carousel",
            re.I,
        )

        def excluded(tag: Tag) -> bool:
            return bool(
                tag.find_parent(class_=excluded_re)
                or tag.find_parent(id=excluded_re)
                or (tag.get("class") and excluded_re.search(" ".join(tag.get("class") or [])))
                or (tag.get("id") and excluded_re.search(str(tag.get("id"))))
            )

        def euro_values(tag: Tag) -> list[Decimal]:
            values: list[Decimal] = []
            for attr in ("value", "content", "data-price", "data-special"):
                raw = tag.get(attr)
                if raw in (None, ""):
                    continue
                currency = normalize_space(
                    tag.get("currency") or tag.get("data-currency") or "EUR"
                ).upper()
                if currency not in ("", "EUR"):
                    continue
                value = safe_decimal(raw)
                if value is not None:
                    values.append(value)
            text = normalize_space(tag.get_text(" ", strip=True))
            for raw in re.findall(r"([0-9][0-9\s.,]*)\s*€", text):
                value = safe_decimal(raw)
                if value is not None:
                    values.append(value)
            unique: list[Decimal] = []
            for value in values:
                if value not in unique:
                    unique.append(value)
            return unique

        def visible_euro_values(tag: Tag) -> list[Decimal]:
            """Return only EUR amounts that are visibly rendered as text.

            Hidden ``value``/``data-price`` attributes on the live Oreshak page
            can contain customer-group or wholesale prices (for example 20%
            below the public price).  Those values must not be mixed into the
            visible price cluster beside the product code.
            """
            text = normalize_space(tag.get_text(" ", strip=True))
            values: list[Decimal] = []
            for raw in re.findall(r"([0-9][0-9\s.,]*)\s*€", text):
                value = safe_decimal(raw)
                if value is not None and value not in values:
                    values.append(value)
            return values

        def text_anchor(pattern: re.Pattern[str]) -> Tag | None:
            for text_node in broad_scope.find_all(string=pattern):
                parent = text_node.parent if isinstance(text_node.parent, Tag) else None
                if parent is not None and not excluded(parent):
                    return parent
            return None

        def exact_product_code_anchor(exact_code: str) -> Tag | None:
            """Return the smallest tag that contains the full code label and value.

            On the live site the label and the numeric value can be separate text
            nodes inside one ``li``.  Looking only at the individual text node
            therefore loses the exact-code anchor and makes unrelated focus-product
            prices look equally close.
            """
            exact_code = normalize_space(exact_code)
            if not exact_code:
                return None
            combined = re.compile(
                rf"(?:Код\s+на\s+продукта|Product\s+Code)\s*[:#]?\s*{re.escape(exact_code)}(?=$|\s|[^0-9A-Za-zА-Яа-я])",
                re.I,
            )
            candidates: list[tuple[int, int, int, Tag]] = []
            for tag in broad_scope.find_all(("li", "span", "p", "small", "strong", "td", "dd", "div")):
                if not isinstance(tag, Tag) or excluded(tag):
                    continue
                text = normalize_space(tag.get_text(" ", strip=True))
                if not combined.search(text):
                    continue
                descendant_tags = sum(1 for child in tag.descendants if isinstance(child, Tag))
                depth = len(list(tag.parents))
                candidates.append((len(text), descendant_tags, -depth, tag))
            return min(candidates, key=lambda item: item[:3])[3] if candidates else None

        code_anchor = exact_product_code_anchor(product_code) or text_anchor(
            re.compile(r"Код\s+на\s+продукта|Product\s+Code", re.I)
        )
        cart_anchor = broad_scope.select_one(
            "#button-cart, button[id*='cart'], input[id*='cart'], button[name*='cart']"
        )
        title_anchor = broad_scope.select_one("h1")
        if title_hint:
            wanted = normalize_space(title_hint).casefold()
            for heading in broad_scope.select("h1, h2"):
                if normalize_space(heading.get_text(" ", strip=True)).casefold() == wanted:
                    title_anchor = heading
                    break

        def contains(node: Tag, target: Tag | None) -> bool:
            if target is None:
                return False
            return node is target or any(desc is target for desc in node.descendants)

        price_probe_selector = (
            "input#price[value], input[name='price'][value], .price-new, .special-price, "
            "[itemprop='price'], .price-old, .old-price, del, s, "
            "ul.list-unstyled h2, .product-price h2, h2.price, .product-price, .price"
        )

        def adjacent_code_price_values(anchor: Tag | None) -> tuple[list[Decimal], str]:
            """Read the price block immediately preceding the exact code block.

            The current Oreshak layout places the product price list and the
            ``Код на продукта`` list in adjacent sibling containers.  A focus
            product can be structurally closer in the ancestor tree, so tree
            distance alone is not reliable.  The nearest preceding sibling with
            EUR values is a much stronger live-page signal.
            """
            if anchor is None or not product_code:
                return [], ""
            current: Tag | None = anchor
            for level in range(7):
                parent = current.parent if isinstance(current, Tag) and isinstance(current.parent, Tag) else None
                if parent is None:
                    break
                siblings = [child for child in parent.children if isinstance(child, Tag)]
                try:
                    index = siblings.index(current)
                except ValueError:
                    index = 0
                # Nearby previous siblings are examined from nearest to farthest.
                for offset, sibling in enumerate(reversed(siblings[:index]), start=1):
                    if offset > 5 or excluded(sibling):
                        continue
                    sibling_text = normalize_space(sibling.get_text(" ", strip=True))
                    if re.search(
                        r"Продукти\s+на\s+фокус|Подобни\s+продукти|Related|Featured|Recommended",
                        sibling_text,
                        re.I,
                    ):
                        continue
                    values = visible_euro_values(sibling)
                    if not values:
                        continue
                    # Product-card grids usually contain several outbound links;
                    # the real price list is normally a compact ul/div without them.
                    product_links = sibling.select(
                        "a[href*='/product/'], a[href*='product_id='], .product-thumb a[href], .product-grid a[href]"
                    )
                    if len(product_links) > 1:
                        continue
                    unique: list[Decimal] = []
                    for value in values:
                        if value not in unique:
                            unique.append(value)
                    if unique:
                        return unique, f"exact product-code adjacent price block (ancestor level {level}, previous sibling {offset})"
                current = parent
                if current is broad_scope:
                    break
            return [], ""

        code_adjacent_values, code_adjacent_source = adjacent_code_price_values(code_anchor)

        def smallest_panel(anchor: Tag | None, companion: Tag | None) -> Tag | None:
            node = anchor
            while isinstance(node, Tag):
                if excluded(node):
                    node = node.parent if isinstance(node.parent, Tag) else None
                    continue
                if (companion is None or contains(node, companion)) and node.select_one(price_probe_selector):
                    return node
                if node is broad_scope:
                    break
                node = node.parent if isinstance(node.parent, Tag) else None
            return None

        # Product code is the strongest anchor; title is the fallback.
        purchase_scope = (
            smallest_panel(code_anchor, cart_anchor)
            or smallest_panel(title_anchor, cart_anchor)
            or smallest_panel(code_anchor, None)
            or broad_scope
        )

        def dom_distance(left: Tag, right: Tag | None) -> int:
            if right is None:
                return 999
            left_path: list[Tag] = []
            node: Tag | None = left
            while isinstance(node, Tag):
                left_path.append(node)
                if node is broad_scope:
                    break
                node = node.parent if isinstance(node.parent, Tag) else None
            right_index: dict[int, int] = {}
            node = right
            depth = 0
            while isinstance(node, Tag):
                right_index[id(node)] = depth
                if node is broad_scope:
                    break
                node = node.parent if isinstance(node.parent, Tag) else None
                depth += 1
            for left_depth, ancestor in enumerate(left_path):
                if id(ancestor) in right_index:
                    return left_depth + right_index[id(ancestor)]
            return 999

        def best_tag(selectors: Sequence[str]) -> tuple[Tag | None, str]:
            best: tuple[int, int, Tag, str] | None = None
            for selector_rank, selector in enumerate(selectors):
                for order, tag in enumerate(purchase_scope.select(selector)):
                    if not isinstance(tag, Tag) or excluded(tag) or not euro_values(tag):
                        continue
                    code_distance = dom_distance(tag, code_anchor)
                    cart_distance = dom_distance(tag, cart_anchor)
                    title_distance = dom_distance(tag, title_anchor)
                    score = (
                        100000
                        - code_distance * 1000
                        - cart_distance * 100
                        - title_distance * 10
                        - selector_rank * 3
                        - order
                    )
                    # Strong bonus when the candidate's nearest list/container also
                    # contains the exact product-code marker.
                    node: Tag | None = tag
                    for depth in range(5):
                        if node is None:
                            break
                        node_text = normalize_space(node.get_text(" ", strip=True))
                        if product_code and re.search(
                            rf"(?:Код\s+на\s+продукта|Product\s+Code)\s*[:#]?\s*{re.escape(product_code)}(?:\D|$)",
                            node_text,
                            re.I,
                        ):
                            score += 5000 - depth * 250
                            break
                        node = node.parent if isinstance(node.parent, Tag) else None
                    candidate = (score, -order, tag, selector)
                    if best is None or candidate[:2] > best[:2]:
                        best = candidate
            if best is None:
                return None, ""
            return best[2], best[3]

        # Product JSON-LD is the strongest source for the currently public
        # offer price because it is tied to the matched Product entity.  Read it
        # before generic DOM inputs, which may contain customer-group prices.
        structured_price: Decimal | None = None
        offers = json_ld.get("offers")
        offer_items = offers if isinstance(offers, list) else [offers] if isinstance(offers, dict) else []
        for offer in offer_items:
            if not isinstance(offer, dict):
                continue
            currency = normalize_space(str(offer.get("priceCurrency") or "EUR")).upper()
            if currency not in ("", "EUR"):
                continue
            value = safe_decimal(offer.get("price") or offer.get("lowPrice"))
            if value is not None:
                structured_price = value
                break

        # Current/discounted price from explicit checkout or special-price markup.
        current_tag, current_selector = best_tag((
            "input#price[value]",
            "input[name='price'][value]",
            ".price-new",
            ".special-price",
            "[itemprop='price'][content]",
            "meta[itemprop='price'][content]",
        ))
        current_values = euro_values(current_tag) if current_tag is not None else []
        explicit_price = current_values[0] if current_values else None
        if structured_price is not None:
            price = structured_price
            price_source = "matched Product JSON-LD offer"
        else:
            price = explicit_price
            price_source = (
                f"anchored main product selector: {current_selector}"
                if price is not None else ""
            )

        old_tag, old_selector = best_tag((".price-old", ".old-price", "del", "s"))
        old_values = euro_values(old_tag) if old_tag is not None else []
        list_price = max(old_values) if old_values else None
        list_source = (
            f"anchored main product selector: {old_selector}"
            if list_price is not None else ""
        )

        # The exact-code adjacent block has priority over generic visible-price
        # selectors.  This handles the live layout where current and old prices
        # are in one ul and the code is in the next ul, while an unrelated focus
        # price sits elsewhere in the same broad product-right column.
        if code_adjacent_values:
            visible_values = code_adjacent_values
            visible_source = code_adjacent_source
        else:
            visible_tag, visible_selector = best_tag((
                "ul.list-unstyled h2",
                ".product-price h2",
                "h2.price",
                ".product-price",
                ".price",
            ))
            visible_values = euro_values(visible_tag) if visible_tag is not None else []
            visible_source = (
                f"anchored main product selector: {visible_selector}"
                if visible_values else ""
            )

        if visible_values:
            visible_current = min(visible_values)
            if price is None:
                price = visible_current
                price_source = visible_source

            # Select the nearest higher visible amount, not the maximum amount
            # in a broad container.  This limits contamination from unrelated
            # focus cards while preserving the genuine public pre-promotion
            # price immediately above the current Product offer.
            if price is not None:
                higher_candidates = sorted(value for value in visible_values if value > price)
                visible_regular = higher_candidates[0] if higher_candidates else None
            else:
                visible_regular = max(visible_values)

            if visible_regular is not None:
                if list_price is None or visible_regular < list_price:
                    list_price = visible_regular
                    if len(visible_values) > 1:
                        list_source = visible_source + " (nearest higher published pre-promotion price)"
                    elif price_source.startswith("anchored main product selector: input"):
                        list_source = visible_source + " (higher than checkout price)"
                    else:
                        list_source = visible_source + " (higher than current Product offer)"

        # Do not manufacture a list price. Temu supports N/A when no genuine
        # previous/list price is published.
        if price is not None and list_price is not None and list_price <= price:
            list_price = None
            list_source = "no genuine higher list price found"
        return price, list_price, price_source, list_source

    @classmethod
    def _parse_stock(cls, soup: BeautifulSoup, json_ld: Mapping[str, Any]) -> tuple[bool, str, int | None]:
        offers = json_ld.get("offers")
        offer_items = offers if isinstance(offers, list) else [offers] if isinstance(offers, dict) else []
        for offer in offer_items:
            if not isinstance(offer, dict):
                continue
            availability = str(offer.get("availability") or "").casefold()
            if "outofstock" in availability or "soldout" in availability:
                return False, "matched Product JSON-LD availability", 0
            if "instock" in availability or "limitedavailability" in availability:
                break

        scope = cls._main_product_scope(soup)
        text = normalize_space(scope.get_text(" ", strip=True))
        if re.search(r"последна\s+бройка|last\s+(?:item|piece)", text, re.I):
            detected_qty = 1
        else:
            qty_match = re.search(r"(?:налични|в наличност|available)\s*[:\-]?\s*(\d+)\s*(?:бр|pcs|pieces)?", text, re.I)
            detected_qty = int(qty_match.group(1)) if qty_match else None

        lowered = text.casefold()
        unavailable = (
            "няма наличност", "неналичен", "неналична", "изчерпан",
            "out of stock", "sold out", "не е наличен", "не е налична",
        )
        if any(marker in lowered for marker in unavailable):
            return False, "main product availability text", 0

        button = scope.select_one("#button-cart, button[id*='cart'], input[id*='cart']")
        if isinstance(button, Tag):
            disabled = button.has_attr("disabled") or canonical(button.get("aria-disabled")) == "true"
            button_text = normalize_space(button.get_text(" ", strip=True) or button.get("value") or "").casefold()
            if disabled or any(word in button_text for word in ("неналич", "изчерпан", "out of stock", "sold out")):
                return False, "main add-to-cart control", 0
            return True, "enabled main add-to-cart control", detected_qty

        return True, "no out-of-stock marker in main product area", detected_qty

    @staticmethod
    def _parse_options(soup: BeautifulSoup) -> list[ProductOption]:
        options: list[ProductOption] = []
        excluded = re.compile(r"гравир|персонализ|текст|място|engraving|personal", re.I)
        for group in soup.select("#product .form-group, .product-options .form-group"):
            label_tag = group.select_one("label.control-label, label")
            name = normalize_space(label_tag.get_text(" ", strip=True) if label_tag else "")
            if not name or excluded.search(name):
                continue
            values: list[str] = []
            select = group.select_one("select")
            if select:
                for option in select.select("option[value]"):
                    value = normalize_space(option.get_text(" ", strip=True))
                    if value and not re.search(r"избер|select", value, re.I):
                        values.append(re.sub(r"\s*\([+\-].*?\)\s*$", "", value).strip())
            else:
                for label in group.select("label.radio, label.checkbox, .radio label, .checkbox label"):
                    value = normalize_space(label.get_text(" ", strip=True))
                    if value:
                        values.append(re.sub(r"\s*\([+\-].*?\)\s*$", "", value).strip())
            values = dedupe(values)
            if values:
                options.append(ProductOption(name=name, values=values))
        return options[:2]


def _number(value: str) -> float:
    return float(value.replace(",", "."))


def _to_cm(value: str, unit: str | None) -> float:
    number = _number(value)
    normalized = canonical(unit or "cm")
    if normalized in {"mm", "мм"}:
        return number / 10.0
    if normalized in {"m", "м"}:
        return number * 100.0
    return number


def parse_weight(text: str) -> float | None:
    # Repair common source typos such as "Т егло" and split words.
    normalized = re.sub(r"т\s*е\s*г\s*л\s*о", "тегло", text, flags=re.I)
    patterns = [
        (r"(?:тегло|weight)(?:\s+на\s+[^:;,➔]{1,60})?\s*:?[\s]*([0-9]+(?:[.,][0-9]+)?)\s*(кг|kg)\b", 1000),
        (r"(?:тегло|weight)(?:\s+на\s+[^:;,➔]{1,60})?\s*:?[\s]*([0-9]+(?:[.,][0-9]+)?)\s*(гр|грама|г|g)\b", 1),
    ]
    for pattern, multiplier in patterns:
        match = re.search(pattern, normalized, re.I)
        if match:
            value = _number(match.group(1)) * multiplier
            return round(max(value, 0.1), 1)
    return None


def parse_dimensions(text: str) -> tuple[float, float, float] | None:
    normalized = text.replace(",", ".").replace("×", "x").replace("Х", "x").replace("х", "x")
    normalized = re.sub(r"(?<=\d)\s+(?=\d\s*[/x])", "", normalized)
    # Repair source typos/split words such as "Р азмер".
    normalized = re.sub(r"р\s+азмер", "размер", normalized, flags=re.I)

    # Explicit length/width/height labels are the most reliable source.
    l = re.search(r"дължина\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)\b", normalized, re.I)
    w = re.search(r"ширина\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)\b", normalized, re.I)
    h = re.search(r"височина\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)\b", normalized, re.I)
    if l and w and h:
        vals = (_to_cm(l.group(1), l.group(2)), _to_cm(w.group(1), w.group(2)), _to_cm(h.group(1), h.group(2)))
        return tuple(round(v, 2) for v in sorted(vals, reverse=True))

    # Some pages state values as "27 см. височина / 16 см. ширина / 5 см. дебелина".
    labelled: dict[str, float] = {}
    for value, unit, label in re.findall(
        r"(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)[.\s]*(височина|ширина|дебелина|дължина)",
        normalized,
        re.I,
    ):
        labelled[canonical(label)] = _to_cm(value, unit)
    if {"височина", "ширина", "дебелина"}.issubset(labelled):
        vals = [labelled["височина"], labelled["ширина"], labelled["дебелина"]]
        return tuple(round(v, 2) for v in sorted(vals, reverse=True))

    thickness_match = re.search(
        r"дебелина(?:\s+на\s+[^:;,➔]{1,50})?\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)",
        normalized,
        re.I,
    )
    thickness = _to_cm(thickness_match.group(1), thickness_match.group(2)) if thickness_match else None
    height_match = re.search(
        r"височина(?:\s+на\s+[^:;,➔]{1,50})?\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)",
        normalized,
        re.I,
    )
    explicit_height = _to_cm(height_match.group(1), height_match.group(2)) if height_match else None
    depth_match = re.search(
        r"(?:дълбочина|depth)(?:\s+на\s+[^:;,➔]{1,50})?\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)",
        normalized,
        re.I,
    )
    explicit_depth = None
    if depth_match:
        # Internal compartment depth is not an outer product/package dimension.
        context_before = normalized[max(0, depth_match.start() - 30):depth_match.start()].casefold()
        if "вътреш" not in context_before and "internal" not in context_before:
            explicit_depth = _to_cm(depth_match.group(1), depth_match.group(2))

    # Prefer outer/closed/frame dimensions. Keep the text between the label and
    # the first number short so the regex cannot drift into a later "inner size".
    priority_patterns = [
        r"външни\s+размери[^0-9]{0,60}?(\d+(?:\.\d+)?)\s*[/x]\s*(\d+(?:\.\d+)?)(?:\s*[/x]\s*(\d+(?:\.\d+)?))?\s*(мм|mm|см|cm|м|m)",
        r"(?:размери?\s+)?(?:на\s+кутията\s*)?(?:в\s+)?затворено\s+състояние[^0-9]{0,40}?(\d+(?:\.\d+)?)\s*[/x]\s*(\d+(?:\.\d+)?)(?:\s*[/x]\s*(\d+(?:\.\d+)?))?\s*(мм|mm|см|cm|м|m)",
        r"кутия\s*:\s*в\s+затворено\s+състояние[^0-9]{0,40}?(\d+(?:\.\d+)?)\s*[/x]\s*(\d+(?:\.\d+)?)(?:\s*[/x]\s*(\d+(?:\.\d+)?))?\s*(мм|mm|см|cm|м|m)",
        r"размер\s+с\s+рамката[^0-9]{0,40}?(\d+(?:\.\d+)?)\s*[/x]\s*(\d+(?:\.\d+)?)(?:\s*[/x]\s*(\d+(?:\.\d+)?))?\s*(мм|mm|см|cm|м|m)",
        r"(?:размер(?:и)?(?:\s+на\s+(?!едно отделение|квадратите|платно)[^:;,➔]{1,50})?|dimensions?)\s*:?\s*(\d+(?:\.\d+)?)\s*[/x]\s*(\d+(?:\.\d+)?)(?:\s*[/x]\s*(\d+(?:\.\d+)?))?\s*(мм|mm|см|cm|м|m)",
    ]
    for pattern in priority_patterns:
        match = re.search(pattern, normalized, re.I)
        if not match:
            continue
        a, b, c, unit = match.groups()
        dims = [_to_cm(a, unit), _to_cm(b, unit)]
        if c:
            dims.append(_to_cm(c, unit))
        else:
            inferred = thickness or explicit_height or explicit_depth
            # Tiny square dice/game pieces are cubes. This is the only safe
            # geometric inference from two equal published measurements.
            if inferred is None and max(dims) <= 1.5 and abs(dims[0] - dims[1]) <= 0.05:
                inferred = dims[0]
            # Do not manufacture a third dimension for flat textiles, paintings,
            # spoons, boxes, etc. The writer will use a category fallback and
            # flag the row for review instead.
            if inferred is None:
                continue
            dims.append(inferred)
        return tuple(round(v, 2) for v in sorted((max(v, 0.1) for v in dims), reverse=True))

    # One overall size plus a height (common for round/square trays and ashtrays).
    single_size = re.search(r"размер(?:и)?\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)", normalized, re.I)
    if single_size and explicit_height:
        side = _to_cm(single_size.group(1), single_size.group(2))
        return tuple(round(v, 2) for v in sorted((side, side, explicit_height), reverse=True))

    diameter = re.search(r"(?:диаметър|ф|Ø)\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)?", normalized, re.I)
    if diameter:
        d = _to_cm(diameter.group(1), diameter.group(2))
        third = thickness or explicit_height or explicit_depth
        if third is None:
            return None
        return round(d, 2), round(d, 2), round(third, 2)

    # Tall narrow objects sometimes publish height and opening only.
    opening = re.search(r"отвор\s*:?\s*(\d+(?:\.\d+)?)\s*(мм|mm|см|cm|м|m)", normalized, re.I)
    if explicit_height and opening:
        d = _to_cm(opening.group(1), opening.group(2))
        return tuple(round(v, 2) for v in sorted((explicit_height, d, d), reverse=True))

    blade = re.search(r"дължина на острието\s*:?\s*(\d+(?:\.\d+)?)\s*см", normalized, re.I)
    handle = re.search(r"дължина на дръжката\s*:?\s*(\d+(?:\.\d+)?)\s*см", normalized, re.I)
    total = re.search(r"обща дължина\s*:?\s*(\d+(?:\.\d+)?)\s*см", normalized, re.I)
    if total or blade:
        length = float((total or blade).group(1))
        if not total and handle:
            length += float(handle.group(1))
        return round(length + 3, 1), 8.0, 5.0
    return None


def choose_valid(values: Sequence[str], candidates: Sequence[str], fallback_first: bool = True) -> str:
    real_values = [v for v in values if v and not v.startswith("Must first finish")]
    if not real_values:
        return ""
    by_canon = {canonical(v): v for v in real_values}
    for candidate in candidates:
        candidate_key = canonical(candidate)
        if candidate_key in by_canon:
            return by_canon[candidate_key]
    for candidate in candidates:
        candidate_key = canonical(candidate)
        if not candidate_key:
            continue
        for value in real_values:
            value_key = canonical(value)
            if candidate_key in value_key or value_key in candidate_key:
                return value
    return real_values[0] if fallback_first else ""


def detect_material_candidates(product: Product, food_contact: bool = False) -> list[str]:
    published = " ".join(re.findall(r"материал\s*:\s*([^,.;➔]{1,80})", product.description, re.I))
    structured = normalize_space(" ".join([
        published,
        product.attributes.get("Материал", ""),
        product.attributes.get("Дърво", ""),
    ])).casefold()
    full_text = normalize_space(" ".join([
        product.title,
        product.description,
        product.source_category_name,
        slug_key(product.source_category_url),
        *product.attributes.values(),
    ])).casefold()
    candidates: list[str] = []

    # Explicit "Материал:" text and structured source attributes take priority
    # over incidental component words (for example a brass latch on a wooden box).
    for source in (structured, full_text):
        for pattern, values in MATERIAL_TRANSLATIONS:
            if pattern.search(source):
                candidates.extend(values)
        if candidates and source == structured:
            break
    if not candidates:
        candidates = ["Wood", "Log"] if "дър" in full_text else ["Other", "Metal"]
    if food_contact and any(canonical(v) in {"wood", "solidwood", "naturalwood"} for v in candidates):
        candidates = ["Log", "Wood", *candidates]
    # "Other" is a safer valid dropdown fallback when the exact published
    # material (for example bone) is not offered by Temu.
    candidates.append("Other")
    return dedupe(candidates)


def ensure_product_description(product: Product) -> bool:
    """Create a source-grounded description when the product page has none.

    The fallback uses only the published title, code, bullet points and
    attributes; it does not invent marketing claims. Returns True when a
    fallback was created so the row can be marked REVIEW.
    """
    if normalize_space(product.description):
        product.description = normalize_space(product.description)
        return False
    parts: list[str] = []
    if normalize_space(product.title):
        parts.append(normalize_space(product.title).rstrip(". ") + ".")
    if normalize_space(product.code):
        parts.append(f"Код на продукта: {normalize_space(product.code)}.")
    for bullet in product.bullet_points[:6]:
        value = normalize_space(bullet)
        if value:
            parts.append(value.rstrip(". ") + ".")
    for key, value in list(product.attributes.items())[:12]:
        key_n, value_n = normalize_space(key), normalize_space(value)
        if key_n and value_n:
            parts.append(f"{key_n}: {value_n}.")
    product.description = normalize_space(" ".join(parts)) or "Информацията за продукта не е публикувана."
    product.warnings.append("Product description generated from the published title/code/attributes because the source description is empty")
    return True


def category_for(product: Product, overrides: Mapping[str, str], schema: TemplateSchema) -> tuple[str, str, str]:
    if product.code in overrides:
        return overrides[product.code], "override by product code", "high"
    if product.url in overrides:
        return overrides[product.url], "override by product URL", "high"
    # Category detection must be driven by the product identity, not generic words such as
    # "gift" or "wood" appearing later in the description.
    title_only = normalize_space(product.title).casefold()
    for pattern, reason in LOW_CONFIDENCE_PRODUCT_RULES:
        if pattern.search(title_only):
            default = SOURCE_DEFAULTS.get(slug_key(product.source_category_url), "13020")
            return default, reason, "low"
    for pattern, category_id, reason in KEYWORD_CATEGORY_RULES:
        if pattern.search(title_only) and category_id in schema.category_names:
            return category_id, reason, "high"
    source_slug = slug_key(product.source_category_url)
    if source_slug == "kuhnenski-aksesoari-ot-darvo-oreshak" and re.search(r"\bдъск", title_only) and "54423" in schema.category_names:
        return "54423", "kitchen-source cutting/serving board", "high"
    default = SOURCE_DEFAULTS.get(source_slug, "13020")
    if source_slug in STRICT_EXPLICIT_MAPPING_SLUGS:
        return default, f"no explicit safe product rule for heterogeneous source category: {source_slug}", "low"
    return default, f"source category default: {source_slug}", "medium"


def load_overrides(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    result: dict[str, str] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            category = normalize_space(row.get("temu_category_id"))
            if not category:
                continue
            for key in (normalize_space(row.get("product_code")), normalize_space(row.get("product_url"))):
                if key:
                    result[key] = category
    return result


def expand_variants(product: Product, max_variants: int = 100) -> list[Variant]:
    if not product.options:
        return [Variant(product, {}, product.code, product.title, product.images[0] if product.images else "")]
    option_names = [option.name for option in product.options]
    combinations = itertools.product(*(option.values for option in product.options))
    variants: list[Variant] = []
    for index, combination in enumerate(combinations, start=1):
        if index > max_variants:
            product.warnings.append(f"Variants truncated to {max_variants}")
            break
        values = dict(zip(option_names, combination))
        suffix = "-".join(re.sub(r"[^A-Z0-9]+", "-", unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode().upper()).strip("-")[:12] for value in combination)
        sku = f"{product.code}-{suffix}"[:80].rstrip("-")
        title = f"{product.title} - {' / '.join(combination)}"[:500]
        variants.append(Variant(product, values, sku, title, product.images[0] if product.images else ""))
    return variants


def variation_fields(variant: Variant, schema: TemplateSchema) -> dict[str, Any]:
    property_columns = {
        "color": "PG", "цвят": "PG", "size": "PH", "размер": "PH", "style": "PI", "стил": "PI",
        "material": "PJ", "материал": "PJ", "quantity": "PQ", "количество": "PQ", "model": "PR", "модел": "PR",
    }
    row: dict[str, Any] = {}
    if not variant.option_values:
        row["PF"] = choose_valid(schema.dropdown_for("PF", variant.product.category_id, row), ["Model"])
        row["PR"] = variant.product.code
        return row
    columns: list[str] = []
    theme_parts: list[str] = []
    for name, value in variant.option_values.items():
        key = canonical(name)
        column = next((col for token, col in property_columns.items() if canonical(token) in key or key in canonical(token)), "PR")
        columns.append(column)
        header = schema.headers.get(column, "Model")
        theme_parts.append(header)
        row[column] = value
    theme = " × ".join(theme_parts[:2])
    row["PF"] = choose_valid(schema.dropdown_for("PF", variant.product.category_id, row), [theme, theme_parts[0], "Model"])
    return row


FABRIC_COMPOSITION_COLUMNS: dict[str, tuple[str, ...]] = {
    "IA": ("акрил", "acrylic"),
    "IB": ("памук", "cotton"),
    "IC": ("лен", "linen"),
    "ID": ("modal", "модал"),
    "IE": ("найлон", "nylon", "полиамид"),
    "IF": ("полиестер", "polyester"),
    "IG": ("коприна", "silk"),
    "IH": ("еластан", "elastane", "spandex"),
    "II": ("вискоза", "viscose"),
    "IJ": ("ацетат", "acetate"),
    "IK": ("лиосел", "lyocell"),
    "IM": ("полиуретан", "polyurethane"),
    "IN": ("полипропилен", "polypropylene"),
    "MG": ("кашмир", "cashmere"),
    "MH": ("пера", "feather"),
    "MI": ("мохер", "mohair"),
    "MJ": ("вълна", "wool"),
    "MK": ("алпака", "alpaca"),
    "ML": ("пух", "down"),
    "MM": ("патешки пух", "duck down"),
    "MN": ("заешки косъм", "rabbit hair"),
    "MO": ("pvc", "пвц"),
    "MP": ("гъши пух", "goose down"),
    "MQ": ("кожа", "leather"),
    "MR": ("изкуствен косъм", "faux fur"),
    "MS": ("изкуствена кожа", "faux leather"),
    "MU": ("полиетилен", "polyethylene"),
    "MV": ("корк", "cork"),
}


def fabric_composition_fields(product: Product, schema: TemplateSchema) -> dict[str, Any]:
    if product.category_id != "39650":
        return {}
    text = normalize_space(" ".join([product.title, product.description, *product.attributes.values()])).casefold()
    row: dict[str, Any] = {column: 0 for column in FABRIC_COMPOSITION_COLUMNS}
    row["IL"] = 0  # Other Fibers
    row["MT"] = 0  # Non-textile Materials

    explicit: list[tuple[str, float]] = []
    for column, names in FABRIC_COMPOSITION_COLUMNS.items():
        for name in names:
            match = re.search(rf"(\d+(?:[.,]\d+)?)\s*%\s*{re.escape(name)}", text, re.I)
            if not match:
                match = re.search(rf"{re.escape(name)}\s*[:\-]?\s*(\d+(?:[.,]\d+)?)\s*%", text, re.I)
            if match:
                explicit.append((column, float(match.group(1).replace(",", "."))))
                break
    if explicit:
        for column, percent in explicit:
            row[column] = percent
    else:
        detected = next((column for column, names in FABRIC_COMPOSITION_COLUMNS.items() if any(name in text for name in names)), None)
        if detected:
            row[detected] = 100
            product.warnings.append("Fabric composition percentage inferred as 100% from the published material")
        else:
            row["IL"] = 100
            product.warnings.append("Fabric composition was not published; set to 100% Other Fibers for review")

    total = sum(float(row.get(column, 0) or 0) for column in [*FABRIC_COMPOSITION_COLUMNS, "IL", "MT"])
    if total and abs(total - 100) > 0.01:
        product.warnings.append(f"Published fabric composition totals {total:g}% and must be reviewed")
    row["MY"] = choose_valid(schema.dropdown_for("MY", product.category_id, row), ["Textile Material"])
    return row


def infer_required_value(column: str, product: Product, variant: Variant, row: dict[str, Any], schema: TemplateSchema, config: Mapping[str, Any]) -> Any:
    category_id = product.category_id
    header = schema.headers.get(column, "")
    dropdown = schema.dropdown_for(column, category_id, row)
    default_weight, default_l, default_w, default_h = DEFAULT_PACKAGE.get(category_id, (500, 30, 20, 10))
    weight = product.weight_g or default_weight
    dims = product.dimensions_cm or (default_l, default_w, default_h)
    text = " ".join([product.title, product.description, *product.attributes.values()]).casefold()

    if column == "E": return category_id
    if column == "L": return variant.title[:500]
    if column == "M": return product.code[:80]
    if column == "N": return variant.sku[:80]
    if column == "O": return choose_valid(dropdown, ["Add", "Add a new product", "New"])
    if column == "T": return product.description[:2000]
    if column == "PT": return variant.image
    if column == "QE": return (product.stock_quantity if product.stock_quantity is not None else int(config.get("default_in_stock_quantity", 10))) if product.in_stock else 0
    if column == "QF": return product.price_eur
    if column == "QG": return product.url
    if column == "QH": return product.list_price_eur
    if column == "QI": return "N/A" if product.list_price_eur is None else ""
    if column == "QJ": return round(weight, 1)
    if column == "QK": return round(max(dims), 1)
    if column == "QL": return round(sorted(dims, reverse=True)[1], 1)
    if column == "QM": return round(min(dims), 1)
    if column == "QO": return choose_valid(dropdown, ["Yes"])
    if column == "QP": return 1
    if column == "QQ": return choose_valid(dropdown, ["piece", "pack"])
    if column == "QY": return normalize_space(config.get("shipping_template")) or choose_valid(dropdown, [])
    if column == "QZ": return normalize_space(config.get("handling_time")) or choose_valid(dropdown, ["1 Day"])
    if column == "RA": return normalize_space(config.get("fulfillment_channel")) or choose_valid(dropdown, ["Seller Fulfilled", "Seller"])
    if column == "RB": return normalize_space(config.get("item_tax_code")) or choose_valid(dropdown, ["GEN STANDARD"])
    if column == "RC": return choose_valid(dropdown, [str(config.get("country_of_origin", "Bulgaria")), "Bulgaria"])
    if column == "RD": return ""  # not required for Bulgaria
    if column == "TD": return f"SKU: {variant.sku}"[:200]
    if column == "TE": return normalize_space(config.get("manufacturer_name"))
    if column == "TF": return normalize_space(config.get("eu_responsible_person"))
    if column == "PF": return variation_fields(variant, schema).get("PF")

    if "Surface Finishing Type" in header:
        candidates: list[str] = []
        if re.search(r"\bматов(?:а|о|и)?\b|\bmatte\b", text, re.I):
            candidates.append("Matte")
        if "полир" in text or "polish" in text:
            candidates.append("Polishing")
        if "четкан" in text or "brushed" in text:
            candidates.append("Brushed")
        if "прахов" in text:
            candidates.append("Powder Coating")
        if "боядис" in text or "painted" in text:
            candidates.append("Painted")
        if "лазер" in text and "грав" in text:
            candidates.append("Laser Engraving")
        if "кован" in text or "hammered" in text:
            candidates.append("Hammered")
        if not candidates and category_id in {"10059", "10072", "10638"}:
            candidates.append("Polishing")
            product.warnings.append("Surface finishing is not published; Polishing selected as the Temu-required knife fallback")
        return choose_valid(dropdown, candidates, fallback_first=False)
    if "Blade Material" in header:
        candidates: list[str] = []
        if "дамаск" in text:
            candidates.append("Damascus Steel")
        if "неръждаем" in text:
            candidates.append("Stainless Steel")
        if "високовъглерод" in text or "въглерод" in text:
            candidates.append("Carbon")
        if "стоман" in text:
            candidates.extend(["Steel", "Alloy"])
        if not candidates:
            candidates.append("Steel")
            product.warnings.append("Blade material is not published; Steel selected as the Temu-required knife fallback")
        return choose_valid(dropdown, candidates, fallback_first=False)
    if "Handle Material" in header:
        if any(token in text for token in ("дърв", "бук", "орех", "дъб", "ясен", "махагон")):
            candidates = ["Wooden handle"]
        elif any(token in text for token in ("пластмас", "plastic")):
            candidates = ["Plastic"]
        elif any(token in text for token in ("смол", "resin")):
            candidates = ["Resin"]
        elif any(token in text for token in ("еленов рог", "рог", "кост")):
            candidates = ["Resin"]
            product.warnings.append("Temu handle-material dropdown has no horn/bone option; Resin selected as the closest permitted value for manual review")
        elif "неръждаем" in text:
            candidates = ["Stainless Steel", "Metal"]
        elif any(token in text for token in ("стоман", "метал")):
            candidates = ["Metal", "Steel"]
        else:
            candidates = ["Resin"]
            product.warnings.append("Handle material is not published; Resin selected as the Temu-required fallback for manual review")
        return choose_valid(dropdown, candidates, fallback_first=False)
    if "Material Type" in header:
        return choose_valid(dropdown, ["Textile Material", "Non-textile Material"], fallback_first=False)
    if column in {"DS", "DT", "DU", "FO", "FP", "FQ"} or "Material" in header:
        return choose_valid(dropdown, detect_material_candidates(product, food_contact="Food Contact" in header), fallback_first=False)
    if "Power Supply" in header:
        return choose_valid(dropdown, ["Use Without Electricity", "Without Electricity", "Non Electric"])
    if "Battery Properties" in header:
        return choose_valid(dropdown, ["Without Battery", "No Battery", "Battery Free"])
    if "Applicable Age Group" in header:
        age = "18 Years+" if any(word in text for word in ("ловен", "нож", "трофей")) else "14 Years+"
        return choose_valid(dropdown, [age, "12 Years+", "8 Years+"])
    if "Can Be Used For Food Contact" in header:
        always_food_categories = {"9998", "9999", "10006", "10059", "10072", "10638", "10628", "54423", "10740", "10741", "10807", "10808", "11514", "10703", "9923"}
        explicit_food_use = any(token in text for token in ("за храна", "хранене", "сервиране", "food contact", "food safe", "подходяща за храна", "подходящ за храна"))
        # Novelty/decorative plates should not be declared food-safe unless the source says so.
        food = category_id in always_food_categories or (category_id == "10853" and explicit_food_use)
        return choose_valid(dropdown, ["Yes" if food else "No"])
    if "Food Contact Material" in header:
        return choose_valid(dropdown, detect_material_candidates(product, food_contact=True), fallback_first=False)
    if "Closure Type" in header:
        return choose_valid(dropdown, ["Magnetic" if "магнит" in text else "Latch", "Flip top"])
    if "Water Resistance Level" in header:
        return choose_valid(dropdown, ["Non-water resistant"])
    if "Frame Type" in header:
        framed = "рамк" in text or "frame" in text
        return choose_valid(dropdown, ["Framed", "With Frame"] if framed else ["Frameless"], fallback_first=False)
    if "Thickness" in header and "value" in schema.internal_keys.get(column, "").casefold():
        return round(min(dims), 1)
    if "Thickness" in header and "unit" in schema.internal_keys.get(column, "").casefold():
        return choose_valid(dropdown, ["cm"])
    if "Wood Type" in header:
        return choose_valid(dropdown, ["Log", "Solid Wood", "Natural Wood", "Wood"], fallback_first=False)
    if "Wood Species" in header:
        species_candidates: list[str] = []
        for bg, candidates in WOOD_SPECIES.items():
            if bg in text:
                species_candidates.extend(candidates)
        return choose_valid(dropdown, species_candidates, fallback_first=False)
    if "Stainless Steel Grade" in header:
        return choose_valid(dropdown, ["304", "18/10", "Other"])
    if "Genuine Leather Type" in header:
        return choose_valid(dropdown, ["Cowhide", "Full Grain Leather", "Genuine Leather"])
    if "Packaging unit" in header:
        return choose_valid(dropdown, ["piece"])
    if dropdown:
        return choose_valid(dropdown, ["Other", "No", "None", "Not Applicable", "Use Without Electricity"], fallback_first=False)
    # Free-text/numeric fallback. It is preferable to provide a traceable SKU
    # rather than leave a Temu-required cell empty.
    if any(token in header.casefold() for token in ("quantity", "number", "count")):
        return 1
    if any(token in header.casefold() for token in ("length", "width", "height", "capacity", "thickness")):
        return 1
    return "Not Applicable"


def ambiguous_set_measurement_reason(product: Product) -> str:
    """Identify when dimensions describe one component rather than the sold set."""
    text = normalize_space(" ".join([product.title, product.description, *product.attributes.values()])).casefold()
    if not product.dimensions_cm:
        return ""
    count_match = re.search(
        r"(?:комплект(?:ът)?\s+(?:се\s+)?съдържа|съдържа)\s*[:\-]?\s*(\d+)\s*(?:броя|бр\.?)",
        text,
        re.I,
    )
    if not count_match:
        return ""
    count = int(count_match.group(1))
    if count <= 1:
        return ""
    per_item_language = any(
        token in text
        for token in (
            "на страна",
            "на всяка от страните",
            "всеки е с размер",
            "размер на всеки",
            "размерът на всеки",
        )
    )
    if per_item_language:
        return f"Published dimensions describe one item, while the offer contains {count} pieces"
    return ""


def package_measurements_for_upload(
    product: Product, config: Mapping[str, Any]
) -> tuple[float, tuple[float, float, float], list[str], str, str]:
    """Return package weight/dimensions plus traceable REVIEW notes.

    Temu requires measurements for the packed SKU, while Oreshak usually
    publishes measurements for the unpacked product. In ``estimate_and_review``
    mode the scraper adds a configurable packing allowance to source-backed
    measurements. Missing or ambiguous measurements use the category fallback.
    Every estimated/fallback value is explicitly reported as a REVIEW note.
    """
    default_weight, default_l, default_w, default_h = DEFAULT_PACKAGE.get(
        product.category_id, (500, 30, 20, 10)
    )
    mode = str(config.get("package_measurement_mode", "strict")).strip().casefold()
    notes: list[str] = []

    if mode != "estimate_and_review":
        weight = float(product.weight_g if product.weight_g is not None else default_weight)
        dims = tuple(float(v) for v in (product.dimensions_cm or (default_l, default_w, default_h)))
        weight_basis = "published product weight" if product.weight_g is not None else "category fallback"
        dims_basis = "published product dimensions" if product.dimensions_cm is not None else "category fallback"
        return weight, dims, notes, weight_basis, dims_basis

    padding_cm = max(0.0, float(config.get("package_dimension_padding_cm", 2.0)))
    weight_percent = max(0.0, float(config.get("package_weight_padding_percent", 10.0)))
    weight_min_g = max(0.0, float(config.get("package_weight_padding_min_g", 50.0)))

    if product.weight_g is not None:
        added_weight = max(float(product.weight_g) * weight_percent / 100.0, weight_min_g)
        weight = float(product.weight_g) + added_weight
        weight_basis = (
            f"estimated package weight: published {float(product.weight_g):g} g + "
            f"{added_weight:g} g packing allowance"
        )
        notes.append(weight_basis)
    else:
        weight = float(default_weight)
        weight_basis = f"estimated package weight: category fallback {weight:g} g (source weight missing)"
        notes.append(weight_basis)

    ambiguous_reason = ambiguous_set_measurement_reason(product)
    if product.dimensions_cm is not None and not ambiguous_reason:
        source_dims = tuple(float(v) for v in product.dimensions_cm)
        dims = tuple(v + padding_cm for v in source_dims)
        dims_basis = (
            "estimated package dimensions: published "
            + " × ".join(f"{v:g}" for v in source_dims)
            + f" cm + {padding_cm:g} cm packing allowance per dimension"
        )
        notes.append(dims_basis)
    else:
        dims = (float(default_l), float(default_w), float(default_h))
        reason = ambiguous_reason or "complete source 3D dimensions missing"
        dims_basis = (
            "estimated package dimensions: category fallback "
            + " × ".join(f"{v:g}" for v in dims)
            + f" cm ({reason})"
        )
        notes.append(dims_basis)

    return round(weight, 1), tuple(round(v, 1) for v in dims), dedupe(notes), weight_basis, dims_basis


def source_measurement_errors(product: Product, config: Mapping[str, Any]) -> list[str]:
    """Validate source-backed measurements before an upload row is accepted."""
    mode = str(config.get("package_measurement_mode", "")).strip().casefold()
    if mode == "estimate_and_review":
        return []

    errors: list[str] = []
    strict_fallback = mode == "strict" or bool(config.get("omit_rows_with_fallback_measurements", True))
    strict_ambiguous = mode == "strict" or bool(config.get("omit_rows_with_ambiguous_set_dimensions", True))
    if strict_fallback:
        if product.weight_g is None:
            errors.append("No source-backed weight; category fallback measurements are disabled")
        if product.dimensions_cm is None:
            errors.append("No complete source-backed 3D dimensions; category fallback measurements are disabled")
    if strict_ambiguous:
        reason = ambiguous_set_measurement_reason(product)
        if reason:
            errors.append(reason)
    return errors


def prices_for_upload(product: Product, config: Mapping[str, Any]) -> tuple[Decimal | None, Decimal | None, str]:
    """Return the Temu selling price, Temu list price and a traceable basis.

    Oreshak may publish both a discounted checkout price and a higher regular
    price. When ``use_pre_promotion_price`` is enabled, the higher genuine
    pre-promotion price becomes the Temu selling price. In that mode no
    separate Temu list price is sent, because the uploaded selling price is
    already the regular price before the promotion.
    """
    multiplier = safe_decimal(config.get("price_multiplier", 1)) or Decimal("1")
    use_pre_promotion = bool(config.get("use_pre_promotion_price", True))

    source_value = product.price_eur
    basis = product.price_source or "current product price"
    if (
        use_pre_promotion
        and product.list_price_eur is not None
        and (product.price_eur is None or product.list_price_eur > product.price_eur)
    ):
        source_value = product.list_price_eur
        basis = product.list_price_source or "published pre-promotion price"

    price_eur = (
        (source_value * multiplier).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if source_value is not None else None
    )
    list_price_eur = None
    if not use_pre_promotion and product.list_price_eur is not None:
        list_price_eur = (product.list_price_eur * multiplier).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    return price_eur, list_price_eur, basis


def build_row(variant: Variant, schema: TemplateSchema, config: Mapping[str, Any]) -> tuple[dict[str, Any], list[str], list[str]]:
    product = variant.product
    price_eur, list_price_eur, _ = prices_for_upload(product, config)

    row: dict[str, Any] = {
        "E": product.category_id,
        "L": variant.title[:500], "M": product.code[:80], "N": variant.sku[:80],
        "T": product.description[:2000],
        "QE": (product.stock_quantity if product.stock_quantity is not None else int(config.get("default_in_stock_quantity", 10))) if product.in_stock else 0,
        "QF": price_eur, "QG": product.url, "QH": list_price_eur,
        "QI": "N/A" if list_price_eur is None else "",
        "QY": normalize_space(config.get("shipping_template")) or choose_valid(schema.dropdown_for("QY", product.category_id, {}), []),
        "QZ": normalize_space(config.get("handling_time")) or "1 Day",
        "RA": normalize_space(config.get("fulfillment_channel")) or choose_valid(schema.dropdown_for("RA", product.category_id, {}), ["Seller Fulfilled", "Seller"]),
        "RB": normalize_space(config.get("item_tax_code")) or "GEN STANDARD",
        "RC": choose_valid(schema.dropdown_for("RC", product.category_id, {}), [str(config.get("country_of_origin", "Bulgaria")), "Bulgaria"]),
        "TD": f"SKU: {variant.sku}"[:200],
        "TE": normalize_space(config.get("manufacturer_name")),
        "TF": normalize_space(config.get("eu_responsible_person")),
    }
    row.update(variation_fields(variant, schema))
    row.update(fabric_composition_fields(product, schema))
    for i, bullet in enumerate(product.bullet_points[:6]):
        row[col_letter(col_number("U") + i)] = bullet[:700]
    detail_limit = min(int(config.get("max_detail_images", 10)), 50)
    for i, image in enumerate(product.images[:detail_limit]):
        row[col_letter(col_number("AA") + i)] = image
    sku_limit = min(int(config.get("max_sku_images", 10)), 10)
    for i, image in enumerate(product.images[:sku_limit]):
        row[col_letter(col_number("PT") + i)] = image
    row["PT"] = variant.image or (product.images[0] if product.images else "")

    package_weight, package_dims, package_review_notes, _, _ = package_measurements_for_upload(product, config)
    row.update({"QJ": round(package_weight, 1), "QK": round(max(package_dims), 1),
                "QL": round(sorted(package_dims, reverse=True)[1], 1), "QM": round(min(package_dims), 1),
                "QO": choose_valid(schema.dropdown_for("QO", product.category_id, row), ["Yes"]),
                "QP": 1, "QQ": choose_valid(schema.dropdown_for("QQ", product.category_id, row), ["piece"])})

    required = schema.required_columns(product.category_id)
    # Data Definitions additionally marks these offer fields required even when
    # the category-mode helper sheet does not list them.
    required.update({"E", "L", "M", "N", "PF", "PT", "QE", "QF", "QJ", "QK", "QL", "QM", "QY", "QZ", "RA", "RC", "TD", "TE", "TF"})
    review_notes: list[str] = list(package_review_notes)
    for column in sorted(required, key=col_number):
        if column == "QH" and canonical(row.get("QI")) == canonical("N/A"):
            continue
        if row.get(column) in (None, ""):
            inferred = infer_required_value(column, product, variant, row, schema, config)
            row[column] = inferred
            if inferred not in (None, ""):
                header = schema.headers.get(column, column)
                review_notes.append(f"Auto-filled required {column} ({header}) with: {inferred}")

    # Known conditional requirements are filled only when their parent choice triggers them.
    conditionals = schema.conditional_columns(product.category_id)
    power = canonical(row.get("DV"))
    material_text = canonical(" ".join(str(row.get(c, "")) for c in ("DS", "DT", "DU", "FO", "FP", "FQ")))
    for column in sorted(conditionals, key=col_number):
        header = schema.headers.get(column, "")
        should_fill = False
        if column == "RD":
            should_fill = canonical(row.get("RC")) == canonical("Mainland China")
        elif "Plug Type" in header or "Operating Voltage" in header or "Acceptable Voltage" in header:
            should_fill = power not in ("", canonical("Use Without Electricity"))
        elif "Food Contact Material" in header:
            should_fill = canonical(row.get("FH")) == canonical("Yes")
        elif "Stainless Steel Grade" in header:
            should_fill = "stainlesssteel" in material_text
        elif "Wood Type" in header or "Wood Species" in header:
            should_fill = "wood" in material_text or "log" in material_text
        elif "Genuine Leather Type" in header:
            should_fill = "leather" in material_text or "genuine" in material_text
        elif "Battery Capacity" in header:
            should_fill = "battery" in power and "without" not in power
        if should_fill and row.get(column) in (None, ""):
            inferred = infer_required_value(column, product, variant, row, schema, config)
            row[column] = inferred
            if inferred not in (None, ""):
                header = schema.headers.get(column, column)
                review_notes.append(f"Auto-filled conditional {column} ({header}) with: {inferred}")

    errors: list[str] = []
    for column in sorted(required, key=col_number):
        if column == "QH" and canonical(row.get("QI")) == canonical("N/A"):
            continue
        if row.get(column) in (None, ""):
            errors.append(f"Missing required {column} ({schema.headers.get(column, '')})")
    if not normalize_space(product.description):
        errors.append("No product description")
    if not product.images:
        errors.append("No gallery images")
    if product.price_eur is None:
        errors.append("No EUR price")
    if product.category_id not in schema.category_names:
        errors.append(f"Unknown Temu category {product.category_id}")
    if product.mapping_confidence == "low":
        errors.append(f"No safe Temu category mapping: {product.mapping_reason}; add category_overrides.csv entry")
    errors.extend(source_measurement_errors(product, config))
    return row, dedupe(errors), dedupe(review_notes)


class XlsxTemplateWriter:
    def __init__(self, schema: TemplateSchema):
        self.schema = schema

    @staticmethod
    def _set_cell(row_node: etree._Element, address: str, value: Any) -> None:
        existing = row_node.find(f"{{{NS_MAIN}}}c[@r='{address}']")
        if existing is not None:
            # Never overwrite formulas. The only formula in data rows is Category Name (F).
            if existing.find(f"{{{NS_MAIN}}}f") is not None:
                return
            row_node.remove(existing)
        cell = etree.Element(f"{{{NS_MAIN}}}c", r=address)
        if isinstance(value, Decimal):
            etree.SubElement(cell, f"{{{NS_MAIN}}}v").text = format(value, "f")
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            number = 0 if isinstance(value, float) and (math.isnan(value) or math.isinf(value)) else value
            etree.SubElement(cell, f"{{{NS_MAIN}}}v").text = str(number)
        else:
            cell.set("t", "inlineStr")
            inline = etree.SubElement(cell, f"{{{NS_MAIN}}}is")
            text = etree.SubElement(inline, f"{{{NS_MAIN}}}t")
            string = str(value)
            if string[:1].isspace() or string[-1:].isspace():
                text.set(f"{{{NS_XML}}}space", "preserve")
            text.text = string
        row_node.append(cell)

    def write(self, rows: Sequence[Mapping[str, Any]], output_path: Path) -> None:
        template_sheet_path = self.schema.sheet_paths["Template"]
        sheet_root = etree.fromstring(self.schema._zip.read(template_sheet_path))
        sheet_data = sheet_root.find(f"{{{NS_MAIN}}}sheetData")
        if sheet_data is None:
            raise RuntimeError("Template sheetData not found")
        row_map = {int(row.get("r")): row for row in sheet_data.findall(f"{{{NS_MAIN}}}row")}
        # Clear previous imported values while preserving formula cells.
        for row_number, row_node in row_map.items():
            if row_number < 5:
                continue
            for cell in list(row_node.findall(f"{{{NS_MAIN}}}c")):
                if cell.find(f"{{{NS_MAIN}}}f") is None:
                    row_node.remove(cell)
        for index, data in enumerate(rows, start=5):
            row_node = row_map.get(index)
            if row_node is None:
                row_node = etree.Element(f"{{{NS_MAIN}}}row", r=str(index))
                sheet_data.append(row_node)
                row_map[index] = row_node
            for column, value in data.items():
                if value not in (None, ""):
                    self._set_cell(row_node, f"{column}{index}", value)
            cells = list(row_node.findall(f"{{{NS_MAIN}}}c"))
            cells.sort(key=lambda cell: col_number(cell_col(cell.get("r"))))
            for cell in cells:
                row_node.remove(cell)
            row_node.extend(cells)

        workbook_root = etree.fromstring(self.schema._zip.read("xl/workbook.xml"))
        calc_pr = workbook_root.find(f"{{{NS_MAIN}}}calcPr")
        if calc_pr is None:
            calc_pr = etree.SubElement(workbook_root, f"{{{NS_MAIN}}}calcPr")
        calc_pr.set("fullCalcOnLoad", "1")
        calc_pr.set("forceFullCalc", "1")
        calc_pr.set("calcMode", "auto")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output_path, "w", ZIP_DEFLATED) as out_zip:
            for info in self.schema._zip.infolist():
                if info.filename == template_sheet_path:
                    data = etree.tostring(sheet_root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif info.filename == "xl/workbook.xml":
                    data = etree.tostring(workbook_root, xml_declaration=True, encoding="UTF-8", standalone=True)
                else:
                    data = self.schema._zip.read(info.filename)
                out_zip.writestr(info, data)


def write_csv(path: Path, fieldnames: Sequence[str], rows: Iterable[Mapping[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def raw_product_row(product: Product, config: Mapping[str, Any]) -> dict[str, Any]:
    upload_price_eur, upload_list_price_eur, upload_price_basis = prices_for_upload(product, config)
    package_weight, package_dims, package_notes, package_weight_basis, package_dimensions_basis = package_measurements_for_upload(product, config)
    return {
        "product_code": product.code, "product_name": product.title, "source_category": product.source_category_name,
        "temu_category_id": product.category_id, "temu_mapping_reason": product.mapping_reason,
        "mapping_confidence": product.mapping_confidence, "description": product.description,
        "price_eur": product.price_eur, "list_price_eur": product.list_price_eur,
        "price_source": product.price_source, "list_price_source": product.list_price_source,
        "upload_price_eur": upload_price_eur, "upload_list_price_eur": upload_list_price_eur,
        "upload_price_basis": upload_price_basis,
        "availability": "in_stock" if product.in_stock else "out_of_stock", "stock_source": product.stock_source, "stock_quantity": product.stock_quantity,
        "weight_g": product.weight_g, "weight_source": product.weight_source,
        "length_cm": product.dimensions_cm[0] if product.dimensions_cm else "",
        "width_cm": product.dimensions_cm[1] if product.dimensions_cm else "",
        "height_cm": product.dimensions_cm[2] if product.dimensions_cm else "",
        "dimensions_source": product.dimensions_source,
        "package_weight_g": package_weight,
        "package_length_cm": max(package_dims),
        "package_width_cm": sorted(package_dims, reverse=True)[1],
        "package_height_cm": min(package_dims),
        "package_weight_basis": package_weight_basis,
        "package_dimensions_basis": package_dimensions_basis,
        "package_review_notes": " | ".join(package_notes),
        "product_url": product.url, "images": " | ".join(product.images),
        "attributes_json": json.dumps(product.attributes, ensure_ascii=False),
        "options_json": json.dumps([option.__dict__ for option in product.options], ensure_ascii=False),
        "warnings": " | ".join(product.warnings),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape Oreshak.bg and populate a Temu XLSX template")
    parser.add_argument("--template", default="template.xlsx", help="Temu XLSX template")
    parser.add_argument("--config", default="config.json", help="JSON configuration")
    parser.add_argument("--output-dir", default="output", help="Output directory")
    parser.add_argument("--limit", type=int, default=10, help="Maximum source products overall; 0 means all")
    parser.add_argument("--limit-per-category", type=int, default=0, help="Maximum products from each category; 0 means all")
    parser.add_argument("--category", action="append", help="Run only URL/slug containing this value")
    parser.add_argument("--schema-only", action="store_true", help="Validate template/config without scraping")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    base_dir = Path(__file__).resolve().parent
    template_path = (base_dir / args.template).resolve() if not Path(args.template).is_absolute() else Path(args.template)
    config_path = (base_dir / args.config).resolve() if not Path(args.config).is_absolute() else Path(args.config)
    output_dir = (base_dir / args.output_dir).resolve() if not Path(args.output_dir).is_absolute() else Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    log_path = output_dir / "scraper.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
    )
    config = json.loads(config_path.read_text(encoding="utf-8"))
    schema = TemplateSchema(template_path)
    try:
        logging.info("Template loaded: %s categories, %s columns", len(schema.category_names), len(schema.headers))
        if args.schema_only:
            logging.info("Schema validation completed successfully")
            return 0
        category_urls = list(config["category_urls"])
        if args.category:
            needles = [needle.casefold() for needle in args.category]
            category_urls = [url for url in category_urls if any(needle in url.casefold() for needle in needles)]
        override_path = base_dir / str(config.get("category_overrides_file", "category_overrides.csv"))
        overrides = load_overrides(override_path)
        client = OreshakClient(config)
        discovered: list[tuple[str, str]] = []
        seen_links: set[str] = set()
        for category_url in category_urls:
            try:
                links = client.discover_product_links(category_url)
            except Exception as exc:
                logging.exception("Category failed: %s: %s", category_url, exc)
                continue
            if args.limit_per_category:
                links = links[:args.limit_per_category]
            for link in links:
                if link not in seen_links:
                    seen_links.add(link)
                    discovered.append((link, category_url))
                    if args.limit and len(discovered) >= args.limit:
                        break
            if args.limit and len(discovered) >= args.limit:
                break
        logging.info("Discovered %s unique products", len(discovered))

        products: list[Product] = []
        failures: list[dict[str, Any]] = []
        for index, (url, category_url) in enumerate(discovered, start=1):
            logging.info("Product %s/%s: %s", index, len(discovered), url)
            try:
                product = client.parse_product(url, category_url)
                ensure_product_description(product)
                product.category_id, product.mapping_reason, product.mapping_confidence = category_for(product, overrides, schema)
                if product.mapping_confidence != "high":
                    product.warnings.append(f"Category mapping {product.mapping_confidence}: {product.mapping_reason}")
                if product.category_id not in schema.category_names:
                    raise ValueError(f"Mapped category {product.category_id} is not in the template")
                products.append(product)
            except Exception as exc:
                logging.exception("Product failed: %s", url)
                failures.append({"product_url": url, "source_category_url": category_url, "error": str(exc)})

        temu_rows: list[dict[str, Any]] = []
        validation_rows: list[dict[str, Any]] = []
        mapping_rows: list[dict[str, Any]] = []
        for product in products:
            mapping_rows.append({
                "product_code": product.code, "product_name": product.title, "product_url": product.url,
                "source_category": product.source_category_name, "temu_category_id": product.category_id,
                "temu_category_name": schema.category_names.get(product.category_id, ""),
                "mapping_reason": product.mapping_reason, "mapping_confidence": product.mapping_confidence,
            })
            if not product.in_stock and not bool(config.get("include_out_of_stock", True)):
                validation_rows.append({
                    "product_code": product.code, "sku": product.code, "product_url": product.url,
                    "temu_category_id": product.category_id, "status": "SKIPPED",
                    "issues": "Out of stock; retained in raw export but excluded from Temu upload",
                })
                continue
            for variant in expand_variants(product):
                if len(temu_rows) >= MAX_TEMU_ROWS:
                    logging.warning("Temu limit of %s rows reached; remaining variants omitted", MAX_TEMU_ROWS)
                    break
                row, errors, row_review_notes = build_row(variant, schema, config)
                all_review_notes = [*product.warnings, *row_review_notes]
                status = "ERROR" if errors else "REVIEW" if all_review_notes else "OK"
                validation_rows.append({
                    "product_code": product.code, "sku": variant.sku, "product_url": product.url,
                    "temu_category_id": product.category_id, "status": status,
                    "issues": " | ".join([*errors, *all_review_notes]),
                })
                if errors:
                    logging.error("Row omitted for %s: %s", variant.sku, "; ".join(errors))
                    continue
                temu_rows.append(row)

        writer = XlsxTemplateWriter(schema)
        xlsx_path = output_dir / "TEMU_ORESHAK_UPLOAD.xlsx"
        writer.write(temu_rows, xlsx_path)
        raw_fields = [
            "product_code", "product_name", "source_category", "temu_category_id", "temu_mapping_reason",
            "mapping_confidence", "description", "price_eur", "list_price_eur", "price_source", "list_price_source",
            "upload_price_eur", "upload_list_price_eur", "upload_price_basis",
            "availability", "stock_source", "stock_quantity", "weight_g", "weight_source", "length_cm", "width_cm", "height_cm",
            "dimensions_source", "package_weight_g", "package_length_cm", "package_width_cm", "package_height_cm",
            "package_weight_basis", "package_dimensions_basis", "package_review_notes",
            "product_url", "images", "attributes_json", "options_json", "warnings",
        ]
        write_csv(output_dir / "oreshak_raw_export.csv", raw_fields, (raw_product_row(p, config) for p in products))
        write_csv(output_dir / "category_mapping_review.csv", [
            "product_code", "product_name", "product_url", "source_category", "temu_category_id",
            "temu_category_name", "mapping_reason", "mapping_confidence",
        ], mapping_rows)
        write_csv(output_dir / "validation_report.csv", [
            "product_code", "sku", "product_url", "temu_category_id", "status", "issues",
        ], validation_rows)
        write_csv(output_dir / "failed_products.csv", ["product_url", "source_category_url", "error"], failures)
        summary = {
            "categories_requested": len(category_urls), "products_discovered": len(discovered),
            "products_parsed": len(products), "temu_rows_written": len(temu_rows),
            "failed_products": len(failures),
            "rows_with_errors": sum(1 for row in validation_rows if row["status"] == "ERROR"),
            "rows_for_review": sum(1 for row in validation_rows if row["status"] == "REVIEW"),
            "out_of_stock_skipped": sum(1 for row in validation_rows if row["status"] == "SKIPPED"),
            "output_file": str(xlsx_path),
        }
        (output_dir / "run_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        logging.info("Completed: %s", json.dumps(summary, ensure_ascii=False))
        return 0 if temu_rows else 2
    finally:
        schema.close()


if __name__ == "__main__":
    raise SystemExit(main())
