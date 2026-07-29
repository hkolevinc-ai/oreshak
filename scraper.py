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
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse
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
    "kutii-za-shah-i-tabla": "12179",
    "kuhnenski-aksesoari-ot-darvo-oreshak": "9923",
    "nojove-ot-balgaria": "10059",
    "profesionalen-shahmat": "51777",
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
    (re.compile(r"(кутия).*(бижут)|бижут.*кутия"), "12735", "jewelry box"),
    (re.compile(r"(кутия|кашон).*(подар|gift)"), "39880", "gift box"),
    (re.compile(r"(стойка|рафт).*(вино|бутил)"), "12686", "wine rack"),
    (re.compile(r"(кутия).*(вино|бутил)"), "39880", "wine gift box"),
    (re.compile(r"\bбуре|бъчв"), "10905", "barrel"),
    (re.compile(r"бъклиц|манерк|кег"), "10888", "flask/keg"),
    (re.compile(r"(трофе|глиги|сръндак|глиган|елен).*(дъск|поставка)|дъск.*трофе"), "32650", "trophy mount"),
    (re.compile(r"(шахматен часовник|фигури за шах|пулове|зарове|аксесоар.*шах|аксесоар.*табла)"), "25613", "game pieces/accessories"),
    (re.compile(r"(комплект|сет).*(шах|табла)|\bтабла\b|шах.*табла"), "25615", "board game"),
    (re.compile(r"професионален.*шах|шахмат.*професион"), "51777", "professional chess"),
    (re.compile(r"(черпак|черпало)"), "9998", "ladle"),
    (re.compile(r"(лъжиц|шпатула)"), "9999", "cooking spoon"),
    (re.compile(r"(комплект|сет).*(прибор|кухненск|лъжиц|шпатула)"), "10006", "utensil set"),
    (re.compile(r"(дъска).*(рязане|сервира|мезе|сирена)|serving board"), "54423", "serving board"),
    (re.compile(r"\bподнос|tray"), "10741", "serving tray"),
    (re.compile(r"\bплато|platter"), "10740", "platter"),
    (re.compile(r"(десертн).*(чини)|чини.*десерт"), "10807", "dessert plate"),
    (re.compile(r"\bчини(я|и)|plate"), "10808", "dinner plate"),
    (re.compile(r"(поставка|органайзер).*(нож|прибор)"), "10328", "utensil rack"),
    (re.compile(r"(нож).*(хранене|масов)"), "10638", "dinner knife"),
    (re.compile(r"(нож).*(плод|универсал|ловен)|\bловен нож"), "10072", "utility knife"),
    (re.compile(r"(кухненски нож|готварски нож|нож.*готвач|сатър|chef)"), "10059", "chef knife"),
    (re.compile(r"(кутия).*(нож|аксесоар|пура|тютюн)"), "12179", "decorative box"),
    (re.compile(r"(отварачка).*(вино|тирбушон)"), "10875", "wine accessory"),
    (re.compile(r"(отварачка|gadget)"), "9923", "kitchen gadget"),
    (re.compile(r"(табел|надпис|плакет|plaque)"), "12193", "decorative plaque"),
    (re.compile(r"(пано|пластик|релеф|дърворезб)"), "12151", "wall sculpture"),
    (re.compile(r"(картина|живопис|painting)"), "12867", "painting"),
    (re.compile(r"(гравюра|дърворит|woodcut)"), "39291", "woodcut"),
    (re.compile(r"(смесена техника|mixed media)"), "39280", "mixed media"),
    (re.compile(r"(бяла заготовка|неоцветен|за оцветяване|blank)"), "39981", "wood art blank"),
    (re.compile(r"(колед|елха|орнамент|украса).*(фигур|статует)|figurine ornament"), "13031", "figurine ornament"),
    (re.compile(r"(висящ|за окачване|hanging ornament)"), "12141", "hanging ornament"),
    (re.compile(r"(фигур|статует|скулптур|figurine)"), "12140", "collectible figurine"),
    (re.compile(r"(тъкан|плат|fabric|текстил)"), "39650", "fabric"),
]

MATERIAL_TRANSLATIONS: list[tuple[re.Pattern[str], list[str]]] = [
    (re.compile(r"дамаск"), ["Damascus Steel", "Stainless Steel", "Steel"]),
    (re.compile(r"неръждаем"), ["Stainless Steel", "Steel"]),
    (re.compile(r"стоман"), ["High carbon steel", "Carbon Steel", "Steel", "Stainless Steel"]),
    (re.compile(r"месинг"), ["Brass", "Copper Alloy", "Metal"]),
    (re.compile(r"медн|\bмед\b"), ["Copper", "Copper Alloy", "Metal"]),
    (re.compile(r"алумини"), ["Aluminum Alloy", "Aluminum", "Metal"]),
    (re.compile(r"метал"), ["Metal", "Iron", "Steel", "Zinc Alloy"]),
    (re.compile(r"кож|еленов рог"), ["Leather", "Genuine", "Buckskin", "Animal Skin, Fur, And Down"]),
    (re.compile(r"керами"), ["Ceramic", "Ceramics", "Clay"]),
    (re.compile(r"стък"), ["Glass"]),
    (re.compile(r"плат|тъкан|текстил"), ["Textile", "Cotton", "Linen", "Fabric"]),
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
    options: list[ProductOption] = field(default_factory=list)
    category_id: str = ""
    mapping_reason: str = ""
    mapping_confidence: str = "medium"
    weight_g: float | None = None
    dimensions_cm: tuple[float, float, float] | None = None
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
        json_ld = self._json_ld_product(soup)
        product.title = normalize_space(
            self._first_text(soup, ["#content h1", "h1", ".product-info h1"]) or json_ld.get("name")
        )
        if not product.title:
            raise ValueError("Product title not found")
        product.code = self._parse_code(soup, html, url)
        product.source_category_name = self._parse_source_category(soup)
        product.description = self._parse_description(soup, json_ld)
        product.attributes = self._parse_attributes(soup, product.description)
        product.bullet_points = self._make_bullets(product)
        product.tags = self._parse_tags(soup)
        product.images = self._parse_images(soup, json_ld, url)
        product.price_eur, product.list_price_eur = self._parse_prices(soup, json_ld)
        product.in_stock = self._parse_stock(soup)
        product.options = self._parse_options(soup)
        product.weight_g = parse_weight(product.description + " " + " ".join(product.attributes.values()))
        product.dimensions_cm = parse_dimensions(product.description + " " + " ".join(product.attributes.values()))
        if not product.code:
            product.code = self._generated_code(url, product.title)
            product.warnings.append("Product code was generated from URL/title")
        if not product.images:
            product.warnings.append("No product gallery image found")
        if product.price_eur is None:
            product.warnings.append("EUR price not found")
        return product

    @staticmethod
    def _json_ld_product(soup: BeautifulSoup) -> dict[str, Any]:
        for script in soup.select('script[type="application/ld+json"]'):
            try:
                data = json.loads(script.string or script.get_text())
            except Exception:
                continue
            candidates = data if isinstance(data, list) else [data]
            for item in candidates:
                if isinstance(item, dict) and item.get("@type") == "Product":
                    return item
                if isinstance(item, dict) and isinstance(item.get("@graph"), list):
                    for graph_item in item["@graph"]:
                        if isinstance(graph_item, dict) and graph_item.get("@type") == "Product":
                            return graph_item
        return {}

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
            # Capture only the main product content, stopping before reviews/footer.
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
        return description[:2000]

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
        for heading in ("Характеристики", "Материал", "Тегло", "Размери", "Дължина", "Диаметър"):
            match = re.search(rf"{heading}\s*:\s*([^.;]+)", description, re.I)
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
    def _parse_images(soup: BeautifulSoup, json_ld: Mapping[str, Any], base_url: str) -> list[str]:
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
        # Fallback restricted to image area; intentionally excludes #tab-description.
        for tag in soup.select("#content .product-info img, #content .col-sm-4 img"):
            if tag.find_parent(id="tab-description"):
                continue
            src = tag.get("data-zoom-image") or tag.get("data-src") or tag.get("src")
            if src:
                urls.append(urljoin(base_url, src))
        cleaned: list[str] = []
        for url in urls:
            lower = url.lower()
            if not lower.startswith(("http://", "https://")):
                continue
            if any(bad in lower for bad in ("logo", "no_image", "placeholder", "facebook", "loader", "icon")):
                continue
            # OpenCart resized images are valid URLs; prefer original when the cache path is obvious.
            cleaned.append(url.replace("/cache/", "/"))
        return dedupe(cleaned)

    @staticmethod
    def _parse_prices(soup: BeautifulSoup, json_ld: Mapping[str, Any]) -> tuple[Decimal | None, Decimal | None]:
        current: list[Decimal] = []
        old: list[Decimal] = []
        for selector in (".price-new", ".product-price .special", "#content h2", ".price"):
            for tag in soup.select(selector):
                text = normalize_space(tag.get_text(" ", strip=True))
                match = re.search(r"([0-9][0-9\s.,]*)\s*€", text)
                if match:
                    value = safe_decimal(match.group(1))
                    if value is not None:
                        current.append(value)
        for selector in (".price-old", "del", "s"):
            for tag in soup.select(selector):
                match = re.search(r"([0-9][0-9\s.,]*)\s*€", normalize_space(tag.get_text(" ", strip=True)))
                if match:
                    value = safe_decimal(match.group(1))
                    if value is not None:
                        old.append(value)
        offers = json_ld.get("offers")
        if isinstance(offers, dict):
            value = safe_decimal(offers.get("price"))
            if value is not None:
                current.append(value)
        # Main product content often contains one current EUR price followed by BGN.
        if not current:
            content = normalize_space((soup.select_one("#content") or soup).get_text(" ", strip=True))
            match = re.search(r"([0-9][0-9\s.,]*)\s*€", content)
            if match:
                value = safe_decimal(match.group(1))
                if value is not None:
                    current.append(value)
        price = min(current) if current else None
        list_price = max(old) if old else price
        return price, list_price

    @staticmethod
    def _parse_stock(soup: BeautifulSoup) -> bool:
        text = normalize_space((soup.select_one("#content") or soup).get_text(" ", strip=True)).casefold()
        unavailable = ("неналичен", "изчерпан", "out of stock", "не е наличен")
        if any(marker in text for marker in unavailable):
            buy_button = soup.select_one("#button-cart:not([disabled]), button[id*='cart']:not([disabled])")
            return buy_button is not None and "неналичен" not in normalize_space(buy_button.get_text()).casefold()
        return True

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


def parse_weight(text: str) -> float | None:
    patterns = [
        (r"(?:тегло|weight)\s*:?\s*([0-9]+(?:[.,][0-9]+)?)\s*(кг|kg)\b", 1000),
        (r"(?:тегло|weight)\s*:?\s*([0-9]+(?:[.,][0-9]+)?)\s*(гр|г|g)\b", 1),
    ]
    for pattern, multiplier in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            value = float(match.group(1).replace(",", ".")) * multiplier
            return round(max(value, 0.1), 1)
    return None


def parse_dimensions(text: str) -> tuple[float, float, float] | None:
    normalized = text.replace(",", ".").replace("×", "x").replace("Х", "x").replace("х", "x")
    patterns = [
        r"(?:размери|размер|dimensions?)\s*:?\s*(\d+(?:\.\d+)?)\s*[x/]\s*(\d+(?:\.\d+)?)(?:\s*[x/]\s*(\d+(?:\.\d+)?))?\s*см",
        r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)(?:\s*x\s*(\d+(?:\.\d+)?))?\s*см",
    ]
    for pattern in patterns:
        match = re.search(pattern, normalized, re.I)
        if match:
            dims = [float(v) for v in match.groups() if v]
            if len(dims) == 2:
                dims.append(3.0)
            return round(max(dims), 1), round(sorted(dims, reverse=True)[1], 1), round(min(dims), 1)
    diameter = re.search(r"(?:диаметър|ф|Ø)\s*:?\s*(\d+(?:\.\d+)?)\s*см?", normalized, re.I)
    if diameter:
        d = float(diameter.group(1))
        return round(d, 1), round(d, 1), 5.0
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
    text = " ".join([product.title, product.description, *product.attributes.values()]).casefold()
    candidates: list[str] = []
    for pattern, values in MATERIAL_TRANSLATIONS:
        if pattern.search(text):
            candidates.extend(values)
    if not candidates:
        candidates = ["Wood", "Log"] if "дър" in text else ["Other", "Metal"]
    if food_contact and any(canonical(v) in {"wood", "solidwood", "naturalwood"} for v in candidates):
        candidates = ["Log", "Wood", *candidates]
    return dedupe(candidates)


def category_for(product: Product, overrides: Mapping[str, str], schema: TemplateSchema) -> tuple[str, str, str]:
    if product.code in overrides:
        return overrides[product.code], "override by product code", "high"
    if product.url in overrides:
        return overrides[product.url], "override by product URL", "high"
    haystack = normalize_space(" ".join([product.title, product.description[:500], product.source_category_name])).casefold()
    for pattern, category_id, reason in KEYWORD_CATEGORY_RULES:
        if pattern.search(haystack) and category_id in schema.category_names:
            return category_id, reason, "high"
    default = SOURCE_DEFAULTS.get(slug_key(product.source_category_url), "13020")
    return default, f"source category default: {slug_key(product.source_category_url)}", "medium"


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
    if column == "QE": return int(config.get("default_in_stock_quantity", 10)) if product.in_stock else 0
    if column == "QF": return product.price_eur
    if column == "QG": return product.url
    if column == "QH": return product.list_price_eur or product.price_eur
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

    if column in {"DS", "DT", "DU", "FO", "FP", "FQ"} or "Material" in header:
        return choose_valid(dropdown, detect_material_candidates(product, food_contact="Food Contact" in header))
    if "Power Supply" in header:
        return choose_valid(dropdown, ["Use Without Electricity", "Without Electricity", "Non Electric"])
    if "Battery Properties" in header:
        return choose_valid(dropdown, ["Without Battery", "No Battery", "Battery Free"])
    if "Applicable Age Group" in header:
        age = "18 Years+" if any(word in text for word in ("ловен", "нож", "трофей")) else "14 Years+"
        return choose_valid(dropdown, [age, "12 Years+", "8 Years+"])
    if "Can Be Used For Food Contact" in header:
        food = category_id in {"9998", "9999", "10006", "10059", "10072", "10638", "10628", "54423", "10740", "10741", "10807", "10808", "10853", "11514", "10905", "10888", "10875", "9923"}
        return choose_valid(dropdown, ["Yes" if food else "No"])
    if "Food Contact Material" in header:
        return choose_valid(dropdown, detect_material_candidates(product, food_contact=True))
    if "Closure Type" in header:
        return choose_valid(dropdown, ["Magnetic" if "магнит" in text else "Latch", "Flip top"])
    if "Water Resistance Level" in header:
        return choose_valid(dropdown, ["Non-water resistant"])
    if "Thickness" in header and "value" in schema.internal_keys.get(column, "").casefold():
        return round(min(dims), 1)
    if "Thickness" in header and "unit" in schema.internal_keys.get(column, "").casefold():
        return choose_valid(dropdown, ["cm"])
    if "Wood Type" in header:
        return choose_valid(dropdown, ["Solid Wood", "Natural Wood", "Wood"])
    if "Wood Species" in header:
        species_candidates: list[str] = []
        for bg, candidates in WOOD_SPECIES.items():
            if bg in text:
                species_candidates.extend(candidates)
        return choose_valid(dropdown, species_candidates or ["Beech", "Walnut", "Oak"])
    if "Stainless Steel Grade" in header:
        return choose_valid(dropdown, ["304", "18/10", "Other"])
    if "Genuine Leather Type" in header:
        return choose_valid(dropdown, ["Cowhide", "Full Grain Leather", "Genuine Leather"])
    if "Packaging unit" in header:
        return choose_valid(dropdown, ["piece"])
    if dropdown:
        return choose_valid(dropdown, ["Other", "No", "None", "Not Applicable", "Use Without Electricity"])
    # Free-text/numeric fallback. It is preferable to provide a traceable SKU
    # rather than leave a Temu-required cell empty.
    if any(token in header.casefold() for token in ("quantity", "number", "count")):
        return 1
    if any(token in header.casefold() for token in ("length", "width", "height", "capacity", "thickness")):
        return 1
    return "Not Applicable"


def build_row(variant: Variant, schema: TemplateSchema, config: Mapping[str, Any]) -> tuple[dict[str, Any], list[str]]:
    product = variant.product
    multiplier = safe_decimal(config.get("price_multiplier", 1)) or Decimal("1")
    price_eur = (product.price_eur * multiplier).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if product.price_eur is not None else None
    list_price_eur = (product.list_price_eur * multiplier).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if product.list_price_eur is not None else None

    row: dict[str, Any] = {
        "E": product.category_id,
        "L": variant.title[:500], "M": product.code[:80], "N": variant.sku[:80],
        "T": product.description[:2000],
        "QE": int(config.get("default_in_stock_quantity", 10)) if product.in_stock else 0,
        "QF": price_eur, "QG": product.url, "QH": list_price_eur or price_eur,
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
    for i, bullet in enumerate(product.bullet_points[:6]):
        row[col_letter(col_number("U") + i)] = bullet[:700]
    detail_limit = min(int(config.get("max_detail_images", 10)), 50)
    for i, image in enumerate(product.images[:detail_limit]):
        row[col_letter(col_number("AA") + i)] = image
    sku_limit = min(int(config.get("max_sku_images", 10)), 10)
    for i, image in enumerate(product.images[:sku_limit]):
        row[col_letter(col_number("PT") + i)] = image
    row["PT"] = variant.image or (product.images[0] if product.images else "")

    default_weight, default_l, default_w, default_h = DEFAULT_PACKAGE.get(product.category_id, (500, 30, 20, 10))
    dims = product.dimensions_cm or (default_l, default_w, default_h)
    row.update({"QJ": round(product.weight_g or default_weight, 1), "QK": round(max(dims), 1),
                "QL": round(sorted(dims, reverse=True)[1], 1), "QM": round(min(dims), 1),
                "QO": choose_valid(schema.dropdown_for("QO", product.category_id, row), ["Yes"]),
                "QP": 1, "QQ": choose_valid(schema.dropdown_for("QQ", product.category_id, row), ["piece"])})

    required = schema.required_columns(product.category_id)
    # Data Definitions additionally marks these offer fields required even when
    # the category-mode helper sheet does not list them.
    required.update({"E", "L", "M", "N", "PF", "PT", "QE", "QF", "QH", "QJ", "QK", "QL", "QM", "QY", "QZ", "RA", "RC", "TD", "TE", "TF"})
    for column in sorted(required, key=col_number):
        if row.get(column) in (None, ""):
            row[column] = infer_required_value(column, product, variant, row, schema, config)

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
            row[column] = infer_required_value(column, product, variant, row, schema, config)

    errors: list[str] = []
    for column in sorted(required, key=col_number):
        if row.get(column) in (None, ""):
            errors.append(f"Missing required {column} ({schema.headers.get(column, '')})")
    if not product.images:
        errors.append("No gallery images")
    if product.price_eur is None:
        errors.append("No EUR price")
    if product.category_id not in schema.category_names:
        errors.append(f"Unknown Temu category {product.category_id}")
    return row, errors


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


def raw_product_row(product: Product) -> dict[str, Any]:
    return {
        "product_code": product.code, "product_name": product.title, "source_category": product.source_category_name,
        "temu_category_id": product.category_id, "temu_mapping_reason": product.mapping_reason,
        "mapping_confidence": product.mapping_confidence, "description": product.description,
        "price_eur": product.price_eur, "list_price_eur": product.list_price_eur,
        "availability": "in_stock" if product.in_stock else "out_of_stock",
        "weight_g": product.weight_g, "length_cm": product.dimensions_cm[0] if product.dimensions_cm else "",
        "width_cm": product.dimensions_cm[1] if product.dimensions_cm else "",
        "height_cm": product.dimensions_cm[2] if product.dimensions_cm else "",
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
    parser.add_argument("--limit", type=int, default=10, help="Maximum source products; 0 means all")
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
                product.category_id, product.mapping_reason, product.mapping_confidence = category_for(product, overrides, schema)
                if product.category_id not in schema.category_names:
                    raise ValueError(f"Mapped category {product.category_id} is not in the template")
                if not product.in_stock and not bool(config.get("include_out_of_stock", True)):
                    logging.info("Skipping out-of-stock product: %s", product.title)
                    continue
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
            for variant in expand_variants(product):
                if len(temu_rows) >= MAX_TEMU_ROWS:
                    logging.warning("Temu limit of %s rows reached; remaining variants omitted", MAX_TEMU_ROWS)
                    break
                row, errors = build_row(variant, schema, config)
                validation_rows.append({
                    "product_code": product.code, "sku": variant.sku, "product_url": product.url,
                    "temu_category_id": product.category_id, "status": "ERROR" if errors else "OK",
                    "issues": " | ".join(errors),
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
            "mapping_confidence", "description", "price_eur", "list_price_eur", "availability", "weight_g",
            "length_cm", "width_cm", "height_cm", "product_url", "images", "attributes_json", "options_json", "warnings",
        ]
        write_csv(output_dir / "oreshak_raw_export.csv", raw_fields, (raw_product_row(p) for p in products))
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
            "failed_products": len(failures), "rows_with_errors": sum(1 for row in validation_rows if row["status"] == "ERROR"),
            "output_file": str(xlsx_path),
        }
        (output_dir / "run_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        logging.info("Completed: %s", json.dumps(summary, ensure_ascii=False))
        return 0 if temu_rows else 2
    finally:
        schema.close()


if __name__ == "__main__":
    raise SystemExit(main())
