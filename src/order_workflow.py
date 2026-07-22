from __future__ import annotations

import hashlib
import io
import json
import math
import posixpath
import re
import sqlite3
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable
from functools import lru_cache
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageChops, ImageOps, UnidentifiedImageError
import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from src.navigation import NavigationItem, render_mobile_navigation, render_sidebar


ORDER_MODE_STONES = "Камни"
ORDER_MODE_PEARLS = "Жемчуг"
ORDER_MODES = (ORDER_MODE_STONES, ORDER_MODE_PEARLS)

CATEGORY_TOP = "Топы продаж"
CATEGORY_MEDIUM = "Средние комплекты"
CATEGORY_WEAK = "Слабые комплекты"
CATEGORY_ZERO = "Нулевые комплекты"
CATEGORY_ORDER = (CATEGORY_TOP, CATEGORY_MEDIUM, CATEGORY_WEAK, CATEGORY_ZERO)
CATEGORY_SHORT = {
    CATEGORY_TOP: "Топы",
    CATEGORY_MEDIUM: "Средние",
    CATEGORY_WEAK: "Слабые",
    CATEGORY_ZERO: "Нулевые",
}
CATEGORY_TONE = {
    CATEGORY_TOP: "🔴",
    CATEGORY_MEDIUM: "🟠",
    CATEGORY_WEAK: "🟡",
    CATEGORY_ZERO: "⚪",
}

RING_SIZES = tuple(range(15, 25))
DRAFT_VERSION = 2

ROOT_DIR = Path(__file__).resolve().parents[1]
RUNTIME_DIR = ROOT_DIR / "data" / "order_runtime"
UPLOAD_DIR = RUNTIME_DIR / "uploads"
DRAFT_DB = RUNTIME_DIR / "order_drafts.sqlite3"
ORDER_EXCLUSIONS_FILE = ROOT_DIR / "data" / "order_exclusions.json"

_XML_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_XML_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_XML_PACKAGE_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_XML_DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
_XML_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# Goods that are mounted internally and therefore must not be ordered as ready items.
STONE_EXCLUSION_PATTERNS: tuple[str, ...] = (
    "BSHQ",
    "BSMQ",
    "BLUE SAPPHIRE HIGH QUALITY",
    "BLUE SAPPHIRE MEDIUM QUALITY",
    "BLUE SAPPHIRE HQ",
    "BLUE SAPPHIRE MQ",
    "EMERALD HIGH QUALITY",
    "EMERALD HQ",
)

SEA_PEARL_PATTERNS: tuple[str, ...] = (
    "SEA PEARL",
    "SOUTH SEA",
    "AKOYA",
    "TAHITI",
    "TAHITIAN",
    "GALATEA",
    "FACETED SEA",
)

ORDER_SECTIONS = (
    ("order-overview", "Сводка"),
    ("order-workspace", "Комплекты"),
    ("order-rings", "Размеры колец"),
    ("order-export", "Excel"),
)


@dataclass(frozen=True)
class OrderItem:
    row: int
    set_id: str
    sku: str
    stone: str
    group: str
    sales: int
    stock_63: int
    stock_20: int
    stores: dict[str, int]
    total_stock: int
    working_stock: int
    ntr2_stock: int
    ntr2_calculated: bool
    tvp_raw: int
    stock_tt: int = 0
    image_path: str | None = None
    ungrouped: bool = False
    visual_match_set_id: str | None = None
    visual_match_sku: str | None = None
    visual_match_category: str | None = None
    visual_match_score: float = 0.0
    visual_match_status: str | None = None
    errors: tuple[str, ...] = ()

    @property
    def key(self) -> str:
        return f"{self.set_id}|{self.sku}|{self.row}"

    @property
    def positive_tvp(self) -> int:
        return max(0, self.tvp_raw)

    @property
    def is_ring(self) -> bool:
        value = normalize_text(self.group)
        if value in {"RING", "RINGS", "КОЛЬЦО", "КОЛЬЦА"}:
            return True
        return bool(re.search(r"(^|[^A-ZА-Я])RINGS?($|[^A-ZА-Я])", value))


@dataclass(frozen=True)
class OrderSet:
    key: str
    set_id: str
    stone: str
    items: tuple[OrderItem, ...]
    category: str
    driver_sku: str
    max_sales: int
    has_positive_tvp: bool
    has_negative_tvp: bool
    zero_segment: str | None = None
    is_ungrouped: bool = False


@dataclass(frozen=True)
class ParsedOrderWorkbook:
    source_name: str
    source_hash: str
    upload_path: str
    period: str
    supplier: str
    store_columns: tuple[str, ...]
    has_actual_ntr2: bool
    items: tuple[OrderItem, ...]
    warnings: tuple[str, ...] = ()


@dataclass
class OrderDraft:
    source_hash: str
    source_name: str
    mode: str
    version: int = DRAFT_VERSION
    orders: dict[str, int] = field(default_factory=dict)
    sizes: dict[str, dict[str, int]] = field(default_factory=dict)
    stock_checked: dict[str, bool] = field(default_factory=dict)
    stage: str = "order"
    selected_stone: str = ""
    updated_at: str = ""

    def touch(self) -> None:
        self.updated_at = datetime.now().isoformat(timespec="seconds")

    def as_payload(self) -> dict[str, Any]:
        self.touch()
        return {
            "version": self.version,
            "source_hash": self.source_hash,
            "source_name": self.source_name,
            "mode": self.mode,
            "orders": {str(k): int(v) for k, v in self.orders.items()},
            "sizes": {
                str(k): {str(size): int(qty) for size, qty in values.items()}
                for k, values in self.sizes.items()
            },
            "stock_checked": {str(k): bool(v) for k, v in self.stock_checked.items()},
            "stage": self.stage,
            "selected_stone": self.selected_stone,
            "updated_at": self.updated_at,
        }


# ---------------------------- pure business logic ----------------------------

@lru_cache(maxsize=1)
def load_order_exclusions() -> dict[str, Any]:
    defaults: dict[str, Any] = {
        "stone_patterns": list(STONE_EXCLUSION_PATTERNS),
        "pearl_patterns": list(SEA_PEARL_PATTERNS),
        "exclude_round_pearl": True,
    }
    try:
        payload = json.loads(ORDER_EXCLUSIONS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        return defaults
    if not isinstance(payload, dict):
        return defaults
    stone_patterns = payload.get("stone_patterns", defaults["stone_patterns"])
    pearl_patterns = payload.get("pearl_patterns", defaults["pearl_patterns"])
    return {
        "stone_patterns": [normalize_text(value) for value in stone_patterns if normalize_text(value)],
        "pearl_patterns": [normalize_text(value) for value in pearl_patterns if normalize_text(value)],
        "exclude_round_pearl": bool(payload.get("exclude_round_pearl", True)),
    }

def normalize_text(value: object) -> str:
    return " ".join(str(value or "").strip().upper().replace("Ё", "Е").split())


def safe_int(value: object) -> int:
    try:
        if value is None or str(value).strip() == "":
            return 0
        return int(round(float(value)))
    except (TypeError, ValueError):
        return 0


def _display_stone_name(text: str) -> str:
    value = text.title()
    replacements = {
        "Cz": "CZ",
        "Bt": "BT",
        "Mlbt": "MLBT",
        "Hq": "HQ",
        "Mq": "MQ",
        "Mop": "MOP",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)
    return value


STONE_EXACT_ALIASES: dict[str, str] = {
    "BS": "BLUE SAPPHIRE",
    "BSHQ": "BLUE SAPPHIRE HIGH QUALITY",
    "BSMQ": "BLUE SAPPHIRE MEDIUM QUALITY",
    "LBT": "LONDON TOPAZ",
    "SWBT": "SWISS TOPAZ",
    "WBT": "WHITE TOPAZ",
    "BT": "BLUE TOPAZ",
    "CE": "CREATED EMERALD",
    "CD": "CHROME DIOPSIDE",
    "EM": "EMERALD",
    "PERI": "PERIDOT",
    "AMST": "AMETHYST",
    "CIT": "CITRINE",
    "RDT": "RHODOLITE",
    "GARN": "GARNET",
    "APAT": "APATITE",
    "TANZ": "TANZANITE",
    "IOL": "IOLITE",
    "WHO": "WHITE HOWLITE",
    "GAM": "GREEN AMETHYST",
    "DJ": "DALMATIAN JASPER",
    "OB": "OBSIDIAN",
    "RJ": "RED JASPER",
    "GA": "GREEN AGATE",
    "LAP": "LAPIS LAZURITE",
    "AMA": "НЕ РАСПОЗНАНО (AMA)",
    "FPW": "FRESHWATER PEARL WHITE",
    "FPC": "FRESHWATER PEARL COLORED",
    "TAH": "TAHITI PEARL",
    "SSP": "SOUTH SEA PEARL",
}


def canonical_stone(value: object, sku: object = "") -> str:
    """Normalize supplier stone names with the same aliases used in reports.

    The order workspace keeps meaningful stone distinctions (for example
    Citrine versus Smoky) while removing spelling variants and abbreviations.
    """
    text = normalize_text(value)
    sku_text = normalize_text(sku)
    if not text:
        text = sku_text
    if text in STONE_EXACT_ALIASES:
        return _display_stone_name(STONE_EXACT_ALIASES[text])

    replacements = (
        ("LAPIS LAZULI", "LAPIS LAZURITE"),
        ("LAPIZ", "LAPIS LAZURITE"),
        ("BLACK AGATE", "AGATE"),
        ("CREATED EMEARLD", "CREATED EMERALD"),
        ("CREATED EMEALD", "CREATED EMERALD"),
        ("CREATE EMERALD", "CREATED EMERALD"),
        ("EMREAL", "EMERALD"),
        ("WHITETOPAZ", "WHITE TOPAZ"),
        ("WHIT TOPAZ", "WHITE TOPAZ"),
        ("IHOLIT", "IOLITE"),
        ("GERY PEARL", "GREY PEARL"),
        ("GREY PARL", "GREY PEARL"),
        ("WHITEE PEARL", "WHITE PEARL"),
        ("WHITE  PEARL", "WHITE PEARL"),
        ("FRESH WATER", "FRESHWATER"),
        ("MYSTMB", "MYSTIC"),
        ("MYST MB", "MYSTIC"),
        ("MOISANITE", "MOISSANITE"),
        ("MOSSANITE", "MOISSANITE"),
        ("MUSSONITE", "MOISSANITE"),
        ("SAPPHRIE", "SAPPHIRE"),
        ("SAPPHIRE", "SAPPHIRE"),
    )
    for old, new in replacements:
        text = text.replace(old, new)
    text = re.sub(r"\s+", " ", text).strip()
    if text == "LAZURITE":
        text = "LAPIS LAZURITE"

    # Priority and abbreviation rules mirror the report logic for mixed names.
    combined = f"{text} {sku_text}".strip()
    if re.search(r"MO+I?S+A?N+I?T|MOSSANIT|MUSSONIT", combined):
        return "Moissanite"
    if "BSHQ" in combined or "BLUE SAPPHIRE HIGH QUALITY" in combined or "BLUE SAPPHIRE HQ" in combined:
        return "Blue Sapphire High Quality"
    if "BSMQ" in combined or "BLUE SAPPHIRE MEDIUM QUALITY" in combined or "BLUE SAPPHIRE MQ" in combined:
        return "Blue Sapphire Medium Quality"
    if re.search(r"SAPP+H?I?R?E|SAPPPHIRE", combined):
        return "Blue Sapphire"
    if "RUBY" in combined and "ZOISITE" not in combined and "CIOSITE" not in combined:
        return "Ruby"
    if ("LONDON" in combined or re.search(r"(?:^|[-_/\s])LBT(?:$|[-_/\s])", combined)) and ("TOPAZ" in combined or "BT" in combined):
        return "London Topaz"
    if ("SWISS" in combined or "SWIS" in combined or re.search(r"(?:^|[-_/\s])SWBT(?:$|[-_/\s])", combined)) and ("TOPAZ" in combined or "BT" in combined):
        return "Swiss Topaz"
    if any(token in combined for token in ("WHITE TOPAZ", "BLUE TOPAZ", "MLBT", "MULTI BT")):
        return "Other Topaz"
    if "CREATED EMERALD" in combined:
        return "Created Emerald"
    if "CHROME DIOPSIDE" in combined or "DIOPOSIDE" in combined:
        return "Chrome Diopside"
    if "EMERALD" in combined:
        return "Emerald"
    if "GREEN AMETHYST" in combined:
        return "Green Amethyst"
    if "AMETHYST" in combined:
        return "Amethyst"
    if "RHODOLITE" in combined or "RODOLITE" in combined:
        return "Rhodolite"
    if "GARNET" in combined:
        return "Garnet"
    if "CITRINE" in combined:
        return "Citrine"
    if "ROSE QUARTZ" in combined:
        return "Rose Quartz"
    if "WHITE QUARTZ" in combined:
        return "White Quartz"
    if "SMOKY" in combined or "SMOKEY" in combined:
        return "Smoky"
    if "HONEY" in combined:
        return "Honey"
    if "MYSTIC" in combined:
        return "Mystic"
    if "GREEN AGATE" in combined:
        return "Green Agate"
    if "AGATE" in combined:
        return "Agate"
    if "BLACK SPINEL" in combined:
        return "Black Spinel"
    if "ONYX" in combined:
        return "Onyx"
    if "OBSIDIAN" in combined:
        return "Obsidian"
    if "IOLITE" in combined:
        return "Iolite"
    if "TANZANITE" in combined or "TANZNITE" in combined:
        return "Tanzanite"
    if "PERIDOT" in combined:
        return "Peridot"
    if "OPAL" in combined:
        return "Opal"
    if "TOURMALINE" in combined:
        return "Tourmaline"

    if text in STONE_EXACT_ALIASES:
        text = STONE_EXACT_ALIASES[text]
    return _display_stone_name(text) if text else "Не указан"


def canonical_group(value: object) -> str:
    text = normalize_text(value)
    aliases = {
        "EARRING": "Earrings",
        "EARRINGS": "Earrings",
        "СЕРЬГИ": "Earrings",
        "RING": "Ring",
        "RINGS": "Ring",
        "КОЛЬЦО": "Ring",
        "КОЛЬЦА": "Ring",
        "PENDANT": "Pendant",
        "PENDANTS": "Pendant",
        "ПОДВЕСКА": "Pendant",
        "ПОДВЕСКИ": "Pendant",
        "BRACELET": "Bracelet",
        "BRACELETS": "Bracelet",
        "БРАСЛЕТ": "Bracelet",
        "БРАСЛЕТЫ": "Bracelet",
        "NECKLACE": "Necklace",
        "NECKLACES": "Necklace",
        "ОЖЕРЕЛЬЕ": "Necklace",
    }
    return aliases.get(text, text.title() if text else "Не указана")


def is_tt_outlet_store(value: object) -> bool:
    text = normalize_text(value)
    if "STOCK" in text or "СКЛАД" in text:
        return False
    return text in {"OUTLET", "TT", "TT OUTLET", "OUTLET TT", "ТТ"}


def is_pearl_name(value: object) -> bool:
    text = normalize_text(value)
    return "PEARL" in text or "PARL" in text


def is_excluded_pearl(value: object) -> bool:
    text = normalize_text(value)
    settings = load_order_exclusions()
    if any(pattern in text for pattern in settings["pearl_patterns"]):
        return True
    # The current report uses both FRESH WATER ROUND PEARL and shorter
    # ROUND ... PEARL spellings for spherical freshwater pearl.
    if settings["exclude_round_pearl"] and "PEARL" in text and "ROUND" in text:
        return True
    return False


def is_excluded_stone(value: object) -> bool:
    text = normalize_text(value)
    return any(pattern in text for pattern in load_order_exclusions()["stone_patterns"])


def item_in_mode(item: OrderItem, mode: str) -> bool:
    pearl = is_pearl_name(item.stone)
    if mode == ORDER_MODE_PEARLS:
        return pearl and not is_excluded_pearl(item.stone)
    return (not pearl) and not is_excluded_stone(item.stone)


def classify_set(items: Iterable[OrderItem]) -> tuple[str, str, int, str | None]:
    materialized = tuple(items)
    if not materialized:
        return CATEGORY_ZERO, "", 0, "0/0"
    driver = max(materialized, key=lambda item: (item.sales, -item.row))
    maximum = max(0, int(driver.sales))
    if maximum >= 5:
        category = CATEGORY_TOP
    elif maximum >= 3:
        category = CATEGORY_MEDIUM
    elif maximum >= 1:
        category = CATEGORY_WEAK
    else:
        category = CATEGORY_ZERO
    zero_segment = None
    if category == CATEGORY_ZERO:
        zero_segment = "Нулевые с остатком" if any(item.working_stock > 0 for item in materialized) else "0/0 — не было остатка"
    return category, driver.sku, maximum, zero_segment


def build_order_sets(items: Iterable[OrderItem], mode: str) -> tuple[OrderSet, ...]:
    normal_groups: dict[str, list[OrderItem]] = {}
    normal_order: list[str] = []
    ungrouped_items: list[OrderItem] = []

    for item in items:
        if not item_in_mode(item, mode):
            continue
        if item.ungrouped or normalize_text(item.set_id) == "БЕЗ КОМПЛЕКТА":
            ungrouped_items.append(item)
            continue
        if item.set_id not in normal_groups:
            normal_groups[item.set_id] = []
            normal_order.append(item.set_id)
        normal_groups[item.set_id].append(item)

    result: list[OrderSet] = []
    for set_id in normal_order:
        group_items = tuple(normal_groups[set_id])
        category, driver_sku, max_sales, zero_segment = classify_set(group_items)
        stone_counts: dict[str, int] = {}
        for item in group_items:
            stone = canonical_stone(item.stone, item.sku)
            stone_counts[stone] = stone_counts.get(stone, 0) + 1
        primary_stone = max(
            stone_counts,
            key=lambda name: (
                stone_counts[name],
                -next(item.row for item in group_items if canonical_stone(item.stone, item.sku) == name),
            ),
        )
        has_positive = any(item.tvp_raw > 0 for item in group_items)
        has_negative = any(item.tvp_raw < 0 for item in group_items)
        result.append(OrderSet(
            key=f"{mode}|{set_id}",
            set_id=set_id,
            stone=primary_stone,
            items=group_items,
            category=category,
            driver_sku=driver_sku,
            max_sales=max_sales,
            has_positive_tvp=has_positive,
            has_negative_tvp=has_negative,
            zero_segment=zero_segment,
            is_ungrouped=False,
        ))

    # Standalone rows are not allowed to promote each other. They are first
    # classified one by one, then collected into one virtual "Без комплекта"
    # block per stone/category. TVP rows are kept in a separate virtual block
    # so one item in transit never hides all other standalone models.
    virtual_groups: dict[tuple[str, str, str | None, str], list[OrderItem]] = {}
    for item in ungrouped_items:
        category, _driver, _maximum, zero_segment = classify_set((item,))
        stone = canonical_stone(item.stone, item.sku)
        transit_bucket = "tvp" if item.tvp_raw > 0 else "regular"
        key = (stone, category, zero_segment, transit_bucket)
        virtual_groups.setdefault(key, []).append(item)

    category_rank = {category: index for index, category in enumerate(CATEGORY_ORDER)}
    for (stone, category, zero_segment, transit_bucket), grouped_items in sorted(
        virtual_groups.items(),
        key=lambda pair: (pair[0][0], category_rank[pair[0][1]], pair[0][2] or "", pair[0][3]),
    ):
        materialized = tuple(sorted(grouped_items, key=lambda item: item.row))
        driver = max(materialized, key=lambda item: (item.sales, -item.row))
        result.append(OrderSet(
            key=f"{mode}|ungrouped|{stone}|{category}|{zero_segment}|{transit_bucket}",
            set_id="Без комплекта",
            stone=stone,
            items=materialized,
            category=category,
            driver_sku=driver.sku,
            max_sales=driver.sales,
            has_positive_tvp=(transit_bucket == "tvp"),
            has_negative_tvp=any(item.tvp_raw < 0 for item in materialized),
            zero_segment=zero_segment,
            is_ungrouped=True,
        ))
    return tuple(result)


def suggested_order_quantity(item: OrderItem) -> int:
    """First practical draft, intentionally transparent and editable.

    Weak/zero models are not pre-ordered. Medium/top models are suggested only
    when the working stock is low (0-3). Sales 3-6 start from 5 pieces; higher
    sales are rounded up to a multiple of five. Positive TVP reduces only this
    new order suggestion. Working stock is an eligibility check, not a hidden
    subtraction from the order quantity.
    """
    if item.sales < 3 or item.working_stock > 3:
        return 0
    base = 5 if item.sales <= 6 else int(math.ceil(item.sales / 5.0) * 5)
    return max(0, base - item.positive_tvp)


def infer_ntr2(total: int, store_values: dict[str, int], has_actual_ntr2: bool) -> tuple[int, bool, str | None]:
    normalized = {normalize_text(name): safe_int(value) for name, value in store_values.items()}
    if has_actual_ntr2:
        actual = normalized.get("NTR2", 0)
        delta = total - sum(normalized.values())
        warning = None if delta == 0 else f"Сумма магазинов отличается от «Всего» на {delta} шт."
        return max(0, actual), False, warning
    inferred = total - sum(normalized.values())
    if inferred < 0:
        return 0, True, f"Сумма магазинов превышает «Всего» на {abs(inferred)} шт."
    return inferred, True, None


@dataclass(frozen=True)
class _ImageSignature:
    digest: str
    dhash: int
    histogram: tuple[int, ...]
    aspect: float


def _normalized_image(payload: bytes) -> Image.Image | None:
    try:
        image = Image.open(io.BytesIO(payload)).convert("RGB")
    except (UnidentifiedImageError, OSError, ValueError):
        return None
    if image.width <= 1 or image.height <= 1:
        return None

    # Remove the mostly white supplier-photo margins before comparison.
    corners = [
        image.getpixel((0, 0)),
        image.getpixel((image.width - 1, 0)),
        image.getpixel((0, image.height - 1)),
        image.getpixel((image.width - 1, image.height - 1)),
    ]
    background = tuple(sum(pixel[channel] for pixel in corners) // len(corners) for channel in range(3))
    background_image = Image.new("RGB", image.size, background)
    difference = ImageChops.difference(image, background_image).convert("L")
    mask = difference.point(lambda value: 255 if value > 18 else 0)
    bbox = mask.getbbox()
    if bbox:
        left, top, right, bottom = bbox
        pad_x = max(2, int((right - left) * 0.05))
        pad_y = max(2, int((bottom - top) * 0.05))
        bbox = (
            max(0, left - pad_x),
            max(0, top - pad_y),
            min(image.width, right + pad_x),
            min(image.height, bottom + pad_y),
        )
        image = image.crop(bbox)

    image = ImageOps.contain(image, (128, 128), method=Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (128, 128), (255, 255, 255))
    offset = ((128 - image.width) // 2, (128 - image.height) // 2)
    canvas.paste(image, offset)
    return canvas


def _flat_image_data(image: Image.Image):
    getter = getattr(image, "get_flattened_data", None)
    return getter() if callable(getter) else image.getdata()


def _image_signature(payload: bytes) -> _ImageSignature | None:
    image = _normalized_image(payload)
    if image is None:
        return None
    digest = hashlib.sha1(image.tobytes()).hexdigest()
    gray = image.convert("L").resize((17, 16), Image.Resampling.LANCZOS)
    pixels = list(_flat_image_data(gray))
    dhash = 0
    for row in range(16):
        base = row * 17
        for column in range(16):
            dhash = (dhash << 1) | int(pixels[base + column] > pixels[base + column + 1])

    small = image.resize((32, 32), Image.Resampling.BILINEAR)
    histogram = [0] * 64
    for red, green, blue in _flat_image_data(small):
        bucket = (red // 64) * 16 + (green // 64) * 4 + (blue // 64)
        histogram[min(63, bucket)] += 1
    return _ImageSignature(
        digest=digest,
        dhash=dhash,
        histogram=tuple(histogram),
        aspect=image.width / max(1, image.height),
    )


def _signature_similarity(left: _ImageSignature, right: _ImageSignature) -> float:
    if left.digest == right.digest:
        return 1.0
    hash_similarity = 1.0 - ((left.dhash ^ right.dhash).bit_count() / 256.0)
    histogram_similarity = sum(min(a, b) for a, b in zip(left.histogram, right.histogram)) / 1024.0
    aspect_similarity = min(left.aspect, right.aspect) / max(left.aspect, right.aspect)
    return 0.74 * hash_similarity + 0.21 * histogram_similarity + 0.05 * aspect_similarity


def _annotate_ungrouped_visual_matches(archive: ZipFile, items: list[OrderItem]) -> list[OrderItem]:
    """Attach conservative photo-match hints to items from <Без комплекта>."""
    grouped_items = [item for item in items if not item.ungrouped and item.image_path]
    ungrouped_items = [item for item in items if item.ungrouped and item.image_path]
    if not grouped_items or not ungrouped_items:
        return items

    needed_paths = {item.image_path for item in grouped_items + ungrouped_items if item.image_path}
    signatures: dict[str, _ImageSignature] = {}
    archive_names = set(archive.namelist())
    for image_path in needed_paths:
        if not image_path or image_path not in archive_names:
            continue
        signature = _image_signature(archive.read(image_path))
        if signature is not None:
            signatures[image_path] = signature

    candidates: dict[tuple[str, str], list[OrderItem]] = {}
    set_categories: dict[str, str] = {}
    by_set: dict[str, list[OrderItem]] = {}
    for item in grouped_items:
        by_set.setdefault(item.set_id, []).append(item)
        key = (canonical_stone(item.stone, item.sku), canonical_group(item.group))
        candidates.setdefault(key, []).append(item)
    for set_id, set_items in by_set.items():
        set_categories[set_id] = classify_set(set_items)[0]

    replacements: dict[str, OrderItem] = {}
    for item in ungrouped_items:
        signature = signatures.get(item.image_path or "")
        if signature is None:
            continue
        key = (canonical_stone(item.stone, item.sku), canonical_group(item.group))
        possible = [candidate for candidate in candidates.get(key, []) if candidate.image_path in signatures]
        if not possible:
            continue
        scored = sorted(
            ((_signature_similarity(signature, signatures[candidate.image_path or ""]), candidate) for candidate in possible),
            key=lambda pair: (-pair[0], pair[1].row),
        )
        best_score, best = scored[0]
        if best_score < 0.94:
            continue
        close_alternatives = {candidate.set_id for score, candidate in scored[1:4] if best_score - score <= 0.012}
        confirmed = best_score >= 0.965 and not close_alternatives
        replacements[item.key] = replace(
            item,
            visual_match_set_id=best.set_id,
            visual_match_sku=best.sku,
            visual_match_category=set_categories.get(best.set_id),
            visual_match_score=round(best_score, 3),
            visual_match_status="confirmed" if confirmed else "possible",
        )
    return [replacements.get(item.key, item) for item in items]


# ---------------------------- XLSX parser ------------------------------------

def _shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in si.iter(_XML_MAIN + "t")) for si in root.findall(_XML_MAIN + "si")]


def _cell_value(cell: ET.Element, strings: list[str]) -> str:
    value = cell.find(_XML_MAIN + "v")
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(_XML_MAIN + "t"))
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return strings[int(value.text)]
        except (ValueError, IndexError):
            return ""
    return value.text


def _read_sheet_rows(archive: ZipFile, sheet_path: str, strings: list[str]) -> dict[int, dict[str, str]]:
    """Stream worksheet rows without keeping the large XML tree in memory."""
    rows: dict[int, dict[str, str]] = {}
    with archive.open(sheet_path) as handle:
        for _event, element in ET.iterparse(handle, events=("end",)):
            if element.tag != _XML_MAIN + "row":
                continue
            row_number = int(element.attrib.get("r", "0") or 0)
            rows[row_number] = {
                _column_letter(cell.attrib.get("r", "")): _cell_value(cell, strings)
                for cell in element.findall(_XML_MAIN + "c")
            }
            element.clear()
    return rows


def _drawing_relationship_id(archive: ZipFile, sheet_path: str) -> str:
    with archive.open(sheet_path) as handle:
        for _event, element in ET.iterparse(handle, events=("end",)):
            if element.tag == _XML_MAIN + "drawing":
                return element.attrib.get(_XML_REL + "id", "")
            element.clear()
    return ""


def _column_letter(cell_ref: str) -> str:
    match = re.match(r"([A-Z]+)", cell_ref or "")
    return match.group(1) if match else ""


def _workbook_sheet_path(archive: ZipFile) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    sheets = workbook.find(_XML_MAIN + "sheets")
    if sheets is None or not list(sheets):
        raise ValueError("В книге нет листов.")
    sheet = list(sheets)[0]
    rel_id = sheet.attrib[_XML_REL + "id"]
    target = relmap[rel_id].lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def _sheet_relationships(archive: ZipFile, sheet_path: str) -> dict[str, str]:
    folder, filename = posixpath.split(sheet_path)
    rel_path = posixpath.join(folder, "_rels", filename + ".rels")
    if rel_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rel_path))
    result: dict[str, str] = {}
    for rel in root:
        target = rel.attrib.get("Target", "")
        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        result[rel.attrib.get("Id", "")] = resolved
    return result


def _image_index(archive: ZipFile, sheet_path: str) -> dict[int, str]:
    drawing_rel_id = _drawing_relationship_id(archive, sheet_path)
    if not drawing_rel_id:
        return {}
    sheet_rels = _sheet_relationships(archive, sheet_path)
    drawing_path = sheet_rels.get(drawing_rel_id)
    if not drawing_path or drawing_path not in archive.namelist():
        return {}

    folder, filename = posixpath.split(drawing_path)
    rel_path = posixpath.join(folder, "_rels", filename + ".rels")
    relmap: dict[str, str] = {}
    if rel_path in archive.namelist():
        root = ET.fromstring(archive.read(rel_path))
        for rel in root:
            target = rel.attrib.get("Target", "")
            resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
            relmap[rel.attrib.get("Id", "")] = resolved

    result: dict[int, str] = {}
    with archive.open(drawing_path) as handle:
        for _event, anchor in ET.iterparse(handle, events=("end",)):
            if anchor.tag not in {_XML_DRAWING + "twoCellAnchor", _XML_DRAWING + "oneCellAnchor"}:
                continue
            start = anchor.find(_XML_DRAWING + "from")
            picture = anchor.find(_XML_DRAWING + "pic")
            if start is not None and picture is not None:
                row_node = start.find(_XML_DRAWING + "row")
                blip = picture.find(".//" + _XML_A + "blip")
                if row_node is not None and row_node.text is not None and blip is not None:
                    media_path = relmap.get(blip.attrib.get(_XML_REL + "embed", ""))
                    if media_path:
                        # Drawing coordinates are zero-based; worksheet rows are one-based.
                        result.setdefault(int(row_node.text) + 1, media_path)
            anchor.clear()
    return result


def _extract_period_and_supplier(rows: dict[int, dict[str, str]]) -> tuple[str, str]:
    period = ""
    supplier = ""
    for row_number in range(1, 8):
        text = " ".join(rows.get(row_number, {}).values())
        if "Продажи товаров за период" in text:
            period = text.replace("Продажи товаров за период", "").strip()
        if "Поставщик(и):" in text:
            supplier = text.split("Поставщик(и):", 1)[-1].strip()
    return period, supplier


def _detect_columns(rows: dict[int, dict[str, str]]) -> tuple[str, str, list[str], dict[str, str]]:
    row7 = rows.get(7, {})
    row8 = rows.get(8, {})
    sales_col = next((col for col, value in row7.items() if normalize_text(value) == "ПРОДАЖИ ЗА ПЕРИОД"), "E")
    stock_start_col = next((col for col, value in row7.items() if normalize_text(value) in {"ОСТАТКИ", "ОСТАТОК"}), "G")
    tvp_col = next((col for col, value in row7.items() if normalize_text(value) == "ТВП"), "O")
    total_col = next((col for col, value in row8.items() if normalize_text(value) in {"ВСЕГО", "TOTAL"}), "N")

    def col_number(letter: str) -> int:
        result = 0
        for character in letter:
            result = result * 26 + (ord(character) - 64)
        return result

    stock_start_number = col_number(stock_start_col)
    total_number = col_number(total_col)
    store_columns: list[str] = []
    store_names: dict[str, str] = {}
    for col, value in row8.items():
        number = col_number(col)
        name = " ".join(str(value or "").strip().split())
        if stock_start_number <= number < total_number and name:
            store_columns.append(col)
            store_names[col] = name
    return sales_col, tvp_col, store_columns, {**store_names, "__total__": total_col}


def parse_order_workbook(path: str | Path, source_name: str | None = None, source_hash: str | None = None) -> ParsedOrderWorkbook:
    workbook_path = Path(path)
    source_name = source_name or workbook_path.name
    if source_hash is None:
        digest = hashlib.sha256()
        with workbook_path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        source_hash = digest.hexdigest()

    with ZipFile(workbook_path) as archive:
        sheet_path = _workbook_sheet_path(archive)
        strings = _shared_strings(archive)
        row_values = _read_sheet_rows(archive, sheet_path, strings)
        if not row_values:
            raise ValueError("В листе нет данных.")

        period, supplier = _extract_period_and_supplier(row_values)
        sales_col, tvp_col, store_cols, names = _detect_columns(row_values)
        total_col = names.pop("__total__")
        image_index = _image_index(archive, sheet_path)
        actual_ntr2 = any(normalize_text(name) == "NTR2" for name in names.values())

        current_set = ""
        in_ungrouped_section = False
        items: list[OrderItem] = []
        workbook_warnings: list[str] = []
        for row_number in sorted(row_values):
            if row_number < 11:
                continue
            values = row_values[row_number]
            first = str(values.get("A", "") or "").strip()
            normalized_first = normalize_text(first).strip("<>")
            if normalized_first.startswith("SET#"):
                current_set = first
                in_ungrouped_section = False
                continue
            if normalized_first == "БЕЗ КОМПЛЕКТА":
                current_set = "Без комплекта"
                in_ungrouped_section = True
                continue
            stone = str(values.get("B", "") or "").strip()
            group = str(values.get("C", "") or "").strip()
            if not first or not stone or not group:
                continue
            if not current_set:
                current_set = "Без комплекта"
                in_ungrouped_section = True

            stores = {names[col]: safe_int(values.get(col)) for col in store_cols if col in names}
            total = safe_int(values.get(total_col))
            stock_63 = next((qty for name, qty in stores.items() if normalize_text(name) == "63"), 0)
            stock_20 = next((qty for name, qty in stores.items() if normalize_text(name) == "20"), 0)
            stock_tt = sum(qty for name, qty in stores.items() if is_tt_outlet_store(name))
            ntr2, calculated, ntr2_warning = infer_ntr2(total, stores, actual_ntr2)
            working_raw = total - stock_63 - stock_20
            errors: list[str] = []
            if working_raw < 0:
                errors.append(f"Рабочий остаток отрицательный: {working_raw}")
            if ntr2_warning:
                errors.append(ntr2_warning)
            tvp = safe_int(values.get(tvp_col))
            if tvp < 0:
                errors.append(f"Ошибка ТВП: {tvp}")
            items.append(OrderItem(
                row=row_number,
                set_id=current_set,
                sku=first,
                stone=stone,
                group=group,
                sales=max(0, safe_int(values.get(sales_col))),
                stock_63=max(0, stock_63),
                stock_20=max(0, stock_20),
                stores=stores,
                total_stock=max(0, total),
                working_stock=max(0, working_raw),
                ntr2_stock=max(0, ntr2),
                ntr2_calculated=calculated,
                tvp_raw=tvp,
                stock_tt=max(0, stock_tt),
                image_path=image_index.get(row_number),
                ungrouped=in_ungrouped_section,
                errors=tuple(errors),
            ))

        if not items:
            raise ValueError("Не найдены строки изделий. Проверьте структуру отчёта.")
        items = _annotate_ungrouped_visual_matches(archive, items)
        if not actual_ntr2:
            workbook_warnings.append("Колонки NTR2 пока нет: остаток NTR2 восстановлен как «Всего минус все явные магазины».")
        return ParsedOrderWorkbook(
            source_name=source_name,
            source_hash=source_hash,
            upload_path=str(workbook_path),
            period=period,
            supplier=supplier,
            store_columns=tuple(names.values()),
            has_actual_ntr2=actual_ntr2,
            items=tuple(items),
            warnings=tuple(workbook_warnings),
        )


def store_uploaded_workbook(name: str, payload: bytes) -> tuple[Path, str]:
    digest = hashlib.sha256(payload).hexdigest()
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    suffix = Path(name).suffix.lower() or ".xlsx"
    target = UPLOAD_DIR / f"{digest}{suffix}"
    if not target.exists() or target.stat().st_size != len(payload):
        temporary = target.with_suffix(target.suffix + ".tmp")
        temporary.write_bytes(payload)
        temporary.replace(target)
    return target, digest


@st.cache_resource(show_spinner=False, max_entries=6)
def cached_parse_order_workbook(path: str, source_name: str, source_hash: str) -> ParsedOrderWorkbook:
    return parse_order_workbook(path, source_name=source_name, source_hash=source_hash)


@st.cache_data(show_spinner=False, max_entries=12)
def load_visible_images(path: str, image_paths: tuple[str, ...]) -> dict[str, bytes]:
    result: dict[str, bytes] = {}
    if not image_paths:
        return result
    with ZipFile(path) as archive:
        names = set(archive.namelist())
        for image_path in image_paths:
            if image_path in names:
                result[image_path] = archive.read(image_path)
    return result


# ---------------------------- draft persistence ------------------------------

def _connect_drafts() -> sqlite3.Connection:
    DRAFT_DB.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DRAFT_DB, timeout=15)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS order_drafts (
            draft_key TEXT PRIMARY KEY,
            source_hash TEXT NOT NULL,
            mode TEXT NOT NULL,
            payload TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    return connection


def draft_key(source_hash: str, mode: str) -> str:
    return hashlib.sha256(f"{source_hash}|{mode}".encode("utf-8")).hexdigest()


def validate_draft_payload(payload: object) -> OrderDraft:
    if not isinstance(payload, dict):
        raise ValueError("Черновик должен быть JSON-объектом.")
    mode = str(payload.get("mode", ""))
    if mode not in ORDER_MODES:
        raise ValueError("В черновике не указан корректный тип заказа.")

    payload_version = max(1, safe_int(payload.get("version", 1)))
    if payload_version < DRAFT_VERSION:
        # Version 1 automatically prefilled recommendations. The new workflow
        # starts every item from zero, therefore legacy auto-seeded quantities
        # must not silently appear in the final Excel.
        return OrderDraft(
            source_hash=str(payload.get("source_hash", "")),
            source_name=str(payload.get("source_name", "")),
            mode=mode,
            version=DRAFT_VERSION,
        )

    orders = {str(k): max(0, safe_int(v)) for k, v in dict(payload.get("orders", {})).items()}
    sizes: dict[str, dict[str, int]] = {}
    for key, values in dict(payload.get("sizes", {})).items():
        if isinstance(values, dict):
            sizes[str(key)] = {str(size): max(0, safe_int(qty)) for size, qty in values.items() if str(size) in {str(x) for x in RING_SIZES}}
    return OrderDraft(
        source_hash=str(payload.get("source_hash", "")),
        source_name=str(payload.get("source_name", "")),
        mode=mode,
        version=DRAFT_VERSION,
        orders=orders,
        sizes=sizes,
        stock_checked={str(k): bool(v) for k, v in dict(payload.get("stock_checked", {})).items()},
        stage=str(payload.get("stage", "order")) if str(payload.get("stage", "order")) in {"order", "rings"} else "order",
        selected_stone=str(payload.get("selected_stone", "")),
        updated_at=str(payload.get("updated_at", "")),
    )


def load_draft(source_hash: str, source_name: str, mode: str) -> OrderDraft:
    key = draft_key(source_hash, mode)
    try:
        with _connect_drafts() as connection:
            row = connection.execute("SELECT payload FROM order_drafts WHERE draft_key = ?", (key,)).fetchone()
    except sqlite3.Error:
        row = None
    if row:
        try:
            draft = validate_draft_payload(json.loads(row[0]))
            draft.source_hash = source_hash
            draft.source_name = source_name
            draft.mode = mode
            return draft
        except (ValueError, TypeError, json.JSONDecodeError):
            pass
    return OrderDraft(source_hash=source_hash, source_name=source_name, mode=mode)


def save_draft(draft: OrderDraft) -> str:
    payload = draft.as_payload()
    serialized = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    key = draft_key(draft.source_hash, draft.mode)
    with _connect_drafts() as connection:
        connection.execute(
            """
            INSERT INTO order_drafts(draft_key, source_hash, mode, payload, updated_at)
            VALUES(?, ?, ?, ?, ?)
            ON CONFLICT(draft_key) DO UPDATE SET
                payload = excluded.payload,
                updated_at = excluded.updated_at
            """,
            (key, draft.source_hash, draft.mode, serialized, draft.updated_at),
        )
        connection.commit()
    return draft.updated_at


def draft_json_bytes(draft: OrderDraft) -> bytes:
    return json.dumps(draft.as_payload(), ensure_ascii=False, indent=2).encode("utf-8")


def import_draft_json(payload: bytes, expected_hash: str, mode: str) -> OrderDraft:
    try:
        raw = json.loads(payload.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("Файл черновика не распознан.") from exc
    draft = validate_draft_payload(raw)
    if draft.mode != mode:
        raise ValueError(f"Этот черновик относится к заказу «{draft.mode}».")
    if draft.source_hash and draft.source_hash != expected_hash:
        raise ValueError("Черновик создан по другому исходному отчёту.")
    draft.source_hash = expected_hash
    save_draft(draft)
    return draft


# ---------------------------- Excel export -----------------------------------

def format_sizes(values: dict[str, int] | None) -> str:
    values = values or {}
    parts = [f"{size} × {safe_int(values.get(str(size), 0))}" for size in RING_SIZES if safe_int(values.get(str(size), 0)) > 0]
    return "; ".join(parts)


def build_supplier_excel(
    parsed: ParsedOrderWorkbook,
    selected_items: Iterable[OrderItem],
    draft: OrderDraft,
) -> bytes:
    items = [item for item in selected_items if draft.orders.get(item.key, 0) > 0]
    image_paths = tuple(sorted({item.image_path for item in items if item.image_path}))
    images = load_visible_images(parsed.upload_path, image_paths)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Order"
    headers = ["Фото", "Артикул", "Камень", "Группа", "Количество к заказу", "Размеры"]
    sheet.append(headers)

    header_fill = PatternFill("solid", fgColor="1C1A17")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin = Side(style="thin", color="D9D2C4")
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)
    sheet.row_dimensions[1].height = 28

    for row_index, item in enumerate(items, start=2):
        quantity = max(0, safe_int(draft.orders.get(item.key, 0)))
        sizes = format_sizes(draft.sizes.get(item.key)) if item.is_ring else ""
        sheet.append(["", item.sku, canonical_stone(item.stone), item.group, quantity, sizes])
        sheet.row_dimensions[row_index].height = 66
        for cell in sheet[row_index]:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="hair", color="E6E0D8"))
        image_data = images.get(item.image_path or "")
        if image_data:
            try:
                image = XLImage(io.BytesIO(image_data))
                image.width = 72
                image.height = 72
                sheet.add_image(image, f"A{row_index}")
            except Exception:
                sheet.cell(row_index, 1).value = "Фото не вставлено"

    widths = {"A": 14, "B": 27, "C": 28, "D": 18, "E": 22, "F": 38}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:F{max(1, sheet.max_row)}"

    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


# ---------------------------- Streamlit UI -----------------------------------

def _draft_state_key(source_hash: str, mode: str) -> str:
    return f"supplier_order_draft::{source_hash}::{mode}"


def _get_session_draft(parsed: ParsedOrderWorkbook, mode: str) -> OrderDraft:
    key = _draft_state_key(parsed.source_hash, mode)
    if key not in st.session_state:
        st.session_state[key] = load_draft(parsed.source_hash, parsed.source_name, mode)
    draft = st.session_state[key]
    if not isinstance(draft, OrderDraft) or getattr(draft, "version", 1) < DRAFT_VERSION:
        draft = load_draft(parsed.source_hash, parsed.source_name, mode)
        st.session_state[key] = draft
    return draft


def _save_session_draft(draft: OrderDraft) -> None:
    try:
        saved_at = save_draft(draft)
        st.session_state["supplier_order_save_status"] = f"Сохранено: {saved_at.replace('T', ' ')}"
    except (sqlite3.Error, OSError) as exc:
        st.session_state["supplier_order_save_status"] = f"Не удалось записать черновик на сервер: {exc}"


def _render_sidebar(parsed: ParsedOrderWorkbook | None, draft: OrderDraft | None) -> None:
    items = [NavigationItem(item_id=section_id, label=label, href=f"#{section_id}") for section_id, label in ORDER_SECTIONS]
    status = st.session_state.get("supplier_order_save_status", "Черновик ещё не сохранён")
    source = parsed.source_name if parsed else "Ожидается Excel-отчёт"
    result = render_sidebar(
        module_title="Заказ поставщику",
        navigation_title="Навигация заказа",
        items=items,
        status_text=status,
        status_tone="success" if str(status).startswith("Сохранено") else "neutral",
        source_text=source,
        action_label="Сохранить черновик" if draft else None,
        action_key="supplier_order_manual_save" if draft else None,
    )
    if draft and result.action_clicked:
        _save_session_draft(draft)
        st.rerun()
    render_mobile_navigation(items)


def _render_upload() -> tuple[ParsedOrderWorkbook | None, bytes | None]:
    uploaded = st.file_uploader(
        "Загрузите отчёт для формирования заказа",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
        key="supplier_order_upload",
        help="Поддерживается текущий отчёт Y&J и будущая версия с отдельной колонкой NTR2.",
    )
    if uploaded is None:
        _render_sidebar(None, None)
        st.info("Загрузите файл «Заказ.xlsx». Прогноз, скорость продаж и готовые рекомендации использоваться не будут.")
        return None, None
    payload = bytes(uploaded.getvalue())
    path, digest = store_uploaded_workbook(uploaded.name, payload)
    with st.spinner("Читаем комплекты, остатки, ТВП и фотографии..."):
        parsed = cached_parse_order_workbook(str(path), uploaded.name, digest)
    return parsed, payload


def _mode_sets(parsed: ParsedOrderWorkbook, mode: str) -> tuple[OrderSet, ...]:
    return build_order_sets(parsed.items, mode)


def _ordered_items(order_sets: Iterable[OrderSet]) -> list[OrderItem]:
    return [item for order_set in order_sets for item in order_set.items]


def _seed_defaults(draft: OrderDraft, order_sets: Iterable[OrderSet]) -> None:
    changed = False
    for item in _ordered_items(order_sets):
        if item.key not in draft.orders:
            # Recommendations remain visible as hints, but the actual order
            # always starts from zero and changes only after a user action.
            draft.orders[item.key] = 0
            changed = True
    if changed:
        _save_session_draft(draft)


def _category_reason(order_set: OrderSet) -> str:
    if order_set.is_ungrouped:
        return f"{len(order_set.items)} отдельных позиций. Каждая отнесена к категории по собственным продажам."
    if order_set.category == CATEGORY_ZERO:
        return "Все изделия комплекта имеют 0 продаж."
    return f"Категорию определил артикул {order_set.driver_sku}: продано {order_set.max_sales} шт."


def _order_input_key(item: OrderItem, mode: str) -> str:
    return "order_qty::" + hashlib.sha1(f"v{DRAFT_VERSION}|{mode}|{item.key}".encode("utf-8")).hexdigest()


def _render_item_row(item: OrderItem, image_data: bytes | None, draft: OrderDraft, mode: str) -> bool:
    changed = False
    with st.container(border=True):
        photo, details, sales_col, stock_col, tt_col, tvp_col, order_col = st.columns(
            [1.0, 2.05, 0.65, 0.85, 0.78, 0.78, 1.0],
            vertical_alignment="center",
        )
        with photo:
            if image_data:
                st.image(image_data, width="stretch")
            else:
                st.caption("Нет фото")
        with details:
            st.markdown(f"**{item.sku}**")
            st.caption(f"{canonical_stone(item.stone, item.sku)} · {item.group}")
            if item.ntr2_stock > 0:
                suffix = "расчётный" if item.ntr2_calculated else "из файла"
                st.caption(f"NTR2: {item.ntr2_stock} ({suffix})")
            if item.visual_match_set_id:
                match_title = "Найдено визуальное совпадение" if item.visual_match_status == "confirmed" else "Возможное визуальное совпадение"
                message = (
                    f"{match_title}: **{item.visual_match_set_id}** · "
                    f"{item.visual_match_sku or 'артикул не указан'} · "
                    f"сходство {item.visual_match_score:.0%}"
                )
                if item.visual_match_status == "confirmed":
                    st.success(message, icon="🔎")
                else:
                    st.warning(message, icon="🔎")
                if st.button(
                    "Показать найденный комплект",
                    key="show_match::" + hashlib.sha1(f"{mode}|{item.key}".encode("utf-8")).hexdigest(),
                    width="stretch",
                ):
                    target_stone = canonical_stone(item.stone, item.sku)
                    draft.selected_stone = target_stone
                    if item.visual_match_category:
                        st.session_state[f"supplier_order_category::{mode}::{target_stone}"] = item.visual_match_category
                    st.session_state[f"supplier_order_focus_set::{mode}"] = item.visual_match_set_id
                    _save_session_draft(draft)
                    st.rerun()
            for error in item.errors:
                st.error(error, icon="⚠️")
        with sales_col:
            st.metric("Продажи", item.sales)
        with stock_col:
            st.metric("Рабочий остаток", item.working_stock)
            st.caption(f"63: {item.stock_63}")
        with tt_col:
            if item.stock_tt > 0:
                st.metric("Из них в ТТ", item.stock_tt)
        with tvp_col:
            if item.tvp_raw > 0:
                st.metric("ТВП", item.tvp_raw)
            elif item.tvp_raw < 0:
                st.markdown(f"**Ошибка ТВП:** :red[{item.tvp_raw}]")
            else:
                st.metric("ТВП", 0)
        with order_col:
            key = _order_input_key(item, mode)
            current = max(0, safe_int(draft.orders.get(item.key, 0)))
            if key not in st.session_state:
                st.session_state[key] = current
            value = st.number_input(
                "К заказу",
                min_value=0,
                max_value=999,
                step=1,
                value=current,
                key=key,
            )
            value = max(0, safe_int(value))
            if value != current:
                draft.orders[item.key] = value
                changed = True
            suggested = suggested_order_quantity(item)
            if suggested > 0:
                st.caption(f"Подсказка: {suggested}")
    return changed


def _render_set_card(order_set: OrderSet, images: dict[str, bytes], draft: OrderDraft, mode: str) -> bool:
    changed = False
    icon = CATEGORY_TONE[order_set.category]
    focused = st.session_state.get(f"supplier_order_focus_set::{mode}") == order_set.set_id
    with st.container(border=True):
        if focused:
            st.info("Найденный визуально похожий комплект", icon="🔎")
        header_left, header_right = st.columns([4, 1])
        with header_left:
            st.markdown(f"### {order_set.set_id}")
            st.caption(_category_reason(order_set))
        with header_right:
            st.markdown(f"**{icon} {order_set.category}**")
            if order_set.has_negative_tvp:
                st.error("Ошибка ТВП", icon="⚠️")
            elif order_set.has_positive_tvp:
                st.info("Есть товар в пути", icon="🚚")
        for item in order_set.items:
            changed = _render_item_row(item, images.get(item.image_path or ""), draft, mode) or changed
    return changed


def _render_sets_group(sets: list[OrderSet], parsed: ParsedOrderWorkbook, draft: OrderDraft, mode: str, prefix: str) -> bool:
    if not sets:
        st.caption("Комплектов в этом сегменте нет.")
        return False
    ordered_sets = sorted(sets, key=lambda order_set: (order_set.is_ungrouped, order_set.items[0].row if order_set.items else 0))
    image_paths = tuple(sorted({item.image_path for order_set in ordered_sets for item in order_set.items if item.image_path}))
    images = load_visible_images(parsed.upload_path, image_paths)
    changed = False
    for index, order_set in enumerate(ordered_sets):
        st.markdown(f'<div id="{prefix}-{index}"></div>', unsafe_allow_html=True)
        changed = _render_set_card(order_set, images, draft, mode) or changed
    return changed


def _render_category(category: str, sets: list[OrderSet], parsed: ParsedOrderWorkbook, draft: OrderDraft, mode: str) -> bool:
    changed = False
    if category == CATEGORY_ZERO:
        for segment in ("Нулевые с остатком", "0/0 — не было остатка"):
            segment_sets = [order_set for order_set in sets if order_set.zero_segment == segment]
            st.markdown(f"#### {segment}")
            changed = _render_category_segment(segment_sets, parsed, draft, mode, f"zero-{segment}") or changed
        return changed
    return _render_category_segment(sets, parsed, draft, mode, category)


def _render_category_segment(sets: list[OrderSet], parsed: ParsedOrderWorkbook, draft: OrderDraft, mode: str, prefix: str) -> bool:
    regular = [order_set for order_set in sets if not order_set.has_positive_tvp or order_set.has_negative_tvp]
    in_transit = [order_set for order_set in sets if order_set.has_positive_tvp and not order_set.has_negative_tvp]
    changed = _render_sets_group(regular, parsed, draft, mode, "regular-" + re.sub(r"\W+", "-", prefix))
    if in_transit:
        total_tvp = sum(item.positive_tvp for order_set in in_transit for item in order_set.items)
        with st.expander(f"Есть товар в пути — {len(in_transit)} комплектов · {total_tvp} шт. ТВП", expanded=False):
            st.caption("Этот блок свёрнут, чтобы сначала проверить позиции без уже оформленного заказа.")
            changed = _render_sets_group(in_transit, parsed, draft, mode, "tvp-" + re.sub(r"\W+", "-", prefix)) or changed
    return changed


def _render_overview(parsed: ParsedOrderWorkbook, order_sets: tuple[OrderSet, ...], mode: str) -> None:
    st.markdown('<div id="order-overview"></div>', unsafe_allow_html=True)
    st.markdown("## Сводка заказа")
    items = _ordered_items(order_sets)
    excluded_count = len(parsed.items) - len(items)
    errors = sum(bool(item.errors) for item in items)
    positive_tvp_sets = sum(order_set.has_positive_tvp and not order_set.has_negative_tvp for order_set in order_sets)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Тип заказа", mode)
    c2.metric("Комплектов", len(order_sets))
    c3.metric("Изделий", len(items))
    c4.metric("С ТВП", positive_tvp_sets)
    c5.metric("Ошибки", errors)
    st.caption(f"Исключено или относится к другому типу заказа: {excluded_count} строк.")
    if parsed.period:
        st.caption(f"Период продаж: {parsed.period}")
    if parsed.supplier:
        st.caption(f"Поставщик из отчёта: {parsed.supplier}")
    for warning in parsed.warnings:
        st.warning(warning)


def _render_order_workspace(parsed: ParsedOrderWorkbook, order_sets: tuple[OrderSet, ...], draft: OrderDraft, mode: str) -> None:
    st.markdown('<div id="order-workspace"></div>', unsafe_allow_html=True)
    st.markdown("## Комплекты по камням")
    stones = sorted({order_set.stone for order_set in order_sets})
    if not stones:
        st.warning("После исключений в выбранном типе заказа не осталось комплектов.")
        return
    default_index = stones.index(draft.selected_stone) if draft.selected_stone in stones else 0
    selected_stone = st.selectbox("Крупный блок", stones, index=default_index, key=f"supplier_order_stone::{mode}")
    if selected_stone != draft.selected_stone:
        draft.selected_stone = selected_stone
        _save_session_draft(draft)

    stone_sets = [order_set for order_set in order_sets if order_set.stone == selected_stone]
    counts = {category: sum(order_set.category == category for order_set in stone_sets) for category in CATEGORY_ORDER}
    cols = st.columns(4)
    for column, category in zip(cols, CATEGORY_ORDER):
        column.metric(CATEGORY_SHORT[category], counts[category])

    selected_category = st.segmented_control(
        "Категория комплектов",
        list(CATEGORY_ORDER),
        default=CATEGORY_TOP,
        key=f"supplier_order_category::{mode}::{selected_stone}",
    ) or CATEGORY_TOP
    st.caption(
        f"{CATEGORY_TONE[selected_category]} {selected_category}: "
        f"{counts[selected_category]} комплектов. Продажи изделий внутри комплекта не суммируются."
    )
    category_sets = [order_set for order_set in stone_sets if order_set.category == selected_category]
    changed = _render_category(selected_category, category_sets, parsed, draft, mode)
    if changed:
        _save_session_draft(draft)
        st.toast("Изменения автоматически сохранены", icon="💾")

    total_ordered = sum(max(0, draft.orders.get(item.key, 0)) for item in _ordered_items(order_sets))
    ordered_positions = sum(draft.orders.get(item.key, 0) > 0 for item in _ordered_items(order_sets))
    st.markdown("---")
    left, middle, right = st.columns([1.5, 1, 1.5])
    left.metric("Заказано позиций", ordered_positions)
    middle.metric("Всего изделий", total_ordered)
    if right.button("Подтвердить количества и перейти к размерам", type="primary", width="stretch", disabled=ordered_positions == 0):
        draft.stage = "rings"
        _save_session_draft(draft)
        st.rerun()


def _size_input_key(item: OrderItem, size: int, mode: str) -> str:
    return "ring_size::" + hashlib.sha1(f"{mode}|{item.key}|{size}".encode("utf-8")).hexdigest()


def _stock_check_key(item: OrderItem, mode: str) -> str:
    return "ring_stock_check::" + hashlib.sha1(f"{mode}|{item.key}".encode("utf-8")).hexdigest()


def ring_validation(item: OrderItem, draft: OrderDraft) -> tuple[int, int, bool, bool]:
    quantity = max(0, draft.orders.get(item.key, 0))
    values = draft.sizes.get(item.key, {})
    allocated = sum(max(0, safe_int(values.get(str(size), 0))) for size in RING_SIZES)
    stock_ok = item.working_stock <= 0 or bool(draft.stock_checked.get(item.key, False))
    return quantity, allocated, allocated == quantity, stock_ok


def _render_ring_sizes(parsed: ParsedOrderWorkbook, order_sets: tuple[OrderSet, ...], draft: OrderDraft, mode: str) -> None:
    st.markdown('<div id="order-rings"></div>', unsafe_allow_html=True)
    st.markdown("## Размеры колец")
    ordered_rings = [item for item in _ordered_items(order_sets) if item.is_ring and draft.orders.get(item.key, 0) > 0]
    if not ordered_rings:
        st.success("В текущем заказе нет колец. Можно переходить к Excel.")
        return
    image_paths = tuple(sorted({item.image_path for item in ordered_rings if item.image_path}))
    images = load_visible_images(parsed.upload_path, image_paths)
    complete = 0
    checked = 0
    changed = False

    for item in ordered_rings:
        quantity = max(0, draft.orders.get(item.key, 0))
        values = draft.sizes.setdefault(item.key, {})
        with st.container(border=True):
            left, right = st.columns([1, 4])
            with left:
                if item.image_path and item.image_path in images:
                    st.image(images[item.image_path], width="stretch")
                st.markdown(f"**{item.sku}**")
                st.caption(f"{canonical_stone(item.stone)} · к заказу {quantity}")
                if item.working_stock > 0:
                    st.warning(f"Свериться с остатком: {item.working_stock} шт.", icon="⚠️")
                    check_key = _stock_check_key(item, mode)
                    if check_key not in st.session_state:
                        st.session_state[check_key] = bool(draft.stock_checked.get(item.key, False))
                    check_value = st.checkbox("С остатком сверился", key=check_key)
                    if bool(draft.stock_checked.get(item.key, False)) != check_value:
                        draft.stock_checked[item.key] = check_value
                        changed = True
            with right:
                columns = st.columns(5)
                for index, size in enumerate(RING_SIZES):
                    size_key = str(size)
                    current = max(0, safe_int(values.get(size_key, 0)))
                    other_total = sum(max(0, safe_int(values.get(str(other), 0))) for other in RING_SIZES if other != size)
                    max_allowed = max(current, quantity - other_total)
                    widget_key = _size_input_key(item, size, mode)
                    if widget_key not in st.session_state:
                        st.session_state[widget_key] = current
                    with columns[index % 5]:
                        entered = st.number_input(
                            str(size),
                            min_value=0,
                            max_value=max(0, max_allowed),
                            step=1,
                            value=current,
                            key=widget_key,
                        )
                    entered = max(0, safe_int(entered))
                    if entered != current:
                        values[size_key] = entered
                        changed = True
                requested, allocated, allocation_ok, stock_ok = ring_validation(item, draft)
                if allocated > requested:
                    st.error(f"Распределено {allocated}, но к заказу доступно только {requested}. Уменьшите один из размеров.")
                elif allocated < requested:
                    st.warning(f"Распределено {allocated} из {requested} · осталось {requested - allocated}")
                else:
                    st.success(f"Распределено {allocated} из {requested}")
                if allocation_ok:
                    complete += 1
                if stock_ok:
                    checked += 1

    if changed:
        _save_session_draft(draft)
        st.toast("Размеры автоматически сохранены", icon="💾")
    st.caption(f"Размеры заполнены: {complete} из {len(ordered_rings)} · сверка с остатком: {checked} из {len(ordered_rings)}")


def _export_readiness(order_sets: tuple[OrderSet, ...], draft: OrderDraft) -> tuple[bool, list[str]]:
    ordered_items = [item for item in _ordered_items(order_sets) if draft.orders.get(item.key, 0) > 0]
    reasons: list[str] = []
    if not ordered_items:
        reasons.append("В заказе нет изделий.")
    rings = [item for item in ordered_items if item.is_ring]
    incomplete = []
    unchecked = []
    for item in rings:
        _, _, allocation_ok, stock_ok = ring_validation(item, draft)
        if not allocation_ok:
            incomplete.append(item.sku)
        if not stock_ok:
            unchecked.append(item.sku)
    if incomplete:
        reasons.append(f"Не завершены размеры для {len(incomplete)} колец.")
    if unchecked:
        reasons.append(f"Не подтверждена сверка с остатком для {len(unchecked)} колец.")
    return not reasons, reasons


def _render_export(parsed: ParsedOrderWorkbook, order_sets: tuple[OrderSet, ...], draft: OrderDraft, mode: str) -> None:
    st.markdown('<div id="order-export"></div>', unsafe_allow_html=True)
    st.markdown("## Итоговый Excel")
    ready, reasons = _export_readiness(order_sets, draft)
    ordered_items = [item for item in _ordered_items(order_sets) if draft.orders.get(item.key, 0) > 0]
    total_quantity = sum(draft.orders.get(item.key, 0) for item in ordered_items)
    rings = [item for item in ordered_items if item.is_ring]
    c1, c2, c3 = st.columns(3)
    c1.metric("Артикулов", len(ordered_items))
    c2.metric("Изделий", total_quantity)
    c3.metric("Колец", len(rings))
    if reasons:
        for reason in reasons:
            st.warning(reason)
        st.button("Скачать заказ в Excel", disabled=True, width="stretch")
        return
    with st.spinner("Формируем Excel с фотографиями..."):
        payload = build_supplier_excel(parsed, ordered_items, draft)
    safe_mode = "stones" if mode == ORDER_MODE_STONES else "pearls"
    st.download_button(
        "Скачать заказ в Excel",
        data=payload,
        file_name=f"supplier_order_{safe_mode}_{datetime.now().date().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
    )
    st.caption("В файле только: фото, артикул, камень, группа, количество к заказу и размеры колец.")


def _render_draft_tools(parsed: ParsedOrderWorkbook, draft: OrderDraft, mode: str) -> None:
    with st.expander("Черновик и резервная копия", expanded=False):
        left, right = st.columns(2)
        with left:
            st.download_button(
                "Скачать резервный JSON",
                data=draft_json_bytes(draft),
                file_name=f"order_draft_{'stones' if mode == ORDER_MODE_STONES else 'pearls'}.json",
                mime="application/json",
                width="stretch",
            )
        with right:
            imported = st.file_uploader("Восстановить JSON", type=["json"], key=f"order_draft_import::{mode}")
            if imported is not None and st.button("Применить резервную копию", key=f"apply_order_draft::{mode}", width="stretch"):
                try:
                    restored = import_draft_json(bytes(imported.getvalue()), parsed.source_hash, mode)
                    st.session_state[_draft_state_key(parsed.source_hash, mode)] = restored
                    st.success("Черновик восстановлен.")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))


def render_supplier_order_dashboard() -> None:
    parsed, _ = _render_upload()
    if parsed is None:
        return

    mode = st.segmented_control(
        "Какой заказ формируем?",
        list(ORDER_MODES),
        default=ORDER_MODE_STONES,
        key="supplier_order_mode",
    ) or ORDER_MODE_STONES
    draft = _get_session_draft(parsed, mode)
    _render_sidebar(parsed, draft)

    order_sets = _mode_sets(parsed, mode)
    _seed_defaults(draft, order_sets)
    _render_draft_tools(parsed, draft, mode)
    _render_overview(parsed, order_sets, mode)

    if draft.stage == "rings":
        if st.button("← Вернуться к количествам", width="stretch"):
            draft.stage = "order"
            _save_session_draft(draft)
            st.rerun()
        _render_ring_sizes(parsed, order_sets, draft, mode)
        _render_export(parsed, order_sets, draft, mode)
    else:
        _render_order_workspace(parsed, order_sets, draft, mode)
        # Export remains visible as a readiness preview, but sizes are completed
        # on the dedicated second stage.
        _render_export(parsed, order_sets, draft, mode)
