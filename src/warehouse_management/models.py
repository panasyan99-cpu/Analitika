from __future__ import annotations

from dataclasses import asdict, dataclass
import re
from typing import Any, Optional

DEFAULT_MIN_BALANCE = 10
LOW_STOCK_NOTICE = 15

STONE_NAME_ALIASES = {
    "lapis lazuli": "Lapis Lazurite",
    "lapis lazurite": "Lapis Lazurite",
    "black agate": "Agate",
    "green agate": "Agate",
    "blue agate": "Agate",
    "blue lace agate": "Agate",
    "red agate": "Agate",
    "yellow agate": "Agate",
    "white agate": "Agate",
    "grey agate": "Agate",
    "gray agate": "Agate",
    "pink agate": "Agate",
    "purple agate": "Agate",
    "brown agate": "Agate",
    "agate": "Agate",
    "black onyx": "Onyx",
    "onyx": "Onyx",
    "white howlite": "Howlite",
    "howlite": "Howlite",
    "green fluorite": "Fluorite",
    "fluorite": "Fluorite",
    "red aventurine": "Aventurine",
    "green aventurine": "Aventurine",
    "aventurine": "Aventurine",
}

CANONICAL_STONES = [
    "Agate", "Amazonite", "Amethyst", "Aventurine", "Citrine", "Coral",
    "Fluorite", "Garnet", "Howlite", "Jasper", "Labradorite",
    "Lapis Lazurite", "Malachite", "Moonstone", "Multistone", "Obsidian",
    "Onyx", "Opalite", "Pearl", "Picture Jasper", "Quartz", "Rhodonite",
    "Rose Quartz", "Smoky Quartz", "Sodalite", "Tiger Eye", "Tourmaline",
    "Turquoise", "Zircon",
]


def split_multi_values(value: str) -> list[str]:
    return [
        item.strip()
        for item in re.split(r"\s*(?:;|\||\+|,|\r?\n)\s*", str(value or ""))
        if item.strip()
    ]


def canonical_stone_name(value: str) -> str:
    item = str(value or "").strip()
    if not item:
        return ""
    return STONE_NAME_ALIASES.get(item.casefold(), item)


def normalize_stone_names(value: str) -> str:
    parts: list[str] = []
    seen: set[str] = set()
    for raw in split_multi_values(value):
        normalized = canonical_stone_name(raw)
        key = normalized.casefold()
        if normalized and key not in seen:
            parts.append(normalized)
            seen.add(key)
    return "; ".join(parts)


@dataclass
class Product:
    number: int
    boxes: str
    sku: str
    qty_document: int
    description: str
    category: str
    material: str
    stone: str
    color: str
    unit_weight_kg: Optional[float]
    image_path: str
    received: bool = False
    actual_manual: Optional[int] = None
    comment: str = ""
    checked: bool = False
    recognition: str = ""
    name: str = ""
    silver_category: str = ""
    silver_925: bool = False
    plating: str = ""
    size: str = ""
    unit_label: str = "шт."
    sellable: bool = False
    original_name: str = ""
    total_weight_g: Optional[float] = None
    silver_rmb_per_g: Optional[float] = None
    labour_rmb_per_g: Optional[float] = None
    price_rmb_per_g: Optional[float] = None
    amount_rmb: Optional[float] = None
    usd_rmb_rate: Optional[float] = None
    cif_percent: Optional[float] = None
    purchase_usd_per_unit: Optional[float] = None
    invoice_sale_usd: Optional[float] = None
    invoice_usd_vnd_rate: Optional[int] = None
    invoice_coefficient: Optional[float] = None
    invoice_sale_vnd: Optional[int] = None

    def __post_init__(self) -> None:
        self.sku = str(self.sku or "").strip()
        self.boxes = str(self.boxes or "").strip()
        self.qty_document = max(int(self.qty_document or 0), 0)
        self.stone = normalize_stone_names(self.stone)
        self.name = str(self.name or self.description or "").strip()
        self.unit_label = str(self.unit_label or "шт.").strip()
        self.plating = str(self.plating or "").strip()
        self.size = str(self.size or "").strip()
        self.silver_category = str(self.silver_category or "").strip()

    @property
    def actual_qty(self) -> Optional[int]:
        if not self.received:
            return None
        return self.actual_manual if self.actual_manual is not None else self.qty_document

    @property
    def variance(self) -> Optional[int]:
        if self.actual_qty is None:
            return None
        return self.actual_qty - self.qty_document

    @property
    def waiting_qty(self) -> int:
        return max(self.qty_document - (self.actual_qty or 0), 0)

    @property
    def status(self) -> str:
        if not self.received:
            return "Ожидается"
        if self.waiting_qty > 0:
            return "Частично получено"
        return "Получено"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "Product":
        allowed = {field.name for field in cls.__dataclass_fields__.values()}
        clean = {key: value.get(key) for key in allowed if key in value}
        if "stone" in clean:
            clean["stone"] = normalize_stone_names(str(clean["stone"] or ""))
        return cls(**clean)


@dataclass(frozen=True)
class CatalogItem:
    row_id: int
    sku: str
    section: str
    balance: int
    boxes: str = ""
    category: str = ""
    material: str = ""
    stone: str = ""
    color: str = ""
    photo: Any = None
    min_balance: int = DEFAULT_MIN_BALANCE
    active: bool = True
    name: str = ""
    silver_category: str = ""
    silver_925: bool = False
    plating: str = ""
    size: str = ""
    unit_label: str = "шт."
    sellable: bool = False
    purchase_usd_per_unit: Optional[float] = None
    raw: dict[str, Any] | None = None

    @property
    def search_text(self) -> str:
        return " ".join(
            part
            for part in [
                self.sku, self.name, self.category, self.silver_category, self.material,
                self.stone, self.color, self.plating, self.size, self.boxes,
            ]
            if part
        )


@dataclass(frozen=True)
class SupplySummary:
    row_id: int
    supply_id: str
    date: str
    supplier: str
    status: str
    sku_total: int
    sku_received: int
    qty_document: int
    qty_received: int
    qty_waiting: int
    raw: dict[str, Any]
