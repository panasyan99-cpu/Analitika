from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO
from math import ceil
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Any

from PIL import Image, ImageChops, ImageOps

from .models import Product

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

SILVER_DEFAULT_USD_VND = 26_500
SILVER_DEFAULT_COEFFICIENT = 10.0
SILVER_CIF_PERCENT = 11.0

SILVER_CATEGORIES = (
    "Бусины",
    "Пусеты",
    "Экстендеры",
    "Замки с экстендером",
    "Замки",
    "Бейлы",
    "Цепочки",
    "Основы для браслетов",
    "Основы для ожерелий",
)

# Business classification approved for the 18.07.2026 silver invoice.
_LINE_RULES: dict[int, dict[str, Any]] = {
    1: {"name": "Серебряная бусина 5 мм", "silver_category": "Бусины", "unit_label": "шт."},
    2: {"name": "Серебряные винтовые пусеты 15 мм", "silver_category": "Пусеты", "unit_label": "пара"},
    3: {"name": "Серебряный экстендер 5 см", "silver_category": "Экстендеры", "unit_label": "шт."},
    4: {"name": "Серебряный замок с экстендером", "silver_category": "Замки с экстендером", "unit_label": "комплект"},
    5: {"name": "Серебряный замок с экстендером", "silver_category": "Замки с экстендером", "unit_label": "комплект"},
    6: {"name": "Серебряный замок с экстендером", "silver_category": "Замки с экстендером", "unit_label": "комплект"},
    7: {"name": "Серебряный экстендер с сердцем 3 см", "silver_category": "Экстендеры", "unit_label": "шт."},
    8: {"name": "Серебряный U-образный замок с закрытым кольцом", "silver_category": "Замки", "unit_label": "комплект"},
    9: {"name": "Серебряный резьбовой бейл", "silver_category": "Бейлы", "unit_label": "шт."},
    10: {"name": "Серебряная бусина 4 мм", "silver_category": "Бусины", "unit_label": "шт."},
    11: {"name": "Серебряная цепочка 50 см", "silver_category": "Цепочки", "unit_label": "шт.", "sellable": True},
    12: {"name": "Серебряная основа браслета под вставку жемчуга", "silver_category": "Основы для браслетов", "unit_label": "шт."},
    13: {"name": "Серебряная основа браслета под вставку жемчуга", "silver_category": "Основы для браслетов", "unit_label": "шт."},
    14: {"name": "Серебряная основа браслета под вставку жемчуга", "silver_category": "Основы для браслетов", "unit_label": "шт."},
    15: {"name": "Серебряная основа ожерелья под вставку жемчуга", "silver_category": "Основы для ожерелий", "unit_label": "шт."},
    16: {"name": "Серебряная основа браслета под вставку жемчуга", "silver_category": "Основы для браслетов", "unit_label": "шт."},
    17: {"name": "Серебряная основа-цепочка с декоративными шариками 45 см", "silver_category": "Основы для ожерелий", "unit_label": "шт."},
    18: {"name": "Серебряная волнистая основа для ожерелья 45 см", "silver_category": "Основы для ожерелий", "unit_label": "шт."},
}


@dataclass(frozen=True)
class SilverInvoiceMeta:
    supplier: str
    invoice_date: str
    usd_rmb: float
    coefficient: float
    usd_vnd: int
    cif_percent: float = SILVER_CIF_PERCENT
    source_type: str = "silver_invoice"

    def to_dict(self) -> dict[str, Any]:
        return {
            "supplier": self.supplier,
            "invoice_date": self.invoice_date,
            "usd_rmb": self.usd_rmb,
            "coefficient": self.coefficient,
            "usd_vnd": self.usd_vnd,
            "cif_percent": self.cif_percent,
            "source_type": self.source_type,
        }


def silver_sale_vnd(purchase_usd: float | None, usd_vnd: int, coefficient: float) -> int:
    value = float(purchase_usd or 0.0) * max(int(usd_vnd or 0), 0) * max(float(coefficient or 0.0), 0.0)
    return int(ceil(value / 1000.0) * 1000) if value > 0 else 0


def _column_index(ref: str) -> int:
    match = re.match(r"[A-Z]+", str(ref or ""))
    if not match:
        return 0
    value = 0
    for char in match.group(0):
        value = value * 26 + ord(char) - 64
    return value - 1


def _as_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(str(value).replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return default


def _as_int(value: Any, default: int = 0) -> int:
    try:
        return int(float(str(value).replace(" ", "").replace(",", ".")))
    except (TypeError, ValueError):
        return default


def _excel_date(value: Any) -> str:
    numeric = _as_float(value, 0.0)
    if numeric <= 0:
        return ""
    try:
        return (datetime(1899, 12, 30) + timedelta(days=numeric)).date().isoformat()
    except (OverflowError, ValueError):
        return ""


def _read_values(archive: zipfile.ZipFile) -> dict[int, list[Any]]:
    ns_main = f"{{{MAIN_NS}}}"
    shared: list[str] = []
    if "xl/sharedStrings.xml" in archive.namelist():
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        for item in root.findall(f"{ns_main}si"):
            shared.append("".join(node.text or "" for node in item.iter(f"{ns_main}t")))

    sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    values_by_row: dict[int, list[Any]] = {}
    for row_node in sheet.findall(f".//{ns_main}row"):
        row_index = int(row_node.attrib["r"])
        values: list[Any] = [None] * 16
        for cell in row_node.findall(f"{ns_main}c"):
            column = _column_index(cell.attrib.get("r", "A1"))
            if column >= len(values):
                continue
            cell_type = cell.attrib.get("t")
            value_node = cell.find(f"{ns_main}v")
            inline = cell.find(f"{ns_main}is")
            value: Any = None
            if cell_type == "s" and value_node is not None:
                value = shared[int(value_node.text or 0)]
            elif cell_type == "inlineStr" and inline is not None:
                value = "".join(node.text or "" for node in inline.iter(f"{ns_main}t"))
            elif value_node is not None:
                raw = value_node.text or ""
                try:
                    numeric = float(raw)
                    value = int(numeric) if numeric.is_integer() else numeric
                except ValueError:
                    value = raw
            values[column] = value
        values_by_row[row_index] = values
    return values_by_row


def is_silver_invoice(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            rows = _read_values(archive)
        headers = [str(value or "").strip().casefold() for value in rows.get(7, [])]
        required = {"no", "photo", "code", "plating", "size", "quantity", "weight", "cif price usd"}
        return required.issubset(set(headers))
    except Exception:
        return False


def _optimized_photo(raw: bytes, target: Path) -> None:
    with Image.open(BytesIO(raw)) as source:
        image = ImageOps.exif_transpose(source).convert("RGB")
        background = Image.new("RGB", image.size, "white")
        difference = ImageChops.difference(image, background).convert("L")
        # Ignore tiny JPEG noise and near-white backgrounds.
        mask = difference.point(lambda value: 255 if value > 18 else 0)
        bbox = mask.getbbox()
        if bbox:
            left, top, right, bottom = bbox
            pad_x = max(16, int((right - left) * 0.08))
            pad_y = max(16, int((bottom - top) * 0.08))
            bbox = (
                max(0, left - pad_x),
                max(0, top - pad_y),
                min(image.width, right + pad_x),
                min(image.height, bottom + pad_y),
            )
            image = image.crop(bbox)
        image.thumbnail((1000, 1000), Image.Resampling.LANCZOS)
        canvas_side = max(image.width, image.height, 500)
        canvas = Image.new("RGB", (canvas_side, canvas_side), "white")
        canvas.paste(image, ((canvas_side - image.width) // 2, (canvas_side - image.height) // 2))
        target.parent.mkdir(parents=True, exist_ok=True)
        canvas.save(target, format="JPEG", quality=86, optimize=True, progressive=True)


def _extract_row_images(archive: zipfile.ZipFile, image_dir: Path) -> dict[int, str]:
    rel_path = "xl/drawings/_rels/drawing1.xml.rels"
    drawing_path = "xl/drawings/drawing1.xml"
    if rel_path not in archive.namelist() or drawing_path not in archive.namelist():
        return {}
    rel_root = ET.fromstring(archive.read(rel_path))
    rels = {node.attrib["Id"]: node.attrib["Target"].replace("../", "xl/") for node in rel_root}
    drawing = ET.fromstring(archive.read(drawing_path))
    ns = {"xdr": XDR_NS, "a": A_NS}
    row_images: dict[int, str] = {}
    for index, anchor in enumerate(list(drawing), start=1):
        from_node = anchor.find("xdr:from", ns)
        picture = anchor.find("xdr:pic", ns)
        if from_node is None or picture is None:
            continue
        row_number = int(from_node.find("xdr:row", ns).text or 0) + 1
        blip = picture.find(".//a:blip", ns)
        if blip is None:
            continue
        relationship = blip.attrib.get(f"{{{REL_NS}}}embed")
        media = rels.get(str(relationship or ""))
        if not media or media not in archive.namelist():
            continue
        target = image_dir / f"silver_row_{row_number}_{index}.jpg"
        try:
            _optimized_photo(archive.read(media), target)
            row_images[row_number] = str(target)
        except Exception:
            # A bad source image must not block the invoice itself.
            continue
    return row_images


def parse_silver_invoice(path: Path, image_dir: Path) -> tuple[list[Product], SilverInvoiceMeta]:
    with zipfile.ZipFile(path) as archive:
        rows = _read_values(archive)
        headers = [str(value or "").strip().casefold() for value in rows.get(7, [])]
        required = {
            "no", "photo", "code", "plating", "size", "quantity", "weight",
            "silver/g", "labour/g", "price/g", "amount", "cif price usd",
            "sell price usd", "sell price vnd",
        }
        missing = sorted(required.difference(set(headers)))
        if missing:
            raise ValueError("В серебряном инвойсе отсутствуют колонки: " + ", ".join(missing))
        row_images = _extract_row_images(archive, image_dir)

    usd_rmb = _as_float(rows.get(2, [None] * 12)[11], 6.71)
    coefficient = _as_float(rows.get(2, [None] * 13)[12], SILVER_DEFAULT_COEFFICIENT)
    usd_vnd = _as_int(rows.get(2, [None] * 14)[13], SILVER_DEFAULT_USD_VND)
    supplier = str(rows.get(1, [""])[0] or "").strip() or "TIAN YI DA JEWELLERY Co.,LTD"
    invoice_date = _excel_date(rows.get(6, [None] * 10)[9])
    meta = SilverInvoiceMeta(
        supplier=supplier,
        invoice_date=invoice_date,
        usd_rmb=usd_rmb,
        coefficient=coefficient,
        usd_vnd=usd_vnd,
    )

    products: list[Product] = []
    for row_number in range(8, 26):
        row = rows.get(row_number, [None] * 16)
        line = _as_int(row[0])
        quantity = _as_int(row[5])
        if line <= 0 or quantity <= 0:
            continue
        rule = _LINE_RULES.get(line, {})
        total_weight_g = _as_float(row[6])
        purchase_usd = _as_float(row[11])
        product = Product(
            number=line,
            boxes="",
            sku="",  # assigned from the next free SIL sequence after Baserow catalog lookup
            qty_document=quantity,
            description=str(rule.get("name") or row[2] or "").strip(),
            category="Аксессуары",
            material="Silver",
            stone="",
            color=str(row[3] or "").strip(),
            unit_weight_kg=(total_weight_g / quantity / 1000.0) if quantity else None,
            image_path=row_images.get(row_number, ""),
            received=False,
            actual_manual=None,
            comment="Серебро 925 · импорт из инвойса",
            checked=True,
            recognition="Серебряный инвойс 18.07.2026",
            name=str(rule.get("name") or row[2] or "").strip(),
            silver_category=str(rule.get("silver_category") or "Основы для ожерелий"),
            silver_925=True,
            plating=str(row[3] or "").strip(),
            size=str(row[4] or "").strip(),
            unit_label=str(rule.get("unit_label") or "шт."),
            sellable=bool(rule.get("sellable", False)),
            original_name=str(row[2] or "").strip(),
            total_weight_g=total_weight_g,
            silver_rmb_per_g=_as_float(row[7]),
            labour_rmb_per_g=_as_float(row[8]),
            price_rmb_per_g=_as_float(row[9]),
            amount_rmb=_as_float(row[10]),
            usd_rmb_rate=usd_rmb,
            cif_percent=SILVER_CIF_PERCENT,
            purchase_usd_per_unit=purchase_usd,
            invoice_sale_usd=_as_float(row[12]),
            invoice_usd_vnd_rate=usd_vnd,
            invoice_coefficient=coefficient,
            invoice_sale_vnd=_as_int(row[13]),
        )
        products.append(product)
    if len(products) != 18:
        raise ValueError(f"Ожидалось 18 позиций серебра, распознано: {len(products)}.")
    return products, meta
