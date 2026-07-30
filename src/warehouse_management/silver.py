from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO
from math import ceil
from pathlib import Path, PurePosixPath
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

SILVER_DEFAULT_USD_RMB = 6.71
SILVER_DEFAULT_USD_VND = 26_500
SILVER_DEFAULT_COEFFICIENT = 10.0
SILVER_CIF_PERCENT = 0.0

# Keep the historical 18.07 groups and add the approved 30.06 groups.
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
    "Замки и концевики",
    "Бусины и разделители",
    "Трубки и проставки",
    "Бирки / декоративные элементы",
    "Кольца и соединители",
    "Цепи / полуфабрикаты",
    "Готовые цепочки / основы",
)

# Business classification approved for the enriched 18.07.2026 silver invoice.
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


def _normalize_code(value: Any) -> str:
    return (
        re.sub(r"\s+", "", str(value or ""))
        .casefold()
        .replace("×", "*")
        .replace("х", "x")
    )


# Business classification approved by the user for the raw 30.06.2026 invoice.
_RAW_CODE_RULES: dict[str, dict[str, Any]] = {
    _normalize_code("2.6m皮绳管+9m水滴扣"): {
        "name": "Серебряный концевик для шнура 2,6 мм с карабином 9 мм",
        "silver_category": "Замки и концевики",
        "unit_label": "комплект",
    },
    _normalize_code("3m弯刀橄榄珠"): {
        "name": "Серебряный изогнутый разделитель «оливка» 3 мм",
        "silver_category": "Бусины и разделители",
        "unit_label": "шт.",
    },
    _normalize_code("1.5*20车花s管"): {
        "name": "Серебряная S-образная гранёная трубка 1,5×20 мм",
        "silver_category": "Трубки и проставки",
        "unit_label": "шт.",
    },
    _normalize_code("1.5*10车花直管"): {
        "name": "Серебряная прямая гранёная трубка 1,5×10 мм",
        "silver_category": "Трубки и проставки",
        "unit_label": "шт.",
    },
    _normalize_code("5.0轻生圈"): {
        "name": "Серебряный круглый замок Spring Ring 5 мм",
        "silver_category": "Замки и концевики",
        "unit_label": "шт.",
    },
    _normalize_code("3x7舌形牌"): {
        "name": "Серебряная бирка-язычок 3×7 мм",
        "silver_category": "Бирки / декоративные элементы",
        "unit_label": "шт.",
    },
    _normalize_code("60x4开口圈"): {
        "name": "Серебряное открытое соединительное кольцо 0,6×4 мм",
        "silver_category": "Кольца и соединители",
        "unit_label": "шт.",
    },
    _normalize_code("5mm西瓜珠"): {
        "name": "Серебряная бусина «арбуз» 5 мм",
        "silver_category": "Бусины и разделители",
        "unit_label": "шт.",
    },
    _normalize_code("5mm猫眼（大孔）"): {
        "name": "Серебряная бусина «кошачий глаз» 5 мм, большое отверстие",
        "silver_category": "Бусины и разделители",
        "unit_label": "шт.",
    },
    _normalize_code("275车侧身-半成品"): {
        "name": "Серебряная цепь-полуфабрикат по метражу",
        "silver_category": "Цепи / полуфабрикаты",
        "unit_label": "м",
    },
    _normalize_code("275车侧身小正心万能"): {
        "name": "Серебряная готовая цепочка 45 см с застёжкой",
        "silver_category": "Готовые цепочки / основы",
        "unit_label": "шт.",
    },
    _normalize_code("30c车十字包方珠滴胶"): {
        "name": "Серебряная декоративная цепочка / основа 50 см",
        "silver_category": "Готовые цепочки / основы",
        "unit_label": "шт.",
    },
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
    source_variant: str = "enriched"
    auto_receive: bool = False
    price_currency: str = "RMB"
    purchase_price_calculated: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "supplier": self.supplier,
            "invoice_date": self.invoice_date,
            "usd_rmb": self.usd_rmb,
            "coefficient": self.coefficient,
            "usd_vnd": self.usd_vnd,
            "cif_percent": self.cif_percent,
            "source_type": self.source_type,
            "source_variant": self.source_variant,
            "auto_receive": self.auto_receive,
            "price_currency": self.price_currency,
            "purchase_price_calculated": self.purchase_price_calculated,
        }


def silver_sale_vnd(purchase_usd: float | None, usd_vnd: int, coefficient: float) -> int:
    value = float(purchase_usd or 0.0) * max(int(usd_vnd or 0), 0) * max(float(coefficient or 0.0), 0.0)
    return int(ceil(value / 1000.0) * 1000) if value > 0 else 0


def refresh_calculated_silver_prices(
    products: list[Product],
    *,
    usd_vnd: int,
    coefficient: float,
) -> list[Product]:
    """Synchronize raw-invoice retail history after an operator price edit.

    The 30.06 supplier file has no USD or VND retail columns. Its import-history
    values must therefore be derived from the final, operator-approved purchase
    USD value and the active warehouse rate/coefficient at the moment the supply
    is created. The enriched 18.07 file is intentionally excluded by the caller
    because its invoice retail columns are historical source data.
    """
    safe_rate = max(int(usd_vnd or 0), 0)
    safe_coefficient = max(float(coefficient or 0.0), 0.0)
    for product in products:
        purchase = max(float(product.purchase_usd_per_unit or 0.0), 0.0)
        product.invoice_sale_usd = purchase * safe_coefficient if purchase > 0 else None
        product.invoice_usd_vnd_rate = safe_rate
        product.invoice_coefficient = safe_coefficient
        product.invoice_sale_vnd = silver_sale_vnd(purchase, safe_rate, safe_coefficient)
    return products


def _raw_price_issue(
    *,
    line: int,
    weight_g: float,
    silver_rmb_per_g: float,
    labour_rmb_per_g: float,
    price_rmb_per_g: float,
    amount_rmb: float,
) -> str:
    """Return a clear error when the supplier purchase arithmetic is broken."""
    if weight_g <= 0 or price_rmb_per_g <= 0 or amount_rmb <= 0:
        return f"строка {line}: вес, цена/г и сумма должны быть больше нуля"
    component_tolerance = 0.011
    if abs((silver_rmb_per_g + labour_rmb_per_g) - price_rmb_per_g) > component_tolerance:
        return (
            f"строка {line}: price/g {price_rmb_per_g:g} не равна "
            f"silver/g + labour/g ({silver_rmb_per_g:g} + {labour_rmb_per_g:g})"
        )
    expected_amount = weight_g * price_rmb_per_g
    amount_tolerance = max(0.02, abs(amount_rmb) * 1e-8)
    if abs(expected_amount - amount_rmb) > amount_tolerance:
        return (
            f"строка {line}: Amount {amount_rmb:g} не равна "
            f"Weight × price/g ({weight_g:g} × {price_rmb_per_g:g} = {expected_amount:g})"
        )
    return ""


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
        match = re.search(r"[-+]?\d+(?:[.,]\d+)?", str(value or ""))
        if not match:
            return default
        try:
            return float(match.group(0).replace(",", "."))
        except ValueError:
            return default


def _as_int(value: Any, default: int = 0) -> int:
    numeric = _as_float(value, float(default))
    try:
        return int(numeric)
    except (TypeError, ValueError, OverflowError):
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
        values: list[Any] = [None] * 20
        for cell in row_node.findall(f"{ns_main}c"):
            column = _column_index(cell.attrib.get("r", "A1"))
            if column >= len(values):
                continue
            cell_type = cell.attrib.get("t")
            value_node = cell.find(f"{ns_main}v")
            formula_node = cell.find(f"{ns_main}f")
            inline = cell.find(f"{ns_main}is")
            value: Any = None
            if cell_type == "s" and value_node is not None and value_node.text is not None:
                value = shared[int(value_node.text or 0)]
            elif cell_type == "inlineStr" and inline is not None:
                value = "".join(node.text or "" for node in inline.iter(f"{ns_main}t"))
            elif value_node is not None and value_node.text not in (None, ""):
                raw = value_node.text or ""
                try:
                    numeric = float(raw)
                    value = int(numeric) if numeric.is_integer() else numeric
                except ValueError:
                    value = raw
            elif formula_node is not None:
                # WPS DISPIMG may be stored as a formula without a cached value.
                value = "=" + str(formula_node.text or "")
            values[column] = value
        values_by_row[row_index] = values
    return values_by_row


def _headers(rows: dict[int, list[Any]]) -> list[str]:
    return [str(value or "").strip().casefold() for value in rows.get(7, [])]


def is_silver_invoice(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            rows = _read_values(archive)
        required = {
            "no", "photo", "code", "plating", "size", "quantity", "weight",
            "silver/g", "labour/g", "price/g", "amount",
        }
        return required.issubset(set(_headers(rows)))
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


def _relationship_map(archive: zipfile.ZipFile, rel_path: str, *, base: str) -> dict[str, str]:
    if rel_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rel_path))
    result: dict[str, str] = {}
    for node in list(root):
        rel_id = str(node.attrib.get("Id") or "")
        target = str(node.attrib.get("Target") or "").replace("\\", "/")
        if not rel_id or not target:
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = str(PurePosixPath(base) / target)
            # PurePosixPath keeps '..'; normalize manually.
            parts: list[str] = []
            for part in resolved.split("/"):
                if part == "..":
                    if parts:
                        parts.pop()
                elif part not in {"", "."}:
                    parts.append(part)
            resolved = "/".join(parts)
        result[rel_id] = resolved
    return result


def _extract_drawing_images(archive: zipfile.ZipFile, image_dir: Path) -> dict[int, str]:
    rel_path = "xl/drawings/_rels/drawing1.xml.rels"
    drawing_path = "xl/drawings/drawing1.xml"
    if drawing_path not in archive.namelist():
        return {}
    rels = _relationship_map(archive, rel_path, base="xl/drawings")
    drawing = ET.fromstring(archive.read(drawing_path))
    ns = {"xdr": XDR_NS, "a": A_NS}
    row_images: dict[int, str] = {}
    for index, anchor in enumerate(list(drawing), start=1):
        from_node = anchor.find("xdr:from", ns)
        picture = anchor.find("xdr:pic", ns)
        if from_node is None or picture is None:
            continue
        row_node = from_node.find("xdr:row", ns)
        if row_node is None:
            continue
        row_number = int(row_node.text or 0) + 1
        blip = picture.find(".//a:blip", ns)
        if blip is None:
            continue
        relationship = blip.attrib.get(f"{{{REL_NS}}}embed")
        media = rels.get(str(relationship or ""))
        if not media or media not in archive.namelist():
            continue
        target = image_dir / f"silver_row_{row_number}_drawing_{index}.jpg"
        try:
            _optimized_photo(archive.read(media), target)
            row_images[row_number] = str(target)
        except Exception:
            continue
    return row_images


def _extract_wps_cell_images(
    archive: zipfile.ZipFile,
    rows: dict[int, list[Any]],
    image_dir: Path,
) -> dict[int, str]:
    xml_path = "xl/cellimages.xml"
    rel_path = "xl/_rels/cellimages.xml.rels"
    if xml_path not in archive.namelist() or rel_path not in archive.namelist():
        return {}

    rels = _relationship_map(archive, rel_path, base="xl")
    root = ET.fromstring(archive.read(xml_path))
    id_to_media: dict[str, str] = {}
    for cell_image in list(root):
        name_node = cell_image.find(f".//{{{XDR_NS}}}cNvPr")
        blip = cell_image.find(f".//{{{A_NS}}}blip")
        if name_node is None or blip is None:
            continue
        image_id = str(name_node.attrib.get("name") or "")
        relationship = str(blip.attrib.get(f"{{{REL_NS}}}embed") or "")
        media = rels.get(relationship)
        if image_id and media:
            id_to_media[image_id] = media

    row_images: dict[int, str] = {}
    for row_number, values in rows.items():
        photo_value = str(values[1] or "") if len(values) > 1 else ""
        match = re.search(r"(ID_[A-Za-z0-9_]+)", photo_value)
        if not match:
            continue
        image_id = match.group(1)
        media = id_to_media.get(image_id)
        if not media or media not in archive.namelist():
            continue
        target = image_dir / f"silver_row_{row_number}_cell.jpg"
        try:
            _optimized_photo(archive.read(media), target)
            row_images[row_number] = str(target)
        except Exception:
            continue
    return row_images


def _extract_row_images(
    archive: zipfile.ZipFile,
    rows: dict[int, list[Any]],
    image_dir: Path,
) -> dict[int, str]:
    # A supplier file may mix one normal Excel drawing with WPS DISPIMG cell images.
    result = _extract_drawing_images(archive, image_dir)
    for row_number, path in _extract_wps_cell_images(archive, rows, image_dir).items():
        result.setdefault(row_number, path)
    return result


def _header_indices(headers: list[str]) -> dict[str, int]:
    return {name: index for index, name in enumerate(headers) if name}


def _row_value(row: list[Any], indexes: dict[str, int], name: str) -> Any:
    index = indexes.get(name)
    if index is None or index >= len(row):
        return None
    return row[index]


def _calculated_purchase_usd(
    amount_rmb: float,
    quantity: int,
    usd_rmb: float,
    cif_percent: float = 0.0,
) -> float:
    """Convert the supplier purchase total from RMB to USD per stock unit.

    The raw 30.06 invoice already contains purchase prices. No extra CIF or other
    uplift is added here. ``cif_percent`` remains in the signature only for
    backwards-compatible saved drafts and Baserow schema fields.
    """
    if amount_rmb <= 0 or quantity <= 0 or usd_rmb <= 0:
        return 0.0
    return amount_rmb / usd_rmb / quantity


def parse_silver_invoice(path: Path, image_dir: Path) -> tuple[list[Product], SilverInvoiceMeta]:
    with zipfile.ZipFile(path) as archive:
        rows = _read_values(archive)
        headers = _headers(rows)
        required = {
            "no", "photo", "code", "plating", "size", "quantity", "weight",
            "silver/g", "labour/g", "price/g", "amount",
        }
        missing = sorted(required.difference(set(headers)))
        if missing:
            raise ValueError("В серебряном инвойсе отсутствуют колонки: " + ", ".join(missing))
        row_images = _extract_row_images(archive, rows, image_dir)

    indexes = _header_indices(headers)
    enriched = {"cif price usd", "sell price usd", "sell price vnd"}.issubset(set(headers))
    source_variant = "enriched_2026_07_18" if enriched else "supplier_raw_2026_06_30"
    # Registration and physical receipt are always separate warehouse steps.
    auto_receive = False

    usd_rmb = _as_float(rows.get(2, [None] * 12)[11], SILVER_DEFAULT_USD_RMB)
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
        source_variant=source_variant,
        auto_receive=auto_receive,
        purchase_price_calculated=not enriched,
    )

    products: list[Product] = []
    unknown_codes: list[str] = []
    price_issues: list[str] = []
    for row_number in sorted(number for number in rows if number >= 8):
        row = rows.get(row_number, [None] * 20)
        line = _as_int(_row_value(row, indexes, "no"))
        if line <= 0:
            continue
        quantity = _as_int(_row_value(row, indexes, "quantity"))
        if quantity <= 0:
            continue

        original_name = str(_row_value(row, indexes, "code") or "").strip()
        if enriched:
            rule = _LINE_RULES.get(line, {})
        else:
            rule = _RAW_CODE_RULES.get(_normalize_code(original_name), {})
            if not rule:
                unknown_codes.append(original_name or f"строка {line}")

        total_weight_g = _as_float(_row_value(row, indexes, "weight"))
        silver_rmb_per_g = _as_float(_row_value(row, indexes, "silver/g"))
        labour_rmb_per_g = _as_float(_row_value(row, indexes, "labour/g"))
        price_rmb_per_g = _as_float(_row_value(row, indexes, "price/g"))
        amount_rmb = _as_float(_row_value(row, indexes, "amount"))
        if not enriched:
            issue = _raw_price_issue(
                line=line,
                weight_g=total_weight_g,
                silver_rmb_per_g=silver_rmb_per_g,
                labour_rmb_per_g=labour_rmb_per_g,
                price_rmb_per_g=price_rmb_per_g,
                amount_rmb=amount_rmb,
            )
            if issue:
                price_issues.append(issue)
        source_purchase_usd = _as_float(_row_value(row, indexes, "cif price usd"))
        purchase_usd = source_purchase_usd or _calculated_purchase_usd(
            amount_rmb,
            quantity,
            usd_rmb,
            SILVER_CIF_PERCENT,
        )
        invoice_sale_usd = _as_float(_row_value(row, indexes, "sell price usd"))
        if not invoice_sale_usd and purchase_usd > 0:
            invoice_sale_usd = purchase_usd * coefficient
        invoice_sale_vnd = _as_int(_row_value(row, indexes, "sell price vnd"))
        if not invoice_sale_vnd and purchase_usd > 0:
            invoice_sale_vnd = silver_sale_vnd(purchase_usd, usd_vnd, coefficient)

        display_name = str(rule.get("name") or original_name).strip()
        product = Product(
            number=line,
            boxes="",
            sku="",  # assigned from the next free SIL sequence after Baserow catalog lookup
            qty_document=quantity,
            description=display_name,
            category="",  # Silver grouping is stored in «Серебряная категория».
            material="Silver",
            stone="",
            color=str(_row_value(row, indexes, "plating") or "").strip(),
            unit_weight_kg=(total_weight_g / quantity / 1000.0) if quantity else None,
            image_path=row_images.get(row_number, ""),
            received=False,
            actual_manual=None,
            comment="Серебро 925 · импорт из инвойса · товар ожидается",
            checked=True,
            recognition=(
                "Серебряный инвойс 18.07.2026"
                if enriched
                else "Серебряный инвойс 30.06.2026"
            ),
            name=display_name,
            silver_category=str(rule.get("silver_category") or ""),
            silver_925=True,
            plating=str(_row_value(row, indexes, "plating") or "").strip(),
            size=str(_row_value(row, indexes, "size") or "").strip(),
            unit_label=str(rule.get("unit_label") or "шт."),
            sellable=bool(rule.get("sellable", False)),
            original_name=original_name,
            total_weight_g=total_weight_g,
            silver_rmb_per_g=silver_rmb_per_g,
            labour_rmb_per_g=labour_rmb_per_g,
            price_rmb_per_g=price_rmb_per_g,
            amount_rmb=amount_rmb,
            usd_rmb_rate=usd_rmb,
            cif_percent=SILVER_CIF_PERCENT,
            purchase_usd_per_unit=purchase_usd,
            invoice_sale_usd=invoice_sale_usd,
            invoice_usd_vnd_rate=usd_vnd,
            invoice_coefficient=coefficient,
            invoice_sale_vnd=invoice_sale_vnd,
        )
        products.append(product)

    if not products:
        raise ValueError("В серебряном инвойсе не найдено ни одной товарной позиции.")
    expected = 18 if enriched else 14
    if len(products) != expected:
        raise ValueError(f"Ожидалось {expected} позиций серебра, распознано: {len(products)}.")
    if unknown_codes:
        raise ValueError(
            "Не удалось автоматически классифицировать позиции: "
            + ", ".join(unknown_codes[:20])
            + ". Добавьте правила перед созданием поставки."
        )
    if price_issues:
        raise ValueError(
            "Проверка закупочных цен не пройдена: " + "; ".join(price_issues[:20])
        )
    return products, meta
