from pathlib import Path

import pandas as pd

from src.sonu import filter_sonu_metal_groups
from src.warehouse import WarehouseBundle, filter_warehouse_bundle

ROOT = Path(__file__).parent


def test_global_filter_is_rendered_for_every_mode_before_fx():
    source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
    main = source[source.index("def main() -> None:"):]
    assert main.index("mode = st.segmented_control(") < main.index("render_metal_filter_control(mode)")
    assert main.index("render_metal_filter_control(mode)") < main.index("render_global_fx_control()")
    assert 'key="global_metal_groups"' in source
    assert 'id="global-metal-filter"' in source
    assert 'render_warehouse_dashboard(selected_metal_groups())' in source
    assert 'render_sonu_order_dashboard()' in source
    assert 'if mode != "Заказ Sonu"' in source


def test_sonu_filter_uses_purity_groups():
    frame = pd.DataFrame({
        "SKU": ["S", "G", "O"],
        "Проба": ["B 925", "AU 585", "OTHER 0"],
    })
    silver = filter_sonu_metal_groups(frame, ["Серебро"])
    assert silver["SKU"].tolist() == ["S"]
    other = filter_sonu_metal_groups(frame, ["Другое"])
    assert other["SKU"].tolist() == ["O"]


def test_warehouse_filter_uses_material_and_filters_linked_operations():
    inventory = pd.DataFrame({
        "Артикул": ["S", "G", "O"],
        "Материал": ["Ag 925", "AU 585", "brass"],
        "Группа металла": ["Серебро", "Золото и платина", "Другое"],
    })
    operations = pd.DataFrame({"SKU": ["S", "G", "O"], "Изменение": [1, 1, 1]})
    bundle = WarehouseBundle(inventory, inventory.iloc[0:0].copy(), operations, pd.DataFrame(), pd.Timestamp.now().to_pydatetime())
    filtered = filter_warehouse_bundle(bundle, ["Серебро"])
    assert filtered.souvenirs["Артикул"].tolist() == ["S"]
    assert filtered.operations["SKU"].tolist() == ["S"]
